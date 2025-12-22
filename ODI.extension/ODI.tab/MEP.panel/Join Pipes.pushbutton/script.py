# -*- coding: utf-8 -*-
"""
Join Pipes Command
Select two pipes to join them.
- Logic:
    1. Select Pipe 1 (to Modify/Extend)
    2. Select Pipe 2 (Reference)
    3. Detects if Parallel/Collinear -> Union
    4. Detects if Intersecting at Corner -> Elbow
    5. Detects if Intersecting along Body -> Tap/Takeoff
"""
import math
import clr
# clr.AddReference('RevitAPI')
# clr.AddReference('RevitAPIUI')

from pyrevit import revit, forms, script
from pyrevit import DB
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

# --- Utilities ---

class Logger:
    def __init__(self):
        self.logs = []
    
    def info(self, msg):
        self.logs.append("INFO: " + str(msg))
        print("INFO: " + str(msg))
        
    def error(self, msg):
        self.logs.append("**ERROR:** " + str(msg))
        print("ERROR: " + str(msg))
    
    def show(self):
        if self.logs:
            output = script.get_output()
            output.print_md("### Join Pipes Operation Log")
            for log in self.logs:
                output.print_md(log)

def get_id_value(element_id):
    """Safely gets the integer value of an ElementId (Revit 2024+ compatible)."""
    if hasattr(element_id, "Value"): # Revit 2024+
        return element_id.Value
    elif hasattr(element_id, "IntegerValue"): # Pre-2024
        return element_id.IntegerValue
    else:
        return int(element_id)

def is_pipe(element):
    """Checks if element is a Pipe or Placeholder Pipe."""
    if not element or not element.Category:
        return False
    try:
        cat_id_val = get_id_value(element.Category.Id)
    except Exception:
        return False
    return cat_id_val == int(DB.BuiltInCategory.OST_PipeCurves) or cat_id_val == int(DB.BuiltInCategory.OST_PlaceHolderPipes)

def get_connector_closest_to(element, point):
    """Returns the connector of the element closest to the given point."""
    closest_conn = None
    min_dist = float('inf')
    
    try:
        connectors = element.ConnectorManager.Connectors
    except AttributeError:
        try:
            connectors = element.MEPModel.ConnectorManager.Connectors
        except:
            return None

    for conn in connectors:
        dist = conn.Origin.DistanceTo(point)
        if dist < min_dist:
            min_dist = dist
            closest_conn = conn
    return closest_conn

def are_lines_parallel(line1, line2):
    return line1.Direction.CrossProduct(line2.Direction).IsZeroLength()

def is_vertical(v):
    return abs(v.Z) > 0.99

def get_pipe_diameter(pipe):
    """Returns the diameter of the pipe or a default small value."""
    try:
        param = pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
        if param:
            return param.AsDouble()
    except:
        pass
    return 0.1 # Default fallback

def get_intersection_unbound(line1, line2):
    """
    Finds intersection of two unbounded lines using Revit API.
    Returns (Point, parameter_on_line2) or None.
    """
    l1_u = line1.Clone()
    l1_u.MakeUnbound()
    
    l2_u = line2.Clone()
    l2_u.MakeUnbound()
    
    # Use clr.Reference to handle 'out' parameter explicitly
    results = clr.Reference[DB.IntersectionResultArray]()
    res_type = l1_u.Intersect(l2_u, results)
    
    if res_type == DB.SetComparisonResult.Overlap:
        # Dereference results.Value to get the actual array
        res_array = results.Value
        if res_array and not res_array.IsEmpty:
            int_res = res_array.get_Item(0)
            int_pt = int_res.XYZPoint
            
            # Helper to find u parameter on original line2
            # p = origin + u * dir
            # u = (p - origin) . dir
            u = (int_pt - line2.Origin).DotProduct(line2.Direction)
            
            return int_pt, u
            
    return None

def get_intersection_xy(line1, line2):
    """
    Finds intersection of two lines projected to XY plane.
    Returns (XYZ_on_plane) or None.
    XYZ_on_plane has Z=0.
    """
    p1 = line1.Origin
    v1 = line1.Direction
    v1_xy = DB.XYZ(v1.X, v1.Y, 0)
    
    p2 = line2.Origin
    v2 = line2.Direction
    v2_xy = DB.XYZ(v2.X, v2.Y, 0)
    
    # Handle Vertical Pipes (Zero length XY vector)
    if v1_xy.IsZeroLength() and v2_xy.IsZeroLength():
        return None # Both vertical, parallel/collinear (handled elsewhere)
        
    if v1_xy.IsZeroLength():
        # Line 1 is vertical (Point in XY)
        # Check if P1 is on Line 2
        p1_xy = DB.XYZ(p1.X, p1.Y, 0)
        v2_xy_norm = v2_xy.Normalize()
        
        # Distance from p1_xy to line2_xy
        # |(p1_xy - p2_xy) x v2_xy_norm|
        vec = p1_xy - DB.XYZ(p2.X, p2.Y, 0)
        cross_z = vec.CrossProduct(v2_xy_norm).Z
        
        if abs(cross_z) < 0.01:
            return p1_xy
        return None

    if v2_xy.IsZeroLength():
        # Line 2 is vertical
        p2_xy = DB.XYZ(p2.X, p2.Y, 0)
        v1_xy_norm = v1_xy.Normalize()
        
        vec = p2_xy - DB.XYZ(p1.X, p1.Y, 0)
        cross_z = vec.CrossProduct(v1_xy_norm).Z
        
        if abs(cross_z) < 0.01:
            return p2_xy
        return None

    # Normal case: Both have XY length
    l1_xy = DB.Line.CreateUnbound(DB.XYZ(p1.X, p1.Y, 0), v1_xy)
    l2_xy = DB.Line.CreateUnbound(DB.XYZ(p2.X, p2.Y, 0), v2_xy)
    
    results = clr.Reference[DB.IntersectionResultArray]()
    res_type = l1_xy.Intersect(l2_xy, results)
    
    if res_type == DB.SetComparisonResult.Overlap:
        res_array = results.Value
        if res_array and not res_array.IsEmpty:
            int_res = res_array.get_Item(0)
            return int_res.XYZPoint
            
    return None

def get_z_at_xy(line, xy_point):
    """
    Given a 3D line and a point (x,y,0), find the Z value on the line at that (x,y).
    """
    p = line.Origin
    v = line.Direction
    
    # Check if line is vertical
    if is_vertical(v):
        return p.Z # Return origin Z (or any Z really)

    if abs(v.X) > abs(v.Y):
        t = (xy_point.X - p.X) / v.X
    else:
        t = (xy_point.Y - p.Y) / v.Y
        
    return p.Z + t * v.Z

def get_closest_points_between_lines(line1, line2):
    """
    Finds the points on line1 and line2 that are closest to each other (Common Perpendicular).
    Returns (pt1, pt2).
    """
    p1 = line1.Origin
    v1 = line1.Direction
    p2 = line2.Origin
    v2 = line2.Direction
    
    a = v1.DotProduct(v1)
    b = v1.DotProduct(v2)
    c = v2.DotProduct(v2)
    
    dp = p2 - p1
    d = dp.DotProduct(v1)
    e = dp.DotProduct(v2)
    
    denom = a*c - b*b
    
    if abs(denom) < 0.0001:
        # Parallel lines
        # Pick p1 as anchor, project to line2
        t = 0
        # q2 = p2 + u * v2
        # (p2 + u*v2 - p1) . v2 = 0
        # dp . v2 + u * c = 0 => e + u = 0 (if c=1)
        u = -e / c
    else:
        # Skew lines
        u = (d*b - a*e) / denom
        t = (d + u*b) / a
        
    pt1 = p1 + t * v1
    pt2 = p2 + u * v2
    
    return pt1, pt2

# --- Main Logic ---

class PipeJoiner:
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.logger = Logger()
    
    def run(self):
        self.logger.info("Starting Join Pipes Command...")
        pair_count = 0
        while True:
            try:
                # 1. Selection
                sel1 = self.select_pipe("Select Pipe to Modify/Extend (ESC to Finish)")
                if not sel1: 
                    break
                pipe1_ref, p1_pick_pt = sel1
                pipe1 = self.doc.GetElement(pipe1_ref)
                
                sel2 = self.select_pipe("Select Pipe to Connect To (Reference) (ESC to Finish)")
                if not sel2: 
                    break
                pipe2_ref, _ = sel2
                pipe2 = self.doc.GetElement(pipe2_ref)

                if pipe1.Id == pipe2.Id:
                    self.logger.error("Selected the same pipe twice. Skipping pair.")
                    continue

                pair_count += 1
                self.logger.info("--- Processing Pair #{} ---".format(pair_count))
                
                # Nested try-except to prevent crash on single join failure
                try:
                    self.join_pipes(pipe1, pipe2, p1_pick_pt)
                    self.uidoc.RefreshActiveView()
                except Exception as op_err:
                    self.logger.error("Failed to join Pair #{}: {}".format(pair_count, op_err))
                    # Continue to next pair

            except Exception as e:
                self.logger.error("Critical Execution Error: {}".format(e))
                break
        
        self.logger.show()

    def select_pipe(self, prompt):
        # Using loop to ensure a valid pipe is selected or user cancels
        while True:
            try:
                ref = self.uidoc.Selection.PickObject(ObjectType.Element, prompt)
                element = self.doc.GetElement(ref)
                
                if is_pipe(element):
                    return ref, ref.GlobalPoint
                else:
                    print("Selection was not a pipe. Please select a Pipe.")
                    continue
                    
            except OperationCanceledException:
                return None
            except Exception as e:
                print("Selection Error: " + str(e))
                return None

    def join_pipes(self, p1, p2, p1_pick_pt):
        l1 = p1.Location.Curve
        l2 = p2.Location.Curve
        
        # Check parallel
        if are_lines_parallel(l1, l2):
            self.logger.info("Pipes are parallel.")
            # Check distance for collinearity
            dist = (l1.Origin - l2.Origin).CrossProduct(l2.Direction).GetLength()
            
            if dist < 0.01: # Tolerance
                self.logger.info("Pipes are collinear. Attempting Union.")
                self.create_union(p1, p2)
                return
            else:
                 pass # Offset parallel -> Closest Point

        # Check Intersection 3D
        res = get_intersection_unbound(l1, l2)
        if res:
            int_pt, u = res
            self.logger.info("3D Intersection found at {}.".format(int_pt))
            self.join_coplanar(p1, p2, int_pt, u, p1_pick_pt)
            return

        # Check Intersection 2D (Skew Plan Intersection)
        int_pt_xy = get_intersection_xy(l1, l2)
        if int_pt_xy:
            self.logger.info("2D (Plan) Intersection found at {} (Z=0). Checking vertical gap...".format(int_pt_xy))
            self.join_skew(p1, p2, int_pt_xy, p1_pick_pt)
            return
            
        # Fallback: Closest Approach
        self.logger.info("Pipes do not intersect in plan. Attempting Shortest Path (Common Perpendicular)...")
        if not self.join_by_closest_points(p1, p2, p1_pick_pt):
            # RETRY MECHANISM: Offset logic
            self.logger.info("Direct bridge failed (likely too short). Attempting OFFSET connection (Loop/Dogleg)...")
            
            # Shift the connection point on Pipe 2 (Dynamic: 3x Dia or min 6")
            dia = get_pipe_diameter(p1)
            offset_dist = max(0.5, 3.0 * dia) 
            l2_dir = p2.Location.Curve.Direction
            
            # Try shifting in positive direction
            if self.join_by_offset_target(p1, p2, p1_pick_pt, offset_dist):
                return
            
            # Try shifting in negative direction
            self.logger.info("Positive offset failed. Attempting Negative offset...")
            if self.join_by_offset_target(p1, p2, p1_pick_pt, -offset_dist):
                return
            
            self.logger.error("All connection strategies failed.")

    def connect_branch_to_main(self, c_branch, p_main, connect_pt):
        """
        Connects a branch connector to a main pipe at a specific point.
        Attempts Takeoff first, then Split+Tee.
        """
        try:
            # Try Tap/Takeoff
            self.doc.Create.NewTakeoffFitting(c_branch, p_main)
            self.logger.info("Created Takeoff Fitting.")
            return True
        except Exception as e:
            # Check if it's a routing preference issue
            msg = str(e)
            if "No routing preference" in msg or "takeoff" in msg.lower():
                self.logger.info("Takeoff failed ({}). Attempting Split and Tee...".format(msg))
                return self.split_and_tee(c_branch, p_main, connect_pt)
            else:
                self.logger.error("Connection failed: {}".format(msg))
                # raise e # Don't raise, let the caller handle graceful continuation
                return False

    def split_and_tee(self, c_branch, p_main, split_pt):
        # 1. Identify Main Pipe endpoints
        c_main = p_main.Location.Curve
        p0 = c_main.GetEndPoint(0)
        p1 = c_main.GetEndPoint(1)
        
        # Check valid split
        min_len = get_pipe_diameter(p_main) * 1.5
        if p0.DistanceTo(split_pt) < min_len or p1.DistanceTo(split_pt) < min_len:
            self.logger.error("Split point too close to pipe ends for Tee (< {}).".format(min_len))
            return False

        # 2. Store properties
        sys_id = p_main.MEPSystem.GetTypeId()
        type_id = p_main.GetTypeId()
        level_id = p_main.ReferenceLevel.Id
        
        try:
            # 3. Create New Segment (split_pt -> p1)
            # Validate length
            if split_pt.DistanceTo(p1) < min_len:
                 self.logger.error("Resulting split segment too short.")
                 return False

            p_new = DB.Plumbing.Pipe.Create(self.doc, sys_id, type_id, level_id, split_pt, p1)
            
            # Match Diameter/params
            try:
                dia_param = p_main.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                if dia_param:
                    p_new.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(dia_param.AsDouble())
            except:
                pass
                
            # 4. Shorten Original (p0 -> split_pt)
            new_curve = DB.Line.CreateBound(p0, split_pt)
            p_main.Location.Curve = new_curve
            
            # 5. Connect
            c_main_end = get_connector_closest_to(p_main, split_pt)
            c_new_start = get_connector_closest_to(p_new, split_pt)
            
            self.doc.Create.NewTeeFitting(c_main_end, c_new_start, c_branch)
            self.logger.info("Created Tee Fitting.")
            return True
            
        except Exception as e:
            self.logger.error("Split and Tee failed: {}".format(e))
            return False

    def extend_pipe_to_point(self, pipe, target_point, guide_point=None):
        """
        Extends/Trims 'pipe' so one end hits 'target_point'.
        If 'guide_point' is provided, it modifies the end closest to 'guide_point'.
        Otherwise, it modifies the end closest to 'target_point'.
        """
        lc = pipe.Location.Curve
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        
        # Decide which end to move
        if guide_point:
            # Move the end closest to where the user clicked (guide_point)
            dist0 = p0.DistanceTo(guide_point)
            dist1 = p1.DistanceTo(guide_point)
            move_p0 = dist0 < dist1
        else:
            # Move the end closest to the target
            dist0 = p0.DistanceTo(target_point)
            dist1 = p1.DistanceTo(target_point)
            move_p0 = dist0 < dist1
        
        # Check for zero length
        fixed_pt = p1 if move_p0 else p0
        new_len = target_point.DistanceTo(fixed_pt)
        min_len = get_pipe_diameter(pipe) * 0.5 # Minimal tolerance
        
        if new_len < min_len:
            self.logger.error("Extension would result in too-short pipe ({} < {}).".format(new_len, min_len))
            return False

        try:
            if move_p0:
                new_curve = DB.Line.CreateBound(target_point, p1)
            else:
                new_curve = DB.Line.CreateBound(p0, target_point)
            
            pipe.Location.Curve = new_curve
            return True
        except Exception as e:
            self.logger.error("Failed to extend pipe: {}".format(e))
            return False

    def join_coplanar(self, p1, p2, int_pt, u, p1_pick_pt):
        l2 = p2.Location.Curve
        dist_p2 = l2.Length
        is_on_segment = 0.001 < u < (dist_p2 - 0.001)
        
        t_transaction = DB.Transaction(self.doc, "Join Pipes Coplanar")
        t_transaction.Start()
        
        try:
            if is_on_segment:
                self.logger.info("Intersection is ON the reference pipe segment. Connecting Branch...")
                if not self.extend_pipe_to_point(p1, int_pt, p1_pick_pt):
                    t_transaction.RollBack()
                    return
                c1 = get_connector_closest_to(p1, int_pt)
                
                self.connect_branch_to_main(c1, p2, int_pt)
            else:
                self.logger.info("Intersection is OUTSIDE reference pipe segment. Creating Elbow.")
                if not self.extend_pipe_to_point(p1, int_pt, p1_pick_pt):
                    t_transaction.RollBack()
                    return
                if not self.extend_pipe_to_point(p2, int_pt):
                    t_transaction.RollBack()
                    return
                c1 = get_connector_closest_to(p1, int_pt)
                c2 = get_connector_closest_to(p2, int_pt)
                try:
                    self.doc.Create.NewElbowFitting(c1, c2)
                    self.logger.info("Created Elbow Fitting.")
                except Exception as e:
                    self.logger.error("Failed to create Elbow (Geometry created): {}".format(e))
                    
            t_transaction.Commit()
        except Exception as e:
            t_transaction.RollBack()
            self.logger.error("Transaction Failed: {}".format(e))

    def join_skew(self, p1, p2, int_pt_xy, p1_pick_pt):
        # Calculate Z values
        z1 = get_z_at_xy(p1.Location.Curve, int_pt_xy)
        z2 = get_z_at_xy(p2.Location.Curve, int_pt_xy)
        pt1 = DB.XYZ(int_pt_xy.X, int_pt_xy.Y, z1)
        pt2 = DB.XYZ(int_pt_xy.X, int_pt_xy.Y, z2)
        v1_is_vert = is_vertical(p1.Location.Curve.Direction)
        v2_is_vert = is_vertical(p2.Location.Curve.Direction)

        t_transaction = DB.Transaction(self.doc, "Join Pipes Skew")
        t_transaction.Start()
        
        try:
            if v1_is_vert:
                if not self.extend_pipe_to_point(p1, pt2, p1_pick_pt): 
                     t_transaction.RollBack(); return
                
                l2 = p2.Location.Curve
                u = (pt2 - l2.Origin).DotProduct(l2.Direction)
                is_on_segment = 0.001 < u < (l2.Length - 0.001)
                if not is_on_segment:
                     if not self.extend_pipe_to_point(p2, pt2):
                          t_transaction.RollBack(); return
                
                c1 = get_connector_closest_to(p1, pt2)
                if is_on_segment:
                     self.connect_branch_to_main(c1, p2, pt2)
                else:
                     c2 = get_connector_closest_to(p2, pt2)
                     try: self.doc.Create.NewElbowFitting(c1, c2)
                     except Exception as e: self.logger.error("Elbow failed: {}".format(e))

            elif v2_is_vert:
                 if not self.extend_pipe_to_point(p1, pt1, p1_pick_pt):
                    t_transaction.RollBack(); return
                 
                 # P2 is vertical ref. Assume it's long enough or user handles extension logic manually for ref?
                 # Extending ref pipe (P2) is tricky if we don't know which end.
                 # Let's assume we connect to P2 at pt1.
                 c1 = get_connector_closest_to(p1, pt1)
                 # Check if pt1 is on P2 segment
                 l2_start_z = p2.Location.Curve.GetEndPoint(0).Z
                 l2_end_z = p2.Location.Curve.GetEndPoint(1).Z
                 min_z = min(l2_start_z, l2_end_z)
                 max_z = max(l2_start_z, l2_end_z)
                 is_on_p2_segment = (min_z + 0.001) < pt1.Z < (max_z - 0.001)
                 
                 if is_on_p2_segment:
                     self.connect_branch_to_main(c1, p2, pt1)
                 else:
                     # Extend P2 to pt1?
                     if not self.extend_pipe_to_point(p2, pt1):
                         t_transaction.RollBack(); return
                     c2 = get_connector_closest_to(p2, pt1)
                     try: self.doc.Create.NewElbowFitting(c1, c2)
                     except Exception as e: self.logger.error("Elbow failed: {}".format(e))

            else:
                 # Neither vertical (Riser case)
                 z_diff = abs(z1 - z2)
                 if z_diff < 0.01:
                      u = (pt2 - p2.Location.Curve.Origin).DotProduct(p2.Location.Curve.Direction)
                      self.join_coplanar(p1, p2, pt2, u, p1_pick_pt)
                      t_transaction.Commit()
                      return

                 self.logger.info("Creating Riser connection.")
                 if not self.extend_pipe_to_point(p1, pt1, p1_pick_pt):
                    t_transaction.RollBack(); return
                 
                 l2 = p2.Location.Curve
                 u = (pt2 - l2.Origin).DotProduct(l2.Direction)
                 is_on_segment = 0.001 < u < (l2.Length - 0.001)
                 
                 if not is_on_segment:
                    if not self.extend_pipe_to_point(p2, pt2):
                        t_transaction.RollBack(); return
                 
                 # Create Riser
                 riser = self.create_pipe_segment(p1, pt1, pt2)
                 if not riser:
                     t_transaction.RollBack(); return
                 self.logger.info("Created Vertical Riser.")

                 # Connect
                 c1 = get_connector_closest_to(p1, pt1)
                 c_riser_1 = get_connector_closest_to(riser, pt1)
                 try: self.doc.Create.NewElbowFitting(c1, c_riser_1)
                 except Exception as e: self.logger.error("Elbow 1 failed: {}".format(e))
                 
                 c_riser_2 = get_connector_closest_to(riser, pt2)
                 if is_on_segment:
                     self.connect_branch_to_main(c_riser_2, p2, pt2)
                 else:
                     c2 = get_connector_closest_to(p2, pt2)
                     try: self.doc.Create.NewElbowFitting(c_riser_2, c2)
                     except Exception as e: self.logger.error("Elbow 2 failed: {}".format(e))

            t_transaction.Commit()
        except Exception as e:
            t_transaction.RollBack()
            self.logger.error("Skew Join Failed: {}".format(e))

    def join_by_closest_points(self, p1, p2, p1_pick_pt):
        """Attempts to join by shortest path. Returns True if successful, False otherwise."""
        l1 = p1.Location.Curve
        l2 = p2.Location.Curve
        pt1, pt2 = get_closest_points_between_lines(l1, l2)
        
        # Validation: Check bridge length
        bridge_len = pt1.DistanceTo(pt2)
        min_len = get_pipe_diameter(p1) * 1.5
        if bridge_len < min_len:
            self.logger.error("Bridge segment too short ({} < {}).".format(bridge_len, min_len))
            return False

        t_transaction = DB.Transaction(self.doc, "Join Pipes Closest Path")
        t_transaction.Start()
        
        try:
            # 1. Extend P1
            if not self.extend_pipe_to_point(p1, pt1, p1_pick_pt):
                t_transaction.RollBack(); return False
            
            # 2. Extend/Check P2
            l2_curr = p2.Location.Curve
            u = (pt2 - l2_curr.Origin).DotProduct(l2_curr.Direction)
            is_on_segment = 0.001 < u < (l2_curr.Length - 0.001)
            
            if not is_on_segment:
                if not self.extend_pipe_to_point(p2, pt2):
                    t_transaction.RollBack(); return False
            
            # 3. Create Bridge
            bridge = self.create_pipe_segment(p1, pt1, pt2)
            if not bridge:
                t_transaction.RollBack(); return False
            self.logger.info("Created Bridge Pipe (90 deg).")
            
            # 4. Connect
            c1 = get_connector_closest_to(p1, pt1)
            c_bridge_1 = get_connector_closest_to(bridge, pt1)
            try: 
                self.doc.Create.NewElbowFitting(c1, c_bridge_1)
            except Exception as e: 
                self.logger.error("Elbow (P1->Bridge) failed: {}".format(e))
                # Don't fail entire op, maybe just geometry created is fine
            
            c_bridge_2 = get_connector_closest_to(bridge, pt2)
            if is_on_segment:
                 self.connect_branch_to_main(c_bridge_2, p2, pt2)
            else:
                 c2 = get_connector_closest_to(p2, pt2)
                 try: 
                    self.doc.Create.NewElbowFitting(c_bridge_2, c2)
                 except Exception as e: 
                    self.logger.error("Elbow (Bridge->P2) failed: {}".format(e))
                    
            t_transaction.Commit()
            return True

        except Exception as e:
            t_transaction.RollBack()
            self.logger.error("Closest Path Join Failed: {}".format(e))
            return False

    def join_by_offset_target(self, p1, p2, p1_pick_pt, offset_dist):
        """
        Attempts to join P1 to a point shifted along P2 by offset_dist.
        This creates a 'Dogleg' or skewed connection.
        """
        l2 = p2.Location.Curve
        pt1_orig, pt2_orig = get_closest_points_between_lines(p1.Location.Curve, l2)
        
        # Calculate new target point on Pipe 2
        shift_vec = l2.Direction * offset_dist
        pt2_new = pt2_orig + shift_vec
        
        # Now we treat this as a "Skew Join" to this specific point
        # But wait, skew join logic calculates vertical risers based on XY intersection.
        # Here we just want to bridge P1_end to pt2_new.
        
        # 1. Find point on P1 closest to pt2_new? 
        # No, we want to extend P1 to be "opposite" pt2_new?
        # Actually, let's try 'join_by_closest_points' logic but FORCING the target point on P2.
        
        t_transaction = DB.Transaction(self.doc, "Join Pipes Offset")
        t_transaction.Start()
        
        try:
            # We need a point on P1. Let's project pt2_new onto P1's line
            l1 = p1.Location.Curve
            # Project pt2_new onto l1
            # p_proj = origin + (vec . dir) * dir
            vec = pt2_new - l1.Origin
            proj_dist = vec.DotProduct(l1.Direction)
            pt1_new = l1.Origin + l1.Direction * proj_dist
            
            # Check resulting bridge length
            bridge_len = pt1_new.DistanceTo(pt2_new)
            min_len = get_pipe_diameter(p1) * 1.5
            if bridge_len < min_len:
                self.logger.error("Offset Bridge too short ({} < {}).".format(bridge_len, min_len))
                t_transaction.RollBack(); return False

            # 1. Extend P1
            if not self.extend_pipe_to_point(p1, pt1_new, p1_pick_pt):
                t_transaction.RollBack(); return False
            
            # 2. Extend/Check P2
            # Check if pt2_new is on segment
            u = (pt2_new - l2.Origin).DotProduct(l2.Direction)
            is_on_segment = 0.001 < u < (l2.Length - 0.001)
            
            if not is_on_segment:
                if not self.extend_pipe_to_point(p2, pt2_new):
                    t_transaction.RollBack(); return False
            
            # 3. Create Bridge
            bridge = self.create_pipe_segment(p1, pt1_new, pt2_new)
            if not bridge:
                t_transaction.RollBack(); return False
            self.logger.info("Created Offset Bridge (Offset {}').".format(offset_dist))

            # 4. Connect
            c1 = get_connector_closest_to(p1, pt1_new)
            c_bridge_1 = get_connector_closest_to(bridge, pt1_new)
            try: self.doc.Create.NewElbowFitting(c1, c_bridge_1)
            except: pass 
            
            c_bridge_2 = get_connector_closest_to(bridge, pt2_new)
            if is_on_segment:
                 self.connect_branch_to_main(c_bridge_2, p2, pt2_new)
            else:
                 c2 = get_connector_closest_to(p2, pt2_new)
                 try: self.doc.Create.NewElbowFitting(c_bridge_2, c2)
                 except: pass
            
            t_transaction.Commit()
            return True
            
        except Exception as e:
            t_transaction.RollBack()
            self.logger.error("Offset Join Failed: {}".format(e))
            return False

    def create_pipe_segment(self, template_pipe, start, end):
        try:
            # Validate length
            min_len = get_pipe_diameter(template_pipe) * 1.5
            if start.DistanceTo(end) < min_len:
                 self.logger.error("New pipe segment too short.")
                 return None

            sys_id = template_pipe.MEPSystem.GetTypeId()
            type_id = template_pipe.GetTypeId()
            level_id = template_pipe.ReferenceLevel.Id
            
            new_pipe = DB.Plumbing.Pipe.Create(self.doc, sys_id, type_id, level_id, start, end)
            
            try:
                dia_param = template_pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                if dia_param:
                    new_pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(dia_param.AsDouble())
            except:
                pass
            return new_pipe
        except Exception as e:
            self.logger.error("Failed to create pipe segment: {}".format(e))
            return None

    def create_union(self, p1, p2):
        c1_closest = None
        c2_closest = None
        min_dist = float('inf')
        
        # Find closest pair of connectors
        for c1 in p1.ConnectorManager.Connectors:
            for c2 in p2.ConnectorManager.Connectors:
                d = c1.Origin.DistanceTo(c2.Origin)
                if d < min_dist:
                    min_dist = d
                    c1_closest = c1
                    c2_closest = c2
        
        t_transaction = DB.Transaction(self.doc, "Union Pipes")
        t_transaction.Start()
        try:
            # Move p1 to meet p2
            move_vec = c2_closest.Origin - c1_closest.Origin
            DB.ElementTransformUtils.MoveElement(self.doc, p1.Id, move_vec)
            
            # Need to regenerate to update connector locations?
            self.doc.Regenerate()
            
            self.doc.Create.NewUnionFitting(c1_closest, c2_closest)
            self.logger.info("Created Union Fitting.")
            t_transaction.Commit()
        except Exception as e:
            t_transaction.RollBack()
            self.logger.error("Union Failed: {}".format(e))

# --- Entry Point ---

if __name__ == '__main__':
    doc = revit.doc
    uidoc = revit.uidoc
    joiner = PipeJoiner(doc, uidoc)
    joiner.run()