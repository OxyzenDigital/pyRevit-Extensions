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
import System
from System.Collections.Generic import List

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

def is_point_occupied(pipe, point):
    """Checks if the connector closest to 'point' on 'pipe' is already connected."""
    conn = get_connector_closest_to(pipe, point)
    if conn and conn.IsConnected:
        return True
    return False

def get_intersector(doc, view3d, exclude_ids=None):
    """
    Creates a ReferenceIntersector targeting obstacles (Structure, MEP) in the 3D view.
    exclude_ids: List of ElementIds to ignore (e.g., the pipes being modified).
    """
    cats = [
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_StructuralColumns,
        DB.BuiltInCategory.OST_DuctCurves,
        DB.BuiltInCategory.OST_CableTray,
        DB.BuiltInCategory.OST_Conduit,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_Walls,
        DB.BuiltInCategory.OST_Floors,
        DB.BuiltInCategory.OST_Ceilings
    ]
    
    # Collect all elements of these categories in the view
    collector = DB.FilteredElementCollector(doc, view3d.Id)
    
    # Use System.Collections.Generic.List
    cat_list = List[DB.BuiltInCategory](cats)
    filter_cats = DB.ElementMulticategoryFilter(cat_list)
    collector.WherePasses(filter_cats)
    
    # Exclude specific IDs
    if exclude_ids:
        exclude_ids_coll = List[DB.ElementId](exclude_ids)
        collector.Excluding(exclude_ids_coll)
        
    target_ids = collector.ToElementIds()
    
    if not target_ids or target_ids.Count == 0:
        return None

    intersector = DB.ReferenceIntersector(target_ids, DB.FindReferenceTarget.Element, view3d)
    intersector.FindReferencesInRevitLinks = False 
    return intersector

def check_clearance(intersector, start_pt, end_pt, radius):
    """
    Checks for collisions along the path from start_pt to end_pt.
    Casts 5 rays: Center + 4 Perimeter rays (Up, Down, Left, Right).
    Returns True if collision detected.
    """
    if not intersector:
        return False
        
    direction = (end_pt - start_pt).Normalize()
    dist = start_pt.DistanceTo(end_pt)
    
    # Basis vectors for offset
    if is_vertical(direction):
        vec_u = DB.XYZ.BasisX
        vec_v = DB.XYZ.BasisY
    else:
        vec_u = DB.XYZ.BasisZ
        vec_v = direction.CrossProduct(vec_u).Normalize()
        
    # Offset points to check (Center, Up, Down, Left, Right)
    # Adding a small tolerance buffer to radius
    check_radius = radius * 1.1 
    offsets = [
        DB.XYZ.Zero,
        vec_u * check_radius,
        vec_u * -check_radius,
        vec_v * check_radius,
        vec_v * -check_radius
    ]
    
    for off in offsets:
        s_p = start_pt + off
        # Offset start slightly to avoid self-intersection
        s_p_offset = s_p + direction * 0.1
        
        context = intersector.FindNearest(s_p_offset, direction)
        if context:
            hit_dist = context.Proximity
            # Check if hit is within segment length
            if hit_dist < (dist - 0.2):
                return True
                
    return False

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

def get_closest_points_between_lines(line1, line2, ref_point=None):
    """
    Finds the points on line1 and line2 that are closest to each other (Common Perpendicular).
    Returns (pt1, pt2).
    If lines are parallel and ref_point is provided, finds closest point on line1 to ref_point, 
    then projects to line2.
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
        if ref_point:
            # Project ref_point onto line1 to get t
            # p(t) = p1 + t*v1
            # (ref_point - p1) . v1 = t * (v1 . v1)
            # t = (ref_point - p1) . v1 / a (assuming normalized v1, a=1)
            t = (ref_point - p1).DotProduct(v1) / a
        else:
            t = 0
            
        # q2 = p2 + u * v2
        # (p2 + u*v2 - (p1 + t*v1)) . v2 = 0
        # (p2 - p1).v2 + u*c - t*v1.v2 = 0
        # e + u*c - t*b = 0
        u = (t*b - e) / c
    else:
        # Skew lines
        u = (d*b - a*e) / denom
        t = (d + u*b) / a
        
    pt1 = p1 + t * v1
    pt2 = p2 + u * v2
    
    return pt1, pt2

def connect_connectors_robust(doc, c1, c2, logger):
    """
    Attempts to connect two connectors using Union or Elbow.
    Returns True if successful, False if failed (but suppresses exception).
    """
    try:
        dist = c1.Origin.DistanceTo(c2.Origin)
        if dist > 0.01:
            logger.error("Connectors too far apart ({})".format(dist))
            return False

        angle = c1.CoordinateSystem.BasisZ.AngleTo(c2.CoordinateSystem.BasisZ) * 180 / math.pi
        logger.info("Connection Angle: {:.2f}".format(angle))

        if abs(angle - 180) < 5.0:
            logger.info("Creating Union...")
            doc.Create.NewUnionFitting(c1, c2)
        else:
            logger.info("Creating Elbow...")
            doc.Create.NewElbowFitting(c1, c2)
        return True
    except Exception as e:
        logger.error("Fitting creation failed: {}".format(e))
        # Check Routing Prefs hint
        try:
            pipe_type = c1.Owner.PipeType
            # accessing routing prefs is complex in python, just hint
            logger.error("Check Routing Preferences for Pipe Type '{}'.".format(pipe_type.Name))
        except:
            pass
        return False

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
            self.logger.error("Extension would result in too-short pipe ({} < {}). Target: {} Fixed: {}".format(new_len, min_len, target_point, fixed_pt))
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
        pt1, pt2 = get_closest_points_between_lines(l1, l2, p1_pick_pt)
        
        self.logger.info("Closest Points - PT1: {} PT2: {}".format(pt1, pt2))
        
        dia = get_pipe_diameter(p1)
        radius = dia / 2.0
        
        # Strategy 1: Direct Bridge
        # Check if target point on Pipe 2 is occupied
        occupied = is_point_occupied(p2, pt2)
        
        # Check bridge length
        bridge_len = pt1.DistanceTo(pt2)
        min_len = dia * 2.0 
        too_short = bridge_len < min_len
        
        # Check Collision
        collision = False
        active_view = self.doc.ActiveView
        if isinstance(active_view, DB.View3D):
            intersector = get_intersector(self.doc, active_view, exclude_ids=[p1.Id, p2.Id])
            if check_clearance(intersector, pt1, pt2, radius):
                collision = True

        if occupied or too_short or collision:
            reason = []
            if occupied: reason.append("Target Occupied")
            if too_short: reason.append("Too Short")
            if collision: reason.append("Collision")
            self.logger.info("Direct Bridge failed ({}). Attempting Slide Bypass...".format(", ".join(reason)))
            
            if self.join_with_slide_bypass(p1, p2, p1_pick_pt, pt1, pt2):
                return True
                
            self.logger.info("Slide Bypass failed. Attempting Goal Post Bypass (XZ/YZ/XY)...")
            return self.join_with_goalpost_bypass(p1, p2, p1_pick_pt, pt1, pt2)

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
            self.logger.info("Created Direct Bridge Pipe.")
            
            self.doc.Regenerate() # CRITICAL
            
            # 4. Connect
            c1 = get_connector_closest_to(p1, pt1)
            c_bridge_1 = get_connector_closest_to(bridge, pt1)
            
            if not connect_connectors_robust(self.doc, c1, c_bridge_1, self.logger):
                self.logger.error("Warning: Could not connect Pipe 1 to Bridge. Geometry preserved.")
            
            c_bridge_2 = get_connector_closest_to(bridge, pt2)
            if is_on_segment:
                 self.connect_branch_to_main(c_bridge_2, p2, pt2)
            else:
                 c2 = get_connector_closest_to(p2, pt2)
                 if not connect_connectors_robust(self.doc, c_bridge_2, c2, self.logger):
                     self.logger.error("Warning: Could not connect Bridge to Pipe 2. Geometry preserved.")
                    
            t_transaction.Commit()
            return True

        except Exception as e:
            t_transaction.RollBack()
            self.logger.error("Closest Path Join Failed: {}".format(e))
            return False

    def join_with_slide_bypass(self, p1, p2, p1_pick_pt, pt1_orig, pt2_orig):
        """
        Strategy A: Slide/Dogleg along the Reference Pipe axis (Z-shape).
        """
        active_view = self.doc.ActiveView
        intersector = None
        if isinstance(active_view, DB.View3D):
            intersector = get_intersector(self.doc, active_view, exclude_ids=[p1.Id, p2.Id])
            
        dia = get_pipe_diameter(p1)
        radius = dia / 2.0
        base_offset = max(0.5, 4.0 * dia)
        l2_dir = p2.Location.Curve.Direction
        
        max_attempts = 10
        
        for i in range(1, max_attempts + 1):
            current_offset = base_offset * i
            
            for direction in [1.0, -1.0]:
                shift_vec = l2_dir * (current_offset * direction)
                
                pt1_new = pt1_orig + shift_vec
                pt2_new = pt2_orig + shift_vec
                
                # Check Collisions
                if check_clearance(intersector, pt1_orig, pt1_new, radius): continue
                if check_clearance(intersector, pt1_new, pt2_new, radius): continue

                self.logger.info("Found clear Slide path at Offset {}'.".format(current_offset * direction))
                
                t_transaction = DB.Transaction(self.doc, "Join Pipes Slide Bypass")
                t_transaction.Start()
                
                try:
                    if not self.extend_pipe_to_point(p1, pt1_orig, p1_pick_pt):
                        t_transaction.RollBack(); continue

                    l2_curr = p2.Location.Curve
                    u = (pt2_new - l2_curr.Origin).DotProduct(l2_curr.Direction)
                    is_on_segment = 0.001 < u < (l2_curr.Length - 0.001)
                    if not is_on_segment:
                        if not self.extend_pipe_to_point(p2, pt2_new):
                             t_transaction.RollBack(); continue
                    
                    offset_pipe = self.create_pipe_segment(p1, pt1_orig, pt1_new)
                    bridge_pipe = self.create_pipe_segment(p1, pt1_new, pt2_new)
                    if not offset_pipe or not bridge_pipe: 
                        t_transaction.RollBack(); continue

                    self.logger.info("Created Slide Geometry.")
                    self.doc.Regenerate() # CRITICAL
                    
                    # Connect
                    c_p1 = get_connector_closest_to(p1, pt1_orig)
                    c_off_1 = get_connector_closest_to(offset_pipe, pt1_orig)
                    connect_connectors_robust(self.doc, c_p1, c_off_1, self.logger)
                    
                    c_off_2 = get_connector_closest_to(offset_pipe, pt1_new)
                    c_bridge_1 = get_connector_closest_to(bridge_pipe, pt1_new)
                    connect_connectors_robust(self.doc, c_off_2, c_bridge_1, self.logger)
                    
                    c_bridge_2 = get_connector_closest_to(bridge_pipe, pt2_new)
                    if is_on_segment:
                        if not self.connect_branch_to_main(c_bridge_2, p2, pt2_new):
                            self.logger.error("Warning: Branch connection failed. Geometry preserved.")
                    else:
                        c_p2 = get_connector_closest_to(p2, pt2_new)
                        connect_connectors_robust(self.doc, c_bridge_2, c_p2, self.logger)
                    
                    t_transaction.Commit()
                    return True
                    
                except Exception as e:
                    t_transaction.RollBack()
                    self.logger.error("Slide Bypass attempt failed: {}".format(e))
                    continue 
            
        return False

    def join_with_goalpost_bypass(self, p1, p2, p1_pick_pt, pt1_orig, pt2_orig):
        """
        Strategy B: Goal Post Bypass (U-shape) to jump over/around obstacles.
        Tries Vertical (Up/Down) and Horizontal (Left/Right) jumps.
        """
        active_view = self.doc.ActiveView
        intersector = None
        if isinstance(active_view, DB.View3D):
            intersector = get_intersector(self.doc, active_view, exclude_ids=[p1.Id, p2.Id])
            
        dia = get_pipe_diameter(p1)
        radius = dia / 2.0
        base_offset = max(0.5, 4.0 * dia)
        l2_dir = p2.Location.Curve.Direction
        
        # Determine Jump Vectors
        jump_dirs = []
        # 1. Vertical (Z)
        if not is_vertical(l2_dir):
            jump_dirs.append(("Up", DB.XYZ.BasisZ))
            jump_dirs.append(("Down", -DB.XYZ.BasisZ))
            
        # 2. Horizontal (Cross Product)
        # Vector perpendicular to Pipe 2 and Up
        if is_vertical(l2_dir):
            side_vec = DB.XYZ.BasisX # Arbitrary for vertical pipe
        else:
            side_vec = l2_dir.CrossProduct(DB.XYZ.BasisZ).Normalize()
            
        jump_dirs.append(("Side A", side_vec))
        jump_dirs.append(("Side B", -side_vec))
        
        for i in range(1, 6): # Try 5 increments
            current_offset = base_offset * i
            
            for name, vec in jump_dirs:
                jump_vec = vec * current_offset
                
                pt1_jump = pt1_orig + jump_vec
                pt2_jump = pt2_orig + jump_vec
                
                # Check Collisions (3 segments)
                # 1. Riser 1 (pt1 -> pt1_jump)
                if check_clearance(intersector, pt1_orig, pt1_jump, radius): continue
                # 2. Bridge (pt1_jump -> pt2_jump)
                if check_clearance(intersector, pt1_jump, pt2_jump, radius): continue
                # 3. Riser 2 (pt2_jump -> pt2_orig)
                if check_clearance(intersector, pt2_jump, pt2_orig, radius): continue

                self.logger.info("Found clear Goal Post path ({}) at Offset {}'.".format(name, current_offset))
                
                t_transaction = DB.Transaction(self.doc, "Join Pipes GoalPost")
                t_transaction.Start()
                
                try:
                    if not self.extend_pipe_to_point(p1, pt1_orig, p1_pick_pt):
                        t_transaction.RollBack(); continue

                    # Check P2
                    l2_curr = p2.Location.Curve
                    u = (pt2_orig - l2_curr.Origin).DotProduct(l2_curr.Direction)
                    is_on_segment = 0.001 < u < (l2_curr.Length - 0.001)
                    if not is_on_segment:
                        if not self.extend_pipe_to_point(p2, pt2_orig):
                             t_transaction.RollBack(); continue
                    
                    riser1 = self.create_pipe_segment(p1, pt1_orig, pt1_jump)
                    bridge = self.create_pipe_segment(p1, pt1_jump, pt2_jump)
                    riser2 = self.create_pipe_segment(p1, pt2_jump, pt2_orig)
                    
                    if not riser1 or not bridge or not riser2:
                         t_transaction.RollBack(); continue

                    self.logger.info("Created Goal Post Geometry.")
                    self.doc.Regenerate() # CRITICAL
                    
                    # Connect Riser 1
                    c_p1 = get_connector_closest_to(p1, pt1_orig)
                    c_r1_start = get_connector_closest_to(riser1, pt1_orig)
                    connect_connectors_robust(self.doc, c_p1, c_r1_start, self.logger)
                    
                    # Connect Riser 1 to Bridge
                    c_r1_end = get_connector_closest_to(riser1, pt1_jump)
                    c_b_start = get_connector_closest_to(bridge, pt1_jump)
                    connect_connectors_robust(self.doc, c_r1_end, c_b_start, self.logger)
                    
                    # Connect Bridge to Riser 2
                    c_b_end = get_connector_closest_to(bridge, pt2_jump)
                    c_r2_start = get_connector_closest_to(riser2, pt2_jump)
                    connect_connectors_robust(self.doc, c_b_end, c_r2_start, self.logger)
                    
                    # Connect Riser 2 to P2
                    c_r2_end = get_connector_closest_to(riser2, pt2_orig)
                    if is_on_segment:
                        if not self.connect_branch_to_main(c_r2_end, p2, pt2_orig):
                             self.logger.error("Warning: Branch connection failed. Geometry preserved.")
                    else:
                        c_p2 = get_connector_closest_to(p2, pt2_orig)
                        connect_connectors_robust(self.doc, c_r2_end, c_p2, self.logger)
                    
                    t_transaction.Commit()
                    return True
                    
                except Exception as e:
                    t_transaction.RollBack()
                    self.logger.error("Goal Post attempt failed: {}".format(e))
                    continue
                    
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