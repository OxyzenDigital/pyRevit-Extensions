# -*- coding: utf-8 -*-
import sys
import os
import json
import math
import clr
import traceback

# --- ASSEMBLIES ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

# --- IMPORTS ---
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    XYZ, Transaction, TransactionGroup, ElementId, BuiltInParameter,
    ReferenceIntersector, FindReferenceTarget, Options, Solid, ViewType, Edge,
    ElementTransformUtils, FamilyInstance, CurveElement, UnitUtils, SpecTypeId
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB.ExtensibleStorage import SchemaBuilder, Schema, Entity, FieldBuilder, AccessLevel
from System import Guid
from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# 0. UNIT HELPER
# ==========================================
class UnitHelper:
    @staticmethod
    def get_project_length_unit():
        # returns ForgeTypeId (UnitTypeId)
        return doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()

    @staticmethod
    def get_unit_symbol():
        try:
            # Try to get symbol from FormatOptions
            opts = doc.GetUnits().GetFormatOptions(SpecTypeId.Length)
            if opts.UseDefault:
                # Need to look up default symbol for this unit, hard to do easily in API without label utils
                # Fallback to simple mapping or labelutils
                pass
            
            # Simple fallback based on TypeId string for common cases
            tid = UnitHelper.get_project_length_unit().TypeId
            if "meters" in tid: return "m"
            if "centimeters" in tid: return "cm"
            if "millimeters" in tid: return "mm"
            if "feet" in tid: return "ft"
            if "inches" in tid: return "in"
        except: pass
        return "units"

    @staticmethod
    def to_internal(value_in_project_units):
        try:
            val = float(value_in_project_units)
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertToInternalUnits(val, unit_id)
        except: return 0.0

    @staticmethod
    def from_internal(value_in_internal_units):
        try:
            val = float(value_in_internal_units)
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertFromInternalUnits(val, unit_id)
        except: return 0.0

# ==========================================
# 1. SETTINGS & LOGGING
# ==========================================
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "grading_settings.json")
RECIPE_SCHEMA_GUID = Guid("A4B9C8D2-1234-4567-8901-ABCDEF123456")

# Tolerances
MIN_DIST_TOLERANCE = 0.25 
BOUNDARY_TOLERANCE = 0.1 

class BatchLogger(object):
    """Accumulates messages to display in a single dialog."""
    def __init__(self):
        self._errors = []
        self._infos = []
    
    def error(self, msg, detail=None):
        self._errors.append(str(msg))
        if detail:
            self._errors.append("Details: " + str(detail))
    
    def info(self, msg):
        self._infos.append(str(msg))

    def show(self, title="Grading Report"):
        if not self._errors and not self._infos:
            return

        out = script.get_output()
        if self._errors:
            out.print_html('<strong>--- ERRORS ---</strong>')
            for e in self._errors:
                out.print_html('<div style="color:red;">{}</div>'.format(e))
            out.print_html('<br>')
        
        if self._infos:
            out.print_html('<strong>--- INFO ---</strong>')
            for i in self._infos:
                out.print_html('<div style="color:gray;">{}</div>'.format(i))

def get_id_val(element):
    if not element: return -1
    try: return element.Id.Value
    except AttributeError: return element.Id.IntegerValue

class GradingRecipe:
    @staticmethod
    def get_schema():
        schema = Schema.Lookup(RECIPE_SCHEMA_GUID)
        if not schema:
            builder = SchemaBuilder(RECIPE_SCHEMA_GUID)
            builder.SetReadAccessLevel(AccessLevel.Public)
            builder.SetWriteAccessLevel(AccessLevel.Public)
            builder.SetSchemaName("OxyzenGradingRecipe")
            builder.AddSimpleField("JsonData", str) 
            schema = builder.Finish()
        return schema

    @staticmethod
    def save_recipe(element, data_dict):
        if not element: return
        try:
            schema = GradingRecipe.get_schema()
            entity = Entity(schema)
            entity.Set("JsonData", json.dumps(data_dict))
            element.SetEntity(entity)
        except: pass

    @staticmethod
    def read_recipe(element):
        if not element: return None
        try:
            schema = GradingRecipe.get_schema()
            entity = element.GetEntity(schema)
            if entity.IsValid():
                return json.loads(entity.Get("JsonData", str))
        except: pass
        return None

def load_settings_from_disk():
    defaults = {
        "width": "6.0", "falloff": "10.0", "grid": "3.0", "slope": "2.0", "mode": "stakes"
    }
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                data = json.load(f)
                if data: defaults.update(data)
    except: pass
    return defaults

def save_state_to_disk(state):
    data = {
        "width": state.width, 
        "falloff": state.falloff, 
        "grid": state.grid, 
        "slope": state.slope_val, 
        "mode": state.mode
    }
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(data, f)
    except: pass

# ==========================================
# 2. STATE & UI
# ==========================================
class GradingState(object):
    def __init__(self):
        sets = load_settings_from_disk()
        
        self.width = sets.get("width", "6.0")
        self.falloff = sets.get("falloff", "10.0")
        self.grid = sets.get("grid", "3.0")
        self.slope_val = sets.get("slope", "2.0")
        self.mode = sets.get("mode", "stakes")
        
        self.start_stake = None
        self.end_stake = None
        self.grading_line = None
        
        self.next_action = None

    @property
    def ready(self):
        # Basic ready check: must have start stake and line.
        # End stake is needed if mode is NOT slope.
        has_start = self.start_stake is not None
        has_line = self.grading_line is not None
        has_end = self.end_stake is not None
        
        if self.mode == "slope":
            return has_start and has_line
        else:
            return has_start and has_end and has_line

class GradingWindow(forms.WPFWindow):
    def __init__(self, state):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.state = state
        self.bind_ui()
        self.setup_events()

    def bind_ui(self):
        from System.Windows import Media
        green = Media.Brushes.Green; red = Media.Brushes.Red
        def fmt(e): return "{} [{}]".format(e.Name, get_id_val(e)) if e else "[None]"

        if self.state.mode == "slope":
            self.Rb_UseSlope.IsChecked = True; self.Tb_Slope.IsEnabled = True
        else:
            self.Rb_MatchStakes.IsChecked = True; self.Tb_Slope.IsEnabled = False

        self.Lb_StartStake.Text = "Start: {}".format(fmt(self.state.start_stake))
        self.Lb_StartStake.Foreground = green if self.state.start_stake else red
        
        self.Lb_EndStake.Text = "End: {}".format(fmt(self.state.end_stake))
        self.Lb_EndStake.Foreground = green if self.state.end_stake else red
        
        self.Lb_Line.Text = "Line: {}".format(fmt(self.state.grading_line))
        self.Lb_Line.Foreground = green if self.state.grading_line else red

        # UNITS: Internal -> Display
        u_sym = UnitHelper.get_unit_symbol()
        
        def disp(val_str):
            try: return "{:.3f}".format(UnitHelper.from_internal(float(val_str)))
            except: return val_str

        self.Tb_Width.Text = disp(self.state.width)
        self.Tb_Falloff.Text = disp(self.state.falloff)
        self.Tb_Grid.Text = disp(self.state.grid)
        self.Tb_Slope.Text = str(self.state.slope_val)

        # Append unit to Title to inform user
        self.Title += " [{}]".format(u_sym)

        self.Btn_Run.IsEnabled = self.state.ready
        self.Btn_Edging.IsEnabled = self.state.ready
        self.Btn_Swap.IsEnabled = bool(self.state.start_stake and self.state.end_stake)

        if self.state.ready:
            self.Lb_Status.Content = "Ready."
            self.Lb_Status.Foreground = green
        else:
            self.Lb_Status.Content = "Incomplete."
            self.Lb_Status.Foreground = Media.Brushes.Gray

    def setup_events(self):
        self.Btn_SelectStakes.Click += self.a_stakes
        self.Btn_SelectLine.Click += self.a_line
        self.Btn_Swap.Click += self.a_swap
        self.Btn_Run.Click += self.a_run
        self.Btn_Edging.Click += self.a_edge
        self.Btn_Stitch.Click += self.a_stitch
        self.Btn_ReadRecipe.Click += self.a_load
        self.Rb_MatchStakes.Checked += self.mode_changed
        self.Rb_UseSlope.Checked += self.mode_changed
        
        self.Btn_SelectStakes.MouseEnter += self.h_stakes_on
        self.Btn_SelectStakes.MouseLeave += self.h_off
        self.Btn_SelectLine.MouseEnter += self.h_line_on
        self.Btn_SelectLine.MouseLeave += self.h_off

    def set_selection(self, elements):
        try:
            ids = List[ElementId]()
            for e in elements:
                if e and e.IsValidObject: ids.Add(e.Id)
            if ids.Count > 0:
                uidoc.Selection.SetElementIds(ids)
                uidoc.RefreshActiveView()
        except: pass

    def h_stakes_on(self, s, a): self.set_selection([self.state.start_stake, self.state.end_stake])
    def h_line_on(self, s, a): self.set_selection([self.state.grading_line])
    def h_off(self, s, a):
        uidoc.Selection.SetElementIds(List[ElementId]())
        uidoc.RefreshActiveView()

    def mode_changed(self, sender, args):
        self.state.mode = "slope" if self.Rb_UseSlope.IsChecked else "stakes"
        self.Tb_Slope.IsEnabled = (self.state.mode == "slope")
        self.Btn_Run.IsEnabled = self.state.ready
        self.Btn_Edging.IsEnabled = self.state.ready

    def update_state_from_ui(self):
        """Pushes UI values to Memory (converting Display -> Internal)."""
        self.state.width = str(UnitHelper.to_internal(self.Tb_Width.Text))
        self.state.falloff = str(UnitHelper.to_internal(self.Tb_Falloff.Text))
        self.state.grid = str(UnitHelper.to_internal(self.Tb_Grid.Text))
        self.state.slope_val = self.Tb_Slope.Text

    def a_stakes(self, s, a): self.update_state_from_ui(); self.state.next_action = "select_stakes"; self.Close()
    def a_line(self, s, a): self.update_state_from_ui(); self.state.next_action = "select_line"; self.Close()
    def a_swap(self, s, a): self.update_state_from_ui(); self.state.next_action = "swap"; self.Close()
    def a_run(self, s, a): self.update_state_from_ui(); self.state.next_action = "sculpt"; self.Close()
    def a_edge(self, s, a): self.update_state_from_ui(); self.state.next_action = "edge"; self.Close()
    def a_stitch(self, s, a): self.update_state_from_ui(); self.state.next_action = "stitch"; self.Close()
    def a_load(self, s, a): self.update_state_from_ui(); self.state.next_action = "load_recipe"; self.Close()

# ==========================================
# 3. HELPERS
# ==========================================
class UniversalFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, r, p): return True

def flatten(pt): return XYZ(pt.X, pt.Y, 0)
def lerp(a, b, t): return a + t * (b - a)

def get_toposolid_max_z(toposolid):
    bb = toposolid.get_BoundingBox(None)
    if bb: return bb.Max.Z
    return 1000.0 

def is_point_on_solid(intersector, pt, start_z):
    if not intersector: return True 
    origin = XYZ(pt.X, pt.Y, start_z + 10.0) 
    try:
        if intersector.FindNearest(origin, XYZ(0, 0, -1)): return True
    except: pass
    return False

def is_too_close(candidate_pt, occupied_points, tolerance=MIN_DIST_TOLERANCE):
    cand_flat = flatten(candidate_pt)
    for existing in occupied_points:
        if cand_flat.DistanceTo(flatten(existing)) < tolerance: return True
    return False

def get_surface_z(intersector, pt, start_z):
    if not intersector: return None
    origin = XYZ(pt.X, pt.Y, start_z + 10.0) 
    try:
        context = intersector.FindNearest(origin, XYZ(0, 0, -1))
        if context: return context.GetReference().GlobalPoint.Z
    except: pass
    return None

def get_boundary_curves(toposolid):
    opt = Options(); opt.ComputeReferences = True
    geom = toposolid.get_Geometry(opt)
    curves = []
    if not geom: return []
    for obj in geom:
        if isinstance(obj, Solid):
            for edge in obj.Edges:
                c = edge.AsCurve()
                if c: curves.append(c)
    return curves

def get_z_from_curve_param(curve, projected_result, z_start, z_end):
    p_min, p_max = curve.GetEndParameter(0), curve.GetEndParameter(1)
    norm_p = (projected_result.Parameter - p_min) / (p_max - p_min)
    norm_p = max(0.0, min(1.0, norm_p))
    return z_start + (norm_p * (z_end - z_start))

def get_line_ends(curve):
    return curve.GetEndPoint(0), curve.GetEndPoint(1)

def validate_input(state, log):
    """Failsafe check for basic inputs before processing."""
    if not state.start_stake:
        log.error("Start Stake is missing.")
        return False
    if not state.grading_line:
        log.error("Grading Line is missing.")
        return False
    
    # Check if elements still exist in doc
    try:
        if not state.start_stake.IsValidObject:
            log.error("Start Stake element is no longer valid.")
            return False
        if not state.grading_line.IsValidObject:
            log.error("Grading Line element is no longer valid.")
            return False
    except:
        log.error("Invalid element reference.")
        return False

    return True

def calculate_and_adjust_stakes(state, log):
    """Calculates Z levels and adjusts end stake if needed."""
    if not state.grading_line or not isinstance(state.grading_line, CurveElement):
        log.error("Invalid Grading Line selected.")
        raise Exception("Invalid Line")

    curve = state.grading_line.GeometryCurve
    l_start, l_end = get_line_ends(curve)
    
    # Failsafe: Stake might not have a LocationPoint if user picked wrong family
    if not hasattr(state.start_stake, "Location") or not hasattr(state.start_stake.Location, "Point"):
        log.error("Start Stake does not have a valid location point.")
        raise Exception("Invalid Stake")

    u_start_pt = state.start_stake.Location.Point
    z_start = u_start_pt.Z
    dist_start = u_start_pt.DistanceTo(l_start)
    dist_end = u_start_pt.DistanceTo(l_end)
    is_flipped = dist_end < dist_start 
    length = curve.Length
    z_end = 0.0
    
    if state.mode == "slope":
        try:
            pct = float(state.slope_val) / 100.0
            z_end = z_start + (length * pct)
            if state.end_stake:
                t_move = Transaction(doc, "Adjust Stake Height")
                t_move.Start()
                try:
                    current_pt = state.end_stake.Location.Point
                    diff_z = z_end - current_pt.Z
                    if abs(diff_z) > 0.001:
                        vec = XYZ(0, 0, diff_z)
                        ElementTransformUtils.MoveElement(doc, state.end_stake.Id, vec)
                except: 
                    # Fallback for families constrained to host
                    try:
                        p = state.end_stake.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                        if p and not p.IsReadOnly: p.Set(z_end)
                    except: pass
                t_move.Commit()
        except Exception as e:
            log.error("Failed to adjust End Stake slope.", e)
            z_end = z_start
    else:
        if not state.end_stake or not hasattr(state.end_stake.Location, "Point"):
             log.error("End Stake is missing or invalid for 'Match Stakes' mode.")
             raise Exception("Invalid End Stake")
        z_end = state.end_stake.Location.Point.Z
        
    return (z_end, z_start) if is_flipped else (z_start, z_end)

# ==========================================
# 5. EXECUTION
# ==========================================
def perform_load_recipe(state):
    log = BatchLogger()
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        if not ref: return
        data = GradingRecipe.read_recipe(doc.GetElement(ref))
        if data:
            state.width = str(data.get("width", "6.0"))
            state.falloff = str(data.get("falloff", "10.0")),
            state.grid = str(data.get("grid", "3.0"))
            log.info("Recipe loaded successfully.")
        else:
            log.info("No grading recipe found on this element.")
    except Exception as e:
        log.error("Failed to load recipe.", e)
    finally:
        log.show()

def get_chain_of_edges(toposolid, start_edge):
    """
    Finds the boundary loop.
    Priority 1: Toposolid Sketch (Base Curves).
    Priority 2: Geometry Edge Graph (Filtered).
    """
    
    # --- STRATEGY 1: SKETCH (Base Curves) ---
    try:
        sketch = None
        # 1. Try SketchId property
        if hasattr(toposolid, "SketchId") and toposolid.SketchId != ElementId.InvalidElementId:
            sketch = doc.GetElement(toposolid.SketchId)
        
        # 2. Try Dependent Elements if property fails
        if not sketch:
            ids = toposolid.GetDependentElements(UniversalFilter()) # Re-use generic filter or None
            for eid in ids:
                el = doc.GetElement(eid)
                if el and "Sketch" in el.GetType().Name:
                    sketch = el
                    break
        
        if sketch and sketch.Profile:
            # Found Sketch. Find the loop matching selected edge.
            
            # Midpoint of selected edge for proximity check
            sel_curve = start_edge.AsCurve()
            mid_pt = sel_curve.Evaluate(0.5, True)
            mid_flat = flatten(mid_pt)
            
            best_loop = None
            min_dist = 1.0 # Tolerance ft
            
            # Iterate CurveArrArray (Profile is list of loops)
            for curve_array in sketch.Profile:
                # Check if this loop is "close" to our selection
                is_match = False
                for i in range(curve_array.Size):
                    sc = curve_array.get_Item(i)
                    
                    # Project selected midpoint to sketch curve (2D check)
                    # We flatten the sketch curve ends to check distance
                    sp0 = flatten(sc.GetEndPoint(0))
                    sp1 = flatten(sc.GetEndPoint(1))
                    
                    # Simple distance to segment check
                    # Or assume sketch curve is planar Z-flat, just ignore Z
                    # We can use our 'flatten' to create a new Line for check? 
                    # No, generic curve might be arc.
                    # Let's project 'mid_flat' onto 'sc' ignoring Z?
                    # Hard to do generically without creating geometry.
                    
                    # Quick check: Is mid_flat close to endpoints?
                    if mid_flat.DistanceTo(sp0) < 5.0 or mid_flat.DistanceTo(sp1) < 5.0:
                        # Closer check
                        pass

                    # Robust 2D Project:
                    # Create flat version of sc?
                    # Since sc is likely flat, just setting Z to 0 might work.
                    try:
                        # Only checking distance to "unbounded" curve might give false positives
                        # But Profile curves are bounded.
                        # Let's just check endpoints for now as a heuristic 
                        # OR check if mid_pt projects onto it.
                        res = sc.Project(mid_pt)
                        if res:
                            # Distance in XY plane
                            p_res = res.XYZPoint
                            d_xy = flatten(p_res).DistanceTo(mid_flat)
                            if d_xy < min_dist:
                                is_match = True
                                break
                    except: pass
                
                if is_match:
                    best_loop = []
                    for i in range(curve_array.Size):
                        best_loop.append(curve_array.get_Item(i))
                    return best_loop

    except Exception as e:
        # Log failure silently or to debug, fallback to method 2
        # print("Sketch lookup failed: " + str(e))
        pass

    # --- STRATEGY 2: GEOMETRY GRAPH (Fallback) ---
    opt = Options()
    opt.ComputeReferences = True
    geom = toposolid.get_Geometry(opt)
    
    all_edges = []
    if geom:
        for obj in geom:
            if isinstance(obj, Solid):
                for e in obj.Edges:
                    all_edges.append(e)
    
    # 1. Build Adjacency Graph (Endpoint -> List of Edges)
    # Key: (X, Y, Z) rounded tuple
    # Value: List of Edge objects
    adj_map = {}
    
    def pt_key(pt):
        return (round(pt.X, 4), round(pt.Y, 4), round(pt.Z, 4))
    
    # Helper to check if an edge is internal (connects two top faces)
    def is_boundary_candidate(edge):
        try:
            # Get the two faces sharing this edge
            f0 = edge.GetFace(0)
            f1 = edge.GetFace(1)
            
            if not f0 or not f1: return True # Keep if we can't determine (safe fallback)
            
            def is_up(face):
                # Using face.ComputeNormal at UV center is robust for Planar/Ruled faces
                bbox = face.GetBoundingBox()
                center = (bbox.Min + bbox.Max) / 2.0
                try:
                    if hasattr(face, "FaceNormal"): return face.FaceNormal.Z > 0.1
                    res = face.Project(center)
                    if res:
                        norm = face.ComputeNormal(res.UVPoint)
                        return norm.Z > 0.1
                except: pass
                return False 
                
            up0 = is_up(f0)
            up1 = is_up(f1)
            
            if up0 and up1: return False
            return True
        except:
            return True 
            
    start_id = start_edge.Id
    
    filtered_count = 0
    for e in all_edges:
        c = e.AsCurve()
        if not c: continue
        
        # Optimization: Always include the user-selected edge without checking
        if e.Id != start_id:
            if not is_boundary_candidate(e):
                continue
        
        filtered_count += 1
        
        p0 = pt_key(c.GetEndPoint(0))
        p1 = pt_key(c.GetEndPoint(1))
        
        if p0 not in adj_map: adj_map[p0] = []
        if p1 not in adj_map: adj_map[p1] = []
        
        adj_map[p0].append(e)
        adj_map[p1].append(e)
        
    # 2. Traverse Graph using BFS/DFS
    visited_ids = set()
    chain = []
    stack = [start_edge]
    visited_ids.add(start_id)
    
    while stack:
        current_edge = stack.pop(0) # BFS
        chain.append(current_edge)
        
        c = current_edge.AsCurve()
        p0 = pt_key(c.GetEndPoint(0))
        p1 = pt_key(c.GetEndPoint(1))
        
        # Check neighbors at both ends
        for p in [p0, p1]:
            neighbors = adj_map.get(p, [])
            for n_edge in neighbors:
                nid = n_edge.Id
                if nid not in visited_ids:
                    visited_ids.add(nid)
                    stack.append(n_edge)

    return [e.AsCurve() for e in chain] 
                
    return [e.AsCurve() for e in chain]

def perform_manual_stitch(state):
    log = BatchLogger()
    log.info("--- MANUAL STITCH STARTED ---")
    
    try:
        # Use Grid size as the search tolerance/radius
        snap_dist = float(state.grid)
        log.info("Snap Tolerance (Grid): {:.3f} ft".format(snap_dist))
        
        if snap_dist <= 0: snap_dist = 1.0

        try:
            ref_edge = uidoc.Selection.PickObject(ObjectType.Edge, "Select Boundary Edge to Stitch")
        except: 
            return # Cancelled
            
        edge_elem = doc.GetElement(ref_edge)
        edge_geom = edge_elem.GetGeometryObjectFromReference(ref_edge)
        
        if not isinstance(edge_geom, Edge): 
            log.error("Selected object is not a valid Edge.")
            log.show(); return

        toposolid = edge_elem
        log.info("Selected Toposolid ID: {}".format(toposolid.Id))
        
        # NEW: Get the full chain of connected edges
        log.info("Tracing connected boundary edges...")
        boundary_curves = get_chain_of_edges(toposolid, edge_geom)
        log.info("Identified {} connected boundary segments.".format(len(boundary_curves)))
        
        if not boundary_curves:
            log.error("Could not find any boundary edges. Check geometry.")
            log.show(); return

        tg = TransactionGroup(doc, "Stitch Edge")
        tg.Start()
        
        t = Transaction(doc, "Stitch")
        t.Start()
        
        try:
            editor = toposolid.GetSlabShapeEditor()
            editor.Enable()
            
            existing_coords = set()
            candidates = []
            
            # Cache existing vertices
            for v in editor.SlabShapeVertices:
                pos = v.Position
                existing_coords.add((round(pos.X, 4), round(pos.Y, 4)))
                candidates.append(v)
            
            log.info("Internal Vertices to Check: {}".format(len(candidates)))
            
            points_to_add = []
            
            # 2. Find Candidates: Internal Points -> Project -> Closest Edge in Chain
            check_count = 0
            match_count = 0
            
            for v in candidates:
                check_count += 1
                best_proj = None
                best_dist = 9999.0
                
                # Iterate all segments in the boundary chain to find the closest one
                # Optimization: Could use a spatial index, but N_edges is usually small (<100)
                for b_curve in boundary_curves:
                    try:
                        proj_res = b_curve.Project(v.Position)
                        if not proj_res: continue
                        
                        p_target = proj_res.XYZPoint
                        v_flat = flatten(v.Position)
                        p_target_flat = flatten(p_target)
                        dist = v_flat.DistanceTo(p_target_flat)
                        
                        if dist < best_dist:
                            best_dist = dist
                            best_proj = p_target
                    except: pass
                
                # Logic: If close enough (but not ON the edge), propose new point
                if best_proj and (0.005 < best_dist < snap_dist):
                    
                    new_pt = XYZ(best_proj.X, best_proj.Y, v.Position.Z)
                    
                    key = (round(new_pt.X, 4), round(new_pt.Y, 4))
                    if key not in existing_coords:
                        points_to_add.append(new_pt)
                        existing_coords.add(key)
                        match_count += 1
            
            log.info("Candidates Found within Tolerance: {}".format(match_count))
            
            # 3. Add Points
            added_count = 0
            for p in points_to_add:
                try:
                    editor.AddPoint(p)
                    added_count += 1
                except Exception as e:
                    pass
            
            log.info("Successfully Added Points: {}".format(added_count))

            t.Commit()
            tg.Assimilate()
            
        except Exception as e:
            t.RollBack()
            tg.RollBack()
            log.error("Stitch transaction failed.", "{}\n{}".format(e, traceback.format_exc()))
            
    except Exception as e:
        log.error("Critical error in stitch tool.", "{}\n{}".format(e, traceback.format_exc()))
    finally:
        log.show()
def perform_swap(state):
    if state.start_stake and state.end_stake:
        state.start_stake, state.end_stake = state.end_stake, state.start_stake

def perform_sculpt(state):
    log = BatchLogger()
    if not validate_input(state, log):
        log.show()
        return

    try:
        # 1. Parse Inputs & Log
        w_int = float(state.width)
        f_int = float(state.falloff)
        g_int = float(state.grid)
        
        log.info("--- SCULPT STARTED ---")
        log.info("Parameters (Internal Ft): Width={:.2f}, Falloff={:.2f}, Grid={:.2f}".format(w_int, f_int, g_int))
        log.info("Parameters (Display): Width={:.2f}, Falloff={:.2f}, Grid={:.2f}".format(
            UnitHelper.from_internal(w_int), 
            UnitHelper.from_internal(f_int), 
            UnitHelper.from_internal(g_int)
        ))

        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid to Grade")
            toposolid = doc.GetElement(ref)
            log.info("Selected Toposolid ID: {}".format(toposolid.Id))
        except:
            return # Cancelled
            
        if not toposolid:
            log.error("Invalid Toposolid selection.")
            log.show(); return

    except Exception as e:
        log.error("Invalid parameter inputs.", e)
        log.show(); return

    tg = TransactionGroup(doc, "Sculpt Terrain")
    tg.Start()
    
    try:
        # Save Recipe
        rec = {"width": state.width, "falloff": state.falloff, "grid": state.grid}
        t_rec = Transaction(doc, "Save Recipe")
        t_rec.Start()
        GradingRecipe.save_recipe(toposolid, rec)
        t_rec.Commit()

        # Calculation
        z_s, z_e = calculate_and_adjust_stakes(state, log)
        log.info("Stake Elevations: Start={:.2f}, End={:.2f}".format(z_s, z_e))
        
        ids = List[ElementId]([toposolid.Id])
        intersector = None
        ray_start_z = get_toposolid_max_z(toposolid)
        
        if doc.ActiveView.ViewType == ViewType.ThreeD:
            intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, doc.ActiveView)
        else:
            log.info("Warning: Active view is not 3D. Raycasting for new points might be less accurate.")

        curve = state.grading_line.GeometryCurve
        log.info("Guide Line Length: {:.2f}".format(curve.Length))

        core_rad = w_int / 2.0
        total_rad = core_rad + f_int
        
        # Phase 1: Densify
        log.info("--- PHASE 1: DENSIFY ---")
        t1 = Transaction(doc, "Densify")
        t1.Start()
        editor = toposolid.GetSlabShapeEditor()
        editor.Enable()
        occupied_points = [v.Position for v in editor.SlabShapeVertices]
        log.info("Initial Vertex Count: {}".format(len(occupied_points)))
        
        bb = state.grading_line.get_BoundingBox(None)
        start_x = math.floor((bb.Min.X - total_rad - g_int) / g_int) * g_int
        end_x   = math.ceil((bb.Max.X + total_rad + g_int) / g_int) * g_int
        start_y = math.floor((bb.Min.Y - total_rad - g_int) / g_int) * g_int
        end_y   = math.ceil((bb.Max.Y + total_rad + g_int) / g_int) * g_int
        
        log.info("Grid Search Bounds: X[{:.1f}, {:.1f}] Y[{:.1f}, {:.1f}]".format(start_x, end_x, start_y, end_y))
        
        grid_pts = []
        candidates_checked = 0
        x = start_x
        while x <= end_x:
            y = start_y
            while y <= end_y:
                candidates_checked += 1
                t_pt = XYZ(x, y, 0)
                res = curve.Project(t_pt)
                
                # Check 2D distance to curve
                if res and flatten(t_pt).DistanceTo(flatten(res.XYZPoint)) < (total_rad + g_int):
                    if is_point_on_solid(intersector, t_pt, ray_start_z):
                        rz = get_surface_z(intersector, t_pt, ray_start_z)
                        if rz is not None: 
                            grid_pts.append(XYZ(x, y, rz))
                y += g_int
            x += g_int
            
        log.info("Grid Points Found on Solid: {} (out of {} checked)".format(len(grid_pts), candidates_checked))
        
        added_count = 0
        for p in grid_pts:
            if not is_too_close(p, occupied_points):
                try: 
                    editor.AddPoint(p)
                    occupied_points.append(p)
                    added_count += 1
                except: pass
        t1.Commit()
        
        # Phase 2: Sculpt
        log.info("--- PHASE 2: SCULPT ---")
        t2 = Transaction(doc, "Sculpt")
        t2.Start()
        updates = []
        
        # Refresh vertices after densify
        current_verts = [v for v in editor.SlabShapeVertices]
        log.info("Total Vertices to Process: {}".format(len(current_verts)))
        
        sample_log = []
        
        for v in current_verts:
            res = curve.Project(v.Position)
            if not res: continue
            
            d = flatten(v.Position).DistanceTo(flatten(res.XYZPoint))
            if d > total_rad: continue
            
            target_z = get_z_from_curve_param(curve, res, z_s, z_e)
            new_z = v.Position.Z
            
            # Logic check
            is_core = d <= core_rad
            
            if is_core: 
                new_z = target_z
            else:
                t_val = (d - core_rad) / f_int
                t_val = 1.0 if t_val > 1.0 else t_val
                smooth_t = t_val * t_val * (3 - 2 * t_val)
                new_z = lerp(target_z, new_z, smooth_t)
            
            if abs(new_z - v.Position.Z) > 0.005:
                updates.append(XYZ(v.Position.X, v.Position.Y, new_z))
                if len(sample_log) < 3:
                    sample_log.append("Pt ({:.1f}, {:.1f}): Z {:.2f} -> {:.2f} (Dist={:.2f}, Core={})".format(
                        v.Position.X, v.Position.Y, v.Position.Z, new_z, d, is_core
                    ))
        
        modified_count = len(updates)
        log.info("Points Identified for Modification: {}".format(modified_count))
        if sample_log:
            log.info("Sample Changes:\n" + "\n".join(sample_log))
        
        for p in updates:
            try: editor.AddPoint(p)
            except: pass
        
        t2.Commit()
        tg.Assimilate()
        
        log.info("Sculpt Transaction Committed.")
        
        log.info("Sculpt Complete.\nPoints Added: {}\\nPoints Adjusted: {}".format(added_count, modified_count))

    except Exception as e:
        tg.RollBack()
        log.error("Sculpting Failed", "{}\n{}".format(e, traceback.format_exc()))
    
    finally:
        log.show()

def perform_edging(state):
    log = BatchLogger()
    if not validate_input(state, log):
        log.show(); return

    try:
        w = float(state.width); g = float(state.grid)
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid for Edging")
            toposolid = doc.GetElement(ref)
        except: return
    except Exception as e:
        log.error("Invalid inputs for edging.", e)
        log.show(); return

    tg = TransactionGroup(doc, "Edging")
    tg.Start()
    
    try:
        z_s, z_e = calculate_and_adjust_stakes(state, log)
        
        ids = List[ElementId]([toposolid.Id])
        intersector = None
        ray_start_z = get_toposolid_max_z(toposolid)
        if doc.ActiveView.ViewType == ViewType.ThreeD:
            intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, doc.ActiveView)
        
        edge_offset = w / 2.0
        edge_res = g * 0.5 
        curve = state.grading_line.GeometryCurve
        editor = toposolid.GetSlabShapeEditor()
        editor.Enable()
        
        to_move = [] 
        to_add = []
        
        all_verts = [v for v in editor.SlabShapeVertices]
        
        # 1. Snap existing nearby points to exact edge
        for v in all_verts: 
            res = curve.Project(v.Position)
            if res and abs(flatten(v.Position).DistanceTo(flatten(res.XYZPoint)) - edge_offset) < 1.0:
                vec = (flatten(v.Position) - flatten(res.XYZPoint)).Normalize()
                exact_xy = flatten(res.XYZPoint) + (vec * edge_offset)
                exact_z = get_z_from_curve_param(curve, res, z_s, z_e)
                to_move.append((v, XYZ(exact_xy.X, exact_xy.Y, exact_z)))
        
        # 2. Add new points along the exact edge
        step_t = edge_res / curve.Length
        step_t = 0.05 if step_t > 0.05 else step_t
        t_val = 0.0
        
        while t_val <= 1.001:
            eval_t = max(0.0, min(1.0, t_val))
            center_pt = curve.Evaluate(eval_t, True)
            tangent = curve.ComputeDerivatives(eval_t, True).BasisX.Normalize()
            normal = tangent.CrossProduct(XYZ.BasisZ)
            
            road_z = get_z_from_curve_param(curve, curve.Project(center_pt), z_s, z_e)
            
            for side in [1.0, -1.0]:
                offset_vec = normal * (side * edge_offset)
                final_pt = center_pt + offset_vec
                final_pt = XYZ(final_pt.X, final_pt.Y, road_z)
                
                # Check if this point is actually on the solid's footprint
                if is_point_on_solid(intersector, final_pt, ray_start_z): 
                    to_add.append(final_pt)
            
            t_val += step_t
            
        t = Transaction(doc, "Apply Edging")
        t.Start()
        
        occupied_points = [v.Position for v in all_verts]
        
        # Apply moves
        for item in to_move:
            try: 
                editor.ModifySlabShapeVertex(item[0], item[1])
                occupied_points.append(item[1])
            except: pass
            
        # Apply adds
        for pt in to_add:
            if not is_too_close(pt, occupied_points):
                try: 
                    editor.AddPoint(pt)
                    occupied_points.append(pt)
                except: pass 
                
        t.Commit()
        tg.Assimilate()
        log.info("Edging Complete.\nSnapped: {}\\nAdded: {}".format(len(to_move), len(to_add)))
        
    except Exception as e:
        tg.RollBack()
        log.error("Edging failed.", e)
    finally:
        log.show()

# ==========================================
# 6. LOOP
# ==========================================
if __name__ == '__main__':
    state = GradingState()
    while True:
        win = GradingWindow(state)
        win.ShowDialog()
        action = state.next_action
        state.next_action = None 
        
        if not action:
            save_state_to_disk(state)
            break 
            
        elif action == "select_stakes":
            try:
                state.start_stake = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Start Stake"))
                state.end_stake = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "End Stake"))
                save_state_to_disk(state)
            except: pass # Cancelled selection
            
        elif action == "select_line":
            try: 
                state.grading_line = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Guide Line"))
                save_state_to_disk(state)
            except: pass # Cancelled selection
            
        elif action == "swap": 
            perform_swap(state)
            save_state_to_disk(state)
            
        elif action == "sculpt": 
            perform_sculpt(state)
            save_state_to_disk(state)
            
        elif action == "edge": 
            perform_edging(state)
            save_state_to_disk(state)
            
        elif action == "stitch": 
            perform_manual_stitch(state)
            
        elif action == "load_recipe": 
            perform_load_recipe(state)
            save_state_to_disk(state)
