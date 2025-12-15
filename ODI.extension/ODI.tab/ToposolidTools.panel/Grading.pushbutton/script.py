# -*- coding: utf-8 -*-
import sys
import os
import json
import math
import clr

# --- ASSEMBLIES ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

# --- IMPORTS ---
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    XYZ, Transaction, TransactionGroup, ElementId, BuiltInParameter,
    ReferenceIntersector, FindReferenceTarget, Options, Solid, ViewType, Edge
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB.ExtensibleStorage import SchemaBuilder, Schema, Entity, FieldBuilder, AccessLevel
from System import Guid
from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# 1. SETTINGS & SCHEMA
# ==========================================
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "grading_settings.json")
RECIPE_SCHEMA_GUID = Guid("A4B9C8D2-1234-4567-8901-ABCDEF123456")

MIN_DIST_TOLERANCE = 0.25 

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

def load_settings():
    defaults = {"width": "6.0", "falloff": "10.0", "grid": "3.0", "slope": "2.0", "mode": "stakes"}
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                defaults.update(json.load(f))
    except: pass
    return defaults

def save_settings(width, falloff, grid, slope, mode):
    data = {"width": str(width), "falloff": str(falloff), "grid": str(grid), "slope": str(slope), "mode": mode}
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(data, f)
    except: pass

# ==========================================
# 2. STATE & UI
# ==========================================
class GradingState(object):
    def __init__(self):
        self.start_stake = None
        self.end_stake = None
        self.grading_line = None
        
        sets = load_settings()
        self.width = sets["width"]
        self.falloff = sets["falloff"]
        self.grid = sets["grid"]
        self.slope_val = sets["slope"]
        self.mode = sets["mode"] 
        self.next_action = None

    @property
    def ready(self):
        if self.mode == "slope":
            return bool(self.start_stake and self.grading_line)
        return bool(self.start_stake and self.end_stake and self.grading_line)

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
            self.Rb_UseSlope.IsChecked = True
            self.Tb_Slope.IsEnabled = True
        else:
            self.Rb_MatchStakes.IsChecked = True
            self.Tb_Slope.IsEnabled = False

        self.Lb_StartStake.Text = "Start: {}".format(fmt(self.state.start_stake))
        self.Lb_StartStake.Foreground = green if self.state.start_stake else red
        self.Lb_EndStake.Text = "End: {}".format(fmt(self.state.end_stake))
        self.Lb_EndStake.Foreground = green if self.state.end_stake else red
        self.Lb_Line.Text = "Line: {}".format(fmt(self.state.grading_line))
        self.Lb_Line.Foreground = green if self.state.grading_line else red

        self.Tb_Width.Text = self.state.width
        self.Tb_Falloff.Text = self.state.falloff
        self.Tb_Grid.Text = self.state.grid
        self.Tb_Slope.Text = self.state.slope_val

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

    def mode_changed(self, sender, args):
        self.state.mode = "slope" if self.Rb_UseSlope.IsChecked else "stakes"
        self.Tb_Slope.IsEnabled = (self.state.mode == "slope")
        self.Btn_Run.IsEnabled = self.state.ready
        self.Btn_Edging.IsEnabled = self.state.ready

    def save(self):
        self.state.width = self.Tb_Width.Text
        self.state.falloff = self.Tb_Falloff.Text
        self.state.grid = self.Tb_Grid.Text
        self.state.slope_val = self.Tb_Slope.Text
        save_settings(self.state.width, self.state.falloff, self.state.grid, self.state.slope_val, self.state.mode)

    def a_stakes(self, s, a): self.save(); self.state.next_action = "select_stakes"; self.Close()
    def a_line(self, s, a): self.save(); self.state.next_action = "select_line"; self.Close()
    def a_swap(self, s, a): self.save(); self.state.next_action = "swap"; self.Close()
    def a_run(self, s, a): self.save(); self.state.next_action = "sculpt"; self.Close()
    def a_edge(self, s, a): self.save(); self.state.next_action = "edge"; self.Close()
    def a_stitch(self, s, a): self.save(); self.state.next_action = "stitch"; self.Close()
    def a_load(self, s, a): self.save(); self.state.next_action = "load_recipe"; self.Close()

# ==========================================
# 3. HELPERS
# ==========================================
class UniversalFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, r, p): return True

def flatten(pt): return XYZ(pt.X, pt.Y, 0)
def lerp(a, b, t): return a + t * (b - a)

def is_point_on_solid(intersector, pt):
    if not intersector: return True 
    origin = XYZ(pt.X, pt.Y, pt.Z + 1000.0) 
    try:
        if intersector.FindNearest(origin, XYZ(0, 0, -1)): return True
    except: pass
    return False

def is_too_close(candidate_pt, occupied_points, tolerance=MIN_DIST_TOLERANCE):
    cand_flat = flatten(candidate_pt)
    for existing in occupied_points:
        if cand_flat.DistanceTo(flatten(existing)) < tolerance: return True
    return False

def get_surface_z(intersector, pt):
    if not intersector: return None
    origin = XYZ(pt.X, pt.Y, pt.Z + 2000.0) 
    try:
        context = intersector.FindNearest(origin, XYZ(0, 0, -1))
        if context: return context.GetReference().GlobalPoint.Z
    except: pass
    return None

def get_z_from_curve_param(curve, projected_result, z_start, z_end):
    p_min, p_max = curve.GetEndParameter(0), curve.GetEndParameter(1)
    norm_p = (projected_result.Parameter - p_min) / (p_max - p_min)
    norm_p = max(0.0, min(1.0, norm_p))
    return z_start + (norm_p * (z_end - z_start))

def get_line_ends(curve):
    return curve.GetEndPoint(0), curve.GetEndPoint(1)

def calculate_slope_params(state):
    curve = state.grading_line.GeometryCurve
    l_start, l_end = get_line_ends(curve)
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
                t_move = Transaction(doc, "Adj Stake")
                t_move.Start()
                try:
                    p = state.end_stake.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                    if p and not p.IsReadOnly: p.Set(z_end)
                    else: state.end_stake.Location.Point = XYZ(state.end_stake.Location.Point.X, state.end_stake.Location.Point.Y, z_end)
                except: pass
                t_move.Commit()
        except: z_end = z_start
    else:
        z_end = state.end_stake.Location.Point.Z
    return (z_end, z_start) if is_flipped else (z_start, z_end)

# ==========================================
# 4. ACTIONS
# ==========================================
def perform_load_recipe(state):
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        data = GradingRecipe.read_recipe(doc.GetElement(ref))
        if data:
            state.width = str(data.get("width", "6.0"))
            state.falloff = str(data.get("falloff", "10.0"))
            state.grid = str(data.get("grid", "3.0"))
            try:
                state.grading_line = doc.GetElement(ElementId(data["line_id"]))
                state.start_stake = doc.GetElement(ElementId(data["start_id"]))
                state.end_stake = doc.GetElement(ElementId(data["end_id"]))
            except: pass
    except: pass

def perform_manual_stitch(state):
    try:
        g = float(state.grid)
        # Select Edge
        ref_edge = uidoc.Selection.PickObject(ObjectType.Edge, "Select Toposolid Boundary Edge to Stitch")
        edge_elem = doc.GetElement(ref_edge)
        
        # Get actual Curve geometry from the Reference
        edge_geom = edge_elem.GetGeometryObjectFromReference(ref_edge)
        if not isinstance(edge_geom, Edge):
            print("Selection was not an Edge.")
            return
        
        # This is the single curve we will walk
        b_curve = edge_geom.AsCurve() 
        toposolid = edge_elem # The element IS the toposolid (or floor)

        tg = TransactionGroup(doc, "Stitch Edge")
        tg.Start()
        
        t = Transaction(doc, "Stitch")
        t.Start()
        
        editor = toposolid.GetSlabShapeEditor()
        editor.Enable()
        
        all_verts = [v.Position for v in editor.SlabShapeVertices]
        points_to_add = []
        
        # Search radius is Grid Resolution (the typical gap size)
        search_dist = g
        step_size = g * 0.5 
        
        length = b_curve.Length
        steps = int(length / step_size)
        if steps < 2: steps = 2
        
        for i in range(steps + 1):
            t_val = float(i) / float(steps)
            b_pt = b_curve.Evaluate(t_val, True)
            
            # 1. Find Nearest Neighbor (2D Search)
            nearest_dist = 9999.0
            nearest_z = 0.0
            nearest_pt_ref = None
            
            for v in all_verts:
                # 2D Distance
                d = flatten(b_pt).DistanceTo(flatten(v))
                if d < nearest_dist:
                    nearest_dist = d
                    nearest_z = v.Z
                    nearest_pt_ref = v
            
            # 2. Logic: Is gap valid?
            # Must be closer than Grid (so it's a neighbor)
            # Must be further than Tolerance (so it's not already ON the boundary)
            if nearest_dist < search_dist and nearest_dist > MIN_DIST_TOLERANCE:
                
                # Check duplication in queue
                if not is_too_close(b_pt, points_to_add, MIN_DIST_TOLERANCE):
                    # ADD: X/Y from Boundary, Z from Neighbor
                    points_to_add.append(XYZ(b_pt.X, b_pt.Y, nearest_z))

        count = 0
        for p in points_to_add:
            try: editor.AddPoint(p); count += 1
            except: pass
            
        t.Commit()
        tg.Assimilate()
        print("Manual Stitch: Added {} points on selected edge.".format(count))

    except Exception as e:
        print("Stitch Cancelled: {}".format(e))

def perform_swap(state):
    state.start_stake, state.end_stake = state.end_stake, state.start_stake

def perform_sculpt(state):
    try:
        w = float(state.width); f = float(state.falloff); g = float(state.grid)
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        toposolid = doc.GetElement(ref)
    except: return

    tg = TransactionGroup(doc, "Sculpt Terrain"); tg.Start()
    try:
        rec = {"width": w, "falloff": f, "grid": g, "line_id": get_id_val(state.grading_line), "start_id": get_id_val(state.start_stake), "end_id": get_id_val(state.end_stake)}
        t_rec = Transaction(doc, "Save Recipe"); t_rec.Start()
        GradingRecipe.save_recipe(toposolid, rec); t_rec.Commit()

        z_s, z_e = calculate_and_adjust_stakes(state)
        ids = List[ElementId]([toposolid.Id])
        intersector = None
        if doc.ActiveView.ViewType == ViewType.ThreeD:
            intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, doc.ActiveView)

        curve = state.grading_line.GeometryCurve
        core_rad = w / 2.0; total_rad = core_rad + f
        
        t1 = Transaction(doc, "Densify"); t1.Start()
        editor = toposolid.GetSlabShapeEditor(); editor.Enable()
        occupied_points = [v.Position for v in editor.SlabShapeVertices]
        bb = state.grading_line.get_BoundingBox(None)
        buffer = total_rad + g 
        start_x = math.floor((bb.Min.X - buffer) / g) * g
        end_x   = math.ceil((bb.Max.X + buffer) / g) * g
        start_y = math.floor((bb.Min.Y - buffer) / g) * g
        end_y   = math.ceil((bb.Max.Y + buffer) / g) * g

        grid_pts = []
        x = start_x
        while x <= end_x:
            y = start_y
            while y <= end_y:
                t_pt = XYZ(x, y, z_s)
                res = curve.Project(t_pt)
                if res and flatten(t_pt).DistanceTo(flatten(res.XYZPoint)) < (total_rad + g):
                    if is_point_on_solid(intersector, t_pt):
                        rz = get_surface_z(intersector, t_pt)
                        if rz: grid_pts.append(XYZ(x, y, rz))
                y += g
            x += g
        for p in grid_pts:
            if not is_too_close(p, occupied_points):
                try: editor.AddPoint(p); occupied_points.append(p)
                except: pass
        t1.Commit()
        
        t2 = Transaction(doc, "Sculpt"); t2.Start()
        updates = []
        for v in editor.SlabShapeVertices:
            res = curve.Project(v.Position)
            if not res: continue
            d = flatten(v.Position).DistanceTo(flatten(res.XYZPoint))
            if d > total_rad: continue
            target_z = get_z_from_curve_param(curve, res, z_s, z_e)
            new_z = v.Position.Z
            if d <= core_rad: new_z = target_z
            else:
                t_val = (d - core_rad) / f; t_val = 1.0 if t_val > 1.0 else t_val
                smooth_t = t_val * t_val * (3 - 2 * t_val)
                new_z = lerp(target_z, new_z, smooth_t)
            if abs(new_z - v.Position.Z) > 0.005:
                updates.append(XYZ(v.Position.X, v.Position.Y, new_z))
        for p in updates:
            try: editor.AddPoint(p)
            except: pass
        t2.Commit()
        tg.Assimilate()
    except Exception as e:
        tg.RollBack(); print("Error: {}".format(e))

def perform_edging(state):
    try:
        w = float(state.width); g = float(state.grid)
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        toposolid = doc.GetElement(ref)
    except: return

    tg = TransactionGroup(doc, "Edging"); tg.Start()
    try:
        z_s, z_e = calculate_and_adjust_stakes(state)
        ids = List[ElementId]([toposolid.Id])
        intersector = None
        if doc.ActiveView.ViewType == ViewType.ThreeD:
            intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, doc.ActiveView)
        
        edge_offset = w / 2.0; edge_res = g * 0.5 
        curve = state.grading_line.GeometryCurve
        editor = toposolid.GetSlabShapeEditor(); editor.Enable()
        to_move, to_add = [], []
        all_verts = [v for v in editor.SlabShapeVertices]
        
        for v in all_verts: # Snap
            res = curve.Project(v.Position)
            if res and abs(flatten(v.Position).DistanceTo(flatten(res.XYZPoint)) - edge_offset) < 1.0:
                vec = (flatten(v.Position) - flatten(res.XYZPoint)).Normalize()
                exact_xy = flatten(res.XYZPoint) + (vec * edge_offset)
                exact_z = get_z_from_curve_param(curve, res, z_s, z_e)
                to_move.append((v, XYZ(exact_xy.X, exact_xy.Y, exact_z)))

        step_t = edge_res / curve.Length; step_t = 0.05 if step_t > 0.05 else step_t
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
                if is_point_on_solid(intersector, final_pt): to_add.append(final_pt)
            t_val += step_t

        t = Transaction(doc, "Apply Edging"); t.Start()
        occupied_points = [v.Position for v in all_verts]
        for item in to_move:
            try: editor.ModifySlabShapeVertex(item[0], item[1]); occupied_points.append(item[1])
            except: pass
        for pt in to_add:
            if not is_too_close(pt, occupied_points):
                try: editor.AddPoint(pt); occupied_points.append(pt)
                except: pass 
        t.Commit(); tg.Assimilate()
    except Exception as e:
        tg.RollBack(); print("Error: {}".format(e))

# ==========================================
# 7. LOOP
# ==========================================
if __name__ == '__main__':
    state = GradingState()
    while True:
        win = GradingWindow(state)
        win.ShowDialog()
        action = state.next_action
        state.next_action = None 
        if not action: break 
        elif action == "select_stakes":
            try:
                state.start_stake = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Start Stake"))
                state.end_stake = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "End Stake"))
            except: pass
        elif action == "select_line":
            try: state.grading_line = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Guide Line"))
            except: pass
        elif action == "swap": perform_swap(state)
        elif action == "sculpt": perform_sculpt(state); break 
        elif action == "edge": perform_edging(state); break
        elif action == "stitch": perform_manual_stitch(state)
        elif action == "load_recipe": perform_load_recipe(state)