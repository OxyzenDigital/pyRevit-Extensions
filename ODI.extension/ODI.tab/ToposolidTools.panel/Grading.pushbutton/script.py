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
    XYZ, Transaction, ElementId, BuiltInParameter,
    ReferenceIntersector, FindReferenceTarget, Options, Solid, ViewType
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# 1. SETTINGS
# ==========================================
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "grading_settings.json")
MIN_DIST_TOLERANCE = 0.25 

def load_settings():
    defaults = {"width": "6.0", "falloff": "10.0", "grid": "3.0"}
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                data = json.load(f)
                defaults.update(data)
    except: pass
    return defaults

def save_settings(width, falloff, grid):
    data = {"width": str(width), "falloff": str(falloff), "grid": str(grid)}
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
        self.next_action = None

    @property
    def ready_to_sculpt(self):
        return bool(self.start_stake and self.end_stake and self.grading_line)

    @property
    def ready_to_edge(self):
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
        def fmt(e): return "{} [{}]".format(e.Name, e.Id) if e else "[None]"

        self.Lb_StartStake.Text = "Start: {}".format(fmt(self.state.start_stake))
        self.Lb_StartStake.Foreground = green if self.state.start_stake else red
        self.Lb_EndStake.Text = "End: {}".format(fmt(self.state.end_stake))
        self.Lb_EndStake.Foreground = green if self.state.end_stake else red
        self.Lb_Line.Text = "Line: {}".format(fmt(self.state.grading_line))
        self.Lb_Line.Foreground = green if self.state.grading_line else red

        self.Tb_Width.Text = self.state.width
        self.Tb_Falloff.Text = self.state.falloff
        self.Tb_Grid.Text = self.state.grid

        self.Btn_Run.IsEnabled = self.state.ready_to_sculpt
        self.Btn_Edging.IsEnabled = self.state.ready_to_edge
        self.Btn_Swap.IsEnabled = bool(self.state.start_stake and self.state.end_stake)

        if self.state.ready_to_sculpt:
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

    def save(self):
        self.state.width = self.Tb_Width.Text
        self.state.falloff = self.Tb_Falloff.Text
        self.state.grid = self.Tb_Grid.Text
        save_settings(self.state.width, self.state.falloff, self.state.grid)

    def a_stakes(self, s, a): self.save(); self.state.next_action = "select_stakes"; self.Close()
    def a_line(self, s, a): self.save(); self.state.next_action = "select_line"; self.Close()
    def a_swap(self, s, a): self.save(); self.state.next_action = "swap"; self.Close()
    def a_run(self, s, a): self.save(); self.state.next_action = "sculpt"; self.Close()
    def a_edge(self, s, a): self.save(); self.state.next_action = "edge"; self.Close()

# ==========================================
# 3. GEOMETRY HELPERS
# ==========================================
class UniversalFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, r, p): return True

def flatten(pt): return XYZ(pt.X, pt.Y, 0)
def lerp(a, b, t): return a + t * (b - a)

def get_surface_z(intersector, pt):
    """
    Casts a ray to find the EXACT height of the terrain at 'pt'.
    Crucial for initializing new points correctly.
    """
    if not intersector: return None
    # Start high above (2000ft) and shoot down
    origin = XYZ(pt.X, pt.Y, pt.Z + 2000.0) 
    try:
        context = intersector.FindNearest(origin, XYZ(0, 0, -1))
        if context: 
            return context.GetReference().GlobalPoint.Z
    except: pass
    return None

def is_too_close(candidate_pt, occupied_points):
    cand_flat = flatten(candidate_pt)
    for existing in occupied_points:
        if cand_flat.DistanceTo(flatten(existing)) < MIN_DIST_TOLERANCE:
            return True
    return False

def get_z_from_curve_param(curve, projected_result, z_start, z_end):
    p_min = curve.GetEndParameter(0)
    p_max = curve.GetEndParameter(1)
    p_range = p_max - p_min
    
    raw_p = projected_result.Parameter
    norm_p = (raw_p - p_min) / p_range
    norm_p = max(0.0, min(1.0, norm_p))
    
    return z_start + (norm_p * (z_end - z_start))

# ==========================================
# 4. ACTIONS
# ==========================================
def perform_swap(state):
    t = Transaction(doc, "Swap Heights")
    t.Start()
    try:
        pid = BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM
        p1 = state.start_stake.get_Parameter(pid)
        p2 = state.end_stake.get_Parameter(pid)
        if p1 and p2:
            v1, v2 = p1.AsDouble(), p2.AsDouble()
            p1.Set(v2); p2.Set(v1)
        t.Commit()
        state.start_stake, state.end_stake = state.end_stake, state.start_stake
    except: t.RollBack()

def perform_sculpt(state):
    try:
        w = float(state.width); f = float(state.falloff); g = float(state.grid)
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        toposolid = doc.GetElement(ref)
    except: return

    # Ray Tracer
    ids = List[ElementId]([toposolid.Id])
    intersector = None
    if doc.ActiveView.ViewType == ViewType.ThreeD:
        intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, doc.ActiveView)

    curve = state.grading_line.GeometryCurve
    l_start = curve.GetEndPoint(0)
    
    u_start = state.start_stake.Location.Point
    u_end = state.end_stake.Location.Point
    
    if u_start.DistanceTo(l_start) < u_end.DistanceTo(l_start):
        z_start, z_end = u_start.Z, u_end.Z
    else:
        z_start, z_end = u_end.Z, u_start.Z
    
    core_rad = w / 2.0
    total_rad = core_rad + f
    
    # 1. DENSIFY (With Z-Sampling)
    t1 = Transaction(doc, "Densify")
    t1.Start()
    editor = toposolid.GetSlabShapeEditor()
    editor.Enable()
    
    occupied_points = [v.Position for v in editor.SlabShapeVertices]
    
    # Grid Bounds
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
            t_pt = XYZ(x, y, u_start.Z) 
            
            res = curve.Project(t_pt)
            if res:
                d = flatten(t_pt).DistanceTo(flatten(res.XYZPoint))
                if d < (total_rad + g):
                    # SAMPLING Z (The Fix)
                    # Instead of assuming 'u_start.Z', we ask the Toposolid
                    # "What is your height here right now?"
                    # This ensures the new point starts flush with the existing ground.
                    
                    real_z = get_surface_z(intersector, t_pt)
                    if real_z is not None:
                        grid_pts.append(XYZ(x, y, real_z))
            y += g
        x += g
        
    # Add points
    added_count = 0
    for p in grid_pts:
        if not is_too_close(p, occupied_points):
            try: 
                editor.AddPoint(p)
                occupied_points.append(p)
                added_count += 1
            except: pass
    t1.Commit()
    print("Densify: Added {} points at Surface Z".format(added_count))
    
    # 2. SCULPT (With Smooth Falloff)
    t2 = Transaction(doc, "Sculpt")
    t2.Start()
    updates = []
    
    for v in editor.SlabShapeVertices:
        res = curve.Project(v.Position)
        if not res: continue
        
        d = flatten(v.Position).DistanceTo(flatten(res.XYZPoint))
        if d > total_rad: continue
        
        # Target Road Z
        target_z = get_z_from_curve_param(curve, res, z_start, z_end)
        
        current_z = v.Position.Z
        
        # Calculate New Z
        if d <= core_rad:
            new_z = target_z
        else:
            # S-Curve / SmoothStep Blending
            # Linear t (0.0 at core edge, 1.0 at falloff edge)
            t = (d - core_rad) / f
            if t > 1.0: t = 1.0
            
            # SmoothStep Formula: t * t * (3 - 2 * t)
            # This creates an ease-in / ease-out curve
            smooth_t = t * t * (3 - 2 * t)
            
            # Lerp using Smooth T
            new_z = lerp(target_z, current_z, smooth_t)
            
        if abs(new_z - current_z) > 0.005:
            updates.append(XYZ(v.Position.X, v.Position.Y, new_z))
            
    for p in updates:
        try: editor.AddPoint(p)
        except: pass
    t2.Commit()


def perform_edging(state):
    try:
        w = float(state.width); g = float(state.grid)
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        toposolid = doc.GetElement(ref)
    except: return

    ids = List[ElementId]([toposolid.Id])
    intersector = None
    if doc.ActiveView.ViewType == ViewType.ThreeD:
        intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, doc.ActiveView)

    edge_offset = w / 2.0
    edge_res = g * 0.5 
    
    curve = state.grading_line.GeometryCurve
    l_start = curve.GetEndPoint(0)
    
    u_start = state.start_stake.Location.Point
    u_end = state.end_stake.Location.Point
    if u_start.DistanceTo(l_start) < u_end.DistanceTo(l_start):
        z_start, z_end = u_start.Z, u_end.Z
    else:
        z_start, z_end = u_end.Z, u_start.Z

    editor = toposolid.GetSlabShapeEditor()
    editor.Enable()
    
    bb = state.grading_line.get_BoundingBox(None)
    buffer = w + 5.0
    min_b = bb.Min - XYZ(buffer, buffer, 0)
    max_b = bb.Max + XYZ(buffer, buffer, 0)
    
    to_move = [] 
    to_add = []  
    all_verts = [v for v in editor.SlabShapeVertices]
    
    # 1. SNAP EXISTING
    for v in all_verts:
        p = v.Position
        if not (min_b.X <= p.X <= max_b.X and min_b.Y <= p.Y <= max_b.Y): continue
        
        res = curve.Project(p)
        if not res: continue
        
        d = flatten(p).DistanceTo(flatten(res.XYZPoint))
        if abs(d - edge_offset) < 1.0:
            vec = (flatten(p) - flatten(res.XYZPoint)).Normalize()
            exact_xy = flatten(res.XYZPoint) + (vec * edge_offset)
            
            exact_z = get_z_from_curve_param(curve, res, z_start, z_end)
            
            to_move.append((v, XYZ(exact_xy.X, exact_xy.Y, exact_z)))

    # 2. GENERATE NEW POINTS
    length = curve.Length
    step_t = edge_res / length
    if step_t > 0.05: step_t = 0.05
    
    t_val = 0.0
    while t_val <= 1.001:
        eval_t = max(0.0, min(1.0, t_val))
        
        center_pt = curve.Evaluate(eval_t, True)
        deriv = curve.ComputeDerivatives(eval_t, True)
        tangent = deriv.BasisX.Normalize()
        normal = tangent.CrossProduct(XYZ.BasisZ)
        
        road_z = z_start + (eval_t * (z_end - z_start))
        
        sides = [1.0, -1.0]
        for side in sides:
            offset_vec = normal * (side * edge_offset)
            final_pt = center_pt + offset_vec
            final_pt = XYZ(final_pt.X, final_pt.Y, road_z)
            
            # Check Z of existing surface to ensure we don't dive into void
            # Actually Edging is strict, so we force road_z.
            # But we must check if XY is valid.
            real_z = get_surface_z(intersector, final_pt)
            if real_z is None: continue # Void

            to_add.append(final_pt)
            
        t_val += step_t

    # 3. EXECUTE
    t = Transaction(doc, "Apply Edging")
    t.Start()
    occupied_points = [v.Position for v in all_verts]
    
    for item in to_move:
        try: editor.ModifySlabShapeVertex(item[0], item[1])
        except: pass
        occupied_points.append(item[1])

    for pt in to_add:
        if not is_too_close(pt, occupied_points):
            try: 
                editor.AddPoint(pt)
                occupied_points.append(pt)
            except: pass 
            
    t.Commit()
    print("Edging: Snapped {} | Added {}".format(len(to_move), len(to_add)))

# ==========================================
# 5. LOOP
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