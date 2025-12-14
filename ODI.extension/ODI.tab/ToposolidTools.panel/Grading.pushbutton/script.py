# -*- coding: utf-8 -*-
from pyrevit import forms, revit, script
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB import (
    BuiltInCategory, ModelLine, ModelCurve, CurveElement, 
    XYZ, Transaction, SubTransaction, ViewType,
    ReferenceIntersector, FindReferenceTarget, ElementId
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows import Media
from System.Collections.Generic import List
import math

# --- CONFIGURATION ---
GRADING_WIDTH = 6.0    # Width of the Road/Path
BANK_HEIGHT   = 0.0    # Banking/Crown
FALLOFF_DIST  = 8.0    # Distance to blend back to existing
STEP_SIZE     = 2.0    # Spacing for new definition points

# --- SETUP ---
doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================
class UniversalFilter(ISelectionFilter):
    def AllowElement(self, elem): return True
    def AllowReference(self, ref, point): return True

def flatten(pt):
    """Drops the Z coordinate for 2D distance checks."""
    return XYZ(pt.X, pt.Y, 0)

def clamp(val, min_val=0.0, max_val=1.0):
    return max(min_val, min(val, max_val))

def get_design_z_at_station(dist_from_start, start_z, slope):
    return start_z + (dist_from_start * slope)

def get_existing_z_at_loc(intersector, xy_point):
    """Ray traces down to find existing ground Z."""
    origin = XYZ(xy_point.X, xy_point.Y, xy_point.Z + 1000)
    try:
        context = intersector.FindNearest(origin, XYZ(0, 0, -1))
        if context: return context.GetReference().GlobalPoint.Z
    except: pass
    return None

# ==========================================
# 2. CORE LOGIC
# ==========================================

def adjust_and_grade(doc, start_stake, end_stake, line, toposolid):
    print("--- Starting Calculation ---")
    
    # A. PREP GEOMETRY
    curve = line.GeometryCurve
    start_pt = start_stake.Location.Point
    end_pt = end_stake.Location.Point
    line_length = curve.Length
    
    slope = (end_pt.Z - start_pt.Z) / line_length

    # Direction Check
    dist_start = start_pt.DistanceTo(curve.GetEndPoint(0))
    dist_end   = start_pt.DistanceTo(curve.GetEndPoint(1))
    is_reversed = dist_end < dist_start

    # Limits
    half_width = GRADING_WIDTH / 2.0
    max_influence = half_width + FALLOFF_DIST

    # Ray Tracer setup
    target_ids = List[ElementId]([toposolid.Id])
    intersector = ReferenceIntersector(target_ids, FindReferenceTarget.Element, doc.ActiveView)

    editor = toposolid.GetSlabShapeEditor()
    existing_vertices = editor.SlabShapeVertices
    
    all_points_to_add = []

    # --- PHASE 1: SCAN EXISTING POINTS (2D Logic) ---
    # Note: If the element is fresh (flat), existing_vertices might be empty.
    print("Scanning {} existing vertices...".format(existing_vertices.Size))
    count_moved = 0
    
    for v in existing_vertices:
        v_pt = v.Position
        
        # 1. Project to Curve (3D)
        result = curve.Project(v_pt)
        if not result: continue
        
        proj_pt_3d = result.XYZPoint
        
        # 2. CONVERT TO 2D DISTANCE
        dist_2d = flatten(v_pt).DistanceTo(flatten(proj_pt_3d))
        
        if dist_2d > max_influence:
            continue 
            
        # 3. Calculate Station
        raw_param = result.Parameter
        curve_pt_at_param = curve.ComputeDerivatives(raw_param, False).Origin
        dist_from_curve_start = curve_pt_at_param.DistanceTo(curve.GetEndPoint(0))
        station_dist = (line_length - dist_from_curve_start) if is_reversed else dist_from_curve_start

        # 4. Determine New Z
        center_z = get_design_z_at_station(station_dist, start_pt.Z, slope)
        bank_z   = center_z + BANK_HEIGHT
        
        new_z = None
        
        if dist_2d <= half_width:
            # ZONE A: ROAD
            new_z = bank_z
            
        elif dist_2d <= max_influence:
            # ZONE B: FALLOFF
            vec = (flatten(v_pt) - flatten(proj_pt_3d)).Normalize()
            tie_in_loc_2d = flatten(proj_pt_3d) + (vec * max_influence)
            tie_in_search_pt = XYZ(tie_in_loc_2d.X, tie_in_loc_2d.Y, 0)
            
            existing_z = get_existing_z_at_loc(intersector, tie_in_search_pt)
            
            if existing_z is not None:
                t = (dist_2d - half_width) / FALLOFF_DIST
                new_z = bank_z + (t * (existing_z - bank_z))
        
        if new_z is not None:
            all_points_to_add.append(XYZ(v_pt.X, v_pt.Y, new_z))
            count_moved += 1

    print("  -> Found {} existing points to adjust.".format(count_moved))


    # --- PHASE 2: GENERATE NEW DEFINITION POINTS ---
    print("Generating new path points...")
    current_dist = 0.0
    count_new = 0
    
    while current_dist <= line_length:
        
        raw_param = clamp(current_dist / line_length)
        param = (1.0 - raw_param) if is_reversed else raw_param
        
        transform = curve.ComputeDerivatives(param, True)
        center_loc = transform.Origin
        tangent = transform.BasisX.Normalize()
        normal = XYZ(-tangent.Y, tangent.X, 0).Normalize()
        
        z_center = start_pt.Z + (current_dist * slope)
        z_bank   = z_center + BANK_HEIGHT
        
        # Center & Banks
        all_points_to_add.append(XYZ(center_loc.X, center_loc.Y, z_center))
        all_points_to_add.append(XYZ(center_loc.X + normal.X * half_width, center_loc.Y + normal.Y * half_width, z_bank))
        all_points_to_add.append(XYZ(center_loc.X - normal.X * half_width, center_loc.Y - normal.Y * half_width, z_bank))
        
        # Tie-Ins
        pt_tie_left_loc  = center_loc + (normal * max_influence)
        pt_tie_right_loc = center_loc - (normal * max_influence)
        
        z_tie_left = get_existing_z_at_loc(intersector, pt_tie_left_loc)
        z_tie_right = get_existing_z_at_loc(intersector, pt_tie_right_loc)
        
        if z_tie_left: all_points_to_add.append(XYZ(pt_tie_left_loc.X, pt_tie_left_loc.Y, z_tie_left))
        if z_tie_right: all_points_to_add.append(XYZ(pt_tie_right_loc.X, pt_tie_right_loc.Y, z_tie_right))
        
        current_dist += STEP_SIZE
        count_new += 1

    print("  -> Generated {} profile slices.".format(count_new))

    # Lock End
    all_points_to_add.append(XYZ(end_pt.X, end_pt.Y, end_pt.Z))

    return all_points_to_add


# ==========================================
# 3. UI CLASS
# ==========================================

class GradingWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.start_stake = None
        self.end_stake = None
        self.grading_line = None
        
        if doc.ActiveView.ViewType != ViewType.ThreeD:
            self.StatusLabel.Content = "Tip: Use a 3D View."

    # --- SELECTION EVENTS ---
    def select_stakes(self, sender, args):
        self.Hide()
        try:
            ref1 = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select START Stake")
            self.start_stake = doc.GetElement(ref1)
            ref2 = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select END Stake")
            self.end_stake = doc.GetElement(ref2)
            self.update_ui()
        except OperationCanceledException: pass
        except Exception as e: forms.alert("Error: {}".format(e))
        finally: self.ShowDialog()

    def select_line(self, sender, args):
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Model Line")
            elem = doc.GetElement(ref)
            if not isinstance(elem, (ModelLine, ModelCurve, CurveElement)):
                forms.alert("Please select a Model Line.")
            else:
                self.grading_line = elem
            self.update_ui()
        except OperationCanceledException: pass
        finally: self.ShowDialog()

    def swap_stakes(self, sender, args):
        self.start_stake, self.end_stake = self.end_stake, self.start_stake
        self.update_ui()

    def update_ui(self):
        green = Media.Brushes.Green
        red = Media.Brushes.Red
        
        # Helper for ID
        def get_name_id(elem):
            return "{} [{}]".format(elem.Name, elem.Id) if elem else "[None]"

        self.StartStakeID.Text = "Start: {}".format(get_name_id(self.start_stake))
        self.StartStakeID.Foreground = green if self.start_stake else red
        
        self.EndStakeID.Text = "End: {}".format(get_name_id(self.end_stake))
        self.EndStakeID.Foreground = green if self.end_stake else red

        self.LineID.Text = "Line: {}".format(get_name_id(self.grading_line))
        self.LineID.Foreground = green if self.grading_line else red

        ready = self.start_stake and self.end_stake and self.grading_line
        self.RunBtn.IsEnabled = ready
        self.SwapBtn.IsEnabled = bool(self.start_stake and self.end_stake)
        self.StatusLabel.Content = "Ready." if ready else "Incomplete Selection."

    # --- EXECUTION ---
    def run_grading(self, sender, args):
        self.Hide()
        
        # 1. Select Toposolid
        toposolid = None
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
            toposolid = doc.GetElement(ref)
            if not hasattr(toposolid, "GetSlabShapeEditor"):
                forms.alert("Element does not support Shape Editing.")
                self.ShowDialog() 
                return
        except OperationCanceledException:
            self.ShowDialog()
            return

        # 2. Compute
        try:
            points_to_add = adjust_and_grade(
                doc, self.start_stake, self.end_stake, self.grading_line, toposolid
            )
        except Exception as e:
            forms.alert("Math Error: {}".format(e))
            self.ShowDialog()
            return

        if not points_to_add:
            forms.alert("Result: 0 points calculated.")
            self.ShowDialog()
            return

        # 3. Modify
        t = Transaction(doc, "Grade Toposolid")
        t.Start()
        st = SubTransaction(doc)
        st.Start()
        
        try:
            editor = toposolid.GetSlabShapeEditor()
            
            # --- CRITICAL FIX: ENABLE SHAPE EDITING ---
            editor.Enable() 
            
            print("Applying {} points...".format(len(points_to_add)))
            
            for pt in points_to_add:
                editor.AddPoint(pt)
                
            st.Commit()
            t.Commit()
            print("Success! Grading Complete.")
            self.Close()
            
        except Exception as e:
            st.RollBack()
            t.Commit()
            print("Grading Failed: {}".format(e))
            self.ShowDialog()

if __name__ == '__main__':
    GradingWindow().ShowDialog()