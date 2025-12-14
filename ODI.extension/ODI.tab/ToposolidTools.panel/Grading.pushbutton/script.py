# -*- coding: utf-8 -*-
from pyrevit import forms, revit, script
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB import (
    BuiltInCategory, ModelLine, ModelCurve, CurveElement, 
    XYZ, Transaction, SubTransaction, ViewType
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows import Media

# --- SETUP ---
doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# 1. PERMISSIVE FILTERS (Let you click, then we check)
# ==========================================

class UniversalFilter(ISelectionFilter):
    """Allows selecting ANYTHING so we can diagnose the category."""
    def AllowElement(self, elem):
        return True
    def AllowReference(self, ref, point):
        return True

# ==========================================
# 2. MATH UTILITIES
# ==========================================
def calculate_spade_points(start_pt, end_pt, curve):
    """Generates the Spade/Swale profile points."""
    points = []
    line_length = curve.Length
    
    # HARDCODED SETTINGS (Adjust if needed)
    width = 6.0
    bank_height = 0.0
    step_size = 2.0
    offset = width / 2.0

    # Direction Check
    dist_start = start_pt.DistanceTo(curve.GetEndPoint(0))
    dist_end   = start_pt.DistanceTo(curve.GetEndPoint(1))
    is_reversed = dist_end < dist_start

    # Slope
    slope = (end_pt.Z - start_pt.Z) / line_length
    
    current_dist = 0.0
    while current_dist <= line_length:
        # Parameter
        raw_param = current_dist / line_length
        param = (1.0 - raw_param) if is_reversed else raw_param
        
        # Geometry
        transform = curve.ComputeDerivatives(param, True)
        center_loc = transform.Origin
        tangent = transform.BasisX.Normalize()
        normal = XYZ(-tangent.Y, tangent.X, 0).Normalize()
        
        # Elevations
        z_center = start_pt.Z + (current_dist * slope)
        z_bank   = z_center + bank_height
        
        # Create 3 Points
        points.append(XYZ(center_loc.X, center_loc.Y, z_center))
        points.append(XYZ(center_loc.X + normal.X * offset, center_loc.Y + normal.Y * offset, z_bank))
        points.append(XYZ(center_loc.X - normal.X * offset, center_loc.Y - normal.Y * offset, z_bank))
        
        current_dist += step_size
        
    # Final Point (Center)
    points.append(XYZ(end_pt.X, end_pt.Y, end_pt.Z))
    return points

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
            self.StatusLabel.Content = "Tip: Please use a 3D View."

    # --- SELECT STAKES (Diagnostic Mode) ---
    def select_stakes(self, sender, args):
        self.Hide()
        try:
            # 1. START STAKE
            ref1 = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Click the START Stake")
            elem1 = doc.GetElement(ref1)
            
            # VALIDATION 1
            # We verify if it's strictly Site or Generic Model
            cat_name = elem1.Category.Name if elem1.Category else "Unknown"
            if cat_name not in ["Site", "Generic Models"]:
                forms.alert("You selected a '{}'.\nPlease select a Site or Generic Model element.".format(cat_name))
                return # Stop here if wrong
            
            self.start_stake = elem1

            # 2. END STAKE
            ref2 = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Click the END Stake")
            elem2 = doc.GetElement(ref2)
            
            # VALIDATION 2
            cat_name2 = elem2.Category.Name if elem2.Category else "Unknown"
            if cat_name2 not in ["Site", "Generic Models"]:
                forms.alert("You selected a '{}'.\nPlease select a Site or Generic Model element.".format(cat_name2))
                self.start_stake = None # Reset
                return

            self.end_stake = elem2
            self.update_ui()
            
        except OperationCanceledException: 
            pass
        except Exception as e: 
            forms.alert("Error: {}".format(e))
        finally: 
            self.ShowDialog()

    # --- SELECT LINE (Diagnostic Mode) ---
    def select_line(self, sender, args):
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select the Model Line")
            elem = doc.GetElement(ref)
            
            # VALIDATION: Is it a line?
            if not isinstance(elem, (ModelLine, ModelCurve, CurveElement)):
                forms.alert("You selected a '{}'.\nPlease select a Model Line or Spline.".format(elem.Category.Name))
                return
                
            self.grading_line = elem
            self.update_ui()

        except OperationCanceledException: 
            pass
        finally: 
            self.ShowDialog()

    # --- UI UTILS ---
    def swap_stakes(self, sender, args):
        self.start_stake, self.end_stake = self.end_stake, self.start_stake
        self.update_ui()

    def update_ui(self):
        # Update Text Colors to Green so you know it worked
        if self.start_stake: 
            self.StartStakeID.Text = "Start: {}".format(self.start_stake.Name)
            self.StartStakeID.Foreground = Media.Brushes.Green
        else:
            self.StartStakeID.Text = "Start: [None]"
            self.StartStakeID.Foreground = Media.Brushes.Red
            
        if self.end_stake: 
            self.EndStakeID.Text = "End: {}".format(self.end_stake.Name)
            self.EndStakeID.Foreground = Media.Brushes.Green
        else:
            self.EndStakeID.Text = "End: [None]"
            self.EndStakeID.Foreground = Media.Brushes.Red

        if self.grading_line: 
            self.LineID.Text = "Line: Selected"
            self.LineID.Foreground = Media.Brushes.Green
        else:
            self.LineID.Text = "Line: [None]"
            self.LineID.Foreground = Media.Brushes.Red

        ready = self.start_stake and self.end_stake and self.grading_line
        self.RunBtn.IsEnabled = ready
        self.SwapBtn.IsEnabled = bool(self.start_stake and self.end_stake)
        self.StatusLabel.Content = "Ready." if ready else "Incomplete Selection."

    # --- RUN GRADING ---
    def run_grading(self, sender, args):
        self.Close()
        
        # 1. Calc
        try:
            points = calculate_spade_points(
                self.start_stake.Location.Point,
                self.end_stake.Location.Point,
                self.grading_line.GeometryCurve
            )
        except Exception as e:
            forms.alert("Math Error: {}".format(e))
            return
        
        # 2. Select Toposolid (Diagnostic)
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid/Floor to Grade")
            toposolid = doc.GetElement(ref)
            
            if not hasattr(toposolid, "GetSlabShapeEditor"):
                forms.alert("Selected element '{}' does not support Shape Editing.".format(toposolid.Category.Name))
                return
        except OperationCanceledException: 
            return

        # 3. Modify
        t = Transaction(doc, "Carve Toposolid")
        t.Start()
        st = SubTransaction(doc)
        st.Start()
        
        try:
            editor = toposolid.GetSlabShapeEditor()
            for pt in points:
                editor.AddPoint(pt)
            st.Commit()
            t.Commit()
            print("Success! Added {} points.".format(len(points)))
        except Exception as e:
            st.RollBack()
            t.Commit()
            print("Grading Failed: {}".format(e))

if __name__ == '__main__':
    GradingWindow().ShowDialog()