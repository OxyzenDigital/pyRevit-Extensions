# -*- coding: utf-8 -*-
from pyrevit import forms, revit, script
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB import (
    BuiltInCategory, ModelLine, ModelCurve, XYZ, 
    Transaction, SubTransaction, ViewType
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows import Media

# --- SETUP ---
doc = revit.doc
uidoc = revit.uidoc

# --- 1. ROBUST FILTERS ---
# We make these filters permissive to avoid blocking valid elements.

class AnyElementFilter(ISelectionFilter):
    """Allows almost any model element to be selected to prevent blocking."""
    def AllowElement(self, elem):
        if not elem.Category: return False
        return True
    def AllowReference(self, ref, point):
        return True

class LineFilter(ISelectionFilter):
    """Strictly allows Model Lines/Curves."""
    def AllowElement(self, elem):
        if isinstance(elem, (ModelLine, ModelCurve)):
            return True
        return False
    def AllowReference(self, ref, point):
        return True


# --- 2. MAIN WINDOW ---

class GradingWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.start_stake = None
        self.end_stake = None
        self.grading_line = None
        
        # Initial check for 3D View (Best practice for grading)
        if doc.ActiveView.ViewType != ViewType.ThreeD:
            self.StatusLabel.Content = "Tip: 3D View is recommended."
            self.StatusLabel.Foreground = Media.Brushes.Gray

    # --- UI: SELECT STAKES ---
    def select_stakes(self, sender, args):
        self.Hide()
        try:
            # We use AnyElementFilter so you can click Generic Models, Site, Topo, anything.
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element, 
                AnyElementFilter(), 
                "Select exactly 2 Stakes (Start & End)"
            )
            
            if len(refs) != 2:
                self.StatusLabel.Content = "Selected {} items. Please select exactly 2.".format(len(refs))
                self.StatusLabel.Foreground = Media.Brushes.Red
            else:
                self.start_stake = doc.GetElement(refs[0])
                self.end_stake = doc.GetElement(refs[1])
                self.update_ui()
                
        except OperationCanceledException:
            pass 
        except Exception as e:
            self.StatusLabel.Content = "Error: {}".format(e)
        finally:
            self.ShowDialog()

    # --- UI: SELECT LINE ---
    def select_line(self, sender, args):
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, 
                LineFilter(), 
                "Select the Grading Model Line"
            )
            self.grading_line = doc.GetElement(ref)
            self.update_ui()

        except OperationCanceledException:
            pass
        except Exception as e:
            self.StatusLabel.Content = "Error: {}".format(e)
        finally:
            self.ShowDialog()

    # --- UI: UTILS ---
    def swap_stakes(self, sender, args):
        self.start_stake, self.end_stake = self.end_stake, self.start_stake
        self.update_ui()

    def update_ui(self):
        # Update Text
        if self.start_stake:
            self.StartStakeID.Text = "Start: {} [{}]".format(self.start_stake.Name, self.start_stake.Category.Name)
            self.StartStakeID.Foreground = Media.Brushes.Black
        if self.end_stake:
            self.EndStakeID.Text = "End: {} [{}]".format(self.end_stake.Name, self.end_stake.Category.Name)
            self.EndStakeID.Foreground = Media.Brushes.Black
        if self.grading_line:
            self.LineID.Text = "Line: Model Line"
            self.LineID.Foreground = Media.Brushes.Black

        # Update Buttons
        ready = self.start_stake and self.end_stake and self.grading_line
        self.SwapBtn.IsEnabled = bool(self.start_stake and self.end_stake)
        self.RunBtn.IsEnabled = ready
        
        if ready:
            self.StatusLabel.Content = "Ready. Click Calculate to pick Toposolid."
            self.StatusLabel.Foreground = Media.Brushes.Green
        else:
            self.StatusLabel.Content = "Incomplete Selection."


    # --- 3. EXECUTE LOGIC ---
    def run_grading(self, sender, args):
        # 1. Close the UI first so user can interact with Revit
        self.Close()
        
        # 2. CALCULATE MATH
        try:
            start_pt = self.start_stake.Location.Point
            end_pt   = self.end_stake.Location.Point
            curve = self.grading_line.GeometryCurve
            line_length = curve.Length

            # Direction Check
            dist_start = start_pt.DistanceTo(curve.GetEndPoint(0))
            dist_end   = start_pt.DistanceTo(curve.GetEndPoint(1))
            is_reversed = dist_end < dist_start

            # Interpolation
            total_rise = end_pt.Z - start_pt.Z
            slope = total_rise / line_length
            step_size = 3.0 
            current_dist = 0.0
            grading_points = []

            while current_dist <= line_length:
                param = (line_length - current_dist) / line_length if is_reversed else current_dist / line_length
                pt_on_curve = curve.Evaluate(param, True)
                new_z = start_pt.Z + (current_dist * slope)
                grading_points.append(XYZ(pt_on_curve.X, pt_on_curve.Y, new_z))
                current_dist += step_size
            
            # Add final point explicitly
            grading_points.append(XYZ(end_pt.X, end_pt.Y, end_pt.Z))
        
        except Exception as e:
            print("Math Error: Could not calculate points. Check stake placement.")
            print(e)
            return


        # 3. MANUAL TOPOSOLID SELECTION
        target_elem = None
        try:
            # We use AnyElementFilter to allow clicking ANYTHING.
            # We will validate if it's a Toposolid *after* selection.
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, 
                AnyElementFilter(), 
                "Select the Toposolid/Floor to Grade"
            )
            target_elem = doc.GetElement(ref)
        
        except OperationCanceledException:
            print("Grading Cancelled by user.")
            return

        
        # 4. EXECUTE TRANSACTION
        # Validate capability first
        if not hasattr(target_elem, "GetSlabShapeEditor"):
            print("Error: The selected element '{}' does not support Shape Editing.".format(target_elem.Category.Name))
            print("Please select a Toposolid or Floor.")
            return

        t = Transaction(doc, "ODI Grading")
        t.Start()
        
        st = SubTransaction(doc)
        st.Start()
        
        try:
            editor = target_elem.GetSlabShapeEditor()
            
            # Apply Points
            for pt in grading_points:
                editor.AddPoint(pt)
            
            st.Commit()
            t.Commit()
            print("Success! Applied {} points to element {}.".format(len(grading_points), target_elem.Id))
            
        except Exception as e:
            st.RollBack()
            t.Commit()
            print("Grading Failed: Geometry Error.")
            print("The points might be overlapping existing points or folding the surface.")
            print("Technical Detail: {}".format(e))

# --- LAUNCH ---
if __name__ == '__main__':
    GradingWindow().ShowDialog()