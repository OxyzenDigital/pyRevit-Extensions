# Import necessary Revit API modules and CLR for .NET interop
import clr
clr.AddReference('System')
from System.Collections.Generic import List as NETList  # .NET List for ICollection
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BeamSystem,
    ElementId,
    Transaction,
    View,
)
from Autodesk.Revit.UI import UIDocument, TaskDialog
from rpw import ui  # Optional: for pyRevit UI selection; remove if not using pyRevit

# Get the active document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Safety check: Ensure an active view exists
active_view = doc.ActiveView
if not active_view:
    TaskDialog.Show("Error", "No active view found. Please open a view and try again.")
    raise Exception("No active view available.")

# Function to get selected BeamSystem
def get_selected_beam_system():
    try:
        # Use pyRevit's selection tool (if using pyRevit)
        selected = ui.Pick.pick_element(msg="Select a Beam System", multiple=False)
        element = doc.GetElement(selected.ElementId)
        
        # Verify it's a BeamSystem
        if not isinstance(element, BeamSystem):
            TaskDialog.Show("Error", "Selected element is not a Beam System.")
            return None
        return element
    except Exception as e:
        TaskDialog.Show("Error", "Please select a Beam System.")
        return None

# Function to toggle visibility of beams
def toggle_beam_visibility(beam_system, view):
    beam_ids = beam_system.GetBeamIds()
    if not beam_ids or len(beam_ids) == 0:
        TaskDialog.Show("Warning", "No beams found in the selected Beam System.")
        return False

    # Convert Python list to .NET List<ElementId> (ICollection)
    beam_id_list = NETList[ElementId](beam_ids)

    with Transaction(doc, "Toggle Beam Visibility") as t:
        t.Start()
        try:
            # Check if the first beam is hidden in the view
            first_beam = doc.GetElement(beam_ids[0])
            is_hidden = first_beam.IsHidden(view)
            
            # Toggle visibility
            if is_hidden:
                # Unhide all beams
                view.UnhideElements(beam_id_list)
                t.Commit()
                return True
            else:
                # Hide all beams
                view.HideElements(beam_id_list)
                t.Commit()
                return True
        except Exception as e:
            t.RollBack()
            TaskDialog.Show("Error", "Failed to toggle visibility: {}".format(str(e)))
            return False

# Main execution
try:
    # Get the selected BeamSystem
    beam_system = get_selected_beam_system()
    if not beam_system:
        raise Exception("No valid Beam System selected.")

    # Toggle visibility in the active view
    success = toggle_beam_visibility(beam_system, active_view)
    if success:
        beam_count = len(beam_system.GetBeamIds())
        first_beam = doc.GetElement(beam_system.GetBeamIds()[0])
        is_hidden = first_beam.IsHidden(active_view)
        status = "hidden" if is_hidden else "unhidden"
        TaskDialog.Show("Success", "Toggled visibility of {} beams to {}.".format(beam_count, status))
    else:
        TaskDialog.Show("Warning", "Visibility toggle failed or no beams were affected.")

except Exception as e:
    TaskDialog.Show("Error", "Script failed: {}".format(str(e)))