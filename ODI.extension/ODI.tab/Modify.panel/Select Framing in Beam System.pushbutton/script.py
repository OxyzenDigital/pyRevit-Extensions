# Import necessary Revit API modules and CLR for .NET interop
import clr
clr.AddReference('System')
from System.Collections.Generic import List as NETList  # .NET List for ICollection
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BeamSystem,
    ElementId,
)
from Autodesk.Revit.UI import UIDocument, TaskDialog
from rpw import ui  # Optional: for pyRevit UI selection; remove if not using pyRevit

# Get the active document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

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

# Function to select beams in the BeamSystem
def select_beam_system_members(beam_system):
    beam_ids = beam_system.GetBeamIds()
    if not beam_ids or len(beam_ids) == 0:
        TaskDialog.Show("Warning", "No beams found in the selected Beam System.")
        return False

    # Convert to .NET List<ElementId> for selection
    beam_id_list = NETList[ElementId](beam_ids)

    try:
        # Set the selection in the Revit UI
        uidoc.Selection.SetElementIds(beam_id_list)
        return True
    except Exception as e:
        TaskDialog.Show("Error", "Failed to select beams: {}".format(str(e)))
        return False

# Main execution
try:
    # Get the selected BeamSystem
    beam_system = get_selected_beam_system()
    if not beam_system:
        raise Exception("No valid Beam System selected.")

    # Select the beams in the UI
    success = select_beam_system_members(beam_system)
    if success:
        beam_count = len(beam_system.GetBeamIds())
        TaskDialog.Show("Success", "Selected {} beams from the Beam System.".format(beam_count))
    else:
        TaskDialog.Show("Warning", "Selection failed or no beams were found.")

except Exception as e:
    TaskDialog.Show("Error", "Script failed: {}".format(str(e)))