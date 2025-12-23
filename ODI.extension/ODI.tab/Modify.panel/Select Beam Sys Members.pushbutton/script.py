# Import necessary Revit API modules and CLR for .NET interop

__title__ = "Select Framing \n in Beam System"
__author__ = "ODI"
__context__ = "active-view-type: FloorPlan,CeilingPlan,EngineeringPlan,AreaPlan,Section,Elevation,ThreeD"

import clr
clr.AddReference('System')
from System.Collections.Generic import List as NETList  # .NET List for ICollection
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BeamSystem,
    ElementId,
    FamilyInstance,
    BuiltInParameter,
)
from Autodesk.Revit.UI import UIDocument, TaskDialog
from rpw import ui  # Optional: for pyRevit UI selection; remove if not using pyRevit

# Get the active document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Minimum length threshold (in feet, Revit's internal unit); 6 inches = 0.5 feet
MIN_BEAM_LENGTH = 0.5  # Adjust this value as needed (e.g., 0.25 for 3 inches)

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

# Function to select beams in the BeamSystem, filtering out tiny members
def select_beam_system_members(beam_system):
    beam_ids = beam_system.GetBeamIds()
    if not beam_ids or len(beam_ids) == 0:
        TaskDialog.Show("Warning", "No beams found in the selected Beam System.")
        return False

    # Filter beams by length
    valid_beam_ids = NETList[ElementId]()
    tiny_beam_count = 0
    
    for beam_id in beam_ids:
        beam = doc.GetElement(beam_id)
        if isinstance(beam, FamilyInstance):
            # Get the Length parameter (BuiltInParameter.CURVE_ELEM_LENGTH)
            length_param = beam.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
            if length_param and length_param.HasValue:
                length = length_param.AsDouble()  # Length in feet (Revit's internal unit)
                if length >= MIN_BEAM_LENGTH:
                    valid_beam_ids.Add(beam_id)
                else:
                    tiny_beam_count += 1
            else:
                # If length parameter is missing, include the beam to avoid exclusion errors
                valid_beam_ids.Add(beam_id)
        else:
            # If not a FamilyInstance, include it (unlikely in a BeamSystem)
            valid_beam_ids.Add(beam_id)

    if valid_beam_ids.Count == 0:
        TaskDialog.Show("Warning", "No beams meet the minimum length threshold ({} feet).".format(MIN_BEAM_LENGTH))
        return False

    try:
        # Select only valid beams
        uidoc.Selection.SetElementIds(valid_beam_ids)
        
        # Report tiny beams excluded
        if tiny_beam_count > 0:
            TaskDialog.Show("Info", "Selected {} beams. Excluded {} tiny beams (< {} feet).".format(
                valid_beam_ids.Count, tiny_beam_count, MIN_BEAM_LENGTH))
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
        filtered_count = uidoc.Selection.GetElementIds().Count
        TaskDialog.Show("Success", "Selected {} of {} beams from the Beam System.".format(filtered_count, beam_count))
    else:
        TaskDialog.Show("Warning", "Selection failed or no beams were found.")

except Exception as e:
    TaskDialog.Show("Error", "Script failed: {}".format(str(e)))