# -*- coding: utf-8 -*-
"""
Finds and selects duplicate elements that are in the same location.

This script groups elements by their Category, Family, Type, and Location.
Groups with more than one element are considered duplicates.
"""

__title__ = 'Find Duplicates'
__author__ = 'Your Name'


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI import TaskDialog

from pyrevit import revit, script

# Get the current Revit document
doc = revit.doc
uidoc = revit.uidoc

# --- Functions ---
def get_element_location_key(element):
    """
    Creates a string representation of the element's location.
    For point-based elements, it uses coordinates.
    For other elements, it uses the center of the bounding box.
    """
    # Try to get the location point for point-based elements (like families)
    if hasattr(element, 'Location') and element.Location and hasattr(element.Location, 'Point'):
        point = element.Location.Point
        # Format the point to a consistent string with limited precision
        return "({:.4f}, {:.4f}, {:.4f})".format(point.X, point.Y, point.Z)

    # For elements without a location point (like walls, floors), use the bounding box center
    bounding_box = element.get_BoundingBox(None)
    if bounding_box:
        center = (bounding_box.Min + bounding_box.Max) / 2
        return "({:.4f}, {:.4f}, {:.4f})".format(center.X, center.Y, center.Z)

    # Return None if no location can be determined
    return None

# --- Main Script ---
def find_duplicates():
    """Main function to find and report duplicate elements."""
    # Dictionary to store element signatures and their IDs
    elements_dict = {}
    duplicates = []
    
    # Use a FilteredElementCollector to get all model elements
    # We exclude non-model elements like views, templates, etc.
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType().WhereElementIsViewIndependent()

    print("Analyzing model elements...")

    # Iterate over all collected elements
    for element in collector:
        # Skip elements that are not valid for this check
        if not element.Category or not hasattr(element, 'Location'):
            continue

        location_key = get_element_location_key(element)
        if not location_key:
            continue
            
        # Create a unique key for the element
        # This key combines the category, type, and location
        try:
            type_id = element.GetTypeId()
            element_type = doc.GetElement(type_id)
            type_name = element_type.Name if hasattr(element_type, 'Name') else str(type_id)
            family_name = element_type.FamilyName if hasattr(element_type, 'FamilyName') else ''
            
            key = (element.Category.Name, family_name, type_name, location_key)

        except Exception as e:
            # Some elements might not have all properties, skip them
            # print("Could not process element {}: {}".format(element.Id, e))
            continue
            
        # If the key is already in the dictionary, it's a duplicate
        if key in elements_dict:
            # Check if this is the first time we've found a duplicate for this key
            if elements_dict[key] not in duplicates:
                duplicates.append(elements_dict[key])
            duplicates.append(element.Id)
        else:
            # If not, add the element to the dictionary
            elements_dict[key] = element.Id
            
    # --- Reporting Results ---
    if duplicates:
        print("\nFound {} duplicate elements!".format(len(duplicates)))
        
        # Select the duplicates in the Revit UI
        revit.get_selection().set_to(duplicates)
        
        # Prepare a message for the user
        dialog = TaskDialog("Duplicates Found")
        dialog.MainInstruction = "Found {} duplicate elements.".format(len(duplicates))
        dialog.MainContent = ("The duplicate elements have been selected in the model. "
                             "Please review them carefully.\n\n"
                             "You can use the 'Selection Box' (BX) tool to isolate them. "
                             "It's recommended to use the 'Manage' tab > 'Select by ID' tool to inspect each one before deleting.")
        dialog.Show()
        
        # Print the IDs of the duplicates
        print("Selected Duplicate Element IDs:")
        for elem_id in sorted(list(set(duplicates))): # Use set to get unique IDs
            print("  - {}".format(elem_id))
            
    else:
        # If no duplicates are found
        print("\nNo duplicate elements found in the model. ✅")
        TaskDialog.Show("Success", "No duplicate elements were found.")

# --- Run the script ---
if __name__ == '__main__':
    # A transaction is not needed for just selecting elements
    find_duplicates()