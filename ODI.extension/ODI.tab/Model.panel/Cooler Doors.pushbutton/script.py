# -*- coding: utf-8 -*-
"""Resize curtain wall based on number of doors needed.
Uses the wall type's Vertical Grid Fixed Distance for spacing."""

__title__ = 'Resize\nCurtain Wall'
__author__ = 'Claude'
__helpurl__ = ''
__min_revit_ver__ = 2019
__max_revit_ver__ = 2024

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms
from pyrevit import script
import math

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_selected_curtain_wall():
    """Get the selected curtain wall element."""
    try:
        selection = uidoc.Selection.GetElementIds()
        if len(selection) != 1:
            forms.alert('Please select exactly one curtain wall.', exitscript=True)
        
        element = doc.GetElement(selection[0])
        if not isinstance(element, Wall) or not element.CurtainGrid:
            forms.alert('Selected element is not a curtain wall.', exitscript=True)
        
        return element
    except Exception as e:
        forms.alert('Error getting curtain wall: {}'.format(str(e)), exitscript=True)
        return None

def get_grid_spacing(wall):
    """Get the vertical grid spacing from wall type."""
    try:
        # Get wall type
        wall_type = doc.GetElement(wall.GetTypeId())
        
        # Get vertical grid spacing parameter
        grid_param = wall_type.get_Parameter(BuiltInParameter.SPACING_LENGTH_VERT)
        if not grid_param:
            forms.alert('Could not find vertical grid spacing parameter.', exitscript=True)
            return None
            
        # Convert to feet
        spacing = grid_param.AsDouble()
        return spacing
        
    except Exception as e:
        forms.alert('Error getting grid spacing: {}'.format(str(e)), exitscript=True)
        return None

def get_user_inputs():
    """Get number of doors from user."""
    try:
        count = forms.ask_for_string(
            default='10',
            prompt='Enter number of doors needed:',
            title='Door Count'
        )
        if not count:
            script.exit()
        door_count = int(count)

        maintain_center = forms.alert(
            'Maintain center position?',
            yes=True, no=True,
            title='Position Option'
        )

        # Validate inputs
        if door_count <= 0:
            forms.alert('Number of doors must be positive.', exitscript=True)
        
        return door_count, maintain_center
            
    except ValueError:
        forms.alert('Invalid input. Please enter a whole number.', exitscript=True)
        return None
    except Exception as e:
        forms.alert('Error getting user input: {}'.format(str(e)), exitscript=True)
        return None

def resize_curtain_wall(wall, door_count, maintain_center):
    """Resize the curtain wall based on number of doors and grid spacing."""
    try:
        # Get grid spacing
        unit_size = get_grid_spacing(wall)
        if not unit_size:
            return
            
        # Start transaction
        with Transaction(doc, 'Resize Curtain Wall') as t:
            t.Start()
            
            # Get wall location line
            wall_line = wall.Location.Curve
            start_point = wall_line.GetEndPoint(0)
            end_point = wall_line.GetEndPoint(1)
            
            # Calculate new length
            new_length = unit_size * door_count
            current_length = wall_line.Length
            
            # Calculate direction vector
            direction = (end_point - start_point).Normalize()
            
            if maintain_center:
                # Find middle point
                mid_point = (start_point + end_point) * 0.5
                # Calculate new start and end points from middle
                half_length = new_length * 0.5
                new_start = mid_point - direction * half_length
                new_end = mid_point + direction * half_length
            else:
                # Keep start point, only move end point
                new_start = start_point
                new_end = start_point + direction * new_length
            
            # Create new line
            new_line = Line.CreateBound(new_start, new_end)
            wall.Location.Curve = new_line
            
            t.Commit()
            
            # Show success message
            message = 'Curtain wall resized successfully!\nNew length: {:.2f} feet\nDoors: {} x {:.2f} feet spacing'.format(
                new_length, door_count, unit_size)
            TaskDialog.Show('Success', message)
            
    except Exception as e:
        forms.alert('Error resizing curtain wall: {}'.format(str(e)), exitscript=True)

def main():
    """Main script execution."""
    # Get selected curtain wall
    wall = get_selected_curtain_wall()
    if not wall:
        return
    
    # Get user inputs
    result = get_user_inputs()
    if not result:
        return
        
    door_count, maintain_center = result
    
    # Resize curtain wall
    resize_curtain_wall(wall, door_count, maintain_center)

if __name__ == '__main__':
    main()