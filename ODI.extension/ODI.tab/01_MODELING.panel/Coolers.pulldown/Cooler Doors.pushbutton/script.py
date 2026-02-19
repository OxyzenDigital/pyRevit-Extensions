# -*- coding: utf-8 -*-
"""Convert and resize curtain wall to Cooler Doors.
Changes Wall Type to 'Cooler Doors' and sets grid spacing based on selection."""

__title__ = 'Cooler\nDoors'
__author__ = 'Claude'
__helpurl__ = ''
__min_revit_ver__ = 2019
__max_revit_ver__ = 2024

def __context__(context):
    from Autodesk.Revit.DB import ViewType
    if not context.doc or context.doc.IsFamilyDocument:
        return False
    view = context.doc.ActiveView
    if not view:
        return False
    allowed_types = [
        ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.EngineeringPlan,
        ViewType.AreaPlan, ViewType.ThreeD, ViewType.Section, ViewType.Elevation
    ]
    return view.ViewType in allowed_types

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import WallType, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import forms
from pyrevit import script
import traceback

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Dictionary of Standard Sizes: "Label": (Width_in_Feet, Height_in_Feet)
DOOR_SIZES = {
    "30\" Standard (2.5')": (2.5, 7.0),
    "24\" Narrow (2.0')": (2.0, 7.0),
    "36\" Wide (3.0')": (3.0, 7.0),
    "Custom 4'": (4.0, 8.0)
}

def get_curtain_wall():
    """Get the selected curtain wall element or pick one."""
    selection = uidoc.Selection.GetElementIds()
    
    # Check pre-selection
    if len(selection) == 1:
        el = doc.GetElement(selection[0])
        if isinstance(el, Wall) and el.CurtainGrid:
            return el
            
    # If no valid pre-selection, prompt to pick
    try:
        with forms.WarningBar(title="Select a Curtain Wall to convert"):
            ref = uidoc.Selection.PickObject(ObjectType.Element, "Select a Curtain Wall to convert")
            el = doc.GetElement(ref)
            if isinstance(el, Wall) and el.CurtainGrid:
                return el
            else:
                forms.alert('Selected element is not a curtain wall.', exitscript=True)
                return None
    except OperationCanceledException:
        return None
    except Exception as e:
        forms.alert('Error selecting wall: {}'.format(str(e)), exitscript=True)
        return None

def get_user_inputs(wall):
    """Get door size and count from user."""
    # 1. Select Size
    selected_size_name = forms.CommandSwitchWindow.show(
        sorted(DOOR_SIZES.keys()),
        message='Select Cooler Door Size:',
    )
    if not selected_size_name:
        return None
        
    door_width, door_height = DOOR_SIZES[selected_size_name]

    # Calculate suggested count based on wall length
    current_length = wall.Location.Curve.Length
    suggested_count = int(round(current_length / door_width))
    if suggested_count < 1: suggested_count = 1
    
    # 2. Select Count
    try:
        count = forms.ask_for_string(
            default=str(suggested_count),
            prompt='Enter number of doors (Calculated from length):',
            title='Door Count'
        )
        if not count:
            return None
        door_count = int(count)
        if door_count <= 0:
            forms.alert('Count must be positive.')
            return None
            
        # 3. Position
        maintain_center = forms.alert(
            'Maintain center position?',
            yes=True, no=True,
            title='Position Option'
        )
        
        return door_width, door_height, door_count, maintain_center
    except ValueError:
        forms.alert('Invalid number.')
        return None

def setup_cooler_type(base_wall, door_width, door_height):
    """Finds or creates 'Cooler Doors' WallType and updates spacing."""
    target_name = "Cooler Doors - {:.0f}\" x {:.0f}\"".format(door_width * 12, door_height * 12)
    
    # Find existing WallType
    wall_types = FilteredElementCollector(doc).OfClass(WallType).ToElements()
    target_type = None
    for wt in wall_types:
        p_name = wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p_name and p_name.AsString() == target_name:
            target_type = wt
            break
            
    # Create if missing by duplicating the selected wall's type
    if not target_type:
        base_wall_type = doc.GetElement(base_wall.GetTypeId())
        if not base_wall_type:
            forms.alert("Could not determine Wall Type.", exitscript=True)
            return None
        target_type = base_wall_type.Duplicate(target_name)
        
    # Update Parameters on the WallType
    # SPACING_LAYOUT_VERT: 1 = Fixed Distance
    layout_param_v = target_type.get_Parameter(BuiltInParameter.SPACING_LAYOUT_VERT)
    if layout_param_v and not layout_param_v.IsReadOnly:
        layout_param_v.Set(1) # 1 = Fixed Distance
        
    # SPACING_LENGTH_VERT: Door Width
    spacing_param_v = target_type.get_Parameter(BuiltInParameter.SPACING_LENGTH_VERT)
    if spacing_param_v and not spacing_param_v.IsReadOnly:
        spacing_param_v.Set(door_width)

    # SPACING_LAYOUT_HORIZ: 1 = Fixed Distance
    layout_param_h = target_type.get_Parameter(BuiltInParameter.SPACING_LAYOUT_HORIZ)
    if layout_param_h and not layout_param_h.IsReadOnly:
        layout_param_h.Set(1) # 1 = Fixed Distance

    # SPACING_LENGTH_HORIZ: Door Height
    spacing_param_h = target_type.get_Parameter(BuiltInParameter.SPACING_LENGTH_HORIZ)
    if spacing_param_h and not spacing_param_h.IsReadOnly:
        spacing_param_h.Set(door_height)
        
    return target_type

def get_mullion_type_width(m_type):
    """Calculates total width of a MullionType."""
    if not m_type: return 0.0
    
    width = 0.0
    
    # Rectangular Mullions have Width 1 and Width 2
    # Use getattr to safely access BuiltInParameter to avoid AttributeErrors in some environments
    
    bip_w1 = getattr(BuiltInParameter, "MULLION_WIDTH1", None)
    bip_w2 = getattr(BuiltInParameter, "MULLION_WIDTH2", None)
    
    if bip_w1 and bip_w2:
        p1 = m_type.get_Parameter(bip_w1)
        p2 = m_type.get_Parameter(bip_w2)
        if p1: width += p1.AsDouble()
        if p2: width += p2.AsDouble()
    
    # Fallback: Circular Mullions (Radius * 2)
    if width <= 0.001:
        bip_rad = getattr(BuiltInParameter, "MULLION_RADIUS", None)
        if bip_rad:
            p_rad = m_type.get_Parameter(bip_rad)
            if p_rad: width = p_rad.AsDouble() * 2.0
            
    # Fallback: Thickness (Depth) - often used for Corner Mullions or if Width is missing
    if width <= 0.001:
        bip_th = getattr(BuiltInParameter, "MULLION_THICKNESS", None)
        if bip_th:
            p_th = m_type.get_Parameter(bip_th)
            if p_th: width = p_th.AsDouble()
    
    return width

def get_assigned_mullion_id(wall_type, param_name):
    """Gets the ElementId of the assigned mullion."""
    bip = getattr(BuiltInParameter, param_name, None)
    if not bip: return ElementId.InvalidElementId
    
    p = wall_type.get_Parameter(bip)
    if not p: return ElementId.InvalidElementId
    
    return p.AsElementId()

def get_existing_mullion_type(doc, wall):
    """Finds the first vertical mullion type used in the wall instance."""
    grid = wall.CurtainGrid
    if not grid: return None
    
    m_ids = grid.GetMullionIds()
    if not m_ids: return None
    
    for mid in m_ids:
        m = doc.GetElement(mid)
        if not m: continue
        
        # Check orientation (Vertical)
        curve = m.Location.Curve
        if isinstance(curve, Line):
            direction = curve.Direction
            if abs(direction.Z) > 0.9: 
                return doc.GetElement(m.GetTypeId())
    
    # Fallback: Return first found
    if m_ids:
        m = doc.GetElement(m_ids[0])
        return doc.GetElement(m.GetTypeId())
        
    return None

def main():
    wall = get_curtain_wall()
    if not wall:
        return
        
    if wall.Pinned:
        forms.alert("The selected wall is pinned. Please unpin it before running this tool.", exitscript=True)
        return
        
    inputs = get_user_inputs(wall)
    if not inputs:
        return
        
    door_width, door_height, door_count, maintain_center = inputs
    
    try:
        with Transaction(doc, 'Create Cooler Doors') as t:
            t.Start()
            
            # 1. Setup Type
            cooler_type = setup_cooler_type(wall, door_width, door_height)
            
            # 2. Check Mullions (Type -> Instance -> Alert)
            id_b1 = get_assigned_mullion_id(cooler_type, "AUTO_MULLION_BORDER1_VERT")
            id_b2 = get_assigned_mullion_id(cooler_type, "AUTO_MULLION_BORDER2_VERT")
            
            if id_b1 == ElementId.InvalidElementId or id_b2 == ElementId.InvalidElementId:
                # Try to get from instance
                inst_m_type = get_existing_mullion_type(doc, wall)
                if inst_m_type:
                    mid = inst_m_type.Id
                    if id_b1 == ElementId.InvalidElementId:
                        cooler_type.get_Parameter(BuiltInParameter.AUTO_MULLION_BORDER1_VERT).Set(mid)
                        id_b1 = mid
                    if id_b2 == ElementId.InvalidElementId:
                        cooler_type.get_Parameter(BuiltInParameter.AUTO_MULLION_BORDER2_VERT).Set(mid)
                        id_b2 = mid
                    
                    # Update Interior too
                    p_int = cooler_type.get_Parameter(BuiltInParameter.AUTO_MULLION_INTERIOR_VERT)
                    if p_int and p_int.AsElementId() == ElementId.InvalidElementId:
                        p_int.Set(mid)
            
            if id_b1 == ElementId.InvalidElementId or id_b2 == ElementId.InvalidElementId:
                p_name = cooler_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                type_name = p_name.AsString() if p_name else "Cooler Doors"
                
                forms.alert(
                    "Border Mullions are not defined for Wall Type '{}'.\n\nPlease Edit Type and assign Vertical Border Mullions, then run the tool again.".format(type_name),
                    title="Missing Mullions"
                )
                t.Commit()
                return

            # 3. Apply Type
            if wall.GetTypeId() != cooler_type.Id:
                wall.ChangeTypeId(cooler_type.Id)

            # 4. Calculate Length
            b1_width = get_mullion_type_width(doc.GetElement(id_b1))
            b2_width = get_mullion_type_width(doc.GetElement(id_b2))
            
            length_adjustment = (b1_width * 0.5) + (b2_width * 0.5)
            
            # 5. Set Height
            height_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
            if height_param and not height_param.IsReadOnly:
                height_param.Set(door_height)

            # 6. Resize Geometry
            wall_line = wall.Location.Curve
            start_point = wall_line.GetEndPoint(0)
            end_point = wall_line.GetEndPoint(1)
            
            new_length = (door_width * door_count) + length_adjustment
            current_length = wall_line.Length
            
            direction = (end_point - start_point).Normalize()
            
            if maintain_center and current_length > 0:
                mid_point = (start_point + end_point) * 0.5
                half_length = new_length * 0.5
                new_start = mid_point - direction * half_length
                new_end = mid_point + direction * half_length
            else:
                new_start = start_point
                new_end = start_point + direction * new_length
                
            new_line = Line.CreateBound(new_start, new_end)
            wall.Location.Curve = new_line
            
            # Set Grid Justification to Beginning and Offset to half mullion width
            # This ensures the first panel starts after the border mullion
            bip_just = getattr(BuiltInParameter, "CURTAIN_VERT_GRID_JUSTIFICATION", None)
            if bip_just:
                p_just = wall.get_Parameter(bip_just)
                if p_just: p_just.Set(0) # Beginning
            
            bip_off = getattr(BuiltInParameter, "CURTAIN_VERT_GRID_OFFSET", None)
            if bip_off:
                p_off = wall.get_Parameter(bip_off)
                if p_off: p_off.Set(b1_width * 0.5)
            
            t.Commit()
            
            TaskDialog.Show('Success', 'Created {} Cooler Doors ({:.2f}\' each).'.format(door_count, door_width))
            
    except Exception as e:
        forms.alert('Error: {}'.format(traceback.format_exc()))

if __name__ == '__main__':
    main()
