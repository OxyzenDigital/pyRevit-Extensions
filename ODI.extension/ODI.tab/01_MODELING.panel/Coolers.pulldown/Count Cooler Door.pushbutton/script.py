# -*- coding: utf-8 -*-
"""Count Cooler Doors.
Counts panels in selected curtain walls or scans model for 'Cooler' walls."""

__title__ = 'Count\nCooler Doors'
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
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_target_walls():
    """
    Determines target walls based on selection.
    - If Selection exists: Returns selected Curtain Walls.
    - If No Selection: Scans model for Curtain Walls with 'Cooler' in Type Name.
    """
    selection = uidoc.Selection.GetElementIds()
    walls = []
    source_type = ""

    if selection:
        source_type = "Current Selection"
        for eid in selection:
            el = doc.GetElement(eid)
            if isinstance(el, Wall) and el.CurtainGrid:
                walls.append(el)
        
        if not walls:
            forms.alert("Selection contains no Curtain Walls.", exitscript=True)
            return None, None
    else:
        source_type = "Active View (Cooler Walls)"
        # Scan for walls with "Cooler" in the type name, filtered by Active View (Phase/Design Option)
        all_walls = FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(Wall).ToElements()
        for w in all_walls:
            if w.CurtainGrid:
                w_type = doc.GetElement(w.GetTypeId())
                if w_type:
                    p_name = w_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p_name and "cooler" in p_name.AsString().lower():
                        walls.append(w)
        
        if not walls:
            forms.alert("No 'Cooler' curtain walls found in the model.", exitscript=True)
            return None, None

    return walls, source_type

def to_feet_inch(val):
    feet = int(val)
    inches = round((val - feet) * 12)
    if inches == 12:
        feet += 1
        inches = 0
    return "{}'-{}\"".format(feet, inches)

def main():
    walls, source = get_target_walls()
    if not walls: return

    output = script.get_output()
    data = []
    grand_total = 0

    for w in walls:
        grid = w.CurtainGrid
        if not grid:
            continue

        # Get Wall Type Name
        w_type = doc.GetElement(w.GetTypeId())
        w_type_name = "Unknown"
        if w_type:
            p_name = w_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p_name: w_type_name = p_name.AsString()

        # Analyze Panels
        panel_ids = grid.GetPanelIds()
        panel_types = set()
        panel_dims = set()
        count = 0

        for pid in panel_ids:
            panel = doc.GetElement(pid)
            if not isinstance(panel, FamilyInstance): continue

            # Panel Type
            p_type = doc.GetElement(panel.GetTypeId())
            if p_type:
                p_name = p_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if p_name:
                    panel_types.add(p_name.AsString())

            # Panel Dimensions (Width x Height)
            # Try Instance Parameters first (Standard Curtain Panels)
            w_param = panel.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH)
            h_param = panel.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT)
            
            width = w_param.AsDouble() if w_param else 0.0
            height = h_param.AsDouble() if h_param else 0.0
            
            # Fallback to Type Parameters if Instance is missing/zero (e.g. Fixed Door Families)
            if width < 0.1 and p_type:
                w_type_p = p_type.get_Parameter(BuiltInParameter.DOOR_WIDTH)
                if w_type_p: width = w_type_p.AsDouble()
            if height < 0.1 and p_type:
                h_type_p = p_type.get_Parameter(BuiltInParameter.DOOR_HEIGHT)
                if h_type_p: height = h_type_p.AsDouble()

            if width > 0.5: # Filter out slivers/mullion joins (approx 6 inches)
                panel_dims.add((round(width, 2), round(height, 2)))
                count += 1
        
        grand_total += count
        
        data.append([
            output.linkify(w.Id),
            w_type_name,
            ", ".join(sorted(panel_types)),
            ", ".join(["{} x {}".format(to_feet_inch(w), to_feet_inch(h)) for w, h in sorted(panel_dims)]),
            count
        ])

    data.append(["", "**GRAND TOTAL**", "", "", "**{}**".format(grand_total)])

    if data:
        output.print_table(
            table_data=data,
            title="Cooler Door Report ({})".format(source),
            columns=["Wall ID", "Wall Type", "Panel Types", "Panel Dims (WxH)", "Door Count"]
        )
    else:
        forms.alert("No panels found in selected walls.")

if __name__ == '__main__':
    main()