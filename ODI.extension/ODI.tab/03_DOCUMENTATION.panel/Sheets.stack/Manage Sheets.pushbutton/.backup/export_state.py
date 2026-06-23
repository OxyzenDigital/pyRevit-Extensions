# -*- coding=utf-8 -*-
"""Manage Sheets - Exporter.
Queries the active Revit document for sheets, views, sheet collections, and loaded
title blocks, and exports the state into a JS variable assignment in the
temp directory before launching the HTML simulator.
"""

import os
import sys
import csv
import json
import shutil
import tempfile
import webbrowser

# Revit API imports
try:
    import Autodesk
    from pyrevit import revit
    doc = revit.doc
    is_revit = True
except ImportError:
    doc = None
    is_revit = False

if is_revit:
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        ViewSheet,
        ViewSheetSet,
        View,
        ViewType,
        BuiltInCategory,
        BuiltInParameter,
        StorageType,
        Level
    )
def get_id_value(element_id):
    if hasattr(element_id, "Value"):
        return int(element_id.Value)
    return int(element_id.IntegerValue)

def get_element_name(elem):
    if not elem:
        return ""
    if hasattr(elem, "name"):
        return elem.name
    if hasattr(elem, "Name"):
        try:
            return elem.Name
        except:
            pass
    # Fallback using BuiltInParameter
    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString() or ""
    except:
        pass
    return ""

def get_title_blocks(doc):
    title_blocks = []
    if not doc:
        return title_blocks
        
    collector = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_TitleBlocks) \
        .WhereElementIsElementType()
        
    for symbol in collector:
        family_name = get_element_name(symbol.Family) if symbol.Family else ""
        type_name = get_element_name(symbol)
        full_name = "{} {}".format(family_name, type_name).lower()
        
        # 1. Try reading parameters first
        width_in_inches = 0.0
        w_param = symbol.get_Parameter(BuiltInParameter.SHEET_WIDTH)
        if w_param:
            width_in_inches = w_param.AsDouble() * 12.0
            
        height_in_inches = 0.0
        h_param = symbol.get_Parameter(BuiltInParameter.SHEET_HEIGHT)
        if h_param:
            height_in_inches = h_param.AsDouble() * 12.0
            
        # 2. If parameter reads 0, parse from symbol/family name or default to 24x36 (36" x 24")
        if width_in_inches < 0.1 or height_in_inches < 0.1:
            if "30x42" in full_name or "42x30" in full_name or "e1" in full_name:
                width_in_inches = 42.0
                height_in_inches = 30.0
            elif "24x36" in full_name or "36x24" in full_name or "d" in full_name:
                width_in_inches = 36.0
                height_in_inches = 24.0
            elif "22x34" in full_name or "34x22" in full_name:
                width_in_inches = 34.0
                height_in_inches = 22.0
            elif "11x17" in full_name or "17x11" in full_name:
                width_in_inches = 17.0
                height_in_inches = 11.0
            elif "a0" in full_name:
                width_in_inches = 46.8
                height_in_inches = 33.1
            elif "a1" in full_name:
                width_in_inches = 33.1
                height_in_inches = 23.4
            elif "a2" in full_name:
                width_in_inches = 23.4
                height_in_inches = 16.5
            elif "a3" in full_name:
                width_in_inches = 16.5
                height_in_inches = 11.7
            else:
                # Default fallback choice 24x36
                width_in_inches = 36.0
                height_in_inches = 24.0
                
        title_blocks.append({
            "FamilyName": family_name,
            "TypeName": type_name,
            "Id": get_id_value(symbol.Id),
            "Width": round(width_in_inches, 2),
            "Height": round(height_in_inches, 2)
        })
    return title_blocks

def get_sheets(doc):
    """Returns list of sheet dicts. Also reads SHEETS - Discipline and SHEETS - Use
    so the UI can pre-populate both browser-organisation parameters."""
    sheets_data = []
    if not doc:
        return sheets_data

    # Parameter name variants — try both common spellings
    DISC_PARAM_NAMES  = ["SHEETS - Discipline", "Sheet Discipline", "Discipline"]
    USE_PARAM_NAMES   = ["SHEETS - Use",        "Sheet Use",        "Use"]
    COLLECTION_NAMES  = ["Sheet Collection", "SheetCollection", "Collection"]

    def _read_param(element, name_candidates):
        for name in name_candidates:
            try:
                p = element.LookupParameter(name)
                if p and p.HasValue:
                    val = p.AsString()
                    if not val:
                        val = p.AsValueString()
                    if val:
                        return val
            except:
                pass
        return ""

    collector = FilteredElementCollector(doc).OfClass(ViewSheet)
    for sheet in collector:
        # Query custom Shared Parameter ODI_Schema_Link
        schema_link = ""
        param = sheet.LookupParameter("ODI_Schema_Link")
        if param:
            schema_link = param.AsString() or ""

        # Read Project Browser organisation parameters
        sheet_discipline = _read_param(sheet, DISC_PARAM_NAMES)
        sheet_use        = _read_param(sheet, USE_PARAM_NAMES)
        
        # Read Sheet Collection exactly by parameter (Revit 2025+ or Custom)
        sheet_collection = ""
        
        try:
            # 1. Native Revit 2025+ via ParameterTypeId
            from Autodesk.Revit.DB import ParameterTypeId
            c_param = sheet.get_Parameter(ParameterTypeId.SheetCollection)
            if c_param and c_param.HasValue:
                c_id = c_param.AsElementId()
                if c_id and c_id.IntegerValue > -1:
                    c_elem = doc.GetElement(c_id)
                    if c_elem:
                        sheet_collection = get_element_name(c_elem)
        except:
            pass
            
        # 2. Strict Project Parameter Lookup
        if not sheet_collection:
            sheet_collection = _read_param(sheet, ["Sheet Collection"])

        sheet_id = get_id_value(sheet.Id)
        
        # Get views placed on this sheet
        placed_views = []
        if hasattr(sheet, "GetAllPlacedViews"):
            try:
                for v_id in sheet.GetAllPlacedViews():
                    view_elem = doc.GetElement(v_id)
                    if view_elem:
                        detail_param = view_elem.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                        detail_num = detail_param.AsString() if detail_param else ""
                        placed_views.append({
                            "Id": get_id_value(v_id),
                            "Name": get_element_name(view_elem),
                            "Type": str(view_elem.ViewType) if hasattr(view_elem, "ViewType") else "",
                            "DetailNumber": detail_num or ""
                        })
            except:
                pass

        sheets_data.append({
            "Number": sheet.SheetNumber,
            "Name": get_element_name(sheet),
            "Id": sheet_id,
            "SchemaLink": schema_link,
            "SheetCollection": sheet_collection,
            "PlacedViews": placed_views,
            "SheetDiscipline": sheet_discipline,
            "SheetUse": sheet_use
        })
    return sheets_data

def get_project_info(doc):
    info_data = {}
    if not doc:
        return info_data
    try:
        project_info = doc.ProjectInformation
        if project_info:
            for p in project_info.Parameters:
                try:
                    p_name = p.Definition.Name
                    val = ""
                    if hasattr(p, "StorageType"):
                        st = p.StorageType
                        if st == StorageType.String:
                            val = p.AsString() or ""
                        elif st == StorageType.Integer:
                            val = p.AsInteger()
                        elif st == StorageType.Double:
                            val = p.AsValueString() or p.AsDouble()
                        elif st == StorageType.ElementId:
                            val = get_id_value(p.AsElementId())
                    else:
                        val = p.AsString() or ""
                    
                    if val is not None:
                        info_data[p_name] = val
                except:
                    pass
    except:
        pass
    return info_data

def get_views(doc):
    views_data = []
    if not doc:
        return views_data
        
    collector = FilteredElementCollector(doc).OfClass(View)
    for view in collector:
        if view.IsTemplate:
            continue
            
        # Select views that can be placed on a sheet
        valid_types = [
            ViewType.FloorPlan,
            ViewType.CeilingPlan,
            ViewType.Elevation,
            ViewType.Section,
            ViewType.ThreeD
        ]
        
        if view.ViewType in valid_types:
            views_data.append({
                "Id": get_id_value(view.Id),
                "Name": get_element_name(view),
                "Type": str(view.ViewType)
            })
    return views_data



def parse_ncs_csv(filename):
    current_dir = os.path.dirname(__file__)
    csv_path = os.path.join(current_dir, "_Resources", filename)
    data = []
    if not os.path.exists(csv_path):
        return data
        
    open_mode = "rb" if sys.version_info[0] < 3 else "r"
    open_kwargs = {} if sys.version_info[0] < 3 else {"encoding": "utf-8", "newline": ""}
    
    with open(csv_path, open_mode, **open_kwargs) as f:
        reader = csv.reader(f)
        header = next(reader, None) # Skip header
        for row in reader:
            if row:
                # Remove empty/whitespace rows
                if any(cell.strip() for cell in row):
                    data.append([cell.decode('utf-8') if sys.version_info[0] < 3 else cell for cell in row])
    return data

def get_levels(doc):
    levels_data = []
    if not doc:
        return levels_data
    try:
        collector = FilteredElementCollector(doc).OfClass(Level)
        for lvl in collector:
            levels_data.append({
                "Id": get_id_value(lvl.Id),
                "Name": get_element_name(lvl),
                "Elevation": lvl.Elevation
            })
        # Sort levels by elevation ascending
        levels_data.sort(key=lambda x: x["Elevation"])
    except Exception as e:
        pass
    return levels_data

def run():
    if not doc:
        print("Error: Revit Document context not found.")
        return
        
    # 2. Extract model elements
    sheets = get_sheets(doc)
    views = get_views(doc)
    levels = get_levels(doc)
    title_blocks = get_title_blocks(doc)
    
    # 3. Read NCS reference data from resources
    disciplines = parse_ncs_csv("Discipline_Designators.csv")
    sheet_types = parse_ncs_csv("Sheet_Types.csv")
    sequence_numbers = parse_ncs_csv("Sequence_Numbers.csv")
    templates = parse_ncs_csv("Sheet_Names_and_Numbers.csv")
    
    # 4. Collect unique SHEETS - Use values from live model
    all_sheet_use_values = sorted(set(
        s["SheetUse"] for s in sheets if s.get("SheetUse")
    ))

    # Collect unique Sheet Collections from live model
    all_sheet_collections = sorted(set(
        s["SheetCollection"] for s in sheets if s.get("SheetCollection")
    ))

    # 5. Consolidate payload
    payload = {
        "projectInfo": get_project_info(doc),
        "levels": levels,
        "sheets": sheets,
        "views": views,
        "sheetCollections": all_sheet_collections,
        "titleBlocks": title_blocks,
        "sheetUseValues": all_sheet_use_values,
        "ncs": {
            "disciplines": disciplines,
            "sheetTypes": sheet_types,
            "sequenceNumbers": sequence_numbers,
            "templates": templates
        }
    }
    
    # 5. Write to revit_state.js in TEMP folder
    temp_dir = tempfile.gettempdir()
    js_path = os.path.join(temp_dir, "revit_state.js")
    
    with open(js_path, "w") as f:
        # Wrap in global variable assignment for CORS bypass
        f.write("window.RevitState = {};".format(json.dumps(payload)))
        
    # 6. Copy index.html to TEMP folder and open in default browser
    current_dir = os.path.dirname(__file__)
    index_src = os.path.join(current_dir, "index.html")
    index_dst = os.path.join(temp_dir, "index.html")
    
    if os.path.exists(index_src):
        shutil.copy(index_src, index_dst)
        url = "file:///" + index_dst.replace("\\", "/")
        webbrowser.open(url)
    else:
        from pyrevit import forms
        forms.alert("Could not locate index.html in the pushbutton directory.", title="Missing HTML UI")

if __name__ == "__main__":
    # If run standalone inside Revit
    run()
