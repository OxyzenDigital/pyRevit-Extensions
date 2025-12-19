# -*- coding: utf-8 -*-
"""
Energy Analysis Master Exporter (v43.0)
- FIX: Solved 'Unknown' Construction Group by using Element Category as a fallback.
- LOGIC: Enum Check -> Category Name Check -> Parameter Check.
- STRUCTURE: 3 Main Headers (Settings, Spaces, Types).
"""
import os
import io
import datetime
from pyrevit import revit, script, forms
import Autodesk.Revit.DB as DB

# --- Configuration ---
doc = revit.doc
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
user_home = os.path.expanduser("~")
export_folder = os.path.join(user_home, "Downloads", "Exported Energy Settings")

if not os.path.exists(export_folder):
    os.makedirs(export_folder)

# ==============================================================================
# PHASE 1: HELPERS & MAPPINGS
# ==============================================================================

# Expanded Map for EnergyAnalysisConstruction.ConstructionType
ENUM_MAP = {
    0: "Exterior Wall",
    1: "Interior Wall",
    2: "Roof",
    3: "Floor",
    4: "Door",
    5: "Window",
    6: "Skylight",
    7: "Underground Wall",
    8: "Underground Slab",
    9: "Ceiling",
    10: "Slab On Grade",
    11: "Shade",
    12: "Air"
}

def get_bip(name):
    if hasattr(DB.BuiltInParameter, name):
        return getattr(DB.BuiltInParameter, name)
    return None

def get_safe_string(p):
    if not p or not p.HasValue: return ""
    return p.AsString() or p.AsValueString() or ""

def get_element_name(element):
    if not element: return "Unknown"
    r_num = get_safe_string(element.get_Parameter(get_bip("ROOM_NUMBER")))
    r_name = get_safe_string(element.get_Parameter(get_bip("ROOM_NAME")))
    if r_num and r_name: return "{} - {}".format(r_num, r_name)
    t_name = get_safe_string(element.get_Parameter(get_bip("ALL_MODEL_TYPE_NAME")))
    if t_name: return t_name
    if hasattr(element, "Name") and element.Name: return element.Name
    return "ID: {}".format(element.Id)

def get_grouped_parameters(element):
    if not element: return {}
    groups = {}
    for p in element.Parameters:
        if not p.Definition or not p.HasValue: continue
        g_name = "Other"
        try:
            if hasattr(p.Definition, "GetGroupTypeId"): 
                g_name = DB.LabelUtils.GetLabelFor(p.Definition.GetGroupTypeId())
            else:
                g_name = str(p.Definition.ParameterGroup).replace("PG_", "").replace("_", " ").title()
        except: pass
        val = get_safe_string(p)
        if val:
            if g_name not in groups: groups[g_name] = []
            groups[g_name].append((p.Definition.Name, val))
    for g in groups: groups[g].sort(key=lambda x: x[0])
    return groups

def resolve_construction_role(element):
    """
    Robustly identifies if an element is a Wall, Roof, Window, etc.
    Tries API Enum first, then falls back to Category Name.
    """
    # 1. Try API Enum
    try:
        if hasattr(element, "ConstructionType"):
            enum_int = int(element.ConstructionType)
            if enum_int in ENUM_MAP:
                return ENUM_MAP[enum_int]
    except: pass

    # 2. Try Category Name (The Fallback)
    if element.Category:
        cat_name = element.Category.Name
        # Map common Revit Category names to our clean roles
        if "Roof" in cat_name: return "Roof"
        if "Exterior Wall" in cat_name: return "Exterior Wall"
        if "Interior Wall" in cat_name: return "Interior Wall"
        if "Floor" in cat_name: return "Floor"
        if "Window" in cat_name: return "Window"
        if "Door" in cat_name: return "Door"
        if "Skylight" in cat_name: return "Skylight"
        if "Slab" in cat_name: return "Slab On Grade"
        if "Underground" in cat_name: return "Underground Wall"
        
    # 3. Try Parameter "Category"
    p = element.LookupParameter("Category")
    if p: return get_safe_string(p)

    return "Unknown"

# ==============================================================================
# PHASE 2: COLLECTORS
# ==============================================================================

def collect_analytical_settings():
    data = {"GlobalSettings": {}, "Constructions": {}, "BuildingTypes": []}
    
    # 1. GLOBAL ENERGY SETTINGS
    es = DB.Analysis.EnergyDataSettings.GetFromDocument(doc)
    if es:
        data["GlobalSettings"] = get_grouped_parameters(es)
        
        # Inject Enums
        if "Energy Analysis" not in data["GlobalSettings"]:
             data["GlobalSettings"]["Energy Analysis"] = []
        try: 
            val = str(es.BuildingType)
            data["GlobalSettings"]["Energy Analysis"].insert(0, ("Building Type", val))
        except: pass
        try:
            val = str(es.ServiceType)
            data["GlobalSettings"]["Energy Analysis"].insert(0, ("Building Service", val))
        except: pass

    # 2. ANALYTIC CONSTRUCTIONS
    if hasattr(DB.Analysis, "EnergyAnalysisConstruction"):
        elems = DB.FilteredElementCollector(doc).OfClass(DB.Analysis.EnergyAnalysisConstruction).ToElements()
        for e in elems:
            # ROBUST RESOLVER
            enum_name = resolve_construction_role(e)
            
            # GET DATA STRING
            val = ""
            for p_name in ["Analytic Construction", "Analytic Construction Type", "Schematic Type", "Construction Name", "Name"]:
                p = e.LookupParameter(p_name)
                if p and p.HasValue:
                    val = get_safe_string(p)
                    break
            
            if enum_name not in data["Constructions"]: data["Constructions"][enum_name] = []
            data["Constructions"][enum_name].append({
                "enum": enum_name, 
                "val": val, 
                "id": e.Id
            })

    # 3. BUILDING TYPES LIBRARY
    bt_cat = getattr(DB.BuiltInCategory, "OST_HVAC_Load_Building_Types", None)
    if bt_cat:
        data["BuildingTypes"] = list(DB.FilteredElementCollector(doc).OfCategory(bt_cat).ToElements())

    return data

def collect_spaces_and_types():
    spaces = list(DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).ToElements())
    used_type_ids = set()
    for s in spaces:
        p = s.get_Parameter(get_bip("ROOM_SPACE_TYPE_PARAM"))
        if p and p.HasValue: used_type_ids.add(p.AsElementId())
    all_types = []
    for cat in ["OST_HVAC_Load_Space_Types", "OST_MEPSpaceType"]:
        if hasattr(DB.BuiltInCategory, cat):
            bic = getattr(DB.BuiltInCategory, cat)
            all_types.extend(list(DB.FilteredElementCollector(doc).OfCategory(bic).ToElements()))
    used_types = [t for t in all_types if t.Id in used_type_ids]
    return spaces, used_types

# ==============================================================================
# PHASE 3: HTML RENDERERS
# ==============================================================================

def render_params_table(html, params):
    html.append(u"<table>")
    for p_name, p_val in params:
        html.append(u"<tr><th>{}</th><td>{}</td></tr>".format(p_name, p_val))
    html.append(u"</table>")

def render_deep_item(html, element):
    name = get_element_name(element)
    html.append(u"<button class='item-accordion'>{} <span style='float:right; font-weight:normal; font-size:0.8em; color:#bbb'>ID: {}</span></button>".format(name, element.Id))
    html.append(u"<div class='item-panel'>")
    groups = get_grouped_parameters(element)
    if not groups:
        html.append(u"<div style='padding:10px; color:#aaa'>No parameters found.</div>")
    else:
        for g_name, params in sorted(groups.items()):
            html.append(u"<button class='sub-accordion'>{}</button>".format(g_name))
            html.append(u"<div class='sub-panel'>")
            render_params_table(html, params)
            html.append(u"</div>")
    html.append(u"</div>")

def render_global_settings(html, grouped_params):
    if not grouped_params:
        html.append(u"<p>No settings found.</p>")
        return
    for g_name, params in sorted(grouped_params.items()):
        html.append(u"<button class='sub-accordion'>{}</button>".format(g_name))
        html.append(u"<div class='sub-panel'>")
        render_params_table(html, params)
        html.append(u"</div>")

def render_constructions(html, data):
    if not data:
        html.append(u"<p style='color:#ccc'>No constructions found.</p>")
        return
    
    html.append(u"<div style='margin-bottom:5px; font-weight:bold; font-size:0.85em; color:#7f8c8d; display:flex; justify-content:space-between; padding:0 5px;'>")
    html.append(u"<span style='width:30%'>ENUMERATION</span>")
    html.append(u"<span style='width:65%'>ANALYTIC DATA STRING</span>")
    html.append(u"</div>")
    
    # Strict Logical Order
    order = ["Exterior Wall", "Interior Wall", "Roof", "Floor", "Door", "Window", "Skylight", "Slab On Grade", "Underground Wall"]
    sorted_keys = sorted(data.keys(), key=lambda k: order.index(k) if k in order else 99)

    for role in sorted_keys:
        html.append(u"<div class='cat-header'>{} Group</div>".format(role))
        html.append(u"<table>")
        for item in data[role]:
            html.append(u"<tr><td style='width:30%; font-weight:bold; color:#2980b9'>{}</td><td style='color:#d35400;'>{}</td></tr>".format(item['enum'], item['val']))
        html.append(u"</table>")

def render_list_deep(html, element_list):
    if not element_list:
        html.append(u"<p style='color:#ccc; padding:10px'>No items found.</p>")
        return
    sorted_list = sorted(element_list, key=lambda x: get_element_name(x))
    for el in sorted_list:
        render_deep_item(html, el)

def generate_html(settings_data, spaces, used_types):
    css = """
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #f4f4f9; padding: 30px; color: #333; }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        
        .accordion {
            background-color: #fff; color: #2c3e50; cursor: pointer; padding: 18px; width: 100%;
            border: 1px solid #ddd; border-left: 6px solid #3498db; text-align: left;
            font-size: 16px; font-weight: bold; margin-top: 15px; display: flex; justify-content: space-between;
        }
        .active, .accordion:hover { background-color: #f0f7ff; }
        .panel { padding: 0 15px; background-color: white; max-height: 0; overflow: hidden; transition: max-height 0.2s ease-out; border: 1px solid #ddd; border-top:none; }
        
        .item-accordion {
            background-color: #fdfdfd; color: #444; cursor: pointer; padding: 10px 15px; width: 100%;
            border: 1px solid #eee; border-left: 4px solid #95a5a6; text-align: left;
            font-size: 14px; font-weight: 600; margin-top: 5px;
        }
        .item-accordion:hover { background-color: #f2f2f2; }
        .item-panel { padding: 0 10px; background-color: #fff; max-height: 0; overflow: hidden; transition: max-height 0.2s ease-out; border-left: 1px solid #eee; border-right: 1px solid #eee; }

        .sub-accordion {
            background: #fafafa; color: #666; cursor: pointer; padding: 6px 12px; width: 100%;
            border: 1px solid #eee; text-align: left; font-size: 13px; font-weight: bold; margin-top: 2px;
        }
        .sub-panel { padding: 10px; background: #fff; display: none; border: 1px solid #eee; border-top:none; }
        
        table { width: 100%; border-collapse: collapse; font-size: 0.9em; margin: 5px 0; }
        th, td { text-align: left; padding: 5px 10px; border-bottom: 1px solid #f0f0f0; }
        th { width: 40%; color: #999; font-weight: normal; }
        
        .cat-header { background: #eaf2f8; padding: 5px 10px; font-weight: bold; color: #2980b9; margin-top: 10px; font-size: 0.9em; }
        .section-label { background: #34495e; color: white; padding: 8px 15px; font-weight: bold; margin-top: 20px; border-radius: 4px; }
    </style>
    """
    js = """
    <script>
        document.addEventListener("DOMContentLoaded", function() {
            var acc = document.querySelectorAll(".accordion, .item-accordion");
            for (var i = 0; i < acc.length; i++) {
                acc[i].addEventListener("click", function() {
                    this.classList.toggle("active");
                    var panel = this.nextElementSibling;
                    if (panel.style.maxHeight) { 
                        panel.style.maxHeight = null; 
                    } else { 
                        panel.style.maxHeight = panel.scrollHeight + "px"; 
                        var parent = this.closest('.panel, .item-panel');
                        if (parent) { parent.style.maxHeight = parent.scrollHeight + panel.scrollHeight + "px"; }
                    } 
                });
            }
            var sub = document.getElementsByClassName("sub-accordion");
            for (var i = 0; i < sub.length; i++) {
                sub[i].addEventListener("click", function() {
                    var p = this.nextElementSibling;
                    if (p.style.display === "block") { p.style.display = "none"; } else { p.style.display = "block"; }
                    var parent1 = this.closest('.item-panel');
                    if(parent1) parent1.style.maxHeight = parent1.scrollHeight + "px";
                    var parent2 = this.closest('.panel');
                    if(parent2) parent2.style.maxHeight = parent2.scrollHeight + "px";
                });
            }
        });
    </script>
    """
    
    html = [u"<html><head><meta charset='utf-8'>{}{}</head><body>".format(css, js)]
    html.append(u"<h1>Energy Analysis Master Report</h1>")
    html.append(u"<p style='color:#888'>Generated: {}</p>".format(timestamp))
    
    html.append(u"<button class='accordion'>1. Analytical Energy Settings</button>")
    html.append(u"<div class='panel'><div style='padding:15px'>")
    
    html.append(u"<div class='section-label'>A. Global Energy Settings</div>")
    render_global_settings(html, settings_data["GlobalSettings"])
    
    html.append(u"<div class='section-label'>B. Analytic Construction Map</div>")
    render_constructions(html, settings_data["Constructions"])
    
    html.append(u"<div class='section-label'>C. Building Types Library (Full Detail)</div>")
    render_list_deep(html, settings_data["BuildingTypes"])
    
    html.append(u"</div></div>")
    
    html.append(u"<button class='accordion'>2. Project Spaces ({})</button>".format(len(spaces)))
    html.append(u"<div class='panel'><div style='padding:15px'>")
    render_list_deep(html, spaces)
    html.append(u"</div></div>")
    
    html.append(u"<button class='accordion'>3. Used Space Types ({})</button>".format(len(used_types)))
    html.append(u"<div class='panel'><div style='padding:15px'>")
    render_list_deep(html, used_types)
    html.append(u"</div></div>")
    
    html.append(u"</body></html>")
    return u"".join(html)

# ==============================================================================
# MAIN
# ==============================================================================

def main():
    try:
        settings_data = collect_analytical_settings()
        spaces, used_types = collect_spaces_and_types()
        filename = "EnergyAnalysis_Final_{}.html".format(timestamp)
        save_path = os.path.join(export_folder, filename)
        html = generate_html(settings_data, spaces, used_types)
        with io.open(save_path, "w", encoding="utf-8") as f: f.write(html)
        forms.alert("Report Generated!\nFile: {}".format(filename), warn_icon=False)
        os.startfile(export_folder)
    except Exception as e:
        forms.alert("Error: {}".format(e))

if __name__ == '__main__':
    main()