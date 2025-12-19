# -*- coding: utf-8 -*-
"""
Energy Analysis Master Exporter (v48.0)
- BUG FIX: Solved "Last Row Cutoff" via Recursive Parent Resizing in JS.
- FEATURE: Smart Filter for Active Building Type.
- OUTPUT: Dual Export (HTML + Flat TXT).
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
# PHASE 1: HELPERS
# ==============================================================================

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

def normalize_name(name):
    if not name: return ""
    return name.replace(" ", "").lower().strip()

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

def get_flat_parameters(element):
    if not element: return []
    params = []
    for p in element.Parameters:
        if not p.Definition or not p.HasValue: continue
        val = get_safe_string(p)
        if val:
            params.append((p.Definition.Name, val))
    params.sort(key=lambda x: x[0])
    return params

# ==============================================================================
# PHASE 2: COLLECTORS
# ==============================================================================

def collect_analytical_settings():
    data = {"GlobalSettings": {}, "Constructions": {}, "BuildingTypes": []}
    
    # 1. GLOBAL ENERGY SETTINGS
    es = DB.Analysis.EnergyDataSettings.GetFromDocument(doc)
    active_b_type_str = "" 
    
    if es:
        data["GlobalSettings"] = get_grouped_parameters(es)
        if "Energy Analysis" not in data["GlobalSettings"]:
             data["GlobalSettings"]["Energy Analysis"] = []
        try: 
            active_b_type_str = str(es.BuildingType) 
            data["GlobalSettings"]["Energy Analysis"].insert(0, ("Building Type", active_b_type_str))
        except: pass
        try:
            val = str(es.ServiceType)
            data["GlobalSettings"]["Energy Analysis"].insert(0, ("Building Service", val))
        except: pass

    # 2. SCHEMATIC CONSTRUCTIONS
    if hasattr(DB.Analysis, "EnergyAnalysisConstruction"):
        elems = DB.FilteredElementCollector(doc).OfClass(DB.Analysis.EnergyAnalysisConstruction).ToElements()
        for e in elems:
            enum_name = "Unknown"
            try: enum_name = str(e.ConstructionType)
            except: pass
            
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
                "element": e
            })

    # 3. BUILDING TYPES LIBRARY (FILTERED)
    bt_cat = getattr(DB.BuiltInCategory, "OST_HVAC_Load_Building_Types", None)
    if bt_cat:
        all_types = list(DB.FilteredElementCollector(doc).OfCategory(bt_cat).ToElements())
        target_name = normalize_name(active_b_type_str)
        filtered_types = []
        for b in all_types:
            b_name = get_element_name(b)
            if normalize_name(b_name) == target_name:
                filtered_types.append(b)
        data["BuildingTypes"] = filtered_types

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
# PHASE 3: HTML RENDERER
# ==============================================================================

def render_params_table(html, params):
    html.append(u"<table>")
    for p_name, p_val in params:
        html.append(u"<tr><th>{}</th><td>{}</td></tr>".format(p_name, p_val))
    html.append(u"</table>")

def render_deep_item(html, element, title_override=None):
    name = title_override if title_override else get_element_name(element)
    html.append(u"<button class='item-accordion'>{} <span style='float:right; font-weight:normal; font-size:0.8em; color:#95a5a6'>ID: {}</span></button>".format(name, element.Id))
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
    html.append(u"<div style='margin-bottom:8px; font-weight:bold; font-size:0.85em; color:#7f8c8d; display:flex; justify-content:space-between; padding:0 10px;'>")
    html.append(u"<span style='width:30%'>ENUMERATION TYPE</span>")
    html.append(u"<span style='width:65%'>SCHEMATIC DEFINITION</span>")
    html.append(u"</div>")
    sorted_keys = sorted(data.keys())
    for role in sorted_keys:
        items = data[role]
        html.append(u"<div class='cat-header'>{} <span style='font-weight:normal; font-size:0.85em; opacity:0.8'>({} Items)</span></div>".format(role, len(items)))
        html.append(u"<div style='padding-bottom:15px'>")
        for item in items:
            title = item['val'] if item['val'] else "Un-named Construction"
            render_deep_item(html, item['element'], title_override=title)
        html.append(u"</div>")

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
        body { font-family: 'Segoe UI', Roboto, Helvetica, sans-serif; background: #ecf0f1; padding: 40px; color: #2c3e50; }
        h1 { color: #2c3e50; border-bottom: 4px solid #3498db; padding-bottom: 15px; margin-bottom: 5px; letter-spacing: -0.5px; }
        
        .accordion {
            background-color: #fff; color: #2c3e50; cursor: pointer; padding: 20px; width: 100%;
            border: 1px solid #bdc3c7; border-left: 8px solid #2980b9; text-align: left;
            font-size: 18px; font-weight: 700; margin-top: 20px; display: flex; justify-content: space-between;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05); transition: all 0.2s ease;
        }
        .accordion:hover { background-color: #f7f9fa; border-left-color: #3498db; transform: translateX(2px); }
        .panel { padding: 0 20px; background-color: white; max-height: 0; overflow: hidden; transition: max-height 0.3s ease-out; border: 1px solid #bdc3c7; border-top:none; }
        
        .item-accordion {
            background-color: #f8f9fa; color: #34495e; cursor: pointer; padding: 12px 15px; width: 100%;
            border: 1px solid #e0e0e0; border-left: 5px solid #bdc3c7; text-align: left;
            font-size: 14px; font-weight: 600; margin-top: 8px; transition: background 0.15s;
        }
        .item-accordion:hover { background-color: #ecf0f1; border-left-color: #95a5a6; }
        .item-panel { padding: 0 15px; background-color: #fff; max-height: 0; overflow: hidden; transition: max-height 0.2s ease-out; border-left: 1px solid #eee; border-right: 1px solid #eee; }

        .sub-accordion {
            background: #ffffff; color: #2980b9; cursor: pointer; padding: 8px 10px; width: 100%;
            border: none; border-bottom: 1px solid #eee; text-align: left; font-size: 13px; font-weight: 700; margin-top: 0;
            text-transform: uppercase; letter-spacing: 0.5px;
        }
        .sub-accordion:hover { background-color: #fbfbfb; color: #3498db; }
        .sub-panel { padding: 10px 15px 15px 15px; background: #fff; display: none; border-bottom: 1px solid #eee; }
        
        table { width: 100%; border-collapse: collapse; font-size: 0.9em; margin: 5px 0; }
        th, td { text-align: left; padding: 6px 10px; border-bottom: 1px solid #f1f1f1; }
        th { width: 35%; color: #7f8c8d; font-weight: 500; } td { color: #2c3e50; } tr:hover { background-color: #fcfcfc; }
        .cat-header { background: #e8f6f3; padding: 10px 15px; font-weight: 700; color: #16a085; margin-top: 20px; font-size: 1.05em; border-left: 4px solid #1abc9c; border-radius: 0 4px 4px 0; }
        .section-label { background: #16a085; color: white; padding: 10px 20px; font-weight: 700; font-size: 1.1em; margin-top: 30px; margin-bottom: 10px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    </style>
    """
    
    # RECURSIVE PARENT RESIZE LOGIC
    js = """<script>
    document.addEventListener("DOMContentLoaded", function() {
        
        function resizeParentPanels(element) {
            var parent = element.parentElement;
            while (parent) {
                if (parent.classList.contains('panel') || parent.classList.contains('item-panel')) {
                    // Force the parent to fit its new scrollHeight
                    parent.style.maxHeight = parent.scrollHeight + "px";
                }
                parent = parent.parentElement;
            }
        }

        var acc = document.querySelectorAll(".accordion, .item-accordion, .sub-accordion");
        for (var i = 0; i < acc.length; i++) {
            acc[i].addEventListener("click", function() {
                // Toggle self
                this.classList.toggle("active");
                var panel = this.nextElementSibling;
                
                if (panel.style.maxHeight && panel.style.maxHeight !== "0px") {
                    // COLLAPSE
                    panel.style.maxHeight = null;
                } else {
                    // EXPAND
                    // 1. Show this panel
                    panel.style.display = "block"; // Ensure display is block for sub-panels
                    panel.style.maxHeight = panel.scrollHeight + "px";
                    
                    // 2. BUBBLE UP: Tell all parents to grow
                    resizeParentPanels(this);
                }
            });
        }
    });</script>"""
    
    html = [u"<html><head><meta charset='utf-8'>{}{}</head><body>".format(css, js)]
    html.append(u"<h1>Energy Analysis Master Report</h1>")
    html.append(u"<p style='color:#7f8c8d; font-style:italic'>Generated: {}</p>".format(timestamp))
    html.append(u"<button class='accordion'>1. Analytical Energy Settings</button><div class='panel'><div style='padding:20px'>")
    html.append(u"<div class='section-label'>A. Global Energy Settings</div>")
    render_global_settings(html, settings_data["GlobalSettings"])
    html.append(u"<div class='section-label'>B. Schematic Construction Types</div>")
    render_constructions(html, settings_data["Constructions"])
    html.append(u"<div class='section-label'>C. Active Building Type (Full Detail)</div>")
    render_list_deep(html, settings_data["BuildingTypes"])
    html.append(u"</div></div>")
    html.append(u"<button class='accordion'>2. Project Spaces ({})</button><div class='panel'><div style='padding:20px'>".format(len(spaces)))
    render_list_deep(html, spaces)
    html.append(u"</div></div>")
    html.append(u"<button class='accordion'>3. Used Space Types ({})</button><div class='panel'><div style='padding:20px'>".format(len(used_types)))
    render_list_deep(html, used_types)
    html.append(u"</div></div></body></html>")
    return u"".join(html)

# ==============================================================================
# PHASE 4: TEXT GENERATOR
# ==============================================================================

def generate_txt(settings_data, spaces, used_types):
    lines = []
    lines.append("ENERGY ANALYSIS MASTER REPORT (AI OPTIMIZED)")
    lines.append("Timestamp: {}".format(timestamp))
    lines.append("="*50)
    
    # 1. SETTINGS
    lines.append("\n1. ANALYTICAL ENERGY SETTINGS")
    
    lines.append("\n  A. GLOBAL ENERGY SETTINGS")
    flat_globals = []
    for g_list in settings_data["GlobalSettings"].values():
        flat_globals.extend(g_list)
    flat_globals.sort(key=lambda x: x[0])
    for p, v in flat_globals:
        lines.append("    {}: {}".format(p, v))
            
    lines.append("\n  B. SCHEMATIC CONSTRUCTION TYPES")
    for role in sorted(settings_data["Constructions"].keys()):
        lines.append("    GROUP: {}".format(role))
        for item in settings_data["Constructions"][role]:
            val = item['val'] if item['val'] else "Un-named"
            lines.append("      ELEMENT: {} (ID: {})".format(val, item['element'].Id))
            flat_params = get_flat_parameters(item['element'])
            for p, v in flat_params:
                lines.append("        {}: {}".format(p, v))
            
    lines.append("\n  C. ACTIVE BUILDING TYPE")
    if not settings_data["BuildingTypes"]:
        lines.append("    (No matching Building Type element found in library for active enum)")
    for b in sorted(settings_data["BuildingTypes"], key=lambda x: get_element_name(x)):
        lines.append("    ELEMENT: {} (ID: {})".format(get_element_name(b), b.Id))
        flat_params = get_flat_parameters(b)
        for p, v in flat_params:
            lines.append("      {}: {}".format(p, v))
    
    # 2. SPACES
    lines.append("\n2. PROJECT SPACES ({})".format(len(spaces)))
    for s in sorted(spaces, key=lambda x: get_element_name(x)):
        lines.append("  SPACE: {} (ID: {})".format(get_element_name(s), s.Id))
        flat_params = get_flat_parameters(s)
        for p, v in flat_params:
            lines.append("    {}: {}".format(p, v))
    
    # 3. TYPES
    lines.append("\n3. USED SPACE TYPES ({})".format(len(used_types)))
    for t in sorted(used_types, key=lambda x: get_element_name(x)):
        lines.append("  TYPE: {} (ID: {})".format(get_element_name(t), t.Id))
        flat_params = get_flat_parameters(t)
        for p, v in flat_params:
            lines.append("    {}: {}".format(p, v))

    return "\n".join(lines)

# ==============================================================================
# MAIN
# ==============================================================================

def main():
    try:
        settings_data = collect_analytical_settings()
        spaces, used_types = collect_spaces_and_types()
        
        # HTML
        html_file = "EnergyAnalysis_Report_{}.html".format(timestamp)
        html_path = os.path.join(export_folder, html_file)
        with io.open(html_path, "w", encoding="utf-8") as f: f.write(generate_html(settings_data, spaces, used_types))
        
        # TXT
        txt_file = "EnergyAnalysis_Data_{}.txt".format(timestamp)
        txt_path = os.path.join(export_folder, txt_file)
        with io.open(txt_path, "w", encoding="utf-8") as f: f.write(generate_txt(settings_data, spaces, used_types))
        
        forms.alert("Export Complete!\n\nHTML: {}\nTXT: {}".format(html_file, txt_file), warn_icon=False)
        os.startfile(export_folder)
        
    except Exception as e:
        forms.alert("Error: {}".format(e))

if __name__ == '__main__':
    main()