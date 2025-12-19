# -*- coding: utf-8 -*-
"""
Energy Analysis Master Exporter (v16.0)
- FIXED: Space Type names now populate correctly in headers.
- NEW: Sub-Headers (Parameter Groups) are now collapsible accordions.
- ARCHITECTURE: Collect -> Map -> UI -> Export.
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
# PHASE 1: HELPERS & SAFETY
# ==============================================================================

def get_bip(name):
    if hasattr(DB.BuiltInParameter, name):
        return getattr(DB.BuiltInParameter, name)
    return None

def get_bip_string(element, bip_name):
    """Safely gets a string parameter by BIP Name."""
    bip = get_bip(bip_name)
    if bip:
        p = element.get_Parameter(bip)
        if p and p.HasValue:
            return p.AsString() or p.AsValueString()
    return None

def get_element_name(element):
    """
    Robust Naming Logic (Updated for Types).
    """
    if not element: return "Unknown"
    
    # 1. Try SPACE/ROOM logic (Number + Name)
    r_num = get_bip_string(element, "ROOM_NUMBER")
    r_name = get_bip_string(element, "ROOM_NAME")
    if r_num and r_name:
        return "{} - {}".format(r_num, r_name)

    # 2. Try TYPE Name (Critical for Space Types)
    # This is usually the best bet for OST_HVAC_Load_Space_Types
    t_name = get_bip_string(element, "ALL_MODEL_TYPE_NAME")
    if t_name: return t_name

    # 3. Try SYMBOL Name
    s_name = get_bip_string(element, "SYMBOL_NAME_PARAM")
    if s_name: return s_name

    # 4. Fallback to Property if API allows
    try:
        if hasattr(element, "Name") and element.Name:
            return element.Name
    except: pass

    # 5. Last Resort
    return "Element ID: {}".format(element.Id)

def get_safe_value(p):
    if not p or not p.HasValue: return ""
    try:
        if p.StorageType == DB.StorageType.ElementId:
            eid = p.AsElementId()
            if eid == DB.ElementId.InvalidElementId: return "<None>"
            e = doc.GetElement(eid)
            return get_element_name(e)
        elif p.StorageType == DB.StorageType.String:
            return p.AsString()
        elif p.StorageType == DB.StorageType.Integer:
            val = p.AsValueString()
            return val if val else str(p.AsInteger())
        elif p.StorageType == DB.StorageType.Double:
            val = p.AsValueString()
            return val if val else "{:.2f}".format(p.AsDouble())
    except:
        return "Error"
    return ""

def get_grouped_parameters(element):
    if not element: return {}
    groups = {}
    
    for p in element.Parameters:
        if not p.Definition: continue
        if not p.HasValue: continue
        
        g_name = "Other Data"
        try:
            if hasattr(p.Definition, "GetGroupTypeId"): # Revit 2024+
                g_name = DB.LabelUtils.GetLabelFor(p.Definition.GetGroupTypeId())
            else:
                g_name = str(p.Definition.ParameterGroup).replace("PG_", "").replace("_", " ").title()
        except: pass
            
        val = get_safe_value(p)
        if val:
            if g_name not in groups: groups[g_name] = []
            groups[g_name].append( (p.Definition.Name, val) )
            
    for g in groups:
        groups[g].sort(key=lambda x: x[0])
        
    return groups

# ==============================================================================
# PHASE 2: DEEP COLLECTION
# ==============================================================================

def collect_database():
    db = { "spaces": {}, "space_types": {}, "building_types": {}, "constructions": {} }
    
    # 1. Spaces
    spaces = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MEPSpaces).ToElements()
    for s in spaces: db["spaces"][s.Id] = s
        
    # 2. Space Types (Library)
    st_cats = ["OST_HVAC_Load_Space_Types", "OST_MEPSpaceType"]
    for cat in st_cats:
        if hasattr(DB.BuiltInCategory, cat):
            bic = getattr(DB.BuiltInCategory, cat)
            elems = DB.FilteredElementCollector(doc).OfCategory(bic).ToElements()
            for e in elems: db["space_types"][e.Id] = e

    # 3. Building Types (Library)
    bt_cat = getattr(DB.BuiltInCategory, "OST_HVAC_Load_Building_Types", None)
    if bt_cat:
        elems = DB.FilteredElementCollector(doc).OfCategory(bt_cat).ToElements()
        for e in elems: db["building_types"][e.Id] = e

    # 4. Constructions (Library)
    c_cats = ["OST_EAConstructions", "OST_MEPBuildingConstruction", "OST_EnergyAnalysisConstruction"]
    for cat in c_cats:
        if hasattr(DB.BuiltInCategory, cat):
            bic = getattr(DB.BuiltInCategory, cat)
            elems = DB.FilteredElementCollector(doc).OfCategory(bic).ToElements()
            for e in elems: db["constructions"][e.Id] = e
            
    return db

# ==============================================================================
# PHASE 3: RELATIONAL MAPPING
# ==============================================================================

def map_usage(db_index):
    used_st = set()
    used_bt = set()
    used_cs = set()
    
    # A. Global Settings
    try:
        es = DB.Analysis.EnergyDataSettings.GetFromDocument(doc)
        if es:
            p = es.get_Parameter(get_bip("RBS_ENERGY_BUILDING_TYPE"))
            if p and p.HasValue: used_bt.add(p.AsElementId())
            
            p = es.get_Parameter(get_bip("RBS_ENERGY_CONSTRUCTION_SET"))
            if p and p.HasValue: used_cs.add(p.AsElementId())
    except: pass
        
    # B. Space Overrides
    for s_id, space in db_index["spaces"].items():
        # Space Type
        p = space.get_Parameter(get_bip("ROOM_SPACE_TYPE_PARAM"))
        if p and p.HasValue: used_st.add(p.AsElementId())
            
        # Construction Set
        for pname in ["RBS_ENERGY_CONSTRUCTION_SET", "ROOM_CONSTRUCTION_SET_PARAM"]:
            bip = get_bip(pname)
            if bip:
                p = space.get_Parameter(bip)
                if p and p.HasValue: used_cs.add(p.AsElementId())
                
    return used_st, used_bt, used_cs

# ==============================================================================
# PHASE 4: HTML GENERATION (NESTED ACCORDIONS)
# ==============================================================================

def generate_html(final_data, mode_title):
    css = """
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #f4f4f9; padding: 40px; color: #333; }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        .timestamp { color: #7f8c8d; margin-bottom: 30px; font-size: 0.9em; }
        
        /* Main Accordion (Level 1) */
        .accordion {
            background-color: #fff; color: #2c3e50; cursor: pointer; padding: 18px; width: 100%;
            border: 1px solid #ddd; border-left: 6px solid #3498db; text-align: left;
            outline: none; font-size: 16px; font-weight: bold; transition: 0.3s;
            display: flex; justify-content: space-between; align-items: center; margin-top: 15px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        }
        .active, .accordion:hover { background-color: #eaf6ff; }
        .accordion:after { content: '+'; font-size: 20px; font-weight: bold; }
        .active:after { content: '-'; }
        
        /* Panel Content (Level 1) */
        .panel {
            padding: 0 18px; background-color: white; max-height: 0; overflow: hidden;
            transition: max-height 0.2s ease-out; border: 1px solid #ddd; border-top: none; margin-bottom: 10px;
        }
        
        /* Item Blocks */
        .item-block {
            border: 1px solid #eee; margin: 20px 0; border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05); overflow: hidden;
        }
        .item-header {
            background: #f8f9fa; padding: 12px 15px; font-weight: bold; color: #333;
            border-bottom: 1px solid #eee; display: flex; justify-content: space-between;
        }

        /* Sub-Accordion (Level 2 - Parameter Groups) */
        .sub-accordion {
            background-color: #fbfbfb; color: #555; cursor: pointer; padding: 10px 15px; width: 100%;
            border: none; border-bottom: 1px solid #eee; text-align: left; outline: none;
            font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
            transition: 0.2s; display: flex; justify-content: space-between; align-items: center;
        }
        .sub-accordion:hover { background-color: #f0f0f0; }
        .sub-accordion:after { content: '+'; font-size: 14px; font-weight: bold; color: #999; }
        .sub-active:after { content: '-'; }
        
        /* Sub-Panel (Level 2) */
        .sub-panel {
            padding: 0 15px; background-color: #fff; max-height: 0; overflow: hidden;
            transition: max-height 0.2s ease-out; border-bottom: 1px solid #eee;
        }
        
        table { width: 100%; border-collapse: collapse; font-size: 0.9em; margin: 10px 0; }
        th, td { text-align: left; padding: 6px 0; border-bottom: 1px solid #f9f9f9; }
        th { width: 40%; color: #888; font-weight: normal; }
        tr:last-child td { border-bottom: none; }
    </style>
    """
    
    js = """
    <script>
        document.addEventListener("DOMContentLoaded", function() {
            // Function to toggle any accordion (Main or Sub)
            function toggleAccordion() {
                this.classList.toggle("active");
                if (this.classList.contains("sub-accordion")) {
                    this.classList.toggle("sub-active");
                }
                
                var panel = this.nextElementSibling;
                if (panel.style.maxHeight) {
                    panel.style.maxHeight = null;
                } else {
                    panel.style.maxHeight = panel.scrollHeight + "px";
                    
                    // If we are opening a sub-panel, we need to grow the parent panel too!
                    var parentPanel = this.closest('.panel');
                    if (parentPanel) {
                        parentPanel.style.maxHeight = parentPanel.scrollHeight + panel.scrollHeight + "px";
                    }
                } 
            }

            var acc = document.getElementsByClassName("accordion");
            for (var i = 0; i < acc.length; i++) {
                acc[i].addEventListener("click", toggleAccordion);
            }
            
            var subAcc = document.getElementsByClassName("sub-accordion");
            for (var j = 0; j < subAcc.length; j++) {
                subAcc[j].addEventListener("click", toggleAccordion);
            }
        });
    </script>
    """
    
    html = [u"<html><head><meta charset='utf-8'>{}{}</head><body>".format(css, js)]
    html.append(u"<h1>Energy Analysis: {}</h1>".format(mode_title))
    html.append(u"<p class='timestamp'>Report Generated: {}</p>".format(timestamp))
    
    order = [
        ("1. Global Settings", "settings"),
        ("2. Project Spaces", "spaces"),
        ("3. Space Types", "space_types"),
        ("4. Building Types", "building_types"),
        ("5. Constructions", "constructions")
    ]
    
    for title, key in order:
        elements = final_data.get(key, [])
        count = len(elements)
        
        html.append(u"<button class='accordion'>{} ({})</button>".format(title, count))
        html.append(u"<div class='panel'>")
        
        if not elements:
            html.append(u"<p style='padding:20px; color:#aaa; font-style:italic;'>No data in this category.</p>")
        else:
            sorted_elems = sorted(elements, key=lambda x: get_element_name(x))
            html.append(u"<div style='padding:15px 0;'>")
            
            for el in sorted_elems:
                name = get_element_name(el)
                groups = get_grouped_parameters(el)
                
                html.append(u"<div class='item-block'>")
                html.append(u"<div class='item-header'><span>{}</span> <span style='color:#999; font-weight:normal; font-size:0.9em;'>ID: {}</span></div>".format(name, el.Id))
                
                for g_name, params in sorted(groups.items()):
                    # Sub-Accordion Button
                    html.append(u"<button class='sub-accordion'>{}</button>".format(g_name))
                    # Sub-Panel Content
                    html.append(u"<div class='sub-panel'>")
                    html.append(u"<table>")
                    for p_name, p_val in params:
                        html.append(u"<tr><th>{}</th><td>{}</td></tr>".format(p_name, p_val))
                    html.append(u"</table>")
                    html.append(u"</div>")
                    
                html.append(u"</div>")
            html.append(u"</div>")
        html.append(u"</div>")
        
    html.append(u"</body></html>")
    return u"".join(html)

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    options = ["Export USED Settings Only (Clean)", "Export ALL Library Items (Audit)"]
    res = forms.ask_for_one_item(
        options,
        default="Export USED Settings Only (Clean)",
        prompt="Select Export Mode:",
        title="Energy Analysis Exporter"
    )
    
    if not res: return
    
    is_used_only = "Clean" in res
    mode_title = "Active Project Settings" if is_used_only else "Full Library Audit"
    
    # 1. Collect
    db_index = collect_database()
    
    # 2. Map Usage
    used_st_ids, used_bt_ids, used_cs_ids = map_usage(db_index)
    
    # 3. Filter
    final_data = {
        "settings": [],
        "spaces": list(db_index["spaces"].values()),
        "space_types": [],
        "building_types": [],
        "constructions": []
    }
    
    try:
        es = DB.Analysis.EnergyDataSettings.GetFromDocument(doc)
        if es: final_data["settings"] = [es]
    except: pass
    
    if is_used_only:
        final_data["space_types"] = [e for eid, e in db_index["space_types"].items() if eid in used_st_ids]
        final_data["building_types"] = [e for eid, e in db_index["building_types"].items() if eid in used_bt_ids]
        final_data["constructions"] = [e for eid, e in db_index["constructions"].items() if eid in used_cs_ids]
    else:
        final_data["space_types"] = list(db_index["space_types"].values())
        final_data["building_types"] = list(db_index["building_types"].values())
        final_data["constructions"] = list(db_index["constructions"].values())

    # 4. Generate & Save
    filename = "EnergySettings_{}_{}.html".format("Used" if is_used_only else "All", timestamp)
    save_path = os.path.join(export_folder, filename)
    
    try:
        html_content = generate_html(final_data, mode_title)
        
        with io.open(save_path, "w", encoding="utf-8") as f:
            f.write(html_content)
            
        res = forms.alert(
            "Export Successful!\nMode: {}\nFile: {}".format(mode_title, filename),
            title="Energy Exporter",
            warn_icon=False
        )
        os.startfile(export_folder)
        
    except Exception as e:
        forms.alert("Error writing file: {}".format(e))

if __name__ == '__main__':
    main()