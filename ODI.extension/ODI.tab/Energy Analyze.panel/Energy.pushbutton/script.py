# -*- coding: utf-8 -*-
"""
Universal Version 5.0 (Revit 2023-2026+ Compatible)
- Replaces '.IntegerValue' with a version-safe ID check.
- Uses BuiltInCategory directly to avoid Integer conversion errors.
- Robust visibility handling for all versions.
"""
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (FilteredElementCollector, ViewPlan, ViewFamily, 
                               BuiltInCategory, CategoryType, ElementId, Level)

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# --- COMPATIBILITY HELPER ---
def get_id_val(element_id):
    """
    Universal helper to get the Integer/Long value of an ElementId.
    Revit 2024+ uses .Value (Long)
    Revit <2024 uses .IntegerValue (Int)
    """
    try:
        # Try newer API first (2024/2025/2026)
        return element_id.Value
    except AttributeError:
        # Fallback to older API
        return element_id.IntegerValue

# --- Configuration ---
# Keeping BICs as Objects, not Ints, to be safe across versions
VIEW_CONFIGS = [
    {"suffix": "Manage Rooms - Energy Analyze",        "bic": BuiltInCategory.OST_Rooms},
    {"suffix": "Manage Spaces - Energy Analyze",       "bic": BuiltInCategory.OST_MEPSpaces},
    {"suffix": "Manage System Zones - Energy Analyze", "bic": BuiltInCategory.OST_HVAC_Zones},
]

MODEL_WHITELIST_BIC = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_CurtainWallMullions,
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_Floors,
]

ANNOTATION_WHITELIST_BIC = [
    BuiltInCategory.OST_Grids,
]

# --- Custom UI Wrapper ---
class LevelUIItem(object):
    def __init__(self, level_element, is_existing):
        self.element = level_element
        if is_existing:
            self.name = "{}   --   [Existing Views]".format(level_element.Name)
        else:
            self.name = "{}   --   [Create New]".format(level_element.Name)
        self.checked = True

    def __repr__(self):
        return self.name

# --- Modules ---

def get_floor_plan_type():
    types = FilteredElementCollector(doc).OfClass(revit.DB.ViewFamilyType).ToElements()
    for t in types:
        if t.ViewFamily == ViewFamily.FloorPlan:
            return t.Id
    return None

def get_safe_model_category_ids():
    """
    Returns a list of ElementIds for all Model Categories.
    """
    safe_ids = []
    categories = doc.Settings.Categories
    for cat in categories:
        try:
            if cat.CategoryType == CategoryType.Model:
                safe_ids.append(cat.Id)
        except:
            continue
    return safe_ids

def check_level_status(level):
    collector = FilteredElementCollector(doc).OfClass(ViewPlan)
    existing_names = set([v.Name for v in collector])
    for config in VIEW_CONFIGS:
        target_name = "{} - {}".format(level.Name, config["suffix"])
        if target_name not in existing_names:
            return False
    return True

def ensure_subcategories_visible(view, parent_bic):
    """
    Robustly turns on subcategories.
    Uses ElementId(BuiltInCategory) to ensure 2026 compatibility.
    """
    try:
        parent_id = ElementId(parent_bic)
        parent_cat = doc.Settings.Categories.get_Item(parent_id)
        if parent_cat and parent_cat.SubCategories:
            for sub_cat in parent_cat.SubCategories:
                try:
                    if view.GetCategoryHidden(sub_cat.Id):
                        view.SetCategoryHidden(sub_cat.Id, False)
                except:
                    pass
    except Exception:
        pass 

def configure_visibility(view, target_bic, all_model_ids):
    # 1. Prepare Whitelist Set (using Universal ID Values)
    whitelist_vals = set()
    
    # Add Standard Model Whitelist
    for bic in MODEL_WHITELIST_BIC:
        try:
            eid = ElementId(bic)
            whitelist_vals.add(get_id_val(eid))
        except: pass
        
    # Add Target Category
    target_eid = ElementId(target_bic)
    whitelist_vals.add(get_id_val(target_eid))
    
    # 2. Hide Non-Whitelisted
    for cat_id in all_model_ids:
        # Compare Values, not Objects
        if get_id_val(cat_id) in whitelist_vals:
            continue
        try:
            if view.CanCategoryBeHidden(cat_id):
                view.SetCategoryHidden(cat_id, True)
        except:
            continue

    # 3. Force ON Whitelist + Target + Annotation
    # Combine lists
    full_on_bics = MODEL_WHITELIST_BIC + ANNOTATION_WHITELIST_BIC + [target_bic]
    
    for bic in full_on_bics:
        try:
            eid = ElementId(bic)
            if view.CanCategoryBeHidden(eid):
                view.SetCategoryHidden(eid, False)
        except:
            continue
            
    # 4. Force Subcategories
    ensure_subcategories_visible(view, target_bic)

def create_or_get_view(level, view_type_id, view_suffix, target_bic, all_model_ids):
    view_name = "{} - {}".format(level.Name, view_suffix)
    is_new = False
    
    target_view = None
    collector = FilteredElementCollector(doc).OfClass(revit.DB.ViewPlan)
    for v in collector:
        if v.Name == view_name and not v.IsTemplate:
            target_view = v
            break

    if not target_view:
        try:
            target_view = ViewPlan.Create(doc, view_type_id, level.Id)
            target_view.Name = view_name
            is_new = True
        except:
            pass

    if target_view:
        configure_visibility(target_view, target_bic, all_model_ids)
    
    return target_view, is_new

def apply_zoom_to_fit(views):
    uiviews = uidoc.GetOpenUIViews()
    # Create set of ID Values for the views we just touched
    target_ids = set([get_id_val(v.Id) for v in views])
    
    for uiv in uiviews:
        # Check if this UI View matches one of ours
        if get_id_val(uiv.ViewId) in target_ids:
            try:
                uiv.ZoomToFit()
            except:
                pass

# --- Main ---

def main():
    # 1. UI Setup
    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    all_levels = sorted(all_levels, key=lambda x: x.Elevation)
    
    if not all_levels:
        forms.alert("No levels found.", exitscript=True)

    ui_items = []
    for lvl in all_levels:
        is_complete = check_level_status(lvl)
        ui_items.append(LevelUIItem(lvl, is_complete))

    selected_items = forms.SelectFromList.show(
        ui_items,
        title="Energy Analysis | Select Levels",
        multiselect=True,
        button_name="Process Views"
    )

    if not selected_items:
        return

    # 2. Processing
    view_type_id = get_floor_plan_type()
    if not view_type_id:
        forms.alert("No Floor Plan Type found.", exitscript=True)

    all_model_ids = get_safe_model_category_ids()
    views_to_open = []
    created_count = 0

    with revit.Transaction("Process Energy Views"):
        for item in selected_items:
            level = item.element
            for config in VIEW_CONFIGS:
                try:
                    view, is_new = create_or_get_view(
                        level, 
                        view_type_id, 
                        config["suffix"], 
                        config["bic"],
                        all_model_ids
                    )
                    if view:
                        views_to_open.append(view)
                        if is_new:
                            created_count += 1
                except Exception as e:
                    logger.error("Error on {}: {}".format(level.Name, e))

    # 3. View Activation & Zoom
    if views_to_open:
        views_to_open.sort(key=lambda x: x.Name)
        
        # Open them
        for v in views_to_open:
            try:
                uidoc.RequestViewChange(v)
            except:
                pass
        
        # Apply Zoom
        apply_zoom_to_fit(views_to_open)
        
        # 4. Final Dialog
        result_message = "Process Complete.\n\nTotal Views Active: {}\nNew Views Created: {}".format(
            len(views_to_open), 
            created_count
        )
        
        forms.alert(
            result_message, 
            title="Energy Analysis Manager", 
            warn_icon=False
        )

if __name__ == '__main__':
    main()