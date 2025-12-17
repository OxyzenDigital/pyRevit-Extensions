# -*- coding: utf-8 -*-
"""
Final Version 4.0: 
- Naming: "Manage Rooms - Energy Analyze"
- Zoom: Forces Zoom-to-Fit on all opened views.
- Visibility: robustly enables Room/Space Interior & Reference subcategories.
"""
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (FilteredElementCollector, ViewPlan, ViewFamily, 
                               BuiltInCategory, CategoryType, ElementId, Level)

# --- Configuration ---
# Updated Naming Convention: Suffix moved to end
VIEW_CONFIGS = [
    {"suffix": "Manage Rooms - Energy Analyze",        "category": BuiltInCategory.OST_Rooms},
    {"suffix": "Manage Spaces - Energy Analyze",       "category": BuiltInCategory.OST_MEPSpaces},
    {"suffix": "Manage System Zones - Energy Analyze", "category": BuiltInCategory.OST_HVAC_Zones},
]

MODEL_WHITELIST = [
    int(BuiltInCategory.OST_Walls),
    int(BuiltInCategory.OST_Doors),
    int(BuiltInCategory.OST_Windows),
    int(BuiltInCategory.OST_StructuralColumns),
    int(BuiltInCategory.OST_StructuralFraming),
    int(BuiltInCategory.OST_CurtainWallMullions),
    int(BuiltInCategory.OST_CurtainWallPanels),
    int(BuiltInCategory.OST_Floors),
]

ANNOTATION_WHITELIST = [
    int(BuiltInCategory.OST_Grids),
]

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

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

def get_safe_model_categories():
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

def ensure_subcategories_visible(view, parent_category_id):
    """
    Robustly turns on 'Interior Fill', 'Reference', etc.
    We skip 'CanCategoryBeHidden' check which can sometimes be false for system subcats.
    """
    try:
        parent_cat = doc.Settings.Categories.get_Item(ElementId(parent_category_id))
        if parent_cat and parent_cat.SubCategories:
            for sub_cat in parent_cat.SubCategories:
                try:
                    # Force Un-hide
                    if view.GetCategoryHidden(sub_cat.Id):
                        view.SetCategoryHidden(sub_cat.Id, False)
                except:
                    # Some subcats are strictly controlled, ignore failures
                    pass
    except Exception:
        pass 

def configure_visibility(view, target_category_id, all_model_cat_ids):
    whitelist_set = set(MODEL_WHITELIST)
    whitelist_set.add(target_category_id.IntegerValue)
    
    # 1. Hide Non-Whitelisted
    for cat_id in all_model_cat_ids:
        if cat_id.IntegerValue in whitelist_set:
            continue
        try:
            if view.CanCategoryBeHidden(cat_id):
                view.SetCategoryHidden(cat_id, True)
        except:
            continue

    # 2. Force ON Whitelist + Target
    full_on_list = MODEL_WHITELIST + ANNOTATION_WHITELIST + [target_category_id.IntegerValue]
    for int_id in full_on_list:
        try:
            eid = ElementId(int_id)
            if view.CanCategoryBeHidden(eid):
                view.SetCategoryHidden(eid, False)
        except:
            continue
            
    # 3. Force Subcategories (Rooms/Spaces internals)
    ensure_subcategories_visible(view, target_category_id.IntegerValue)

def create_or_get_view(level, view_type_id, view_suffix, target_category_enum, all_model_cat_ids):
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
        target_id = ElementId(target_category_enum)
        configure_visibility(target_view, target_id, all_model_cat_ids)
    
    return target_view, is_new

def apply_zoom_to_fit(views):
    """
    Finds the UI window for each view and applies ZoomToFit.
    """
    # Get all open UI Views
    uiviews = uidoc.GetOpenUIViews()
    view_ids = [v.Id.IntegerValue for v in views]
    
    for uiv in uiviews:
        if uiv.ViewId.IntegerValue in view_ids:
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

    all_model_cat_ids = get_safe_model_categories()
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
                        config["category"],
                        all_model_cat_ids
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
        
        # Apply Zoom (Must be done after they are opened)
        apply_zoom_to_fit(views_to_open)
        
        # 4. Final MODAL Dialog
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