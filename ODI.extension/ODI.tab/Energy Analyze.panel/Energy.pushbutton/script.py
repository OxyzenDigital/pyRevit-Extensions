# -*- coding: utf-8 -*-
"""
Smart Version 3.2: 
- Uses a custom wrapper class to handle UI display names safely.
- Resolves 'readonly attribute' and 'unexpected keyword' errors.
- Ensures Interior Fill and Reference visibility.
"""
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (FilteredElementCollector, ViewPlan, ViewFamily, 
                               BuiltInCategory, CategoryType, ElementId, Level)

# --- Configuration ---
VIEW_CONFIGS = [
    {"suffix": "Energy Analyze - Manage Rooms",        "category": BuiltInCategory.OST_Rooms},
    {"suffix": "Energy Analyze - Manage Spaces",       "category": BuiltInCategory.OST_MEPSpaces},
    {"suffix": "Energy Analyze - Manage System Zones", "category": BuiltInCategory.OST_HVAC_Zones},
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
output = script.get_output()

# --- Custom UI Wrapper Class ---
class LevelUIItem(object):
    """
    A simple wrapper class to control how Levels appear in the list.
    We use this instead of TemplateListItem to avoid readonly/constructor errors.
    """
    def __init__(self, level_element, is_existing):
        self.element = level_element
        # This 'name' attribute is what pyRevit displays in the list
        if is_existing:
            self.name = "{}   --   [Existing Views]".format(level_element.Name)
        else:
            self.name = "{}   --   [Create New]".format(level_element.Name)
        
        # Default all items to checked
        self.checked = True

    # This ensures the name is displayed correctly in all contexts
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
    """Returns True if all 3 Energy views exist for this level."""
    collector = FilteredElementCollector(doc).OfClass(ViewPlan)
    existing_names = set([v.Name for v in collector])
    
    for config in VIEW_CONFIGS:
        target_name = "{} - {}".format(level.Name, config["suffix"])
        if target_name not in existing_names:
            return False
    return True

def ensure_subcategories_visible(view, parent_category_id):
    """Forces 'Interior Fill' and 'Reference' to be visible."""
    try:
        parent_cat = doc.Settings.Categories.get_Item(ElementId(parent_category_id))
        if parent_cat and parent_cat.SubCategories:
            for sub_cat in parent_cat.SubCategories:
                if view.CanCategoryBeHidden(sub_cat.Id):
                    if view.GetCategoryHidden(sub_cat.Id):
                        view.SetCategoryHidden(sub_cat.Id, False)
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
            
    # 3. Subcategories
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

def print_green_header(title, subtitle):
    html = """
    <div style="background-color:#27ae60; padding:15px; margin-bottom:15px; border-radius:4px; box-shadow:0 2px 5px rgba(0,0,0,0.2);">
        <h2 style="color:white; margin:0; font-family:Segoe UI;">{}</h2>
        <p style="color:white; margin:5px 0 0 0; opacity:0.9;">{}</p>
    </div>
    """.format(title, subtitle)
    output.print_html(html)

# --- Main ---

def main():
    # 1. Collect Levels
    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    all_levels = sorted(all_levels, key=lambda x: x.Elevation)
    
    if not all_levels:
        forms.alert("No levels found.", exitscript=True)

    # 2. Build List with Custom UI Item
    ui_items = []
    for lvl in all_levels:
        is_complete = check_level_status(lvl)
        # Create our custom wrapper object
        ui_items.append(LevelUIItem(lvl, is_complete))

    # 3. Show Dialog
    # SelectFromList returns a list of the LevelUIItem objects we created
    selected_items = forms.SelectFromList.show(
        ui_items,
        title="Select Levels for Energy Analysis",
        multiselect=True,
        button_name="Process Views"
    )

    if not selected_items:
        return

    # 4. Process
    view_type_id = get_floor_plan_type()
    if not view_type_id:
        forms.alert("No Floor Plan Type found.", exitscript=True)

    all_model_cat_ids = get_safe_model_categories()
    views_to_open = []
    created_count = 0

    with revit.Transaction("Process Energy Views"):
        for item in selected_items:
            # Unwrap the level from our custom item
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

    # 5. Open Views
    if views_to_open:
        views_to_open.sort(key=lambda x: x.Name)
        
        for v in views_to_open:
            try:
                uidoc.RequestViewChange(v)
            except:
                pass
        
        print_green_header("Energy Analysis Views", "Process Complete")
        print("Total Views Active: {}".format(len(views_to_open)))
        print("New Views Created: {}".format(created_count))

if __name__ == '__main__':
    main()