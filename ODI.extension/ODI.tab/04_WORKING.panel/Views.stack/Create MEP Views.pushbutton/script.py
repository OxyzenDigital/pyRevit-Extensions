# -*- coding: utf-8 -*-
from pyrevit import revit, forms, script, DB
import System.Collections.Generic as List

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# --- Configurations ---

VIEW_DEFINITIONS = [
    "_Work HVAC",
    "_Work Pipes", # All combined
    "_Work Pipes Sanitary",
    "_Work Pipes Water Supply",
    "_Work Pipes Gas",
    "_Work 3D Pipes",
    "_Work 3D HVAC"
]

# Halftone Categories (Background context)
HALFTONE_CATEGORIES = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_StructuralFraming,
    DB.BuiltInCategory.OST_StructuralColumns,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_Doors
]

# Map View Names to Disciplines/Config
VIEW_CONFIG = {
    "_Work HVAC": {
        "Discipline": DB.ViewDiscipline.Mechanical,
        "Categories": [
            DB.BuiltInCategory.OST_DuctCurves,
            DB.BuiltInCategory.OST_DuctFitting,
            DB.BuiltInCategory.OST_DuctAccessory,
            DB.BuiltInCategory.OST_DuctTerminal,
            DB.BuiltInCategory.OST_FlexDuctCurves,
            DB.BuiltInCategory.OST_MechanicalEquipment
        ]
    },
    "_Work Pipes": {
        "Discipline": DB.ViewDiscipline.Plumbing,
        "Categories": [
            DB.BuiltInCategory.OST_PipeCurves,
            DB.BuiltInCategory.OST_PipeFitting,
            DB.BuiltInCategory.OST_PipeAccessory,
            DB.BuiltInCategory.OST_FlexPipeCurves,
            DB.BuiltInCategory.OST_PlumbingFixtures,
            DB.BuiltInCategory.OST_Sprinklers,
            DB.BuiltInCategory.OST_MechanicalEquipment # Pumps etc often needed
        ]
    },
    "_Work Pipes Sanitary": {
        "Discipline": DB.ViewDiscipline.Plumbing,
        "Categories": [
            DB.BuiltInCategory.OST_PipeCurves,
            DB.BuiltInCategory.OST_PipeFitting,
            DB.BuiltInCategory.OST_PipeAccessory,
            DB.BuiltInCategory.OST_FlexPipeCurves,
            DB.BuiltInCategory.OST_PlumbingFixtures
        ]
    },
    "_Work Pipes Water Supply": {
        "Discipline": DB.ViewDiscipline.Plumbing,
        "Categories": [
            DB.BuiltInCategory.OST_PipeCurves,
            DB.BuiltInCategory.OST_PipeFitting,
            DB.BuiltInCategory.OST_PipeAccessory,
            DB.BuiltInCategory.OST_FlexPipeCurves,
            DB.BuiltInCategory.OST_PlumbingFixtures
        ]
    },
    "_Work Pipes Gas": {
        "Discipline": DB.ViewDiscipline.Plumbing,
        "Categories": [
            DB.BuiltInCategory.OST_PipeCurves,
            DB.BuiltInCategory.OST_PipeFitting,
            DB.BuiltInCategory.OST_PipeAccessory,
            DB.BuiltInCategory.OST_FlexPipeCurves,
            DB.BuiltInCategory.OST_PlumbingFixtures
        ]
    },
    "_Work 3D Pipes": {
        "Discipline": DB.ViewDiscipline.Plumbing,
        "Categories": [
            DB.BuiltInCategory.OST_PipeCurves,
            DB.BuiltInCategory.OST_PipeFitting,
            DB.BuiltInCategory.OST_PipeAccessory,
            DB.BuiltInCategory.OST_FlexPipeCurves,
            DB.BuiltInCategory.OST_PlumbingFixtures,
            DB.BuiltInCategory.OST_Sprinklers,
            DB.BuiltInCategory.OST_MechanicalEquipment
        ]
    },
    "_Work 3D HVAC": {
        "Discipline": DB.ViewDiscipline.Mechanical,
        "Categories": [
            DB.BuiltInCategory.OST_DuctCurves,
            DB.BuiltInCategory.OST_DuctFitting,
            DB.BuiltInCategory.OST_DuctAccessory,
            DB.BuiltInCategory.OST_DuctTerminal,
            DB.BuiltInCategory.OST_FlexDuctCurves,
            DB.BuiltInCategory.OST_MechanicalEquipment
        ]
    }
}

FILTER_CATEGORIES = [
    DB.BuiltInCategory.OST_PipeCurves,
    DB.BuiltInCategory.OST_PipeFitting,
    DB.BuiltInCategory.OST_PipeAccessory,
    DB.BuiltInCategory.OST_FlexPipeCurves
]

FILTERS_DEF = {
    "Cold Water": {
        "Param": DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
        "Value": "Domestic Cold Water",
        "Rule": "Contains" 
    },
    "Hot Water": {
        "Param": DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
        "Value": "Domestic Hot Water",
        "Rule": "Contains"
    },
    "Gas": {
        "Param": DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM, 
        "Value": "GAS",
        "Rule": "Contains"
    },
    "Sanitary": {
        "Param": DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
        "Value": "Sanitary",
        "Rule": "Contains"
    },
    "Vent": {
        "Param": DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
        "Value": "Vent",
        "Rule": "Contains"
    },
    "Storm Water": {
        "Param": DB.BuiltInParameter.RBS_SYSTEM_NAME_PARAM,
        "Value": "Storm",
        "Rule": "Contains"
    },
    "Hydronic": {
         "Param": DB.BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
         "Value": "Hydronic",
         "Rule": "Contains"
    }
}

def get_id_value(element_id):
    # Helper to support Revit 2024+ (Value) and older (IntegerValue)
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue

def get_view_family_type(doc, view_type=DB.ViewFamily.FloorPlan):
    try:
        return next((x for x in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType) if x.ViewFamily == view_type), None)
    except Exception:
        return None

def get_view_template():
    try:
        templates = [v for v in DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements() if v.IsTemplate]
        templates.sort(key=lambda x: x.Name)
        class NoTemplate:
            Name = "<None>"
        options = [NoTemplate()] + templates
        return forms.SelectFromList.show(options, name_attr='Name', title='Select View Template (Optional)', multiselect=False)
    except Exception:
        return None

def set_view_depth(view, level_id, offset=-5.0):
    if not view or not level_id: return
    try:
        view_range = view.GetViewRange()
        view_range.SetLevelId(DB.PlanViewPlane.ViewDepthPlane, level_id)
        view_range.SetOffset(DB.PlanViewPlane.ViewDepthPlane, offset)
        view_range.SetLevelId(DB.PlanViewPlane.BottomClipPlane, level_id)
        view_range.SetOffset(DB.PlanViewPlane.BottomClipPlane, offset)
        view.SetViewRange(view_range)
    except Exception as e:
        logger.error("Failed to set View Depth for {}: {}".format(view.Name, e))

def get_or_create_filter(doc, name, definition):
    try:
        filters = DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement).ToElements()
        existing = next((f for f in filters if f.Name == name), None)
        if existing:
            return existing
        
        cat_ids = List.List[DB.ElementId]([DB.ElementId(c) for c in FILTER_CATEGORIES])
        param_id = DB.ElementId(definition["Param"])
        
        if definition["Rule"] == "Contains":
            provider = DB.ParameterValueProvider(param_id)
            evaluator = DB.FilterStringContains()
            rule = DB.FilterStringRule(provider, evaluator, definition["Value"])
            epf = DB.ElementParameterFilter(rule)
            
            pf = DB.ParameterFilterElement.Create(doc, name, cat_ids, epf)
            return pf
    except Exception as e:
        logger.error("Failed to create filter '{}': {}".format(name, e))
        return None
    return None

def apply_filters_to_view(view, filters_map):
    if not view or not filters_map: return
    for name, pfilter in filters_map.items():
        if not pfilter: continue
        try:
            if not view.IsFilterApplied(pfilter.Id):
                view.AddFilter(pfilter.Id)
                view.SetFilterVisibility(pfilter.Id, True)
        except Exception:
            pass

def apply_visibility_overrides(doc, view, config):
    if not view or not config: return
    # Get all Model Categories
    try:
        categories = doc.Settings.Categories
    except Exception:
        return
    
    target_cats = set([int(c) for c in config.get("Categories", [])])
    halftone_cats = set([int(c) for c in HALFTONE_CATEGORIES])
    
    # Create Halftone Settings
    halftone_settings = DB.OverrideGraphicSettings()
    halftone_settings.SetHalftone(True)
    
    # Default Settings (Reset)
    reset_settings = DB.OverrideGraphicSettings()
    
    for cat in categories:
        try:
            if cat.CategoryType == DB.CategoryType.Model and cat.CanAddSubcategory:
                cat_id = cat.Id
                cat_int = get_id_value(cat_id)
                
                # Check if view can control this category
                if not view.CanCategoryBeHidden(cat_id):
                    continue

                if cat_int in target_cats:
                    # Visible, Not Halftone
                    view.SetCategoryHidden(cat_id, False)
                    view.SetCategoryOverrides(cat_id, reset_settings)
                    
                elif cat_int in halftone_cats:
                    # Visible, Halftone
                    view.SetCategoryHidden(cat_id, False)
                    view.SetCategoryOverrides(cat_id, halftone_settings)
                    
                else:
                    # Hidden (if it's not a target or background)
                    if not view.GetCategoryHidden(cat_id):
                        view.SetCategoryHidden(cat_id, True)
        except Exception:
            continue

def main():
    try:
        # 1. Determine Level
        level = None
        active_view = doc.ActiveView
        if active_view and not active_view.IsTemplate and active_view.ViewType in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan]:
             level = active_view.GenLevel
        
        all_levels = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements(), key=lambda l: l.Elevation)

        if not level:
            if not all_levels:
                 forms.alert('No levels found in the project.', exitscript=True)
            level = forms.SelectFromList.show(all_levels, name_attr='Name', title='Select Level for New Views', multiselect=False)
        
        if not level:
            return

        # 2. Select Views
        selected_view_types = forms.SelectFromList.show(
            VIEW_DEFINITIONS,
            title='Select View Types to Create/Open',
            multiselect=True,
            button_name='Select'
        )
        if not selected_view_types:
            return

        # 3. Select Template
        selected_template = get_view_template()
        is_template_selected = selected_template and hasattr(selected_template, 'Id')

        floor_plan_type = get_view_family_type(doc, DB.ViewFamily.FloorPlan)
        if not floor_plan_type:
            forms.alert("No Floor Plan View Family Type found.", exitscript=True)

        view3d_type = get_view_family_type(doc, DB.ViewFamily.ThreeDimensional)
        if not view3d_type and any("3D" in x for x in selected_view_types):
             logger.warning("No 3D View Type found. 3D views might fail to create.")

        created_views = []
        existing_views = []
        
        all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
        should_create_single_3d = len(all_levels) <= 2
        
        with revit.Transaction("Create Work Views"):
            # 3a. Ensure Filters Exist
            project_filters = {}
            for fname, fdef in FILTERS_DEF.items():
                project_filters[fname] = get_or_create_filter(doc, fname, fdef)

            for base_name in selected_view_types:
                try:
                    is_3d = "3D" in base_name
                    
                    if is_3d:
                        if should_create_single_3d:
                            view_name = base_name 
                        else:
                            view_name = "{} - {}".format(base_name, level.Name)
                    else:
                        view_name = "{} - {}".format(base_name, level.Name)

                    # Check if exists
                    existing_view = next((v for v in all_views if v.Name == view_name and not v.IsTemplate), None)
                    target_view = existing_view
                    
                    if not target_view:
                        if is_3d:
                            if view3d_type:
                                target_view = DB.View3D.CreateIsometric(doc, view3d_type.Id)
                            else:
                                logger.error("Cannot create 3D view '{}' because no 3D View Type is available.".format(view_name))
                                continue
                        else:
                            target_view = DB.ViewPlan.Create(doc, floor_plan_type.Id, level.Id)
                            set_view_depth(target_view, level.Id, offset=-5.0)

                        if target_view:
                            target_view.Name = view_name
                            created_views.append(target_view)
                            
                            # Apply Template (Properties only)
                            if is_template_selected:
                                try:
                                    target_view.ApplyViewTemplateParameters(selected_template)
                                except Exception as e:
                                    logger.error("Failed to apply template to {}: {}".format(view_name, e))
                            
                            if not is_template_selected:
                                 # Set Detail Level
                                try: target_view.DetailLevel = DB.ViewDetailLevel.Fine
                                except: pass
                                
                                apply_visibility_overrides(doc, target_view, VIEW_CONFIG.get(base_name, {}))

                    if target_view:
                        if existing_view:
                            existing_views.append(existing_view)
                        
                        # --- ENFORCE DISCIPLINE ---
                        config = VIEW_CONFIG.get(base_name)
                        if config:
                            disc = config.get("Discipline")
                            if disc:
                                try: target_view.Discipline = disc
                                except: pass
                        
                        # --- APPLY FILTERS ---
                        apply_filters_to_view(target_view, project_filters)
                except Exception as e:
                     logger.error("Error processing view '{}': {}".format(base_name, e))

        # 4. Report and Open
        all_target_views = created_views + existing_views
        
        if not all_target_views:
            forms.alert("No views created or found.")
            return

        # Open ALL views
        for v in all_target_views:
            try:
                uidoc.ActiveView = v
            except Exception:
                pass 
            
        # Summary Output
        output = script.get_output()
        if created_views:
            output.print_md("### Created Views:")
            for v in created_views:
                print("- " + v.Name)
        
        if existing_views:
            output.print_md("### Existing Views (Opened):")
        for v in existing_views:
            print("- " + v.Name)

    except Exception as e:
        logger.error("Critical Error in Main Loop: {}".format(e))
        forms.alert("An unexpected error occurred. See output for details.", exitscript=True)

if __name__ == '__main__':
    main()