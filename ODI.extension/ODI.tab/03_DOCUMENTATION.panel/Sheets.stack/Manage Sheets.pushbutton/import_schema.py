# -*- coding=utf-8 -*-
"""Manage Sheets - Importer.
Loads the execution intent JSON, runs the conflict diff engine, checks and binds
the shared parameter, launches the WPF reviewer window, and executes the
validated transactions.
"""

import os
import sys
import json
import System
from pyrevit import forms, script

# Revit API imports
try:
    import Autodesk
    from pyrevit import revit
    doc = revit.doc
    app = revit.doc.Application
    is_revit = True
except ImportError:
    doc = None
    app = None
    is_revit = False

if is_revit:
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        ViewSheet,
        ViewSheet,
        Transaction,
        ElementId,
        FamilySymbol,
        BuiltInCategory,
        BuiltInParameter,
        InstanceBinding,
        ExternalDefinitionCreationOptions
    )
    
    # Handle Revit 2025 parameter group binding changes
    try:
        from Autodesk.Revit.DB import GroupTypeId
        has_group_type_id = True
    except ImportError:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        has_group_type_id = False

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
    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString() or ""
    except:
        pass
    return ""

# --- AIA DISCIPLINE MAPPING ---

DISCIPLINE_MAP = {
    "CS": "General",
    "G":  "General",
    "H":  "Hazardous Materials",
    "V":  "Survey / Mapping",
    "B":  "Geotechnical",
    "C":  "Civil",
    "L":  "Landscape",
    "S":  "Structural",
    "A":  "Architectural",
    "I":  "Interiors",
    "Q":  "Equipment",
    "F":  "Fire Protection",
    "P":  "Plumbing",
    "D":  "Process",
    "M":  "Mechanical",
    "E":  "Electrical",
    "T":  "Telecommunications",
    "R":  "Resource",
    "X":  "Other Disciplines",
    "Z":  "Contractor / Shop Drawings",
    "O":  "Operations",
}

def get_aia_discipline(designator):
    """Returns the AIA full-name discipline for a given designator code.
    Handles compound codes like 'CS' before single-letter fallback.
    """
    if not designator:
        return "General"
    d = designator.strip().upper()
    # Try full code first (e.g. 'CS'), then first letter
    return DISCIPLINE_MAP.get(d, DISCIPLINE_MAP.get(d[:1], "General"))


# --- MVVM BINDING CLASSES ---

class CreateItem(object):
    def __init__(self, number, name, titleblock, titleblock_id, sheet_collection,
                 schema_link, designator_code, sheet_discipline="", sheet_use=""):
        self.number = number
        self.name = name
        self.titleblock = titleblock
        self.titleblock_id = titleblock_id
        self.sheet_collection = sheet_collection
        self.schema_link = schema_link
        self.designator_code = designator_code
        self.sheet_discipline = sheet_discipline
        self.sheet_use = sheet_use
        
    @property
    def Number(self): return self.number
    @property
    def Name(self): return self.name
    @property
    def TitleBlock(self): return self.titleblock
    @property
    def SheetCollection(self): return self.sheet_collection

class ConflictItem(object):
    def __init__(self, element, number, model_name, proposed_name, reason, proposed_sheet, schema_link):
        self.element = element
        self.number = number
        self.model_name = model_name
        self.proposed_name = proposed_name
        self.reason = reason
        self.proposed_sheet = proposed_sheet # dict from json
        self.schema_link = schema_link
        self._selected_action = "Overwrite"
        
    @property
    def Number(self): return self.number
    @property
    def ModelName(self): return self.model_name
    @property
    def ProposedName(self): return self.proposed_name
    @property
    def Reason(self): return self.reason
    
    @property
    def SelectedAction(self): return self._selected_action
    @SelectedAction.setter
    def SelectedAction(self, value):
        self._selected_action = value

class PurgeItem(object):
    def __init__(self, element, number, name):
        self.element = element
        self.number = number
        self.name = name
        self._should_purge = True
        
    @property
    def Number(self): return self.number
    @property
    def Name(self): return self.name
    
    @property
    def ShouldPurge(self): return self._should_purge
    @ShouldPurge.setter
    def ShouldPurge(self, value):
        self._should_purge = value

# --- WPF WINDOW CLASS ---

class ReviewerWindow(forms.WPFWindow):
    def __init__(self, xaml_path, create_items, conflict_items, purge_items):
        forms.WPFWindow.__init__(self, xaml_path)
        self.create_items = create_items
        self.conflict_items = conflict_items
        self.purge_items = purge_items
        self.cancelled = True
        
        # MVVM Data Bindings
        self.CreateListView.ItemsSource = self.create_items
        self.ConflictListView.ItemsSource = self.conflict_items
        self.PurgeListView.ItemsSource = self.purge_items
        
        # Programmatic Event Bindings for CPython compatibility
        self.CancelBtn.Click += self.CancelClick
        self.ExecuteBtn.Click += self.ExecuteClick
        
    def CancelClick(self, sender, e):
        self.cancelled = True
        self.Close()
        
    def ExecuteClick(self, sender, e):
        self.cancelled = False
        self.Close()

# --- HELPER FUNCTIONS ---

def ensure_shared_parameter(doc, app):
    """Verifies if the shared parameters ODI_Schema_Link, ODI_Building_Type,
    ODI_Designator_Code, SHEETS - Discipline, and SHEETS - Use exist and are
    bound to the correct categories. Creates and binds any that are missing.
    """
    if not doc or not app:
        return False

    # Check bindings
    binding_map = doc.ParameterBindings
    it = binding_map.ForwardIterator()
    bound_to_sheets = False
    bound_to_proj_info = False
    bound_designator_code = False
    bound_disc = False
    bound_use = False

    while it.MoveNext():
        definition = it.Key
        if definition.Name == "ODI_Schema_Link":
            binding = it.Current
            if isinstance(binding, InstanceBinding):
                for cat in binding.Categories:
                    if get_id_value(cat.Id) == int(BuiltInCategory.OST_Sheets):
                        bound_to_sheets = True
                        break
        elif definition.Name == "ODI_Building_Type":
            binding = it.Current
            if isinstance(binding, InstanceBinding):
                for cat in binding.Categories:
                    if get_id_value(cat.Id) == int(BuiltInCategory.OST_ProjectInformation):
                        bound_to_proj_info = True
                        break
        elif definition.Name == "ODI_Designator_Code":
            binding = it.Current
            if isinstance(binding, InstanceBinding):
                for cat in binding.Categories:
                    if get_id_value(cat.Id) == int(BuiltInCategory.OST_Sheets):
                        bound_designator_code = True
                        break
        elif definition.Name == "SHEETS - Discipline":
            binding = it.Current
            if isinstance(binding, InstanceBinding):
                for cat in binding.Categories:
                    if get_id_value(cat.Id) == int(BuiltInCategory.OST_Sheets):
                        bound_disc = True
                        break
        elif definition.Name == "SHEETS - Use":
            binding = it.Current
            if isinstance(binding, InstanceBinding):
                for cat in binding.Categories:
                    if get_id_value(cat.Id) == int(BuiltInCategory.OST_Sheets):
                        bound_use = True
                        break

    if bound_to_sheets and bound_to_proj_info and bound_designator_code and bound_disc and bound_use:
        return True
        
    logger.info("Parameters missing or unbound. Creating required bindings...")
    
    # 1. Setup shared parameter file
    sp_file = None
    try:
        sp_file = app.OpenSharedParameterFile()
    except System.Exception as e:
        logger.warning("CLR Exception reading current Shared Parameter File: {}. Re-creating fallback...".format(e.Message))
        sp_file = None
    except Exception as e:
        logger.warning("Error reading current Shared Parameter File: {}. Re-creating fallback...".format(e))
        sp_file = None
        
    if not sp_file:
        import tempfile
        temp_sp_path = os.path.join(tempfile.gettempdir(), "ODI_SharedParameters.txt")
        try:
            # Force UTF-16 LE with BOM (strictly required by Revit's internal parameter parser)
            content = (
                u"# This is a Revit shared parameter file.\r\n"
                u"# Do not edit manually.\r\n"
                u"*META\tVERSION\tMINVERSION\r\n"
                u"META\t2\t1\r\n"
                u"*GROUP\tID\tNAME\r\n"
                u"GROUP\t1\tOxyzenDigital\r\n"
                u"*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\r\n"
                u"PARAM\t7f8a9b1c-3d2e-4f5a-6b7c-8d9e0a1b2c3d\tODI_Schema_Link\tTEXT\t\t1\t1\tUnique hash linking sheet to schema\t1\r\n"
                u"PARAM\t8a7b6c5d-4e3d-2c1b-0a9b-8c7d6e5f4a3b\tODI_Building_Type\tTEXT\t\t1\t1\tAssigned project library type\t1\r\n"
                u"PARAM\t9b8c7d6e-5f4a-3b2c-1a0b-9c8d7e6f5a4b\tODI_Designator_Code\tTEXT\t\t1\t1\tDesignator sorting and sequencing code\t1\r\n"
                u"PARAM\tac1b2c3d-4e5f-6a7b-8c9d-0e1f2a3b4c5d\tSHEETS - Discipline\tTEXT\t\t1\t1\tAIA discipline name for Project Browser grouping\t1\r\n"
                u"PARAM\tbd2c3d4e-5f6a-7b8c-9d0e-1f2a3b4c5d6e\tSHEETS - Use\tTEXT\t\t1\t1\tUser-defined sheet use category for Project Browser\t1\r\n"
            )
            # Remove existing file if any to prevent write conflicts or lockups
            if os.path.exists(temp_sp_path):
                try:
                    os.remove(temp_sp_path)
                except Exception:
                    pass
            with open(temp_sp_path, "wb") as f:
                f.write(content.encode('utf-16'))
            logger.info("Successfully wrote fallback shared parameter file to: {}".format(temp_sp_path))
        except Exception as e:
            logger.error("Failed to write fallback shared parameter file contents: {}".format(e))
            
        try:
            app.SharedParametersFilename = temp_sp_path
            sp_file = app.OpenSharedParameterFile()
        except System.Exception as e2:
            logger.error("Failed to load fallback shared parameter file (CLR): {}".format(e2.Message))
            sp_file = None
        except Exception as e2:
            logger.error("Failed to load fallback shared parameter file: {}".format(e2))
            sp_file = None
        
    if not sp_file:
        forms.alert("Could not load/create Shared Parameter File.", title="Shared Parameter Error")
        return False
        
    # 2. Get Group
    group_name = "OxyzenDigital"
    group = sp_file.Groups.get_Item(group_name) or sp_file.Groups.Create(group_name)
    
    # 3. Get or Create Definitions
    def_schema_link = group.Definitions.get_Item("ODI_Schema_Link")
    if not def_schema_link:
        try:
            from Autodesk.Revit.DB import SpecTypeId
            opt = ExternalDefinitionCreationOptions("ODI_Schema_Link", SpecTypeId.String.Text)
        except ImportError:
            from Autodesk.Revit.DB import ParameterType
            opt = ExternalDefinitionCreationOptions("ODI_Schema_Link", ParameterType.Text)
            
        opt.GUID = System.Guid("7f8a9b1c-3d2e-4f5a-6b7c-8d9e0a1b2c3d")
        opt.Description = "Unique hash linking sheet to schema"
        def_schema_link = group.Definitions.Create(opt)
        
    def_building_type = group.Definitions.get_Item("ODI_Building_Type")
    if not def_building_type:
        try:
            from Autodesk.Revit.DB import SpecTypeId
            opt = ExternalDefinitionCreationOptions("ODI_Building_Type", SpecTypeId.String.Text)
        except ImportError:
            from Autodesk.Revit.DB import ParameterType
            opt = ExternalDefinitionCreationOptions("ODI_Building_Type", ParameterType.Text)
            
        opt.GUID = System.Guid("8a7b6c5d-4e3d-2c1b-0a9b-8c7d6e5f4a3b")
        opt.Description = "Assigned project library type"
        def_building_type = group.Definitions.Create(opt)
        
    def_designator_code = group.Definitions.get_Item("ODI_Designator_Code")
    if not def_designator_code:
        try:
            from Autodesk.Revit.DB import SpecTypeId
            opt = ExternalDefinitionCreationOptions("ODI_Designator_Code", SpecTypeId.String.Text)
        except ImportError:
            from Autodesk.Revit.DB import ParameterType
            opt = ExternalDefinitionCreationOptions("ODI_Designator_Code", ParameterType.Text)

        opt.GUID = System.Guid("9b8c7d6e-5f4a-3b2c-1a0b-9c8d7e6f5a4b")
        opt.Description = "Designator sorting and sequencing code"
        def_designator_code = group.Definitions.Create(opt)

    def_disc = group.Definitions.get_Item("SHEETS - Discipline")
    if not def_disc:
        try:
            from Autodesk.Revit.DB import SpecTypeId
            opt = ExternalDefinitionCreationOptions("SHEETS - Discipline", SpecTypeId.String.Text)
        except ImportError:
            from Autodesk.Revit.DB import ParameterType
            opt = ExternalDefinitionCreationOptions("SHEETS - Discipline", ParameterType.Text)
        opt.GUID = System.Guid("ac1b2c3d-4e5f-6a7b-8c9d-0e1f2a3b4c5d")
        opt.Description = "AIA discipline name for Project Browser grouping"
        def_disc = group.Definitions.Create(opt)

    def_use = group.Definitions.get_Item("SHEETS - Use")
    if not def_use:
        try:
            from Autodesk.Revit.DB import SpecTypeId
            opt = ExternalDefinitionCreationOptions("SHEETS - Use", SpecTypeId.String.Text)
        except ImportError:
            from Autodesk.Revit.DB import ParameterType
            opt = ExternalDefinitionCreationOptions("SHEETS - Use", ParameterType.Text)
        opt.GUID = System.Guid("bd2c3d4e-5f6a-7b8c-9d0e-1f2a3b4c5d6e")
        opt.Description = "User-defined sheet use category for Project Browser"
        def_use = group.Definitions.Create(opt)

    if not def_schema_link or not def_building_type or not def_designator_code or not def_disc or not def_use:
        forms.alert("Failed to create parameter definitions.", title="Shared Parameter Error")
        return False
        
    # 4. Bind in a single transaction
    t = Transaction(doc, "Add Manage Sheets Shared Parameters")
    t.Start()
    try:
        if not bound_to_sheets and def_schema_link:
            categories = app.Create.NewCategorySet()
            cat_sheets = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets)
            categories.Insert(cat_sheets)
            binding = app.Create.NewInstanceBinding(categories)
            if has_group_type_id:
                doc.ParameterBindings.Insert(def_schema_link, binding, GroupTypeId.IdentityData)
            else:
                doc.ParameterBindings.Insert(def_schema_link, binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
            logger.info("Bound 'ODI_Schema_Link' to Sheets category.")

        if not bound_to_proj_info and def_building_type:
            categories = app.Create.NewCategorySet()
            cat_proj_info = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ProjectInformation)
            categories.Insert(cat_proj_info)
            binding = app.Create.NewInstanceBinding(categories)
            if has_group_type_id:
                doc.ParameterBindings.Insert(def_building_type, binding, GroupTypeId.IdentityData)
            else:
                doc.ParameterBindings.Insert(def_building_type, binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
            logger.info("Bound 'ODI_Building_Type' to Project Information category.")

        if not bound_designator_code and def_designator_code:
            categories = app.Create.NewCategorySet()
            cat_sheets = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets)
            categories.Insert(cat_sheets)
            binding = app.Create.NewInstanceBinding(categories)
            if has_group_type_id:
                doc.ParameterBindings.Insert(def_designator_code, binding, GroupTypeId.IdentityData)
            else:
                doc.ParameterBindings.Insert(def_designator_code, binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
            logger.info("Bound 'ODI_Designator_Code' to Sheets category.")

        if not bound_disc and def_disc:
            categories = app.Create.NewCategorySet()
            cat_sheets = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets)
            categories.Insert(cat_sheets)
            binding = app.Create.NewInstanceBinding(categories)
            if has_group_type_id:
                doc.ParameterBindings.Insert(def_disc, binding, GroupTypeId.IdentityData)
            else:
                doc.ParameterBindings.Insert(def_disc, binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
            logger.info("Bound 'SHEETS - Discipline' to Sheets category.")

        if not bound_use and def_use:
            categories = app.Create.NewCategorySet()
            cat_sheets = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets)
            categories.Insert(cat_sheets)
            binding = app.Create.NewInstanceBinding(categories)
            if has_group_type_id:
                doc.ParameterBindings.Insert(def_use, binding, GroupTypeId.IdentityData)
            else:
                doc.ParameterBindings.Insert(def_use, binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
            logger.info("Bound 'SHEETS - Use' to Sheets category.")

        t.Commit()
        return True
    except System.Exception as e:
        t.RollBack()
        forms.alert("Error binding parameters (CLR):\n{}".format(e.Message), title="Shared Parameter Error")
        return False
    except Exception as e:
        t.RollBack()
        forms.alert("Error binding parameters:\n{}".format(e), title="Shared Parameter Error")
        return False

def assign_sheet_to_collection(doc, sheet, collection_name):
    """Assigns the sheet to a Sheet Collection by name."""
    if not doc or not collection_name:
        return
        
    try:
        # 1. Try native Revit 2025 SheetCollection via ParameterTypeId
        try:
            from Autodesk.Revit.DB import SheetCollection, ParameterTypeId
            collector = FilteredElementCollector(doc).OfClass(SheetCollection)
            target_collection = None
            for c in collector:
                if get_element_name(c) == collection_name:
                    target_collection = c
                    break
            
            if not target_collection:
                target_collection = SheetCollection.Create(doc, collection_name)
                
            c_param = sheet.get_Parameter(ParameterTypeId.SheetCollection)
            if c_param and not c_param.IsReadOnly:
                c_param.Set(target_collection.Id)
                return
        except:
            pass # Pre-2025
            
        # 2. Strict Project Parameter
        param = sheet.LookupParameter("Sheet Collection")
        if param and not param.IsReadOnly:
            param.Set(collection_name)
            return

    except Exception as e:
        logger.warning("Could not assign sheet to collection '{}': {}".format(collection_name, e))

# --- MAIN EXECUTION ---

def run():
    if not doc:
        print("Error: Revit document context not found.")
        return
        
    # 1. Prompt user to select execution intent JSON
    json_path = forms.pick_file(file_ext="json", title="Select execution_intent.json")
    if not json_path:
        script.exit()
        
    # 2. Parse JSON
    try:
        with open(json_path, "r") as f:
            intent = json.load(f)
    except Exception as e:
        forms.alert("Failed to parse JSON file:\n{}".format(e), title="JSON Error")
        return
        
    schema_sheets = intent.get("Sheets", [])
    if not schema_sheets:
        forms.alert("No sheets defined in schema intent.", title="Invalid Schema")
        return
        
    # 3. Ensure Shared Parameters exist and are bound
    if not ensure_shared_parameter(doc, app):
        return
        
    # 4. Query live sheets & sets
    live_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    
    # Index live sheets
    live_by_id = {}
    live_by_number = {}
    for sheet in live_sheets:
        live_by_id[get_id_value(sheet.Id)] = sheet
        live_by_number[sheet.SheetNumber] = sheet

    # 5. Build MVVM item lists from JSON Action Intents
    create_items = []
    conflict_items = []
    purge_items = []

    # Map schema sheets
    for p_sheet in schema_sheets:
        p_num = p_sheet.get("Number")
        p_name = p_sheet.get("Name")
        p_link = p_sheet.get("SchemaLink")
        p_tb = p_sheet.get("TitleBlock", "")
        p_tb_id = p_sheet.get("TitleBlockId")
        p_set = p_sheet.get("SheetCollection", "")
        p_action = p_sheet.get("Action", "Skip")
        
        # Look up live match
        live_sheet = None
        revit_id_val = p_sheet.get("RevitId")
        if revit_id_val:
            live_sheet = live_by_id.get(int(revit_id_val))
        if not live_sheet:
            live_sheet = live_by_number.get(p_num)
            
        if p_action == "Create" and not live_sheet:
            create_items.append(
                CreateItem(
                    number=p_num,
                    name=p_name,
                    titleblock=p_tb,
                    titleblock_id=p_tb_id,
                    sheet_collection=p_set,
                    schema_link=p_link,
                    designator_code=p_sheet.get("DesignatorCode", ""),
                    sheet_discipline=p_sheet.get("SheetDiscipline", ""),
                    sheet_use=p_sheet.get("SheetUse", "")
                )
            )
        elif live_sheet:
            view_count = len(p_sheet.get("RenameViews", {}))
            reasons = []
            if live_sheet.SheetNumber != p_num:
                reasons.append("Renumber")
            if get_element_name(live_sheet) != p_name:
                reasons.append("Rename")
            if view_count > 0:
                reasons.append("Update {} Views".format(view_count))
                
            reason_str = " & ".join(reasons) if reasons else "Sync Schema Metadata"
            
            item = ConflictItem(
                element=live_sheet,
                number=p_num,
                model_name=get_element_name(live_sheet),
                proposed_name=p_name,
                reason=reason_str,
                proposed_sheet=p_sheet,
                schema_link=p_link
            )
            item.SelectedAction = "Overwrite" if p_action == "Overwrite" else "Skip"
            conflict_items.append(item)

    # Process purges from JSON intent
    purged_ids = set()
    purges_list = intent.get("Purges", [])
    for p_purge in purges_list:
        p_num = p_purge.get("Number")
        live_sheet = None
        revit_id_val = p_purge.get("RevitId")
        if revit_id_val:
            live_sheet = live_by_id.get(int(revit_id_val))
        if not live_sheet:
            live_sheet = live_by_number.get(p_num)
            
        if live_sheet:
            purge_items.append(
                PurgeItem(
                    element=live_sheet,
                    number=p_num,
                    name=get_element_name(live_sheet)
                )
            )
            purged_ids.add(get_id_value(live_sheet.Id))

    # Also display other unassigned sheets that are not in schema and not in sheet sets as un-checked purge options
    matched_ids = set()
    for item in conflict_items:
        matched_ids.add(get_id_value(item.element.Id))
        
    def _get_sheet_collection(sht):
        # Native Revit 2025+
        try:
            from Autodesk.Revit.DB import ParameterTypeId
            c_param = sht.get_Parameter(ParameterTypeId.SheetCollection)
            if c_param and c_param.HasValue:
                c_id = c_param.AsElementId()
                if c_id and c_id.IntegerValue > -1:
                    c_elem = doc.GetElement(c_id)
                    if c_elem: return get_element_name(c_elem)
        except: pass
        
        # Strict parameter
        p = sht.LookupParameter("Sheet Collection")
        if p and p.HasValue:
            v = p.AsString()
            if not v: v = p.AsValueString()
            if v: return v
        return ""

    for sheet in live_sheets:
        sid = get_id_value(sheet.Id)
        if sid in matched_ids or sid in purged_ids:
            continue
        if any(s.get("Number") == sheet.SheetNumber for s in schema_sheets):
            continue
            
        collection_val = _get_sheet_collection(sheet)
        if not collection_val:
            item = PurgeItem(
                element=sheet,
                number=sheet.SheetNumber,
                name=get_element_name(sheet)
            )
            item.ShouldPurge = False
            purge_items.append(item)

    # 6. Load WPF Reviewer Window
    current_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(current_dir, "ReviewerWindow.xaml")
    
    if not os.path.exists(xaml_path):
        forms.alert("Could not locate ReviewerWindow.xaml in pushbutton directory.", title="Missing UI")
        return
        
    window = ReviewerWindow(xaml_path, create_items, conflict_items, purge_items)
    window.ShowDialog()
    
    if window.cancelled:
        logger.info("Schema execution cancelled by user.")
        return

    # 7. Execute transaction
    logger.info("Executing Schema Sync...")
    
    loaded_tbs = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_TitleBlocks) \
        .WhereElementIsElementType() \
        .ToElements()

    views_renamed_count = 0
    failed_renames_count = 0
    t = Transaction(doc, "Sync Proposed Sheet Schema")
    t.Start()
    try:
        # Create sheets
        for item in create_items:
            tb_symbol = None
            if item.titleblock_id and item.titleblock_id != "fallback":
                try:
                    tb_el = doc.GetElement(ElementId(int(item.titleblock_id)))
                    if isinstance(tb_el, FamilySymbol):
                        tb_symbol = tb_el
                except Exception:
                    pass

            if not tb_symbol and item.titleblock:
                if ":" in item.titleblock:
                    fam_name, type_name = item.titleblock.split(":", 1)
                    for symbol in loaded_tbs:
                        if symbol.Family and get_element_name(symbol.Family) == fam_name and get_element_name(symbol) == type_name:
                            tb_symbol = symbol
                            break

            if not tb_symbol:
                if loaded_tbs:
                    options = { "{}:{}".format(get_element_name(tb.Family) if tb.Family else "Default", get_element_name(tb)): tb for tb in loaded_tbs }
                    selected_key = forms.SelectFromList.show(
                        options.keys(),
                        title="Sheet {} - Fallback Title Block".format(item.number),
                        multiselect=False
                    )
                    if selected_key:
                        tb_symbol = options[selected_key]
                    else:
                        tb_symbol = loaded_tbs[0]
                else:
                    forms.alert("No Title Blocks loaded. Creating sheet without Title Block.", title="Warning")
                    
            tb_id = tb_symbol.Id if tb_symbol else ElementId.InvalidElementId
            
            # Create ViewSheet
            new_sheet = ViewSheet.Create(doc, tb_id)
            new_sheet.SheetNumber = item.number
            new_sheet.Name = item.name
            
            # Write unique Schema Link
            param = new_sheet.LookupParameter("ODI_Schema_Link")
            if param and not param.IsReadOnly:
                param.Set(item.schema_link)
                
            # Write Designator Code
            param_code = new_sheet.LookupParameter("ODI_Designator_Code")
            if param_code and not param_code.IsReadOnly:
                param_code.Set(str(item.designator_code))

            # Write SHEETS - Discipline (AIA Project Browser grouping)
            param_disc = new_sheet.LookupParameter("SHEETS - Discipline")
            if param_disc and not param_disc.IsReadOnly:
                disc_val = item.sheet_discipline or get_aia_discipline(item.number.split("-")[0].rstrip("0123456789"))
                param_disc.Set(disc_val)

            # Write SHEETS - Use (user-defined, may be blank)
            param_use = new_sheet.LookupParameter("SHEETS - Use")
            if param_use and not param_use.IsReadOnly and item.sheet_use:
                param_use.Set(item.sheet_use)

            # Add to Sheet Set
            if item.sheet_collection:
                assign_sheet_to_collection(doc, new_sheet, item.sheet_collection)

        # Process Conflicts
        for item in conflict_items:
            if item.SelectedAction == "Overwrite":
                # Update Number and Name
                item.element.SheetNumber = item.proposed_sheet["Number"]
                item.element.Name = item.proposed_sheet["Name"]
                
                # Write Schema Link
                param = item.element.LookupParameter("ODI_Schema_Link")
                if param and not param.IsReadOnly:
                    param.Set(item.schema_link)
                    
                # Write Designator Code
                param_code = item.element.LookupParameter("ODI_Designator_Code")
                if param_code and not param_code.IsReadOnly:
                    param_code.Set(str(item.proposed_sheet.get("DesignatorCode", "")))

                # Write SHEETS - Discipline
                param_disc = item.element.LookupParameter("SHEETS - Discipline")
                if param_disc and not param_disc.IsReadOnly:
                    disc_val = item.proposed_sheet.get("SheetDiscipline", "")
                    if not disc_val:
                        # Fall back to deriving from the sheet number prefix
                        p_num_local = item.proposed_sheet.get("Number", "")
                        disc_val = get_aia_discipline(p_num_local.split("-")[0].rstrip("0123456789"))
                    param_disc.Set(disc_val)

                # Write SHEETS - Use
                param_use = item.element.LookupParameter("SHEETS - Use")
                if param_use and not param_use.IsReadOnly:
                    use_val = item.proposed_sheet.get("SheetUse", "")
                    if use_val:
                        param_use.Set(use_val)

                # Add to set
                set_name = item.proposed_sheet.get("SheetCollection")
                if set_name:
                    assign_sheet_to_collection(doc, item.element, set_name)
                    
                # Rename views placed on this sheet
                rename_views_map = item.proposed_sheet.get("RenameViews", {})
                for v_id_str, proposed_v_name in rename_views_map.items():
                    try:
                        v_id = ElementId(int(v_id_str))
                        view_elem = doc.GetElement(v_id)
                        if view_elem and proposed_v_name:
                            old_v_name = get_element_name(view_elem)
                            if old_v_name != proposed_v_name:
                                view_elem.Name = proposed_v_name
                                logger.info("Renamed view '{}' to '{}'".format(old_v_name, proposed_v_name))
                                views_renamed_count += 1
                    except System.Exception as ve:
                        logger.warning("Could not rename view ID {} (CLR): {}".format(v_id_str, ve.Message))
                        failed_renames_count += 1
                    except Exception as ve:
                        logger.warning("Could not rename view ID {}: {}".format(v_id_str, ve))
                        failed_renames_count += 1

        # Process Collection Renames
        collection_renames = intent.get("CollectionRenames", {})
        renamed_collection_sheets_count = 0
        if collection_renames:
            for sheet in live_sheets:
                old_collection = _get_sheet_collection(sheet)
                if old_collection and old_collection in collection_renames:
                    new_collection = collection_renames[old_collection]
                    assign_sheet_to_collection(doc, sheet, new_collection)
                    logger.info("Renamed sheet collection for '{}' from '{}' to '{}'".format(sheet.SheetNumber, old_collection, new_collection))
                    renamed_collection_sheets_count += 1

        # Process Purges
        purged_count = 0
        for item in purge_items:
            if item.ShouldPurge:
                doc.Delete(item.element.Id)
                purged_count += 1

        # Write Library type tag back to Project Information
        lib_type = intent.get("LibraryType", "")
        if lib_type:
            try:
                proj_info = doc.ProjectInformation
                if proj_info:
                    param = proj_info.LookupParameter("ODI_Building_Type")
                    if param and not param.IsReadOnly:
                        param.Set(lib_type)
                        logger.info("Updated Project Information Building Type tag to: {}".format(lib_type))
            except Exception as e:
                logger.warning("Could not write Building Type tag: {}".format(e))

        t.Commit()
        
        # Display Sync Completed popup
        msg = "Sync Completed!\n- Created: {} sheets\n- Updated: {} sheets\n- Renamed: {} views\n- Purged: {} sheets\n- Re-assigned Collections: {} sheets".format(
            len(create_items),
            len([x for x in conflict_items if x.SelectedAction == "Overwrite"]),
            views_renamed_count,
            purged_count,
            renamed_collection_sheets_count
        )
        if failed_renames_count > 0:
            msg += "\n\n⚠️ Note: {} views could not be renamed due to duplicate names or naming rules in Revit (see logs).".format(failed_renames_count)
            
        forms.alert(msg, title="Success")
        
        # Log to %APPDATA%
        try:
            import datetime
            appdata = os.getenv("APPDATA")
            log_dir = os.path.join(appdata, "OxyzenDigital", "ManageSheets")
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            log_path = os.path.join(log_dir, "sync_log.json")
            
            log_data = []
            if os.path.exists(log_path):
                try:
                    with open(log_path, "r") as lf:
                        log_data = json.load(lf)
                        if not isinstance(log_data, list):
                            log_data = []
                except:
                    pass
            
            proj_name = ""
            try:
                pi = doc.ProjectInformation
                p_name_param = pi.LookupParameter("Project Name")
                if p_name_param:
                    proj_name = p_name_param.AsString() or ""
            except:
                pass
                
            entry = {
                "timestamp": datetime.datetime.now().isoformat(),
                "project_name": proj_name,
                "library_type": lib_type,
                "created_sheets": len(create_items),
                "modified_sheets": len([x for x in conflict_items if x.SelectedAction == "Overwrite"]),
                "views_renamed": views_renamed_count,
                "purged_sheets": purged_count,
                "failed_renames": failed_renames_count
            }
            log_data.append(entry)
            log_data = log_data[-50:] # keep only last 50 entries
            
            with open(log_path, "w") as lf:
                json.dump(log_data, lf, indent=2)
            logger.info("Successfully logged sync summary to APPDATA.")
        except Exception as le:
            logger.warning("Could not write sync log: {}".format(le))
            
    except System.Exception as e:
        t.RollBack()
        forms.alert("Sync transaction failed (CLR):\n{}".format(e.Message), title="Transaction Error")
        logger.error("Sync transaction failed (CLR): {}".format(e.Message))
    except Exception as e:
        t.RollBack()
        forms.alert("Sync transaction failed:\n{}".format(e), title="Transaction Error")
        logger.error("Sync transaction failed: {}".format(e))

if __name__ == "__main__":
    run()
