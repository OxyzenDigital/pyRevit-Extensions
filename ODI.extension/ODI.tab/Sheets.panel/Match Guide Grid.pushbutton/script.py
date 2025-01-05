from pyrevit import revit, DB, script, forms, output

logger = script.get_logger()
out = script.get_output()
logger.set_level(20)  # INFO or DEBUG
out.set_height(600)


def log_debug(msg):
    """Helper to print debug info to the pyRevit output panel (IronPython 2.7 style)."""
    logger.debug(msg)
    out.print_md("**Debug:** {}".format(msg))


def get_doc_title(doc):
    """Return a doc title or fallback if not available."""
    try:
        return doc.Title
    except:
        return "Untitled Document"


def get_sheets_from_doc(doc):
    """
    Collect all ViewSheet objects in 'doc'.
    Returns a dict: { "SheetNumber - SheetName": ViewSheetObject }
    """
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
    sheets_dict = {}
    for sheet in collector:
        key = "{} - {}".format(sheet.SheetNumber, sheet.Name)
        sheets_dict[key] = sheet
    return sheets_dict


# ------------------------------------------------------------------------------
# DEBUG FUNCTIONS
# ------------------------------------------------------------------------------

def list_guide_grid_parameters(guide_grid):
    """
    Lists all parameters (Name + Value) of the given GuideGrid element.
    """
    if not guide_grid:
        log_debug("No GuideGrid element to list parameters from.")
        return

    log_debug("Listing parameters for GuideGrid: '{}' (ID: {})".format(
        guide_grid.Name, guide_grid.Id
    ))
    for param in guide_grid.Parameters:
        defn = param.Definition
        if not defn:
            continue

        p_name = defn.Name
        stype = param.StorageType
        val_str = "<unknown>"

        if stype == DB.StorageType.String:
            val_str = param.AsString()
        elif stype == DB.StorageType.Double:
            val_str = str(param.AsDouble())  # raw internal double
        elif stype == DB.StorageType.Integer:
            val_str = str(param.AsInteger())
        elif stype == DB.StorageType.ElementId:
            eid = param.AsElementId()
            if eid and eid != DB.ElementId.InvalidElementId:
                val_str = "ElementId: {}".format(eid.IntegerValue)
            else:
                val_str = "InvalidElementId"
        else:
            val_str = "<none>"

        log_debug("     - {} = {}".format(p_name, val_str))


# ------------------------------------------------------------------------------
# GUIDE GRID HELPER FUNCTIONS
# ------------------------------------------------------------------------------

def get_guide_grid_from_sheet(sheet):
    """
    Retrieves the GuideGrid element assigned to 'sheet'.
    Handles cases where DB.GuideGrid might not be available.
    Returns the GuideGrid element or None if not found/assigned.
    """
    if not sheet or not isinstance(sheet, DB.ViewSheet):
        log_debug("Invalid or missing sheet.")
        return None

    doc = sheet.Document
    param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_GUIDE_GRID)
    if not param:
        log_debug("Sheet '{}' has no SHEET_GUIDE_GRID parameter.".format(sheet.Name))
        return None

    storage_type = param.StorageType
    if storage_type == DB.StorageType.ElementId:
        # Param is an ElementId referencing the GuideGrid
        gg_id = param.AsElementId()
        if gg_id and gg_id != DB.ElementId.InvalidElementId:
            guide_grid = doc.GetElement(gg_id)
            
            # Check if the element is a Guide Grid (without relying on DB.GuideGrid)
            if guide_grid and guide_grid.Category.Name == "Guide Grid":  
                log_debug("Retrieved GuideGrid by ElementId: '{}' (ID: {})".format(
                    guide_grid.Name, guide_grid.Id
                ))
                return guide_grid
            else:
                log_debug("ElementId from SHEET_GUIDE_GRID is not a valid GuideGrid.")
        else:
            log_debug("SHEET_GUIDE_GRID param is an invalid ElementId.")
        return None
    elif storage_type == DB.StorageType.String:
        # Param is the string name of the GuideGrid
        grid_name = param.AsString()
        if not grid_name:
            log_debug("Sheet '{}' has an empty guide grid name.".format(sheet.Name))
            return None

        # Find the GuideGrid by name
        grids_collector = DB.FilteredElementCollector(doc).OfClass(DB.GuideGrid)
        for gg in grids_collector:
            if gg.Name == grid_name:
                log_debug("Retrieved GuideGrid by Name: '{}' (ID: {})".format(
                    gg.Name, gg.Id
                ))
                return gg

        log_debug("No GuideGrid found by name '{}' in doc.".format(grid_name))
        return None

    else:
        # Unknown or unexpected storage type
        log_debug("SHEET_GUIDE_GRID has unexpected storage type: {}".format(storage_type))
        return None


def create_guide_grid(doc, sheet, guide_grid_name):
    """
    Create a new GuideGrid on the given sheet with the specified name.
    Handles different Revit API versions for Guide Grid creation.
    """
    try:
        with revit.Transaction("Create Guide Grid"):
            # Use the older API method directly (without trying NewGuideGrid() first)
            new_grid = DB.GuideGrid.Create(doc, sheet.Id, guide_grid_name) 

            log_debug("Created new GuideGrid '{}' (ID: {}) on sheet '{}'.".format(
                guide_grid_name, new_grid.Id, sheet.Name
            ))
            return new_grid
    except Exception as e:
        log_debug("Failed to create guide grid on sheet '{}': {}".format(sheet.Name, e))
        return None


def copy_guide_grid_properties(source_grid, target_grid):
    """
    Copy basic properties (like name) from source_grid to target_grid.
    Handles cases where DB.BuiltInParameter might not be fully accessible.
    """
    if not source_grid or not target_grid:
        log_debug("Invalid source or target grid. Cannot copy properties.")
        return False

    try:
        with revit.Transaction("Copy Guide Grid Properties"):
            # Copy the grid name using LookupParameter()
            src_name_param = source_grid.LookupParameter("Name")  
            tgt_name_param = target_grid.LookupParameter("Name")

            if src_name_param and tgt_name_param:
                old_name = tgt_name_param.AsString()
                new_name = src_name_param.AsString()
                tgt_name_param.Set(new_name)
                log_debug("GuideGrid name changed from '{}' to '{}'.".format(old_name, new_name))

            return True
    except Exception as e:
        log_debug("Failed to copy grid properties: {}".format(e))
        return False


def match_guide_grid_from_source_sheet(source_sheet, target_sheet):
    """
    1) Debug-list the source sheet's guide grid parameters.
    2) Retrieve the source grid.
    3) Debug-list the target sheet's guide grid parameters (before creation).
    4) If none, create a new one with the same name.
    5) Copy properties from source to target.
    6) Debug-list the target sheet's guide grid parameters (after copying).
    """
    log_debug("------ DEBUG: Listing SOURCE sheet guide grid params ------")
    source_grid = get_guide_grid_from_sheet(source_sheet)
    if source_grid:
        list_guide_grid_parameters(source_grid)
    else:
        forms.alert(
            "Source sheet '{}' does not have a valid GuideGrid.".format(source_sheet.Name),
            exitscript=True
        )
        return

    # Debug target sheet guide grid before we do anything
    log_debug("------ DEBUG: Listing TARGET sheet guide grid params (BEFORE) ------")
    target_grid_before = get_guide_grid_from_sheet(target_sheet)
    if target_grid_before:
        list_guide_grid_parameters(target_grid_before)
    else:
        log_debug("No guide grid on target sheet '{}' yet.".format(target_sheet.Name))

    # If there's no target grid, create one with the same name as the source
    target_grid = target_grid_before
    if not target_grid:
        # Get the Guide Grid name using LookupParameter()
        src_name_param = source_grid.LookupParameter("Name")  
        src_name = src_name_param.AsString() if src_name_param else "Default Guide Grid Name"
        target_grid = create_guide_grid(target_sheet.Document, target_sheet, src_name)
        if not target_grid:
            forms.alert(
                "Failed to create a new guide grid on target sheet '{}'. "
                "Check logs for details.".format(target_sheet.Name),
                exitscript=True
            )
            return

    # Copy properties
    success = copy_guide_grid_properties(source_grid, target_grid)
    if not success:
        forms.alert("Failed to copy guide grid properties.", exitscript=True)
        return

    # Debug target sheet guide grid after copying
    log_debug("------ DEBUG: Listing TARGET sheet guide grid params (AFTER) ------")
    target_grid_after = get_guide_grid_from_sheet(target_sheet)
    if target_grid_after:
        list_guide_grid_parameters(target_grid_after)
        forms.alert(
            "Successfully matched guide grid from source sheet '{}' to target sheet '{}'. \
Check pyRevit output for logs.".format(source_sheet.Name, target_sheet.Name),
            exitscript=False
        )
    else:
        forms.alert(
            "Something went wrong. No valid GuideGrid on the target sheet after copying.",
            exitscript=True
        )


# ------------------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------------------

def main():
    # 1) Verify the active view is a sheet in the current doc (the "target" doc)
    if not isinstance(revit.active_view, DB.ViewSheet):
        forms.alert("Please open a sheet view (target sheet) before running this script.",
                    exitscript=True)

    target_sheet = revit.active_view
    log_debug("Target sheet: '{} - {}' (ID: {})".format(
        target_sheet.SheetNumber, target_sheet.Name, target_sheet.Id
    ))

    # 2) Prompt the user to pick a "source" document (open in same session)
    other_docs = [d for d in revit.docs if d != revit.doc]
    if not other_docs:
        forms.alert("No other open documents found. Please open the source document.", exitscript=True)

    doc_options = {}
    for d in other_docs:
        doc_title = get_doc_title(d)
        doc_options[doc_title] = d

    doc_titles_sorted = sorted(doc_options.keys())
    chosen_doc_title = forms.SelectFromList.show(
        doc_titles_sorted,
        title="Select Source Document",
        multiselect=False,
        button_name="Select Document",
        width=400,
        height=300
    )
    if not chosen_doc_title:
        script.exit()

    source_doc = doc_options[chosen_doc_title]
    log_debug("Selected source doc: '{}'".format(source_doc.Title))

    # 3) Prompt the user to pick the source sheet
    sheets_dict = get_sheets_from_doc(source_doc)
    if not sheets_dict:
        forms.alert("No sheets found in the chosen source document.", exitscript=True)

    sheets_sorted = sorted(sheets_dict.keys())
    chosen_sheet_key = forms.SelectFromList.show(
        sheets_sorted,
        title="Select Source Sheet",
        multiselect=False,
        button_name="Select Sheet",
        width=500,
        height=400
    )
    if not chosen_sheet_key:
        script.exit()

    source_sheet = sheets_dict[chosen_sheet_key]
    log_debug("Selected source sheet: '{} - {}' (ID: {})".format(
        source_sheet.SheetNumber, source_sheet.Name, source_sheet.Id
    ))

    # 4) Match the guide grid from the source sheet to the target (active) sheet
    match_guide_grid_from_source_sheet(source_sheet, target_sheet)


if __name__ == "__main__":
    main()