from pyrevit import revit, DB, script, forms, output

logger = script.get_logger()
out = script.get_output()


def log_debug(msg):
    """Helper to print debug info to the pyRevit output panel."""
    logger.debug(msg)
    out.print_md("**Debug:** {}".format(msg))


def get_doc_title(doc):
    """Return a doc title or fallback if not available."""
    try:
        return doc.Title
    except:
        return "Untitled Document"


def get_sheets_from_doc(doc):
    """Collect all ViewSheet objects in 'doc'."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
    return { "{} - {}".format(sheet.SheetNumber, sheet.Name): sheet for sheet in collector }


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
            log_debug("SHEET_GUIDE_GRID param is an invalid ElementId.")  # This is where the log message occurs
        return None

    elif storage_type == DB.StorageType.String:
        # Param is the string name of the GuideGrid
        grid_name = param.AsString()
        if not grid_name:
            log_debug("Sheet '{}' has an empty guide grid name.".format(sheet.Name))
            return None

        # Find the GuideGrid by name (iterating through all Guide Grids)
        grids_collector = DB.FilteredElementCollector(doc).OfClass(DB.GuideGrid)
        for gg in grids_collector:
            if gg.get_Parameter(DB.BuiltInParameter.GUIDE_GRID_NAME).AsString() == grid_name:
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
    Attempts various methods for Guide Grid creation to handle API differences.
    """
    try:
        with revit.Transaction("Create Guide Grid"):
            new_grid = None

            try:
                # Attempt to use doc.Create.NewGuideGrid() (Revit 2019+)
                new_grid = doc.Create.NewGuideGrid(sheet.Id, guide_grid_name)
            except AttributeError:
                pass  # Ignore if NewGuideGrid() is not available

            if not new_grid:
                try:
                    # Try alternative NewGuideGrid() overload with origin and basis vectors
                    origin = DB.XYZ(0, 0, 0)
                    basisX = DB.XYZ(1, 0, 0)
                    basisY = DB.XYZ(0, 1, 0)
                    new_grid = doc.Create.NewGuideGrid(sheet.Id, guide_grid_name, origin, basisX, basisY)
                except AttributeError:
                    pass  # Ignore if this overload is also not available

            if not new_grid:
                try:
                    # Fallback to older API method if the above fail
                    new_grid = DB.GuideGrid.Create(doc, sheet.Id, guide_grid_name)
                except AttributeError:
                    pass  # Ignore if DB.GuideGrid.Create() is not available

            if new_grid:
                log_debug("Created new GuideGrid '{}' (ID: {}) on sheet '{}'.".format(
                    guide_grid_name, new_grid.Id, sheet.Name
                ))
                return new_grid
            else:
                log_debug("Failed to create guide grid on sheet '{}'. No suitable method found.".format(sheet.Name))
                return None

    except Exception as e:
        log_debug("Failed to create guide grid on sheet '{}': {}".format(sheet.Name, e))
        return None

def copy_guide_grid_properties(source_grid, target_grid):
    """Copy properties from source_grid to target_grid."""
    try:
        with revit.Transaction("Copy Guide Grid Properties"):
            # Example property copy: Guide Grid Name
            if source_grid and target_grid:
                target_grid.Name = source_grid.Name
                log_debug("GuideGrid name copied from '{}' to '{}'.".format(source_grid.Name, target_grid.Name))
            else:
                log_debug("Invalid source or target grid. Cannot copy properties.")
            return True
    except Exception as e:
        log_debug("Failed to copy grid properties: {}".format(e))
        return False
        
def main():
    # Ensure the active view is a sheet
    if not isinstance(revit.active_view, DB.ViewSheet):
        forms.alert("Please open a sheet view before running this script.", exitscript=True)

    target_sheet = revit.active_view
    target_doc = revit.doc
    log_debug("Target sheet: '{}' (ID: {})".format(target_sheet.Name, target_sheet.Id))

    # Prompt user to select a source document
    other_docs = [d for d in revit.docs if d.Title != target_doc.Title]
    if not other_docs:
        forms.alert("No other open documents found. Please open the source document.", exitscript=True)

    doc_options = {get_doc_title(d): d for d in other_docs}
    chosen_doc_title = forms.SelectFromList.show(sorted(doc_options.keys()),
                                                title="Select Source Document",
                                                button_name="Select Document")
    if not chosen_doc_title:
        script.exit()

    source_doc = doc_options[chosen_doc_title]
    log_debug("Selected source doc: '{}'".format(source_doc.Title))

    # Prompt user to select the source sheet
    sheets_dict = get_sheets_from_doc(source_doc)
    if not sheets_dict:
        forms.alert("No sheets found in the chosen source document.", exitscript=True)

    chosen_sheet_key = forms.SelectFromList.show(sorted(sheets_dict.keys()),
                                                title="Select Source Sheet",
                                                button_name="Select Sheet")
    if not chosen_sheet_key:
        script.exit()

    source_sheet = sheets_dict[chosen_sheet_key]
    log_debug("Selected source sheet: '{}' (ID: {})".format(source_sheet.Name, source_sheet.Id))

    # Get the Guide Grid from the source sheet
    source_guide_grid = get_guide_grid_from_sheet(source_sheet)
    if not source_guide_grid:
        forms.alert("No Guide Grid found on the source sheet.", exitscript=True)

    # Check if there's an existing Guide Grid on the target sheet
    target_guide_grid = get_guide_grid_from_sheet(target_sheet)
    
    # Handle Guide Grid creation or update
    with revit.TransactionGroup("Process Guide Grid", target_doc):
        if not target_guide_grid:
            # Create a new Guide Grid with the same name as the source and assign it to the target sheet
            target_guide_grid = create_guide_grid(target_doc, target_sheet, source_guide_grid.Name)
            if not target_guide_grid:
                forms.alert("Failed to create a new Guide Grid.", exitscript=True)

        # Copy properties from source to target Guide Grid
        if not copy_guide_grid_properties(source_guide_grid, target_guide_grid):
            forms.alert("Failed to copy properties to the target Guide Grid.", exitscript=True)

    forms.alert("Guide Grid processing completed. Check the debug output for details.", exitscript=False)

if __name__ == "__main__":
    main()