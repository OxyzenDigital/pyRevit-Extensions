from pyrevit import revit, DB
from pyrevit.revit import Transaction

# Ensure the active view is a Sheet View
doc = revit.doc
active_view = doc.ActiveView
if not isinstance(active_view, DB.ViewSheet):
    print("Error: Active view must be a Sheet View.")
    sys.exit()

# Get the currently opened documents
doc_manager = revit.doc_manager
background_docs = [d for d in doc_manager.docs if d != doc]
if not background_docs:
    print("Error: No background documents found.")
    sys.exit()

# Get the active view of the first background document
background_doc = background_docs[0]
background_active_view = background_doc.ActiveView
if not isinstance(background_active_view, DB.ViewSheet):
    print("Error: Background active view must be a Sheet View.")
    sys.exit()

# Find the Guide Grid in the background active view
def get_guide_grid(view, doc):
    guide_grids = DB.FilteredElementCollector(doc).OfClass(DB.GuideGrid).ToElements()
    for guide_grid in guide_grids:
        if guide_grid.ViewId == view.Id:
            return guide_grid
    return None

source_guide_grid = get_guide_grid(background_active_view, background_doc)
if not source_guide_grid:
    print("Error: No Guide Grid found in the background document's active view.")
    sys.exit()

# Find or create a Guide Grid in the active view
def create_or_get_guide_grid(view, doc):
    guide_grids = DB.FilteredElementCollector(doc).OfClass(DB.GuideGrid).ToElements()
    for guide_grid in guide_grids:
        if guide_grid.Name == source_guide_grid.Name:
            if guide_grid.ViewId != view.Id:
                with Transaction(doc, "Assign Guide Grid to Sheet") as t:
                    t.Start()
                    guide_grid.ViewId = view.Id
                    t.Commit()
            return guide_grid
    with Transaction(doc, "Create Guide Grid") as t:
        t.Start()
        new_guide_grid = DB.GuideGrid.Create(doc, view.Id)
        new_guide_grid.Name = source_guide_grid.Name
        # Set default outline to fit the sheet if no outline exists
        view_outline = view.CropBox
        if view_outline:
            min_point = view_outline.Min
            max_point = view_outline.Max
            outline = DB.Outline(min_point, max_point)
            new_guide_grid.SetOutline(outline)
        t.Commit()
        return new_guide_grid

# Transfer properties to the Guide Grid
def transfer_guide_grid_properties(source, target):
    with Transaction(doc, "Transfer Guide Grid Properties") as t:
        t.Start()
        target.GridSpacing = source.GridSpacing
        outline = source.GetOutline()
        target.SetOutline(outline)
        t.Commit()

# Perform the transfer
target_guide_grid = create_or_get_guide_grid(active_view, doc)
transfer_guide_grid_properties(source_guide_grid, target_guide_grid)

print("Guide Grid successfully transferred.")
