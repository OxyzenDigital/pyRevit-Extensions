from pyrevit import revit, DB, script
from pyrevit.revit import Transaction

def ensure_sheet_view(view):
    """Ensure the active view is a Sheet View."""
    if not isinstance(view, DB.ViewSheet):
        script.exit()

def get_background_sheet_view():
    """Get the active sheet view of the first background document."""
    background_docs = [doc for doc in revit.docs if doc != revit.doc]
    if not background_docs:
        script.exit()

    background_doc = background_docs[0]
    background_active_view = background_doc.ActiveView
    if not isinstance(background_active_view, DB.ViewSheet):
        script.exit()
    return background_active_view, background_doc

def get_guide_grid(view, doc):
    """Find the Guide Grid associated with a specific view."""
    guide_grids = DB.FilteredElementCollector(doc).OfClass(DB.GuideGrid).ToElements()
    for guide_grid in guide_grids:
        if guide_grid.ViewId == view.Id:
            return guide_grid
    return None

def create_or_get_guide_grid(view, doc, source_guide_grid):
    """Find or create a Guide Grid for the active view."""
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
        view_outline = view.CropBox
        if view_outline:
            min_point = view_outline.Min
            max_point = view_outline.Max
            min_width = abs(max_point.X - min_point.X)
            min_height = abs(max_point.Y - min_point.Y)
            expanded_min_point = DB.XYZ(min_point.X - min_width / 2, min_point.Y - min_height / 2, min_point.Z)
            expanded_max_point = DB.XYZ(max_point.X + min_width / 2, max_point.Y + min_height / 2, max_point.Z)
            outline = DB.Outline(expanded_min_point, expanded_max_point)
            new_guide_grid.SetOutline(outline)
        t.Commit()
        return new_guide_grid

def transfer_guide_grid_properties(source, target):
    """Transfer properties from the source Guide Grid to the target Guide Grid."""
    with Transaction(doc, "Transfer Guide Grid Properties") as t:
        t.Start()
        target.GridSpacing = source.GridSpacing
        source_outline = source.GetOutline()
        target_outline = target.GetOutline()
        if not source_outline.IsAlmostEqualTo(target_outline):
            target.SetOutline(source_outline)
        t.Commit()

# Main execution
active_view = revit.doc.ActiveView
ensure_sheet_view(active_view)

background_active_view, background_doc = get_background_sheet_view()
source_guide_grid = get_guide_grid(background_active_view, background_doc)
if not source_guide_grid:
    script.exit()

target_guide_grid = create_or_get_guide_grid(active_view, revit.doc, source_guide_grid)
transfer_guide_grid_properties(source_guide_grid, target_guide_grid)

script.exit()
