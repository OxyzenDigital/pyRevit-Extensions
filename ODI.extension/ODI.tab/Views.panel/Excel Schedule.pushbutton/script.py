import json
from pyrevit import revit, DB

# Load JSON data
with open('c:\\tmp\\save_data.json', 'r') as json_file:
    excel_data = json.load(json_file)

def setup_legend_view():
    doc = revit.doc
    view = next((v for v in DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements() if v.ViewType == DB.ViewType.Legend), None)
    if view:
        print "Found legend view: " + view.Name
        with DB.Transaction(doc, "Duplicate Legend View") as t:
            t.Start()
            new_view_id = view.Duplicate(DB.ViewDuplicateOption.Duplicate)
            new_view = doc.GetElement(new_view_id)
            if new_view:
                print "New view created: " + new_view.Name
                new_view.Scale = 48  # 1/4" = 1'-0"
                t.Commit()
                return new_view
            else:
                print "Failed to create new view"
                t.RollBack()
                return None
    else:
        print "No Legend View found to duplicate."
        return None

def create_legend_elements(view, data):
    if view is None:
        print "Cannot proceed: View not created or found."
        return

    doc = revit.doc
    with DB.Transaction(doc, "Create Legend Elements") as t:
        t.Start()
        
        for cell in data['cells']:
            # Get row height and column width
            row_height = data['row_heights'].get(str(cell['row']), 12.75)  # Default to smallest height if not found
            col_width = data['column_widths'].get(str(cell['col']), 13.0)  # Default to smallest width if not found

            # Convert from Excel points to Revit units (feet)
            height = row_height * 0.013888888888889  # 1 point = 1/72 inch, 1 inch = 1/12 ft
            width = col_width * 0.013888888888889

            # Positioning - Assuming top-left corner of view is origin (0,0)
            x = (cell['col'] - 1) * width
            y = -(cell['row'] - 1) * height  # Negative because Revit's Y-axis points downward

            # Create Filled Region for cell background
            bg_fill = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Model, "Solid Fill")
            if bg_fill:
                # Create lines to form a rectangle
                lines = [
                    DB.Line.CreateBound(DB.XYZ(x, y, 0), DB.XYZ(x + width, y, 0)),  # Bottom line
                    DB.Line.CreateBound(DB.XYZ(x + width, y, 0), DB.XYZ(x + width, y - height, 0)),  # Right line
                    DB.Line.CreateBound(DB.XYZ(x + width, y - height, 0), DB.XYZ(x, y - height, 0)),  # Top line
                    DB.Line.CreateBound(DB.XYZ(x, y - height, 0), DB.XYZ(x, y, 0))  # Left line
                ]
                
                filled_region = DB.FilledRegion.Create(doc, view.Id, bg_fill.Id, lines)
            else:
                print "Failed to find solid fill pattern for cell background."

            # Add text note for cell content
            text = DB.TextNote.Create(doc, view.Id, DB.XYZ(x + width/2, y - height/2, 0), cell['value'])
            text_type = text.TextNoteType
            text_type.FontSize = cell['font']['size']  # Revit will handle conversion from points to feet

        t.Commit()

# Main execution
legend_view = setup_legend_view()
if legend_view:
    create_legend_elements(legend_view, excel_data)
else:
    print "Failed to create or find Legend View. Script execution stopped."