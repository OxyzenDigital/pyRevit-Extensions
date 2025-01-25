#!Python3
from Autodesk.Revit.DB import *
from pyrevit import revit, DB, UI

def create_legend_from_excel(json_path):
    """Creates a legend view in Revit based on extracted Excel data."""
    try:
        doc = revit.doc
        with open(json_path, "r") as json_file:
            excel_data = json.load(json_file)

        # Duplicate Legend View
        legend_view = FilteredElementCollector(doc).OfClass(View).WhereElementIsViewTemplate().Where(lambda v: v.Name == "Legend").First() #Get Legend View Template
        if legend_view:
            new_legend_view = View.Create(doc, legend_view.Id)
            new_legend_view.Name = "Excel Import Legend"
            new_legend_view.Scale = 4 # 1/4" scale

        else:
            print("No Legend View Template was found")
            return

        # Start a transaction
        TransactionManager.Instance.EnsureInTransaction(doc)

        # Revit units are in feet, Excel dimensions are in points. Conversion factor needed. 72 points = 1 inch
        points_to_feet = 1.0 / (72.0 * 12.0)

        # Use the JSON data to create Revit elements
        x_offset = 0
        y_offset = 0
        for row_index, row_data in enumerate(excel_data["cells"]):
            x_offset = 0
            for cell_index, cell_data in enumerate(row_data):

                #Filled Region
                fill_color = cell_data.get("fill_color")
                if fill_color:
                    r = int(fill_color[2:4], 16)
                    g = int(fill_color[4:6], 16)
                    b = int(fill_color[6:8], 16)
                    color = Color(r, g, b)
                    # Create a solid fill pattern (you might need to create it if it doesn't exist)
                    fill_pattern_element = FilteredElementCollector(doc).OfClass(FillPatternElement).Where(lambda x: x.Name == "Solid fill").First()
                    if fill_pattern_element:
                        fill_type = FillPatternElement.GetFillPattern(doc, fill_pattern_element.Id)
                        # Create a filled region
                        curve_loop = CurveLoop()
                        width = cell_data["width"] * points_to_feet
                        height = cell_data["height"] * points_to_feet
                        curve_loop.Add(Line.CreateBound(XYZ(x_offset, y_offset, 0), XYZ(x_offset + width, y_offset, 0)))
                        curve_loop.Add(Line.CreateBound(XYZ(x_offset + width, y_offset, 0), XYZ(x_offset + width, y_offset - height, 0)))
                        curve_loop.Add(Line.CreateBound(XYZ(x_offset + width, y_offset - height, 0), XYZ(x_offset, y_offset - height, 0)))
                        curve_loop.Add(Line.CreateBound(XYZ(x_offset, y_offset - height, 0), XYZ(x_offset, y_offset, 0)))

                        filled_region = FilledRegion.Create(doc, new_legend_view.Id, fill_type.Id, [curve_loop])
                        filled_region.Color = color

                #Text Note
                text_value = cell_data.get("value")
                if text_value:
                    text_note_options = TextNoteOptions(cell_data["font_name"], cell_data["font_size"] * points_to_feet)
                    text_note = TextNote.Create(doc, new_legend_view.Id, XYZ(x_offset + width/2, y_offset - height/2, 0), text_value, text_note_options)

                x_offset += cell_data["width"] * points_to_feet
            y_offset -= excel_data["dimensions"]["max_row_height"] * points_to_feet #Move down to the next row

        TransactionManager.Instance.TransactionTaskDone()
        print("Excel data imported to Revit Legend View.")

    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage (within pyRevit command):
json_file_path = r"C:\tmp\excel_data.json" # Replace with the actual path
create_legend_from_excel(json_file_path)