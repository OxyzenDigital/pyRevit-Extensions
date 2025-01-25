import openpyxl
import json
import os
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.styles import Font

# Default dimensions in points (based on Calibri 11)
DEFAULT_ROW_HEIGHT = 12.75
DEFAULT_COLUMN_WIDTH = 8.43  # This is an approximation in character units, not points

def extract_excel_data(excel_path):
    """Extracts data from Excel tables with robust error handling."""
    try:
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"File not found: {excel_path}")

        try:
            workbook = openpyxl.load_workbook(excel_path)
        except openpyxl.utils.exceptions.InvalidFileException:
            raise ValueError(f"Invalid Excel file: {excel_path}")

        sheet = workbook.active

        tables_data = {}

        if sheet.tables:
            for table_name in sheet.tables:
                try:
                    table = sheet.tables[table_name]
                    table_range = table.ref

                    if isinstance(table_range, str):
                        min_col, min_row, max_col, max_row = range_boundaries(table_range)
                    else:
                        min_col, min_row, max_col, max_row = table_range.min_col, table_range.min_row, table_range.max_col, table_range.max_row

                    table_data = []
                    for row_index in range(min_row, max_row + 1):
                        row_data = []
                        for col_index in range(min_col, max_col + 1):
                            try:
                                cell = sheet.cell(row=row_index, column=col_index)
                                col_letter = get_column_letter(col_index)

                                fill_color = cell.fill.start_color.rgb if cell.fill and cell.fill.start_color else "FFFFFFFF"
                                font = cell.font if cell.font else openpyxl.styles.Font()
                                border = cell.border if cell.border else openpyxl.styles.Border()

                                row_height = sheet.row_dimensions[row_index].height
                                col_width = sheet.column_dimensions[col_letter].width

                                cell_info = {
                                    "value": str(cell.value) if cell.value is not None else "",
                                    "font_name": font.name if font.name else "Calibri",
                                    "font_size": font.size if font.size else 11,
                                    "fill_color": fill_color,
                                    "border": {
                                        "left": border.left.style if border.left else None,
                                        "right": border.right.style if border.right else None,
                                        "top": border.top.style if border.top else None,
                                        "bottom": border.bottom.style if border.bottom else None,
                                    },
                                    "width": col_width if col_width is not None else DEFAULT_COLUMN_WIDTH,
                                    "height": row_height if row_height is not None else DEFAULT_ROW_HEIGHT,
                                }
                                row_data.append(cell_info)
                            except Exception as cell_err:
                                print(f"Error processing cell ({row_index}, {col_index}) in table '{table_name}': {cell_err}")
                                continue

                        table_data.append(row_data)

                    tables_data[table_name] = table_data

                except Exception as table_err:
                    print(f"Error processing table '{table_name}': {table_err}")
                    continue
        else:
            print("No Tables found in the excel file")
            return {"tables": {}}

        return {"tables": tables_data}

    except (FileNotFoundError, ValueError) as file_err:
        print(file_err)
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

# Example usage
excel_file = r"C:\tmp\LegendTransferTemplate.xlsx"
extracted_data = extract_excel_data(excel_file)

if extracted_data:
    output_path = os.path.join(os.path.dirname(excel_file), "excel_data.json")
    with open(output_path, "w") as json_file:
        json.dump(extracted_data, json_file, indent=4)
    print(f"Data saved to {output_path}")
else:
    print("Failed to extract data.")