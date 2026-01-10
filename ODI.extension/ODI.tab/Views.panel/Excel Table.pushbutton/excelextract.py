import clr
import sys
import os

# Setup path to bundled NPOI libraries
script_dir = os.path.dirname(__file__)
lib_dir = os.path.join(script_dir, "lib")

if os.path.exists(lib_dir):
    sys.path.append(lib_dir)
    try:
        # Load Dependencies first using FULL PATHS
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "BouncyCastle.Crypto.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "ICSharpCode.SharpZipLib.dll"))
        
        # Load NPOI using FULL PATHS
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.OOXML.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.OpenXml4Net.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.OpenXmlFormats.dll"))
    except Exception:
        # Fallback
        try:
            sys.path.append(lib_dir)
            clr.AddReference("BouncyCastle.Crypto")
            clr.AddReference("ICSharpCode.SharpZipLib")
            clr.AddReference("NPOI")
            clr.AddReference("NPOI.OOXML")
            clr.AddReference("NPOI.OpenXml4Net")
            clr.AddReference("NPOI.OpenXmlFormats")
        except Exception:
             pass
    except Exception:
        pass

from NPOI.SS.UserModel import WorkbookFactory, CellType, DateUtil, FillPattern
from NPOI.SS.Util import AreaReference, CellReference
from NPOI.SS import SpreadsheetVersion
from NPOI.XSSF.UserModel import XSSFWorkbook, XSSFColor
from System.IO import FileStream, FileMode, FileAccess, FileShare

# Helper for Colors
def get_rgb(color_obj, cell_style=None):
    rgb_bytes = None
    alpha = 255
    tint = 0.0
    
    # 1. Try XSSFColor (ARGB or RGB + Tint)
    if color_obj:
        # Check for Auto
        if hasattr(color_obj, "IsAuto") and color_obj.IsAuto:
            return None, 0

        # Get Tint
        if hasattr(color_obj, "Tint"):
            tint = color_obj.Tint

        # Try RGB (Red, Green, Blue) - Priority
        if hasattr(color_obj, "RGB"):
            try:
                b = color_obj.RGB
                if b and len(b) == 3:
                    rgb_bytes = [int(b[0]) & 0xFF, int(b[1]) & 0xFF, int(b[2]) & 0xFF]
            except: pass

        # Try ARgb (Alpha, Red, Green, Blue) - Fallback
        if not rgb_bytes and hasattr(color_obj, "ARgb"):
            try:
                b = color_obj.ARgb
                if b and len(b) == 4:
                    alpha = int(b[0]) & 0xFF
                    rgb_bytes = [int(b[1]) & 0xFF, int(b[2]) & 0xFF, int(b[3]) & 0xFF]
                elif b and len(b) == 3:
                    rgb_bytes = [int(b[0]) & 0xFF, int(b[1]) & 0xFF, int(b[2]) & 0xFF]
            except: pass

    # 2. Try HSSFColor (RGB only, no tint)
    if not rgb_bytes and color_obj and hasattr(color_obj, "RGB"):
        try:
            b = color_obj.RGB 
            if b and len(b) == 3:
                rgb_bytes = [int(b[0]) & 0xFF, int(b[1]) & 0xFF, int(b[2]) & 0xFF]
        except: pass

    # 3. Process RGB
    if rgb_bytes:
        red, green, blue = rgb_bytes
        
        # Apply Tint
        if tint != 0:
            def apply_tint(c, t):
                val = float(c)
                if t > 0: val = val * (1.0 - t) + 255.0 * t
                else: val = val * (1.0 + t)
                return int(round(val))
            
            red = apply_tint(red, tint)
            green = apply_tint(green, tint)
            blue = apply_tint(blue, tint)

        if alpha == 0: return None, 0 # Fully transparent

        rgb_str = "{},{},{}".format(red, green, blue)
        transparency = int(100 - (alpha / 255.0 * 100))
        
        return rgb_str, transparency

    # 4. Indexed Color Fallback
    if cell_style:
        try:
            idx = cell_style.FillForegroundColor
            if idx == 64: return None, 0 # Auto
            if idx == 0 or idx == 8: return "0,0,0", 0 # Black
            if idx == 1 or idx == 9: return "255,255,255", 0 # White
            if idx == 10: return "255,0,0", 0 # Red
            if idx == 11: return "0,255,0", 0 # Green
            if idx == 12: return "0,0,255", 0 # Blue
            if idx == 13: return "255,255,0", 0 # Yellow
            if idx == 14: return "255,0,255", 0 # Magenta
            if idx == 15: return "0,255,255", 0 # Cyan
            # Extended Palette (Standard Excel Colors)
            if idx == 16: return "128,0,0", 0      # Dark Red
            if idx == 17: return "0,128,0", 0      # Dark Green
            if idx == 18: return "0,0,128", 0      # Dark Blue
            if idx == 19: return "128,128,0", 0    # Dark Yellow
            if idx == 20: return "128,0,128", 0    # Dark Magenta
            if idx == 21: return "0,128,128", 0    # Teal
            if idx == 22: return "192,192,192", 0  # Silver
            if idx == 23: return "128,128,128", 0  # Grey
            # Add more standard index mappings if needed
        except: pass
    
    return None, 0

def get_sheet_names(file_path):
    names = []
    if not os.path.exists(file_path): return names
    try:
        with FileStream(file_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite) as fs:
            wb = WorkbookFactory.Create(fs)
            for i in range(wb.NumberOfSheets):
                names.append(wb.GetSheetName(i))
    except Exception as e:
        print("Error getting sheets: " + str(e))
    return names

def get_print_areas(file_path):
    areas = []
    if not os.path.exists(file_path): return areas
    try:
        with FileStream(file_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite) as fs:
            wb = WorkbookFactory.Create(fs)
            for i in range(wb.NumberOfNames):
                name = wb.GetNameAt(i)
                if name.IsDeleted: continue
                
                # Check for Print_Area or user defined ranges
                # Print_Area is usually specific to a sheet, identified by SheetIndex
                sheet_name = ""
                if name.SheetIndex >= 0 and name.SheetIndex < wb.NumberOfSheets:
                    sheet_name = wb.GetSheetName(name.SheetIndex)
                
                # We return a dict for the UI
                areas.append({
                    'name': name.NameName,
                    'sheet': sheet_name,
                    'formula': name.RefersToFormula
                })
    except Exception as e:
        print("Error getting ranges: " + str(e))
    return areas

def get_excel_data(file_path, sheet_name=None, range_name=None):
    data = {
        'sheet_name': '',
        'row_heights': {},
        'column_widths': {}, 
        'cells': []
    }
    
    if not os.path.exists(file_path): return None
        
    try:
        with FileStream(file_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite) as fs:
            wb = WorkbookFactory.Create(fs)
            
            target_sheet = None
            first_row, last_row = 0, 0
            first_col, last_col = 0, 0
            valid_range_found = False
            
            # Resolve Range
            if range_name:
                found_name = None
                for i in range(wb.NumberOfNames):
                    try:
                        name_obj = wb.GetNameAt(i)
                        if name_obj.IsDeleted: continue
                        if name_obj.NameName == range_name:
                            # If sheet_name is specified, ensure match
                            if sheet_name:
                                if name_obj.SheetIndex >= 0:
                                    if wb.GetSheetName(name_obj.SheetIndex) == sheet_name:
                                        found_name = name_obj
                                        break
                                else:
                                    # Name is global, might refer to our sheet
                                    found_name = name_obj
                            else:
                                found_name = name_obj
                                break
                    except: continue
                
                if found_name:
                    try:
                        version = SpreadsheetVersion.EXCEL2007
                        area_ref = AreaReference(found_name.RefersToFormula, version)
                        refs = area_ref.GetAllReferencedCells()
                        if refs and len(refs) > 0:
                            s_name = refs[0].SheetName
                            target_sheet = wb.GetSheet(s_name)
                            first_row = min(c.Row for c in refs)
                            last_row = max(c.Row for c in refs)
                            first_col = min(c.Col for c in refs)
                            last_col = max(c.Col for c in refs)
                            valid_range_found = True
                    except: pass
            
            # Fallback to Sheet
            if not target_sheet:
                if sheet_name: target_sheet = wb.GetSheet(sheet_name)
                else: target_sheet = wb.GetSheetAt(0)
                
                if target_sheet:
                    first_row = target_sheet.FirstRowNum
                    last_row = target_sheet.LastRowNum
                    # Initialize cols logic later per row or scan
                    valid_range_found = False 

            if not target_sheet: return None
            data['sheet_name'] = target_sheet.SheetName
            
            # Pre-process Merged Regions into a map for O(1) lookup
            merge_map = {}
            for i in range(target_sheet.NumMergedRegions):
                region = target_sheet.GetMergedRegion(i)
                for r in range(region.FirstRow, region.LastRow + 1):
                    for c in range(region.FirstColumn, region.LastColumn + 1):
                        merge_map[(r, c)] = region

            # Read Data
            for i in range(first_row, last_row + 1):
                row = target_sheet.GetRow(i)
                if not row:
                    # Capture default height for empty rows to maintain vertical spacing
                    data['row_heights'][str(i + 1)] = target_sheet.DefaultRowHeightInPoints
                    continue
                
                data['row_heights'][str(i + 1)] = row.HeightInPoints if row.HeightInPoints >= 0 else target_sheet.DefaultRowHeightInPoints
                
                c_start = first_col if valid_range_found else row.FirstCellNum
                c_end = last_col + 1 if valid_range_found else row.LastCellNum
                
                if c_start < 0: c_start = 0
                if c_end < c_start: continue 

                for j in range(c_start, c_end):
                    # Check Merge Status
                    region = merge_map.get((i, j))
                    is_head = False
                    
                    if region:
                        if region.FirstRow == i and region.FirstColumn == j:
                            is_head = True
                        else:
                            # Skip cells that are part of a merge but not the head
                            continue

                    cell = row.GetCell(j)
                    if not cell: cell = row.CreateCell(j)

                    width_px = target_sheet.GetColumnWidthInPixels(j)
                    data['column_widths'][str(j + 1)] = width_px 
                    
                    # Extract Value
                    val = ""
                    ctype = cell.CellType
                    if ctype == CellType.Formula:
                        try: ctype = cell.CachedFormulaResultType
                        except: pass
                    
                    if ctype == CellType.String: val = cell.StringCellValue or ""
                    elif ctype == CellType.Numeric:
                        try:
                            if DateUtil.IsCellDateFormatted(cell): val = str(cell.DateCellValue)
                            else: val = str(cell.NumericCellValue)
                        except: val = str(cell.NumericCellValue)
                    elif ctype == CellType.Boolean: val = str(cell.BooleanCellValue)
                    elif ctype == CellType.Error: val = ""
                    
                    # Extract Style
                    style = cell.CellStyle
                    font = style.GetFont(wb)
                    
                    bg_color_str = None
                    if style.FillPattern != FillPattern.NoFill:
                        # 1. Try Foreground (High Precision) - No Index Fallback
                        bg_color_str = get_rgb(style.FillForegroundColorColor, None)
                        
                        # 2. Try Background (High Precision) - NPOI sometimes swaps these for Solid fills
                        if (not bg_color_str or not bg_color_str[0]) and style.FillPattern == FillPattern.SolidForeground:
                            bg_color_str = get_rgb(style.FillBackgroundColorColor, None)
                            
                        # 3. Fallback to Indexed Color (Low Precision)
                        if not bg_color_str or not bg_color_str[0]:
                            bg_color_str = get_rgb(style.FillForegroundColorColor, style)
                    
                    f_color_obj = None
                    if hasattr(font, "GetXSSFColor"):
                        # XSSF Font Color
                        f_color_obj = font.GetXSSFColor()
                    elif hasattr(font, "GetHSSFColor"):
                        # HSSF Font Color
                        f_color_obj = font.GetHSSFColor(wb)
                    font_color = get_rgb(f_color_obj)
                    
                    font_data = {
                        'name': font.FontName,
                        'size': font.FontHeightInPoints,
                        'bold': font.IsBold,
                        'italic': font.IsItalic,
                        'underline': font.Underline != 0, 
                        'color': font_color
                    }
                    
                    b_left, b_right, b_top, b_bottom = str(style.BorderLeft), str(style.BorderRight), str(style.BorderTop), str(style.BorderBottom)
                    
                    # Merge Border Logic
                    if is_head and region:
                        # Right Border
                        if region.LastColumn > j:
                            r_row = target_sheet.GetRow(i)
                            if r_row:
                                # Check the cell at the actual edge of the merge
                                r_cell = r_row.GetCell(region.LastColumn) 
                                if r_cell and r_cell.CellStyle:
                                    b_right = str(r_cell.CellStyle.BorderRight)
                                # Fallback: if edge cell is null, use head cell style (common in some Excel writers)
                                elif style: b_right = str(style.BorderRight)
                        # Bottom Border
                        if region.LastRow > i:
                            b_row = target_sheet.GetRow(region.LastRow)
                            if b_row:
                                b_cell = b_row.GetCell(j) # Check bottom-left cell of the column
                                if b_cell and b_cell.CellStyle:
                                    b_bottom = str(b_cell.CellStyle.BorderBottom)
                                elif style: b_bottom = str(style.BorderBottom)

                    borders = {
                        'left': b_left,
                        'right': b_right,
                        'top': b_top,
                        'bottom': b_bottom,
                    }
                    
                    cell_data = {
                        'row': i + 1,
                        'col': j + 1,
                        'value': val,
                        'font': font_data,
                        'fill': {'color': bg_color_str},
                        'borders': borders,
                        'align': str(style.Alignment),
                        'v_align': str(style.VerticalAlignment),
                        'wrap_text': style.WrapText
                    }
                    
                    if is_head and region:
                        cell_data['r_span'] = region.LastRow - region.FirstRow + 1
                        cell_data['c_span'] = region.LastColumn - region.FirstColumn + 1
                        
                    data['cells'].append(cell_data)
                    
    except Exception as e:
        print("Error reading Excel: " + str(e))
        return None
        
    return data
