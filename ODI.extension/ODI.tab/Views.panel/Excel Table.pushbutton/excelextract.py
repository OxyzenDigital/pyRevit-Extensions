import clr
import sys
import os

# Setup path to bundled NPOI libraries
script_dir = os.path.dirname(__file__)
lib_dir = os.path.join(script_dir, "lib")

if os.path.exists(lib_dir):
    sys.path.append(lib_dir)
    
    try:
        # Load Dependencies first using FULL PATHS with AddReferenceToFileAndPath
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "BouncyCastle.Crypto.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "ICSharpCode.SharpZipLib.dll"))
        
        # Load NPOI using FULL PATHS with AddReferenceToFileAndPath
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.OOXML.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.OpenXml4Net.dll"))
        clr.AddReferenceToFileAndPath(os.path.join(lib_dir, "NPOI.OpenXmlFormats.dll"))
    except AttributeError:
        # Fallback if AddReferenceToFileAndPath is missing (some implementations)
        # Try adding to path and loading by filename
        try:
            sys.path.append(lib_dir)
            clr.AddReference("BouncyCastle.Crypto")
            clr.AddReference("ICSharpCode.SharpZipLib")
            clr.AddReference("NPOI")
            clr.AddReference("NPOI.OOXML")
            clr.AddReference("NPOI.OpenXml4Net")
            clr.AddReference("NPOI.OpenXmlFormats")
        except Exception as e2:
             print("Fallback load failed: " + str(e2))
    except Exception as e:
        print("Error loading bundled NPOI libraries: " + str(e))
else:
    print("Error: 'lib' directory with NPOI libraries not found at: " + lib_dir)

from NPOI.SS.UserModel import WorkbookFactory, CellType, DateUtil
from NPOI.SS.Util import AreaReference, CellReference
from NPOI.SS import SpreadsheetVersion
from NPOI.XSSF.UserModel import XSSFWorkbook
from System.IO import FileStream, FileMode, FileAccess

def get_sheet_names(file_path):
    """Returns a list of sheet names in the Excel file."""
    if not os.path.exists(file_path): return []
    try:
        with FileStream(file_path, FileMode.Open, FileAccess.Read) as fs:
            wb = WorkbookFactory.Create(fs)
            return [wb.GetSheetName(i) for i in range(wb.NumberOfSheets)]
    except:
        return []

def get_print_areas(file_path):
    """
    Returns a list of Named Ranges (including Print_Area) found in the workbook.
    Format: [{'name': 'Print_Area', 'sheet': 'Sheet1', 'formula': 'Sheet1!$A$1:$F$20'}, ...]
    """
    if not os.path.exists(file_path): return []
    areas = []
    try:
        with FileStream(file_path, FileMode.Open, FileAccess.Read) as fs:
            wb = WorkbookFactory.Create(fs)
            for i in range(wb.NumberOfNames):
                try:
                    name = wb.GetNameAt(i)
                    if name.IsDeleted: continue
                    
                    # Skip if formula is complex/invalid/dynamic (starts with OFFSET etc or contains errors)
                    # Simple check: try to see if it looks like a range reference
                    formula = name.RefersToFormula
                    if not formula or "REF!" in formula: continue
                    
                    # Resolve sheet name
                    sheet_name = ""
                    if name.SheetIndex >= 0 and name.SheetIndex < wb.NumberOfSheets:
                        sheet_name = wb.GetSheetName(name.SheetIndex)
                    
                    # Use NameName for display
                    areas.append({
                        'name': name.NameName,
                        'sheet': sheet_name, 
                        'formula': formula
                    })
                except Exception as ex_name:
                    # Skip problematic names
                    continue
    except Exception as e:
        print("Error reading named ranges: " + str(e))
    return areas

def get_excel_data(file_path, sheet_name=None, range_name=None):
    """
    Reads excel using NPOI with high fidelity (Colors, Merges, Dimensions).
    """
    data = {
        'sheet_name': '',
        'row_heights': {},
        'column_widths': {}, # In pixels
        'cells': [],
        'merges': {} # Key: "r,c", Value: {'r_span': int, 'c_span': int}
    }
    
    if not os.path.exists(file_path):
        return None
        
    try:
        with FileStream(file_path, FileMode.Open, FileAccess.Read) as fs:
            wb = WorkbookFactory.Create(fs)
            
            # 1. Resolve Sheet and Range (Same logic as before)
            target_sheet = None
            first_row, last_row = 0, 0
            first_col, last_col = 0, 0
            valid_range_found = False
            
            if range_name:
                found_name = None
                for i in range(wb.NumberOfNames):
                    try:
                        name_obj = wb.GetNameAt(i)
                        if name_obj.IsDeleted: continue
                        if name_obj.NameName == range_name:
                            found_name = name_obj
                            break
                    except: continue
                
                if found_name:
                    name_obj = found_name
                    formula = name_obj.RefersToFormula
                    try:
                        version = SpreadsheetVersion.EXCEL2007
                        try:
                            area_ref = AreaReference(formula, version)
                            refs = area_ref.GetAllReferencedCells()
                            if refs and len(refs) > 0:
                                s_name = refs[0].SheetName
                                target_sheet = wb.GetSheet(s_name)
                                first_row = min(c.Row for c in refs)
                                last_row = max(c.Row for c in refs)
                                first_col = min(c.Col for c in refs)
                                last_col = max(c.Col for c in refs)
                                valid_range_found = True
                        except: return None 
                    except: return None
            
            if not target_sheet:
                if sheet_name: target_sheet = wb.GetSheet(sheet_name)
                else: target_sheet = wb.GetSheetAt(0)
                if target_sheet:
                    first_row = target_sheet.FirstRowNum
                    last_row = target_sheet.LastRowNum
                    first_col = 0
                    last_col = 0 
            
            if not target_sheet: return None

            data['sheet_name'] = target_sheet.SheetName
            
            # --- Pre-process Merged Regions ---
            # Map: "row,col" -> (row_span, col_span) for HEAD
            # Map: "row,col" -> "SKIP" for others
            merge_map = {}
            for i in range(target_sheet.NumMergedRegions):
                region = target_sheet.GetMergedRegion(i)
                # Check if region intersects our target range
                # Simple check: overlap
                # Only strictly necessary if we are filtering by range, but good optimization.
                # Actually, we just map all, and look them up.
                r_min, r_max = region.FirstRow, region.LastRow
                c_min, c_max = region.FirstColumn, region.LastColumn
                
                # Head
                key = "{},{}".format(r_min, c_min)
                data['merges'][key] = {
                    'r_span': r_max - r_min + 1,
                    'c_span': c_max - c_min + 1
                }
                
                # Mark others as skip
                for r in range(r_min, r_max + 1):
                    for c in range(c_min, c_max + 1):
                        if r == r_min and c == c_min: continue
                        merge_map["{},{}".format(r, c)] = "SKIP"

            # Helper for Colors
            def get_rgb(xssf_color):
                if not xssf_color: return None
                # Check if RGB is available
                # NPOI 2.5: GetARGB() -> bytes
                try:
                    argb = xssf_color.ARgb # Returns byte[]
                    if argb and len(argb) >= 3:
                        # ARGB or RGB? Usually [A, R, G, B] or [R, G, B]
                        if len(argb) == 4:
                            # Alpha ignored for Revit usually, or map to transparency?
                            # Revit fills are opaque.
                            return "{},{},{}".format(argb[1], argb[2], argb[3])
                        elif len(argb) == 3:
                            return "{},{},{}".format(argb[0], argb[1], argb[2])
                except: pass
                return None

            # 2. Iterate Data
            for i in range(first_row, last_row + 1):
                row = target_sheet.GetRow(i)
                if not row: continue
                
                data['row_heights'][str(i + 1)] = row.HeightInPoints
                
                c_start = first_col if valid_range_found else row.FirstCellNum
                if c_start < 0: c_start = 0
                c_end = last_col + 1 if valid_range_found else row.LastCellNum
                if c_end < c_start: c_end = c_start
                
                for j in range(c_start, c_end):
                    if valid_range_found and (j < first_col or j > last_col): continue
                    
                    # Check Merge Skip
                    if merge_map.get("{},{}".format(i, j)) == "SKIP":
                        continue

                    try:
                        cell = row.GetCell(j)
                        
                        # Dimensions (Pixels are better for width)
                        # GetColumnWidthInPixels returns float
                        width_px = target_sheet.GetColumnWidthInPixels(j)
                        data['column_widths'][str(j + 1)] = width_px 
                        
                        # If cell is null, we might still need to draw background if it's styled?
                        # NPOI row.GetCell(j) returns null if empty.
                        # But we might have a merged region starting here or a background.
                        # If null, create dummy to hold style? No, style is on cell.
                        # If cell is missing, we check if there's a default column style? Too complex.
                        # We proceed only if cell exists OR if it's a merge head (handled by logic above, but head must exist?)
                        # Actually, if cell is null, it has no style.
                        if not cell: 
                            # But if we have width/height, we should register it?
                            # data['cells'].append({'row': i+1, 'col': j+1, 'value': ''})
                            continue
                        
                        # Value
                        val = ""
                        ctype = cell.CellType
                        if ctype == CellType.Formula:
                            try: ctype = cell.CachedFormulaResultType
                            except: pass
                        
                        if ctype == CellType.String: val = cell.StringCellValue
                        elif ctype == CellType.Numeric:
                            try:
                                if DateUtil.IsCellDateFormatted(cell): val = str(cell.DateCellValue)
                                else: val = str(cell.NumericCellValue)
                            except: val = str(cell.NumericCellValue)
                        elif ctype == CellType.Boolean: val = str(cell.BooleanCellValue)
                        elif ctype == CellType.Error: val = ""
                        
                        # Style
                        style = cell.CellStyle
                        font = style.GetFont(wb)
                        
                        # Colors
                        fg_color = get_rgb(style.FillForegroundColorColor)
                        font_color = get_rgb(font.GetXSSFColor())
                        
                        font_data = {
                            'name': font.FontName,
                            'size': font.FontHeightInPoints,
                            'bold': font.IsBold,
                            'italic': font.IsItalic,
                            'underline': font.Underline != 0, 
                            'color': font_color
                        }
                        
                        borders = {
                            'left': str(style.BorderLeft),
                            'right': str(style.BorderRight),
                            'top': str(style.BorderTop),
                            'bottom': str(style.BorderBottom),
                            # Colors? style.LeftBorderColor (index) or ...Color (XSSF)
                            # 'left_color': get_rgb(style.LeftBorderColorColor) # NPOI property naming varies
                        }
                        
                        # Alignment
                        align = str(style.Alignment) # "Center", "Left", etc.
                        v_align = str(style.VerticalAlignment)
                        wrap_text = style.WrapText
                        
                        cell_data = {
                            'row': i + 1,
                            'col': j + 1,
                            'value': val,
                            'font': font_data,
                            'fill': {'color': fg_color},
                            'borders': borders,
                            'align': align,
                            'v_align': v_align,
                            'wrap_text': wrap_text
                        }
                        data['cells'].append(cell_data)
                    except Exception as cell_ex:
                        print("Warning: Failed to read cell [{}, {}]: {}".format(i+1, j+1, str(cell_ex)))
                        continue
                    
    except Exception as e:
        print("Error reading Excel: " + str(e))
        import traceback
        traceback.print_exc()
        return None
        
    return data