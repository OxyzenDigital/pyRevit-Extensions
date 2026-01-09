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
    Reads excel using NPOI. 
    Supports filtering by sheet_name (str) and range_name (str, e.g. "Print_Area").
    """
    data = {
        'sheet_name': '',
        'row_heights': {},
        'column_widths': {},
        'cells': []
    }
    
    if not os.path.exists(file_path):
        return None
        
    try:
        with FileStream(file_path, FileMode.Open, FileAccess.Read) as fs:
            wb = WorkbookFactory.Create(fs)
            
            # 1. Resolve Sheet and Range
            target_sheet = None
            first_row, last_row = 0, 0
            first_col, last_col = 0, 0
            
            # If range name provided, try to parse it
            if range_name:
                # Find the named range
                n_idx = wb.GetNameIndex(range_name)
                
                if n_idx != -1:
                    name_obj = wb.GetNameAt(n_idx)
                    formula = name_obj.RefersToFormula
                    
                    try:
                        # Parse area reference
                        # Use SpreadsheetVersion.EXCEL2007 for XSSF
                        version = wb.GetSpreadsheetVersion()
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
                        except Exception as parse_ex:
                            print("Skipping named range '{}'. Only static ranges are supported. Formula: {}".format(range_name, formula))
                            # return None to indicate failure to load this specific range
                            return None 
                            
                    except Exception as ex_area:
                        print("Warning: Could not parse named range formula '{}': {}".format(formula, str(ex_area)))
                        return None
            
            # Fallback if no range or parsing failed: Use Sheet Name
            if not target_sheet:
                if sheet_name:
                    target_sheet = wb.GetSheet(sheet_name)
                else:
                    target_sheet = wb.GetSheetAt(0)
                
                if target_sheet:
                    first_row = target_sheet.FirstRowNum
                    last_row = target_sheet.LastRowNum
                    
                    # Reset cols (will determine per row)
                    first_col = 0
                    last_col = 0 
            
            if not target_sheet:
                return None

            data['sheet_name'] = target_sheet.SheetName
            
            # 2. Iterate Data
            for i in range(first_row, last_row + 1):
                row = target_sheet.GetRow(i)
                if not row: continue
                
                # Height
                data['row_heights'][str(i + 1)] = row.HeightInPoints
                
                # Determine loop range for cells
                c_start = first_col if range_name else row.FirstCellNum
                # Ensure c_start is valid (FirstCellNum can be -1 for empty rows)
                if c_start < 0: c_start = 0
                
                c_end = last_col + 1 if range_name else row.LastCellNum
                if c_end < c_start: c_end = c_start # Handle empty
                
                for j in range(c_start, c_end):
                    if range_name and (j < first_col or j > last_col):
                        continue
                        
                    try:
                        cell = row.GetCell(j)
                        if not cell: continue 
                        
                        # Column Width
                        width_units = target_sheet.GetColumnWidth(j)
                        data['column_widths'][str(j + 1)] = width_units / 256.0 
                        
                        # Value extraction
                        val = ""
                        ctype = cell.CellType
                        if ctype == CellType.Formula:
                            try:
                                ctype = cell.CachedFormulaResultType
                            except: pass
                        
                        if ctype == CellType.String:
                            val = cell.StringCellValue
                        elif ctype == CellType.Numeric:
                            try:
                                if DateUtil.IsCellDateFormatted(cell):
                                    val = str(cell.DateCellValue)
                                else:
                                    val = str(cell.NumericCellValue)
                            except:
                                # Fallback if date check or numeric fetch fails
                                val = str(cell.NumericCellValue)
                        elif ctype == CellType.Boolean:
                            val = str(cell.BooleanCellValue)
                        elif ctype == CellType.Error:
                            val = "ERROR"
                        
                        # Style extraction
                        cell_style = cell.CellStyle
                        font = cell_style.GetFont(wb)
                        
                        font_data = {
                            'name': font.FontName,
                            'size': font.FontHeightInPoints,
                            'bold': font.IsBold,
                            'italic': font.IsItalic,
                            'underline': font.Underline != 0, 
                            'color': None 
                        }
                        
                        borders = {
                            'left': str(cell_style.BorderLeft),
                            'right': str(cell_style.BorderRight),
                            'top': str(cell_style.BorderTop),
                            'bottom': str(cell_style.BorderBottom),
                        }
                        
                        cell_data = {
                            'row': i + 1,
                            'col': j + 1,
                            'value': val,
                            'font': font_data,
                            'fill': {'color': None},
                            'borders': borders
                        }
                        data['cells'].append(cell_data)
                    except Exception as cell_ex:
                        # Skip bad cells without crashing
                        print("Warning: Failed to read cell [{}, {}]: {}".format(i+1, j+1, str(cell_ex)))
                        continue
                    
    except Exception as e:
        print("Error reading Excel: " + str(e))
        import traceback
        traceback.print_exc()
        return None
        
    return data