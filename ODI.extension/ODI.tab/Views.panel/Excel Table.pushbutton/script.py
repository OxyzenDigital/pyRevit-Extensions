import os
import sys
import clr
import System
from pyrevit import revit, DB, forms, script
import excelextract

# --- Constants & Schema ---
SCHEMA_GUID = System.Guid("72945C5C-002F-433A-9610-061093123456")
SCHEMA_NAME = "ODISmartExcelTable"

def get_schema():
    schema = DB.ExtensibleStorage.Schema.Lookup(SCHEMA_GUID)
    if not schema:
        builder = DB.ExtensibleStorage.SchemaBuilder(SCHEMA_GUID)
        builder.SetReadAccessLevel(DB.ExtensibleStorage.AccessLevel.Public)
        builder.SetWriteAccessLevel(DB.ExtensibleStorage.AccessLevel.Public)
        builder.SetSchemaName(SCHEMA_NAME)
        builder.AddSimpleField("SourcePath", str)
        builder.AddSimpleField("SheetName", str)
        builder.AddSimpleField("RangeName", str)
        schema = builder.Finish()
    return schema

def save_metadata(element, source_path, sheet_name, range_name):
    if not element: return
    try:
        schema = get_schema()
        entity = DB.ExtensibleStorage.Entity(schema)
        entity.Set("SourcePath", source_path or "")
        entity.Set("SheetName", sheet_name or "")
        entity.Set("RangeName", range_name or "")
        with DB.Transaction(element.Document, "Save Table Metadata") as t:
            t.Start()
            element.SetEntity(entity)
            t.Commit()
    except Exception as e:
        print("Failed to save metadata: " + str(e))

def get_metadata(element):
    try:
        schema = get_schema()
        entity = element.GetEntity(schema)
        if entity.IsValid():
            return {
                "SourcePath": entity.Get[str]("SourcePath"),
                "SheetName": entity.Get[str]("SheetName"),
                "RangeName": entity.Get[str]("RangeName")
            }
    except: pass
    return None

# --- Helpers ---
def points_to_feet(pts):
    return pts * (1.0 / 72.0) * (1.0 / 12.0)

def width_units_to_feet(w):
    if w <= 0: return 0.01
    return w * (1.0 / 10.0) * (1.0 / 12.0) * 8.0 

def get_family_template():
    app = revit.doc.Application
    versions = ["2020", "2021", "2022", "2023", "2024", "2025"]
    base_paths = [r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English\Annotations",
                  r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English-Imperial\Annotations"]
    for v in versions:
        for b in base_paths:
            p = os.path.join(b.format(v), "Generic Annotation.rft")
            if os.path.exists(p): return p
    return forms.pick_file(file_ext='rft', title="Select 'Generic Annotation' Template")

def draw_schedule_in_family(fam_doc, data):
    """Draws the schedule geometry into the given family document."""
    # Find active view
    ref_view = fam_doc.ActiveView
    if not ref_view:
        ref_view = next((v for v in DB.FilteredElementCollector(fam_doc).OfClass(DB.View).ToElements() if v.Name == "Ref. Level"), None)
    if not ref_view:
         ref_view = DB.FilteredElementCollector(fam_doc).OfClass(DB.ViewPlan).FirstElement()
    
    if not ref_view: return False

    with DB.Transaction(fam_doc, "Draw Excel Data") as t:
        t.Start()
        
        # Cleanup existing (for updates)
        collector = DB.FilteredElementCollector(fam_doc, ref_view.Id)
        ids_to_del = []
        for el in collector.ToElements():
            if isinstance(el, (DB.DetailCurve, DB.TextNote, DB.FilledRegion)):
                ids_to_del.append(el.Id)
        
        if ids_to_del:
            # Convert python list to ICollection[ElementId]
            col = System.Collections.Generic.List[DB.ElementId](ids_to_del)
            fam_doc.Delete(col)

        # Setup Fonts
        def_text_type = DB.FilteredElementCollector(fam_doc).OfClass(DB.TextNoteType).FirstElement()
        
        # Pre-calc offsets
        sorted_rows = sorted([int(k) for k in data['row_heights'].keys()])
        row_y = {}
        curr_y = 0.0
        for r in sorted_rows:
            row_y[r] = curr_y
            rh = data['row_heights'].get(str(r), 12.75)
            curr_y -= max(points_to_feet(rh), 0.01)
        
        sorted_cols = sorted([int(k) for k in data['column_widths'].keys()])
        col_x = {}
        curr_x = 0.0
        for c in sorted_cols:
            col_x[c] = curr_x
            cw = data['column_widths'].get(str(c), 10.0)
            curr_x += max(width_units_to_feet(cw), 0.01)

        # Draw Cells
        for cell in data['cells']:
            r = cell['row']
            c = cell['col']
            if r not in row_y or c not in col_x: continue
                
            x = col_x[c]
            y = row_y[r]
            
            rh_pts = data['row_heights'].get(str(r), 12.75)
            cw_units = data['column_widths'].get(str(c), 10.0)
            h = max(points_to_feet(rh_pts), 0.01)
            w = max(width_units_to_feet(cw_units), 0.01)
            
            # Draw Box
            rect = [
                DB.Line.CreateBound(DB.XYZ(x, y, 0), DB.XYZ(x+w, y, 0)),
                DB.Line.CreateBound(DB.XYZ(x+w, y, 0), DB.XYZ(x+w, y-h, 0)),
                DB.Line.CreateBound(DB.XYZ(x+w, y-h, 0), DB.XYZ(x, y-h, 0)),
                DB.Line.CreateBound(DB.XYZ(x, y-h, 0), DB.XYZ(x, y, 0))
            ]
            
            for line in rect:
                try: fam_doc.FamilyCreate.NewDetailCurve(ref_view, line)
                except: pass
            
            # Draw Text
            val = cell['value']
            if val and def_text_type:
                center = DB.XYZ(x + w/2.0, y - h/2.0, 0)
                try:
                    DB.TextNote.Create(fam_doc, ref_view.Id, center, val, def_text_type.Id)
                except:
                    try: fam_doc.FamilyCreate.NewTextNote(val, center, def_text_type)
                    except: pass
        
        t.Commit()
    return True

# --- Core Logic ---
def create_new_schedule(file_path, sheet_name, range_name, schedule_name, template_path):
    data = excelextract.get_excel_data(file_path, sheet_name, range_name)
    if not data: return None
    
    app = revit.doc.Application
    fam_doc = app.NewFamilyDocument(template_path)
    if not fam_doc: return None
    
    success = draw_schedule_in_family(fam_doc, data)
    
    if success:
        # Load
        class FamLoadOpt(DB.IFamilyLoadOptions):
            def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                overwriteParameterValues.Value = True
                return True
            def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
                return True
        
        # Handle Naming: SaveAs not strictly needed if we rename via FamilyName parameter?
        # Actually NewFamilyDocument creates "Family1". We should SaveAs to set name.
        # But we can't easily save to temp without path.
        # Alternative: Load it, then rename the symbol in project.
        
        fam = fam_doc.LoadFamily(revit.doc, FamLoadOpt())
        fam_doc.Close(False)
        
        if fam:
            # Rename Family
            try:
                with DB.Transaction(revit.doc, "Rename Family") as t:
                    t.Start()
                    fam.Name = schedule_name
                    t.Commit()
            except: pass # Might fail if name exists
            
            # Save Metadata to the FAMILY element
            save_metadata(fam, file_path, sheet_name, range_name)
            
            # Notify User
            forms.alert("Excel Table family '{}' created and loaded.\nYou can now place it from the Annotations browser.".format(fam.Name))
            return fam
    else:
        fam_doc.Close(False)
    return None

def update_selected_instance(instance):
    # Metadata is now on the Family, not the instance
    fam = instance.Symbol.Family
    meta = get_metadata(fam)
    
    if not meta:
        # Backward compatibility: check instance (if any old ones exist)
        meta = get_metadata(instance)
        
    if not meta:
        forms.alert("This element is not linked to an Excel file.")
        return
    
    f_path = meta["SourcePath"]
    if not os.path.exists(f_path):
        forms.alert("Source file not found: " + f_path)
        return
        
    data = excelextract.get_excel_data(f_path, meta["SheetName"], meta["RangeName"])
    if not data: return
    
    # Edit Family
    fam_doc = revit.doc.EditFamily(fam)
    
    success = draw_schedule_in_family(fam_doc, data)
    
    if success:
        class FamLoadOpt(DB.IFamilyLoadOptions):
            def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                overwriteParameterValues.Value = True
                return True
            def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
                return True
        fam_doc.LoadFamily(revit.doc, FamLoadOpt())
        fam_doc.Close(False)
        print("Table updated successfully.")
    else:
        fam_doc.Close(False)

class NamedRangeItem(object):
    def __init__(self, data):
        self.name = data['name']
        self.sheet = data['sheet'] if data['sheet'] else "Workbook"
        self.formula = data['formula']

# --- UI Class ---
class ExcelScheduleWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.excel_path = None
        self.sheets = []
        self.ranges = []
        self.template_path = None

    def Btn_Browse_Click(self, sender, args):
        f = forms.pick_file(file_ext='xlsx')
        if f:
            self.excel_path = f
            self.Tb_FilePath.Text = f
            self.refresh_data()

    def refresh_data(self):
        if not self.excel_path: return
        self.Lb_Status.Text = "Reading file..."
        
        # Load Sheets
        self.sheets = excelextract.get_sheet_names(self.excel_path)
        self.Cb_Sheets.ItemsSource = self.sheets
        if self.sheets: self.Cb_Sheets.SelectedIndex = 0
        
        # Load Ranges
        raw_ranges = excelextract.get_print_areas(self.excel_path)
        self.ranges = [NamedRangeItem(r) for r in raw_ranges]
        self.Cb_Ranges.ItemsSource = self.ranges
        # DisplayMemberPath removed to avoid conflict with XAML ItemTemplate
        
        self.Lb_Status.Text = "Ready"

    def Cb_Sheets_SelectionChanged(self, sender, args):
        if self.Cb_Sheets.SelectedItem:
            # Default: Sheet Name + Table
            self.Tb_Name.Text = "{} Table".format(self.Cb_Sheets.SelectedItem)

    def Cb_Ranges_SelectionChanged(self, sender, args):
        sel = self.Cb_Ranges.SelectedItem
        if sel:
            raw_name = sel.name # Access property
            sheet_name = sel.sheet
            
            # Smart naming
            if raw_name == "Print_Area":
                clean_name = "{} Print Area".format(sheet_name)
            else:
                # Replace underscores with spaces for readability
                clean_name = raw_name.replace("_", " ")
                
            self.Tb_Name.Text = clean_name

    def Btn_Import_Click(self, sender, args):
        if not self.excel_path: return
        
        s_name = self.Cb_Sheets.SelectedItem
        r_name = None
        if self.Cb_Ranges.SelectedItem:
            r_name = self.Cb_Ranges.SelectedItem.name # Access property
            
        n_name = self.Tb_Name.Text
        if not n_name: n_name = "Excel Table"
        
        # Get Template
        if not self.template_path:
             self.template_path = get_family_template()
             
        if not self.template_path:
            forms.alert("Template required.")
            return

        self.Close()
        
        # Run Creation
        create_new_schedule(self.excel_path, s_name, r_name, n_name, self.template_path)

    def Btn_Close_Click(self, sender, args):
        self.Close()

# --- Main Entry Point ---
sel = revit.get_selection()
if len(sel) == 1 and isinstance(sel[0], DB.FamilyInstance):
    # Check if it's our smart table
    # Metadata is on the Family
    fam = sel[0].Symbol.Family
    meta = get_metadata(fam)
    
    # Fallback for older instances
    if not meta:
        meta = get_metadata(sel[0])
        
    if meta and meta["SourcePath"]:
        res = forms.alert("Selected element is a Smart Excel Table.\nUpdate from source?", options=["Update", "Create New..."])
        if res == "Update":
            update_selected_instance(sel[0])
            script.exit()

# Default: Open UI
win = ExcelScheduleWindow()
win.ShowDialog()
