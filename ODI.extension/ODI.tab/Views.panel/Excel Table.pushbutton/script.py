import os
import sys
import clr
import System
from pyrevit import revit, DB, forms, script
import excelextract

# --- Constants ---
SCHEMA_GUID = System.Guid("72945C5C-002F-433A-9610-061093123458")
SCHEMA_NAME = "ODISmartExcelTable"
FAMILY_PREFIX = "ODI_Excel_"

# --- Schema Manager ---
class TableMetadataManager(object):
    @staticmethod
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
            builder.AddSimpleField("TextScale", str)
            schema = builder.Finish()
        return schema

    @staticmethod
    def save(element, source_path, sheet_name, range_name, text_scale=100.0):
        if not element: return
        try:
            schema = TableMetadataManager.get_schema()
            entity = DB.ExtensibleStorage.Entity(schema)
            entity.Set("SourcePath", source_path or "")
            entity.Set("SheetName", sheet_name or "")
            entity.Set("RangeName", range_name or "")
            entity.Set("TextScale", str(text_scale))
            with DB.Transaction(element.Document, "Save Table Metadata") as t:
                t.Start()
                element.SetEntity(entity)
                t.Commit()
        except Exception as e:
            print("Failed to save metadata: " + str(e))

    @staticmethod
    def get(element):
        try:
            schema = TableMetadataManager.get_schema()
            entity = element.GetEntity(schema)
            if entity.IsValid():
                # Safe get for new field
                scale = 100.0
                try: scale = float(entity.Get[str]("TextScale"))
                except: pass
                
                return {
                    "SourcePath": entity.Get[str]("SourcePath"),
                    "SheetName": entity.Get[str]("SheetName"),
                    "RangeName": entity.Get[str]("RangeName"),
                    "TextScale": scale
                }
        except: pass
        return None

# --- Family Generation Logic ---
class FamilyGenerator(object):
    @staticmethod
    def pixels_to_feet(px):
        # NPOI Pixels -> 1/96 inch (Standard Screen DPI).
        if px <= 0: return 0.001
        return px * (1.0 / 96.0) * (1.0 / 12.0)

    @staticmethod
    def points_to_feet(pts):
        # 1 pt = 1/72 inch (Standard). 
        # User confirmed 1/96 failed for rows, so reverting to standard.
        return pts * (1.0 / 72.0) * (1.0 / 12.0)

    @staticmethod
    def text_points_to_feet(pts):
        # User Request: Treat Points as 96 DPI (Pixels) for consistency.
        return pts * (1.0 / 96.0) * (1.0 / 12.0)

    @staticmethod
    def get_color_from_string(rgb_str):
        if not rgb_str: return None
        try:
            parts = [int(x) for x in rgb_str.split(',')]
            if len(parts) == 3:
                return DB.Color(parts[0], parts[1], parts[2])
        except: pass
        return None

    @staticmethod
    def get_or_create_fill_type(doc, color, type_cache, solid_pat_id):
        name = "Solid_{}_{}_{}".format(color.Red, color.Green, color.Blue)
        
        if name in type_cache:
            return type_cache[name]
        
        # Base type for duplication (should be passed or found once, but for now find once if needed)
        
        base = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()
        if not base: return None
        
        try:
            new_type = base.Duplicate(name)
        except Exception:
            # Race condition or name collision? Try to find it again just in case
            for x in DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).ToElements():
                if x.Name == name:
                    type_cache[name] = x
                    return x
            return None
        
        # Set Color (Support New and Old API)
        try: new_type.ForegroundPatternColor = color
        except: 
            try: new_type.Color = color
            except: pass
        
        if solid_pat_id:
            try: new_type.ForegroundPatternId = solid_pat_id
            except: 
                try: new_type.FillPatternId = solid_pat_id
                except: pass
        
        type_cache[name] = new_type
        return new_type

    @staticmethod
    def get_or_create_text_type(doc, font_data, type_cache):
        color = FamilyGenerator.get_color_from_string(font_data['color']) or DB.Color(0,0,0)
        
        name = "{}_{}_{}{}_{}_{}_{}".format(
            font_data['name'], 
            int(font_data['size']),
            "B" if font_data['bold'] else "",
            "I" if font_data['italic'] else "",
            color.Red, color.Green, color.Blue
        ).replace(" ", "")
        
        if name in type_cache:
            return type_cache[name]
        
        base = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).FirstElement()
        if not base: return None
        
        try:
            new_type = base.Duplicate(name)
        except Exception:
             # Fallback lookup
            for x in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).ToElements():
                if x.Name == name:
                    type_cache[name] = x
                    return x
            return None
        
        # Set Props
        # Font Name
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_FONT).Set(font_data['name'])
        # Size (pts to ft) - Use corrected text size
        size_ft = FamilyGenerator.text_points_to_feet(font_data['size'])
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE).Set(size_ft)
        # Bold/Italic
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_BOLD).Set(1 if font_data['bold'] else 0)
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_ITALIC).Set(1 if font_data['italic'] else 0)
        # Color
        new_type.get_Parameter(DB.BuiltInParameter.LINE_COLOR).Set(color.Red + (color.Green << 8) + (color.Blue << 16))
        
        # Width Factor & Background
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE).Set(1.0)
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_BACKGROUND).Set(1) # 1 = Transparent
        
        type_cache[name] = new_type
        return new_type

    @staticmethod
    def get_template_path():
        versions = ["2020", "2021", "2022", "2023", "2024", "2025"]
        base_paths = [r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English\Annotations",
                      r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English-Imperial\Annotations"]
        for v in versions:
            for b in base_paths:
                p = os.path.join(b.format(v), "Generic Annotation.rft")
                if os.path.exists(p): return p
        return forms.pick_file(file_ext='rft', title="Select 'Generic Annotation' Template")

    @staticmethod
    def ensure_line_styles(doc):
        style_map = {}
        try:
            # Safe Access to Category
            cat = None
            try: cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_GenericAnnotation)
            except: pass
            
            if not cat: return {}
            
            styles = {
                "ODI_Wide": 5,
                "ODI_Medium": 3,
                "ODI_Thin": 1
            }
            
            for name, weight in styles.items():
                subcat = None
                if cat.SubCategories.Contains(name):
                    subcat = cat.SubCategories.get_Item(name)
                else:
                    try:
                        subcat = doc.Settings.Categories.NewSubcategory(cat, name)
                    except: pass
                
                if subcat:
                    try: subcat.LineWeight = weight
                    except: pass
                    
                    # Find GraphicsStyle
                    gs_col = DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle).ToElements()
                    found_gs = next((g for g in gs_col if g.GraphicsStyleCategory.Name == name), None)
                    if found_gs:
                        style_map[name] = found_gs.Id
        except: pass
        return style_map

    @staticmethod
    def draw_content(fam_doc, data, text_scale=1.0):
        # View Finding
        ref_view = next((v for v in DB.FilteredElementCollector(fam_doc).OfClass(DB.View).ToElements() 
                         if not v.IsTemplate and v.ViewType != DB.ViewType.ProjectBrowser), None)
        if not ref_view: return False

        with DB.Transaction(fam_doc, "Draw Excel Data") as t:
            t.Start()
            
            # Clean existing
            collector = DB.FilteredElementCollector(fam_doc, ref_view.Id)
            ids_to_del = System.Collections.Generic.List[DB.ElementId]()
            for el in collector.ToElements():
                if isinstance(el, (DB.DetailCurve, DB.TextNote, DB.FilledRegion)):
                    ids_to_del.Add(el.Id)
            if ids_to_del.Count > 0:
                fam_doc.Delete(ids_to_del)

            # --- Optimization: Cache Types ---
            def get_type_name(e):
                p = e.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                return p.AsString() if p else ""

            text_type_cache = {get_type_name(t): t for t in DB.FilteredElementCollector(fam_doc).OfClass(DB.TextNoteType).ToElements()}
            fill_type_cache = {get_type_name(t): t for t in DB.FilteredElementCollector(fam_doc).OfClass(DB.FilledRegionType).ToElements()}
            
            solid_pat_id = None
            sp = next((p for p in DB.FilteredElementCollector(fam_doc).OfClass(DB.FillPatternElement).ToElements() if p.GetFillPattern().IsSolidFill), None)
            if sp: solid_pat_id = sp.Id

            # Ensure Styles
            line_style_map = FamilyGenerator.ensure_line_styles(fam_doc)

            # Pre-calc dimensions
            row_y = {}
            col_x = {}
            
            curr_y = 0.0
            sorted_rows = sorted([int(k) for k in data['row_heights'].keys()])
            for r in sorted_rows:
                row_y[r] = curr_y
                rh = data['row_heights'].get(str(r), 12.75)
                curr_y -= max(FamilyGenerator.points_to_feet(rh), 0.005) # Min height
            
            curr_x = 0.0
            sorted_cols = sorted([int(k) for k in data['column_widths'].keys()])
            for c in sorted_cols:
                col_x[c] = curr_x
                cw_px = data['column_widths'].get(str(c), 64.0)
                curr_x += max(FamilyGenerator.pixels_to_feet(cw_px), 0.005) # Min width

            # --- Geometry Optimizer (Style Aware) ---
            class StyleAwareOptimizer:
                def __init__(self):
                    # Key: Coordinate (Y or X), Value: List of (start, end, weight)
                    self.horiz_lines = {}
                    self.vert_lines = {}
                    # Fills: key=ColorStr, val=[(x, y, w, h)]
                    self.fills = {}

                def add_line(self, p1, p2, weight=1):
                    # Normalize checks
                    if abs(p1.Y - p2.Y) < 0.0001: # Horizontal
                        y = round(p1.Y, 5)
                        xs, xe = sorted([p1.X, p2.X])
                        if y not in self.horiz_lines: self.horiz_lines[y] = []
                        self.horiz_lines[y].append((xs, xe, weight))
                    elif abs(p1.X - p2.X) < 0.0001: # Vertical
                        x = round(p1.X, 5)
                        ys, ye = sorted([p1.Y, p2.Y])
                        if x not in self.vert_lines: self.vert_lines[x] = []
                        self.vert_lines[x].append((ys, ye, weight))

                def add_fill(self, color_str, x, y, w, h):
                    if not color_str: return
                    if color_str not in self.fills: self.fills[color_str] = []
                    self.fills[color_str].append((x, y, w, h))
                
                def resolve_intervals(self, segments):
                    """
                    Resolves overlapping segments using a priority (weight) system.
                    Input: List of (start, end, weight)
                    Output: List of (start, end, weight) with no overlaps and merged neighbors.
                    """
                    if not segments: return []
                    
                    # 1. Collect all unique split points
                    points = set()
                    for s, e, w in segments:
                        points.add(s)
                        points.add(e)
                    sorted_points = sorted(list(points))
                    
                    if len(sorted_points) < 2: return []
                    
                    # 2. Create atomic intervals and find max weight for each
                    resolved = []
                    for i in range(len(sorted_points) - 1):
                        p1 = sorted_points[i]
                        p2 = sorted_points[i+1]
                        mid = (p1 + p2) / 2.0
                        
                        max_w = 0 # 0 means no line
                        for s, e, w in segments:
                            if s <= mid and e >= mid:
                                if w > max_w: max_w = w
                        
                        if max_w > 0:
                            resolved.append({'s': p1, 'e': p2, 'w': max_w})
                            
                    # 3. Merge adjacent intervals with same weight
                    if not resolved: return []
                    
                    merged = []
                    curr = resolved[0]
                    
                    for i in range(1, len(resolved)):
                        next_int = resolved[i]
                        # Check continuity and same weight
                        if abs(curr['e'] - next_int['s']) < 0.0001 and curr['w'] == next_int['w']:
                            curr['e'] = next_int['e'] # Extend
                        else:
                            merged.append(curr)
                            curr = next_int
                    merged.append(curr)
                    
                    return [(m['s'], m['e'], m['w']) for m in merged]

                def get_line_style_id(self, doc, weight, cache):
                    # Map weight to subcategory ID
                    if weight >= 3: return cache.get("ODI_Wide", DB.ElementId.InvalidElementId)
                    elif weight == 2: return cache.get("ODI_Medium", DB.ElementId.InvalidElementId)
                    else: return cache.get("ODI_Thin", DB.ElementId.InvalidElementId)

                def draw_lines(self, doc, view, style_cache):
                    # Draw Horizontal
                    for y, segments in self.horiz_lines.items():
                        final_segments = self.resolve_intervals(segments)
                        for s, e, w in final_segments:
                            if abs(e - s) > 0.001:
                                l = DB.Line.CreateBound(DB.XYZ(s, y, 0), DB.XYZ(e, y, 0))
                                try:
                                    dc = doc.FamilyCreate.NewDetailCurve(view, l)
                                    gs_id = self.get_line_style_id(doc, w, style_cache)
                                    if gs_id != DB.ElementId.InvalidElementId:
                                        dc.LineStyle = doc.GetElement(gs_id)
                                except: pass

                    # Draw Vertical
                    for x, segments in self.vert_lines.items():
                        final_segments = self.resolve_intervals(segments)
                        for s, e, w in final_segments:
                            if abs(e - s) > 0.001:
                                l = DB.Line.CreateBound(DB.XYZ(x, s, 0), DB.XYZ(x, e, 0))
                                try:
                                    dc = doc.FamilyCreate.NewDetailCurve(view, l)
                                    gs_id = self.get_line_style_id(doc, w, style_cache)
                                    if gs_id != DB.ElementId.InvalidElementId:
                                        dc.LineStyle = doc.GetElement(gs_id)
                                except: pass

                def draw_fills(self, doc, view, type_cache, solid_id):
                    for color_str, rects in self.fills.items():
                        color = FamilyGenerator.get_color_from_string(color_str)
                        if not color: continue
                        
                        # Improved FillType Creation
                        name = "Solid_{}_{}_{}".format(color.Red, color.Green, color.Blue)
                        ftype = None
                        if name in type_cache:
                            ftype = type_cache[name]
                        else:
                             # Create New
                             base = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()
                             if base:
                                 try: ftype = base.Duplicate(name)
                                 except: 
                                     # Try finding again
                                     for x in DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).ToElements():
                                         if x.Name == name: ftype = x; break
                                 
                                 if ftype:
                                     # Modern API (2019+)
                                     try: ftype.ForegroundPatternColor = color
                                     except: 
                                         try: ftype.Color = color
                                         except: pass
                                     
                                     if solid_id:
                                         try: ftype.ForegroundPatternId = solid_id
                                         except: 
                                              try: ftype.FillPatternId = solid_id
                                              except: pass
                                     
                                     type_cache[name] = ftype
                        
                        if not ftype: continue

                        # Batch loops. Limit to 500 per region to be safe
                        batch_size = 500
                        chunks = [rects[i:i + batch_size] for i in range(0, len(rects), batch_size)]
                        
                        for chunk in chunks:
                            loops = System.Collections.Generic.List[DB.CurveLoop]()
                            for (x, y, w, h) in chunk:
                                lines = [
                                    DB.Line.CreateBound(DB.XYZ(x, y, 0), DB.XYZ(x+w, y, 0)),
                                    DB.Line.CreateBound(DB.XYZ(x+w, y, 0), DB.XYZ(x+w, y-h, 0)),
                                    DB.Line.CreateBound(DB.XYZ(x+w, y-h, 0), DB.XYZ(x, y-h, 0)),
                                    DB.Line.CreateBound(DB.XYZ(x, y-h, 0), DB.XYZ(x, y, 0))
                                ]
                                try:
                                    loop = DB.CurveLoop.Create(lines)
                                    loops.Add(loop)
                                except: pass
                            
                            if loops.Count > 0:
                                try: DB.FilledRegion.Create(doc, ftype.Id, view.Id, loops)
                                except: pass

            optimizer = StyleAwareOptimizer()

            # Border Weight Mapping (NPOI Strings -> Int Priority)
            BORDER_WEIGHTS = {
                "NONE": 0,
                "THIN": 1,
                "HAIR": 1,
                "DOTTED": 1,
                "DASHED": 1,
                "MEDIUM": 2,
                "MEDIUM_DASHED": 2,
                "THICK": 3,
                "DOUBLE": 3
            }

            # Iterate Cells
            for cell in data['cells']:
                r, c = cell['row'], cell['col']
                
                if r not in row_y or c not in col_x: continue
                
                # Dimensions with Span
                x = col_x[c]
                y = row_y[r]
                
                # Check spans
                r_span = 1
                c_span = 1
                merge_key = "{},{}".format(r, c)
                if merge_key in data['merges']:
                    r_span = data['merges'][merge_key]['r_span']
                    c_span = data['merges'][merge_key]['c_span']
                
                # Calculate W/H based on span
                w = 0.0
                for cs in range(c_span):
                    target_c = c + cs
                    if target_c in col_x:
                        cw_px = data['column_widths'].get(str(target_c), 64.0)
                        w += max(FamilyGenerator.pixels_to_feet(cw_px), 0.005)
                        
                h = 0.0
                for rs in range(r_span):
                    target_r = r + rs
                    if target_r in row_y:
                        rh = data['row_heights'].get(str(target_r), 12.75)
                        h += max(FamilyGenerator.points_to_feet(rh), 0.005)
                
                # Register Fill
                fill_color = cell.get('fill', {}).get('color')
                if fill_color:
                    optimizer.add_fill(fill_color, x, y, w, h)

                # Register Borders
                borders = cell.get('borders', {})
                
                def get_w(b_val):
                    return BORDER_WEIGHTS.get(str(b_val).upper(), 1) if b_val and str(b_val) != "0" else 0

                wt = get_w(borders.get('top'))
                if wt > 0: optimizer.add_line(DB.XYZ(x, y, 0), DB.XYZ(x+w, y, 0), wt)
                
                wb = get_w(borders.get('bottom'))
                if wb > 0: optimizer.add_line(DB.XYZ(x, y-h, 0), DB.XYZ(x+w, y-h, 0), wb)
                
                wl = get_w(borders.get('left'))
                if wl > 0: optimizer.add_line(DB.XYZ(x, y, 0), DB.XYZ(x, y-h, 0), wl)
                
                wr = get_w(borders.get('right'))
                if wr > 0: optimizer.add_line(DB.XYZ(x+w, y, 0), DB.XYZ(x+w, y-h, 0), wr)

                # Draw Text (Keep as is, but improved checking)
                val = cell['value']
                if val:
                    scaled_font = cell['font'].copy()
                    scaled_font['size'] = scaled_font['size'] * text_scale
                    
                    text_type = FamilyGenerator.get_or_create_text_type(fam_doc, scaled_font, text_type_cache)
                    
                    align_h = cell.get('align', 'Left')
                    align_v = cell.get('v_align', 'Bottom')
                    
                    revit_h_align = DB.HorizontalTextAlignment.Left
                    ins_x = x + 0.002
                    
                    if 'Center' in align_h:
                        revit_h_align = DB.HorizontalTextAlignment.Center
                        ins_x = x + w / 2.0
                    elif 'Right' in align_h:
                        revit_h_align = DB.HorizontalTextAlignment.Right
                        ins_x = x + w - 0.002
                    
                    revit_v_align = DB.VerticalTextAlignment.Bottom
                    ins_y = y - h + 0.002
                    
                    if 'Center' in align_v:
                        revit_v_align = DB.VerticalTextAlignment.Middle
                        ins_y = y - h / 2.0
                    elif 'Top' in align_v:
                        revit_v_align = DB.VerticalTextAlignment.Top
                        ins_y = y - 0.002
                        
                    insertion_point = DB.XYZ(ins_x, ins_y, 0)

                    try:
                        tn = DB.TextNote.Create(fam_doc, ref_view.Id, insertion_point, val, text_type.Id)
                        tn.HorizontalAlignment = revit_h_align
                        tn.VerticalAlignment = revit_v_align
                        
                        should_wrap = cell.get('wrap_text', False)
                        if should_wrap:
                            tn.Width = max(w - 0.004, 0.001) 
                    except: pass
            
            # Execute Optimization
            optimizer.draw_fills(fam_doc, ref_view, fill_type_cache, solid_pat_id)
            optimizer.draw_lines(fam_doc, ref_view, line_style_map)

            t.Commit()
        return True

    @staticmethod
    def load_family(fam_doc):
        class FamLoadOpt(DB.IFamilyLoadOptions):
            def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                overwriteParameterValues.Value = True
                return True
            def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
                source.Value = DB.FamilySource.Family
                overwriteParameterValues.Value = True
                return True
        
        try:
            return fam_doc.LoadFamily(revit.doc)
        except:
            try:
                return fam_doc.LoadFamily(revit.doc, FamLoadOpt())
            except Exception as e:
                print("Error Loading Family: " + str(e))
                return None

# --- Main App Logic ---
def perform_family_update(family, data, text_scale=1.0):
    fam_doc = revit.doc.EditFamily(family)
    success = FamilyGenerator.draw_content(fam_doc, data, text_scale)
    
    result = None
    if success:
        result = FamilyGenerator.load_family(fam_doc)
        
    fam_doc.Close(False)
    return result

def create_table(file_path, sheet_name, range_name, table_name, text_scale_percent=100.0):
    # Check Exists
    existing_fam = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements() if f.Name == table_name), None)
    
    data = excelextract.get_excel_data(file_path, sheet_name, range_name)
    if not data: return
    
    scale_factor = text_scale_percent / 100.0

    if existing_fam:
        res = forms.alert("Table '{}' already exists.\nUpdate it with new data?".format(table_name), options=["Update", "Cancel"])
        if res == "Update":
            updated_fam = perform_family_update(existing_fam, data, scale_factor)
            if updated_fam:
                TableMetadataManager.save(updated_fam, file_path, sheet_name, range_name, text_scale_percent)
                print("Table updated successfully.")
        return

    # Create New
    template = FamilyGenerator.get_template_path()
    if not template: return

    app = revit.doc.Application
    fam_doc = app.NewFamilyDocument(template)
    if not fam_doc: return
    
    success = FamilyGenerator.draw_content(fam_doc, data, scale_factor)
    
    if success:
        fam = FamilyGenerator.load_family(fam_doc)
        fam_doc.Close(False)
        
        if fam:
            # Rename
            try:
                with revit.Transaction("Rename Table Family"):
                    fam.Name = table_name
                    
                    # Rename Type to match Family Name
                    symbol_ids = fam.GetFamilySymbolIds()
                    if symbol_ids:
                        for sid in symbol_ids:
                            symbol = revit.doc.GetElement(sid)
                            if symbol:
                                symbol.Name = table_name
                            break # Rename the first/default type only
                        
            except Exception as e_rename:
                print("Warning: Rename failed. Family is named '{}'. Error: {}".format(fam.Name, str(e_rename)))
            
            TableMetadataManager.save(fam, file_path, sheet_name, range_name, text_scale_percent)
            
            forms.alert("Excel Table '{}' loaded successfully.".format(fam.Name))
    else:
        fam_doc.Close(False)

def update_table(instance):
    fam = instance.Symbol.Family
    meta = TableMetadataManager.get(fam) or TableMetadataManager.get(instance)
    
    if not meta:
        forms.alert("Not a linked Excel Table.")
        return
    
    if not os.path.exists(meta["SourcePath"]):
        forms.alert("Source file missing: " + meta["SourcePath"])
        return
        
    data = excelextract.get_excel_data(meta["SourcePath"], meta["SheetName"], meta["RangeName"])
    if not data: return
    
    scale_pct = meta.get("TextScale", 100.0)
    scale_factor = scale_pct / 100.0
    
    if perform_family_update(fam, data, scale_factor):
        print("Table updated.")

# --- UI ---
class NamedRangeItem(object):
    def __init__(self, data):
        self.name = data['name']
        self.sheet = data['sheet'] if data['sheet'] else "Workbook"
        self.formula = data['formula']

class ExcelScheduleWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.excel_path = None
        self.Loaded += self.window_loaded

    def window_loaded(self, sender, args):
        self.Tb_Scale.Text = "80"
        self.Btn_Browse_Click(None, None)

    def Btn_Browse_Click(self, sender, args):
        f = forms.pick_file(file_ext='xlsx')
        if f:
            self.excel_path = f
            self.Tb_FilePath.Text = f
            self.refresh_data()

    def refresh_data(self):
        if not self.excel_path: return
        self.Lb_Status.Text = "Reading..."
        self.Cb_Sheets.ItemsSource = excelextract.get_sheet_names(self.excel_path)
        if self.Cb_Sheets.ItemsSource: self.Cb_Sheets.SelectedIndex = 0
        
        raw_ranges = excelextract.get_print_areas(self.excel_path)
        self.Cb_Ranges.ItemsSource = [NamedRangeItem(r) for r in raw_ranges]
        if self.Cb_Ranges.ItemsSource: self.Cb_Ranges.SelectedIndex = 0
        self.Lb_Status.Text = "Ready"

    def check_existing_scale(self):
        try:
            t_name = self.Tb_Name.Text
            if not t_name: return
            
            # Find existing family
            fam = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements() if f.Name == t_name), None)
            if fam:
                meta = TableMetadataManager.get(fam)
                if meta and meta.get("TextScale"):
                    self.Tb_Scale.Text = str(meta["TextScale"])
        except: pass

    def Cb_Sheets_SelectionChanged(self, sender, args):
        if self.Cb_Sheets.SelectedItem:
            self.Tb_Name.Text = "{}{}_Table".format(FAMILY_PREFIX, self.Cb_Sheets.SelectedItem)
            self.check_existing_scale()

    def Cb_Ranges_SelectionChanged(self, sender, args):
        sel = self.Cb_Ranges.SelectedItem
        if sel:
            clean = sel.name.replace("_", " ") if sel.name != "Print_Area" else "{} Print Area".format(sel.sheet)
            self.Tb_Name.Text = "{}{}".format(FAMILY_PREFIX, clean)
            self.check_existing_scale()

    def Btn_Import_Click(self, sender, args):
        if not self.excel_path: return
        s = self.Cb_Sheets.SelectedItem
        r = self.Cb_Ranges.SelectedItem.name if self.Cb_Ranges.SelectedItem else None
        n = self.Tb_Name.Text or FAMILY_PREFIX + "Table"
        
        scale_pct = 80.0
        try:
            scale_pct = float(self.Tb_Scale.Text)
        except: pass
        
        self.Close()
        create_table(self.excel_path, s, r, n, scale_pct)

    def Btn_Close_Click(self, sender, args):
        self.Close()

# --- Entry ---
sel = revit.get_selection()
if len(sel) == 1 and isinstance(sel[0], DB.FamilyInstance):
    fam = sel[0].Symbol.Family
    meta = TableMetadataManager.get(fam) or TableMetadataManager.get(sel[0])
    if meta and meta["SourcePath"]:
        if forms.alert("Update this Excel Table?", options=["Update", "Create New..."]) == "Update":
            update_table(sel[0])
            script.exit()

ExcelScheduleWindow().ShowDialog()