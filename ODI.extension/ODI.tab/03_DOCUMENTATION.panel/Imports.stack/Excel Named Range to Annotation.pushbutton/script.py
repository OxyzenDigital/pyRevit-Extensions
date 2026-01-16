# -*- coding: utf-8 -*-
__title__ = "Excel Named \nRange to Annotation"
__version__ = "1.0.0"
__doc__ = """Imports Excel ranges as Revit Annotation Families.
Features:
- Preserves Formatting (Fonts, Colors, Borders)
- One-click Updates
- Persistent Settings"""

import os
import sys
import json
import clr
import System
import tempfile
import shutil
from pyrevit import revit, DB, forms, script
import excelextract

# --- Debug Helper ---
def log(msg):
    print("[Debug] " + str(msg))

# --- Constants ---
FAMILY_PREFIX = "ODI_Excel_"

# --- Load Options ---
class FamLoadOpt(DB.IFamilyLoadOptions, System.Object):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        if overwriteParameterValues is None:
            log("FamLoadOpt: overwriteParameterValues is None")
            return False
        try:
            overwriteParameterValues.Value = True
            return True
        except Exception as e:
            log("FamLoadOpt Error: " + str(e))
            return False

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        if source is None or overwriteParameterValues is None:
            log("FamLoadOpt Shared: source or overwriteParameterValues is None")
            return False
        try:
            source.Value = DB.FamilySource.Family
            overwriteParameterValues.Value = True
            return True
        except Exception as e:
            log("FamLoadOpt Shared Error: " + str(e))
            return False

# --- Warning Swallower ---
class WarningSwallower(DB.IFailuresPreprocessor, System.Object):
    def PreprocessFailures(self, failuresAccessor):
        if failuresAccessor is None:
            return DB.FailureProcessingResult.Continue
        try:
            failures = failuresAccessor.GetFailureMessages()
            if not failures: return DB.FailureProcessingResult.Continue
            
            for f in failures:
                if f.GetSeverity() == DB.FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(f)
            return DB.FailureProcessingResult.Continue
        except Exception as e:
            log("WarningSwallower Error: " + str(e))
            return DB.FailureProcessingResult.Continue

# --- Local Settings Manager (JSON) ---
SETTINGS_DIR = os.path.join(os.getenv('APPDATA'), 'ODI_ExcelTable')
if not os.path.exists(SETTINGS_DIR):
    try: os.makedirs(SETTINGS_DIR)
    except: pass
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "excel_table_map.json")

class TableDataManager(object):
    @staticmethod
    def load_db():
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    return json.load(f)
            except: pass
        return {}

    @staticmethod
    def save_db(db):
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(db, f, indent=2)
        except Exception as e:
            log("Error saving DB: " + str(e))

    @staticmethod
    def save_metadata(family, source_path, sheet_name, range_name, text_scale):
        fam_name = family.Name if hasattr(family, "Name") else str(family)
        db = TableDataManager.load_db()
        db[fam_name] = {
            "SourcePath": source_path,
            "SheetName": sheet_name,
            "RangeName": range_name,
            "TextScale": text_scale
        }
        TableDataManager.save_db(db)

    @staticmethod
    def get_metadata(family_or_name):
        fam_name = family_or_name
        if hasattr(family_or_name, "Name"):
            fam_name = family_or_name.Name
        db = TableDataManager.load_db()
        return db.get(fam_name)

# --- Family Generation Logic ---
class FamilyGenerator(object):
    @staticmethod
    def pixels_to_feet(px):
        if px <= 0: return 0.001
        return px * (1.0 / 96.0) * (1.0 / 12.0)

    @staticmethod
    def points_to_feet(pts):
        return pts * (1.0 / 72.0) * (1.0 / 12.0)

    @staticmethod
    def text_points_to_feet(pts):
        val = pts * (1.0 / 96.0) * (1.0 / 12.0)
        return max(val, 0.0005) # Ensure non-zero positive value

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
    def safe_name(element):
        try: return element.Name
        except: return ""

    @staticmethod
    def get_or_create_fill_type(doc, color, type_cache, solid_pat_id):
        # Revit treats pure white (255,255,255) as black/inverse. 
        # Force it to near-white (255,255,254) to ensure it stays white.
        if int(color.Red) == 255 and int(color.Green) == 255 and int(color.Blue) == 255:
            color = DB.Color(255, 255, 254)

        name = "Solid_{}_{}_{}".format(color.Red, color.Green, color.Blue)
        
        # 1. Check Cache / Existing
        if name in type_cache:
            existing = type_cache[name]
            # Verify properties (If it's the same, use it)
            try:
                # Check Color
                e_color = existing.ForegroundPatternColor
                if not e_color.IsValid: e_color = existing.Color
                color_match = (e_color.Red == color.Red and e_color.Green == color.Green and e_color.Blue == color.Blue)
                
                # Check Pattern
                e_pat_id = existing.ForegroundPatternId
                if e_pat_id == DB.ElementId.InvalidElementId: e_pat_id = existing.FillPatternId
                pat_match = (not solid_pat_id) or (e_pat_id == solid_pat_id)
                
                if color_match and pat_match:
                    return existing
            except: pass
            # If we are here, name exists but properties differ. We need a new unique name.
            name = "{}_{}".format(name, str(System.Guid.NewGuid())[:8])

        base = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()
        if not base: return None
        
        new_type = None
        try: new_type = base.Duplicate(name)
        except Exception as e:
            # Try unique name if duplicate failed
            try:
                unique_name = "{}_{}".format(name, str(System.Guid.NewGuid())[:8])
                new_type = base.Duplicate(unique_name)
            except:
                log("Warning: Could not create FillType '{}'. Using base. Error: {}".format(name, e))
                return base
        
        # Set Color
        try: new_type.ForegroundPatternColor = color
        except: 
            try: new_type.Color = color
            except: pass
        
        # Set Pattern
        if solid_pat_id:
            try: new_type.ForegroundPatternId = solid_pat_id
            except: 
                try: new_type.FillPatternId = solid_pat_id
                except: pass
        
        # Cache using the requested name (or the unique one) to avoid re-creation
        type_cache["Solid_{}_{}_{}".format(color.Red, color.Green, color.Blue)] = new_type
        return new_type

    @staticmethod
    def get_or_create_text_type(doc, font_data, type_cache):
        c_data = font_data.get('color')
        # Handle tuple (rgb_str, transparency) or string
        rgb_str = c_data[0] if c_data and isinstance(c_data, tuple) else (c_data if c_data else "0,0,0")
        color = FamilyGenerator.get_color_from_string(rgb_str) or DB.Color(0,0,0)
        
        # Revit treats pure white (255,255,255) as black/inverse. 
        # Force it to near-white (255,255,254) to ensure it stays white.
        if int(color.Red) == 255 and int(color.Green) == 255 and int(color.Blue) == 255:
            color = DB.Color(255, 255, 254)

        name = "{}_{}_{}{}_{}_{}_{}".format(
            font_data.get('name', 'Arial'), 
            int(font_data.get('size', 10)),
            "B" if font_data.get('bold') else "",
            "I" if font_data.get('italic') else "",
            int(color.Red), int(color.Green), int(color.Blue)
        ).replace(" ", "")
        
        # 1. Check Cache / Existing
        if name in type_cache:
            existing = type_cache[name]
            # Verify properties
            try:
                matches = True
                # Font
                p_font = existing.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
                if p_font and p_font.AsString() != font_data.get('name', 'Arial'): matches = False
                
                # Size
                p_size = existing.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
                target_size = FamilyGenerator.text_points_to_feet(font_data.get('size', 10))
                if p_size and abs(p_size.AsDouble() - target_size) > 0.001: matches = False
                
                # Bold/Italic
                p_bold = existing.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_BOLD)
                if p_bold and p_bold.AsInteger() != (1 if font_data.get('bold') else 0): matches = False
                
                p_italic = existing.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_ITALIC)
                if p_italic and p_italic.AsInteger() != (1 if font_data.get('italic') else 0): matches = False
                
                # Color
                p_color = existing.get_Parameter(DB.BuiltInParameter.LINE_COLOR)
                target_int = color.Red + (color.Green << 8) + (color.Blue << 16)
                if p_color and p_color.AsInteger() != target_int: matches = False
                
                if matches: return existing
            except: pass
            # Properties differ, use unique name
            name = "{}_{}".format(name, str(System.Guid.NewGuid())[:8])
        
        base = DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).FirstElement()
        if not base: return None
        
        new_type = None
        try: new_type = base.Duplicate(name)
        except Exception as e:
            # Try unique name
            try:
                unique_name = "{}_{}".format(name, str(System.Guid.NewGuid())[:8])
                new_type = base.Duplicate(unique_name)
            except:
                log("Warning: Could not create TextType '{}'. Using base. Error: {}".format(name, e))
                return base
        
        # Explicitly set parameters
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_FONT).Set(font_data.get('name', 'Arial'))
        size_ft = FamilyGenerator.text_points_to_feet(font_data.get('size', 10))
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_SIZE).Set(size_ft)
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_BOLD).Set(1 if font_data.get('bold') else 0)
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_ITALIC).Set(1 if font_data.get('italic') else 0)
        new_type.get_Parameter(DB.BuiltInParameter.LINE_COLOR).Set(color.Red + (color.Green << 8) + (color.Blue << 16))
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE).Set(1.0)
        new_type.get_Parameter(DB.BuiltInParameter.TEXT_BACKGROUND).Set(1) 
        
        # Cache the result mapped to the *original requested name* logic
        # This ensures that next time we ask for "Arial_10...", we get this new valid type
        original_name_key = "{}_{}_{}{}_{}_{}_{}".format(
            font_data.get('name', 'Arial'), 
            int(font_data.get('size', 10)),
            "B" if font_data.get('bold') else "",
            "I" if font_data.get('italic') else "",
            int(color.Red), int(color.Green), int(color.Blue)
        ).replace(" ", "")
        type_cache[original_name_key] = new_type
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
            cat = None
            try: cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_GenericAnnotation)
            except: pass
            
            if not cat: 
                log("Warning: Could not find Generic Annotation category.")
                return {}
            
            styles = {
                "ODI_Wide": 5,
                "ODI_Medium": 3,
                "ODI_Thin": 1
            }
            
            cats_created = False
            for name, weight in styles.items():
                if not cat.SubCategories.Contains(name):
                    try:
                        doc.Settings.Categories.NewSubcategory(cat, name)
                        cats_created = True
                    except Exception as e:
                        log("Error creating subcategory '{}': {}".format(name, e))
            
            if cats_created:
                doc.Regenerate()

            for name, weight in styles.items():
                if cat.SubCategories.Contains(name):
                    subcat = cat.SubCategories.get_Item(name)
                    try: subcat.LineWeight = weight
                    except: pass
                    
                    gs_col = DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle).ToElements()
                    found_gs = next((g for g in gs_col if g.GraphicsStyleCategory and g.GraphicsStyleCategory.Id == subcat.Id), None)
                    
                    # Fallback: Try matching by name if ID match fails
                    if not found_gs:
                        found_gs = next((g for g in gs_col if FamilyGenerator.safe_name(g) == name), None)
                    
                    if found_gs:
                        style_map[name] = found_gs.Id
                    else:
                        log("Warning: GraphicsStyle not found for subcategory '{}'".format(name))
        except Exception as e:
            log("Error ensuring line styles: " + str(e))
        return style_map
    
    @staticmethod
    def get_invisible_style_id(doc):
        # 1. Try Generic Annotation Category (Specific to Annotation Families)
        try:
            cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_GenericAnnotation)
            if cat and cat.SubCategories.Contains("Invisible Lines"):
                return cat.SubCategories.get_Item("Invisible Lines").GetGraphicsStyle(DB.GraphicsStyleType.Projection).Id
        except: pass

        # 2. Try Lines Category (Standard)
        try:
            cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
            if cat and cat.SubCategories.Contains("Invisible Lines"):
                return cat.SubCategories.get_Item("Invisible Lines").GetGraphicsStyle(DB.GraphicsStyleType.Projection).Id
        except: pass

        # 3. Search all GraphicsStyles by name (Fallback)
        for gs in DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
            name = FamilyGenerator.safe_name(gs)
            if name == "<Invisible lines>" or name == "Invisible lines":
                return gs.Id

        return DB.ElementId.InvalidElementId

    # --- Task-Based Managers ---
    class FillManager(object):
        def __init__(self):
            self.fills = {} # Key: (rgb_str, transparency), Value: list of rects (x,y,w,h)

        def add(self, rgb_str, transparency, x, y, w, h):
            if not rgb_str: return
            key = (rgb_str, transparency)
            if key not in self.fills: self.fills[key] = []
            self.fills[key].append((x, y, w, h))

        def draw(self, doc, view, type_cache, solid_id, invisible_id):
            log("Task: Drawing Fills ({} groups)...".format(len(self.fills)))
            created_regions = []
            eps = 0.0005

            for (rgb_str, transparency), rects in self.fills.items():
                color = FamilyGenerator.get_color_from_string(rgb_str)
                if not color: continue
                
                ftype = FamilyGenerator.get_or_create_fill_type(doc, color, type_cache, solid_id)
                if not ftype: continue

                # 1. Merge Horizontally
                sorted_rects = sorted(rects, key=lambda r: (-round(r[1], 5), round(r[0], 5)))
                merged_h = []
                if sorted_rects:
                    curr_x, curr_y, curr_w, curr_h = sorted_rects[0]
                    for i in range(1, len(sorted_rects)):
                        nx, ny, nw, nh = sorted_rects[i]
                        if abs(ny - curr_y) < 0.0001 and abs(nh - curr_h) < 0.0001 and abs(nx - (curr_x + curr_w)) < 0.0001:
                            curr_w += nw
                        else:
                            merged_h.append((curr_x, curr_y, curr_w, curr_h))
                            curr_x, curr_y, curr_w, curr_h = nx, ny, nw, nh
                    merged_h.append((curr_x, curr_y, curr_w, curr_h))

                # 2. Merge Vertically
                sorted_v = sorted(merged_h, key=lambda r: (round(r[0], 5), -round(r[1], 5)))
                final_rects = []
                if sorted_v:
                    curr_x, curr_y, curr_w, curr_h = sorted_v[0]
                    for i in range(1, len(sorted_v)):
                        nx, ny, nw, nh = sorted_v[i]
                        if abs(nx - curr_x) < 0.0001 and abs(nw - curr_w) < 0.0001 and abs(ny - (curr_y - curr_h)) < 0.0001:
                            curr_h += nh
                        else:
                            final_rects.append((curr_x, curr_y, curr_w, curr_h))
                            curr_x, curr_y, curr_w, curr_h = nx, ny, nw, nh
                    final_rects.append((curr_x, curr_y, curr_w, curr_h))

                # 3. Create Regions
                batch_size = 500
                chunks = [final_rects[i:i + batch_size] for i in range(0, len(final_rects), batch_size)]
                
                for chunk in chunks:
                    loops = System.Collections.Generic.List[DB.CurveLoop]()
                    for (x, y, w, h) in chunk:
                        sx, sy = x + eps, y - eps
                        sw, sh = w - (2 * eps), h - (2 * eps)
                        if sw <= 0 or sh <= 0: continue

                        lines = System.Collections.Generic.List[DB.Curve]()
                        p0, p1 = DB.XYZ(sx, sy, 0), DB.XYZ(sx + sw, sy, 0)
                        p2, p3 = DB.XYZ(sx + sw, sy - sh, 0), DB.XYZ(sx, sy - sh, 0)
                        
                        lines.Add(DB.Line.CreateBound(p0, p1))
                        lines.Add(DB.Line.CreateBound(p1, p2))
                        lines.Add(DB.Line.CreateBound(p2, p3))
                        lines.Add(DB.Line.CreateBound(p3, p0))
                        try: loops.Add(DB.CurveLoop.Create(lines))
                        except: pass

                    if loops.Count > 0:
                        try:
                            fr = DB.FilledRegion.Create(doc, ftype.Id, view.Id, loops)
                            if transparency > 0:
                                p = fr.get_Parameter(DB.BuiltInParameter.TRANSPARENCY)
                                if p and not p.IsReadOnly: p.Set(transparency)
                            created_regions.append(fr)
                        except Exception as e: log("Fill Create Error: " + str(e))
            
            # 4. Apply Invisible Lines (Requires Regeneration first)
            if invisible_id != DB.ElementId.InvalidElementId and created_regions:
                log("Regenerating to access sketches...")
                doc.Regenerate()
                
                invisible_gs = doc.GetElement(invisible_id)
                if invisible_gs:
                    count_fixed = 0
                    for fr in created_regions:
                        try:
                            # Try to find sketch
                            sketch = None
                            if hasattr(fr, "SketchId"):
                                try:
                                    sid = fr.SketchId
                                    if sid != DB.ElementId.InvalidElementId:
                                        sketch = doc.GetElement(sid)
                                except: pass
                            
                            if not sketch:
                                # Fallback search
                                sketch = next((s for s in DB.FilteredElementCollector(doc).OfClass(DB.Sketch).ToElements() if s.OwnerId == fr.Id), None)
                            
                            if sketch:
                                sketch_lines = DB.FilteredElementCollector(doc).OfClass(DB.CurveElement).WherePasses(DB.ElementOwnerFilter(sketch.Id)).ToElements()
                                for line in sketch_lines:
                                    try: line.LineStyle = invisible_gs
                                    except: pass
                                count_fixed += 1
                        except: pass
                    log("Applied invisible lines to {} regions.".format(count_fixed))
                else:
                    log("Invisible GraphicsStyle not found from ID.")

    class BorderManager(object):
        def __init__(self):
            self.horiz = {}
            self.vert = {}
            self.styles = {
                "NONE": 0, "THIN": 1, "HAIR": 1, "DOTTED": 1, "DASHED": 1,
                "MEDIUM": 2, "MEDIUM_DASHED": 2, "THICK": 3, "DOUBLE": 3
            }

        def add(self, p1, p2, weight_key):
            weight = self.styles.get(str(weight_key).upper(), 1) if weight_key and str(weight_key) != "0" else 0
            if weight == 0: return

            if abs(p1.Y - p2.Y) < 0.0001:
                y = round(p1.Y, 5)
                xs, xe = sorted([p1.X, p2.X])
                if y not in self.horiz: self.horiz[y] = []
                self.horiz[y].append((xs, xe, weight))
            elif abs(p1.X - p2.X) < 0.0001:
                x = round(p1.X, 5)
                ys, ye = sorted([p1.Y, p2.Y])
                if x not in self.vert: self.vert[x] = []
                self.vert[x].append((ys, ye, weight))

        def _resolve(self, segments):
            if not segments: return []
            # Split segments into points and find max weight for each interval.
            # Round points to ensure continuity between adjacent cells.
            raw_points = [s for s,e,w in segments] + [e for s,e,w in segments]
            points = sorted(list(set([round(p, 5) for p in raw_points])))
            
            if len(points) < 2: return []
            
            resolved = []
            for i in range(len(points) - 1):
                p1, p2 = points[i], points[i+1]
                mid = (p1 + p2) / 2.0
                max_w = 0
                for s, e, w in segments:
                    if s <= mid and e >= mid:
                        max_w = max(max_w, w)
                if max_w > 0:
                    resolved.append({'s': p1, 'e': p2, 'w': max_w})
            
            # Merge adjacent equal weights
            if not resolved: return []
            merged = []
            curr = resolved[0]
            for i in range(1, len(resolved)):
                nxt = resolved[i]
                if abs(curr['e'] - nxt['s']) < 0.0001 and curr['w'] == nxt['w']:
                    curr['e'] = nxt['e']
                else:
                    merged.append(curr)
                    curr = nxt
            merged.append(curr)
            return merged

        def _get_style_id(self, weight, cache):
            if weight >= 3: return cache.get("ODI_Wide", DB.ElementId.InvalidElementId)
            elif weight == 2: return cache.get("ODI_Medium", DB.ElementId.InvalidElementId)
            else: return cache.get("ODI_Thin", DB.ElementId.InvalidElementId)

        def draw(self, doc, view, style_cache):
            log("Task: Drawing Borders...")
            cnt = 0
            # Horizontal
            for y, segs in self.horiz.items():
                for r in self._resolve(segs):
                    if abs(r['e'] - r['s']) > 0.0025:
                        l = DB.Line.CreateBound(DB.XYZ(r['s'], y, 0), DB.XYZ(r['e'], y, 0))
                        try:
                            dc = doc.FamilyCreate.NewDetailCurve(view, l)
                            sid = self._get_style_id(r['w'], style_cache)
                            if sid != DB.ElementId.InvalidElementId: dc.LineStyle = doc.GetElement(sid)
                            cnt += 1
                        except: pass
            # Vertical
            for x, segs in self.vert.items():
                for r in self._resolve(segs):
                    if abs(r['e'] - r['s']) > 0.0025:
                        l = DB.Line.CreateBound(DB.XYZ(x, r['s'], 0), DB.XYZ(x, r['e'], 0))
                        try:
                            dc = doc.FamilyCreate.NewDetailCurve(view, l)
                            sid = self._get_style_id(r['w'], style_cache)
                            if sid != DB.ElementId.InvalidElementId: dc.LineStyle = doc.GetElement(sid)
                            cnt += 1
                        except: pass
            log("Borders Drawn: {}".format(cnt))

    class TextManager(object):
        def __init__(self):
            self.texts = []

        def add(self, val, font_data, x, y, w, h, align_h, align_v, indent, wrap):
            if val:
                self.texts.append({
                    'val': val, 'font': font_data, 
                    'x': x, 'y': y, 'w': w, 'h': h,
                    'ah': align_h, 'av': align_v, 'indent': indent, 'wrap': wrap
                })

        def draw(self, doc, view, type_cache, scale):
            log("Task: Drawing Text ({} items)...".format(len(self.texts)))
            cnt = 0
            for t in self.texts:
                scaled_font = t['font'].copy()
                scaled_font['size'] = scaled_font.get('size', 10) * scale
                
                ttype = FamilyGenerator.get_or_create_text_type(doc, scaled_font, type_cache)
                if not ttype: continue

                # Alignment
                indent_level = t.get('indent', 0)
                indent_offset = indent_level * 0.009
                width_reduction = 0.0

                r_align_h = DB.HorizontalTextAlignment.Left
                ins_x = t['x'] + 0.002
                if 'Center' in t['ah']:
                    r_align_h = DB.HorizontalTextAlignment.Center
                    ins_x = t['x'] + t['w'] / 2.0
                elif 'Right' in t['ah']:
                    r_align_h = DB.HorizontalTextAlignment.Right
                    ins_x = t['x'] + t['w'] - 0.002 - indent_offset
                    width_reduction = indent_offset
                else:
                    ins_x = t['x'] + 0.002 + indent_offset
                    width_reduction = indent_offset
                
                r_align_v = DB.VerticalTextAlignment.Bottom
                ins_y = t['y'] - t['h'] + 0.002
                if 'Center' in t['av']:
                    r_align_v = DB.VerticalTextAlignment.Middle
                    ins_y = t['y'] - t['h'] / 2.0
                elif 'Top' in t['av']:
                    r_align_v = DB.VerticalTextAlignment.Top
                    ins_y = t['y'] - 0.002

                try:
                    tn = DB.TextNote.Create(doc, view.Id, DB.XYZ(ins_x, ins_y, 0), t['val'], ttype.Id)
                    tn.HorizontalAlignment = r_align_h
                    tn.VerticalAlignment = r_align_v
                    if t['wrap']:
                        tn.Width = max(t['w'] - 0.004 - width_reduction, 0.005) # Min width ~1.5mm
                    cnt += 1
                except Exception as e: log("Text Create Failed: " + str(e))
            log("Text Notes Created: {}".format(cnt))

    @staticmethod
    def draw_content(fam_doc, data, text_scale=1.0):
        log("Starting draw_content...")
        views = [v for v in DB.FilteredElementCollector(fam_doc).OfClass(DB.View).ToElements() if not v.IsTemplate and v.ViewType != DB.ViewType.ProjectBrowser]
        ref_view = next((v for v in views if v.Name == "Ref. Level"), views[0] if views else None)
        if not ref_view: 
            log("No valid view found in family.")
            return False

        with DB.Transaction(fam_doc, "Draw Excel Data") as t:
            t.Start()
            
            # Clean existing
            collector = DB.FilteredElementCollector(fam_doc, ref_view.Id)
            ids_to_del = System.Collections.Generic.List[DB.ElementId]()
            for el in collector.ToElements():
                if isinstance(el, (DB.DetailCurve, DB.TextNote, DB.FilledRegion)):
                    ids_to_del.Add(el.Id)
            if ids_to_del.Count > 0:
                try: fam_doc.Delete(ids_to_del)
                except: pass

            # Caches (Use safe_name to prevent crashes on invalid elements)
            text_type_cache = {FamilyGenerator.safe_name(t): t for t in DB.FilteredElementCollector(fam_doc).OfClass(DB.TextNoteType).ToElements()}
            fill_type_cache = {FamilyGenerator.safe_name(t): t for t in DB.FilteredElementCollector(fam_doc).OfClass(DB.FilledRegionType).ToElements()}
            
            # Solid Pattern
            solid_pat_id = None
            pats = DB.FilteredElementCollector(fam_doc).OfClass(DB.FillPatternElement).ToElements()
            sp = next((p for p in pats if p.GetFillPattern().IsSolidFill), None)
            if not sp:
                sp = next((p for p in pats if FamilyGenerator.safe_name(p) in ["<Solid>", "Solid Fill"]), None)
            if sp: solid_pat_id = sp.Id
            else: log("Warning: No Solid Fill Pattern found.")

            # Line Styles
            line_style_map = FamilyGenerator.ensure_line_styles(fam_doc)
            invisible_style_id = FamilyGenerator.get_invisible_style_id(fam_doc)
            log("Line Styles Mapped: {}".format(line_style_map.keys()))

            # --- Task Managers ---
            fill_mgr = FamilyGenerator.FillManager()
            border_mgr = FamilyGenerator.BorderManager()
            text_mgr = FamilyGenerator.TextManager()

            # --- 1. Coordinate Mapping ---
            row_y, col_x = {}, {}
            curr_y, curr_x = 0.0, 0.0
            
            for r in sorted([int(k) for k in data['row_heights'].keys()]):
                row_y[r] = curr_y
                rh = data['row_heights'].get(str(r), 12.75)
                curr_y -= max(FamilyGenerator.points_to_feet(rh), 0.005)
            
            for c in sorted([int(k) for k in data['column_widths'].keys()]):
                col_x[c] = curr_x
                cw_px = data['column_widths'].get(str(c), 64.0)
                curr_x += max(FamilyGenerator.pixels_to_feet(cw_px), 0.005)

            # --- 2. Data Collection ---
            for cell in data['cells']:
                r, c = cell['row'], cell['col']
                if r not in row_y or c not in col_x: continue
                
                x = col_x[c]
                y = row_y[r]
                
                r_span = cell.get('r_span', 1)
                c_span = cell.get('c_span', 1)
                
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
                
                # Fill
                rgb_info = cell.get('fill', {}).get('color')
                if rgb_info and rgb_info[0]:
                    fill_mgr.add(rgb_info[0], rgb_info[1], x, y, w, h)

                # Borders
                borders = cell.get('borders', {})
                border_mgr.add(DB.XYZ(x, y, 0), DB.XYZ(x+w, y, 0), borders.get('top'))
                border_mgr.add(DB.XYZ(x, y-h, 0), DB.XYZ(x+w, y-h, 0), borders.get('bottom'))
                border_mgr.add(DB.XYZ(x, y, 0), DB.XYZ(x, y-h, 0), borders.get('left'))
                border_mgr.add(DB.XYZ(x+w, y, 0), DB.XYZ(x+w, y-h, 0), borders.get('right'))

                # Text
                if cell['value']:
                    text_mgr.add(cell['value'], cell['font'], x, y, w, h, 
                                 cell.get('align', 'Left'), cell.get('v_align', 'Bottom'), 
                                 cell.get('indent', 0),
                                 cell.get('wrap_text', False))

            # --- 3. Execution Phase ---
            fill_mgr.draw(fam_doc, ref_view, fill_type_cache, solid_pat_id, invisible_style_id)
            border_mgr.draw(fam_doc, ref_view, line_style_map)
            text_mgr.draw(fam_doc, ref_view, text_type_cache, text_scale)

            log("Regenerating Family Document...")
            fam_doc.Regenerate()
            log("Committing Transaction...")
            if t.Commit() != DB.TransactionStatus.Committed:
                log("Draw Content Transaction Failed")
                return False
            log("Transaction Committed.")
        return True

    @staticmethod
    def load_family_to_project(fam_doc, target_name, existing_family_id=None):
        """
        Loads the family into the project.
        Prioritizes Memory Load to avoid 'DocumentSaving' event triggers (SaveAs) which cause crashes 
        with some external add-ins.
        """
        # Check if project is modifiable before starting
        if revit.doc.IsReadOnly:
            log("Error: Project document is Read-Only.")
            return None

        loaded_fam = None
        
        # --- Attempt 1: Memory Load (Primary) ---
        try:
            # Memory Load requires the document to NOT be modifiable (no open transaction).
            # It handles its own transaction internally.
            load_opt = FamLoadOpt()
            loaded_fam = fam_doc.LoadFamily(revit.doc, load_opt)
                
        except Exception as e:
            log("Memory Load Failed: " + str(e))
            loaded_fam = None

        if loaded_fam:
            try: fam_doc.Close(False)
            except: pass
            return loaded_fam

        # --- Attempt 2: Disk Load (Fallback) ---
        log("Attempting Disk Load Fallback...")
        app = revit.doc.Application
        temp_folder = None
        try:
            temp_folder = tempfile.mkdtemp()
            safe_name = "".join([c for c in target_name if c.isalnum() or c in (' ', '-', '_', '(', ')', '.')]).strip()
            if not safe_name: safe_name = "ExcelTable"
            if not safe_name.lower().endswith(".rfa"): safe_name += ".rfa"
            temp_path = os.path.join(temp_folder, safe_name)
            
            # Save Family
            save_opt = DB.SaveAsOptions()
            save_opt.OverwriteExistingFile = True
            fam_doc.SaveAs(temp_path, save_opt)
            
            # Close Memory Doc
            try: fam_doc.Close(False)
            except: pass
            
            # Load from Disk
            loaded_fam_ref = clr.Reference[DB.Family]()
            with DB.Transaction(revit.doc, "Load Excel Table (Disk)") as t:
                t.Start()
                
                swallower = WarningSwallower()
                fho = t.GetFailureHandlingOptions()
                fho.SetFailuresPreprocessor(swallower)
                t.SetFailureHandlingOptions(fho)
                
                load_opt = FamLoadOpt()
                load_success = revit.doc.LoadFamily(temp_path, load_opt, loaded_fam_ref)
                revit.doc.Regenerate()
                
                if t.Commit() != DB.TransactionStatus.Committed:
                    return None
                
                if loaded_fam_ref.Value:
                    loaded_fam = loaded_fam_ref.Value
                elif load_success:
                    if existing_family_id:
                        try: loaded_fam = revit.doc.GetElement(existing_family_id)
                        except: pass
                    if not loaded_fam:
                        loaded_fam = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements() if f.Name == target_name), None)

        except Exception as e:
            log("Disk Load Failed: " + str(e))
        finally:
            if temp_folder and os.path.exists(temp_folder):
                try: shutil.rmtree(temp_folder)
                except: pass
        
        return loaded_fam

# --- Main App Logic ---
def perform_family_update(family, data, text_scale=1.0):
    if not family.IsEditable:
        log("Cannot edit System Family: " + family.Name)
        return None
        
    fam_doc = None
    try:
        log("Editing Family: " + family.Name)
        fam_doc = revit.doc.EditFamily(family)
        
        log("Drawing Content...")
        success = FamilyGenerator.draw_content(fam_doc, data, text_scale)
        
        if success:
            # Force timestamp to ensure change detection
            try:
                with DB.Transaction(fam_doc, "Update Timestamp") as t_time:
                    t_time.Start()
                    p = fam_doc.OwnerFamily.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                    if p: p.Set("Updated: " + str(System.DateTime.Now))
                    t_time.Commit()
            except Exception as e:
                log("Timestamp update failed: " + str(e))

            return FamilyGenerator.load_family_to_project(fam_doc, family.Name, family.Id)
        else:
            fam_doc.Close(False)
            return None
    except Exception as e:
        log("Error during update: " + str(e))
        if fam_doc:
            try: fam_doc.Close(False)
            except: pass
        return None

def create_table(file_path, sheet_name, range_name, table_name, text_scale_percent=100.0):
    existing_fam = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements() if f.Name == table_name), None)
    
    data = excelextract.get_excel_data(file_path, sheet_name, range_name)
    if not data: return
    
    scale_factor = text_scale_percent / 100.0

    if existing_fam:
        res = forms.alert("Table '{}' already exists.\nUpdate it with new data?".format(table_name), options=["Update", "Cancel"])
        if res == "Update":
            # Deselect to prevent silent rejection during family update
            try:
                revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]())
                revit.uidoc.RefreshActiveView()
            except: pass

            with DB.TransactionGroup(revit.doc, "Update Excel Table") as tg:
                tg.Start()
                updated_fam = perform_family_update(existing_fam, data, scale_factor)
                if updated_fam:
                    TableDataManager.save_metadata(updated_fam, file_path, sheet_name, range_name, text_scale_percent)
                    log("Table updated successfully.")
                    tg.Assimilate()
                    try: revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]([updated_fam.Id]))
                    except: pass
                else:
                    tg.RollBack()
                    forms.alert("Failed to update table. Check log for details.")
        return

    template = FamilyGenerator.get_template_path()
    if not template: return

    app = revit.doc.Application
    fam_doc = app.NewFamilyDocument(template)
    if not fam_doc: return
    
    success = FamilyGenerator.draw_content(fam_doc, data, scale_factor)
    
    if success:
        # Load using the unified method (handles saving/naming/loading)
        fam = FamilyGenerator.load_family_to_project(fam_doc, table_name)
        
        if fam:
            # Ensure name matches (in case fallback memory load was used)
            if fam.Name != table_name:
                try:
                    with revit.Transaction("Rename Table Family"):
                        fam.Name = table_name
                        # Also rename the type/symbol if possible
                        for sid in fam.GetFamilySymbolIds():
                            s = revit.doc.GetElement(sid)
                            if s: s.Name = table_name
                            break
                except: pass
            
            TableDataManager.save_metadata(fam, file_path, sheet_name, range_name, text_scale_percent)
            try: revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]([fam.Id]))
            except: pass
        else:
            # If load failed, fam_doc might still be open if it wasn't closed by load_family_to_project
            try: fam_doc.Close(False)
            except: pass
    else:
        fam_doc.Close(False)

def update_table(instance):
    # Deselect to prevent silent rejection during family update
    try:
        revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]())
        revit.uidoc.RefreshActiveView()
    except: pass

    fam = instance.Symbol.Family
    meta = TableDataManager.get_metadata(fam)
    
    if not meta:
        forms.alert("Not a linked Excel Table (No local settings found).")
        return
    
    if not os.path.exists(meta["SourcePath"]):
        forms.alert("Source file missing: " + meta["SourcePath"])
        return
        
    data = excelextract.get_excel_data(meta["SourcePath"], meta["SheetName"], meta["RangeName"])
    if not data: 
        forms.alert("Failed to read Excel data. Check file path and permissions.")
        return
    
    scale_pct = meta.get("TextScale", 80.0)
    scale_factor = scale_pct / 100.0
    
    updated_fam = perform_family_update(fam, data, scale_factor)
    if updated_fam:
        TableDataManager.save_metadata(updated_fam, meta["SourcePath"], meta["SheetName"], meta["RangeName"], scale_pct)
        log("Table updated.")
        revit.uidoc.RefreshActiveView()
        try: revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]([instance.Id]))
        except: pass
    else:
        log("Update failed (Family object not returned).")
        forms.alert("Failed to update table family.")

# --- UI ---
class NamedRangeItem(object):
    def __init__(self, data):
        self.name = data['name']
        self.sheet = data['sheet'] if data['sheet'] else "Global"
        self.formula = data['formula']

class ExcelScheduleWindow(forms.WPFWindow):
    def __init__(self, metadata=None):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.excel_path = None
        self.initial_metadata = metadata
        self.Loaded += self.window_loaded

    def window_loaded(self, sender, args):
        if self.initial_metadata:
            self.load_initial_settings(self.initial_metadata)
        else:
            self.Tb_Scale.Text = "80"
            self.Btn_Browse_Click(None, None)

    def load_initial_settings(self, meta):
        path = meta.get("SourcePath")
        if path and os.path.exists(path):
            self.excel_path = path
            self.Tb_FilePath.Text = path
            self.Tb_Scale.Text = str(meta.get("TextScale", "80"))
            if meta.get("FamilyName"):
                self.Tb_Name.Text = meta["FamilyName"]
            self.refresh_data()
            
            target_sheet = meta.get("SheetName")
            if target_sheet and self.Cb_Sheets.ItemsSource:
                for item in self.Cb_Sheets.ItemsSource:
                    if item == target_sheet:
                        self.Cb_Sheets.SelectedItem = item
                        break
            
            target_range = meta.get("RangeName")
            if target_range and self.Cb_Ranges.ItemsSource:
                for item in self.Cb_Ranges.ItemsSource:
                    if item.name == target_range:
                        self.Cb_Ranges.SelectedItem = item
                        break
        else:
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
            meta = TableDataManager.get_metadata(t_name)
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
        try: scale_pct = float(self.Tb_Scale.Text)
        except: pass
        
        self.Close()
        create_table(self.excel_path, s, r, n, scale_pct)

    def Btn_Close_Click(self, sender, args):
        self.Close()

# --- Entry ---
sel = revit.get_selection()
initial_meta = None

if len(sel) == 1 and isinstance(sel[0], DB.FamilyInstance):
    fam = sel[0].Symbol.Family
    meta = TableDataManager.get_metadata(fam)
    if meta and meta["SourcePath"]:
        initial_meta = meta
        initial_meta["FamilyName"] = fam.Name

ExcelScheduleWindow(initial_meta).ShowDialog()