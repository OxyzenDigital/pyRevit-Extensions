# -*- coding: utf-8 -*-
__title__ = "Excel Named \nRange to Annotation"
__version__ = "1.1.0"
__doc__ = """Imports Excel ranges as Revit Annotation Families.
Features:
- Preserves Formatting (Fonts, Colors, Borders)
- One-click Updates & Batch Refresh
- ExtensibleStorage & Local Settings Persistence
- Interactive Instance Placement
- Dark & Light Theme Adaptive UI"""

import os
import sys
import json
import clr
import System
import tempfile
import shutil

from pyrevit import HOST_APP, revit, DB, forms, script
from Autodesk.Revit.DB.ExtensibleStorage import SchemaBuilder, Schema, Entity, DataStorage, AccessLevel
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows import SystemColors
from System.Windows.Media import BrushConverter

import data_model
import excelextract

# --- Debug Helper ---
def log(msg):
    print("[Debug] " + str(msg))

# --- Constants ---
FAMILY_PREFIX = "ODI_Excel_"
EXCEL_TABLE_SCHEMA_GUID = System.Guid("A4E8C912-5B3D-4E7A-9C1F-8D2E6F0A3B5C")

# --- Revit Theme Detection ---
def is_dark_theme():
    try:
        from Autodesk.Revit.UI import UIThemeManager, UITheme
        return UIThemeManager.CurrentTheme == UITheme.Dark
    except Exception:
        return False

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
            if not failures:
                return DB.FailureProcessingResult.Continue
            
            for f in failures:
                if f.GetSeverity() == DB.FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(f)
            return DB.FailureProcessingResult.Continue
        except Exception as e:
            log("WarningSwallower Error: " + str(e))
            return DB.FailureProcessingResult.Continue

# --- Settings & ExtensibleStorage Manager ---
SETTINGS_DIR = os.path.join(os.getenv('APPDATA'), 'ODI_ExcelTable')
if not os.path.exists(SETTINGS_DIR):
    try:
        os.makedirs(SETTINGS_DIR)
    except Exception:
        pass
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "excel_table_map.json")

class TableDataManager(object):
    META_KEY = "__meta__"

    @staticmethod
    def get_or_create_schema():
        try:
            schema = Schema.Lookup(EXCEL_TABLE_SCHEMA_GUID)
            if not schema:
                builder = SchemaBuilder(EXCEL_TABLE_SCHEMA_GUID)
                builder.SetReadAccessLevel(AccessLevel.Public)
                builder.SetWriteAccessLevel(AccessLevel.Public)
                builder.SetSchemaName("ODIExcelTableLinks")
                builder.AddSimpleField("TableLinksJson", clr.GetClrType(System.String))
                schema = builder.Finish()
            return schema
        except Exception as e:
            log("SchemaBuilder Error: " + str(e))
            return None

    @staticmethod
    def load_from_extensible_storage(doc=None):
        if not doc:
            try:
                doc = revit.doc
            except Exception:
                pass
        if not doc:
            return {}

        try:
            schema = TableDataManager.get_or_create_schema()
            if not schema:
                return {}
            collector = DB.FilteredElementCollector(doc).OfClass(DataStorage)
            for ds in collector:
                entity = ds.GetEntity(schema)
                if entity and entity.IsValid():
                    raw_json = entity.Get[System.String]("TableLinksJson")
                    if raw_json:
                        return json.loads(raw_json)
        except Exception as e:
            log("ExtensibleStorage load error: " + str(e))
        return {}

    @staticmethod
    def save_to_extensible_storage(doc, db):
        if not doc:
            try:
                doc = revit.doc
            except Exception:
                pass
        if not doc:
            return False

        try:
            schema = TableDataManager.get_or_create_schema()
            if not schema:
                return False

            raw_json = json.dumps(db)

            def _write():
                collector = DB.FilteredElementCollector(doc).OfClass(DataStorage)
                target_ds = None
                for ds in collector:
                    entity = ds.GetEntity(schema)
                    if entity and entity.IsValid():
                        target_ds = ds
                        break
                if not target_ds:
                    target_ds = DataStorage.Create(doc)
                entity = Entity(schema)
                entity.Set[System.String]("TableLinksJson", raw_json)
                target_ds.SetEntity(entity)

            if doc.IsModifiable:
                _write()
            else:
                with DB.Transaction(doc, "Save Excel Table Links") as t:
                    t.Start()
                    _write()
                    t.Commit()
            return True
        except Exception as e:
            log("ExtensibleStorage save error: " + str(e))
            return False

    @staticmethod
    def load_db(doc=None):
        db = {}
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    db = json.load(f)
            except Exception:
                db = {}

        # Overlay ExtensibleStorage (project ground truth)
        es_db = TableDataManager.load_from_extensible_storage(doc)
        if es_db:
            db.update(es_db)
        return db

    @staticmethod
    def save_db(db, doc=None):
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(db, f, indent=2)
        except Exception as e:
            log("Error saving DB: " + str(e))

        TableDataManager.save_to_extensible_storage(doc, db)

    @staticmethod
    def save_metadata(family, source_path, sheet_name, range_name, text_scale, doc=None):
        fam_name = family.Name if hasattr(family, "Name") else str(family)
        db = TableDataManager.load_db(doc)
        db[fam_name] = {
            "SourcePath": source_path,
            "SheetName": sheet_name,
            "RangeName": range_name,
            "TextScale": text_scale
        }

        # Maintain __meta__ history
        meta_rec = db.get(TableDataManager.META_KEY)
        if not isinstance(meta_rec, dict):
            meta_rec = {}
        if source_path:
            meta_rec["LastFilePath"] = source_path
            recent = meta_rec.get("RecentFiles", [])
            if not isinstance(recent, list):
                recent = []
            try:
                norm_source = os.path.normcase(os.path.abspath(source_path))
            except Exception:
                norm_source = str(source_path).lower()
            new_recent = [source_path]
            for p in recent:
                if p:
                    try:
                        norm_p = os.path.normcase(os.path.abspath(p))
                    except Exception:
                        norm_p = str(p).lower()
                    if norm_p != norm_source:
                        new_recent.append(p)
            meta_rec["RecentFiles"] = new_recent[:15]
            db[TableDataManager.META_KEY] = meta_rec

        TableDataManager.save_db(db, doc)

    @staticmethod
    def get_recent_files(doc=None):
        db = TableDataManager.load_db(doc)
        meta_rec = db.get(TableDataManager.META_KEY)
        if not isinstance(meta_rec, dict):
            meta_rec = {}

        raw_candidates = []
        last_p = meta_rec.get("LastFilePath")
        if last_p:
            raw_candidates.append(last_p)
        for p in meta_rec.get("RecentFiles", []):
            if p:
                raw_candidates.append(p)
        for k, v in db.items():
            if k != TableDataManager.META_KEY and isinstance(v, dict):
                sp = v.get("SourcePath")
                if sp:
                    raw_candidates.append(sp)

        seen = set()
        unique_paths = []
        for p in raw_candidates:
            if not p:
                continue
            try:
                norm_p = os.path.normcase(os.path.abspath(p))
            except Exception:
                norm_p = str(p).lower()
            if norm_p not in seen:
                seen.add(norm_p)
                unique_paths.append(p)

        existing = [p for p in unique_paths if os.path.exists(p)]
        missing = [p for p in unique_paths if not os.path.exists(p)]
        return existing + missing

    @staticmethod
    def get_last_file_path(doc=None):
        db = TableDataManager.load_db(doc)
        meta_rec = db.get(TableDataManager.META_KEY)
        if isinstance(meta_rec, dict):
            last_p = meta_rec.get("LastFilePath")
            if last_p:
                return last_p
        recent = TableDataManager.get_recent_files(doc)
        return recent[0] if recent else ""

    @staticmethod
    def infer_hint_metadata(fam_name, db, recent_files):
        def _norm(s):
            if not s:
                return ""
            val = str(s).lower()
            if val.startswith("odi_excel_"):
                val = val[len("odi_excel_"):]
            val = val.replace("table", "").replace("_", "").replace(" ", "").strip()
            return val

        target_norm = _norm(fam_name)

        # 1. Fuzzy match against existing DB records
        for k, v in db.items():
            if k == TableDataManager.META_KEY or not isinstance(v, dict):
                continue
            k_norm = _norm(k)
            r_norm = _norm(v.get("RangeName", ""))
            s_norm = _norm(v.get("SheetName", ""))
            if target_norm and (k_norm == target_norm or r_norm == target_norm or s_norm == target_norm):
                return {
                    "HintPath": v.get("SourcePath", ""),
                    "SheetName": v.get("SheetName", ""),
                    "RangeName": v.get("RangeName", ""),
                    "TextScale": v.get("TextScale", 80.0)
                }

        # 2. Most frequent SourcePath in DB
        path_counts = {}
        for k, v in db.items():
            if k != TableDataManager.META_KEY and isinstance(v, dict):
                sp = v.get("SourcePath")
                if sp:
                    path_counts[sp] = path_counts.get(sp, 0) + 1
        if path_counts:
            sorted_paths = sorted(path_counts.keys(), key=lambda p: (path_counts[p], 1 if os.path.exists(p) else 0), reverse=True)
            return {"HintPath": sorted_paths[0]}

        # 3. Fallback to first available recent file
        if recent_files:
            return {"HintPath": recent_files[0]}

        return {}

    @staticmethod
    def get_metadata(family_or_name, doc=None):
        fam_name = family_or_name
        if hasattr(family_or_name, "Name"):
            fam_name = family_or_name.Name
        if fam_name == TableDataManager.META_KEY:
            return None
        db = TableDataManager.load_db(doc)
        meta = db.get(fam_name)

        if not meta and hasattr(family_or_name, "get_Parameter"):
            try:
                p = family_or_name.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)
                if p and p.AsString() and p.AsString().startswith("ODI_EXCEL_META:"):
                    raw = p.AsString()[len("ODI_EXCEL_META:"):]
                    meta = json.loads(raw)
            except Exception:
                pass
        return meta

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
        
        # Cache the result mapped to the original requested name
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
        # 1. Try Revit Application FamilyTemplatePath setting
        try:
            fam_tpl_path = revit.doc.Application.FamilyTemplatePath
            if fam_tpl_path and os.path.exists(fam_tpl_path):
                # Standard subfolder
                p1 = os.path.join(fam_tpl_path, "Annotations", "Generic Annotation.rft")
                if os.path.exists(p1):
                    return p1
                # Root template folder
                p2 = os.path.join(fam_tpl_path, "Generic Annotation.rft")
                if os.path.exists(p2):
                    return p2
                # Shallow walk
                for root, dirs, files in os.walk(fam_tpl_path):
                    for f in files:
                        if f.lower() == "generic annotation.rft":
                            return os.path.join(root, f)
                    break
        except Exception as e:
            log("FamilyTemplatePath inspection error: " + str(e))

        # 2. Known Autodesk RVT version paths
        current_v = str(HOST_APP.version) if hasattr(HOST_APP, "version") else "2026"
        versions = [current_v, "2027", "2026", "2025", "2024", "2023", "2022", "2021", "2020"]
        # Deduplicate while preserving order
        seen = set()
        ordered_versions = []
        for v in versions:
            if v and v not in seen:
                seen.add(v)
                ordered_versions.append(v)

        base_templates = [
            r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English\Annotations",
            r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English-Imperial\Annotations",
            r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English_I\Annotations"
        ]

        for v in ordered_versions:
            for b in base_templates:
                p = os.path.join(b.format(v), "Generic Annotation.rft")
                if os.path.exists(p):
                    return p

        # 3. User prompt fallback
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
        try:
            cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_GenericAnnotation)
            if cat and cat.SubCategories.Contains("Invisible Lines"):
                return cat.SubCategories.get_Item("Invisible Lines").GetGraphicsStyle(DB.GraphicsStyleType.Projection).Id
        except: pass

        try:
            cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
            if cat and cat.SubCategories.Contains("Invisible Lines"):
                return cat.SubCategories.get_Item("Invisible Lines").GetGraphicsStyle(DB.GraphicsStyleType.Projection).Id
        except: pass

        for gs in DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
            name = FamilyGenerator.safe_name(gs)
            if name == "<Invisible lines>" or name == "Invisible lines":
                return gs.Id

        return DB.ElementId.InvalidElementId

    # --- Task-Based Managers ---
    class FillManager(object):
        def __init__(self):
            self.fills = {}

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

                # Merge Horizontally
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

                # Merge Vertically
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

                # Create Regions
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
            
            # Apply Invisible Lines
            if invisible_id != DB.ElementId.InvalidElementId and created_regions:
                log("Regenerating to access sketches...")
                doc.Regenerate()
                
                invisible_gs = doc.GetElement(invisible_id)
                if invisible_gs:
                    count_fixed = 0
                    for fr in created_regions:
                        try:
                            sketch = None
                            if hasattr(fr, "SketchId"):
                                try:
                                    sid = fr.SketchId
                                    if sid != DB.ElementId.InvalidElementId:
                                        sketch = doc.GetElement(sid)
                                except: pass
                            
                            if not sketch:
                                sketch = next((s for s in DB.FilteredElementCollector(doc).OfClass(DB.Sketch).ToElements() if s.OwnerId == fr.Id), None)
                            
                            if sketch:
                                sketch_lines = DB.FilteredElementCollector(doc).OfClass(DB.CurveElement).WherePasses(DB.ElementOwnerFilter(sketch.Id)).ToElements()
                                for line in sketch_lines:
                                    try: line.LineStyle = invisible_gs
                                    except: pass
                                count_fixed += 1
                        except: pass
                    log("Applied invisible lines to {} regions.".format(count_fixed))

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
                        tn.Width = max(t['w'] - 0.004 - width_reduction, 0.005)
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

            # Caches
            text_type_cache = {FamilyGenerator.safe_name(ty): ty for ty in DB.FilteredElementCollector(fam_doc).OfClass(DB.TextNoteType).ToElements()}
            fill_type_cache = {FamilyGenerator.safe_name(ty): ty for ty in DB.FilteredElementCollector(fam_doc).OfClass(DB.FilledRegionType).ToElements()}
            
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

            fill_mgr = FamilyGenerator.FillManager()
            border_mgr = FamilyGenerator.BorderManager()
            text_mgr = FamilyGenerator.TextManager()

            # 1. Coordinate Mapping
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

            # 2. Data Collection
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

            # 3. Execution Phase
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
        if revit.doc.IsReadOnly:
            log("Error: Project document is Read-Only.")
            return None

        loaded_fam = None
        
        # --- Attempt 1: Memory Load (Primary) ---
        try:
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
        temp_folder = None
        try:
            temp_folder = tempfile.mkdtemp()
            safe_name = "".join([c for c in target_name if c.isalnum() or c in (' ', '-', '_', '(', ')', '.')]).strip()
            if not safe_name: safe_name = "ExcelTable"
            if not safe_name.lower().endswith(".rfa"): safe_name += ".rfa"
            temp_path = os.path.join(temp_folder, safe_name)
            
            save_opt = DB.SaveAsOptions()
            save_opt.OverwriteExistingFile = True
            fam_doc.SaveAs(temp_path, save_opt)
            
            try: fam_doc.Close(False)
            except: pass
            
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

# --- Interactive Placement ---
def place_table_instance(family):
    if not family:
        return None
    doc = revit.doc
    uidoc = revit.uidoc
    symbol_ids = family.GetFamilySymbolIds()
    if not symbol_ids or len(list(symbol_ids)) == 0:
        log("No symbols found in family.")
        return None
    
    symbol_id = list(symbol_ids)[0]
    symbol = doc.GetElement(symbol_id)
    if not symbol:
        return None

    if not symbol.IsActive:
        try:
            with DB.Transaction(doc, "Activate Table Symbol") as t:
                t.Start()
                symbol.Activate()
                doc.Regenerate()
                t.Commit()
        except Exception as e:
            log("Symbol activation error: " + str(e))

    active_view = doc.ActiveView
    if active_view:
        try:
            uidoc.PromptForFamilyInstancePlacement(symbol)
        except OperationCanceledException:
            log("User canceled placement.")
        except Exception as e:
            log("Placement prompt error: " + str(e))

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
            try:
                with DB.Transaction(fam_doc, "Update Timestamp") as t_time:
                    t_time.Start()
                    p = fam_doc.OwnerFamily.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
                    if p and not p.IsReadOnly:
                        p.Set("Updated: " + str(System.DateTime.Now))
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

def create_table(file_path, sheet_name, range_name, table_name, text_scale_percent=100.0, place_instance=True):
    existing_fam = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements() if f.Name == table_name), None)
    
    clean_range = None if range_name == "<Entire Worksheet>" else range_name
    data = excelextract.get_excel_data(file_path, sheet_name, clean_range)
    if not data:
        forms.alert("Failed to extract data from Excel file.")
        return
    
    scale_factor = text_scale_percent / 100.0

    if existing_fam:
        res = forms.alert("Table '{}' already exists.\nUpdate it with new data?".format(table_name), options=["Update", "Cancel"])
        if res == "Update":
            try:
                revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]())
                revit.uidoc.RefreshActiveView()
            except: pass

            with DB.TransactionGroup(revit.doc, "Update Excel Table") as tg:
                tg.Start()
                updated_fam = perform_family_update(existing_fam, data, scale_factor)
                if updated_fam:
                    TableDataManager.save_metadata(updated_fam, file_path, sheet_name, range_name, text_scale_percent, revit.doc)
                    log("Table updated successfully.")
                    tg.Assimilate()
                    revit.uidoc.RefreshActiveView()
                    try:
                        revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]([updated_fam.Id]))
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
    
    # Embed compact metadata backup inside family
    try:
        with DB.Transaction(fam_doc, "Embed Metadata") as t_meta:
            t_meta.Start()
            p_c = fam_doc.OwnerFamily.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)
            if p_c and not p_c.IsReadOnly:
                meta_dict = {
                    "SourcePath": file_path,
                    "SheetName": sheet_name or "",
                    "RangeName": range_name or "",
                    "TextScale": text_scale_percent
                }
                p_c.Set("ODI_EXCEL_META:" + json.dumps(meta_dict))
            t_meta.Commit()
    except Exception as e:
        log("Metadata embedding warning: " + str(e))

    success = FamilyGenerator.draw_content(fam_doc, data, scale_factor)
    
    if success:
        fam = FamilyGenerator.load_family_to_project(fam_doc, table_name)
        if fam:
            if fam.Name != table_name:
                try:
                    with revit.Transaction("Rename Table Family"):
                        fam.Name = table_name
                        for sid in fam.GetFamilySymbolIds():
                            s = revit.doc.GetElement(sid)
                            if s: s.Name = table_name
                            break
                except: pass
            
            TableDataManager.save_metadata(fam, file_path, sheet_name, range_name, text_scale_percent, revit.doc)
            try:
                revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]([fam.Id]))
            except: pass

            if place_instance:
                place_table_instance(fam)
        else:
            try: fam_doc.Close(False)
            except: pass
    else:
        fam_doc.Close(False)

def update_table(family_or_instance, silent=False):
    try:
        revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]())
        revit.uidoc.RefreshActiveView()
    except: pass

    fam = None
    if isinstance(family_or_instance, DB.FamilyInstance):
        fam = family_or_instance.Symbol.Family
    elif isinstance(family_or_instance, DB.Family):
        fam = family_or_instance
    elif isinstance(family_or_instance, (str, System.String)):
        fam = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements() if f.Name == str(family_or_instance)), None)

    if not fam:
        if not silent: forms.alert("Could not locate family in project.")
        return False

    meta = TableDataManager.get_metadata(fam, revit.doc)
    if not meta:
        if not silent: forms.alert("Not a linked Excel Table (No settings found).")
        return False
    
    source_path = meta.get("SourcePath", "")
    if not source_path or not os.path.exists(source_path):
        if not silent: forms.alert("Source file missing: " + str(source_path))
        return False
        
    range_n = meta.get("RangeName")
    clean_range = None if range_n == "<Entire Worksheet>" else range_n
    data = excelextract.get_excel_data(source_path, meta.get("SheetName"), clean_range)
    if not data: 
        if not silent: forms.alert("Failed to read Excel data. Check file path and permissions.")
        return False
    
    scale_pct = meta.get("TextScale", 80.0)
    try:
        scale_pct = float(scale_pct)
    except:
        scale_pct = 80.0
    scale_factor = scale_pct / 100.0
    
    updated_fam = perform_family_update(fam, data, scale_factor)
    if updated_fam:
        TableDataManager.save_metadata(updated_fam, source_path, meta.get("SheetName"), meta.get("RangeName"), scale_pct, revit.doc)
        log("Table updated: " + fam.Name)
        revit.uidoc.RefreshActiveView()
        return True
    else:
        log("Update failed for: " + fam.Name)
        if not silent: forms.alert("Failed to update table family.")
        return False

def batch_update_tables(linked_items):
    if not linked_items:
        return

    try:
        revit.uidoc.Selection.SetElementIds(System.Collections.Generic.List[DB.ElementId]())
        revit.uidoc.RefreshActiveView()
    except Exception:
        pass

    updated_count = 0
    missing_count = 0
    failed_count = 0

    with DB.TransactionGroup(revit.doc, "Batch Update Excel Tables") as tg:
        tg.Start()
        for item in linked_items:
            if not item.FileExists:
                missing_count += 1
                continue
            res = update_table(item.FamilyName, silent=True)
            if res:
                updated_count += 1
            else:
                failed_count += 1
        tg.Assimilate()

    revit.uidoc.RefreshActiveView()
    msg = "Batch update completed.\n\nUpdated: {}\nMissing File: {}\nFailed: {}".format(
        updated_count, missing_count, failed_count
    )
    forms.alert(msg, title="Batch Update Results")

# --- Directory Helper ---
def get_best_init_dir(*candidate_paths):
    for p in candidate_paths:
        if not p:
            continue
        try:
            curr = p if os.path.isdir(p) else os.path.dirname(p)
            while curr and not os.path.exists(curr):
                parent = os.path.dirname(curr)
                if parent == curr:
                    curr = ""
                    break
                curr = parent
            if curr and os.path.exists(curr):
                return curr
        except Exception:
            pass
    return ""

# --- MVVM Window Controller ---
class ExcelScheduleWindow(forms.WPFWindow):
    def __init__(self, initial_family_name=None, metadata=None):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.vm = data_model.ExcelImportViewModel()
        self.DataContext = self.vm
        self.initial_family_name = initial_family_name
        self.initial_metadata = metadata
        self._suppress_recent_event = False

        self.apply_revit_theme()
        self.wire_events()

    def wire_events(self):
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_WinClose.Click += self.Btn_Close_Click
        self.Btn_Close.Click += self.Btn_Close_Click
        self.Btn_Browse.Click += self.Btn_Browse_Click
        self.Btn_Import.Click += self.Btn_Import_Click
        self.Btn_UpdateChecked.Click += self.Btn_UpdateChecked_Click
        self.Btn_PlaceSelected.Click += self.Btn_PlaceSelected_Click
        self.Btn_SelectAllLinked.Click += self.Btn_SelectAllLinked_Click
        self.Btn_SelectNoneLinked.Click += self.Btn_SelectNoneLinked_Click
        self.Cb_RecentFiles.SelectionChanged += self.Cb_RecentFiles_SelectionChanged
        self.Cb_Sheets.SelectionChanged += self.Cb_Sheets_SelectionChanged
        self.Cb_Ranges.SelectionChanged += self.Cb_Ranges_SelectionChanged
        self.Grid_LinkedTables.SelectionChanged += self.Grid_LinkedTables_SelectionChanged
        self.Loaded += self.window_loaded

    def apply_revit_theme(self):
        if is_dark_theme():
            dark_palette = {
                "WindowBrush": "#1F2937",
                "ToolbarBrush": "#1F2937",
                "ControlBrush": "#111827",
                "FooterBrush": "#111827",
                "AltRowBrush": "#1F2937",
                "TextBrush": "#F9FAFB",
                "TextLightBrush": "#9CA3AF",
                "HeaderTextBrush": "#F9FAFB",
                "BorderBrush": "#4B5563",
                "ButtonBrush": "#374151",
                "HoverBrush": "#374151",
                "AccentBrush": "#3B82F6",
                "SelectionBrush": "#1E3A8A",
                "SelectionBorderBrush": "#3B82F6",
                "SelectionTextBrush": "#FFFFFF",
                "InactiveSelectionBrush": "#374151",
                "CardBrush": "#374151",
                "CardBorderBrush": "#4B5563",
                "CardTextBrush": "#FFFFFF",
                "CardSubTextBrush": "#D1D5DB",
                "CardLabelBrush": "#9CA3AF",
                "CardValueBrush": "#FFFFFF",
                "CardAccentBrush": "#60A5FA",
                "PrimaryButtonTextBrush": "#FFFFFF"
            }
            conv = BrushConverter()
            for k, v in dark_palette.items():
                self.Resources[k] = conv.ConvertFromString(v)

            self.Resources[SystemColors.HighlightBrushKey] = conv.ConvertFromString("#1E3A8A")
            self.Resources[SystemColors.HighlightTextBrushKey] = conv.ConvertFromString("#FFFFFF")
            self.Resources[SystemColors.InactiveSelectionHighlightBrushKey] = conv.ConvertFromString("#374151")
            self.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = conv.ConvertFromString("#F9FAFB")

    def drag_window(self, sender, args):
        try:
            self.DragMove()
        except Exception:
            pass

    def window_loaded(self, sender, args):
        self.refresh_linked_tables(select_family_name=self.initial_family_name)

        if self.initial_metadata:
            self.load_metadata_into_editor(self.initial_metadata, family_name=self.initial_family_name)
        elif self.initial_family_name:
            meta = TableDataManager.get_metadata(self.initial_family_name, revit.doc)
            if meta:
                self.load_metadata_into_editor(meta, family_name=self.initial_family_name)
            else:
                db = TableDataManager.load_db(revit.doc)
                hint = TableDataManager.infer_hint_metadata(self.initial_family_name, db, self.vm.RecentFiles)
                self.load_metadata_into_editor(hint, family_name=self.initial_family_name, hint_path=hint.get("HintPath", ""))
        elif self.vm.LinkedTables:
            self.vm.SelectedLinkedTable = self.vm.LinkedTables[0]
            self.Grid_LinkedTables_SelectionChanged(None, None)
        elif self.vm.RecentFiles:
            first_recent = self.vm.RecentFiles[0]
            self._suppress_recent_event = True
            self.vm.SelectedRecentFile = first_recent
            self._suppress_recent_event = False
            if os.path.exists(first_recent):
                self.vm.FilePath = first_recent
                self.vm.LinkedFileHint = "Loaded last saved source: {}".format(first_recent)
                self.refresh_data()
            else:
                self.vm.LinkedFileHint = "Last saved source (missing): {}  (Click Browse to locate)".format(first_recent)
                self.vm.StatusText = "Missing File"
        else:
            self.Btn_Browse_Click(None, None)

    def refresh_linked_tables(self, select_family_name=None):
        all_families = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements()
        db = TableDataManager.load_db(revit.doc)
        recent = TableDataManager.get_recent_files(revit.doc)

        self._suppress_recent_event = True
        self.vm.RecentFiles = recent
        self._suppress_recent_event = False

        existing_names = set(f.Name for f in all_families if f.Name != TableDataManager.META_KEY)
        self.vm.existing_family_names = existing_names

        items = []
        for f in all_families:
            if f.Name == TableDataManager.META_KEY:
                continue
            meta = db.get(f.Name)
            if not meta and f.Name.startswith(FAMILY_PREFIX):
                meta = TableDataManager.get_metadata(f, revit.doc)

            if meta:
                source = meta.get("SourcePath", "")
                sheet = meta.get("SheetName", "")
                range_n = meta.get("RangeName", "")
                scale = meta.get("TextScale", 80.0)
                hint_path = ""
            elif f.Name.startswith(FAMILY_PREFIX):
                hint = TableDataManager.infer_hint_metadata(f.Name, db, recent)
                source = ""
                hint_path = hint.get("HintPath", "")
                sheet = hint.get("SheetName", "")
                range_n = hint.get("RangeName", "")
                scale = hint.get("TextScale", 80.0)
            else:
                continue

            item = data_model.LinkedTableItem(
                family_name=f.Name,
                source_path=source,
                sheet_name=sheet,
                range_name=range_n,
                text_scale=scale,
                family_id=f.Id,
                hint_path=hint_path
            )
            items.append(item)

        items.sort(key=lambda x: x.FamilyName)
        self.vm.LinkedTables = items

        if select_family_name:
            target_item = next((i for i in items if i.FamilyName == select_family_name), None)
            if target_item:
                self.vm.SelectedLinkedTable = target_item

    def load_metadata_into_editor(self, meta, family_name=None, hint_path=""):
        path = (meta.get("SourcePath", "") if meta else "")
        eff_hint = hint_path or (meta.get("HintPath", "") if meta else "")

        if family_name:
            self.vm.TableName = family_name
        elif meta and meta.get("FamilyName"):
            self.vm.TableName = meta.get("FamilyName")

        if meta and meta.get("TextScale"):
            self.vm.TextScale = str(meta.get("TextScale"))

        self.vm.SelectedSheet = None
        self.vm.SelectedRange = None

        if path and os.path.exists(path):
            self.vm.FilePath = path
            self.vm.LinkedFileHint = "Linked to: {}".format(path)
            self.refresh_data()

            target_sheet = meta.get("SheetName")
            if target_sheet and self.vm.Sheets:
                for s in self.vm.Sheets:
                    if s == target_sheet:
                        self.vm.SelectedSheet = s
                        break

            target_range = meta.get("RangeName")
            if target_range and self.vm.Ranges:
                matched_range = None
                for r in self.vm.Ranges:
                    if r.Name == target_range or r.name == target_range:
                        matched_range = r
                        break
                if matched_range:
                    self.vm.SelectedRange = matched_range
                elif self.vm.Ranges:
                    self.vm.SelectedRange = self.vm.Ranges[0]
            elif self.vm.Ranges:
                self.vm.SelectedRange = self.vm.Ranges[0]
        else:
            self.vm.FilePath = path or ""
            self.vm.Sheets = []
            self.vm.Ranges = []
            if path:
                self.vm.StatusText = "Missing File"
                self.vm.LinkedFileHint = "Missing original file: {}  (Click Browse to locate)".format(path)
            elif eff_hint:
                self.vm.StatusText = "Hint Only"
                self.vm.LinkedFileHint = "Hint (nearest match / history): {}  (Select source or Browse to relink)".format(eff_hint)
            else:
                self.vm.StatusText = "No Link"
                self.vm.LinkedFileHint = "No source file recorded for this table. Pick a Saved Source or click Browse."

        target_match = path or eff_hint
        if target_match and self.vm.RecentFiles:
            try:
                norm_match = os.path.normcase(os.path.abspath(target_match))
            except Exception:
                norm_match = str(target_match).lower()
            matched_recent = None
            for rf in self.vm.RecentFiles:
                try:
                    norm_rf = os.path.normcase(os.path.abspath(rf))
                except Exception:
                    norm_rf = str(rf).lower()
                if norm_rf == norm_match:
                    matched_recent = rf
                    break
            self._suppress_recent_event = True
            self.vm.SelectedRecentFile = matched_recent
            self._suppress_recent_event = False
        else:
            self._suppress_recent_event = True
            self.vm.SelectedRecentFile = None
            self._suppress_recent_event = False

    def Btn_Browse_Click(self, sender, args):
        sel_table = self.vm.SelectedLinkedTable
        init_dir = get_best_init_dir(
            self.vm.FilePath,
            sel_table.SourcePath if sel_table else None,
            sel_table.HintPath if sel_table else None,
            self.vm.SelectedRecentFile,
            TableDataManager.get_last_file_path(revit.doc)
        )
        f = forms.pick_file(file_ext='xlsx', init_dir=init_dir, title="Select Excel File")
        if f:
            if sel_table and (not sel_table.FileExists or not sel_table.SourcePath):
                self.vm.FilePath = f
                self.vm.LinkedFileHint = "Relinking '{}' -> {}".format(sel_table.FamilyName, f)
                self.refresh_data()
            else:
                self.vm.SelectedLinkedTable = None
                self.vm.SelectedSheet = None
                self.vm.SelectedRange = None
                self.vm.FilePath = f
                self.vm.LinkedFileHint = "Loaded: {}".format(f)
                self.refresh_data()

    def refresh_data(self):
        if not self.vm.FilePath or not os.path.exists(self.vm.FilePath):
            return

        self.vm.StatusText = "Reading..."
        sheets = excelextract.get_sheet_names(self.vm.FilePath)
        self.vm.Sheets = sheets

        raw_ranges = excelextract.get_print_areas(self.vm.FilePath)
        range_items = [data_model.NamedRangeItem({
            'name': '<Entire Worksheet>',
            'sheet': '',
            'formula': '',
            'is_full_sheet': True
        })]
        for r in raw_ranges:
            range_items.append(data_model.NamedRangeItem(r))

        self.vm.Ranges = range_items
        if sheets and not self.vm.SelectedSheet:
            self.vm.SelectedSheet = sheets[0]
        if range_items and not self.vm.SelectedRange:
            self.vm.SelectedRange = range_items[0]

        self.vm.StatusText = "Ready"

    def Cb_RecentFiles_SelectionChanged(self, sender, args):
        if self._suppress_recent_event:
            return
        picked = self.vm.SelectedRecentFile
        if not picked:
            return
        if os.path.exists(picked):
            self.vm.SelectedSheet = None
            self.vm.SelectedRange = None
            self.vm.FilePath = picked
            self.vm.LinkedFileHint = "Selected saved source: {}".format(picked)
            self.refresh_data()
        else:
            self.vm.LinkedFileHint = "Saved source missing on disk: {}  (Click Browse to open nearest folder)".format(picked)
            self.vm.StatusText = "Missing File"

    def Cb_Sheets_SelectionChanged(self, sender, args):
        if self.vm.SelectedSheet and not self.vm.SelectedLinkedTable:
            self.vm.TableName = "{}{}_Table".format(FAMILY_PREFIX, self.vm.SelectedSheet)

    def Cb_Ranges_SelectionChanged(self, sender, args):
        sel = self.vm.SelectedRange
        if sel and not self.vm.SelectedLinkedTable:
            if getattr(sel, 'IsFullSheet', False):
                clean = "{}_Table".format(self.vm.SelectedSheet if self.vm.SelectedSheet else "Worksheet")
            else:
                name_val = sel.Name if hasattr(sel, 'Name') else sel.name
                sheet_val = sel.Sheet if hasattr(sel, 'Sheet') else sel.sheet
                clean = name_val.replace("_", " ") if name_val != "Print_Area" else "{} Print Area".format(sheet_val)
            self.vm.TableName = "{}{}".format(FAMILY_PREFIX, clean)

    def Grid_LinkedTables_SelectionChanged(self, sender, args):
        sel = self.vm.SelectedLinkedTable
        if sel:
            meta = TableDataManager.get_metadata(sel.FamilyName, revit.doc)
            if meta:
                self.load_metadata_into_editor(meta, family_name=sel.FamilyName, hint_path=sel.HintPath)
            else:
                hint_meta = {
                    "SourcePath": sel.SourcePath,
                    "SheetName": sel.SheetName,
                    "RangeName": sel.RangeName,
                    "TextScale": sel.TextScale,
                    "HintPath": sel.HintPath
                }
                self.load_metadata_into_editor(hint_meta, family_name=sel.FamilyName, hint_path=sel.HintPath)

    def Btn_SelectAllLinked_Click(self, sender, args):
        self.vm.select_all_linked(True)

    def Btn_SelectNoneLinked_Click(self, sender, args):
        self.vm.select_all_linked(False)

    def Btn_UpdateChecked_Click(self, sender, args):
        checked = [t for t in self.vm.LinkedTables if t.IsChecked]
        if not checked:
            forms.alert("No linked tables selected for update.")
            return
        batch_update_tables(checked)
        self.refresh_linked_tables()

    def Btn_PlaceSelected_Click(self, sender, args):
        sel = self.vm.SelectedLinkedTable
        if not sel:
            forms.alert("Please select a table from the list.")
            return

        fam = next((f for f in DB.FilteredElementCollector(revit.doc).OfClass(DB.Family).ToElements() if f.Name == sel.FamilyName), None)
        if not fam:
            forms.alert("Table family '{}' not found in project.".format(sel.FamilyName))
            return

        self.Close()
        place_table_instance(fam)

    def Btn_Import_Click(self, sender, args):
        if not self.vm.FilePath:
            return

        s = self.vm.SelectedSheet
        r = self.vm.SelectedRange.Name if self.vm.SelectedRange else None
        n = self.vm.TableName or FAMILY_PREFIX + "Table"

        scale_pct = 80.0
        try:
            scale_pct = float(self.vm.TextScale)
        except Exception:
            scale_pct = 80.0

        place_inst = self.vm.PlaceOnView

        self.Close()
        create_table(self.vm.FilePath, s, r, n, scale_pct, place_instance=place_inst)

    def Btn_Close_Click(self, sender, args):
        self.Close()

# --- Entry Point ---
if __name__ == '__main__':
    sel = revit.get_selection()
    initial_meta = None
    initial_fam_name = None

    if len(sel) == 1 and isinstance(sel[0], DB.FamilyInstance):
        fam = sel[0].Symbol.Family
        if fam.Name.startswith(FAMILY_PREFIX) or TableDataManager.get_metadata(fam, revit.doc):
            initial_fam_name = fam.Name
            initial_meta = TableDataManager.get_metadata(fam, revit.doc)
            if initial_meta:
                initial_meta["FamilyName"] = fam.Name

    ExcelScheduleWindow(initial_family_name=initial_fam_name, metadata=initial_meta).ShowDialog()
