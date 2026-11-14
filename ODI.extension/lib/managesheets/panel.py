# -*- coding: utf-8 -*-
__title__ = "Manage Sheets"
__version__ = "4.2"
__doc__ = """A Modal WPF tool to align active Revit sheets against dynamically generated AIA UDS schemas using a Card-Style Tree Grid UI.
Features:
- Dynamically generates AIA standard sheets based on user inputs.
- Validates Sheet Names, Numbers, and Collections.
- Intelligent Fuzzy matching based on Discipline codes and Sheet Type synonyms.
- Editable Card-Style Tree Grid UI with inline editing and character diffs."""
__author__ = "ODI"
__context__ = "doc-project"

import os
import re
import difflib
import clr

clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Data import CollectionViewSource, PropertyGroupDescription
from System.Windows.Media import SolidColorBrush, Color as WpfColor, Colors
import System
from System.Windows import MessageBox, Visibility, SystemColors, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows.Controls import CheckBox
from System.Windows.Input import ICommand

from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory, 
    ElementId, ViewSheet, View, ViewType, BuiltInParameter, ParameterTypeId,
    Level, ForgeTypeId, ViewFamilyType, ViewFamily, ViewDrafting, ViewPlan, Viewport, XYZ, BoundingBoxXYZ
)
from pyrevit import revit, forms, script, HOST_APP
import classification
import project_settings

logger = script.get_logger()
cfg = script.get_config("ODI_ManageSheets")

# --- Constants & Rules ---

# Comprehensive AIA Discipline List
DISCIPLINE_CODES = {
  "CS": "COVER SHEET", "G": "GENERAL", "H": "HAZARDOUS MATERIALS",
  "V": "SURVEY", "B": "GEOTECHNICAL", "C": "CIVIL", "L": "LANDSCAPE",
  "S": "STRUCTURAL", "A": "ARCHITECTURAL", "I": "INTERIORS",
  "Q": "EQUIPMENT", "F": "FIRE PROTECTION", "P": "PLUMBING",
  "D": "PROCESS", "M": "MECHANICAL", "E": "ELECTRICAL",
  "W": "DISTRIBUTED ENERGY", "T": "TELECOMMUNICATIONS", "R": "RESOURCE",
  "X": "OTHER DISCIPLINES", "Z": "CONTRACTOR SHOP DRAWINGS",
  "O": "OPERATIONS", "AD": "ARCHITECTURAL DEMOLITION", "AF": "ARCHITECTURAL FINISHES",
  "AG": "ARCHITECTURAL GRAPHICS", "AI": "ARCHITECTURAL INTERIORS",
  "FA": "FIRE ALARM", "MH": "HVAC", "MP": "HVAC PIPING",
  "EL": "ELECTRICAL LIGHTING", "EP": "ELECTRICAL POWER",
  "RA": "EXISTING ARCHITECTURAL", "RS": "EXISTING STRUCTURAL",
  "RP": "EXISTING PLUMBING", "RM": "EXISTING MECHANICAL"
}

SERIES_MAP = {
    "0": "General", "1": "Plans", "2": "Elevations", "3": "Sections",
    "5": "Details", "6": "Schedules", "7": "Diagrams", "9": "ThreeD"
}

MODIFIERS = [
    "Overall", "Dimensions", "Construction", "Enlarged", 
    "Finishes", "Furniture", "Reflected Ceiling", "Framing"
]


from managesheets.data_model import *

# --- Revit Utility Functions ---

def get_sheet_collection_name(doc, sheet):
    if hasattr(ParameterTypeId, "SheetCollection"):
        param = sheet.GetParameter(ParameterTypeId.SheetCollection)
        if param and param.AsElementId() != ElementId.InvalidElementId:
            col_elem = doc.GetElement(param.AsElementId())
            if col_elem: return col_elem.Name
    else:
        param = sheet.LookupParameter(" Sheet Collection")
        if param and param.HasValue: return param.AsString()
    return "Undefined"

def assign_sheet_to_collection(doc, sheet, collection_name):
    if not doc or not collection_name: return
    try:
        try:
            from Autodesk.Revit.DB import SheetCollection, ParameterTypeId
            collector = FilteredElementCollector(doc).OfClass(SheetCollection)
            target_collection = next((c for c in collector if c.Name == collection_name), None)
            
            if not target_collection:
                target_collection = SheetCollection.Create(doc, collection_name)
                
            c_param = sheet.GetParameter(ParameterTypeId.SheetCollection)
            if c_param and not c_param.IsReadOnly:
                c_param.Set(target_collection.Id)
                return
        except:
            pass # Pre-2025
            
        param = sheet.LookupParameter(" Sheet Collection")
        if param and not param.IsReadOnly:
            param.Set(collection_name)
    except Exception as e:
        pass

def set_sheet_parameter(sheet, param_name, value):
    param = sheet.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        param.Set(value)

def ensure_sheet_parameter(doc, param_name):
    from Autodesk.Revit.DB import FilteredElementCollector, SharedParameterElement, BuiltInCategory, Category
    import os, tempfile, System
    
    # Check if parameter already exists on sheets
    sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
    if sheets:
        for test_sheet in sheets:
            if test_sheet.LookupParameter(param_name):
                return True # Parameter exists
            break
            
    # Try to find an existing Shared Parameter Element by name
    spe = None
    for sp in FilteredElementCollector(doc).OfClass(SharedParameterElement):
        if sp.Name == param_name:
            spe = sp
            break
            
    app = doc.Application
    original_file = app.SharedParametersFilename
    
    try:
        temp_file = os.path.join(tempfile.gettempdir(), "temp_shared_params_{}.txt".format(System.Guid.NewGuid()))
        if not spe:
            with open(temp_file, "w") as f:
                f.write("# This is a Revit shared parameter file.\n")
                f.write("# Do not edit manually.\n")
                f.write("*META\tVERSION\tMINVERSION\n")
                f.write("META\t2\t1\n")
                f.write("*GROUP\tID\tNAME\n")
                f.write("GROUP\t1\tManageSheets\n")
                f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n")
                f.write("PARAM\t{}\t{}\tTEXT\t\t1\t1\t\t1\t0\n".format(System.Guid.NewGuid(), param_name))
                
            app.SharedParametersFilename = temp_file
            sp_file = app.OpenSharedParameterFile()
            if not sp_file:
                from System.Windows import MessageBox
                MessageBox.Show("Error: app.OpenSharedParameterFile() returned None for temp file:\n" + temp_file, "Debug")
                return False
            group = sp_file.Groups.get_Item("ManageSheets")
            definition = group.Definitions.get_Item(param_name)
        else:
            definition = spe.GetDefinition()
            
        # Try to get the existing binding
        existing_binding = doc.ParameterBindings.Item[definition]
        if existing_binding:
            cat_set = existing_binding.Categories
            if not cat_set.Contains(Category.GetCategory(doc, BuiltInCategory.OST_Sheets)):
                cat_set.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Sheets))
                import Autodesk.Revit.DB as DB
                try:
                    res = doc.ParameterBindings.ReInsert(definition, existing_binding, DB.GroupTypeId.IdentityData)
                except Exception as ex1:
                    try:
                        res = doc.ParameterBindings.ReInsert(definition, existing_binding, DB.BuiltInParameterGroup.PG_IDENTITY_DATA)
                    except Exception as ex2:
                        from System.Windows import MessageBox
                        MessageBox.Show("Failed to ReInsert parameter due to version mismatch.\n2024 error: {}\nPre-2024 error: {}".format(ex1, ex2), "Debug")
                        return False
                if not res:
                    from System.Windows import MessageBox
                    MessageBox.Show("Failed to ReInsert parameter: " + param_name, "Debug")
                    return False
            return True
        else:
            cat_set = app.Create.NewCategorySet()
            cat_set.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Sheets))
            binding = app.Create.NewInstanceBinding(cat_set)
            
            import Autodesk.Revit.DB as DB
            try:
                res = doc.ParameterBindings.Insert(definition, binding, DB.GroupTypeId.IdentityData)
            except Exception as ex1:
                try:
                    res = doc.ParameterBindings.Insert(definition, binding, DB.BuiltInParameterGroup.PG_IDENTITY_DATA)
                except Exception as ex2:
                    from System.Windows import MessageBox
                    MessageBox.Show("Failed to Insert parameter due to version mismatch.\n2024 error: {}\nPre-2024 error: {}".format(ex1, ex2), "Debug")
                    return False
            if not res:
                from System.Windows import MessageBox
                MessageBox.Show("Failed to Insert parameter: " + param_name, "Debug")
                return False
                
            return True
    except Exception as e:
        import traceback
        from System.Windows import MessageBox
        MessageBox.Show("Error creating parameter '{}':\n\n{}".format(param_name, traceback.format_exc()), "Debug")
        return False
    finally:
        if original_file:
            try: app.SharedParametersFilename = original_file
            except: pass

# --- Generation Logic ---

def generate_suffixes(rows, cols, naming_scheme="Segment-Based", custom_schemes=None):
    if custom_schemes is None:
        custom_schemes = classification.NAMING_SCHEMES
        
    total = rows * cols
    if total <= 1: return [("", "")]
    
    custom_list = custom_schemes.get(naming_scheme, [])
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    suffixes = []
    
    for i in range(total):
        if i < 26: let = alphabet[i]
        else: let = "A" + alphabet[i-26]
        
        if i < len(custom_list):
            name_part = custom_list[i]
        else:
            prefix = naming_scheme if naming_scheme not in custom_schemes else "Segment"
            name_part = "{} {}".format(prefix, let)
            
        suffixes.append((let, name_part))
        
    return suffixes

def infer_modifier_level_scope(mod_name, series_code):
    """
    Classifies a modifier by its level multiplication scope so Cartoon Sheet Sets
    generate the right number of sheets instead of multiplying Roof/Site/Foundation
    across every building story:
      - 'series_once': Series != '1', or global single-sheet plans (Key Plan)
      - 'site_once': Site / Civil / Grading / Utility / Landscape plans -> 1 sheet
      - 'roof_only': Roof / Penthouse plans -> only Roof-matching levels (or 1 sheet)
      - 'foundation_only': Foundation / Basement / Slab / Footing plans -> only lowest/foundation level
      - 'per_floor': Standard Floor / Ceiling / Dimension / Demolition / Finish / Furniture / Framing / Power / Lighting / HVAC / Plumbing plans -> occupied floor levels
    """
    if str(series_code) != "1":
        return "series_once"
    m_low = (mod_name or "").lower()
    if any(k in m_low for k in ("site", "grading", "utility", "erosion", "landscape", "planting", "irrigation", "paving", "key plan", "overall plan")):
        return "site_once"
    if any(k in m_low for k in ("roof", "penthouse")):
        return "roof_only"
    if any(k in m_low for k in ("foundation", "footing", "basement", "subgrade", "slab on grade", "pile")):
        return "foundation_only"
    return "per_floor"

def is_enlarged_or_segmented_modifier(mod_name, modifier_overrides=None):
    """
    Only apply R x C grid segment expansion (A, B, C, D...) to modifiers that represent
    Enlarged / Area / Sector plans, OR when the user explicitly configured GridOverrides for that modifier.
    """
    if modifier_overrides and mod_name in modifier_overrides and modifier_overrides[mod_name] is not None:
        return True
    m_low = (mod_name or "").lower()
    return any(k in m_low for k in ("enlarged", "segment", "area", "sector", "zone", "quadrant", "unit"))

def filter_levels_for_scope(levels, scope):
    """
    Returns the list of level names (or [None] for non-level-bound sheets) appropriate for `scope`.
    """
    if not levels:
        return [None]
    if scope in ("series_once", "site_once"):
        return [None]
    if scope == "roof_only":
        roof_lvls = [l for l in levels if any(k in l.lower() for k in ("roof", "penthouse", "parapet", "high roof", "low roof"))]
        return roof_lvls if roof_lvls else [None]
    if scope == "foundation_only":
        fnd_lvls = [l for l in levels if any(k in l.lower() for k in ("foundation", "fnd", "footing", "basement", "bsmt", "cellar", "b1", "b2", "sub"))]
        return fnd_lvls if fnd_lvls else [levels[0]]
    # 'per_floor': exclude pure Roof / Parapet / Foundation / Footing datum levels when standard floor levels exist
    floor_lvls = [
        l for l in levels
        if not any(k in l.lower() for k in ("roof", "parapet", "t.o.", "top of", "foundation", "footing", "fnd", "subgrade", "pile"))
    ]
    return floor_lvls if floor_lvls else list(levels)

def generate_discipline_sheets(disc_code, levels, active_series, active_modifiers, global_cover, rows, cols, naming_scheme, custom_schemes=None, excluded_targets=None, shuffle_on_exclude=False, modifier_overrides=None):
    if excluded_targets is None: excluded_targets = set()
    if modifier_overrides is None: modifier_overrides = {}
    targets = []
    suffixes = generate_suffixes(rows, cols, naming_scheme, custom_schemes)
    disc_dict = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors", "CS": "Cover Sheet" }
    target_disc_name = disc_dict.get(disc_code, None)

    # 1. Pre-calculate series_mods for all active series
    all_series_mods = {}
    max_mods = 0
    
    for series in active_series:
        if series == "0" and global_cover and disc_code != "CS":
            continue
            
        series_mods = []
        known_names_and_cgs = set()
        for disc, groups in classification.CLASSIFICATION_DICT.items():
            if target_disc_name in classification.CLASSIFICATION_DICT:
                if disc != target_disc_name and disc != "Custom":
                    continue
                    
            for cg, types in groups.items():
                known_names_and_cgs.add(cg)
                for item in types:
                    if len(item) >= 2:
                        m_name, m_code = item[0], str(item[1])
                        known_names_and_cgs.add(m_name)
                        if (cg in active_modifiers or m_name in active_modifiers) and m_code.startswith(series):
                            if (m_name, m_code, cg) not in series_mods:
                                series_mods.append((m_name, m_code, cg))
                                
        # Add Custom Modifiers (those not in the known dictionary)
        for mod in active_modifiers:
            if mod not in known_names_and_cgs:
                c_info = classification.classify_sheet("000", mod)
                guess_code = c_info.get("drawingTypeCode", "99")
                
                if guess_code != "99":
                    if guess_code.startswith(series):
                        if (mod, series + "99", "Custom") not in series_mods:
                            series_mods.append((mod, series + "99", "Custom"))
                else:
                    if len(active_series) > 0 and series == active_series[0]:
                        if (mod, series + "99", "Custom") not in series_mods:
                            series_mods.append((mod, series + "99", "Custom"))
        
        if not series_mods:
            series_mods.append((SERIES_MAP.get(series, series).upper(), series + "01", "0{}. {}".format(series, SERIES_MAP.get(series, series))))
        elif series == "1":
            def _plan_mod_sort_key(item_tuple):
                mn = (item_tuple[0] or "").lower()
                if "site" in mn or "key plan" in mn: return (0, mn)
                if "demo" in mn: return (1, mn)
                if "foundation" in mn or "basement" in mn or "footing" in mn: return (2, mn)
                if "floor plan" in mn and "enlarged" not in mn: return (3, mn)
                if "dimension" in mn or "partition" in mn: return (4, mn)
                if "roof" in mn: return (5, mn)
                if "ceiling" in mn or "rcp" in mn: return (6, mn)
                if "finish" in mn or "furniture" in mn or "equipment" in mn: return (7, mn)
                if "enlarged" in mn or "sector" in mn or "area" in mn: return (8, mn)
                return (4, mn)
            series_mods.sort(key=_plan_mod_sort_key)
            
        all_series_mods[series] = series_mods
        if len(series_mods) > max_mods:
            max_mods = len(series_mods)
            
    # 2. Determine Universal Multiplier and Padding format
    seq_mult = 100 if max_mods > 10 else 10
    pad_fmt = "{:03d}" if max_mods > 10 else "{:02d}"
    
    # 3. Generate right-sized targets using Level Scope & Enlargement Scope
    for series in active_series:
        if series not in all_series_mods:
            continue
        series_mods = all_series_mods[series]
        
        if series == "1":
            # Separate per-floor modifiers from single-instance/scoped plan modifiers (Site, Foundation, Roof)
            running_seq = 1
            actual_seq = 1
            for mod_idx, (m_name, m_code, cg) in enumerate(series_mods):
                scope = infer_modifier_level_scope(m_name, series)
                scoped_levels = filter_levels_for_scope(levels, scope)
                use_segments = (len(suffixes) > 1) and is_enlarged_or_segmented_modifier(m_name, modifier_overrides)

                for lvl_idx, lvl in enumerate(scoped_levels):
                    baseline_seq = running_seq
                    running_seq += 1

                    baseline_num = "{}-{}{}".format(disc_code, series, pad_fmt.format(baseline_seq))
                    is_excluded = baseline_num in excluded_targets
                    
                    if shuffle_on_exclude:
                        if is_excluded:
                            actual_num = baseline_num + " [Skipped]"
                        else:
                            actual_num = "{}-{}{}".format(disc_code, series, pad_fmt.format(actual_seq))
                            actual_seq += 1
                    else:
                        actual_num = baseline_num

                    # Compose clean Sheet Name without redundant level duplication (e.g. avoid "ROOF PLAN ROOF")
                    base_title = m_name.upper()
                    if lvl and lvl.upper() not in base_title:
                        composed_name = "{} - {}".format(base_title, lvl.upper())
                    else:
                        composed_name = base_title

                    if not use_segments:
                        targets.append({
                            "num": actual_num,
                            "baseline_num": baseline_num,
                            "name": composed_name,
                            "series_name": "01. Plans",
                            "cg": cg,
                            "type": "1",
                            "level_name": lvl or ""
                        })
                    else:
                        for let, name_part in suffixes:
                            if m_name in modifier_overrides and modifier_overrides[m_name] is not None:
                                if let not in modifier_overrides[m_name]:
                                    continue
                                    
                            baseline_num_suffixed = "{}{}".format(baseline_num, let)
                            is_suffixed_excluded = baseline_num_suffixed in excluded_targets
                            
                            if shuffle_on_exclude and is_suffixed_excluded:
                                final_actual = baseline_num_suffixed + " [Skipped]"
                            else:
                                final_actual = "{}{}".format(actual_num.replace(" [Skipped]", ""), let)
                                
                            prefix_mod = composed_name
                            if not prefix_mod.startswith("ENLARGED"):
                                prefix_mod = "ENLARGED " + prefix_mod
                                
                            targets.append({
                                "num": final_actual,
                                "baseline_num": baseline_num_suffixed,
                                "name": "{} - {}".format(prefix_mod, name_part.upper()),
                                "series_name": "01. Plans",
                                "cg": cg,
                                "type": "1",
                                "level_name": lvl or ""
                            })
        else:
            actual_mod_seq = 1
            for mod_idx, (m_name, m_code, cg) in enumerate(series_mods):
                baseline_num = "{}-{}{}".format(disc_code, series, pad_fmt.format(mod_idx + 1))
                is_excluded = baseline_num in excluded_targets
                
                if shuffle_on_exclude:
                    if is_excluded:
                        actual_num = baseline_num + " [Skipped]"
                    else:
                        actual_num = "{}-{}{}".format(disc_code, series, pad_fmt.format(actual_mod_seq))
                        actual_mod_seq += 1
                else:
                    actual_num = baseline_num
                    
                coll_name = "0{}. {}".format(series, SERIES_MAP.get(series, series))
                targets.append({
                    "num": actual_num,
                    "baseline_num": baseline_num,
                    "name": m_name.upper(),
                    "series_name": coll_name,
                    "cg": cg,
                    "type": series,
                    "level_name": ""
                })
                
    return targets

# --- Intelligent Fuzzy Match Algorithm ---
def calculate_smart_score(sheet_row, target):
    """Calculates a match score using structural parsing, view names, and keyword weighting."""
    score = 0.0
    live_num = sheet_row.OriginalNumber
    live_name = sheet_row.OriginalName
    live_name_upper = live_name.upper()
    target_name_upper = target["name"].upper()
    
    # Extract View names
    view_names_upper = [v.Name.upper() for v in sheet_row.Views] if hasattr(sheet_row, "Views") else []
    
    # 1. Base Sequence Matcher Score (0.0 to 1.0)
    base_score = difflib.SequenceMatcher(None, live_name_upper, target_name_upper).ratio()
    score += base_score * 0.4 # Weight base score at 40%
    
    # 2. Exact Number Match (Massive Bonus)
    live_num_norm = re.sub(r'[^A-Z0-9]', '', live_num.upper())
    target_num_norm = re.sub(r'[^A-Z0-9]', '', target["num"].upper())
    
    if live_num_norm == target_num_norm:
        score += 1.0
        return score
        
    # 3. Keyword/Synonym Matching
    target_type = target["type"]
    synonyms = classification.SHEET_TYPE_SYNONYMS.get(target_type, [])
    has_synonym = False
    
    for syn in synonyms:
        if syn in live_name_upper:
            has_synonym = True
            score += 0.5 # 50% bonus for hitting a specific keyword
            break
        for vn in view_names_upper:
            if syn in vn:
                has_synonym = True
                score += 0.5
                break
        if has_synonym: break
            
    # Penalty: If it hit no synonyms for its type, but hits synonyms for a DIFFERENT type
    if not has_synonym:
        for t_type, syn_list in classification.SHEET_TYPE_SYNONYMS.items():
            if t_type == target_type: continue
            for syn in syn_list:
                if syn in live_name_upper or any(syn in vn for vn in view_names_upper):
                    score -= 0.3 # 30% penalty for contradicting keywords
                    break
                    
    # 4. Partial Number Match Bonus (e.g., A-101 vs A-101A)
    if len(target_num_norm) > 0 and (live_num_norm.startswith(target_num_norm) or target_num_norm.startswith(live_num_norm)):
        score += 0.2
        
    # 5. Viewport Token Overlap (Boost score if view names explicitly share words with target name)
    target_tokens = set(re.findall(r'[A-Z0-9]+', target_name_upper))
    for vn in view_names_upper:
        view_tokens = set(re.findall(r'[A-Z0-9]+', vn))
        overlap = target_tokens.intersection(view_tokens)
        if overlap:
            score += 0.15 * len(overlap)
            
    # 6. Explicit Revit View Level Match (Immense Boost)
    lvl_match = re.search(r'(LEVEL\s*0*(\d+)|ROOF|FOUNDATION|BASEMENT|MEZZANINE)', target_name_upper)
    if lvl_match and hasattr(sheet_row, "Views"):
        if lvl_match.group(2): 
            target_lvl = "LEVEL" + lvl_match.group(2)
        else:
            target_lvl = lvl_match.group(1).replace(" ", "")
            
        for v in sheet_row.Views:
            v_lvl = getattr(v, "LevelName", "")
            if v_lvl:
                v_lvl_upper = v_lvl.upper()
                v_lvl_norm = v_lvl_upper.replace(" ", "")
                v_num_match = re.search(r'0*(\d+)', v_lvl_upper)
                
                # If both have numbers, compare the numbers directly
                if v_num_match and lvl_match.group(2):
                    if v_num_match.group(1) == lvl_match.group(2):
                        score += 2.0
                        break
                # Otherwise string match (for Roof, Foundation, etc.)
                else:
                    if target_lvl in v_lvl_norm or v_lvl_norm in target_lvl:
                        score += 2.0
                        break
            
    return score

def is_dark_theme():
    try:
        from Autodesk.Revit.UI import UIThemeManager, UITheme
        return UIThemeManager.CurrentTheme == UITheme.Dark
    except Exception:
        return False

# --- Main Window ---

class NamingSchemeSettingsDialog(forms.WPFWindow):
    def __init__(self, current_schemes):
        xaml_path = os.path.join(os.path.dirname(__file__), "naming_settings.xaml")
        forms.WPFWindow.__init__(self, xaml_path)
        self.apply_theme()
        
        self.schemes_dict = {k: list(v) for k, v in current_schemes.items()}
        
        self.List_Schemes.SelectionChanged += self.on_scheme_selected
        self.Btn_AddScheme.Click += self.on_add_scheme
        self.Btn_RemoveScheme.Click += self.on_remove_scheme
        self.Btn_AddSegment.Click += self.on_add_segment
        self.Btn_RemoveSegment.Click += self.on_remove_segment
        self.Btn_Save.Click += self.on_save
        self.Btn_Cancel.Click += self.on_cancel
        
        self.refresh_schemes()
        
    def apply_theme(self):
        from System.Windows.Media import BrushConverter
        from System.Windows import SystemColors
        bc = BrushConverter()
        
        if is_dark_theme():
            colors = {
                "WindowBrush": "#1F2937",
                "ControlBrush": "#111827",
                "TextBrush": "#F9FAFB",
                "BorderBrush": "#4B5563",
                "ButtonBrush": "#374151",
                "AccentBrush": "#3B82F6",
            }
        else:
            colors = {
                "WindowBrush": "#F3F3F3",
                "ControlBrush": "#FFFFFF",
                "TextBrush": "#333333",
                "BorderBrush": "#CCCCCC",
                "ButtonBrush": "#DDDDDD",
                "AccentBrush": "#0078D7",
            }
            
        for key, hex_val in colors.items():
            if self.Resources.Contains(key):
                self.Resources[key] = bc.ConvertFromString(hex_val)
            else:
                self.Resources.Add(key, bc.ConvertFromString(hex_val))

    def refresh_schemes(self):
        self.List_Schemes.ItemsSource = None
        self.List_Schemes.ItemsSource = sorted(self.schemes_dict.keys())
        
    def on_scheme_selected(self, sender, e):
        if hasattr(e, "OriginalSource") and e.OriginalSource != sender: return
        sel = self.List_Schemes.SelectedItem
        self.List_Segments.ItemsSource = None
        if sel and sel in self.schemes_dict:
            self.List_Segments.ItemsSource = self.schemes_dict[sel]
            
    def on_add_scheme(self, sender, e):
        name = forms.ask_for_string(default="", prompt="Enter new Naming Scheme category name:", title="Add Scheme")
        if name and name not in self.schemes_dict:
            self.schemes_dict[name] = []
            self.refresh_schemes()
            self.List_Schemes.SelectedItem = name
            
    def on_remove_scheme(self, sender, e):
        sel = self.List_Schemes.SelectedItem
        if sel and sel in self.schemes_dict:
            del self.schemes_dict[sel]
            self.refresh_schemes()
            
    def on_add_segment(self, sender, e):
        sel = self.List_Schemes.SelectedItem
        if sel and sel in self.schemes_dict:
            name = forms.ask_for_string(default="", prompt="Enter new Segment suffix:", title="Add Segment")
            if name:
                self.schemes_dict[sel].append(name)
                self.on_scheme_selected(None, None)
                
    def on_remove_segment(self, sender, e):
        sel_scheme = self.List_Schemes.SelectedItem
        sel_seg = self.List_Segments.SelectedItem
        if sel_scheme and sel_seg and sel_seg in self.schemes_dict[sel_scheme]:
            self.schemes_dict[sel_scheme].remove(sel_seg)
            self.on_scheme_selected(None, None)
            
    def on_save(self, sender, e):
        self.DialogResult = True
        self.Close()
        
    def on_cancel(self, sender, e):
        self.DialogResult = False
        self.Close()


class GridOverrideDialog(forms.WPFWindow):
    def __init__(self, node_name, suffixes, current_overrides):
        xaml_path = os.path.join(os.path.dirname(__file__), "grid_override.xaml")
        forms.WPFWindow.__init__(self, xaml_path)
        self.apply_theme()
        
        self.Txt_Title.Text = "Configure Grids for: {}".format(node_name)
        
        self.grid_items = ObservableCollection[SelectableNode]()
        
        for let, name_part in suffixes:
            is_checked = True
            if current_overrides is not None:
                is_checked = let in current_overrides
            
            node = SelectableNode(let, is_checked=is_checked, display_name="{} ({})".format(name_part, let))
            self.grid_items.Add(node)
            
        self.List_Grids.ItemsSource = self.grid_items
        
        self.Btn_Save.Click += self.on_save
        self.Btn_Cancel.Click += self.on_cancel
        self.selected_overrides = None
        self.save_clicked = False
        
    def apply_theme(self):
        from System.Windows.Media import BrushConverter
        bc = BrushConverter()
        
        if is_dark_theme():
            self.Resources["WindowBrush"] = bc.ConvertFrom("#1F2937")
            self.Resources["ControlBrush"] = bc.ConvertFrom("#111827")
            self.Resources["TextBrush"] = bc.ConvertFrom("#F9FAFB")
            self.Resources["TextLightBrush"] = bc.ConvertFrom("#9CA3AF")
            self.Resources["BorderBrush"] = bc.ConvertFrom("#4B5563")
        else:
            self.Resources["WindowBrush"] = bc.ConvertFrom("#F3F3F3")
            self.Resources["ControlBrush"] = bc.ConvertFrom("#FFFFFF")
            self.Resources["TextBrush"] = bc.ConvertFrom("#333333")
            self.Resources["TextLightBrush"] = bc.ConvertFrom("#888888")
            self.Resources["BorderBrush"] = bc.ConvertFrom("#CCCCCC")
            
    def on_save(self, sender, args):
        all_checked = all([n.IsChecked for n in self.grid_items])
        if all_checked:
            self.selected_overrides = None
        else:
            self.selected_overrides = [n.Name for n in self.grid_items if n.IsChecked]
        self.save_clicked = True
        self.Close()
        
    def on_cancel(self, sender, args):
        self.Close()


class ModifierSettingsDialog(forms.WPFWindow):
    def __init__(self, current_dict):
        xaml_path = os.path.join(os.path.dirname(__file__), "modifier_settings.xaml")
        forms.WPFWindow.__init__(self, xaml_path)
        self.apply_theme()
        
        # Deep copy the dict so we don't modify it until saved
        import json
        self.class_dict = json.loads(json.dumps(current_dict))
        
        self.List_Disciplines.SelectionChanged += self.on_disc_selected
        self.List_Modifiers.SelectionChanged += self.on_mod_selected
        
        self.Btn_AddDiscipline.Click += self.on_add_disc
        self.Btn_RemoveDiscipline.Click += self.on_remove_disc
        self.Btn_AddModifier.Click += self.on_add_mod
        self.Btn_RemoveModifier.Click += self.on_remove_mod
        self.Btn_AddSheet.Click += self.on_add_sheet
        self.Btn_RemoveSheet.Click += self.on_remove_sheet
        if hasattr(self, 'Btn_MoveSheetUp'): self.Btn_MoveSheetUp.Click += self.on_move_sheet_up
        if hasattr(self, 'Btn_MoveSheetDown'): self.Btn_MoveSheetDown.Click += self.on_move_sheet_down
        
        self.Btn_Save.Click += self.on_save
        self.Btn_Cancel.Click += self.on_cancel
        
        self.refresh_disciplines()
        
    def apply_theme(self):
        from System.Windows.Media import BrushConverter
        bc = BrushConverter()
        
        if is_dark_theme():
            colors = {
                "WindowBrush": "#1F2937",
                "ControlBrush": "#111827",
                "TextBrush": "#F9FAFB",
                "BorderBrush": "#4B5563",
                "ButtonBrush": "#374151",
                "AccentBrush": "#3B82F6",
            }
        else:
            colors = {
                "WindowBrush": "#F3F3F3",
                "ControlBrush": "#FFFFFF",
                "TextBrush": "#333333",
                "BorderBrush": "#CCCCCC",
                "ButtonBrush": "#DDDDDD",
                "AccentBrush": "#0078D7",
            }
            
        for key, hex_val in colors.items():
            if self.Resources.Contains(key):
                self.Resources[key] = bc.ConvertFromString(hex_val)
            else:
                self.Resources.Add(key, bc.ConvertFromString(hex_val))

    def refresh_disciplines(self):
        self.List_Disciplines.ItemsSource = None
        self.List_Disciplines.ItemsSource = sorted(self.class_dict.keys())
        
    def on_disc_selected(self, sender, e):
        if hasattr(e, "OriginalSource") and e.OriginalSource != sender: return
        sel = self.List_Disciplines.SelectedItem
        self.List_Modifiers.ItemsSource = None
        self.List_Sheets.ItemsSource = None
        if sel and sel in self.class_dict:
            # Sort modifiers according to AIA code (0-9) by looking at their sheets
            def sort_key(mod_name):
                sheets = self.class_dict[sel][mod_name]
                if sheets and len(sheets) > 0 and len(sheets[0]) >= 2:
                    return str(sheets[0][1])
                return "9" # fallback
            self.List_Modifiers.ItemsSource = sorted(self.class_dict[sel].keys(), key=sort_key)
            
    def on_mod_selected(self, sender, e):
        if hasattr(e, "OriginalSource") and e.OriginalSource != sender: return
        disc = self.List_Disciplines.SelectedItem
        mod = self.List_Modifiers.SelectedItem
        self.List_Sheets.ItemsSource = None
        if disc and mod and disc in self.class_dict and mod in self.class_dict[disc]:
            sheets_raw = self.class_dict[disc][mod]
            display_sheets = ["{} [{}]".format(s[0], s[1]) if len(s) >= 2 else str(s) for s in sheets_raw]
            self.List_Sheets.ItemsSource = display_sheets
            
    def on_add_disc(self, sender, e):
        name = forms.ask_for_string(default="New Discipline", prompt="Enter Discipline Name:", title="Add Discipline")
        if name and name not in self.class_dict:
            self.class_dict[name] = {}
            self.refresh_disciplines()
            
    def on_remove_disc(self, sender, e):
        sel = self.List_Disciplines.SelectedItem
        if sel and sel in self.class_dict:
            del self.class_dict[sel]
            self.refresh_disciplines()
            
    def on_add_mod(self, sender, e):
        disc = self.List_Disciplines.SelectedItem
        if not disc: return
        name = forms.ask_for_string(default="New Sheet Series", prompt="Enter Sheet Series (Content Group) Name:", title="Add Sheet Series")
        if name and name not in self.class_dict[disc]:
            self.class_dict[disc][name] = []
            self.on_disc_selected(None, None)
            
    def on_remove_mod(self, sender, e):
        disc = self.List_Disciplines.SelectedItem
        mod = self.List_Modifiers.SelectedItem
        if disc and mod and mod in self.class_dict[disc]:
            del self.class_dict[disc][mod]
            self.on_disc_selected(None, None)
            
    def on_row_click(self, sender, e):
        # sender is the DataGridRow
        # The DataContext of the row is the SheetViewModel
        try:
            row_vm = sender.DataContext
            if row_vm and hasattr(row_vm, 'IsExpanded'):
                # Toggle expansion on row click
                row_vm.IsExpanded = not row_vm.IsExpanded
        except:
            pass

    def on_add_sheet(self, sender, e):
        disc = self.List_Disciplines.SelectedItem
        mod = self.List_Modifiers.SelectedItem
        if not disc or not mod: return
        name = forms.ask_for_string(default="Sheet Type, 1", prompt="Enter Modifier (Sheet Type) Name and Code separated by comma:\nCodes: 0=General, 1=Plans, 2=Elevs, 3=Sects, 5=Details, 6=Schedules", title="Add Modifier (Sheet Type)")
        if name:
            parts = [p.strip() for p in name.split(',')]
            if len(parts) >= 2:
                self.class_dict[disc][mod].append([parts[0], parts[1]])
                self.on_mod_selected(None, None)
            else:
                forms.alert("Please provide both name and code separated by a comma. (e.g. 'Floor Plan, 1')")
                
    def on_remove_sheet(self, sender, e):
        disc = self.List_Disciplines.SelectedItem
        mod = self.List_Modifiers.SelectedItem
        idx = self.List_Sheets.SelectedIndex
        if disc and mod and idx >= 0:
            del self.class_dict[disc][mod][idx]
            self.on_mod_selected(None, None)
            if self.List_Sheets.Items.Count > 0:
                self.List_Sheets.SelectedIndex = min(idx, self.List_Sheets.Items.Count - 1)
                
    def on_move_sheet_up(self, sender, e):
        disc = self.List_Disciplines.SelectedItem
        mod = self.List_Modifiers.SelectedItem
        idx = self.List_Sheets.SelectedIndex
        if disc and mod and idx > 0:
            lst = self.class_dict[disc][mod]
            lst[idx], lst[idx - 1] = lst[idx - 1], lst[idx]
            self.on_mod_selected(None, None)
            self.List_Sheets.SelectedIndex = idx - 1
            
    def on_move_sheet_down(self, sender, e):
        disc = self.List_Disciplines.SelectedItem
        mod = self.List_Modifiers.SelectedItem
        idx = self.List_Sheets.SelectedIndex
        if disc and mod and idx >= 0:
            lst = self.class_dict[disc][mod]
            if idx < len(lst) - 1:
                lst[idx], lst[idx + 1] = lst[idx + 1], lst[idx]
                self.on_mod_selected(None, None)
                self.List_Sheets.SelectedIndex = idx + 1
            
    def on_save(self, sender, e):
        # Update the in-memory classification dict
        classification.reload_classification(self.class_dict)
        classification.save_config_to_disk(class_dict=self.class_dict)
        # Flag the parent window to perform the Revit transaction upon exit
        if hasattr(self, 'Parent') and self.Parent:
            self.Parent.classification_needs_saving = True
        self.DialogResult = True
        self.Close()
        
    def on_cancel(self, sender, e):
        self.DialogResult = False
        self.Close()

class ManageSheetsPanel(forms.WPFWindow):
    def on_unhandled_exception(self, sender, e):
        e.Handled = True
        try:
            from pyrevit import forms
            forms.alert("An unexpected error occurred in Manage Sheets UI:\n\n" + str(e.Exception), title="WPF Error")
        except:
            pass

    def __init__(self):
        forms.WPFWindow.__init__(self, os.path.join(os.path.dirname(__file__), "ui.xaml"))
        self.Dispatcher.UnhandledException += self.on_unhandled_exception
        
        self.apply_theme()
        self.last_loaded_doc_hash = None
        
        # Restore Window Location
        from System.Windows import WindowStartupLocation
        try:
            w_left = cfg.get_option("window_left", None)
            w_top = cfg.get_option("window_top", None)
            if w_left is not None and w_top is not None:
                self.WindowStartupLocation = WindowStartupLocation.Manual
                self.Left = float(w_left)
                self.Top = float(w_top)
        except:
            pass
        
        # Subscribe to visibility to know when opened
        self.IsVisibleChanged += self.on_visible_changed
        
        # Subscribe to Revit ViewActivated to catch project switching
        try:
            from pyrevit import HOST_APP
            HOST_APP.uiapp.ViewActivated += self.on_view_activated
        except:
            pass
            

        self.Btn_RefreshData.Click += self.on_refresh_clicked
        self.Btn_ResetAll.Click += self.on_refresh_clicked
        self.Btn_Push.Click += self.sync_to_revit
        self.Btn_PrintLog.Click += self.print_debug_log
        self.Btn_EditNamingSchemes.Click += self.on_edit_naming_schemes
        self.Btn_EditModifiers.Click += self.on_edit_modifiers
        
        if hasattr(self, 'MainTabControl') and self.MainTabControl:
            self.MainTabControl.SelectionChanged += self.on_tab_changed
        if hasattr(self, 'Btn_QuickAddSheet') and self.Btn_QuickAddSheet:
            self.Btn_QuickAddSheet.Click += self.on_quick_add_sheet
        if hasattr(self, 'Btn_ToggleCartoonDrawer') and self.Btn_ToggleCartoonDrawer:
            self.Btn_ToggleCartoonDrawer.Click += self.on_toggle_cartoon_drawer
        if hasattr(self, 'Btn_AutoSuggestViews') and self.Btn_AutoSuggestViews:
            self.Btn_AutoSuggestViews.Click += self.on_auto_populate_all_views
        if hasattr(self, 'Chk_AddMissing') and self.Chk_AddMissing:
            self.Chk_AddMissing.Checked += self.on_include_cartoon_changed
            self.Chk_AddMissing.Unchecked += self.on_include_cartoon_changed
        if hasattr(self, 'Chk_ShowInvalid') and self.Chk_ShowInvalid:
            self.Chk_ShowInvalid.Checked += self.on_filter_changed
            self.Chk_ShowInvalid.Unchecked += self.on_filter_changed
        self._suppress_disc_filter = False
        if hasattr(self, 'Cmb_DisciplineFilter') and self.Cmb_DisciplineFilter:
            self.Cmb_DisciplineFilter.SelectionChanged += self.on_discipline_filter_changed
        
        # Setup Lists Select All
        self.Btn_DiscAll.Click += self.on_disc_all
        self.Btn_DiscNone.Click += self.on_disc_none
        self.Btn_LevelAll.Click += self.on_level_all
        self.Btn_LevelNone.Click += self.on_level_none
        self.Btn_SeriesAll.Click += self.on_series_all
        self.Btn_SeriesNone.Click += self.on_series_none
        self.Btn_ModAll.Click += self.on_mod_all
        self.Btn_ModNone.Click += self.on_mod_none
        
        # Expand/Collapse Handlers
        self.Btn_ExpandNav.Click += self.on_expand_nav
        self.Btn_CollapseNav.Click += self.on_collapse_nav
        self.Btn_ExpandEditor.Click += self.on_expand_editor
        self.Btn_CollapseEditor.Click += self.on_collapse_editor
        
        # Real-time Generator binds
        self.Sld_GridRows.ValueChanged += self.trigger_generation
        self.Sld_GridCols.ValueChanged += self.trigger_generation
        self.Chk_GlobalCover.Checked += self.trigger_generation
        self.Chk_GlobalCover.Unchecked += self.trigger_generation

        # Search Filters
        self.Txt_SearchLevels.TextChanged += self.filter_levels
        self.Txt_SearchDisciplines.TextChanged += self.filter_disciplines
        self.Txt_SearchSeries.TextChanged += self.filter_series
        self.Txt_SearchModifiers.TextChanged += self.filter_modifiers

        # MVVM Reactive State
        from data_model import ManageSheetsViewModel
        self.main_vm = ManageSheetsViewModel()
        self.main_vm.fill_floor_plans_callback = self.on_fill_floor_plans_column
        self.main_vm.fill_ceiling_plans_callback = self.on_fill_ceiling_plans_column
        self.main_vm.auto_fix_messy_callback = self.on_auto_fix_all_messy_sheets
        self.DataContext = self.main_vm

        # Data Models
        self.NavRoot = self.main_vm.NavRoot
        self.NavTree.ItemsSource = self.NavRoot
        self.NavTree.SelectedItemChanged += self.on_tree_selection_changed
        
        self.TargetSchemaRoot = ObservableCollection[NavTreeNode]()
        self.TargetSchemaTree.ItemsSource = self.TargetSchemaRoot
        self.TargetSchemaTree.SelectedItemChanged += self.on_tree_selection_changed
        if hasattr(self, 'Cmb_TargetCollection'):
            self.Cmb_TargetCollection.SelectionChanged += self.trigger_generation
        
        self.EditorItems = self.main_vm.EditorItems
        self.ViewTypes = ObservableCollection[str](["FloorPlan", "CeilingPlan", "Elevation", "Section", "Detail", "DraftingView", "3D", "Legend"])
        self.Scales = ObservableCollection[str](["1/16\" = 1'-0\"", "1/8\" = 1'-0\"", "1/4\" = 1'-0\"", "1/2\" = 1'-0\"", "1\" = 1'-0\"", "3/4\" = 1'-0\"", "1 1/2\" = 1'-0\"", "3\" = 1'-0\"", "NTS", "As indicated"])
        
        from data_model import AssignableViewNode
        self.AssignableViews = ObservableCollection[AssignableViewNode]()
        
        self.LevelNodes = ObservableCollection[SelectableNode]()
        self.List_Levels.ItemsSource = self.LevelNodes
        
        self.SeriesNodes = ObservableCollection[SelectableNode]()
        self.List_Series.ItemsSource = self.SeriesNodes
        
        self.ModifierRoot = ObservableCollection[NavTreeNode]()
        self.Tree_Modifiers.ItemsSource = self.ModifierRoot
        self.DisciplineNodes = ObservableCollection[SelectableNode]()
        self.List_Disciplines.ItemsSource = self.DisciplineNodes
        
        # Internal Storage
        self.all_grid_nodes = []
        self.generated_targets = []
        self.excluded_target_sheets = set()
        self._cached_harvest_data = None
        
        # Attach event handlers for the toggle switch
        self.Tgl_ShowExcluded.Checked += self.on_toggle_show_excluded
        self.Tgl_ShowExcluded.Unchecked += self.on_toggle_show_excluded
        self.Chk_ShuffleOnExclude.Checked += self.trigger_generation
        self.Chk_ShuffleOnExclude.Unchecked += self.trigger_generation
        
        self.Txt_SearchSchema.TextChanged += self.on_search_schema_text_changed
        self.Btn_ExpandSchema.Click += self.on_expand_schema
        self.Btn_CollapseSchema.Click += self.on_collapse_schema
        
        self.Btn_ExpandModifiers.Click += self.on_expand_modifiers
        self.Btn_CollapseModifiers.Click += self.on_collapse_modifiers
        
        # Title bar controls
        if hasattr(self, 'Btn_MinimizeWindow'): self.Btn_MinimizeWindow.Click += self.on_minimize_click
        if hasattr(self, 'Btn_SaveProjectSettings'): self.Btn_SaveProjectSettings.Click += self.save_to_project
        if hasattr(self, 'Btn_ClearProjectSettings'): self.Btn_ClearProjectSettings.Click += self.clear_project_settings
        if hasattr(self, 'Btn_ExportSchema'): self.Btn_ExportSchema.Click += self.export_schema
        if hasattr(self, 'Btn_ImportSchema'): self.Btn_ImportSchema.Click += self.import_schema
        if hasattr(self, 'Btn_CloseWindow'): self.Btn_CloseWindow.Click += self.on_close_click
        if hasattr(self, 'TitleBarGrid'): self.TitleBarGrid.MouseLeftButtonDown += self.on_drag_window
        
        # Resizing Handles
        if hasattr(self, 'ResizeRight'): self.ResizeRight.DragDelta += self.on_resize_right
        if hasattr(self, 'ResizeBottom'): self.ResizeBottom.DragDelta += self.on_resize_bottom
        if hasattr(self, 'ResizeLeft'): self.ResizeLeft.DragDelta += self.on_resize_left
        if hasattr(self, 'ResizeTop'): self.ResizeTop.DragDelta += self.on_resize_top
        if hasattr(self, 'ResizeBottomRight'): self.ResizeBottomRight.DragDelta += self.on_resize_bottom_right
        if hasattr(self, 'ResizeBottomLeft'): self.ResizeBottomLeft.DragDelta += self.on_resize_bottom_left
        if hasattr(self, 'ResizeTopRight'): self.ResizeTopRight.DragDelta += self.on_resize_top_right
        if hasattr(self, 'ResizeTopLeft'): self.ResizeTopLeft.DragDelta += self.on_resize_top_left
        
        if hasattr(self, 'FooterBorder') and self.FooterBorder:
            self.FooterBorder.Visibility = Visibility.Visible

        self.load_settings()
        self.generate_target_schema() # Initial run
        self.check_and_load_data()

    def on_toggle_cartoon_drawer(self, sender, e):
        self.main_vm.IsCartoonDrawerOpen = not self.main_vm.IsCartoonDrawerOpen
        if hasattr(self, 'CartoonDrawerPanel') and self.CartoonDrawerPanel:
            self.CartoonDrawerPanel.Visibility = Visibility.Visible if self.main_vm.IsCartoonDrawerOpen else Visibility.Collapsed

    def on_include_cartoon_changed(self, sender, e):
        self.refresh_active_workbench()

    def refresh_active_workbench(self):
        node = getattr(self, "_current_selected_node", None)
        if hasattr(self.TargetSchemaTree, "SelectedItem") and self.TargetSchemaTree.SelectedItem:
            if hasattr(self.TargetSchemaTree.SelectedItem, "NodeType"):
                node = self.TargetSchemaTree.SelectedItem
        elif hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
            if hasattr(self.NavTree.SelectedItem, "NodeType"):
                node = self.NavTree.SelectedItem
        if not node and self.TargetSchemaRoot.Count > 0:
            node = self.TargetSchemaRoot[0]
        elif not node and self.NavRoot.Count > 0:
            node = self.NavRoot[0]
        if node:
            self._last_selected_node = None
            self._fake_selection(node)

    def on_refresh_clicked(self, sender, e):
        try:
            self.load_revit_data()
        except Exception as ex:
            import traceback
            forms.alert(traceback.format_exc(), title="Error Refreshing Data")

    def on_drag_window(self, sender, e):
        self.DragMove()

    def on_resize_right(self, sender, e):
        self.Width = max(800, self.Width + e.HorizontalChange)

    def on_resize_bottom(self, sender, e):
        self.Height = max(600, self.Height + e.VerticalChange)

    def on_resize_left(self, sender, e):
        new_width = self.Width - e.HorizontalChange
        if new_width > 800:
            self.Width = new_width
            self.Left += e.HorizontalChange

    def on_resize_top(self, sender, e):
        new_height = self.Height - e.VerticalChange
        if new_height > 600:
            self.Height = new_height
            self.Top += e.VerticalChange

    def on_resize_bottom_right(self, sender, e):
        self.on_resize_bottom(sender, e)
        self.on_resize_right(sender, e)

    def on_resize_bottom_left(self, sender, e):
        self.on_resize_bottom(sender, e)
        self.on_resize_left(sender, e)

    def on_resize_top_right(self, sender, e):
        self.on_resize_top(sender, e)
        self.on_resize_right(sender, e)

    def on_resize_top_left(self, sender, e):
        self.on_resize_top(sender, e)
        self.on_resize_left(sender, e)

    def on_minimize_click(self, sender, e):
        import System.Windows
        self.WindowState = System.Windows.WindowState.Minimized

    def on_visible_changed(self, sender, e):
        if self.IsVisible:
            self.check_and_load_data()

    def on_view_activated(self, sender, e):
        # When switching tabs, check if the doc changed
        if self.IsVisible:
            self.check_and_load_data()



    def check_and_load_data(self):
        try:
            uidoc = HOST_APP.uiapp.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            if not doc:
                return
            
            # Use path + title to identify the document instance uniquely
            doc_id = doc.PathName + "_" + doc.Title
            if self.last_loaded_doc_hash == doc_id:
                return # Already loaded this doc
                
            # Auto-load unconditionally
            from System.Windows import Visibility
            self.AutoRefreshWarningPanel.Visibility = Visibility.Collapsed
            self.load_revit_data()
        except Exception as e:
            import traceback
            forms.alert(traceback.format_exc(), title="Error in check_and_load_data")


    
    def apply_theme(self):
        from System.Windows.Media import BrushConverter
        bc = BrushConverter()
        
        if is_dark_theme():
            colors = {
                "WindowBrush": "#1F2937",
                "ToolbarBrush": "#1F2937",
                "ControlBrush": "#111827",
                "FooterBrush": "#111827",
                "AltRowBrush": "#1F2937",
                "TextBrush": "#F9FAFB",
                "TextLightBrush": "#9CA3AF",
                "BorderBrush": "#4B5563",
                "ButtonBrush": "#374151",
                "HoverBrush": "#4B5563",
                "AccentBrush": "#3B82F6",
                "OnAccentTextBrush": "#FFFFFF",
                "SelectionBrush": "#1E3A8A",
                "SelectionBorderBrush": "#3B82F6",
                "SelectionTextBrush": "White",
                "InactiveSelectionBrush": "#374151",
                "CardBrush": "#374151",
                "CardBorderBrush": "#4B5563",
                "CardTextBrush": "#FFFFFF",
                "CardSubTextBrush": "#D1D5DB",
                "CardLabelBrush": "#9CA3AF",
                "CardValueBrush": "#FFFFFF",
                "CardAccentBrush": "#60A5FA",
                "ErrorBrush": "#7F1D1D",      # Dark Muted Burgundy
                "ErrorTextBrush": "#FECACA",  # Light pink/red text for contrast
                "SuccessBrush": "#064E3B",    # Dark Emerald
                "SuccessTextBrush": "#6EE7B7",
                "UpdateBrush": "#78350F",
                "ExtraBrush": "#4B5563",  # Gray-600 for Dark Theme     # Dark Amber/Yellow
                "UpdateTextBrush": "#FDE047"
            }
            for key, hex_val in colors.items():
                if self.Resources.Contains(key):
                    self.Resources[key] = bc.ConvertFromString(hex_val)
                else:
                    self.Resources.Add(key, bc.ConvertFromString(hex_val))

            # Inject aggressive system color overrides for DataGrid Row Selection States
            self.Resources[SystemColors.HighlightBrushKey] = bc.ConvertFromString(colors["SelectionBrush"])
            self.Resources[SystemColors.HighlightTextBrushKey] = bc.ConvertFromString(colors["SelectionTextBrush"])
            self.Resources[SystemColors.InactiveSelectionHighlightBrushKey] = bc.ConvertFromString(colors["InactiveSelectionBrush"])
            self.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = bc.ConvertFromString(colors["TextBrush"])
        else:
            if not self.Resources.Contains("OnAccentTextBrush"):
                self.Resources.Add("OnAccentTextBrush", bc.ConvertFromString("#FFFFFF"))

    def on_tab_changed(self, sender, e):
        if not hasattr(self, 'MainTabControl') or e.OriginalSource != self.MainTabControl:
            return
        if hasattr(self, 'FooterBorder') and self.FooterBorder:
            self.FooterBorder.Visibility = Visibility.Visible
        if self.MainTabControl.SelectedIndex == 1:
            self.build_matrix_board()
            
    def toggle_list(self, coll, state):
        for node in coll: node.IsChecked = state
        self.trigger_generation(None, None)
            
    def check_tree(self, coll, state):
        for node in coll:
            node.IsChecked = state
            if hasattr(node, "Children") and node.Children:
                self.check_tree(node.Children, state)
        self.trigger_generation(None, None)

    # Hard bound methods to prevent IronPython delegate garbage collection on lambdas
    def on_disc_all(self, sender, e): self.toggle_list(self.DisciplineNodes, True)
    def on_disc_none(self, sender, e): self.toggle_list(self.DisciplineNodes, False)
    def on_level_all(self, sender, e): self.toggle_list(self.LevelNodes, True)
    def on_level_none(self, sender, e): self.toggle_list(self.LevelNodes, False)
    def on_series_all(self, sender, e): self.toggle_list(self.SeriesNodes, True)
    def on_series_none(self, sender, e): self.toggle_list(self.SeriesNodes, False)
    def on_mod_all(self, sender, e): self.check_tree(self.ModifierRoot, True)
    def on_mod_none(self, sender, e): self.check_tree(self.ModifierRoot, False)
    def toggle_tree(self, coll, state):
        for node in coll:
            node.IsExpanded = state
            if hasattr(node, "Children") and node.Children:
                self.toggle_tree(node.Children, state)
                
    def toggle_editor(self, state):
        for node in self.EditorItems:
            node.IsExpanded = state


    def filter_levels(self, sender, e):
        view = CollectionViewSource.GetDefaultView(self.LevelNodes)
        txt = self.Txt_SearchLevels.Text.lower()
        if not txt: view.Filter = None
        else: view.Filter = System.Predicate[object](lambda item: txt in item.Name.lower())
        
    def filter_disciplines(self, sender, e):
        view = CollectionViewSource.GetDefaultView(self.DisciplineNodes)
        txt = self.Txt_SearchDisciplines.Text.lower()
        if not txt: view.Filter = None
        else: view.Filter = System.Predicate[object](lambda item: txt in item.Name.lower())
        
    def filter_modifiers(self, sender, e):
        view = CollectionViewSource.GetDefaultView(self.ModifierNodes)
        txt = self.Txt_SearchModifiers.Text.lower()
        if not txt: view.Filter = None
        else: view.Filter = System.Predicate[object](lambda item: txt in item.Name.lower())
    def filter_series(self, sender, e):
        view = CollectionViewSource.GetDefaultView(self.SeriesNodes)
        txt = self.Txt_SearchSeries.Text.lower()
        if not txt: view.Filter = None
        else: view.Filter = System.Predicate[object](lambda item: txt in item.Name.lower())

    def check_tree(self, root_collection, is_checked):
        for node in root_collection:
            self._recursive_check(node, is_checked)
        self.trigger_generation(None, None)
            
    def _recursive_check(self, node, is_checked):
        node.IsChecked = is_checked
        if hasattr(node, "Children"):
            for child in node.Children:
                self._recursive_check(child, is_checked)

    def load_settings(self):
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        p_setup = project_settings.load_project_setup(doc) or {}
        
        # Setup Target Collection ComboBox
        existing_collections = set([
            "SCHEMATIC DESIGN",
            "DESIGN DEVELOPMENT",
            "PERMIT SET",
            "CONSTRUCTION DOCUMENTS",
            "BID SET",
            "ISSUED FOR CONSTRUCTION"
        ])
        if doc:
            from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet
            sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
            for s in sheets:
                if s.IsPlaceholder: continue
                c_name = get_sheet_collection_name(doc, s)
                if c_name and c_name != "Undefined":
                    existing_collections.add(c_name)
            
        if self.Cmb_TargetCollection.ItemsSource is None and self.Cmb_TargetCollection.Items.Count > 0:
            self.Cmb_TargetCollection.Items.Clear()
        self.Cmb_TargetCollection.ItemsSource = sorted(list(existing_collections))
        self.Cmb_TargetCollection.Text = p_setup.get("target_collection", cfg.get_option("target_collection", "PERMIT SET"))
        self.Cmb_TargetCollection.LostFocus += self.trigger_generation
        self.Cmb_TargetCollection.SelectionChanged += self.trigger_generation
        
        self.Sld_GridRows.Value = float(p_setup.get("grid_rows", cfg.get_option("grid_rows", 1)))
        self.Sld_GridCols.Value = float(p_setup.get("grid_cols", cfg.get_option("grid_cols", 1)))
        
        saved_discs = p_setup.get("disciplines", cfg.get_option("disciplines", ["A"]))
        aia_order = ['CS', 'G', 'H', 'V', 'B', 'C', 'L', 'S', 'A', 'I', 'Q', 'F', 'P', 'D', 'M', 'E', 'W', 'T', 'R', 'X', 'Z', 'O']
        sorted_disciplines = sorted(DISCIPLINE_CODES.items(), key=lambda x: aia_order.index(x[0]) if x[0] in aia_order else 999)
        def handle_discipline_change():
            self.sync_modifier_disciplines()
            self.generate_target_schema()
            
        for k, v in sorted_disciplines:
            is_chk = k in saved_discs
            self.DisciplineNodes.Add(SelectableNode("{} - {}".format(k, v), is_checked=is_chk, callback=handle_discipline_change))
            
        saved_series = p_setup.get("series", cfg.get_option("series", ["0", "1", "2", "3"]))
        sorted_series = sorted(SERIES_MAP.items(), key=lambda x: x[0])
        
        def handle_series_change():
            self.sync_modifier_series()
            self.generate_target_schema()
            
        for k, v in sorted_series:
            is_chk = k in saved_series
            self.SeriesNodes.Add(SelectableNode("{} - {}".format(k, v), is_checked=is_chk, callback=handle_series_change))
            
        # Load Naming Schemes first so event triggers don't fail
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        if doc:
            self._loaded_naming_schemes = project_settings.load_naming_schemes(doc)
        else:
            self._loaded_naming_schemes = project_settings.load_naming_schemes(None)
        if not self._loaded_naming_schemes:
            self._loaded_naming_schemes = classification.NAMING_SCHEMES

        saved_modifiers = p_setup.get("selected_modifiers", cfg.get_option("selected_modifiers", []))
        
        p_custom = p_setup.get("custom_modifiers", [])
        g_custom = cfg.get_option("custom_modifiers", [])
        # Merge lists and remove duplicates while preserving order where possible
        custom_modifiers = list(dict.fromkeys(p_custom + g_custom))
        
        modifier_grid_overrides = p_setup.get("modifier_grid_overrides", cfg.get_option("modifier_grid_overrides", {}))
        self.load_modifiers_from_cfg(saved_modifiers, custom_modifiers, modifier_grid_overrides)
        self.Chk_GlobalCover.IsChecked = p_setup.get("global_cover", cfg.get_option("global_cover", False))
        
        self.excluded_target_sheets = set(p_setup.get("excluded_target_sheets", cfg.get_option("excluded_target_sheets", [])))
        self.Chk_ShuffleOnExclude.IsChecked = p_setup.get("shuffle_on_exclude", cfg.get_option("shuffle_on_exclude", False))

        self.Cmb_NamingScheme.ItemsSource = self._loaded_naming_schemes.keys()
        self.Cmb_NamingScheme.SelectedItem = p_setup.get("naming_scheme", cfg.get_option("naming_scheme", "Segment-Based"))
        self.Cmb_NamingScheme.LostFocus += self.trigger_generation
        self.Cmb_NamingScheme.SelectionChanged += self.trigger_generation
        
        self.ModifierRoot.Clear()

        c_dict = classification.CLASSIFICATION_DICT
        
        # Sort Disciplines first based on AIA order mapping if possible, else alphabetically
        disc_keys = sorted(c_dict.keys())
        
        for disc in disc_keys:
            groups = c_dict[disc]
            disc_node = NavTreeNode(disc, "Discipline")
            self.ModifierRoot.Add(disc_node)
            
            # Sort modifiers by their drawing code
            def get_code(mod):
                shs = groups[mod]
                if shs and len(shs) > 0 and len(shs[0]) >= 2: return str(shs[0][1])
                return "9"
            
            sorted_groups = sorted(groups.keys(), key=get_code)
            
            for cg in sorted_groups:
                cg_node = NavTreeNode(cg, "ContentGroup", parent=disc_node)
                cg_node.IsExpanded = False # Start collapsed
                cg_node.callback = self.generate_target_schema
                
                for sheet_info in groups[cg]:
                    if sheet_info and len(sheet_info) > 0:
                        m_name = sheet_info[0]
                        m_code = str(sheet_info[1]) if len(sheet_info) >= 2 else "9"
                        m_node = NavTreeNode(m_name, "Modifier", tag=m_code, parent=cg_node)
                        m_node.callback = self.generate_target_schema
                        m_node.grid_config_callback = self.open_grid_config
                        m_node.GridOverrides = modifier_grid_overrides.get(m_name, None)
                        m_node.IsSegmentable = ("Plan" in cg) or ("Plan" in m_name)
                        # Bypass the IsChecked setter logic temporarily to avoid mass triggering during load
                        m_node._is_checked = m_name in saved_modifiers
                        cg_node.Children.Add(m_node)
                
                # Evaluate cg_node state after adding children
                cg_node._evaluate_checked_state()
                
                # Always add the content group even if it's currently empty, 
                # so users can see newly added groups before sheets are assigned.
                disc_node.Children.Add(cg_node)
            
            disc_node._evaluate_checked_state()
            
        # Add custom modifiers to a "Custom" group
        if custom_modifiers:
            custom_node = NavTreeNode("Custom", "ContentGroup")
            custom_node.IsExpanded = False
            custom_node.callback = self.generate_target_schema
            for mod in custom_modifiers:
                m_node = NavTreeNode(mod, "Modifier", tag="9", parent=custom_node)
                m_node.callback = self.generate_target_schema
                m_node.grid_config_callback = self.open_grid_config
                m_node.GridOverrides = modifier_grid_overrides.get(mod, None)
                m_node.IsSegmentable = True
                m_node._is_checked = True # Active by definition if it's in this list
                custom_node.Children.Add(m_node)
            self.ModifierRoot.Add(custom_node)
            
        self.sync_modifier_disciplines()
        self.sync_modifier_series()

    def load_modifiers_from_cfg(self, saved_modifiers=None, custom_modifiers=None, modifier_grid_overrides=None):
        if saved_modifiers is None: saved_modifiers = cfg.get_option("selected_modifiers", [])
        if custom_modifiers is None: custom_modifiers = cfg.get_option("custom_modifiers", [])
        if modifier_grid_overrides is None: modifier_grid_overrides = getattr(cfg, 'modifier_grid_overrides', {})
        
        for disc_node in self.ModifierRoot:
            if disc_node.Name == "Custom": continue
            for cg_node in disc_node.Children:
                for m_node in cg_node.Children:
                    m_node._is_checked = m_node.Name in saved_modifiers
                    m_node.GridOverrides = modifier_grid_overrides.get(m_node.Name, None)
                    m_node.OnPropertyChanged("IsChecked")
                cg_node._evaluate_checked_state()
            disc_node._evaluate_checked_state()
            
        custom_node = None
        for dn in self.ModifierRoot:
            if dn.Name == "Custom":
                custom_node = dn
                break
                
        if custom_node:
            self.ModifierRoot.Remove(custom_node)
            
        if custom_modifiers:
            custom_node = NavTreeNode("Custom", "ContentGroup")
            custom_node.IsExpanded = False
            custom_node.callback = self.generate_target_schema
            for mod in custom_modifiers:
                m_node = NavTreeNode(mod, "Modifier", tag="9", parent=custom_node)
                m_node.callback = self.generate_target_schema
                m_node.grid_config_callback = self.open_grid_config
                m_node.GridOverrides = modifier_grid_overrides.get(mod, None)
                m_node.IsSegmentable = True
                m_node._is_checked = True 
                custom_node.Children.Add(m_node)
            self.ModifierRoot.Add(custom_node)
            
        self.sync_modifier_disciplines()
        self.sync_modifier_series()

    def sync_modifier_disciplines(self):
        selected_discs = []
        for n in self.DisciplineNodes:
            if n.IsChecked:
                code = n.Name.split(' - ')[0]
                selected_discs.append(code)
                
        disc_dict = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors", "CS": "Cover Sheet" }
        selected_disc_names = [disc_dict.get(c, c).upper() for c in selected_discs]
        
        for disc_node in self.ModifierRoot:
            if disc_node.Name == "Custom": continue
            
            is_active = disc_node.Name.upper() in selected_disc_names
            disc_node.IsVisible = is_active
            
            if not is_active:
                if disc_node.IsChecked is not False:
                    disc_node.IsChecked = False
                    
    def sync_modifier_series(self):
        active_series = []
        for n in self.SeriesNodes:
            if n.IsChecked:
                code = n.Name.split(' - ')[0]
                active_series.append(code)
                
        for disc_node in self.ModifierRoot:
            for cg_node in disc_node.Children:
                for m_node in cg_node.Children:
                    m_code = str(m_node.Tag)
                    # Enable if any active series code matches the start of the modifier's code
                    m_node.IsEnabled = any(m_code.startswith(s) for s in active_series)

    def save_settings(self):
        try:
            cfg.grid_rows = int(self.Sld_GridRows.Value)
            cfg.grid_cols = int(self.Sld_GridCols.Value)
        except: pass
        cfg.global_cover = bool(self.Chk_GlobalCover.IsChecked)
        
        # Save Window Location
        try:
            cfg.window_left = self.Left
            cfg.window_top = self.Top
        except: pass
        
        ns = None
        if self.Cmb_NamingScheme.SelectedItem:
            ns = str(self.Cmb_NamingScheme.SelectedItem)
        if not ns:
            ns = self.Cmb_NamingScheme.Text
        if not ns:
            ns = "Segment-Based"
        cfg.naming_scheme = ns
        selected_discs = []
        for n in self.DisciplineNodes:
            if n.IsChecked:
                code = n.Name.split(' - ')[0]
                selected_discs.append(code)
        cfg.disciplines = selected_discs
        
        selected_series = []
        for n in self.SeriesNodes:
            if n.IsChecked:
                code = n.Name.split(' - ')[0]
                selected_series.append(code)
        cfg.series = selected_series
        
        selected_modifiers = []
        all_custom = []
        
        for disc_node in self.ModifierRoot:
            if disc_node.Name == "Custom":
                for m_node in disc_node.Children:
                    if m_node.IsChecked:
                        selected_modifiers.append(m_node.Name)
                    all_custom.append(m_node.Name)
            else:
                for cg_node in disc_node.Children:
                    for m_node in cg_node.Children:
                        if m_node.IsChecked:
                            selected_modifiers.append(m_node.Name)
                        
        cfg.selected_modifiers = selected_modifiers
        cfg.custom_modifiers = all_custom
        cfg.global_cover = bool(self.Chk_GlobalCover.IsChecked)
        cfg.shuffle_on_exclude = bool(self.Chk_ShuffleOnExclude.IsChecked)
        cfg.excluded_target_sheets = list(self.excluded_target_sheets)
        
        modifier_grid_overrides = {}
        for disc_node in self.ModifierRoot:
            if disc_node.Name == "Custom":
                for m_node in disc_node.Children:
                    if m_node.GridOverrides is not None:
                        modifier_grid_overrides[m_node.Name] = m_node.GridOverrides
            else:
                for cg_node in disc_node.Children:
                    for m_node in cg_node.Children:
                        if m_node.GridOverrides is not None:
                            modifier_grid_overrides[m_node.Name] = m_node.GridOverrides
        cfg.modifier_grid_overrides = modifier_grid_overrides
        
        script.save_config()
        
    def save_to_project(self, sender, e):
        self.save_settings()
        setup_dict = {
            "target_collection": self.Cmb_TargetCollection.Text,
            "grid_rows": int(self.Sld_GridRows.Value),
            "grid_cols": int(self.Sld_GridCols.Value),
            "global_cover": bool(self.Chk_GlobalCover.IsChecked),
            "naming_scheme": self.Cmb_NamingScheme.SelectedItem if self.Cmb_NamingScheme.SelectedItem else self.Cmb_NamingScheme.Text,
            "disciplines": [n.Name.split(' - ')[0] for n in self.DisciplineNodes if n.IsChecked],
            "series": [n.Name.split(' - ')[0] for n in self.SeriesNodes if n.IsChecked],
            "selected_modifiers": cfg.selected_modifiers,
            "custom_modifiers": cfg.custom_modifiers,
            "shuffle_on_exclude": bool(self.Chk_ShuffleOnExclude.IsChecked),
            "excluded_target_sheets": list(self.excluded_target_sheets),
            "modifier_grid_overrides": cfg.modifier_grid_overrides
        }
        
        # Flag the window to perform the Revit transaction upon exit
        self.setup_needs_saving = True
        self.pending_setup_dict = setup_dict
        forms.alert("Project settings saved to memory. They will be committed to the Revit model when you close this window.", title="Settings Saved")
            
    def clear_project_settings(self, sender, e):
        import System.Windows.MessageBox as MessageBox
        import System.Windows.MessageBoxButton as MessageBoxButton
        import System.Windows.MessageBoxImage as MessageBoxImage
        
        result = MessageBox.Show("Are you sure you want to delete the saved schema from the Revit project?", "Clear Schema", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        if result == System.Windows.MessageBoxResult.Yes:
            uidoc = HOST_APP.uiapp.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            if doc:
                if project_settings.clear_project_setup(doc):
                    forms.alert("Schema deleted successfully. You may want to 'Reset All' to clear the current UI state.")
                else:
                    forms.alert("No schema data found in the project to delete.")
                    
    def export_schema(self, sender, e):
        import json
        import Microsoft.Win32 as Win32
        
        self.save_settings()
        setup_dict = {
            "target_collection": self.Cmb_TargetCollection.Text,
            "grid_rows": int(self.Sld_GridRows.Value),
            "grid_cols": int(self.Sld_GridCols.Value),
            "global_cover": bool(self.Chk_GlobalCover.IsChecked),
            "naming_scheme": self.Cmb_NamingScheme.SelectedItem if self.Cmb_NamingScheme.SelectedItem else self.Cmb_NamingScheme.Text,
            "disciplines": [n.Name.split(' - ')[0] for n in self.DisciplineNodes if n.IsChecked],
            "series": [n.Name.split(' - ')[0] for n in self.SeriesNodes if n.IsChecked],
            "selected_modifiers": cfg.selected_modifiers,
            "custom_modifiers": cfg.custom_modifiers,
            "shuffle_on_exclude": bool(self.Chk_ShuffleOnExclude.IsChecked),
            "excluded_target_sheets": list(self.excluded_target_sheets),
            "modifier_grid_overrides": getattr(cfg, 'modifier_grid_overrides', {})
        }
        
        dlg = Win32.SaveFileDialog()
        dlg.FileName = "ProjectSchema"
        dlg.DefaultExt = ".json"
        dlg.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        
        if dlg.ShowDialog() == True:
            try:
                with open(dlg.FileName, 'w') as f:
                    json.dump(setup_dict, f, indent=4)
                forms.alert("Schema exported successfully.", title="Export Success")
            except Exception as ex:
                forms.alert("Failed to export schema: {}".format(ex), title="Export Error")

    def import_schema(self, sender, e):
        import json
        import Microsoft.Win32 as Win32
        
        dlg = Win32.OpenFileDialog()
        dlg.DefaultExt = ".json"
        dlg.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        
        if dlg.ShowDialog() == True:
            try:
                with open(dlg.FileName, 'r') as f:
                    setup_dict = json.load(f)
                
                # Apply the loaded dict (simulate loading project settings)
                if "target_collection" in setup_dict: self.Cmb_TargetCollection.Text = setup_dict["target_collection"]
                if "grid_rows" in setup_dict: self.Sld_GridRows.Value = float(setup_dict["grid_rows"])
                if "grid_cols" in setup_dict: self.Sld_GridCols.Value = float(setup_dict["grid_cols"])
                if "global_cover" in setup_dict: self.Chk_GlobalCover.IsChecked = setup_dict["global_cover"]
                if "naming_scheme" in setup_dict: self.Cmb_NamingScheme.SelectedItem = setup_dict["naming_scheme"]
                if "shuffle_on_exclude" in setup_dict: self.Chk_ShuffleOnExclude.IsChecked = setup_dict["shuffle_on_exclude"]
                
                i_custom = setup_dict.get("custom_modifiers", [])
                g_custom = cfg.get_option("custom_modifiers", [])
                merged_custom = list(dict.fromkeys(i_custom + g_custom))
                
                cfg.selected_modifiers = setup_dict.get("selected_modifiers", [])
                cfg.custom_modifiers = merged_custom
                cfg.modifier_grid_overrides = setup_dict.get("modifier_grid_overrides", {})
                self.excluded_target_sheets = set(setup_dict.get("excluded_target_sheets", []))
                
                # Re-check disciplines and series
                self.toggle_list(self.DisciplineNodes, False)
                if "disciplines" in setup_dict:
                    for d_code in setup_dict["disciplines"]:
                        for node in self.DisciplineNodes:
                            if node.Name.startswith(d_code + " - "):
                                node.IsChecked = True
                                
                self.toggle_list(self.SeriesNodes, False)
                if "series" in setup_dict:
                    for s_code in setup_dict["series"]:
                        for node in self.SeriesNodes:
                            if node.Name.startswith(s_code + " - "):
                                node.IsChecked = True
                                
                self.load_modifiers_from_cfg(
                    saved_modifiers=setup_dict.get("selected_modifiers", []),
                    custom_modifiers=merged_custom,
                    modifier_grid_overrides=setup_dict.get("modifier_grid_overrides", {})
                )
                self.generate_target_schema()
                self.save_settings() # Save loaded config to user config immediately
            except Exception as ex:
                forms.alert("Failed to import schema: {}".format(ex), title="Import Error")
            
    def on_close_click(self, sender, e):
        res = MessageBox.Show("Do you want to save the current Project Setup to the Revit model before closing?", "Save Settings", MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
        if res == MessageBoxResult.Yes:
            self.save_to_project(None, None)
            self.Close()
        elif res == MessageBoxResult.No:
            self.Close()
        # if Cancel, do nothing

    def on_add_modifier(self, sender, e):
        pass # Function removed from UI
        
    def trigger_generation(self, sender, e):
        if getattr(self, "_is_generating", False): return
        self._is_generating = True
        try:
            self.generate_target_schema()
        finally:
            self._is_generating = False
        
    def on_checkbox_click(self, sender, e):
        # Workaround for WPF TabControl recycling CheckBoxes and forcing them to False
        try:
            node = sender.DataContext
            if node and hasattr(node, "IsChecked"):
                node.IsChecked = bool(sender.IsChecked)
        except: pass
        
    def on_target_checkbox_click(self, sender, e):
        try:
            node = sender.DataContext
            if node and hasattr(node, "IsTargetIncluded"):
                node.IsTargetIncluded = bool(sender.IsChecked)
        except: pass

    def on_edit_naming_schemes(self, sender, e):
        try:
            dialog = NamingSchemeSettingsDialog(self._loaded_naming_schemes)
            try: dialog.Owner = self.Parent
            except: pass
            if dialog.ShowDialog():
                uidoc = HOST_APP.uiapp.ActiveUIDocument
                doc = uidoc.Document if uidoc else None
                if doc:
                    # Update Document with new schemes
                    project_settings.save_naming_schemes(doc, dialog.schemes_dict)
                # Also save globally to JSON resource
                classification.save_config_to_disk(naming_schemes=dialog.schemes_dict)
                
                self._loaded_naming_schemes = dialog.schemes_dict
                
                # Refresh ComboBox
                prev_sel = self.Cmb_NamingScheme.SelectedItem or self.Cmb_NamingScheme.Text
                self.Cmb_NamingScheme.ItemsSource = None
                self.Cmb_NamingScheme.ItemsSource = self._loaded_naming_schemes.keys()
                if prev_sel in self._loaded_naming_schemes:
                    self.Cmb_NamingScheme.SelectedItem = prev_sel
                else:
                    self.Cmb_NamingScheme.SelectedItem = "Segment-Based"
                self.generate_target_schema()
        except Exception as ex:
            import traceback
            forms.alert(traceback.format_exc(), title="Error Editing Naming Schemes")
            
    def on_edit_modifiers(self, sender, e):
        try:
            dialog = ModifierSettingsDialog(classification.CLASSIFICATION_DICT)
            try: dialog.Owner = self.Parent
            except: pass
            if dialog.ShowDialog():
                # Already saved in dialog, just refresh the tree
                self.load_settings()
        except Exception as ex:
            import traceback
            forms.alert(traceback.format_exc(), title="Error Editing Modifiers")

    def open_grid_config(self, node):
        try:
            r = int(self.Txt_GridRows.Text)
            c = int(self.Txt_GridCols.Text)
        except:
            r, c = 1, 1
            
        scheme_idx = self.Cmb_NamingScheme.SelectedIndex
        if scheme_idx < 0: scheme_idx = 0
        scheme_name = list(classification.NAMING_SCHEMES.keys())[scheme_idx] if scheme_idx < len(classification.NAMING_SCHEMES) else "Segment-Based"
        
        suffixes = generate_suffixes(r, c, naming_scheme=scheme_name)
        if len(suffixes) <= 1:
            forms.alert("Global segmentation is currently set to 1x1. Increase rows/cols to configure segments.", title="No Segments Available")
            return
            
        dialog = GridOverrideDialog(node.Name, suffixes, node.GridOverrides)
        try: dialog.Owner = self.Parent
        except: pass
        
        dialog.ShowDialog()
        
        if dialog.save_clicked:
            node.GridOverrides = dialog.selected_overrides
            self.generate_target_schema()

    def generate_target_schema(self):
        self.save_settings()
        try:
            r = int(self.Txt_GridRows.Text)
            c = int(self.Txt_GridCols.Text)
        except:
            r, c = 1, 1
            
        selected_discs = []
        for n in self.DisciplineNodes:
            if n.IsChecked:
                code = n.Name.split(' - ')[0]
                selected_discs.append(code)
                
        active_series = []
        for n in self.SeriesNodes:
            if n.IsChecked:
                code = n.Name.split(' - ')[0]
                active_series.append(code)
                
        active_levels = [lvl.Name for lvl in self.LevelNodes if lvl.IsChecked]
        
        active_modifiers = []
        def get_checked(nodes):
            for n in nodes:
                if n.IsChecked and n.NodeType == "Modifier":
                    active_modifiers.append(n.Name)
                if hasattr(n, "Children"):
                    get_checked(n.Children)
        get_checked(self.ModifierRoot)
        global_cover = bool(self.Chk_GlobalCover.IsChecked)
        
        # Read SelectedItem first, fallback to Text, then default.
        # This fixes the WPF bug where Text lags behind SelectionChanged.
        naming_scheme = None
        if self.Cmb_NamingScheme.SelectedItem:
            naming_scheme = str(self.Cmb_NamingScheme.SelectedItem)
        if not naming_scheme:
            naming_scheme = self.Cmb_NamingScheme.Text
        if not naming_scheme:
            naming_scheme = "Segment-Based"
                    
        shuffle = bool(self.Chk_ShuffleOnExclude.IsChecked)
        
        # Collect Grid Overrides
        modifier_overrides = {}
        def collect_overrides(node):
            if node.NodeType == "Modifier" and getattr(node, "GridOverrides", None) is not None:
                modifier_overrides[node.Name] = node.GridOverrides
            for child in node.Children:
                collect_overrides(child)
        for disc_node in self.ModifierRoot:
            collect_overrides(disc_node)
            
        self.generated_targets = []
        for d in selected_discs:
            self.generated_targets.extend(generate_discipline_sheets(d, active_levels, active_series, active_modifiers, global_cover, r, c, naming_scheme, self._loaded_naming_schemes, self.excluded_target_sheets, shuffle, modifier_overrides))
            
        self.TargetSchemaRoot.Clear()
        t_root = NavTreeNode("All Collections & AIA Schema", "Root", is_expanded=True)
        t_root.target_callback = self.on_target_node_toggled
        self.TargetSchemaRoot.Add(t_root)
        
        target_col_name = "PERMIT SET"
        if hasattr(self, 'Cmb_TargetCollection') and self.Cmb_TargetCollection:
            if self.Cmb_TargetCollection.SelectedItem:
                sel_item = self.Cmb_TargetCollection.SelectedItem
                target_col_name = sel_item.Content if hasattr(sel_item, "Content") else str(sel_item)
            elif self.Cmb_TargetCollection.Text:
                target_col_name = self.Cmb_TargetCollection.Text
                
        col_nodes = {}
        active_col_node = NavTreeNode(target_col_name, "Collection", tag=target_col_name, parent=t_root, is_expanded=True)
        active_col_node.target_callback = self.on_target_node_toggled
        col_nodes[target_col_name] = active_col_node
        t_root.Children.Add(active_col_node)
        
        d_map = {}
        cg_map = {}
        
        for t in self.generated_targets:
            match = re.match(r"^([A-Z]+)[- ]?(\d+)", t["num"].upper())
            disc_code = match.group(1) if match else "Other"
            disc_dict = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors", "CS": "Cover Sheet" }
            disc_name = "{} - {}".format(disc_code, disc_dict.get(disc_code, "Discipline")) if disc_code != "Other" else "Uncategorized"
            
            d_key = (target_col_name, disc_name)
            if d_key not in d_map:
                dn = NavTreeNode(disc_name, "Discipline", tag=(target_col_name, disc_name), parent=active_col_node)
                dn.target_callback = self.on_target_node_toggled
                d_map[d_key] = dn
                active_col_node.Children.Add(dn)
                
            cg_name = t.get("cg", "Unknown Content Group")
            cg_key = (target_col_name, disc_name, cg_name)
            
            if cg_key not in cg_map:
                cn = NavTreeNode(cg_name, "ContentGroup", tag=(target_col_name, disc_name, cg_name), parent=d_map[d_key])
                cn.target_callback = self.on_target_node_toggled
                cg_map[cg_key] = cn
                d_map[d_key].Children.Add(cn)
                
            sn = NavTreeNode("{} - {}".format(t["num"], t["name"]), "Sheet", tag=t["baseline_num"], parent=cg_map[cg_key])
            if t["baseline_num"] in self.excluded_target_sheets:
                sn.IsTargetIncluded = False
                if not self.Tgl_ShowExcluded.IsChecked:
                    sn.IsVisible = False
            sn.target_callback = self.on_target_node_toggled
            cg_map[cg_key].Children.Add(sn)
            
            t_root.Count += 1
            active_col_node.Count += 1
            d_map[d_key].Count += 1
            cg_map[cg_key].Count += 1
            
        # Also include any other existing Revit Sheet Collections from all_grid_nodes (Rule 1)
        for existing_sh in getattr(self, "all_grid_nodes", []):
            ex_col = getattr(existing_sh, "OriginalCollectionName", existing_sh.CollectionName) or "Default"
            if ex_col not in col_nodes:
                ex_col_node = NavTreeNode(ex_col, "Collection", tag=ex_col, parent=t_root, is_expanded=False)
                ex_col_node.target_callback = self.on_target_node_toggled
                col_nodes[ex_col] = ex_col_node
                t_root.Children.Add(ex_col_node)
            if ex_col != target_col_name:
                col_nodes[ex_col].Count += 1
                t_root.Count += 1
                ex_disc = existing_sh.DisciplineName or "Uncategorized"
                ex_cg = existing_sh.ContentGroupName or "Uncategorized"
                d_key = (ex_col, ex_disc)
                if d_key not in d_map:
                    dn = NavTreeNode(ex_disc, "Discipline", tag=(ex_col, ex_disc), parent=col_nodes[ex_col], is_expanded=False)
                    dn.target_callback = self.on_target_node_toggled
                    d_map[d_key] = dn
                    col_nodes[ex_col].Children.Add(dn)
                d_map[d_key].Count += 1
                cg_key = (ex_col, ex_disc, ex_cg)
                if cg_key not in cg_map:
                    cn = NavTreeNode(ex_cg, "ContentGroup", tag=(ex_col, ex_disc, ex_cg), parent=d_map[d_key], is_expanded=False)
                    cn.target_callback = self.on_target_node_toggled
                    cg_map[cg_key] = cn
                    d_map[d_key].Children.Add(cn)
                cg_map[cg_key].Count += 1
            
        # Only enable Sheet Reviewer tab if a schema is actively generated
        if hasattr(self, 'Tab_SheetReviewer'):
            self.main_vm.IsReviewerReady = (len(self.generated_targets) > 0)
            
        # Auto-expand the target schema tree so user sees the new combinations immediately
        if hasattr(self, 'toggle_tree'):
            self.toggle_tree(self.TargetSchemaRoot, True)

        # In the Unified Workbench, immediately refresh the active sheet list when Cartoon Set rules change
        if hasattr(self, 'all_grid_nodes') and not getattr(self, '_suppress_workbench_refresh', False):
            self.refresh_active_workbench()

    def on_target_node_toggled(self, node):
        import System
        from System.Windows.Input import Keyboard, ModifierKeys
        
        is_top_call = not getattr(self, '_in_target_toggle', False)
        self._in_target_toggle = True
        try:
            # Shift-Select Logic
            if is_top_call and Keyboard.Modifiers == ModifierKeys.Shift:
                if hasattr(self, '_last_toggled_schema_node') and self._last_toggled_schema_node:
                    last_node = self._last_toggled_schema_node
                    flat_list = []
                    def _flatten(n):
                        if n.IsVisible:
                            flat_list.append(n)
                            if n.IsExpanded:
                                for c in n.Children:
                                    _flatten(c)
                    for root in self.TargetSchemaRoot:
                        _flatten(root)
                        
                    if last_node in flat_list and node in flat_list:
                        idx1 = flat_list.index(last_node)
                        idx2 = flat_list.index(node)
                        start, end = min(idx1, idx2), max(idx1, idx2)
                        for i in range(start, end + 1):
                            n = flat_list[i]
                            if n != node and n != last_node:
                                old_cb = n.target_callback
                                n.target_callback = None
                                n.IsTargetIncluded = node.IsTargetIncluded
                                n.target_callback = old_cb
                                
                                if n.NodeType == "Sheet":
                                    if not n.IsTargetIncluded:
                                        self.excluded_target_sheets.add(n.Tag)
                                        if not self.Tgl_ShowExcluded.IsChecked:
                                            n.IsVisible = False
                                    else:
                                        if n.Tag in self.excluded_target_sheets:
                                            self.excluded_target_sheets.remove(n.Tag)
                                        n.IsVisible = True
            
            if is_top_call:
                self._last_toggled_schema_node = node

            if node.NodeType == "Sheet":
                if not node.IsTargetIncluded:
                    self.excluded_target_sheets.add(node.Tag)
                    if not self.Tgl_ShowExcluded.IsChecked:
                        node.IsVisible = False
                else:
                    if node.Tag in self.excluded_target_sheets:
                        self.excluded_target_sheets.remove(node.Tag)
                    node.IsVisible = True
            else:
                for child in node.Children:
                    old_cb = child.target_callback
                    child.target_callback = None
                    child.IsTargetIncluded = node.IsTargetIncluded
                    child.target_callback = old_cb
                    self.on_target_node_toggled(child)
        finally:
            if is_top_call:
                self._in_target_toggle = False
                if hasattr(self, 'Chk_ShuffleOnExclude') and self.Chk_ShuffleOnExclude.IsChecked:
                    self.generate_target_schema()
                else:
                    self.save_settings()
                    self.refresh_active_workbench()
                
    def on_toggle_show_excluded(self, sender, e):
        show_excluded = bool(self.Tgl_ShowExcluded.IsChecked)
        def _update_visibility(node):
            if not node.IsTargetIncluded:
                node.IsVisible = show_excluded
            else:
                node.IsVisible = True
            for child in node.Children:
                _update_visibility(child)
        for root in self.TargetSchemaRoot:
            _update_visibility(root)
            
    def on_search_schema_text_changed(self, sender, e):
        search_term = self.Txt_SearchSchema.Text.strip().lower()
        
        def filter_node(node):
            if not search_term:
                if not node.IsTargetIncluded and not self.Tgl_ShowExcluded.IsChecked:
                    node.IsVisible = False
                else:
                    node.IsVisible = True
                for child in node.Children:
                    filter_node(child)
                return True
                
            if not node.Children:
                match = search_term in node.Name.lower()
                if match:
                    if not node.IsTargetIncluded and not self.Tgl_ShowExcluded.IsChecked:
                        node.IsVisible = False
                    else:
                        node.IsVisible = True
                else:
                    node.IsVisible = False
                return match
                
            any_child_visible = False
            for child in node.Children:
                if filter_node(child):
                    any_child_visible = True
                    
            if search_term in node.Name.lower():
                any_child_visible = True
                
            node.IsVisible = any_child_visible
            if any_child_visible:
                node.IsExpanded = True
            return any_child_visible

        for root_node in self.TargetSchemaRoot:
            filter_node(root_node)


    def on_expand_schema(self, sender, e):
        self.toggle_tree(self.TargetSchemaRoot, True)

    def on_collapse_schema(self, sender, e):
        self.toggle_tree(self.TargetSchemaRoot, False)

    def on_expand_modifiers(self, sender, e):
        self.toggle_tree(self.ModifierRoot, True)

    def on_collapse_modifiers(self, sender, e):
        self.toggle_tree(self.ModifierRoot, False)
        
    def on_expand_nav(self, sender, e):
        self.toggle_tree(self.NavRoot, True)
        
    def on_collapse_nav(self, sender, e):
        self.toggle_tree(self.NavRoot, False)
        
    def on_expand_editor(self, sender, e):
        self.toggle_editor(True)
        
    def on_collapse_editor(self, sender, e):
        self.toggle_editor(False)

    def refresh_global_cache(self, doc):
        self._cached_global_sheets = {}
        self._cached_global_views = set()
        
        from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet, View
        all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        for s in all_sheets:
            if not s.IsTemplate:
                s_coll = get_sheet_collection_name(doc, s) or "Default"
                k = (s_coll, s.SheetNumber.strip().lower())
                if k not in self._cached_global_sheets:
                    self._cached_global_sheets[k] = []
                s_id_val = s.Id.IntegerValue if hasattr(s.Id, 'IntegerValue') else s.Id.Value
                self._cached_global_sheets[k].append(s_id_val)
                
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        for v in all_views:
            if not v.IsTemplate:
                self._cached_global_views.add(v.Name.lower())
                
    def load_revit_data(self):
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        if not doc or not uidoc: return
        
        c_dict = project_settings.load_classification_dict(doc)
        if c_dict:
            classification.reload_classification(c_dict)
        else:
            classification.reload_classification()
        
        # Cache this document to prevent duplicate loading
        self.last_loaded_doc_hash = doc.PathName + "_" + doc.Title
        
        self.refresh_global_cache(doc)
        try:
            import reconciliation
            self._cached_harvest_data = reconciliation.harvest_project(doc)
        except Exception:
            self._cached_harvest_data = None
        
        # Populate TitleBlocks ComboBox
        class TitleBlockOption:
            def __init__(self, tb):
                self.Id = tb.Id
                try:
                    self.Name = "{} - {}".format(tb.FamilyName, tb.Name)
                except:
                    try:
                        from Autodesk.Revit.DB import BuiltInParameter
                        p = tb.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        self.Name = p.AsString() if p else "Unknown TitleBlock"
                    except:
                        self.Name = "Unknown TitleBlock"
                        
        titleblocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
        tb_options = [TitleBlockOption(t) for t in titleblocks]
        self.Combo_TitleBlocks.ItemsSource = sorted(tb_options, key=lambda t: t.Name)
        if tb_options:
            self.Combo_TitleBlocks.SelectedIndex = 0
            
        # Hide the warning if manually refreshed
        from System.Windows import Visibility
        self.AutoRefreshWarningPanel.Visibility = Visibility.Collapsed
        
        self.LevelNodes.Clear()
        self.all_grid_nodes = []
        self.NavRoot.Clear()
        self.EditorItems.Clear()
        self.AssignableViews.Clear()
        
        import data_model
        from data_model import AssignableViewNode, format_revit_scale
        
        # Collect all assignable views with ViewType, Level, and real Scale
        from Autodesk.Revit.DB import View, Viewport, FamilyInstance
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        viewports = FilteredElementCollector(doc).OfClass(Viewport).ToElements()
        
        # Map sheet -> annotation symbols on sheet to detect existing Graphic Scales & North Arrows
        sheet_has_gs = set()
        sheet_has_na = set()
        try:
            annos = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType().ToElements()
            for a in annos:
                owner_id = a.OwnerViewId
                if not owner_id or owner_id == ElementId.InvalidElementId:
                    continue
                owner_val = owner_id.IntegerValue if hasattr(owner_id, "IntegerValue") else owner_id.Value
                fam_name = ""
                try:
                    fam_name = (a.Symbol.Family.Name or "").lower()
                except Exception:
                    pass
                if "graphic scale" in fam_name or "scale bar" in fam_name:
                    sheet_has_gs.add(owner_val)
                elif "north" in fam_name:
                    sheet_has_na.add(owner_val)
        except Exception:
            pass
        
        placed_views_map = {}
        for vp in viewports:
            try:
                sheet = doc.GetElement(vp.SheetId)
                if sheet:
                    placed_views_map[vp.ViewId.ToString()] = sheet.SheetNumber
            except: pass
            
        assignable_list = []
        for v in all_views:
            if not v.IsTemplate and v.CanBePrinted and v.ViewType != ViewType.DrawingSheet:
                v_id_str = v.Id.ToString()
                is_placed = v_id_str in placed_views_map
                sh_num = placed_views_map.get(v_id_str, "")
                v_lvl = v.GenLevel.Name if hasattr(v, "GenLevel") and v.GenLevel else ""
                v_scale = getattr(v, "Scale", 48)
                node_av = AssignableViewNode(v.Id, v.Name, is_placed, sh_num, view_type=str(v.ViewType), level_name=v_lvl, scale_int=v_scale)
                self.AssignableViews.Add(node_av)
                assignable_list.append(node_av)
        data_model.AVAILABLE_ASSIGNABLE_VIEWS = assignable_list
        data_model.AVAILABLE_EXISTING_VIEWS = assignable_list
        
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        
        def get_elev(l):
            try: return l.ProjectElevation
            except AttributeError: return l.Elevation
            
        self._suppress_workbench_refresh = True
        try:
            for idx, lvl in enumerate(sorted(levels, key=get_elev)):
                elev_val = get_elev(lvl)

                try:
                    from Autodesk.Revit.DB import UnitFormatUtils, SpecTypeId
                    elev_str = UnitFormatUtils.Format(doc.GetUnits(), SpecTypeId.Length, elev_val, False)
                except ImportError:
                    try:
                        from Autodesk.Revit.DB import UnitFormatUtils, UnitType
                        elev_str = UnitFormatUtils.Format(doc.GetUnits(), UnitType.UT_Length, elev_val, False, False)
                    except Exception:
                        elev_str = str(elev_val)

                display_name = "{:02d} - {} ({})".format(idx + 1, lvl.Name, elev_str)
                self.LevelNodes.Add(SelectableNode(lvl.Name, True, self.generate_target_schema, display_name=display_name))
        finally:
            self._suppress_workbench_refresh = False

        data_model.AVAILABLE_LEVELS = [lvl.Name for lvl in sorted(levels, key=get_elev)]
        vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        
        from Autodesk.Revit.DB import ViewFamily
        valid_families = [ViewFamily.FloorPlan, ViewFamily.CeilingPlan, ViewFamily.StructuralPlan, ViewFamily.Drafting, ViewFamily.Elevation, ViewFamily.Section, ViewFamily.Detail]
        
        vft_names = []
        for vft in vfts:
            if hasattr(vft, 'ViewFamily') and vft.ViewFamily in valid_families:
                try:
                    name = vft.Name
                    if name: vft_names.append(name)
                except AttributeError:
                    try:
                        from Autodesk.Revit.DB import BuiltInParameter
                        p = vft.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if p and p.AsString(): vft_names.append(p.AsString())
                    except: pass
        data_model.AVAILABLE_VIEW_FAMILY_TYPES = sorted(list(set(vft_names)))
        
        selected_ids = uidoc.Selection.GetElementIds()
        scope_ids = [i for i in selected_ids if isinstance(doc.GetElement(i), ViewSheet)]
        if not scope_ids: sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        else: sheets = [doc.GetElement(i) for i in scope_ids]
            
        for s in sheets:
            if s.IsPlaceholder: continue
            
            c_name = get_sheet_collection_name(doc, s)
            
            c_result = classification.classify_sheet(s.SheetNumber, s.Name)
            disc_name = c_result.get("discipline", "Unknown")
            cg_name = c_result.get("contentGroup", "Uncategorized")
            
            sh_row = SheetViewModel(
                s.Id, s.SheetNumber, s.Name, c_name,
                discipline_name=disc_name, content_group_name=cg_name,
                validation_callback=self.run_validation,
                number_changed_callback=self.on_sheet_number_changed,
                context_callback=self.get_sheet_context,
                clone_sheet_callback=self.on_clone_sheet,
                auto_populate_callback=self.auto_populate_sheet_views,
                select_sheet_callback=self.on_select_sheet_vm,
                accept_suggestion_callback=self.on_accept_sheet_suggestion
            )
            self.all_grid_nodes.append(sh_row)
            
            s_id_val = s.Id.IntegerValue if hasattr(s.Id, "IntegerValue") else s.Id.Value
            has_gs_on_sheet = s_id_val in sheet_has_gs
            has_na_on_sheet = s_id_val in sheet_has_na
            
            for v_id in s.GetAllPlacedViews():
                v = doc.GetElement(v_id)
                if not v: continue
                
                lvl_name = ""
                if hasattr(v, "GenLevel") and v.GenLevel:
                    lvl_name = v.GenLevel.Name
                    
                v_number = ""
                param_num = v.get_Parameter(BuiltInParameter.VIEWER_DETAIL_NUMBER)
                if param_num: v_number = param_num.AsString() or ""
                
                v_name = v.Name
                param_title = v.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                if param_title and param_title.AsString():
                    v_name = param_title.AsString()
                    
                v_scale_str = format_revit_scale(getattr(v, "Scale", 48))
                is_plan_type = any(k in str(v.ViewType).lower() for k in ("floor", "ceiling", "engineering", "structural", "area"))
                
                v_row = ViewViewModel(
                    v.Id, v_name, str(v.ViewType),
                    level_name=lvl_name,
                    scale=v_scale_str,
                    view_number=v_number,
                    parent_sheet=sh_row,
                    has_graphic_scale=has_gs_on_sheet if v_scale_str != "NTS" else False,
                    has_north_arrow=has_na_on_sheet if is_plan_type else False
                )
                sh_row.Views.Add(v_row)

        # Build full 4-tier Navigation Tree: All Sheets -> Collection -> Discipline -> ContentGroup
        self.rebuild_nav_tree()

        # 2D Matrix Board Default Mode:
        # Keep the Advanced Rule Book drawer collapsed by default so the 2D Building x Series Matrix
        # takes center stage, and Ground-Truth First: start with Chk_AddMissing = False.
        self.main_vm.IsCartoonDrawerOpen = False
        if hasattr(self, 'CartoonDrawerPanel') and self.CartoonDrawerPanel:
            self.CartoonDrawerPanel.Visibility = Visibility.Collapsed
        if hasattr(self, 'Chk_AddMissing') and self.Chk_AddMissing:
            self.Chk_AddMissing.Checked -= self.on_include_cartoon_changed
            self.Chk_AddMissing.Unchecked -= self.on_include_cartoon_changed
            self.Chk_AddMissing.IsChecked = False
            self.Chk_AddMissing.Checked += self.on_include_cartoon_changed
            self.Chk_AddMissing.Unchecked += self.on_include_cartoon_changed

        # Prepopulate AIA Schema
        self._suppress_workbench_refresh = True
        try:
            self.generate_target_schema()
        finally:
            self._suppress_workbench_refresh = False

        # Wire tree selection handlers and select NavRoot[0] first (Ground-Truth First)
        if hasattr(self, 'NavTree') and self.NavTree:
            self.NavTree.SelectedItemChanged -= self.on_tree_selection_changed
            self.NavTree.SelectedItemChanged += self.on_tree_selection_changed
        if hasattr(self, 'TargetSchemaTree') and self.TargetSchemaTree:
            self.TargetSchemaTree.SelectedItemChanged -= self.on_tree_selection_changed
            self.TargetSchemaTree.SelectedItemChanged += self.on_tree_selection_changed

        if self.NavRoot.Count > 0:
            self._last_selected_node = None
            self._fake_selection(self.NavRoot[0])
        elif self.TargetSchemaRoot.Count > 0:
            self._last_selected_node = None
            self._fake_selection(self.TargetSchemaRoot[0])

    def rebuild_nav_tree(self):
        from data_model import NavTreeNode
        from collections import OrderedDict
        self.NavRoot.Clear()
        root_node = NavTreeNode("All Sheets", "Root", is_expanded=True)
        
        col_groups = OrderedDict()
        for sh in self.all_grid_nodes:
            c_name = getattr(sh, 'CollectionName', 'PERMIT SET') or 'PERMIT SET'
            d_name = getattr(sh, 'DisciplineName', 'Unknown') or 'Unknown'
            cg_name = getattr(sh, 'ContentGroupName', 'Uncategorized') or 'Uncategorized'
            
            if c_name not in col_groups:
                col_groups[c_name] = OrderedDict()
            if d_name not in col_groups[c_name]:
                col_groups[c_name][d_name] = OrderedDict()
            if cg_name not in col_groups[c_name][d_name]:
                col_groups[c_name][d_name][cg_name] = []
            col_groups[c_name][d_name][cg_name].append(sh)
            
        total_sheets = 0
        for c_name, disc_dict in col_groups.items():
            col_node = NavTreeNode(c_name, "Collection", tag=c_name, parent=root_node, is_expanded=True)
            col_count = 0
            for d_name, cg_dict in disc_dict.items():
                disc_node = NavTreeNode(d_name, "Discipline", tag=(c_name, d_name), parent=col_node, is_expanded=False)
                disc_count = 0
                for cg_name, sh_list in cg_dict.items():
                    cg_node = NavTreeNode(cg_name, "ContentGroup", tag=(c_name, d_name, cg_name), parent=disc_node, is_expanded=False)
                    cg_count = len(sh_list)
                    cg_node.Count = cg_count
                    disc_count += cg_count
                    disc_node.Children.Add(cg_node)
                disc_node.Count = disc_count
                col_count += disc_count
                col_node.Children.Add(disc_node)
            col_node.Count = col_count
            total_sheets += col_count
            root_node.Children.Add(col_node)
            
        root_node.Count = total_sheets
        self.NavRoot.Add(root_node)

    def on_quick_add_sheet(self, sender=None, e=None):
        if self.main_vm.SelectedSheet:
            self.on_clone_sheet(self.main_vm.SelectedSheet)
            return
        target_cell = None
        for r in self.main_vm.MatrixRows:
            if r.IsLevelRow:
                target_cell = r.CellFloorPlans
                break
        if not target_cell and self.main_vm.MatrixRows.Count > 0:
            target_cell = self.main_vm.MatrixRows[0].CellFloorPlans
        if target_cell:
            self.on_matrix_add_sheet(target_cell)

    def _fake_selection(self, node):
        class DummyArgs(object): pass
        self._current_selected_node = node
        tree_sender = self.NavTree if (hasattr(self, 'NavTree') and self.NavTree) else self.TargetSchemaTree
        self.on_tree_selection_changed(tree_sender, DummyArgs())

            
    def on_tree_selection_changed(self, sender, e):
        if getattr(self, "_ignore_tree_event", False): return
        if hasattr(e, "OriginalSource") and e.OriginalSource != sender: return
        
        if getattr(e, "__class__", None).__name__ == "DummyArgs":
            new_node = getattr(self, "_current_selected_node", None)
        else:
            new_node = getattr(self, "_current_selected_node", None)
            if hasattr(sender, "SelectedItem") and sender.SelectedItem and hasattr(sender.SelectedItem, "NodeType"):
                new_node = sender.SelectedItem
            elif hasattr(self, 'NavTree') and self.NavTree and hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem and hasattr(self.NavTree.SelectedItem, "NodeType"):
                new_node = self.NavTree.SelectedItem
            elif hasattr(self, 'TargetSchemaTree') and self.TargetSchemaTree and hasattr(self.TargetSchemaTree, "SelectedItem") and self.TargetSchemaTree.SelectedItem and hasattr(self.TargetSchemaTree.SelectedItem, "NodeType"):
                new_node = self.TargetSchemaTree.SelectedItem
                
        if not new_node: return
        
        last_node = getattr(self, "_last_selected_node", None)
        
        # If user clicks the exact same item, no need to refresh
        if last_node and new_node == last_node:
            return
            
        # Alert user about data loss if they are switching away from a previously selected node
        if last_node is not None and getattr(e, "__class__", None).__name__ != "DummyArgs":
            has_dirty = any(getattr(vm, "IsDirty", False) for vm in self.EditorItems)
            if has_dirty:
                from pyrevit import forms
                res = forms.alert("You have unsaved manual edits in the current view.\n\nSwitching collections will discard them. Do you want to continue?", title="Unsaved Changes", yes=True, no=True)
                if not res:
                    # User cancelled. Revert selection.
                    self._ignore_tree_event = True
                    last_node.IsSelected = True
                    self._ignore_tree_event = False
                    return
                
        self._last_selected_node = new_node
        self._current_selected_node = new_node
        node = new_node
        
        valid_sheets = []
        
        c_name = self.Cmb_TargetCollection.Text if hasattr(self, 'Cmb_TargetCollection') and self.Cmb_TargetCollection.Text else "PERMIT SET"
            
        if not hasattr(node, "NodeType"): return
        
        active_discipline = None
        active_cg = None
        focused_sheet_prefix = None
        
        if node.NodeType == "Root":
            valid_sheets = list(self.all_grid_nodes)
            self.Txt_GridTitle.Text = "2D Spatial Cartoon Matrix — All Collections"
        elif node.NodeType == "Collection":
            c_name = node.Tag
            self.Txt_GridTitle.Text = "2D Spatial Cartoon Matrix — Collection: " + str(c_name)
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name or s.OriginalCollectionName in ("Undefined", "Default")]
        elif node.NodeType == "Discipline":
            if isinstance(node.Tag, tuple) and len(node.Tag) == 2:
                c_name, d_name = node.Tag
                valid_sheets = [s for s in self.all_grid_nodes if (getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name or s.OriginalCollectionName in ("Undefined", "Default")) and s.DisciplineName == d_name]
            else:
                c_name = None
                d_name = str(node.Tag or node.Name)
                valid_sheets = [s for s in self.all_grid_nodes if s.DisciplineName == d_name]
            active_discipline = d_name
            if c_name:
                self.Txt_GridTitle.Text = "Collection: {} | Discipline: {}".format(c_name, d_name)
            else:
                self.Txt_GridTitle.Text = "Discipline: {}".format(d_name)
        elif node.NodeType == "ContentGroup":
            if isinstance(node.Tag, tuple) and len(node.Tag) == 3:
                c_name, d_name, cg_name = node.Tag
                valid_sheets = [s for s in self.all_grid_nodes if (getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name or s.OriginalCollectionName in ("Undefined", "Default")) and s.DisciplineName == d_name and s.ContentGroupName == cg_name]
            else:
                c_name = None
                d_name = str(node.Tag) if node.Tag else None
                cg_name = node.Name
                valid_sheets = [s for s in self.all_grid_nodes if (not d_name or s.DisciplineName == d_name) and s.ContentGroupName == cg_name]
            active_discipline = d_name
            active_cg = cg_name
            if c_name:
                self.Txt_GridTitle.Text = "Collection: {} | Group: {}".format(c_name, cg_name)
            else:
                self.Txt_GridTitle.Text = "Group: {}".format(cg_name)
        elif node.NodeType == "Sheet":
            focused_sheet_prefix = node.Name.split(" - ")[0].strip().upper()
            parent_cg = getattr(node, "Parent", None)
            if parent_cg and isinstance(getattr(parent_cg, "Tag", None), tuple) and len(parent_cg.Tag) == 3:
                c_name, d_name, cg_name = parent_cg.Tag
                active_discipline = d_name
                active_cg = cg_name
                self.Txt_GridTitle.Text = "Sheet: {}".format(node.Name)
                valid_sheets = [s for s in self.all_grid_nodes if (getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name or s.OriginalCollectionName in ("Undefined", "Default")) and s.DisciplineName == d_name and s.ContentGroupName == cg_name]
            else:
                valid_sheets = list(self.all_grid_nodes)
            
        self.update_grid_title()
        
        # Highlight active collection in TargetSchemaRoot and NavRoot
        for root_node in self.TargetSchemaRoot:
            for c_node in root_node.Children:
                c_node.IsActiveContext = (c_node.Tag == c_name)
        for root_node in self.NavRoot:
            for c_node in root_node.Children:
                c_node.IsActiveContext = (c_node.Tag == c_name)
                
        self.execute_schema_match(valid_sheets, active_collection=c_name, active_discipline=active_discipline, active_cg=active_cg)
        if focused_sheet_prefix:
            for vm in self.EditorItems:
                if (vm.SheetNumber or "").strip().upper() == focused_sheet_prefix:
                    self.on_select_sheet_vm(vm)

    def execute_schema_match(self, valid_sheets, active_collection="Undefined", active_discipline=None, active_cg=None):
        raw_targets = getattr(self, "generated_targets", [])
        include_missing = True
        if hasattr(self, 'Chk_AddMissing') and self.Chk_AddMissing is not None:
            include_missing = bool(self.Chk_AddMissing.IsChecked)
        
        # Filter generated targets based on the active node context
        generated_targets = []
        for t in raw_targets:
            if t.get("baseline_num") in self.excluded_target_sheets:
                continue
            t_copy = t.copy()
            match = re.match(r"^([A-Z]+)[- ]?(\d+)", t_copy["num"].upper())
            disc_code = match.group(1) if match else "Other"
            disc_dict = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors", "CS": "Cover Sheet" }
            disc_name = "{} - {}".format(disc_code, disc_dict.get(disc_code, "Discipline")) if disc_code != "Other" else "Uncategorized"
            cg_name = t_copy.get("cg", "Unknown Content Group")
            
            if active_discipline and disc_name != active_discipline:
                continue
            if active_cg and cg_name != active_cg:
                continue
            
            t_copy["collection"] = active_collection
            t_copy["disc_name"] = disc_name
            t_copy["cg_name"] = cg_name
            generated_targets.append(t_copy)
            
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        if not doc: return
        
        try:
            import reconciliation
            
            valid_sheet_ids = []
            for s in valid_sheets:
                if hasattr(s.ElementId, "IntegerValue"):
                    valid_sheet_ids.append(s.ElementId.IntegerValue)
                elif hasattr(s.ElementId, "Value"):
                    valid_sheet_ids.append(s.ElementId.Value)
                    
            try:
                plan = reconciliation.run_pipeline(
                    doc,
                    generated_targets,
                    existing_sheet_ids=valid_sheet_ids,
                    h_data=getattr(self, "_cached_harvest_data", None),
                    include_missing_slots=include_missing
                )
            except Exception as e:
                self.Txt_GridTitle.Text = "ERROR IN PIPELINE: " + str(e)
                forms.alert("ERROR IN PIPELINE: " + str(e))
                plan = {"rows": []}
            
            from Autodesk.Revit.DB import ElementId
            # Rule 6: Keep self.all_grid_nodes as a strictly read-only baseline!
            baseline_by_id = {}
            for n in self.all_grid_nodes:
                if n.ElementId and n.ElementId != ElementId.InvalidElementId:
                    try:
                        key = n.ElementId.IntegerValue if hasattr(n.ElementId, "IntegerValue") else n.ElementId.Value
                        baseline_by_id[key] = n
                    except Exception:
                        pass
                        
            self.EditorItems.Clear()
            
            for row in plan["rows"]:
                is_template = (row["status"] == "MISSING")
                sh_id = row["sheet_element_id"]
                
                if sh_id in baseline_by_id and sh_id != -1:
                    # Clone from read-only baseline so edits in EditorItems never corrupt all_grid_nodes
                    vm = baseline_by_id[sh_id].clone_for_sandbox(
                        validation_callback=self.run_validation,
                        number_changed_callback=self.on_sheet_number_changed,
                        move_up_callback=self.move_item_up,
                        move_down_callback=self.move_item_down,
                        context_callback=self.get_sheet_context,
                        clone_sheet_callback=self.on_clone_sheet,
                        auto_populate_callback=self.auto_populate_sheet_views,
                        select_sheet_callback=self.on_select_sheet_vm,
                        accept_suggestion_callback=self.on_accept_sheet_suggestion
                    )
                    vm.SheetSeries = row.get("series_name", "Unknown")
                    row_disc = row.get("discipline")
                    if row_disc and row_disc != "Unknown":
                        vm.DisciplineName = row_disc
                    elif not vm.DisciplineName or vm.DisciplineName == "Unknown":
                        vm.reclassify()

                    row_cg = row.get("cg")
                    if row_cg and row_cg != "Uncategorized":
                        vm.ContentGroupName = row_cg
                    
                    # Always record AIA Rule Book suggestions
                    sug_num = row.get("target_number", "")
                    sug_name = row.get("target_name", "")
                    vm.SuggestedNumber = sug_num
                    vm.SuggestedName = sug_name
                    vm.SuggestedSeries = row.get("series_name", "01. Plans")

                    vm._sheet_number = sug_num if sug_num else vm.OriginalNumber
                    vm._sheet_name = sug_name if sug_name else vm.OriginalName
                    vm._is_user_modified = False
                    vm._is_checked = False
                    vm.update_diffs()
                    vm.update_action()
                    vm.notify_reconciliation_state()
                    
                    # Re-check semantic mismatches on placed views against the proposed sheet name
                    for v_item in vm.Views:
                        v_item.check_semantic_mismatch()
                    self.EditorItems.Add(vm)
                else:
                    tgt_num = row["target_number"] if is_template else row["existing_number"]
                    real_id = ElementId(sh_id) if sh_id != -1 else ElementId.InvalidElementId
                    vm = SheetViewModel(
                        real_id, tgt_num,
                        row["target_name"] if is_template else row["existing_name"],
                        active_collection, row["discipline"], row["cg"],
                        series_name=row.get("series_name", "Unknown"),
                        is_template=is_template,
                        validation_callback=self.run_validation,
                        number_changed_callback=self.on_sheet_number_changed,
                        move_up_callback=self.move_item_up,
                        move_down_callback=self.move_item_down,
                        context_callback=self.get_sheet_context,
                        clone_sheet_callback=self.on_clone_sheet,
                        auto_populate_callback=self.auto_populate_sheet_views,
                        select_sheet_callback=self.on_select_sheet_vm,
                        accept_suggestion_callback=self.on_accept_sheet_suggestion
                    )
                    vm.OriginalCollectionName = active_collection
                    vm.IsChecked = True
                    vm._action = "CREATE" if row["status"] in ["CREATE", "MISSING"] else row["status"]
                    # Auto-populate a smart view slot on newly generated Cartoon Set sheets
                    self.auto_populate_sheet_views(vm, run_val=False)
                    vm.notify_reconciliation_state()
                    self.EditorItems.Add(vm)
            
            for vm in self.EditorItems:
                vm.IsDirty = False
            
            self.update_discipline_filter_options()
            
            self.build_matrix_board()
            self.update_grid_title()
            self.run_validation()
            
        except Exception as big_e:
            import traceback
            forms.alert("CRITICAL CRASH IN SCHEMA MATCH:\n" + traceback.format_exc())

    def _classify_sheet_into_matrix_coords(self, sheet_vm, level_names_ordered):
        """
        Determines the exact (row_key, col_key) in the 2D Building x Series Matrix for `sheet_vm`.
        """
        sh_name_up = (sheet_vm.SheetName or "").upper()
        series_str = (sheet_vm.SheetSeries or "").upper()
        cg_up = (sheet_vm.ContentGroupName or "").upper()

        # 1. Determine Column Key (AIA Series)
        col_key = "FLOOR_PLANS"
        if "00." in series_str or "COVER" in sh_name_up or "GENERAL" in sh_name_up or "INDEX" in sh_name_up or "SYMBOL" in sh_name_up:
            col_key = "GENERAL"
        elif "02." in series_str or "ELEVATION" in sh_name_up:
            col_key = "ELEVATIONS"
        elif "03." in series_str or "SECTION" in sh_name_up:
            col_key = "SECTIONS"
        elif "05." in series_str or "06." in series_str or "07." in series_str or "08." in series_str or "09." in series_str or "DETAIL" in sh_name_up or "SCHEDULE" in sh_name_up or "3D" in sh_name_up:
            col_key = "DETAILS_SCHEDULES"
        elif "04." in series_str or "ENLARGED" in sh_name_up or "QUADRANT" in sh_name_up or "SECTOR" in sh_name_up or "PARTIAL" in sh_name_up:
            col_key = "ENLARGED_PLANS"
        elif "CEILING" in sh_name_up or "RCP" in sh_name_up:
            col_key = "CEILING_PLANS"
        else:
            col_key = "FLOOR_PLANS"

        # 2. Determine Spatial Row Key (Level / Roof / Foundation / Site / Project-Wide)
        # Check if sheet explicitly mentions a project Level
        matched_lvl = None
        for lvl_name in level_names_ordered:
            if lvl_name.upper() in sh_name_up:
                matched_lvl = lvl_name
                break
        if not matched_lvl and sheet_vm.Views.Count > 0:
            v_lvl = sheet_vm.Views[0].LevelName
            if v_lvl and v_lvl in level_names_ordered:
                matched_lvl = v_lvl
        if not matched_lvl:
            m_lvl = re.search(r'LEVEL\s*0*(\d+)|(\d+)(ST|ND|RD|TH)\s*FLOOR', sh_name_up)
            if m_lvl:
                num_str = m_lvl.group(1) or m_lvl.group(2)
                for lvl_name in level_names_ordered:
                    if re.search(r'0*' + re.escape(num_str) + r'\b', lvl_name):
                        matched_lvl = lvl_name
                        break

        if matched_lvl:
            return ("LEVEL::" + matched_lvl, col_key)

        if "SITE" in sh_name_up or "LANDSCAPE" in sh_name_up or "CIVIL" in cg_up:
            return ("SCOPE::SITE", col_key)
        if "FOUNDATION" in sh_name_up or "BASEMENT" in sh_name_up or "FOOTING" in sh_name_up:
            return ("SCOPE::FOUNDATION", col_key)
        if "ROOF" in sh_name_up:
            # If there is a Roof level in Revit, map to it; else SCOPE::ROOF
            for lvl_name in level_names_ordered:
                if "roof" in lvl_name.lower():
                    return ("LEVEL::" + lvl_name, col_key)
            return ("SCOPE::ROOF", col_key)

        return ("SCOPE::PROJECT", col_key)

    def build_matrix_board(self):
        """
        Projects `self.EditorItems` into the 2D `self.main_vm.MatrixRows` board and populates `self.main_vm.UnreconciledItems`.
        """
        import data_model
        from data_model import MatrixRowViewModel

        show_invalid_only = False
        if hasattr(self, 'Chk_ShowInvalid') and self.Chk_ShowInvalid and self.Chk_ShowInvalid.IsChecked:
            show_invalid_only = True

        avail_levels = list(getattr(data_model, "AVAILABLE_LEVELS", []))
        if not avail_levels:
            avail_levels = [n.Name for n in self.LevelNodes]

        # Build rows in Top-to-Bottom Building Section order:
        # 1. Project / Building-Wide
        # 2. Roof (if no explicit Roof level)
        # 3. Building Levels in descending elevation order (top floor -> ground floor)
        # 4. Foundation / Basement (if no explicit Foundation level)
        # 5. Site / Civil
        row_Lookup = {}
        ordered_rows = []

        def _get_or_create_row(r_key, r_title, r_sub, is_lvl=True):
            if r_key not in row_Lookup:
                r_vm = MatrixRowViewModel(
                    r_key, r_title, r_sub, is_level_row=is_lvl,
                    add_sheet_callback=self.on_matrix_add_sheet,
                    split_sectors_callback=self.on_matrix_split_sectors
                )
                row_Lookup[r_key] = r_vm
                ordered_rows.append(r_vm)
            return row_Lookup[r_key]

        _get_or_create_row("SCOPE::PROJECT", u"🏢 Project / Building-Wide", "Cover, Elevations, Sections, Details", is_lvl=False)
        has_roof_lvl = any("roof" in l.lower() for l in avail_levels)
        if not has_roof_lvl:
            _get_or_create_row("SCOPE::ROOF", u"🏔️ Roof", "Roof Plans & Details", is_lvl=True)

        # Add Revit Levels top-to-bottom (reversed elevation order)
        lvl_display_map = {n.Name: n.DisplayName for n in self.LevelNodes}
        for lvl_name in reversed(avail_levels):
            disp = lvl_display_map.get(lvl_name, lvl_name)
            elev_sub = ""
            if "(" in disp and ")" in disp:
                elev_sub = "Elev " + disp[disp.rfind("("):]
            else:
                elev_sub = "Building Story"
            icon = u"🏔️ " if "roof" in lvl_name.lower() else (u"🧱 " if "found" in lvl_name.lower() or "base" in lvl_name.lower() else u"📐 ")
            _get_or_create_row("LEVEL::" + lvl_name, icon + lvl_name, elev_sub, is_lvl=True)

        has_found_lvl = any("found" in l.lower() or "base" in l.lower() for l in avail_levels)
        if not has_found_lvl:
            _get_or_create_row("SCOPE::FOUNDATION", u"🧱 Foundation / Substructure", "Foundation & Slab Plans", is_lvl=True)
        _get_or_create_row("SCOPE::SITE", u"🌲 Site / Civil", "Site & Context Plans", is_lvl=True)

        self.main_vm.UnreconciledItems.Clear()

        for sh_vm in self.EditorItems:
            if show_invalid_only and sh_vm.Action == "MATCHED" and not sh_vm.HasWarning:
                continue

            if sh_vm.Action in ("UNRECONCILED", "EXTRA") or sh_vm.HasSuggestion:
                self.main_vm.UnreconciledItems.Add(sh_vm)

            r_key, c_key = self._classify_sheet_into_matrix_coords(sh_vm, avail_levels)
            target_row = row_Lookup.get(r_key) or row_Lookup["SCOPE::PROJECT"]

            cell_map = {
                "FLOOR_PLANS": target_row.CellFloorPlans,
                "CEILING_PLANS": target_row.CellCeilingPlans,
                "ENLARGED_PLANS": target_row.CellEnlargedPlans,
                "GENERAL": target_row.CellGeneral,
                "ELEVATIONS": target_row.CellElevations,
                "SECTIONS": target_row.CellSections,
                "DETAILS_SCHEDULES": target_row.CellDetails,
            }
            target_cell = cell_map.get(c_key, target_row.CellFloorPlans)
            target_cell.add_sheet_vm(sh_vm)

        self.main_vm.MatrixRows.Clear()
        for r_vm in ordered_rows:
            r_vm.OnPropertyChanged("RowBadge")
            self.main_vm.MatrixRows.Add(r_vm)

        self.main_vm.OnPropertyChanged("HasUnreconciledItems")
        self.main_vm.OnPropertyChanged("UnreconciledCountText")

        # Auto-select first sheet in Inspector if none selected
        if (not self.main_vm.SelectedSheet or self.main_vm.SelectedSheet not in self.EditorItems) and self.EditorItems.Count > 0:
            self.on_select_sheet_vm(self.EditorItems[0])

        # Update top Matrix KPI summary
        create_cnt = sum(1 for s in self.EditorItems if s.Action in ("CREATE", "MISSING") and s.IsChecked)
        match_cnt = sum(1 for s in self.EditorItems if s.Action == "MATCHED")
        update_cnt = sum(1 for s in self.EditorItems if "RENAME" in s.Action or s.Action == "UPDATE")
        messy_cnt = self.main_vm.UnreconciledItems.Count
        self.main_vm.MatrixSummaryText = u"{} Total Sheets  ·  {} Planned Create  ·  {} Matched  ·  {} Updates  ·  {} Messy".format(
            self.EditorItems.Count, create_cnt, match_cnt, update_cnt, messy_cnt
        )

    def on_select_sheet_vm(self, sheet_vm):
        if not sheet_vm:
            return
        self.main_vm.SelectedSheet = sheet_vm

    def on_editor_row_clicked(self, sender, e):
        if hasattr(sender, "DataContext") and sender.DataContext:
            self.on_select_sheet_vm(sender.DataContext)

    def on_accept_sheet_suggestion(self, sheet_vm):
        """
        Accepts the proposal or user modifications for a sheet,
        marking it ready to push to Revit.
        """
        if not sheet_vm:
            return
        if not getattr(sheet_vm, "_is_user_modified", False) and (sheet_vm.SheetNumber == sheet_vm.OriginalNumber) and (sheet_vm.SheetName == sheet_vm.OriginalName):
            if sheet_vm.SuggestedNumber:
                sheet_vm.SheetNumber = sheet_vm.SuggestedNumber
            if sheet_vm.SuggestedName:
                sheet_vm.SheetName = sheet_vm.SuggestedName
        if sheet_vm.SuggestedSeries and sheet_vm.SheetSeries in ("Unknown", ""):
            sheet_vm.SheetSeries = sheet_vm.SuggestedSeries
        sheet_vm.Action = "CREATE" if sheet_vm.IsTemplate else "UPDATE"
        sheet_vm.IsChecked = True
        sheet_vm.notify_reconciliation_state()
        self.build_matrix_board()
        self.on_select_sheet_vm(sheet_vm)
        self.run_validation()

    def on_auto_fix_all_messy_sheets(self):
        """
        Applies suggestions or confirms modifications for all candidate sheets.
        """
        candidates = [s for s in self.EditorItems if s.CanAcceptProposal or s.HasSuggestion]
        if not candidates:
            return
        for sh_vm in candidates:
            if not getattr(sh_vm, "_is_user_modified", False) and (sh_vm.SheetNumber == sh_vm.OriginalNumber) and (sh_vm.SheetName == sh_vm.OriginalName):
                if sh_vm.SuggestedNumber:
                    sh_vm.SheetNumber = sh_vm.SuggestedNumber
                if sh_vm.SuggestedName:
                    sh_vm.SheetName = sh_vm.SuggestedName
            if sh_vm.SuggestedSeries and sh_vm.SheetSeries in ("Unknown", ""):
                sh_vm.SheetSeries = sh_vm.SuggestedSeries
            sh_vm.Action = "CREATE" if sh_vm.IsTemplate else "UPDATE"
            sh_vm.IsChecked = True
            sh_vm.notify_reconciliation_state()
        self.build_matrix_board()
        self.run_validation()

    def _next_available_aia_number(self, disc_code, series_digit):
        used_nums = set((s.SheetNumber or "").strip().upper() for s in self.EditorItems)
        seq = 1
        while True:
            cand = "{}-{}{:02d}".format(disc_code, series_digit, seq)
            if cand.upper() not in used_nums:
                return cand
            seq += 1

    def on_matrix_add_sheet(self, cell_vm):
        """
        Direct 1-Click Cell Action (`+ Floor Plan`, `+ Ceiling Plan`, `+ Elevation`, etc.):
        Creates a properly numbered & named AIA sheet directly inside the clicked [Level, Series] Matrix cell!
        """
        if not cell_vm:
            return
        from Autodesk.Revit.DB import ElementId

        # If cell already has a sheet, clone the last sheet in that cell (+1)
        if cell_vm.Sheets.Count > 0:
            self.on_clone_sheet(cell_vm.Sheets[cell_vm.Sheets.Count - 1])
            return

        row_key = cell_vm.RowKey
        col_key = cell_vm.ColKey
        lvl_label = ""
        if row_key.startswith("LEVEL::"):
            lvl_label = row_key.replace("LEVEL::", "")
        elif row_key == "SCOPE::ROOF":
            lvl_label = "ROOF"
        elif row_key == "SCOPE::FOUNDATION":
            lvl_label = "FOUNDATION"
        elif row_key == "SCOPE::SITE":
            lvl_label = "SITE"

        col_config = {
            "FLOOR_PLANS": ("1", "01. Plans", "10 - General Floor Plans", "FLOOR PLAN - {}".format(lvl_label.upper()) if lvl_label and lvl_label not in ("SITE", "FOUNDATION", "ROOF") else ("{} PLAN".format(lvl_label.upper()) if lvl_label else "FLOOR PLAN")),
            "CEILING_PLANS": ("1", "01. Plans", "15 - Reflected Ceiling Plans", "REFLECTED CEILING PLAN - {}".format(lvl_label.upper()) if lvl_label else "REFLECTED CEILING PLAN"),
            "ENLARGED_PLANS": ("4", "04. Large Scale Views", "40 - Enlarged Plans & Callouts", "ENLARGED FLOOR PLAN - {}".format(lvl_label.upper()) if lvl_label else "ENLARGED FLOOR PLAN"),
            "GENERAL": ("0", "00. General", "01 - Cover, Title & Index", "GENERAL NOTES & SYMBOLS"),
            "ELEVATIONS": ("2", "02. Elevations", "20 - Exterior Elevations", "EXTERIOR ELEVATIONS" if not lvl_label else "INTERIOR ELEVATIONS - {}".format(lvl_label.upper())),
            "SECTIONS": ("3", "03. Sections", "30 - Building Sections", "BUILDING SECTIONS"),
            "DETAILS_SCHEDULES": ("5", "05. Details", "50 - Exterior & Envelope Details", "ARCHITECTURAL DETAILS"),
        }
        s_digit, series_name, cg_name, sh_name = col_config.get(col_key, ("1", "01. Plans", "10 - General Floor Plans", "FLOOR PLAN"))
        sh_num = self._next_available_aia_number("A", s_digit)
        active_col = self.Cmb_TargetCollection.Text if hasattr(self, 'Cmb_TargetCollection') and self.Cmb_TargetCollection.Text else "PERMIT SET"

        new_vm = SheetViewModel(
            ElementId.InvalidElementId,
            sh_num,
            sh_name,
            active_col,
            discipline_name="A - Architectural",
            content_group_name=cg_name,
            series_name=series_name,
            is_template=True,
            validation_callback=self.run_validation,
            number_changed_callback=self.on_sheet_number_changed,
            move_up_callback=self.move_item_up,
            move_down_callback=self.move_item_down,
            context_callback=self.get_sheet_context,
            clone_sheet_callback=self.on_clone_sheet,
            auto_populate_callback=self.auto_populate_sheet_views,
            select_sheet_callback=self.on_select_sheet_vm,
            accept_suggestion_callback=self.on_accept_sheet_suggestion
        )
        new_vm.OriginalCollectionName = active_col
        new_vm.IsChecked = True
        new_vm._action = "CREATE"
        self.auto_populate_sheet_views(new_vm, run_val=False)
        self.EditorItems.Add(new_vm)
        self.build_matrix_board()
        self.on_select_sheet_vm(new_vm)
        self.run_validation()

    def on_matrix_split_sectors(self, cell_vm):
        """
        One-Click `⊞ 2x2` Sector Split on any Plan cell:
        Splits (or un-splits) that specific Level x Plan cell into 4 Sector Sheets (A, B, C, D)
        without causing combinatorial explosion across other levels!
        """
        if not cell_vm:
            return
        from Autodesk.Revit.DB import ElementId

        existing_in_cell = list(cell_vm.Sheets)
        # If already split into >= 4 sector sheets, collapse back to 1 base sheet
        if len(existing_in_cell) >= 4:
            first_sh = existing_in_cell[0]
            for extra_sh in existing_in_cell[1:]:
                if extra_sh.IsTemplate and extra_sh in self.EditorItems:
                    self.EditorItems.Remove(extra_sh)
            first_sh.SheetNumber = re.sub(r'[A-D]$', '', first_sh.SheetNumber)
            first_sh.SheetName = re.sub(r'\s*-\s*(SECTOR|QUADRANT)\s+[A-D]$', '', first_sh.SheetName, flags=re.IGNORECASE)
            self.build_matrix_board()
            self.on_select_sheet_vm(first_sh)
            self.run_validation()
            return

        # Ensure at least 1 base sheet exists in the cell
        if not existing_in_cell:
            self.on_matrix_add_sheet(cell_vm)
            existing_in_cell = list(cell_vm.Sheets)
            if not existing_in_cell:
                return

        base_sh = existing_in_cell[0]
        base_num = re.sub(r'[A-Za-z]$', '', (base_sh.SheetNumber or "A-101").strip())
        base_name = re.sub(r'\s*-\s*(SECTOR|QUADRANT)\s+[A-D]$', '', (base_sh.SheetName or "FLOOR PLAN").strip(), flags=re.IGNORECASE)

        base_sh.SheetNumber = base_num + "A"
        base_sh.SheetName = base_name + " - SECTOR A"
        if base_sh.Views.Count > 0:
            base_sh.Views[0].Name = base_sh.SheetName

        idx = self.EditorItems.IndexOf(base_sh)
        for offset, sector_char in enumerate(["B", "C", "D"], start=1):
            sec_num = base_num + sector_char
            sec_name = "{} - SECTOR {}".format(base_name, sector_char)
            sec_vm = SheetViewModel(
                ElementId.InvalidElementId,
                sec_num,
                sec_name,
                base_sh.CollectionName,
                discipline_name=base_sh.DisciplineName,
                content_group_name=base_sh.ContentGroupName,
                series_name=base_sh.SheetSeries,
                is_template=True,
                validation_callback=self.run_validation,
                number_changed_callback=self.on_sheet_number_changed,
                move_up_callback=self.move_item_up,
                move_down_callback=self.move_item_down,
                context_callback=self.get_sheet_context,
                clone_sheet_callback=self.on_clone_sheet,
                auto_populate_callback=self.auto_populate_sheet_views,
                select_sheet_callback=self.on_select_sheet_vm,
                accept_suggestion_callback=self.on_accept_sheet_suggestion
            )
            sec_vm.OriginalCollectionName = base_sh.CollectionName
            sec_vm.IsChecked = True
            sec_vm._action = "CREATE"
            self.auto_populate_sheet_views(sec_vm, run_val=False)
            if idx >= 0 and idx + offset <= self.EditorItems.Count:
                self.EditorItems.Insert(idx + offset, sec_vm)
            else:
                self.EditorItems.Add(sec_vm)

        self.build_matrix_board()
        self.on_select_sheet_vm(base_sh)
        self.run_validation()

    def on_fill_floor_plans_column(self):
        for r_vm in list(self.main_vm.MatrixRows):
            if r_vm.IsLevelRow and r_vm.CellFloorPlans.IsEmpty:
                self.on_matrix_add_sheet(r_vm.CellFloorPlans)

    def on_fill_ceiling_plans_column(self):
        for r_vm in list(self.main_vm.MatrixRows):
            if r_vm.IsLevelRow and r_vm.RowKey.startswith("LEVEL::") and r_vm.CellCeilingPlans.IsEmpty:
                self.on_matrix_add_sheet(r_vm.CellCeilingPlans)

    def on_clone_sheet(self, sheet_vm):
        """
        Interactive Cartoon Set Builder (+1 button on any Sheet Pill or Inspector):
        Inserts an additional sheet in the same Discipline, Collection, and AIA Series immediately after `sheet_vm`,
        auto-incrementing the sheet number (e.g. A-201 -> A-202) and proposing a matching view slot.
        """
        if not sheet_vm:
            return
        from Autodesk.Revit.DB import ElementId
        base_num = (sheet_vm.SheetNumber or "A-101").strip()
        used_nums = set((s.SheetNumber or "").strip().upper() for s in self.EditorItems if s.CollectionName == sheet_vm.CollectionName)

        m = re.match(r"^([A-Za-z]+[- ]?\d)(\d+)([A-Za-z]*)$", base_num)
        if m:
            prefix, digits, suffix = m.group(1), m.group(2), m.group(3)
            pad_len = len(digits)
            next_val = int(digits) + 1
            candidate = "{}{:0{}d}{}".format(prefix, next_val, pad_len, suffix)
            while candidate.upper() in used_nums:
                next_val += 1
                candidate = "{}{:0{}d}{}".format(prefix, next_val, pad_len, suffix)
        else:
            idx_try = 2
            candidate = "{}-{}".format(base_num, idx_try)
            while candidate.upper() in used_nums:
                idx_try += 1
                candidate = "{}-{}".format(base_num, idx_try)

        new_vm = SheetViewModel(
            ElementId.InvalidElementId,
            candidate,
            sheet_vm.SheetName,
            sheet_vm.CollectionName,
            discipline_name=sheet_vm.DisciplineName,
            content_group_name=sheet_vm.ContentGroupName,
            series_name=sheet_vm.SheetSeries,
            is_template=True,
            validation_callback=self.run_validation,
            number_changed_callback=self.on_sheet_number_changed,
            move_up_callback=self.move_item_up,
            move_down_callback=self.move_item_down,
            context_callback=self.get_sheet_context,
            clone_sheet_callback=self.on_clone_sheet,
            auto_populate_callback=self.auto_populate_sheet_views,
            select_sheet_callback=self.on_select_sheet_vm,
            accept_suggestion_callback=self.on_accept_sheet_suggestion
        )
        new_vm.OriginalCollectionName = sheet_vm.CollectionName
        new_vm.IsChecked = True
        new_vm._action = "CREATE"
        self.auto_populate_sheet_views(new_vm, run_val=False)

        idx = self.EditorItems.IndexOf(sheet_vm)
        if idx >= 0 and idx < self.EditorItems.Count - 1:
            self.EditorItems.Insert(idx + 1, new_vm)
        else:
            self.EditorItems.Add(new_vm)
        self.build_matrix_board()
        self.on_select_sheet_vm(new_vm)
        self.run_validation()

    def auto_populate_sheet_views(self, sheet_vm, run_val=True):
        """
        Inspects an empty or template sheet and:
        1. Searches unplaced project views in `AVAILABLE_ASSIGNABLE_VIEWS` for the best semantic match (by Level + ViewType + Name).
        2. If a strong unplaced view exists (and isn't already assigned in `EditorItems`), links it as an `Existing` view!
        3. Otherwise proposes a `New` View with inferred Level, ViewFamilyType, standard AIA Scale, Detail Number '1',
           Graphic Scale, and North Arrow (for Plan views).
        """
        if not sheet_vm:
            return
        # Skip Series 0 General / Cover sheets unless the user explicitly clicked the button on that sheet
        series_str = (sheet_vm.SheetSeries or "").lower()
        sh_name_upper = (sheet_vm.SheetName or "").upper()
        if not run_val and ("00." in series_str or "cover" in sh_name_upper or "index" in sh_name_upper or "general" in sh_name_upper or "schedule" in sh_name_upper):
            return

        if sheet_vm.Views.Count > 0 and not run_val:
            return

        import data_model
        from Autodesk.Revit.DB import ElementId

        # Collect already-claimed ViewIds across EditorItems so we don't assign the same view twice
        claimed_view_ids = set()
        for s in self.EditorItems:
            for v in s.Views:
                if v.ViewId and v.ViewId != ElementId.InvalidElementId:
                    val = v.ViewId.IntegerValue if hasattr(v.ViewId, "IntegerValue") else v.ViewId.Value
                    claimed_view_ids.add(val)

        # Infer target Level from SheetName
        inferred_level = ""
        avail_levels = getattr(data_model, "AVAILABLE_LEVELS", [])
        for lvl in avail_levels:
            if lvl.upper() in sh_name_upper:
                inferred_level = lvl
                break
        if not inferred_level:
            m_lvl = re.search(r'LEVEL\s*0*(\d+)|(\d+)(ST|ND|RD|TH)\s*FLOOR', sh_name_upper)
            if m_lvl:
                num_str = m_lvl.group(1) or m_lvl.group(2)
                for lvl in avail_levels:
                    if re.search(r'0*' + re.escape(num_str) + r'\b', lvl):
                        inferred_level = lvl
                        break
        if not inferred_level and avail_levels:
            if "ROOF" in sh_name_upper:
                roof_cands = [l for l in avail_levels if "roof" in l.lower()]
                inferred_level = roof_cands[-1] if roof_cands else avail_levels[-1]
            elif "FOUNDATION" in sh_name_upper or "BASEMENT" in sh_name_upper:
                inferred_level = avail_levels[0]
            else:
                inferred_level = avail_levels[0]

        # Infer desired ViewFamilyType & standard Scale
        avail_vfts = getattr(data_model, "AVAILABLE_VIEW_FAMILY_TYPES", [])
        inferred_vft = avail_vfts[0] if avail_vfts else "Floor Plan"
        inferred_scale = "1/8\" = 1'-0\""
        want_ceiling = "CEILING" in sh_name_upper or "RCP" in sh_name_upper
        want_section = "SECTION" in sh_name_upper or "03." in series_str
        want_elev = "ELEVATION" in sh_name_upper or "02." in series_str
        want_detail = "DETAIL" in sh_name_upper or "05." in series_str
        want_enlarged = "ENLARGED" in sh_name_upper or "04." in series_str

        for vft_name in avail_vfts:
            vl = vft_name.lower()
            if want_ceiling and "ceiling" in vl:
                inferred_vft = vft_name
                break
            elif want_section and "section" in vl:
                inferred_vft = vft_name
                break
            elif want_elev and "elevation" in vl:
                inferred_vft = vft_name
                break
            elif want_detail and ("detail" in vl or "drafting" in vl):
                inferred_vft = vft_name
                break
            elif not (want_ceiling or want_section or want_elev or want_detail) and "floor" in vl:
                inferred_vft = vft_name
                break

        if want_enlarged:
            inferred_scale = "1/4\" = 1'-0\""
        elif want_detail:
            inferred_scale = "1 1/2\" = 1'-0\""
        elif want_elev or want_section:
            inferred_scale = "1/4\" = 1'-0\""
        elif "SITE" in sh_name_upper:
            inferred_scale = "1/16\" = 1'-0\""

        # Search unplaced views in AVAILABLE_ASSIGNABLE_VIEWS for a match
        best_av = None
        best_score = 0.0
        import reconciliation
        norm_target = reconciliation.normalize_name(sh_name_upper)

        for av in getattr(data_model, "AVAILABLE_ASSIGNABLE_VIEWS", []):
            if getattr(av, "IsPlaced", False):
                continue
            av_vid = getattr(av, "ViewId", getattr(av, "Id", None))
            if not av_vid:
                continue
            av_id_val = av_vid.IntegerValue if hasattr(av_vid, "IntegerValue") else getattr(av_vid, "Value", av_vid)
            if av_id_val in claimed_view_ids:
                continue
            vt_low = (getattr(av, "ViewType", "") or "").lower()
            if want_ceiling and "ceiling" not in vt_low:
                continue
            if not want_ceiling and "ceiling" in vt_low:
                continue
            if want_elev and "elevation" not in vt_low:
                continue
            if want_section and "section" not in vt_low:
                continue

            v_name = getattr(av, "ViewName", getattr(av, "Name", ""))
            score = reconciliation.token_set_ratio(norm_target, reconciliation.normalize_name(v_name))
            lvl_name = getattr(av, "LevelName", "") or ""
            if inferred_level and lvl_name and lvl_name.lower() == inferred_level.lower():
                score += 0.45
            if score > best_score:
                best_score = score
                best_av = av

        next_dn = str(sheet_vm.Views.Count + 1)
        is_plan = not (want_elev or want_section or want_detail)

        if best_av and best_score >= 0.65:
            best_vid = getattr(best_av, "ViewId", getattr(best_av, "Id", ElementId.InvalidElementId))
            new_v = ViewViewModel(
                best_vid,
                sheet_vm.SheetName,
                getattr(best_av, "ViewType", None) or inferred_vft,
                level_name=getattr(best_av, "LevelName", None) or inferred_level,
                scale=getattr(best_av, "ScaleStr", None) or inferred_scale,
                view_number=next_dn,
                is_new=False,
                parent_sheet=sheet_vm,
                has_graphic_scale=True,
                has_north_arrow=is_plan
            )
            new_v._selected_existing_view = best_av
        else:
            new_v = ViewViewModel(
                ElementId.InvalidElementId,
                sheet_vm.SheetName,
                inferred_vft,
                level_name=inferred_level,
                scale=inferred_scale,
                view_number=next_dn,
                is_new=True,
                parent_sheet=sheet_vm,
                has_graphic_scale=True,
                has_north_arrow=is_plan
            )

        sheet_vm.Views.Add(new_v)
        sheet_vm.IsExpanded = True
        sheet_vm.OnPropertyChanged("HasNoViews")
        sheet_vm.OnPropertyChanged("ViewsSummaryBadge")
        sheet_vm.OnPropertyChanged("MatrixPillSummary")
        sheet_vm.update_action()
        sheet_vm.notify_reconciliation_state()
        if run_val:
            self.run_validation()

    def on_auto_populate_all_views(self, sender, e):
        """
        Toolbar action (`Btn_AutoSuggestViews`):
        Populates missing views across all empty sheets currently in the workbench and standardizes View Names to match Sheet Names.
        """
        populated = 0
        for sh in self.EditorItems:
            series_str = (sh.SheetSeries or "").lower()
            sh_name_upper = (sh.SheetName or "").upper()
            if "00." in series_str or "cover" in sh_name_upper or "index" in sh_name_upper or "schedule" in sh_name_upper:
                continue
            if sh.Views.Count == 0:
                self.auto_populate_sheet_views(sh, run_val=False)
                populated += 1
            elif sh.Views.Count == 1:
                # Suggest matching the single view's Title on Sheet to the standardized Sheet Name
                v = sh.Views[0]
                if v.Name != sh.SheetName:
                    v.Name = sh.SheetName
                    sh.update_action()
                    sh.notify_reconciliation_state()
                    populated += 1
            sh.OnPropertyChanged("MatrixPillSummary")
        self.build_matrix_board()
        self.run_validation()
            
    def get_sheet_context(self):
        append = True
        if hasattr(self, 'Chk_FillGaps') and getattr(self.Chk_FillGaps, 'IsChecked'):
            append = False
            
        all_sh = list(self.all_grid_nodes)
        if hasattr(self, 'EditorItems'):
            for item in self.EditorItems:
                if item not in all_sh:
                    all_sh.append(item)
                    
        is_100_based = False
        try:
            if hasattr(self, '_loaded_naming_schemes') and self._loaded_naming_schemes:
                for b_name, b_data in self._loaded_naming_schemes.get("disciplines", {}).items():
                    for s_code, s_data in b_data.get("series", {}).items():
                        if len(s_data.get("models", [])) > 10:
                            is_100_based = True
                            break
                    if is_100_based: break
        except: pass
        
        return {
            "all_sheets": all_sh,
            "append_sequence": append,
            "is_100_based": is_100_based
        }

    def move_item_up(self, item):
        items_to_move = [item]
        
        for itm in items_to_move:
            idx = self.EditorItems.IndexOf(itm)
            if idx > 0:
                self.EditorItems.Move(idx, idx - 1)
                
    def move_item_down(self, item):
        items_to_move = [item]
        
        for itm in reversed(items_to_move):
            idx = self.EditorItems.IndexOf(itm)
            if idx < len(self.EditorItems) - 1:
                self.EditorItems.Move(idx, idx + 1)
            
    def update_grid_title(self):
        node = getattr(self, "_current_selected_node", None)
        if not node:
            if hasattr(self, 'NavTree') and self.NavTree and hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
                node = self.NavTree.SelectedItem
            elif hasattr(self, 'TargetSchemaTree') and self.TargetSchemaTree and hasattr(self.TargetSchemaTree, "SelectedItem") and self.TargetSchemaTree.SelectedItem:
                node = self.TargetSchemaTree.SelectedItem
        if not node: return
        
        if node.NodeType == "Root":
            title = "All Sheets & Views"
        elif node.NodeType == "Collection":
            title = "Collection: " + str(node.Tag or node.Name)
        elif node.NodeType == "Discipline":
            if isinstance(node.Tag, tuple) and len(node.Tag) == 2:
                c_name, d_name = node.Tag
                title = "Collection: {} | Discipline: {}".format(c_name, d_name)
            else:
                title = "Discipline: {}".format(node.Tag or node.Name)
        elif node.NodeType == "ContentGroup":
            if isinstance(node.Tag, tuple) and len(node.Tag) == 3:
                c_name, d_name, cg_name = node.Tag
                title = "Collection: {} | Group: {}".format(c_name, cg_name)
            elif isinstance(node.Tag, tuple) and len(node.Tag) == 2:
                c_name, d_name = node.Tag
                title = "Group: {}".format(node.Name)
            else:
                title = "Group: {}".format(node.Name)
        else:
            title = "Filtered Results"
            
        # Append diagnostic count
        title += " ({} Rows Generated)".format(self.EditorItems.Count)
        if hasattr(self, 'Txt_GridTitle') and self.Txt_GridTitle:
            self.Txt_GridTitle.Text = title

    def on_sheet_number_changed(self, sheet, old_val, new_val):
        if getattr(self, "_is_auto_sequencing", False):
            return
            
        self._is_auto_sequencing = True
        try:
            # Slot Stealing Logic: check if new_val exactly matches another sheet in EditorItems
            target_slot = None
            for r in self.EditorItems:
                if r != sheet and r.SheetNumber == new_val:
                    target_slot = r
                    break
                    
            if target_slot:
                # Absorb target's metadata
                sheet.CollectionName = target_slot.CollectionName
                sheet.DisciplineName = target_slot.DisciplineName
                sheet.ContentGroupName = target_slot.ContentGroupName
                sheet.SheetName = target_slot.SheetName
                
                if target_slot.Action == "CREATE":
                    # Shift Down: Temporarily release lock to allow recursive cascading
                    self._is_auto_sequencing = False
                    
                    # Compute next number
                    match_ts = re.search(r'(\d+)$', target_slot.SheetNumber)
                    if match_ts:
                        ts_prefix = target_slot.SheetNumber[:match_ts.start()]
                        ts_num = int(match_ts.group(1))
                        ts_len = len(match_ts.group(1))
                        # Setting SheetNumber will recursively trigger on_sheet_number_changed for target_slot!
                        target_slot.SheetNumber = "{}{:0{}d}".format(ts_prefix, ts_num + 1, ts_len)
                    else:
                        # Fallback if no numeric suffix
                        target_slot.SheetNumber += "-1"
                        
                    # Re-acquire lock to finish current sheet update
                    self._is_auto_sequencing = True
                    sheet.Action = "UPDATE" if sheet.Action != "UNRECONCILED" else "UPDATE"
                else:
                    # Evict the target_slot (downgrade to UNRECONCILED)
                    target_slot.CollectionName = "Unreconciled"
                    target_slot.DisciplineName = "Unknown"
                    target_slot.ContentGroupName = "Unknown"
                    target_slot.Action = "UNRECONCILED"
                    # Try to reset its number to original, or append "-CONFLICT"
                    target_slot.SheetNumber = target_slot.OriginalNumber if hasattr(target_slot, 'OriginalNumber') else target_slot.SheetNumber + "-CONFLICT"
                    sheet.Action = "UPDATE"
                    
            idx = -1
            for i, r in enumerate(self.EditorItems):
                if r == sheet:
                    idx = i
                    break
                    
            if idx != -1 and idx < len(self.EditorItems) - 1:
                match_new = re.search(r'(\d+)$', new_val)
                match_old = re.search(r'(\d+)$', old_val)
                if match_new and match_old:
                    new_prefix = new_val[:match_new.start()]
                    new_num = int(match_new.group(1))
                    new_len = len(match_new.group(1))
                    
                    old_prefix = old_val[:match_old.start()]
                    old_num = int(match_old.group(1))
                    
                    delta = new_num - old_num
                    cascaded = False
                    
                    for i in range(idx + 1, len(self.EditorItems)):
                        next_sheet = self.EditorItems[i]
                        # Only auto-sequence if it belongs to the exact same Discipline and Group
                        if next_sheet.DisciplineName != sheet.DisciplineName or next_sheet.ContentGroupName != sheet.ContentGroupName:
                            continue
                            
                        ns_num_str = next_sheet.SheetNumber
                        if not ns_num_str: continue
                        
                        match_ns = re.search(r'(\d+)$', ns_num_str)
                        if not match_ns: continue
                            
                        ns_prefix = ns_num_str[:match_ns.start()]
                        ns_num = int(match_ns.group(1))
                        
                        if ns_prefix != old_prefix:
                            continue
                            
                        if not cascaded:
                            if new_prefix == old_prefix and new_num < ns_num:
                                break
                            cascaded = True
                            
                        proposed_num = ns_num + delta
                        next_sheet.SheetNumber = "{}{:0{}d}".format(new_prefix, proposed_num, new_len)
        finally:
            self._is_auto_sequencing = False
            self.build_matrix_board()
            self.run_validation()
            
    def update_discipline_filter_options(self):
        if not hasattr(self, 'Cmb_DisciplineFilter') or not self.Cmb_DisciplineFilter:
            return
        current_sel = self.Cmb_DisciplineFilter.SelectedItem
        self._suppress_disc_filter = True
        try:
            if hasattr(self.Cmb_DisciplineFilter.Items, "Clear"):
                self.Cmb_DisciplineFilter.Items.Clear()
            elif hasattr(self.Cmb_DisciplineFilter.Items, "clear"):
                self.Cmb_DisciplineFilter.Items.clear()
            else:
                del self.Cmb_DisciplineFilter.Items[:]
            self.Cmb_DisciplineFilter.Items.Add(u"All Disciplines")
            disc_set = set()
            for item in self.EditorItems:
                code = getattr(item, 'DisciplineCode', '')
                short_lbl = getattr(item, 'DisciplineShortLabel', '')
                if short_lbl:
                    disc_set.add(short_lbl)
                elif code:
                    disc_set.add(code)
            for d in sorted(list(disc_set)):
                self.Cmb_DisciplineFilter.Items.Add(d)
            if current_sel and current_sel in self.Cmb_DisciplineFilter.Items:
                self.Cmb_DisciplineFilter.SelectedItem = current_sel
            else:
                self.Cmb_DisciplineFilter.SelectedIndex = 0
        finally:
            self._suppress_disc_filter = False

    def on_discipline_filter_changed(self, sender, e):
        if getattr(self, '_suppress_disc_filter', False):
            return
        self.apply_filters()
        self.build_matrix_board()

    def apply_filters(self):
        try:
            from System.Windows.Data import CollectionViewSource, PropertyGroupDescription
            from System import Predicate, Object
            from System.ComponentModel import SortDescription, ListSortDirection
            
            view = CollectionViewSource.GetDefaultView(self.EditorItems)
            if not view: return
            
            view.GroupDescriptions.Clear()
            
            node = getattr(self, "_current_selected_node", None)
            if hasattr(self, "NavTree") and hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
                node = self.NavTree.SelectedItem
                
            if node and hasattr(node, "NodeType") and node.NodeType == "Root":
                view.GroupDescriptions.Add(PropertyGroupDescription("CollectionAndSeries"))
            else:
                view.GroupDescriptions.Add(PropertyGroupDescription("SheetSeries"))
            
            view.SortDescriptions.Clear()
            view.SortDescriptions.Add(SortDescription("CollectionName", ListSortDirection.Ascending))
            view.SortDescriptions.Add(SortDescription("SheetSeries", ListSortDirection.Ascending))
            view.SortDescriptions.Add(SortDescription("SheetNumber", ListSortDirection.Ascending))
            
            show_invalid = getattr(self, 'Chk_ShowInvalid', None) and self.Chk_ShowInvalid.IsChecked
            sel_disc = None
            if hasattr(self, 'Cmb_DisciplineFilter') and self.Cmb_DisciplineFilter and self.Cmb_DisciplineFilter.SelectedItem:
                raw_val = self.Cmb_DisciplineFilter.SelectedItem
                val = raw_val.strip() if raw_val else u""
                if val and val != u"All Disciplines":
                    sel_disc = val
            
            def filter_func(item):
                if show_invalid:
                    has_error = getattr(item, 'IsNameUnique', True) == False or bool(getattr(item, 'ValidationWarning', ""))
                    for v in getattr(item, 'Views', []):
                        if getattr(v, 'ValidationWarning', ""):
                            has_error = True
                            break
                    if not has_error:
                        return False
                
                if sel_disc:
                    item_lbl = getattr(item, 'DisciplineShortLabel', '')
                    item_code = getattr(item, 'DisciplineCode', '')
                    if sel_disc != item_lbl and sel_disc != item_code and not item_lbl.startswith(sel_disc):
                        return False
                        
                return True
                
            view.Filter = Predicate[Object](filter_func)
        except Exception as e:
            pass

    def on_filter_changed(self, sender, e):
        self.apply_filters()
        self.build_matrix_board()

    def run_validation(self):
        all_numbers = {}
        has_valid_work = False
        
        # Determine the active collection from NavTree
        active_collection = self.Cmb_TargetCollection.Text if hasattr(self, 'Cmb_TargetCollection') and self.Cmb_TargetCollection.Text else "PERMIT SET"
        if hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
            node = self.NavTree.SelectedItem
            if hasattr(node, "NodeType") and node.NodeType == "Collection":
                active_collection = node.Name
                
        # Rule 6: EditorItems is the testing sandbox; include unedited baseline sheets from OTHER collections for cross-check
        editor_ids = set()
        for item in self.EditorItems:
            if item.ElementId and item.ElementId != ElementId.InvalidElementId:
                val = item.ElementId.IntegerValue if hasattr(item.ElementId, "IntegerValue") else item.ElementId.Value
                editor_ids.add(val)

        validation_pool = list(self.EditorItems)
        for base_node in self.all_grid_nodes:
            if base_node.ElementId and base_node.ElementId != ElementId.InvalidElementId:
                val = base_node.ElementId.IntegerValue if hasattr(base_node.ElementId, "IntegerValue") else base_node.ElementId.Value
                if val not in editor_ids:
                    validation_pool.append(base_node)
            
        for r in validation_pool:
            if r.Action == "PURGE" or not r.IsChecked: continue
            
            # Enforce Collection Name for newly created sheets
            if r.Action == "CREATE" and r.CollectionName != active_collection:
                r.CollectionName = active_collection
                
            r.IsNameUnique = True
            r.ValidationWarning = ""
            r.ValidationBrush = "Transparent"
            num = str(r.SheetNumber).strip().lower()
            if not num: continue
            
            # Group by (Collection, Number) instead of just globally by Number
            coll = getattr(r, "CollectionName", "Default")
            key = (coll, num)
            
            if key not in all_numbers:
                all_numbers[key] = []
            all_numbers[key].append(r)
            
        # Check existing sheet numbers globally in the project
        import re
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        
        global_sheet_keys = getattr(self, '_cached_global_sheets', {})
            
        palette = ["#FCA5A5", "#FCD34D", "#86EFAC", "#93C5FD", "#F9A8D4", "#FDBA74", "#6EE7B7", "#67E8F9"]
        color_idx = 0
        has_error = False
        
        all_actively_modified_ids = set()
        for r in validation_pool:
            if not getattr(r, "IsChecked", False) or getattr(r, "Action", "") == "PURGE":
                continue
            if hasattr(r, 'ElementId') and r.ElementId != ElementId.InvalidElementId:
                r_id = r.ElementId
                r_val = r_id.IntegerValue if hasattr(r_id, 'IntegerValue') else r_id.Value
                all_actively_modified_ids.add(r_val)

        for key, items in all_numbers.items():
            is_clash = False
            conflicts = []
            
            # Internal UI Clash
            if len(items) > 1:
                is_clash = True
                for other in items:
                    name = other.SheetName or "Unnamed"
                    conflicts.append(name + " (in grid)")
            
            # Global Document Clash
            if key in global_sheet_keys:
                for global_id in global_sheet_keys[key]:
                    if global_id not in all_actively_modified_ids:
                        is_clash = True
                        conflicts.append("Existing Project Sheet (ID: {})".format(global_id))
            
            if is_clash:
                has_error = True
                brush = palette[color_idx % len(palette)]
                color_idx += 1
                for i in items:
                    i.IsNameUnique = False
                    i.ValidationBrush = brush
                    actual_conflicts = [c for c in conflicts if c != (i.SheetName or "Unnamed") + " (in grid)"]
                    i.ValidationWarning = "Sheet Number '{}' is not unique within Collection '{}'! Conflicts with:\n- {}".format(
                        i.SheetNumber, key[0], "\n- ".join(actual_conflicts))
                        
        # Scale Uniformity Audit per (CollectionName, SheetSeries)
        series_scale_counts = {}
        for r in self.EditorItems:
            if r.Action == "PURGE":
                continue
            s_key = (r.CollectionName, r.SheetSeries)
            for v in r.Views:
                sc = getattr(v, "Scale", "")
                if sc and sc not in ("NTS", "As indicated"):
                    if s_key not in series_scale_counts:
                        series_scale_counts[s_key] = {}
                    series_scale_counts[s_key][sc] = series_scale_counts[s_key].get(sc, 0) + 1

        dominant_series_scale = {}
        for s_key, counts in series_scale_counts.items():
            total_in_series = sum(counts.values())
            if total_in_series >= 2 and len(counts) > 1:
                dom_scale, dom_cnt = max(counts.items(), key=lambda x: x[1])
                if dom_cnt > 1:
                    dominant_series_scale[s_key] = dom_scale

        # Setup for global view name validation
        existing_view_names = getattr(self, '_cached_global_views', set())

        illegal_chars = r'[\\:\{\}\[\]\|;\<\>\?\~]'
        all_grid_view_names = set()
                
        # Validate view names, numbers, scale uniformity, and illegal characters
        for r in self.EditorItems:
            s_key = (r.CollectionName, r.SheetSeries)
            dom_scale = dominant_series_scale.get(s_key, None)
            for v in r.Views:
                if dom_scale and v.Scale and v.Scale not in ("NTS", "As indicated") and v.Scale != dom_scale:
                    v.ScaleWarning = "Scale outlier: '{}' differs from series majority ({})".format(v.Scale, dom_scale)
                else:
                    v.ScaleWarning = ""
                v.check_semantic_mismatch()

        for r in validation_pool:
            if r.Action == "PURGE" or not r.IsChecked: continue
            
            # 1. Illegal characters in Sheet
            if r.SheetName and re.search(illegal_chars, r.SheetName):
                has_error = True
                r.ValidationWarning += "\nSheet Name contains illegal characters!"
                r.ValidationBrush = "#EF4444"
                
            if r.SheetNumber and re.search(illegal_chars, r.SheetNumber):
                has_error = True
                r.ValidationWarning += "\nSheet Number contains illegal characters!"
                r.ValidationBrush = "#EF4444"
            
            sheet_view_numbers = set()
            
            for v in r.Views:
                v.ValidationWarning = ""
                
                # Check 2: Illegal characters in View Name
                if v.Name and re.search(illegal_chars, v.Name):
                    has_error = True
                    v.ValidationWarning = "View name contains illegal characters."
                    
                # Check 3: Missing PlanType or Level for New Views
                if v.IsNew:
                    if not v.PlanType:
                        has_error = True
                        v.ValidationWarning += " Missing Plan Type."
                    if not v.LevelName and v.IsPlanView:
                        has_error = True
                        v.ValidationWarning += " Missing Level."
                        
                # Revit allows Legends and Schedules to be placed on multiple sheets
                is_legend_or_schedule = False
                pt = getattr(v, 'PlanType', '').lower()
                if 'legend' in pt or 'schedule' in pt or 'drafting' in pt:
                    is_legend_or_schedule = True

                # Check 4: Duplicate View Numbers on the SAME sheet
                v_num = getattr(v, 'ViewNumber', '').strip().lower()
                
                if v_num and not is_legend_or_schedule:
                    if v_num in sheet_view_numbers:
                        has_error = True
                        v.ValidationWarning += " Duplicate View Number on sheet."
                    sheet_view_numbers.add(v_num)
                        
                # Check 5: Duplicate View Names GLOBALLY (only for newly created views without detail number disambiguation)
                v_name_lower = v.Name.lower() if v.Name else ""
                if v_name_lower and not v_num and not is_legend_or_schedule:
                    if v.IsNew and v_name_lower in existing_view_names:
                        has_error = True
                        v.ValidationWarning += " Name already exists in project."
                    elif v_name_lower in all_grid_view_names:
                        has_error = True
                        v.ValidationWarning += " Name duplicated in this grid."
                    all_grid_view_names.add(v_name_lower)
            
            # Check if this row is completely valid
            r_has_error = getattr(r, 'IsNameUnique', True) == False or bool(r.ValidationWarning)
            for v in r.Views:
                if getattr(v, 'ValidationWarning', ""):
                    r_has_error = True
                    r.ValidationBrush = "#FCA5A5"
                    if not r.ValidationWarning:
                        r.ValidationWarning = "View Error: " + v.ValidationWarning
                    elif "View Error:" not in r.ValidationWarning:
                        r.ValidationWarning += "\nView Error: " + v.ValidationWarning
                    break
            r.OnPropertyChanged("HasWarning")
            r.OnPropertyChanged("WarningSummary")
            r.OnPropertyChanged("MatrixPillSummary")
                    
            has_view_changes = any(v.has_pending_changes() for v in r.Views)
            if r.IsChecked and (r.Action in ["CREATE", "UPDATE", "PURGE", "RENAME_NAME", "RENAME_NUM", "RENAME_NUMBER", "RENAME_BOTH"] or has_view_changes) and not r_has_error:
                has_valid_work = True
                
        self.main_vm.IsPushEnabled = has_valid_work
        
        # Determine status and generate a synopsis
        has_errors = False
        has_unchecked_work = False
        total_work_items = 0
        
        matched_count = 0
        create_count = 0
        update_count = 0
        mismatch_count = 0
        
        for r in self.EditorItems:
            for v in r.Views:
                if getattr(v, "SemanticWarning", "") or getattr(v, "ScaleWarning", ""):
                    mismatch_count += 1
            if r.Action == "MATCHED" and not any(v.has_pending_changes() for v in r.Views):
                matched_count += 1
            elif r.Action in ["CREATE", "MISSING"]:
                create_count += 1
            elif r.Action in ["UPDATE", "RENAME_NAME", "RENAME_NUM", "RENAME_NUMBER", "RENAME_BOTH"] or any(v.has_pending_changes() for v in r.Views):
                update_count += 1
                
            if r.Action in ["CREATE", "UPDATE", "PURGE", "RENAME_NAME", "RENAME_NUM", "RENAME_NUMBER", "RENAME_BOTH"] or any(v.has_pending_changes() for v in r.Views):
                total_work_items += 1
                r_has_error = getattr(r, 'IsNameUnique', True) == False or bool(r.ValidationWarning)
                for v in r.Views:
                    if getattr(v, 'ValidationWarning', ""):
                        r_has_error = True
                        break
                if r.IsChecked:
                    if r_has_error:
                        has_errors = True
                else:
                    has_unchecked_work = True
                    
        synopsis = "Synopsis: {} Matched | {} New | {} Updating".format(matched_count, create_count, update_count)
        if mismatch_count > 0:
            synopsis += " | ⚠️ {} Audit Warnings".format(mismatch_count)
        
        if has_valid_work:
            self.Txt_Status.Text = "{}  »  Ready to Push".format(synopsis)
        else:
            if total_work_items == 0:
                self.Txt_Status.Text = "{}  »  All Sheets Up To Date (100% Matched)".format(synopsis)
            elif has_errors:
                self.Txt_Status.Text = "{}  »  Cannot push: Resolve validation errors (red text) first.".format(synopsis)
            elif has_unchecked_work:
                self.Txt_Status.Text = "{}  »  Check the items you wish to push to Revit.".format(synopsis)
            else:
                self.Txt_Status.Text = "{}  »  No valid pending actions.".format(synopsis)
                
        self.apply_filters()
        self.update_grid_title()

    def print_debug_log(self, sender, e):
        from pyrevit import script
        from Autodesk.Revit.DB import ElementId
        out = script.get_output()
        out.print_md("### Manage Sheets: Debug Log (Ordered by AIA Schema)")
        
        valid_nodes = []
        for r in self.EditorItems:
            if not r.IsChecked or r.Action == "PURGE": continue
            r_has_error = getattr(r, 'IsNameUnique', True) == False or bool(r.ValidationWarning)
            for v in r.Views:
                if getattr(v, 'ValidationWarning', ""):
                    r_has_error = True
                    break
            if not r_has_error:
                valid_nodes.append(r)
                
        out.print_md("**Total valid checked nodes to sync:** {}".format(len(valid_nodes)))
        
        node_map = { r.SheetNumber: r for r in valid_nodes }
        
        out.print_md("#### AIA Schema Mapping")
        if hasattr(self, "generated_targets") and self.generated_targets:
            for t in self.generated_targets:
                target_num = t.get("num", "")
                pure_name = t.get("name", "")
                
                if target_num in node_map:
                    r = node_map.pop(target_num)
                    elem_id_str = "NEW_SHEET_PLACEHOLDER"
                    if hasattr(r, 'ElementId') and r.ElementId != ElementId.InvalidElementId:
                        elem_id_str = str(r.ElementId.IntegerValue if hasattr(r.ElementId, 'IntegerValue') else r.ElementId.Value)
                    
                    orig_name = getattr(r, 'OriginalName', '')
                    orig_num = getattr(r, 'OriginalNumber', '')
                    prop_name = r.SheetName
                    prop_num = r.SheetNumber
                    
                    out.print_md("- **Element ID:** `{}` | **Original:** `{} - {}` --> **Target:** `{} - {}` | **AIA Pure:** `{} - {}`".format(
                        elem_id_str, orig_num, orig_name, prop_num, prop_name, target_num, pure_name))
                else:
                    out.print_md("- **Element ID:** `[Unmatched/Unchecked]` | **Original:** `None` --> **Target:** `None` | **AIA Pure:** `{} - {}`".format(
                        target_num, pure_name))
        
        if node_map:
            out.print_md("---\n#### Additional/Custom Sheets (Not in base schema order)")
            for r in node_map.values():
                elem_id_str = "NEW_SHEET_PLACEHOLDER"
                if hasattr(r, 'ElementId') and r.ElementId != ElementId.InvalidElementId:
                    elem_id_str = str(r.ElementId.IntegerValue if hasattr(r.ElementId, 'IntegerValue') else r.ElementId.Value)
                
                orig_name = getattr(r, 'OriginalName', '')
                orig_num = getattr(r, 'OriginalNumber', '')
                prop_name = r.SheetName
                prop_num = r.SheetNumber
                
                out.print_md("- **Element ID:** `{}` | **Original:** `{} - {}` --> **Target:** `{} - {}` | **AIA Pure:** `[None]`".format(
                    elem_id_str, orig_num, orig_name, prop_num, prop_name))

    def _ensure_doc_annotation_symbol(self, doc, keyword, rfa_filename):
        """
        Finds an existing Generic Annotation FamilySymbol matching `keyword` (e.g., 'Graphic Scale' or 'North Arrow')
        or loads it from the Manage Graphic Scales Resources folder.
        """
        from Autodesk.Revit.DB import FilteredElementCollector, FamilySymbol, BuiltInCategory
        symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_GenericAnnotation).ToElements()
        kw_low = keyword.lower()
        for sym in symbols:
            try:
                fam_name = (sym.Family.Name or "").lower()
                if kw_low in fam_name:
                    if not sym.IsActive:
                        sym.Activate()
                    return sym
            except Exception:
                pass

        # Attempt to load from Manage Graphic Scales Resources folder
        try:
            ext_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
            rfa_path = os.path.join(
                ext_root,
                "ODI.tab", "03_DOCUMENTATION.panel", "Manage.pulldown",
                "Manage Graphic Scales.pushbutton", "Resources", rfa_filename
            )
            if os.path.exists(rfa_path):
                doc.LoadFamily(rfa_path)
                symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_GenericAnnotation).ToElements()
                for sym in symbols:
                    try:
                        fam_name = (sym.Family.Name or "").lower()
                        if kw_low in fam_name:
                            if not sym.IsActive:
                                sym.Activate()
                            return sym
                    except Exception:
                        pass
        except Exception:
            pass
        return None

    def _place_or_update_viewport_annotations(self, doc, sheet_elem, vp_elem, view_elem, want_gs, want_na, gs_sym, na_sym, debug_log):
        """
        Places and aligns Graphic Scale and North Arrow annotation instances on `sheet_elem`
        linked to `vp_elem` & `view_elem` (updating Scale Value to `view_elem.Scale`).
        """
        if not sheet_elem or not vp_elem or not view_elem:
            return
        from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, XYZ

        # Determine anchor point below/beside the viewport label or box outline
        anchor_gs = XYZ(1.0, 0.25, 0.0)
        anchor_na = XYZ(1.35, 0.25, 0.0)
        try:
            lbl_outline = None
            if hasattr(vp_elem, "GetLabelOutline"):
                try:
                    lbl_outline = vp_elem.GetLabelOutline()
                except Exception:
                    lbl_outline = None
            if lbl_outline and lbl_outline.MinimumPoint and lbl_outline.MaximumPoint and lbl_outline.MinimumPoint.X < 1e10:
                anchor_gs = XYZ(lbl_outline.MinimumPoint.X, lbl_outline.MinimumPoint.Y - (0.375 / 12.0), 0.0)
                anchor_na = XYZ(lbl_outline.MaximumPoint.X + (0.625 / 12.0), (lbl_outline.MinimumPoint.Y + lbl_outline.MaximumPoint.Y) * 0.5, 0.0)
            else:
                box = vp_elem.GetBoxOutline()
                if box and box.MinimumPoint and box.MaximumPoint:
                    anchor_gs = XYZ(box.MinimumPoint.X + (0.5 / 12.0), box.MinimumPoint.Y - (0.5 / 12.0), 0.0)
                    anchor_na = XYZ(box.MaximumPoint.X - (0.5 / 12.0), box.MinimumPoint.Y - (0.5 / 12.0), 0.0)
        except Exception:
            pass

        # Check existing annotations on this sheet
        existing_gs = None
        existing_na = None
        try:
            annos_on_sheet = FilteredElementCollector(doc, sheet_elem.Id).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType().ToElements()
            for a in annos_on_sheet:
                try:
                    fn = (a.Symbol.Family.Name or "").lower()
                    if "graphic scale" in fn or "scale bar" in fn:
                        existing_gs = a
                    elif "north" in fn:
                        existing_na = a
                except Exception:
                    pass
        except Exception:
            pass

        v_id_val = view_elem.Id.IntegerValue if hasattr(view_elem.Id, "IntegerValue") else view_elem.Id.Value
        vp_id_val = vp_elem.Id.IntegerValue if hasattr(vp_elem.Id, "IntegerValue") else vp_elem.Id.Value

        # 1. Graphic Scale
        if want_gs and gs_sym:
            try:
                inst = existing_gs if existing_gs else doc.Create.NewFamilyInstance(anchor_gs, gs_sym, sheet_elem)
                if inst:
                    for p_name in ("Scale Value", "View Scale", "Scale"):
                        p = inst.LookupParameter(p_name)
                        if p and not p.IsReadOnly:
                            try:
                                p.Set(float(view_elem.Scale))
                            except Exception:
                                try:
                                    p.Set(int(view_elem.Scale))
                                except Exception:
                                    pass
                    p_vid = inst.LookupParameter("View Id")
                    if p_vid and not p_vid.IsReadOnly:
                        try: p_vid.Set(int(v_id_val))
                        except Exception: pass
                    p_vpid = inst.LookupParameter("Viewport Id")
                    if p_vpid and not p_vpid.IsReadOnly:
                        try: p_vpid.Set(int(vp_id_val))
                        except Exception: pass
                    c_param = inst.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                    if c_param and not c_param.IsReadOnly:
                        try: c_param.Set("LINKED_VIEW:{}".format(v_id_val))
                        except Exception: pass
                    debug_log.append("  [Annotation] Synced Graphic Scale (1:{}) on sheet {}".format(view_elem.Scale, sheet_elem.SheetNumber))
            except Exception as ex_gs:
                debug_log.append("  [Annotation WARN] Graphic Scale failed: {}".format(ex_gs))

        # 2. North Arrow
        if want_na and na_sym:
            try:
                inst_na = existing_na if existing_na else doc.Create.NewFamilyInstance(anchor_na, na_sym, sheet_elem)
                if inst_na:
                    p_vid = inst_na.LookupParameter("View Id")
                    if p_vid and not p_vid.IsReadOnly:
                        try: p_vid.Set(int(v_id_val))
                        except Exception: pass
                    p_vpid = inst_na.LookupParameter("Viewport Id")
                    if p_vpid and not p_vpid.IsReadOnly:
                        try: p_vpid.Set(int(vp_id_val))
                        except Exception: pass
                    c_param = inst_na.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                    if c_param and not c_param.IsReadOnly:
                        try: c_param.Set("LINKED_VIEW:{}".format(v_id_val))
                        except Exception: pass
                    debug_log.append("  [Annotation] Synced North Arrow on sheet {}".format(sheet_elem.SheetNumber))
            except Exception as ex_na:
                debug_log.append("  [Annotation WARN] North Arrow failed: {}".format(ex_na))

    def sync_to_revit(self, sender, e):
        import tempfile
        import os
        debug_log_path = os.path.join(os.path.expanduser('~'), 'Downloads', "pyrevit_sync_debug.log")
        debug_log = ["=== MANAGE SHEETS SYNC DEBUG LOG ==="]
        
        self.main_vm.IsPushEnabled = False
        
        tb_item = self.Combo_TitleBlocks.SelectedItem
        tb_id_val = None
        if tb_item:
            try:
                tb_id_val = tb_item.Id.IntegerValue if hasattr(tb_item.Id, "IntegerValue") else tb_item.Id.Value
            except: pass
            
        def _sync_action():
            uidoc = HOST_APP.uiapp.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            if not doc: return
            
            from Autodesk.Revit.DB import TransactionGroup, FilteredElementCollector, ViewSheet, View, ElementId, Transaction, BuiltInCategory, Viewport, XYZ, StorageType
            from data_model import parse_revit_scale
            
            # Collect existing sheets to infer majority parameter values for browser organization
            all_existing_sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
            
            param_value_counts = {}
            for sheet in all_existing_sheets:
                for param in sheet.Parameters:
                    if not param.IsReadOnly and param.StorageType == StorageType.String:
                        p_name = param.Definition.Name
                        if p_name in ["Sheet Name", "Sheet Number", "Discipline", "Content Group", "Sheet Series", "Appears In Sheet List", "Drawn By", "Checked By", "Designed By", "Approved By", "Sheet Issue Date"]:
                            continue
                        p_val = param.AsString()
                        if p_val:
                            if p_name not in param_value_counts:
                                param_value_counts[p_name] = {}
                            param_value_counts[p_name][p_val] = param_value_counts[p_name].get(p_val, 0) + 1
            
            majority_param_values = {}
            for p_name, counts in param_value_counts.items():
                if counts:
                    majority_value = max(counts.items(), key=lambda x: x[1])[0]
                    majority_param_values[p_name] = majority_value
                    
            debug_log.append("[Inference] Majority Parameter Values: {}".format(majority_param_values))
            
            # Rule 6: Sync strictly from the active EditorItems sandbox!
            validation_pool = list(self.EditorItems)
                
            needs_tb = any(r.IsChecked and r.Action == "CREATE" for r in validation_pool)
            if needs_tb and not tb_id_val:
                from pyrevit import script
                out = script.get_output()
                out.print_md("# Manage Sheets: Sync Aborted 🛑")
                out.print_md("You are attempting to create new sheets, but **no TitleBlock** is selected.")
                out.print_md("Please select a valid TitleBlock from the dropdown in the bottom right corner, or load a TitleBlock family into your project first.")
                
                self.main_vm.IsPushEnabled = True
                return
            
            # --- CROSS-COLLECTION DUPLICATE VALIDATION ---
            collection_numbers = {}
            for r in validation_pool:
                if not r.IsChecked or r.Action == "PURGE": continue
                coll = getattr(r, "CollectionName", "Default") or "Default"
                num = (r.SheetNumber or "").strip().lower()
                key = (coll, num)
                if key in collection_numbers:
                    r.ValidationWarning = "Duplicate Sheet Number in Collection"
                    collection_numbers[key].ValidationWarning = "Duplicate Sheet Number in Collection"
                else:
                    collection_numbers[key] = r
            
            # --- FILTER VALID NODES ---
            valid_nodes = []
            for r in validation_pool:
                if not r.IsChecked: continue
                r_has_error = getattr(r, 'IsNameUnique', True) == False or bool(r.ValidationWarning)
                for v in r.Views:
                    if getattr(v, 'ValidationWarning', ""):
                        r_has_error = True
                        break
                if not r_has_error:
                    valid_nodes.append(r)
                else:
                    debug_log.append("Node dropped due to error: {} - {}".format(getattr(r, 'OriginalNumber', 'NEW'), r.ValidationWarning))
                    
            debug_log.append("Total valid checked nodes: {}".format(len(valid_nodes)))
                    
            if not valid_nodes:
                self.main_vm.IsPushEnabled = True
                return
            
            # --- DEEP PREFLIGHT: Live Model Collision Check ---
            live_sheet_keys = {}
            all_live_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
            for s in all_live_sheets:
                if not s.IsTemplate:
                    s_coll = get_sheet_collection_name(doc, s) or "Default"
                    k = (s_coll, s.SheetNumber.strip().lower())
                    if k not in live_sheet_keys: live_sheet_keys[k] = []
                    s_id_val = s.Id.IntegerValue if hasattr(s.Id, 'IntegerValue') else s.Id.Value
                    live_sheet_keys[k].append(s_id_val)
                    
            live_view_names = set()
            all_live_views = FilteredElementCollector(doc).OfClass(View).ToElements()
            for v in all_live_views:
                if not v.IsTemplate:
                    live_view_names.add(v.Name.lower())
                    
            error_log = []
            auto_park_ids = []
            
            active_renaming_ids = set()
            for n in valid_nodes:
                if hasattr(n, 'ElementId') and n.ElementId != ElementId.InvalidElementId:
                    val = n.ElementId.IntegerValue if hasattr(n.ElementId, 'IntegerValue') else n.ElementId.Value
                    active_renaming_ids.add(val)
                    
            # Auto-park any unmatched native sheet blocking a target slot (Rule 6)
            for r in valid_nodes:
                if not r.IsChecked or r.Action == "PURGE": continue
                c_name = getattr(r, 'CollectionName', 'Default') or 'Default'
                key = (c_name, (r.SheetNumber or "").strip().lower())
                
                if key in live_sheet_keys:
                    r_id_val = None
                    if hasattr(r, 'ElementId') and r.ElementId != ElementId.InvalidElementId:
                        r_id_val = r.ElementId.IntegerValue if hasattr(r.ElementId, 'IntegerValue') else r.ElementId.Value
                    
                    for live_id in live_sheet_keys[key]:
                        if live_id != r_id_val and live_id not in active_renaming_ids:
                            if live_id not in auto_park_ids:
                                auto_park_ids.append(live_id)
            
            log_created = []
            log_updated = []
            log_purged = []
            log_view_renamed = []
            renames, creates, purges = 0, 0, 0
            
            with TransactionGroup(doc, "AIA Reconciliation") as tg:
                tg.Start()
                try:
                    if auto_park_ids:
                        with Transaction(doc, "Auto-Park Unmatched Sheets") as t_park:
                            t_park.Start()
                            for pid in auto_park_ids:
                                p_elem = doc.GetElement(ElementId(pid))
                                if p_elem:
                                    old_num = p_elem.SheetNumber
                                    import System
                                    new_num = old_num + "_OLD_" + System.Guid.NewGuid().ToString().Substring(0, 4)
                                    try:
                                        p_elem.SheetNumber = new_num
                                        debug_log.append("Auto-parked unmapped blocking sheet: {} -> {}".format(old_num, new_num))
                                    except Exception as ex: 
                                        debug_log.append("Failed to auto-park sheet: {} -> {}: {}".format(old_num, new_num, str(ex)))
                            t_park.Commit()
                            
                    with Transaction(doc, "Ensure Parameters") as t_param:
                        t_param.Start()
                        ensure_sheet_parameter(doc, "Discipline")
                        ensure_sheet_parameter(doc, "Content Group")
                        ensure_sheet_parameter(doc, "Sheet Series")
                        ensure_sheet_parameter(doc, " Sheet Collection")
                        t_param.Commit()
                    
                    # Phase 1: Topological Sequencing
                    renumber_nodes = []
                    current_number_to_node = {}
                    
                    for r in valid_nodes:
                        if not r.IsChecked: continue
                        if hasattr(r, 'OriginalNumber') and r.OriginalNumber:
                            c_name = getattr(r, 'OriginalCollectionName', 'Default') or 'Default'
                            current_number_to_node[(c_name, r.OriginalNumber)] = r
                            
                        match_stat = getattr(r, 'MatchStatus', None) or r.Action
                        if r.Action in ["UPDATE", "MATCHED", "RENAME_NAME", "RENAME_NUMBER", "RENAME_BOTH"] or match_stat in ["RENAME_NUMBER", "RENAME_BOTH"]:
                            if hasattr(r, 'ElementId') and r.ElementId != ElementId.InvalidElementId:
                                if r.SheetNumber != r.OriginalNumber:
                                    renumber_nodes.append(r)
                                    debug_log.append("[Phase 1] Added to renumber_nodes: {} -> {} (Action: {}, MatchStatus: {})".format(r.OriginalNumber, r.SheetNumber, r.Action, match_stat))
                                    
                    debug_log.append("[Phase 1] Total renumber_nodes: {}".format(len(renumber_nodes)))
                                
                    node_by_id = {}
                    adj = {}
                    in_degree = {}
                    
                    for r in renumber_nodes:
                        r_key = r.ElementId.IntegerValue if hasattr(r.ElementId, 'IntegerValue') else r.ElementId.Value
                        node_by_id[r_key] = r
                        adj[r_key] = []
                        in_degree[r_key] = 0
                        
                    for r in renumber_nodes:
                        target_num = r.SheetNumber
                        c_name = getattr(r, 'CollectionName', 'Default') or 'Default'
                        blocking_node = current_number_to_node.get((c_name, target_num))
                        
                        r_key = r.ElementId.IntegerValue if hasattr(r.ElementId, 'IntegerValue') else r.ElementId.Value
                        
                        if blocking_node and blocking_node.ElementId != r.ElementId:
                            blocking_key = blocking_node.ElementId.IntegerValue if hasattr(blocking_node.ElementId, 'IntegerValue') else blocking_node.ElementId.Value
                            if blocking_key in node_by_id:
                                adj[blocking_key].append(r_key)
                                in_degree[r_key] += 1
                                
                    queue = [node_id for node_id in in_degree if in_degree[node_id] == 0]
                    ordered_sequence = []
                    
                    while queue:
                        curr_id = queue.pop(0)
                        ordered_sequence.append(node_by_id[curr_id])
                        for neighbor_id in adj[curr_id]:
                            in_degree[neighbor_id] -= 1
                            if in_degree[neighbor_id] == 0:
                                queue.append(neighbor_id)
                                
                    cycle_nodes = [node_by_id[nid] for nid, deg in in_degree.items() if deg > 0]
                    
                    debug_log.append("[Phase 1] Ordered Sequence: {}".format([getattr(r, 'OriginalNumber', '') for r in ordered_sequence]))
                    debug_log.append("[Phase 1] Cycle Nodes: {}".format([getattr(r, 'OriginalNumber', '') for r in cycle_nodes]))
                    
                    # Phase 2: Finalize Sheets, Views, Scales, Graphic Scales & North Arrows
                    with Transaction(doc, "Phase 2 - Finalize") as t2:
                        t2.Start()
                        try:
                            # Pre-load Graphic Scale and North Arrow symbols if any view requests them
                            any_gs = any(getattr(v, "HasGraphicScale", False) for r in valid_nodes for v in getattr(r, "Views", []))
                            any_na = any(getattr(v, "HasNorthArrow", False) for r in valid_nodes for v in getattr(r, "Views", []))
                            gs_sym = self._ensure_doc_annotation_symbol(doc, "Graphic Scale", "Symbol - Graphic Scale.rfa") if any_gs else None
                            na_sym = self._ensure_doc_annotation_symbol(doc, "North Arrow", "Symbol - North Arrow.rfa") if any_na else None

                            # 1. Break Cycles
                            for r in cycle_nodes:
                                s_elem = doc.GetElement(r.ElementId)
                                if s_elem and s_elem.SheetNumber != r.SheetNumber:
                                    import System
                                    temp_num = r.SheetNumber + "_TMP_" + System.Guid.NewGuid().ToString().Substring(0, 4)
                                    s_elem.SheetNumber = temp_num
                                    
                            # 2. Execute Linear Sequence
                            for r in ordered_sequence:
                                s_elem = doc.GetElement(r.ElementId)
                                if s_elem:
                                    try:
                                        s_elem.SheetNumber = r.SheetNumber
                                        debug_log.append("[Phase 2] Renamed {} -> {}".format(r.OriginalNumber, r.SheetNumber))
                                    except Exception as ex:
                                        debug_log.append("[Phase 2 ERROR] Failed to rename {} to {}: {}".format(r.OriginalNumber, r.SheetNumber, str(ex)))
                                    
                            # 3. Resolve Parked Cycles
                            for r in cycle_nodes:
                                s_elem = doc.GetElement(r.ElementId)
                                if s_elem:
                                    s_elem.SheetNumber = r.SheetNumber

                            for r in valid_nodes:
                                if not r.IsChecked: continue
                                new_sheet = None
                                target_sheet_elem = None
                                match_stat = getattr(r, 'MatchStatus', None) or r.Action
                                debug_log.append("[Phase 2 Loop 4] Processing: {} -> {} (Action: {}, MatchStat: {})".format(getattr(r, 'OriginalNumber', 'NEW'), r.SheetNumber, r.Action, match_stat))
                                
                                if r.Action == "CREATE":
                                    tb_id = ElementId(tb_id_val)
                                    new_sheet = ViewSheet.Create(doc, tb_id)
                                    new_sheet.SheetNumber = r.SheetNumber
                                    new_sheet.Name = r.SheetName
                                    assign_sheet_to_collection(doc, new_sheet, r.CollectionName)
                                    c_res = classification.classify_sheet(r.SheetNumber, r.SheetName)
                                    disc_name = c_res.get("discipline", "Unknown")
                                    cg_name = c_res.get("contentGroup", "Uncategorized")
                                    set_sheet_parameter(new_sheet, "Discipline", disc_name)
                                    set_sheet_parameter(new_sheet, "Content Group", cg_name)
                                    set_sheet_parameter(new_sheet, "Sheet Series", getattr(r, "SheetSeries", "General"))
                                    
                                    for p_name, p_val in majority_param_values.items():
                                        try:
                                            param = new_sheet.LookupParameter(p_name)
                                            if param and not param.IsReadOnly and not param.AsString():
                                                param.Set(p_val)
                                        except Exception: pass
                                        
                                    target_sheet_elem = new_sheet
                                    log_created.append("{} - {}".format(r.SheetNumber, r.SheetName))
                                    creates += 1
                                elif r.Action in ["UPDATE", "MATCHED", "RENAME_NAME", "RENAME_NUMBER", "RENAME_BOTH"] or match_stat in ["RENAME_NUMBER", "RENAME_BOTH"]:
                                    s_elem = doc.GetElement(r.ElementId)
                                    if s_elem:
                                        if r.SheetName != r.OriginalName: s_elem.Name = r.SheetName
                                        assign_sheet_to_collection(doc, s_elem, r.CollectionName)
                                        c_res = classification.classify_sheet(r.SheetNumber, r.SheetName)
                                        disc_name = c_res.get("discipline", "Unknown")
                                        cg_name = c_res.get("contentGroup", "Uncategorized")
                                        set_sheet_parameter(s_elem, "Discipline", disc_name)
                                        set_sheet_parameter(s_elem, "Content Group", cg_name)
                                        set_sheet_parameter(s_elem, "Sheet Series", r.SheetSeries)
                                        target_sheet_elem = s_elem
                                        log_updated.append("{} - {}".format(r.SheetNumber, r.SheetName))
                                        renames += 1
                                elif r.Action == "PURGE":
                                    s_elem = doc.GetElement(r.ElementId)
                                    if s_elem:
                                        assign_sheet_to_collection(doc, s_elem, "PURGE")
                                    log_purged.append("{} - {} (Quarantined)".format(r.SheetNumber, r.SheetName))
                                    purges += 1
                                        
                                if r.Action != "PURGE" and r.IsChecked and target_sheet_elem:
                                    target_sheet_id = target_sheet_elem.Id
                                    
                                    def park_conflicting_dn(doc_obj, sht_id, target_dn, skip_vp_id=None):
                                        from Autodesk.Revit.DB import FilteredElementCollector, Viewport, BuiltInParameter
                                        import uuid
                                        if not target_dn: return
                                        vps_on_sheet = FilteredElementCollector(doc_obj).OwnedByView(sht_id).OfClass(Viewport).ToElements()
                                        for v_on_sht in vps_on_sheet:
                                            if skip_vp_id and v_on_sht.Id == skip_vp_id: continue
                                            dn_p = v_on_sht.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                            if dn_p and not dn_p.IsReadOnly and dn_p.AsString() == str(target_dn):
                                                try: dn_p.Set(str(target_dn) + "_TEMP_" + str(uuid.uuid4())[:4])
                                                except: pass

                                    for v in getattr(r, 'Views', []):
                                        shorthand = ""
                                        if r.CollectionName and r.CollectionName != "Undefined":
                                            words = r.CollectionName.split()
                                            if len(words) > 1:
                                                shorthand = "".join(w[0] for w in words if w).upper()
                                            else:
                                                shorthand = r.CollectionName[:3].upper()
                                                
                                        view_number_str = v.ViewNumber or "1"
                                        unique_name_parts = [v.Name]
                                        if shorthand: unique_name_parts.append(shorthand)
                                        unique_name_parts.append(view_number_str)
                                        if r.SheetNumber: unique_name_parts.append(r.SheetNumber)
                                        
                                        unique_backend_name = " - ".join(unique_name_parts)
                                        
                                        from Autodesk.Revit.DB import BuiltInParameter
                                        if getattr(v, 'ViewId', ElementId.InvalidElementId) != ElementId.InvalidElementId and not getattr(v, 'IsNew', False):
                                            v_elem = doc.GetElement(v.ViewId)
                                            if v_elem:
                                                # 1. Update View Scale if user changed it in the workbench
                                                desired_scale_int = parse_revit_scale(v.Scale)
                                                if desired_scale_int and getattr(v_elem, "Scale", None) != desired_scale_int:
                                                    try:
                                                        v_elem.Scale = desired_scale_int
                                                        debug_log.append("  [View Scale] Updated '{}' scale to 1:{}".format(v.Name, desired_scale_int))
                                                    except Exception:
                                                        pass

                                                # 2. Set Title on Sheet (Desired Name)
                                                title_param = v_elem.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                                                if title_param and not title_param.IsReadOnly:
                                                    if title_param.AsString() != v.Name:
                                                        try: title_param.Set(v.Name)
                                                        except: pass
                                                        
                                                # 3. Set Backend Unique Name
                                                if v.ViewNumber and v_elem.Name != unique_backend_name:
                                                    try:
                                                        v_elem.Name = unique_backend_name
                                                        renames += 1
                                                        log_view_renamed.append(v.Name)
                                                    except: pass
                                            
                                            from Autodesk.Revit.DB import Viewport
                                            vps = FilteredElementCollector(doc).OfClass(Viewport).ToElements()
                                            target_vp = None
                                            for vp in vps:
                                                if vp.ViewId == v.ViewId:
                                                    target_vp = vp
                                                    break
                                            
                                            final_vp = None
                                            if target_vp:
                                                if target_vp.SheetId != target_sheet_id:
                                                    center = target_vp.GetBoxCenter()
                                                    doc.Delete(target_vp.Id)
                                                    if Viewport.CanAddViewToSheet(doc, target_sheet_id, v.ViewId):
                                                        final_vp = Viewport.Create(doc, target_sheet_id, v.ViewId, center)
                                                        park_conflicting_dn(doc, target_sheet_id, v.ViewNumber, final_vp.Id)
                                                        dn_param = final_vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                                        if dn_param and not dn_param.IsReadOnly and v.ViewNumber:
                                                            try: dn_param.Set(str(v.ViewNumber))
                                                            except: pass
                                                else:
                                                    final_vp = target_vp
                                                    park_conflicting_dn(doc, target_sheet_id, v.ViewNumber, target_vp.Id)
                                                    dn_param = target_vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                                    if dn_param and not dn_param.IsReadOnly and v.ViewNumber:
                                                        if dn_param.AsString() != str(v.ViewNumber):
                                                            try: dn_param.Set(str(v.ViewNumber))
                                                            except: pass
                                            else:
                                                # Unplaced Existing View selected from picker -> place onto target_sheet_id!
                                                if Viewport.CanAddViewToSheet(doc, target_sheet_id, v.ViewId):
                                                    final_vp = Viewport.Create(doc, target_sheet_id, v.ViewId, XYZ(1.5, 1.0, 0))
                                                    park_conflicting_dn(doc, target_sheet_id, v.ViewNumber, final_vp.Id)
                                                    dn_param = final_vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                                    if dn_param and not dn_param.IsReadOnly and v.ViewNumber:
                                                        try: dn_param.Set(str(v.ViewNumber))
                                                        except: pass

                                            if final_vp and v_elem:
                                                self._place_or_update_viewport_annotations(
                                                    doc, target_sheet_elem, final_vp, v_elem,
                                                    getattr(v, "HasGraphicScale", False),
                                                    getattr(v, "HasNorthArrow", False),
                                                    gs_sym, na_sym, debug_log
                                                )
                                        elif getattr(v, 'IsNew', False):
                                            from Autodesk.Revit.DB import ViewFamilyType, ViewFamily, Level, ViewPlan, ViewDrafting
                                            view_to_place = None
                                            
                                            if v.PlanType:
                                                vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
                                                vft_target = None
                                                for vft in vfts:
                                                    try:
                                                        v_name = vft.Name
                                                    except AttributeError:
                                                        try:
                                                            from Autodesk.Revit.DB import BuiltInParameter
                                                            p = vft.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                                                            v_name = p.AsString() if p else None
                                                        except:
                                                            v_name = None
                                                            
                                                    if v_name == v.PlanType:
                                                        vft_target = vft
                                                        break
                                                
                                                if not vft_target and vfts:
                                                    vft_target = vfts[0]
                                                    
                                                if vft_target:
                                                    vft_id = vft_target.Id
                                                    vft_family = vft_target.ViewFamily
                                                    
                                                    if vft_family in [ViewFamily.FloorPlan, ViewFamily.CeilingPlan, ViewFamily.StructuralPlan, ViewFamily.AreaPlan]:
                                                        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                                                        target_lvl_id = None
                                                        if v.LevelName:
                                                            for lvl in levels:
                                                                if lvl.Name == v.LevelName:
                                                                    target_lvl_id = lvl.Id
                                                                    break
                                                        if not target_lvl_id and levels:
                                                            target_lvl_id = levels[0].Id
                                                            
                                                        if target_lvl_id:
                                                            view_to_place = ViewPlan.Create(doc, vft_id, target_lvl_id)
                                                    elif vft_family == ViewFamily.Drafting:
                                                        view_to_place = ViewDrafting.Create(doc, vft_id)
                                                        
                                                    if view_to_place:
                                                        desired_scale_int = parse_revit_scale(v.Scale)
                                                        if desired_scale_int:
                                                            try:
                                                                view_to_place.Scale = desired_scale_int
                                                            except Exception:
                                                                pass
                                            
                                            if view_to_place:
                                                title_param = view_to_place.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                                                if title_param and not title_param.IsReadOnly:
                                                    try: title_param.Set(v.Name)
                                                    except: pass
                                                
                                                try:
                                                    view_to_place.Name = unique_backend_name if v.ViewNumber else v.Name
                                                    renames += 1
                                                    log_view_renamed.append(v.Name)
                                                except: pass
                                                
                                                if Viewport.CanAddViewToSheet(doc, target_sheet_id, view_to_place.Id):
                                                    new_vp = Viewport.Create(doc, target_sheet_id, view_to_place.Id, XYZ(1.5, 1.0, 0))
                                                    park_conflicting_dn(doc, target_sheet_id, v.ViewNumber, new_vp.Id)
                                                    from Autodesk.Revit.DB import BuiltInParameter
                                                    dn_param = new_vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                                    if dn_param and not dn_param.IsReadOnly and v.ViewNumber:
                                                        try: dn_param.Set(str(v.ViewNumber))
                                                        except: pass
                                                    self._place_or_update_viewport_annotations(
                                                        doc, target_sheet_elem, new_vp, view_to_place,
                                                        getattr(v, "HasGraphicScale", False),
                                                        getattr(v, "HasNorthArrow", False),
                                                        gs_sym, na_sym, debug_log
                                                    )
                                                    creates += 1
                            t2.Commit()
                        except:
                            if t2.HasStarted() and not t2.HasEnded(): t2.RollBack()
                            raise
                        
                    tg.Assimilate()
                    from pyrevit import script
                    out = script.get_output()
                    out.print_md("# Manage Sheets: Sync Complete 🚀")
                    out.print_md("---")
                    
                    if log_created:
                        out.print_md("## ✨ Created Sheets ({})".format(len(log_created)))
                        for item in log_created: out.print_md("- {}".format(item))
                    if log_updated:
                        out.print_md("## 📝 Updated/Renamed Sheets ({})".format(len(log_updated)))
                        for item in log_updated: out.print_md("- {}".format(item))
                    if log_purged:
                        out.print_md("## 🗑️ Purged Sheets ({})".format(len(log_purged)))
                        if len(log_purged) > 0:
                            out.print_md("### Purged")
                            for l in log_purged: out.print_md("- " + l)
                        
                    if len(log_view_renamed) > 0:
                        out.print_md("### Views Renamed")
                        for l in log_view_renamed: out.print_md("- " + l)
                        
                    out.print_md("---")
                    out.print_md("**Total Operations:** {}".format(len(log_created) + len(log_updated) + len(log_purged)))
                    
                    if error_log:
                        out.print_md("---")
                        out.print_md("## ⚠️ Collisions Detected")
                        out.print_md("The following items were skipped due to late-stage live model collisions:")
                        for err in error_log:
                            out.print_md("- {}".format(err))
                            
                    try:
                        with open(debug_log_path, 'w') as f:
                            f.write("\n".join(debug_log))
                        out.print_md("---")
                        out.print_md("**Debug Log Saved:** `{}`".format(debug_log_path))
                    except: pass
                            
                    has_dropped_nodes = any(getattr(r, 'ValidationWarning', False) for r in validation_pool if r.IsChecked)
                    if has_dropped_nodes or error_log:
                        out.print_md("---")
                        out.print_md("## ⚠️ WARNING: Some Sheets Were Dropped")
                        out.print_md("Validation errors prevented some sheets from syncing. The Manage Sheets window will now close. Please re-run the tool to resolve the remaining errors.")
                    
                    self.Close()
                except Exception as ex:
                    if tg.HasStarted() and not tg.HasEnded(): tg.RollBack()
                    import traceback
                    err_msg = traceback.format_exc()
                    
                    from pyrevit import script
                    out = script.get_output()
                    out.print_md("# Manage Sheets: Sync Failed ❌")
                    out.print_md("All changes have been safely rolled back.")
                    out.print_md("```\n{}\n```".format(err_msg))
                    
                    debug_log.append("CRASH OCCURRED: " + str(ex))
                    debug_log.append(err_msg)
                    try:
                        with open(debug_log_path, 'w') as f:
                            f.write("\n".join(debug_log))
                        out.print_md("---")
                        out.print_md("**Debug Log Saved:** `{}`".format(debug_log_path))
                    except: pass
                    
                    self.main_vm.IsPushEnabled = True
        
        _sync_action()

    def on_textbox_keydown(self, sender, e):
        from System.Windows.Input import Key, TraversalRequest, FocusNavigationDirection
        if e.Key == Key.Enter:
            from System.Windows.Controls import TextBox, ComboBox
            if isinstance(sender, TextBox):
                binding = sender.GetBindingExpression(TextBox.TextProperty)
                if binding: binding.UpdateSource()
            elif isinstance(sender, ComboBox):
                binding = sender.GetBindingExpression(ComboBox.TextProperty)
                if binding: binding.UpdateSource()
                
            req = TraversalRequest(FocusNavigationDirection.Next)
            req.Wrapped = True
            sender.MoveFocus(req)
