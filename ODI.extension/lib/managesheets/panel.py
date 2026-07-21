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

def generate_discipline_sheets(disc_code, levels, active_series, active_modifiers, global_cover, rows, cols, naming_scheme, custom_schemes=None, excluded_targets=None, shuffle_on_exclude=False, modifier_overrides=None):
    if excluded_targets is None: excluded_targets = set()
    if modifier_overrides is None: modifier_overrides = {}
    targets = []
    suffixes = generate_suffixes(rows, cols, naming_scheme, custom_schemes)
    disc_dict = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors", "CS": "Cover Sheet" }
    target_disc_name = disc_dict.get(disc_code, None)

    # 1. Pre-calculate series_mods for all active series to determine universal padding and multiplier
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
                    # Prevent duplicating completely unknown custom modifiers across every series
                    if len(active_series) > 0 and series == active_series[0]:
                        if (mod, series + "99", "Custom") not in series_mods:
                            series_mods.append((mod, series + "99", "Custom"))
        
        if not series_mods:
            series_mods.append((SERIES_MAP.get(series, series).upper(), series + "01", "0{}. {}".format(series, SERIES_MAP.get(series, series))))
            
        all_series_mods[series] = series_mods
        if len(series_mods) > max_mods:
            max_mods = len(series_mods)
            
    # 2. Determine Universal Multiplier and Padding format
    seq_mult = 100 if max_mods > 10 else 10
    pad_fmt = "{:03d}" if max_mods > 10 else "{:02d}"
    
    # 3. Generate targets using the pre-calculated mods and universal formatting
    for series in active_series:
        if series not in all_series_mods:
            continue
        series_mods = all_series_mods[series]
        
        if series == "1":
            for idx, lvl in enumerate(levels):
                lvl_seq = (idx + 1) * seq_mult
                actual_mod_seq = 0
                for mod_idx, (m_name, m_code, cg) in enumerate(series_mods):
                    baseline_seq = lvl_seq + mod_idx
                    baseline_num = "{}-{}{}".format(disc_code, series, pad_fmt.format(baseline_seq))
                    is_excluded = baseline_num in excluded_targets
                    
                    if shuffle_on_exclude:
                        if is_excluded:
                            actual_num = baseline_num + " [Skipped]"
                        else:
                            actual_num = "{}-{}{}".format(disc_code, series, pad_fmt.format(lvl_seq + actual_mod_seq))
                            actual_mod_seq += 1
                    else:
                        actual_num = baseline_num
                        
                    if len(suffixes) == 1:
                        targets.append({"num": actual_num, "baseline_num": baseline_num, "name": "{} {}".format(m_name.upper(), lvl.upper()), "series_name": "01. Plans", "cg": cg, "type": "1"})
                    else:
                        targets.append({"num": actual_num, "baseline_num": baseline_num, "name": "{} {} - OVERALL".format(m_name.upper(), lvl.upper()), "series_name": "01. Plans", "cg": cg, "type": "1"})
                        for let, name_part in suffixes:
                            if m_name in modifier_overrides and modifier_overrides[m_name] is not None:
                                if let not in modifier_overrides[m_name]:
                                    continue
                                    
                            baseline_num_suffixed = "{}{}".format(baseline_num, let)
                            is_suffixed_excluded = baseline_num_suffixed in excluded_targets
                            
                            if shuffle_on_exclude:
                                if is_suffixed_excluded:
                                    final_actual = baseline_num_suffixed + " [Skipped]"
                                else:
                                    final_actual = "{}{}".format(actual_num.replace(" [Skipped]", ""), let)
                            else:
                                final_actual = baseline_num_suffixed
                                
                            prefix_mod = m_name.upper()
                            if not prefix_mod.startswith("ENLARGED"):
                                # e.g. "FLOOR PLAN" -> "ENLARGED FLOOR PLAN"
                                prefix_mod = "ENLARGED " + prefix_mod
                                
                            targets.append({"num": final_actual, "baseline_num": baseline_num_suffixed, "name": "{} {} - {}".format(prefix_mod, lvl.upper(), name_part.upper()), "series_name": "01. Plans", "cg": cg, "type": "1"})
                            
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
                targets.append({"num": actual_num, "baseline_num": baseline_num, "name": m_name.upper(), "series_name": coll_name, "cg": cg, "type": series})
                
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
    # TEMPORARY DEBUG: Forcing to True to verify if the apply_theme function is working correctly
    # If the UI turns dark after this, then our Revit API detection logic was failing.
    # If it stays light, then the Resource override logic is failing.
    return True

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
        
        self.MainTabControl.SelectionChanged += self.on_tab_changed
        
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
        self.DataContext = self.main_vm

        # Data Models
        self.NavRoot = self.main_vm.NavRoot
        self.NavTree.ItemsSource = self.NavRoot
        self.NavTree.SelectedItemChanged += self.on_tree_selection_changed
        
        self.TargetSchemaRoot = ObservableCollection[NavTreeNode]()
        self.TargetSchemaTree.ItemsSource = self.TargetSchemaRoot
        
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
        

        self.load_settings()
        self.generate_target_schema() # Initial run
        self.check_and_load_data()

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

    def on_tab_changed(self, sender, e):
        # Ignore SelectionChanged events that bubble up from child controls (like DataGrid or ComboBox)
        if e.OriginalSource != self.MainTabControl:
            return
            
        if self.MainTabControl.SelectedIndex == 1:
            self.FooterBorder.Visibility = Visibility.Visible
            try:
                # Force rebind just in case TabControl unloaded it
                # Force a schema regeneration in case settings changed in Tab 0
                self.generate_target_schema()
                
                # Force a refresh of the grid based on the latest schema when switching tabs
                node = getattr(self, "_current_selected_node", None)
                if hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
                    node = self.NavTree.SelectedItem
                elif self.NavRoot.Count > 0:
                    node = self.NavRoot[0]
                    
                if node:
                    self._current_selected_node = node
                    
                    class DummyArgs: pass
                    self.on_tree_selection_changed(self.NavTree, DummyArgs())
                else:
                    self.Txt_GridTitle.Text = "WARNING: No node selected and NavRoot is empty!"
            except Exception as ex:
                self.Txt_GridTitle.Text = "TAB ERROR: " + str(ex)
        else:
            self.FooterBorder.Visibility = Visibility.Collapsed
            
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
        existing_collections = set()
        if doc:
            from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet
            sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
            for s in sheets:
                if s.IsPlaceholder: continue
                c_name = get_sheet_collection_name(doc, s)
                if c_name and c_name != "Undefined":
                    existing_collections.add(c_name)
        if "PERMIT SET" not in existing_collections:
            existing_collections.add("PERMIT SET")
            
        self.Cmb_TargetCollection.ItemsSource = sorted(list(existing_collections))
        self.Cmb_TargetCollection.Text = p_setup.get("target_collection", cfg.get_option("target_collection", "PERMIT SET"))
        self.Cmb_TargetCollection.LostFocus += self.trigger_generation
        self.Cmb_TargetCollection.SelectionChanged += self.trigger_generation
        
        self.Sld_GridRows.Value = float(p_setup.get("grid_rows", cfg.get_option("grid_rows", 1)))
        self.Sld_GridCols.Value = float(p_setup.get("grid_cols", cfg.get_option("grid_cols", 1)))
        
        saved_discs = p_setup.get("disciplines", cfg.get_option("disciplines", ["A", "M", "E", "P"]))
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
        custom_modifiers = p_setup.get("custom_modifiers", cfg.get_option("custom_modifiers", []))
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
                
                cfg.selected_modifiers = setup_dict.get("selected_modifiers", [])
                cfg.custom_modifiers = setup_dict.get("custom_modifiers", [])
                cfg.modifier_grid_overrides = setup_dict.get("modifier_grid_overrides", {})
                self.excluded_target_sheets = set(setup_dict.get("excluded_target_sheets", []))
                
                # Re-check disciplines and series
                self.toggle_list(self.DisciplineNodes, False)
                if "disciplines" in setup_dict:
                    for d_code in setup_dict["disciplines"]:
                        for d_node in self.DisciplineNodes:
                            if d_node.Name.startswith(d_code + " -"):
                                d_node.IsChecked = True
                                
                self.toggle_list(self.SeriesNodes, False)
                if "series" in setup_dict:
                    for s_code in setup_dict["series"]:
                        for s_node in self.SeriesNodes:
                            if s_node.Name.startswith(s_code + " -"):
                                s_node.IsChecked = True
                                
                self.load_modifiers_from_cfg(
                    saved_modifiers=setup_dict.get("selected_modifiers", []),
                    custom_modifiers=setup_dict.get("custom_modifiers", []),
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
        t_root = NavTreeNode("AIA Schema", "Root")
        self.TargetSchemaRoot.Add(t_root)
        
        d_map = {}
        cg_map = {}
        
        for t in self.generated_targets:
            match = re.match(r"^([A-Z]+)[- ]?(\d+)", t["num"].upper())
            disc_code = match.group(1) if match else "Other"
            disc_dict = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors", "CS": "Cover Sheet" }
            disc_name = "{} - {}".format(disc_code, disc_dict.get(disc_code, "Discipline")) if disc_code != "Other" else "Uncategorized"
            
            if disc_name not in d_map:
                dn = NavTreeNode(disc_name, "Discipline", tag=disc_name)
                dn.target_callback = self.on_target_node_toggled
                d_map[disc_name] = dn
                t_root.Children.Add(dn)
                
            cg_name = t.get("cg", "Unknown Content Group")
            cg_key = (disc_name, cg_name)
            
            if cg_key not in cg_map:
                cn = NavTreeNode(cg_name, "ContentGroup", tag=disc_name)
                cn.target_callback = self.on_target_node_toggled
                cg_map[cg_key] = cn
                d_map[disc_name].Children.Add(cn)
                
            sn = NavTreeNode("{} - {}".format(t["num"], t["name"]), "Sheet", tag=t["baseline_num"])
            if t["baseline_num"] in self.excluded_target_sheets:
                sn.IsTargetIncluded = False
                if not self.Tgl_ShowExcluded.IsChecked:
                    sn.IsVisible = False
            sn.target_callback = self.on_target_node_toggled
            cg_map[cg_key].Children.Add(sn)
            
            t_root.Count += 1
            d_map[disc_name].Count += 1
            cg_map[cg_key].Count += 1
        if not self.generated_targets:
            pass
        else:
            pass
            
        # Only enable Sheet Reviewer tab if a schema is actively generated
        if hasattr(self, 'Tab_SheetReviewer'):
            self.main_vm.IsReviewerReady = (len(self.generated_targets) > 0)
            
        # Auto-expand the target schema tree so user sees the new combinations immediately
        if hasattr(self, 'toggle_tree'):
            self.toggle_tree(self.TargetSchemaRoot, True)

    def on_target_node_toggled(self, node):
        import System
        from System.Windows.Input import Keyboard, ModifierKeys
        
        # Shift-Select Logic
        if Keyboard.Modifiers == ModifierKeys.Shift:
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
            # If a parent is toggled, recursively toggle all its children
            # Temporarily unhook the callback to prevent recursive flood
            for child in node.Children:
                old_cb = child.target_callback
                child.target_callback = None
                child.IsTargetIncluded = node.IsTargetIncluded
                child.target_callback = old_cb
                self.on_target_node_toggled(child)
                
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
                s_coll = "Default"
                try:
                    p = s.LookupParameter(" Sheet Collection")
                    if p and p.HasValue:
                        s_coll = p.AsString()
                except: pass
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
        
        # Add default option for new view
        from data_model import AssignableViewNode
        
        # Collect all assignable views
        from Autodesk.Revit.DB import View, Viewport
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        viewports = FilteredElementCollector(doc).OfClass(Viewport).ToElements()
        
        placed_views_map = {}
        for vp in viewports:
            try:
                sheet = doc.GetElement(vp.SheetId)
                if sheet:
                    placed_views_map[vp.ViewId.ToString()] = sheet.SheetNumber
            except: pass
            
        for v in all_views:
            if not v.IsTemplate and v.CanBePrinted:
                v_id_str = v.Id.ToString()
                is_placed = v_id_str in placed_views_map
                sh_num = placed_views_map.get(v_id_str, "")
                self.AssignableViews.Add(AssignableViewNode(v.Id, v.Name, is_placed, sh_num))
        
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        
        def get_elev(l):
            try: return l.ProjectElevation
            except AttributeError: return l.Elevation
            
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

        import data_model
        data_model.AVAILABLE_LEVELS = [lvl.Name for lvl in sorted(levels, key=get_elev)]
        vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        
        from Autodesk.Revit.DB import ViewFamily
        valid_families = [ViewFamily.FloorPlan, ViewFamily.CeilingPlan, ViewFamily.StructuralPlan]
        
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
            
        root_node = NavTreeNode("All Sheets", "Root")
        root_node.IsExpanded = True
        self.NavRoot.Add(root_node)
        
        col_map = {}
        disc_map = {}
        cg_map = {}
        
        for s in sheets:
            if s.IsPlaceholder: continue
            
            c_name = get_sheet_collection_name(doc, s)
            
            c_result = classification.classify_sheet(s.SheetNumber, s.Name)
            disc_name = c_result.get("discipline", "Unknown")
            cg_name = c_result.get("contentGroup", "Uncategorized")
            draw_type = c_result.get("drawingType", "Unknown")
            
            # 1. Sheet Collection Level
            if c_name not in col_map:
                col_node = NavTreeNode(c_name, "Collection", tag=c_name)
                col_map[c_name] = col_node
                root_node.Children.Add(col_node)
            else:
                col_node = col_map[c_name]
                
            sh_row = SheetViewModel(s.Id, s.SheetNumber, s.Name, c_name, discipline_name=disc_name, content_group_name=cg_name, validation_callback=self.run_validation, number_changed_callback=self.on_sheet_number_changed, context_callback=self.get_sheet_context)
            self.all_grid_nodes.append(sh_row)
            
            root_node.Count += 1
            col_node.Count += 1
            
            views = 0
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
                    
                v_row = ViewViewModel(v.Id, v_name, str(v.ViewType), level_name=lvl_name, view_number=v_number)
                sh_row.Views.Add(v_row)
                views += 1
        self.run_validation()
        
        # Prepopulate AIA Schema
        self.generate_target_schema()
        # Automatically select the Root node to populate the initial grid
        self.NavTree.SelectedItemChanged -= self.on_tree_selection_changed
        self.NavTree.SelectedItemChanged += self.on_tree_selection_changed
        if self.NavRoot.Count > 0:
            self._fake_selection(self.NavRoot[0])
            
    def _fake_selection(self, node):
        class DummyArgs: pass
        self._current_selected_node = node
        self.on_tree_selection_changed(self.NavTree, DummyArgs())

            
    def on_tree_selection_changed(self, sender, e):
        if getattr(self, "_ignore_tree_event", False): return
        if hasattr(e, "OriginalSource") and e.OriginalSource != sender: return
        
        new_node = getattr(self, "_current_selected_node", None)
        if hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
            # Prevent WPF DisconnectedItem crashes
            if hasattr(self.NavTree.SelectedItem, "NodeType"):
                new_node = self.NavTree.SelectedItem
                
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
        
        c_name = "Undefined"
        has_real_collections = False
        if self.NavRoot.Count > 0:
            for child in self.NavRoot[0].Children:
                if child.Name not in ["Undefined", "Default"]:
                    has_real_collections = True
                    break
        if not has_real_collections:
            c_name = self.Cmb_TargetCollection.Text if hasattr(self, 'Cmb_TargetCollection') and self.Cmb_TargetCollection.Text else "PERMIT SET"
            
        if not hasattr(node, "NodeType"): return
        
        active_discipline = None
        active_cg = None
        
        if node.NodeType == "Root":
            # Pass all sheets to match against the raw AIA schema so existing sheets in the project are detected
            # For the Root node, we must target the intended AIA schema collection from the UI
            c_name = self.Cmb_TargetCollection.Text if hasattr(self, 'Cmb_TargetCollection') and self.Cmb_TargetCollection.Text else "PERMIT SET"
            valid_sheets = list(self.all_grid_nodes)
            self.Txt_GridTitle.Text = "All Generated Sheets"
        elif node.NodeType == "Collection":
            c_name = node.Tag
            self.Txt_GridTitle.Text = "Collection: " + c_name
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name]
        elif node.NodeType == "Discipline":
            c_name, d_name = node.Tag
            active_discipline = d_name
            self.Txt_GridTitle.Text = "Collection: {} | Discipline: {}".format(c_name, d_name)
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name and s.DisciplineName == d_name]
        elif node.NodeType == "ContentGroup":
            c_name, d_name, cg_name = node.Tag
            active_discipline = d_name
            active_cg = cg_name
            self.Txt_GridTitle.Text = "Collection: {} | Group: {}".format(c_name, cg_name)
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name and s.DisciplineName == d_name and s.ContentGroupName == cg_name]
            
        self.update_grid_title()
        
        # Highlight active collection in NavTree
        for root_node in self.NavRoot:
            if c_name is None:
                root_node.IsActiveContext = True
            else:
                root_node.IsActiveContext = False
            for c_node in root_node.Children:
                c_node.IsActiveContext = (c_node.Tag == c_name)
                
        self.execute_schema_match(valid_sheets, active_collection=c_name, active_discipline=active_discipline, active_cg=active_cg)

    def execute_schema_match(self, valid_sheets, active_collection="Undefined", active_discipline=None, active_cg=None):
        raw_targets = getattr(self, "generated_targets", [])
        
        # Filter generated targets based on the active node context
        generated_targets = []
        for t in raw_targets:
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
                plan = reconciliation.run_pipeline(doc, generated_targets, existing_sheet_ids=valid_sheet_ids)
            except Exception as e:
                self.Txt_GridTitle.Text = "ERROR IN PIPELINE: " + str(e)
                forms.alert("ERROR IN PIPELINE: " + str(e))
                plan = {"rows": []}
            
            from Autodesk.Revit.DB import ElementId
            existing_nodes = {}
            existing_unmatched_nodes = {}
            for n in self.all_grid_nodes:
                if n.ElementId and n.ElementId != ElementId.InvalidElementId:
                    try:
                        key = n.ElementId.IntegerValue if hasattr(n.ElementId, "IntegerValue") else n.ElementId.Value
                        existing_nodes[key] = n
                    except Exception:
                        pass
                else:
                    # Cache unmatched/missing nodes by (collection, original_number) to prevent duplicates
                    key = (n.OriginalCollectionName, n.OriginalNumber)
                    existing_unmatched_nodes[key] = n
                        
            self.EditorItems.Clear()
            
            for row in plan["rows"]:
                is_template = False
                if row["status"] == "MISSING":
                    is_template = True
                
                sh_id = row["sheet_element_id"]
                
                if sh_id in existing_nodes and sh_id != -1:
                    vm = existing_nodes[sh_id]
                    # DO NOT overwrite CollectionName with the AIA generated schema (which doesn't have native Revit collection context).
                    # Keep it bound to its existing Native Revit Collection (e.g. 'Permit Set')
                    vm.SheetSeries = row.get("series_name", "Unknown")
                    vm.DisciplineName = row["discipline"]
                    vm.ContentGroupName = row["cg"]
                    if row["status"] != "UNRECONCILED" and row["status"] != "EXTRA":
                        vm.SheetNumber = row["target_number"]
                        vm.SheetName = row["target_name"]
                    
                    vm.move_up_callback = self.move_item_up
                    vm.move_down_callback = self.move_item_down

                    
                    vm._action = row["status"]
                    vm.IsChecked = (vm._action != "UNRECONCILED")
                    self.EditorItems.Add(vm)
                else:
                    tgt_num = row["target_number"] if is_template else row["existing_number"]
                    key = (active_collection, tgt_num)
                    
                    if key in existing_unmatched_nodes:
                        vm = existing_unmatched_nodes[key]
                        # Ensure UI state matches latest generation
                        vm.IsChecked = True
                        if row["status"] in ["CREATE", "MISSING"]:
                            vm._action = "CREATE"
                        else:
                            vm._action = row["status"]
                        self.EditorItems.Add(vm)
                    else:
                        real_id = ElementId(sh_id) if sh_id != -1 else ElementId.InvalidElementId
                        vm = SheetViewModel(real_id, tgt_num, 
                                            row["target_name"] if is_template else row["existing_name"], 
                                            active_collection, row["discipline"], row["cg"], 
                                            series_name=row.get("series_name", "Unknown"),
                                            is_template=is_template, validation_callback=self.run_validation, 
                                            number_changed_callback=self.on_sheet_number_changed,
                                            move_up_callback=self.move_item_up, move_down_callback=self.move_item_down,
                                            context_callback=self.get_sheet_context)
                        vm.OriginalCollectionName = active_collection
                        vm.IsChecked = True
                        if row["status"] in ["CREATE", "MISSING"]:
                            vm._action = "CREATE"
                        else:
                            vm._action = row["status"]
                        self.EditorItems.Add(vm)
            
            for vm in self.EditorItems:
                vm.IsDirty = False
            
            self.update_grid_title()
            self.Txt_GridTitle.Text += " | Items: " + str(len(self.EditorItems))
            self.run_validation()
            
        except Exception as big_e:
            forms.alert("CRITICAL CRASH IN SCHEMA MATCH:\n" + str(big_e))
            
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
        if hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
            node = self.NavTree.SelectedItem
            self._current_selected_node = node
            
        if not node: return
        
        if node.NodeType == "Root":
            title = "All Sheets & Views"
        elif node.NodeType == "Collection":
            title = "Collection: " + node.Tag
        elif node.NodeType == "Discipline":
            c_name, d_name = node.Tag
            title = "Collection: {} | Discipline: {}".format(c_name, d_name)
        elif node.NodeType == "ContentGroup":
            c_name, d_name, cg_name = node.Tag
            title = "Collection: {} | Group: {}".format(c_name, cg_name)
        else:
            title = "Filtered Results"
            
        # Append diagnostic count
        title += " ({} Rows Generated)".format(self.EditorItems.Count)
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
                            
            # Sort the EditorItems list by SheetNumber
            items_list = list(self.EditorItems)
            
            def get_sort_key(s):
                num = s.SheetNumber if s.SheetNumber else ""
                # Pad numbers so string sorting works correctly (e.g. A-010 before A-100)
                # We split the string by numbers and text
                parts = re.split(r'(\d+)', num)
                key = []
                for p in parts:
                    if p.isdigit():
                        key.append(p.zfill(10))
                    else:
                        key.append(p.lower())
                return key
                
            items_list.sort(key=get_sort_key)
            
            self.EditorItems.Clear()
            for item in items_list:
                self.EditorItems.Add(item)
                
        finally:
            self._is_auto_sequencing = False
            self.run_validation()
            
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
            
            def filter_func(item):
                if not show_invalid:
                    return True
                
                has_error = getattr(item, 'IsNameUnique', True) == False or bool(getattr(item, 'ValidationWarning', ""))
                for v in getattr(item, 'Views', []):
                    if getattr(v, 'ValidationWarning', ""):
                        has_error = True
                        break
                return has_error
                
            view.Filter = Predicate[Object](filter_func)
        except Exception as e:
            pass

    def on_filter_changed(self, sender, e):
        self.apply_filters()

    def run_validation(self):
        all_numbers = {}
        has_valid_work = False
        
        # Determine the active collection from NavTree
        active_collection = self.Cmb_TargetCollection.Text if hasattr(self, 'Cmb_TargetCollection') and self.Cmb_TargetCollection.Text else "PERMIT SET"
        if hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
            node = self.NavTree.SelectedItem
            if hasattr(node, "NodeType") and node.NodeType == "Collection":
                active_collection = node.Name
                
        # Combine all real Revit sheets + any active CREATE templates currently in the grid
        validation_pool = set(self.all_grid_nodes)
        for item in self.EditorItems:
            validation_pool.add(item)
            
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
        
        # Collect all actively modified IDs across the entire grid to prevent false clashes
        # when swapping numbers or when an existing sheet vacates its number.
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
                        
        # Setup for global view name validation
        
        existing_view_names = getattr(self, '_cached_global_views', set())

        illegal_chars = r'[\\:\{\}\[\]\|;\<\>\?\~]'
        all_grid_view_names = set()
                
        # Validate view names, numbers, and illegal characters
        has_work = False
        for r in validation_pool:
            if r.IsChecked and r.Action in ["CREATE", "UPDATE", "PURGE"]:
                has_work = True
                
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
                    if not v.LevelName:
                        has_error = True
                        v.ValidationWarning += " Missing Level."
                        
                # Revit allows Legends and Schedules to be placed on multiple sheets
                is_legend_or_schedule = False
                pt = getattr(v, 'PlanType', '').lower()
                if 'legend' in pt or 'schedule' in pt:
                    is_legend_or_schedule = True

                # Check 4: Duplicate View Numbers on the SAME sheet
                v_num = getattr(v, 'ViewNumber', '').strip().lower()
                
                if v_num and not is_legend_or_schedule:
                    if v_num in sheet_view_numbers:
                        has_error = True
                        v.ValidationWarning += " Duplicate View Number on sheet."
                    sheet_view_numbers.add(v_num)
                        
                # Check 5: Duplicate View Names GLOBALLY
                v_name_lower = v.Name.lower() if v.Name else ""
                
                if v_name_lower and v_num and not is_legend_or_schedule:
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
                    r.ValidationBrush = "#FCA5A5"  # Highlight the sheet row in red
                    if not r.ValidationWarning:
                        r.ValidationWarning = "View Error: " + v.ValidationWarning
                    elif "View Error:" not in r.ValidationWarning:
                        r.ValidationWarning += "\nView Error: " + v.ValidationWarning
                    break
                    
            if r.IsChecked and r.Action in ["CREATE", "UPDATE", "PURGE"] and not r_has_error:
                has_valid_work = True
                
        self.main_vm.IsPushEnabled = has_valid_work
        
        # Determine status and generate a synopsis
        has_errors = False
        has_unchecked_work = False
        total_work_items = 0
        
        matched_count = 0
        create_count = 0
        update_count = 0
        
        for r in self.EditorItems:
            if r.Action == "MATCHED":
                matched_count += 1
            elif r.Action in ["CREATE", "MISSING"]:
                create_count += 1
            elif r.Action in ["UPDATE", "RENAME_NAME", "RENAME_NUM", "RENAME_NUMBER", "RENAME_BOTH"]:
                update_count += 1
                
            if r.Action in ["CREATE", "UPDATE", "PURGE", "RENAME_NUM", "RENAME_NUMBER", "RENAME_BOTH"]:
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
        
        validation_pool = set(self.all_grid_nodes)
        for item in self.EditorItems:
            validation_pool.add(item)
            
        valid_nodes = []
        for r in validation_pool:
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
        
        # Print using AIA Schema order
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
        
        # Print any remaining ones that didn't match the schema targets exactly
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
            
            validation_pool = set(self.all_grid_nodes)
            for item in self.EditorItems:
                validation_pool.add(item)
                
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
                    s_coll = "Default"
                    try:
                        p = s.LookupParameter("Sheet Collection")
                        if not p:
                            p = s.LookupParameter(" Sheet Collection")
                        if p and p.HasValue:
                            s_coll = p.AsString()
                    except: pass
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
                    
            # Skip any nodes that clash with the live model
            for r in validation_pool:
                if not r.IsChecked or r.Action == "PURGE": continue
                is_clash = False
                
                # Sheet Number collision
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
                            
                for v in getattr(r, 'Views', []):
                        if getattr(v, 'IsNew', False) and getattr(v, 'Name', '') and v.Name.lower() in live_view_names:
                            is_clash = True
                            error_log.append("Skipped Sheet '{}': View name '{}' already taken by another user.".format(r.SheetNumber, v.Name))
                            break
                            
                if is_clash:
                    r.IsChecked = False
            # ------------------------------------------------
            
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
                            
                    ensure_sheet_parameter(doc, "Discipline")
                    ensure_sheet_parameter(doc, "Content Group")
                    ensure_sheet_parameter(doc, "Sheet Series")
                    ensure_sheet_parameter(doc, " Sheet Collection")
                    
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
                    
                    # Phase 2: Finalize
                    with Transaction(doc, "Phase 2 - Finalize") as t2:
                        t2.Start()
                        try:
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
                                match_stat = getattr(r, 'MatchStatus', None) or r.Action
                                debug_log.append("[Phase 2 Loop 4] Processing: {} -> {} (Action: {}, MatchStat: {})".format(getattr(r, 'OriginalNumber', 'NEW'), r.SheetNumber, r.Action, match_stat))
                                
                                if r.Action == "CREATE":
                                    tb_id = ElementId(tb_id_val)
                                    debug_log.append("  [API] ViewSheet.Create(doc, tb_id={})".format(tb_id_val))
                                    new_sheet = ViewSheet.Create(doc, tb_id)
                                    debug_log.append("  [API] new_sheet.SheetNumber = '{}'".format(r.SheetNumber))
                                    new_sheet.SheetNumber = r.SheetNumber
                                    debug_log.append("  [API] new_sheet.Name = '{}'".format(r.SheetName))
                                    new_sheet.Name = r.SheetName
                                    assign_sheet_to_collection(doc, new_sheet, r.CollectionName)
                                    c_res = classification.classify_sheet(r.SheetNumber, r.SheetName)
                                    disc_name = c_res.get("discipline", "Unknown")
                                    cg_name = c_res.get("contentGroup", "Uncategorized")
                                    set_sheet_parameter(new_sheet, "Discipline", disc_name)
                                    set_sheet_parameter(new_sheet, "Content Group", cg_name)
                                    set_sheet_parameter(new_sheet, "Sheet Series", getattr(r, "SheetSeries", "General"))
                                    
                                    # Apply inferred majority parameters for browser organization
                                    for p_name, p_val in majority_param_values.items():
                                        try:
                                            param = new_sheet.LookupParameter(p_name)
                                            if param and not param.IsReadOnly and not param.AsString():
                                                param.Set(p_val)
                                        except Exception: pass
                                        
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
                                        log_updated.append("{} - {}".format(r.SheetNumber, r.SheetName))
                                        renames += 1
                                elif r.Action == "PURGE":
                                    s_elem = doc.GetElement(r.ElementId)
                                    if s_elem:
                                        assign_sheet_to_collection(doc, s_elem, "PURGE")
                                    log_purged.append("{} - {} (Quarantined)".format(r.SheetNumber, r.SheetName))
                                    purges += 1
                                        
                                if r.Action != "PURGE" and r.IsChecked:
                                    target_sheet_id = r.ElementId if r.Action != "CREATE" else new_sheet.Id
                                    
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
                                        # Determine Shorthand
                                        shorthand = ""
                                        if r.CollectionName and r.CollectionName != "Undefined":
                                            words = r.CollectionName.split()
                                            if len(words) > 1:
                                                shorthand = "".join(w[0] for w in words if w).upper()
                                            else:
                                                shorthand = r.CollectionName[:3].upper()
                                                
                                        view_number_str = v.ViewNumber or "00"
                                        unique_name_parts = [v.Name]
                                        if shorthand: unique_name_parts.append(shorthand)
                                        unique_name_parts.append(view_number_str)
                                        if r.SheetNumber: unique_name_parts.append(r.SheetNumber)
                                        
                                        unique_backend_name = " - ".join(unique_name_parts)
                                        
                                        from Autodesk.Revit.DB import BuiltInParameter
                                        if getattr(v, 'ViewId', ElementId.InvalidElementId) != ElementId.InvalidElementId:
                                            v_elem = doc.GetElement(v.ViewId)
                                            if v_elem:
                                                # Set Title on Sheet (Desired Name)
                                                title_param = v_elem.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                                                if title_param and not title_param.IsReadOnly:
                                                    if title_param.AsString() != v.Name:
                                                        try: title_param.Set(v.Name)
                                                        except: pass
                                                        
                                                # Set Backend Unique Name
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
                                            
                                            if target_vp:
                                                if target_vp.SheetId != target_sheet_id:
                                                    center = target_vp.GetBoxCenter()
                                                    doc.Delete(target_vp.Id)
                                                    if Viewport.CanAddViewToSheet(doc, target_sheet_id, v.ViewId):
                                                        debug_log.append("  [API] Viewport.Create(doc, target_sheet_id, v.ViewId={}, center)".format(v.ViewId.IntegerValue if hasattr(v.ViewId, 'IntegerValue') else v.ViewId.Value))
                                                        new_vp = Viewport.Create(doc, target_sheet_id, v.ViewId, center)
                                                        park_conflicting_dn(doc, target_sheet_id, v.ViewNumber, new_vp.Id)
                                                        dn_param = new_vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                                        if dn_param and not dn_param.IsReadOnly and v.ViewNumber:
                                                            try: 
                                                                debug_log.append("  [API] new_vp.DetailNumber.Set('{}')".format(v.ViewNumber))
                                                                dn_param.Set(str(v.ViewNumber))
                                                            except: pass
                                                else:
                                                    park_conflicting_dn(doc, target_sheet_id, v.ViewNumber, target_vp.Id)
                                                    dn_param = target_vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                                    if dn_param and not dn_param.IsReadOnly and v.ViewNumber:
                                                        if dn_param.AsString() != str(v.ViewNumber):
                                                            try: 
                                                                debug_log.append("  [API] target_vp.DetailNumber.Set('{}')".format(v.ViewNumber))
                                                                dn_param.Set(str(v.ViewNumber))
                                                            except: pass
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
                                                        scale_map = {
                                                            "12\" = 1'-0\"": 1, "6\" = 1'-0\"": 2, "3\" = 1'-0\"": 4, "1 1/2\" = 1'-0\"": 8, "1\" = 1'-0\"": 12, "3/4\" = 1'-0\"": 16, "1/2\" = 1'-0\"": 24, "3/8\" = 1'-0\"": 32, "1/4\" = 1'-0\"": 48, "3/16\" = 1'-0\"": 64, "1/8\" = 1'-0\"": 96, "1\" = 10'-0\"": 120, "3/32\" = 1'-0\"": 128, "1/16\" = 1'-0\"": 192, "1\" = 20'-0\"": 240, "3/64\" = 1'-0\"": 256, "1\" = 30'-0\"": 360, "1/32\" = 1'-0\"": 384, "1\" = 40'-0\"": 480, "1\" = 50'-0\"": 600, "1\" = 60'-0\"": 720, "1/64\" = 1'-0\"": 768, "1\" = 80'-0\"": 960, "1\" = 100'-0\"": 1200, "1\" = 160'-0\"": 1920, "1\" = 200'-0\"": 2400, "1\" = 300'-0\"": 3600, "1\" = 400'-0\"": 4800
                                                        }
                                                        if v.Scale in scale_map:
                                                            view_to_place.Scale = scale_map[v.Scale]
                                            
                                            if view_to_place:
                                                # Set Title on Sheet (Desired Name)
                                                title_param = view_to_place.get_Parameter(BuiltInParameter.VIEW_DESCRIPTION)
                                                if title_param and not title_param.IsReadOnly:
                                                    try: title_param.Set(v.Name)
                                                    except: pass
                                                
                                                # Set Backend Unique Name
                                                try:
                                                    view_to_place.Name = unique_backend_name if v.ViewNumber else v.Name
                                                    renames += 1
                                                    log_view_renamed.append(v.Name)
                                                except: pass
                                                
                                                if Viewport.CanAddViewToSheet(doc, target_sheet_id, view_to_place.Id):
                                                    debug_log.append("  [API] Viewport.Create(doc, target_sheet_id, view_to_place.Id={}, default_center)".format(view_to_place.Id.IntegerValue if hasattr(view_to_place.Id, 'IntegerValue') else view_to_place.Id.Value))
                                                    new_vp = Viewport.Create(doc, target_sheet_id, view_to_place.Id, XYZ(1.5, 1.0, 0))
                                                    park_conflicting_dn(doc, target_sheet_id, v.ViewNumber, new_vp.Id)
                                                    from Autodesk.Revit.DB import BuiltInParameter
                                                    dn_param = new_vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                                    if dn_param and not dn_param.IsReadOnly and v.ViewNumber:
                                                        try: 
                                                            debug_log.append("  [API] new_vp.DetailNumber.Set('{}')".format(v.ViewNumber))
                                                            dn_param.Set(str(v.ViewNumber))
                                                        except: pass
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
