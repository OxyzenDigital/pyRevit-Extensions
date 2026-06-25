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
from System.Windows import MessageBox, Visibility, SystemColors
from System.Windows.Controls import CheckBox
from System.Windows.Input import ICommand

from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory, 
    ElementId, ViewSheet, View, ViewType, BuiltInParameter, ParameterTypeId,
    Level, ForgeTypeId
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
    return "00. General"

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

def generate_discipline_sheets(disc_code, levels, active_series, active_modifiers, global_cover, rows, cols, naming_scheme, custom_schemes=None):
    targets = []
    suffixes = generate_suffixes(rows, cols, naming_scheme, custom_schemes)
    
    if "0" in active_series:
        if global_cover:
            if disc_code == "CS":
                targets.append({"num": "CS-001", "name": "Cover Sheet", "collection": "00. General", "cg": "00. General", "type": "0"})
                targets.append({"num": "CS-002", "name": "General Notes", "collection": "00. General", "cg": "00. General", "type": "0"})
        else:
            targets.append({"num": "{}-001".format(disc_code), "name": "{} Cover Sheet / Index".format(disc_code), "collection": "00. General", "cg": "00. General", "type": "0"})
    
    for series in active_series:
        if series == "0": continue
        
        series_mods = []
        for disc, groups in classification.CLASSIFICATION_DICT.items():
            for cg, types in groups.items():
                for item in types:
                    if len(item) >= 2:
                        m_name, m_code = item[0], str(item[1])
                        if m_name in active_modifiers and m_code.startswith(series):
                            series_mods.append((m_name, m_code, cg))
        
        if not series_mods:
            series_mods.append(("{} (Default)".format(SERIES_MAP.get(series, series)), series + "01", "0{}. {}".format(series, SERIES_MAP.get(series, series))))
            
        if series == "1":
            for idx, lvl in enumerate(levels):
                lvl_seq = (idx + 1) * 10
                for mod_idx, (m_name, m_code, cg) in enumerate(series_mods):
                    mod_seq = lvl_seq + mod_idx
                    base_num = "{}-{}{:02d}".format(disc_code, series, mod_seq)
                    
                    if len(suffixes) == 1:
                        targets.append({"num": base_num, "name": "{} {}".format(m_name, lvl), "collection": "01. Plans", "cg": cg, "type": "1"})
                    else:
                        targets.append({"num": base_num, "name": "{} {} - OVERALL".format(m_name, lvl), "collection": "01. Plans", "cg": cg, "type": "1"})
                        for let, name_part in suffixes:
                            targets.append({"num": "{}{}".format(base_num, let), "name": "{} {} - {}".format(m_name, lvl, name_part), "collection": "01. Plans", "cg": cg, "type": "1"})
                            
        else:
            for mod_idx, (m_name, m_code, cg) in enumerate(series_mods):
                base_num = "{}-{}{:02d}".format(disc_code, series, mod_idx + 1)
                coll_name = "0{}. {}".format(series, SERIES_MAP.get(series, series))
                targets.append({"num": base_num, "name": m_name, "collection": coll_name, "cg": cg, "type": series})
                
    return targets

# --- Intelligent Fuzzy Match Algorithm ---
def calculate_smart_score(live_num, live_name, target):
    """Calculates a match score using structural parsing and keyword weighting."""
    score = 0.0
    live_name_upper = live_name.upper()
    target_name_upper = target["name"].upper()
    
    # 1. Base Sequence Matcher Score (0.0 to 1.0)
    base_score = difflib.SequenceMatcher(None, live_name_upper, target_name_upper).ratio()
    score += base_score * 0.4 # Weight base score at 40%
    
    # 2. Exact Number Match (Massive Bonus)
    if live_num.upper() == target["num"].upper():
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
            
    # Penalty: If it hit no synonyms for its type, but hits synonyms for a DIFFERENT type
    if not has_synonym:
        for t_type, syn_list in classification.SHEET_TYPE_SYNONYMS.items():
            if t_type == target_type: continue
            for syn in syn_list:
                if syn in live_name_upper:
                    score -= 0.3 # 30% penalty for contradicting keywords (e.g. Demolition PLAN shouldn't match ELEVATION)
                    break
                    
    # 4. Partial Number Match Bonus (e.g., A-101 vs A-101A)
    if live_num.upper()[:4] == target["num"].upper()[:4]:
        score += 0.2
        
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

class ManageSheetsPanel(forms.WPFPanel):
    panel_title = "Manage Sheets"
    panel_id = "f4c9c1a5-8e3b-4b1a-a123-0d6b7c8e9f22"
    panel_source = os.path.join(os.path.dirname(__file__), "ui.xaml")

    def __init__(self):
        forms.WPFPanel.__init__(self)
        
        self.apply_theme()
        self.last_loaded_doc_hash = None
        
        # Subscribe to visibility to know when opened
        self.IsVisibleChanged += self.on_visible_changed
        
        # Subscribe to Revit ViewActivated to catch project switching
        try:
            from pyrevit import HOST_APP
            HOST_APP.uiapp.ViewActivated += self.on_view_activated
        except:
            pass
            

        self.Btn_Refresh.Click += self.on_refresh_clicked
        self.Btn_RefreshData.Click += self.on_refresh_clicked
        self.Btn_ResetAll.Click += self.on_refresh_clicked
        self.Btn_RunMatch.Click += self.run_fuzzy_match
        self.Btn_AddMissing.Click += self.add_missing
        self.Btn_Push.Click += self.sync_to_revit
        self.Btn_EditNamingSchemes.Click += self.on_edit_naming_schemes
        
        self.MainTabControl.SelectionChanged += self.on_tab_changed
        
        # Setup Lists Select All
        self.Btn_DiscAll.Click += lambda s,e: self.toggle_list(self.DisciplineNodes, True)
        self.Btn_DiscNone.Click += lambda s,e: self.toggle_list(self.DisciplineNodes, False)
        self.Btn_LevelAll.Click += lambda s,e: self.toggle_list(self.LevelNodes, True)
        self.Btn_LevelNone.Click += lambda s,e: self.toggle_list(self.LevelNodes, False)
        self.Btn_SeriesAll.Click += lambda s,e: self.toggle_list(self.SeriesNodes, True)
        self.Btn_SeriesNone.Click += lambda s,e: self.toggle_list(self.SeriesNodes, False)
        self.Btn_ModAll.Click += lambda s,e: self.check_tree(self.ModifierRoot, True)
        self.Btn_ModNone.Click += lambda s,e: self.check_tree(self.ModifierRoot, False)
        self.Btn_AddModifier.Click += self.on_add_modifier
        
        # Expand/Collapse Handlers
        self.Btn_ExpandNav.Click += lambda s,e: self.toggle_tree(self.NavRoot, True)
        self.Btn_CollapseNav.Click += lambda s,e: self.toggle_tree(self.NavRoot, False)
        self.Btn_ExpandEditor.Click += lambda s,e: self.toggle_editor(True)
        self.Btn_CollapseEditor.Click += lambda s,e: self.toggle_editor(False)
        
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

        # Data Models
        self.NavRoot = ObservableCollection[NavTreeNode]()
        self.NavTree.ItemsSource = self.NavRoot
        self.NavTree.SelectedItemChanged += self.on_tree_selection_changed
        
        self.TargetSchemaRoot = ObservableCollection[NavTreeNode]()
        self.TargetSchemaTree.ItemsSource = self.TargetSchemaRoot
        
        self.EditorItems = ObservableCollection[object]()
        self.EditorGrid.ItemsSource = self.EditorItems
        
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
        
        # self.load_revit_data() is now called via the Refresh button
        self.load_settings()
        self.generate_target_schema() # Initial run

    def on_refresh_clicked(self, sender, e):
        self.load_revit_data()

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
                
            # Count sheets
            from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet
            from System.Windows import Visibility
            sheet_count = FilteredElementCollector(doc).OfClass(ViewSheet).ToElementIds().Count
            
            if sheet_count > 150:
                # Pause auto-load, show warning
                self.AutoRefreshWarningPanel.Visibility = Visibility.Visible
            else:
                # Auto-load
                self.AutoRefreshWarningPanel.Visibility = Visibility.Collapsed
                self.load_revit_data()
        except Exception as e:
            pass


    
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
                "ErrorTextBrush": "#FECACA"   # Light pink/red text for contrast
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
        if self.MainTabControl.SelectedIndex == 1:
            self.FooterBorder.Visibility = Visibility.Visible
        else:
            self.FooterBorder.Visibility = Visibility.Collapsed
            
    def toggle_list(self, coll, state):
        for node in coll: node.IsChecked = state
            
    def toggle_tree(self, coll, state):
        for node in coll:
            node.IsExpanded = state
            if hasattr(node, "Children") and node.Children:
                self.toggle_tree(node.Children, state)
                
    def toggle_editor(self, state):
        for node in self.all_grid_nodes:
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
        self.generate_target_schema()
            
    def _recursive_check(self, node, is_checked):
        node.IsChecked = is_checked
        if hasattr(node, "Children"):
            for child in node.Children:
                self._recursive_check(child, is_checked)

    def load_settings(self):
        self.Sld_GridRows.Value = float(cfg.get_option("grid_rows", 1))
        self.Sld_GridCols.Value = float(cfg.get_option("grid_cols", 1))
        
        saved_discs = cfg.get_option("disciplines", ["A", "M", "E", "P"])
        aia_order = ['CS', 'G', 'H', 'V', 'B', 'C', 'L', 'S', 'A', 'I', 'Q', 'F', 'P', 'D', 'M', 'E', 'W', 'T', 'R', 'X', 'Z', 'O']
        sorted_disciplines = sorted(DISCIPLINE_CODES.items(), key=lambda x: aia_order.index(x[0]) if x[0] in aia_order else 999)
        for k, v in sorted_disciplines:
            is_chk = k in saved_discs
            self.DisciplineNodes.Add(SelectableNode("{} - {}".format(k, v), is_checked=is_chk, callback=self.generate_target_schema))
            
        saved_series = cfg.get_option("series", ["0", "1", "2", "3"])
        sorted_series = sorted(SERIES_MAP.items(), key=lambda x: x[0])
        for k, v in sorted_series:
            is_chk = k in saved_series
            self.SeriesNodes.Add(SelectableNode("{} - {}".format(k, v), is_checked=is_chk, callback=self.generate_target_schema))
            
        saved_modifiers = cfg.get_option("selected_modifiers", [])
        custom_modifiers = cfg.get_option("custom_modifiers", [])
        self.Chk_GlobalCover.IsChecked = cfg.get_option("global_cover", False)
        
        # Load Naming Schemes
        self._loaded_naming_schemes = project_settings.load_naming_schemes(revit.doc)
        if not self._loaded_naming_schemes:
            self._loaded_naming_schemes = classification.NAMING_SCHEMES
            
        self.Cmb_NamingScheme.ItemsSource = self._loaded_naming_schemes.keys()
        self.Cmb_NamingScheme.SelectedItem = cfg.get_option("naming_scheme", "Segment-Based")
        self.Cmb_NamingScheme.LostFocus += self.trigger_generation
        self.Cmb_NamingScheme.SelectionChanged += self.trigger_generation
        
        self.ModifierRoot.Clear()
        for disc, groups in classification.CLASSIFICATION_DICT.items():
            disc_node = NavTreeNode(disc, "ContentGroup")
            for cg, types in groups.items():
                for item in types:
                    if len(item) >= 2:
                        t_name, t_code = item[0], str(item[1])
                        m_node = NavTreeNode(t_name, "Modifier")
                        m_node.Tag = t_code
                        m_node.callback = self.generate_target_schema
                        m_node.IsChecked = t_name in saved_modifiers
                        disc_node.Children.Add(m_node)
            if disc_node.Children:
                self.ModifierRoot.Add(disc_node)
            
        # Add custom modifiers to a "Custom" group
        if custom_modifiers:
            custom_node = NavTreeNode("Custom", "ContentGroup")
            for mod in custom_modifiers:
                m_node = NavTreeNode(mod, "Modifier")
                m_node.callback = self.generate_target_schema
                m_node.IsChecked = mod in saved_modifiers
                custom_node.Children.Add(m_node)
            self.ModifierRoot.Add(custom_node)

    def save_settings(self):
        try:
            cfg.grid_rows = int(self.Sld_GridRows.Value)
            cfg.grid_cols = int(self.Sld_GridCols.Value)
        except: pass
        cfg.global_cover = bool(self.Chk_GlobalCover.IsChecked)
        
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
        
        for cg_node in self.ModifierRoot:
            if cg_node.Name == "Custom":
                for m_node in cg_node.Children:
                    if m_node.IsChecked:
                        selected_modifiers.append(m_node.Name)
                    all_custom.append(m_node.Name)
            else:
                for m_node in cg_node.Children:
                    if m_node.IsChecked:
                        selected_modifiers.append(m_node.Name)
                        
        cfg.selected_modifiers = selected_modifiers
        cfg.custom_modifiers = all_custom
        cfg.global_cover = bool(self.Chk_GlobalCover.IsChecked)
        
        script.save_config()
        
    def on_add_modifier(self, sender, e):
        txt = self.Txt_NewModifier.Text.strip()
        if not txt: return
        
        custom_node = None
        for cg_node in self.ModifierRoot:
            if cg_node.Name == "Custom":
                custom_node = cg_node
            for m_node in cg_node.Children:
                if m_node.Name.lower() == txt.lower():
                    m_node.IsChecked = True
                    self.Txt_NewModifier.Text = ""
                    self.generate_target_schema()
                    return
                    
        if not custom_node:
            custom_node = NavTreeNode("Custom", "ContentGroup")
            self.ModifierRoot.Add(custom_node)
            
        m_node = NavTreeNode(txt, "Modifier")
        m_node.callback = self.generate_target_schema
        m_node.IsChecked = True
        custom_node.Children.Add(m_node)
        self.Txt_NewModifier.Text = ""
        self.save_settings()
        self.generate_target_schema()
        
    def trigger_generation(self, sender, e):
        self.generate_target_schema()

    def on_edit_naming_schemes(self, sender, e):
        dialog = NamingSchemeSettingsDialog(self._loaded_naming_schemes)
        dialog.Owner = self.Parent
        if dialog.ShowDialog():
            # Update Document with new schemes
            project_settings.save_naming_schemes(revit.doc, dialog.schemes_dict)
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
        for cg_node in self.ModifierRoot:
            for m_node in cg_node.Children:
                if m_node.IsChecked:
                    active_modifiers.append(m_node.Name)
                    
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
                    
        self.generated_targets = []
        for d in selected_discs:
            self.generated_targets.extend(generate_discipline_sheets(d, active_levels, active_series, active_modifiers, global_cover, r, c, naming_scheme, self._loaded_naming_schemes))
            
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
                d_map[disc_name] = dn
                t_root.Children.Add(dn)
                
            cg_name = t.get("cg", "Unknown Content Group")
            cg_key = (disc_name, cg_name)
            
            if cg_key not in cg_map:
                cn = NavTreeNode(cg_name, "ContentGroup", tag=disc_name)
                cg_map[cg_key] = cn
                d_map[disc_name].Children.Add(cn)
                
            sn = NavTreeNode("{} - {}".format(t["num"], t["name"]), "Sheet", tag=t["num"])
            cg_map[cg_key].Children.Add(sn)
            
            t_root.Count += 1
            d_map[disc_name].Count += 1
            cg_map[cg_key].Count += 1
            
        if not self.generated_targets:
            self.Btn_RunMatch.IsEnabled = False
        else:
            self.Btn_RunMatch.IsEnabled = True
            
        # Auto-expand the target schema tree so user sees the new combinations immediately
        if hasattr(self, 'toggle_tree'):
            self.toggle_tree(self.TargetSchemaRoot, True)

    def load_revit_data(self):
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        if not doc or not uidoc: return
        
        # Cache this document to prevent duplicate loading
        self.last_loaded_doc_hash = doc.PathName + "_" + doc.Title
        
        # Hide the warning if manually refreshed
        from System.Windows import Visibility
        self.AutoRefreshWarningPanel.Visibility = Visibility.Collapsed

        
        self.LevelNodes.Clear()
        self.all_grid_nodes = []
        self.NavRoot.Clear()
        self.EditorItems.Clear()
        
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        for lvl in sorted(levels, key=lambda l: l.Elevation):
            self.LevelNodes.Add(SelectableNode(lvl.Name, True, self.generate_target_schema))

        selected_ids = uidoc.Selection.GetElementIds()
        scope_ids = [i for i in selected_ids if isinstance(doc.GetElement(i), ViewSheet)]
        if not scope_ids: sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        else: sheets = [doc.GetElement(i) for i in scope_ids]
            
        root_node = NavTreeNode("All Sheets", "Root")
        self.NavRoot.Add(root_node)
        
        cg_map = {}
        disc_map = {}
        for s in sheets:
            if s.IsPlaceholder: continue
            
            c_name = get_sheet_collection_name(doc, s)
            
            disc_name, cg_name, draw_type = classification.classify_sheet(s.SheetNumber, s.Name)
            
            if disc_name not in disc_map:
                d_node = NavTreeNode(disc_name, "Discipline", tag=disc_name)
                disc_map[disc_name] = d_node
                root_node.Children.Add(d_node)
            else:
                d_node = disc_map[disc_name]
                
            cg_key = (disc_name, cg_name)
            if cg_key not in cg_map:
                cg_node = NavTreeNode(cg_name, "ContentGroup", tag=disc_name)
                cg_map[cg_key] = cg_node
                d_node.Children.Add(cg_node)
            else:
                cg_node = cg_map[cg_key]
                
            sh_row = SheetViewModel(s.Id, s.SheetNumber, s.Name, c_name, validation_callback=self.run_validation)
            self.all_grid_nodes.append(sh_row)
            
            root_node.Count += 1
            d_node.Count += 1
            cg_node.Count += 1
            
            views = 0
            for v_id in s.GetAllPlacedViews():
                v = doc.GetElement(v_id)
                if not v or v.ViewType in [ViewType.Schedule, ViewType.Legend, ViewType.PanelSchedule]: continue
                v_row = ViewViewModel(v.Id, v.Name, str(v.ViewType))
                sh_row.Views.Add(v_row)
                views += 1
        self.run_validation()
            
    def on_tree_selection_changed(self, sender, e):
        node = self.NavTree.SelectedItem
        if not node: return
        
        valid_sheets = []
        if node.NodeType == "Root":
            self.Txt_GridTitle.Text = "All Sheets & Views"
            valid_sheets = self.all_grid_nodes
                
        elif node.NodeType == "Collection":
            self.Txt_GridTitle.Text = "Collection: " + node.Name
            valid_sheets = [s for s in self.all_grid_nodes if s.CollectionName == node.Name]
                    
        elif node.NodeType == "Discipline":
            self.Txt_GridTitle.Text = "Discipline: " + node.Name
            valid_sheets = [s for s in self.all_grid_nodes if s.DisciplineName == node.Name and s.CollectionName == node.Tag]
            
        self.EditorItems.Clear()
        for s in valid_sheets:
            self.EditorItems.Add(s)

    def run_fuzzy_match(self, sender, e):
        if not self.generated_targets: return
        
        mapped_count = 0
        for r in self.all_grid_nodes:
            if r.IsTemplate: continue
            
            live_num = r.OriginalNumber
            live_name = r.OriginalName
            
            # 1. Parse Discipline Prefix from Live Sheet Number (e.g. 'A' from 'A-101' or 'MH' from 'MH201')
            match = re.match(r"^([A-Z]+)[- ]?(\d+)", live_num.upper())
            live_disc = match.group(1) if match else None
            
            # 2. Filter Targets based on Discipline Prefix
            valid_targets = []
            if live_disc:
                valid_targets = [t for t in self.generated_targets if t["num"].startswith(live_disc)]
            
            # Fallback: if no valid targets found with that discipline, scan all (user might have messed up numbers)
            if not valid_targets:
                valid_targets = self.generated_targets

            best_match = None
            best_score = 0.0
            
            for t in valid_targets:
                score = calculate_smart_score(live_num, live_name, t)
                if score > best_score:
                    best_score = score
                    best_match = t
                    
            if best_match and best_score > 0.4: # Lowered threshold slightly because penalties can drag scores down
                r.SheetName = best_match["name"]
                r.SheetNumber = best_match["num"]
                r.CollectionName = best_match["collection"]
                mapped_count += 1
                
        self.run_validation()
        MessageBox.Show("Smart Fuzzy match complete!\nSuccessfully mapped {} live sheets to Target Schema.".format(mapped_count), "Match Results")
                    
    def add_missing(self, sender, e):
        if not self.generated_targets: return
        existing_numbers = set([r.SheetNumber for r in self.all_grid_nodes])
                
        added = 0
        for t in self.generated_targets:
            if t["num"] not in existing_numbers:
                new_sh = SheetViewModel(ElementId.InvalidElementId, t["num"], t["name"], t["collection"], is_template=True, validation_callback=self.run_validation)
                new_sh.IsChecked = True
                self.all_grid_nodes.append(new_sh)
                
                node = self.NavTree.SelectedItem
                if node and (node.NodeType == "Root" or (node.NodeType == "Collection" and node.Name == t["collection"])):
                    self.EditorItems.Add(new_sh)
                    
                existing_numbers.add(t["num"])
                added += 1
                
        if added > 0:
            MessageBox.Show("Added {} missing sheets from AIA Schema to the Grid.".format(added), "Info")
        self.run_validation()

    def run_validation(self):
        all_numbers = {}
        for r in self.all_grid_nodes:
            if r.Action == "PURGE": continue
            num = str(r.SheetNumber).strip().lower()
            if num not in all_numbers:
                all_numbers[num] = []
            all_numbers[num].append(r)
            
        has_error = False
        for num, items in all_numbers.items():
            if len(items) > 1:
                has_error = True
                for i in items: i.IsNameUnique = False
            else:
                for i in items: i.IsNameUnique = True
                
        self.Btn_Push.IsEnabled = not has_error

    def sync_to_revit(self, sender, e):
        def _sync_action():
            uidoc = __revit__.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            if not doc: return
            renames, creates, purges = 0, 0, 0
            with Transaction(doc, "Manage Sheets Sync") as t:
                t.Start()
                for r in self.all_grid_nodes:
                    if r.IsChecked:
                        if r.Action == "UPDATE" or r.Action == "MATCHED":
                            s_elem = doc.GetElement(r.ElementId)
                            if s_elem:
                                if r.SheetNumber != r.OriginalNumber: s_elem.SheetNumber = r.SheetNumber
                                if r.SheetName != r.OriginalName: s_elem.Name = r.SheetName
                                assign_sheet_to_collection(doc, s_elem, r.CollectionName)
                                disc_name, cg_name, _ = classification.classify_sheet(r.SheetNumber, r.SheetName)
                                set_sheet_parameter(s_elem, "Discipline", disc_name)
                                set_sheet_parameter(s_elem, "Content Group", cg_name)
                                renames += 1
                        elif r.Action == "CREATE":
                            titleblocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
                            if titleblocks:
                                new_sheet = ViewSheet.Create(doc, titleblocks[0].Id)
                                new_sheet.SheetNumber = r.SheetNumber
                                new_sheet.Name = r.SheetName
                                assign_sheet_to_collection(doc, new_sheet, r.CollectionName)
                                disc_name, cg_name, _ = classification.classify_sheet(r.SheetNumber, r.SheetName)
                                set_sheet_parameter(new_sheet, "Discipline", disc_name)
                                set_sheet_parameter(new_sheet, "Content Group", cg_name)
                                creates += 1
                        elif r.Action == "PURGE":
                            doc.Delete(r.ElementId)
                            purges += 1
                            
                    if r.Action != "PURGE":
                        for v in r.Views:
                            if v.ViewId != ElementId.InvalidElementId:
                                v_elem = doc.GetElement(v.ViewId)
                                if v_elem and v_elem.Name != v.Name:
                                    try:
                                        v_elem.Name = v.Name
                                        renames += 1
                                    except: pass
                            elif v._is_new:
                                # TODO: Phase 2: Create new view using v.ViewType and v.Scale
                                # TODO: Phase 3: Place viewport on sheet
                                pass
                                
                t.Commit()
                
                MessageBox.Show("Sync Complete!\nRenamed: {}\nCreated: {}\nPurged: {}".format(renames, creates, purges), "Success")
            self.all_grid_nodes = []
            self.EditorItems.Clear()
            self.NavRoot.Clear()
            self.load_revit_data()
        
        from pyrevit.revit.events import execute_in_revit_context
        execute_in_revit_context("Manage Sheets Sync", _sync_action)

