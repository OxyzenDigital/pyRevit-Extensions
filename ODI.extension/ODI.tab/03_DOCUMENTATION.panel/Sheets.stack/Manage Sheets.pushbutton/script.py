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
from pyrevit import revit, forms, script

doc = revit.doc
uidoc = revit.uidoc
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

# Intelligent Keyword Synonym Dictionary based on UDS Sheet Types
SHEET_TYPE_SYNONYMS = {
    "0": ["COVER", "INDEX", "GENERAL", "NOTES", "SYMBOLS"],
    "1": ["PLAN", "DEMOLITION", "LAYOUT", "RCP", "REFLECTED", "CEILING", "FRAMING", "LEVEL", "FLOOR", "ROOF", "SITE", "GRADING", "OVERALL", "ENLARGED PLAN"],
    "2": ["ELEVATION", "EXTERIOR", "INTERIOR", "FACADE", "PROFILE"],
    "3": ["SECTION", "CROSS SECTION", "WALL SECTION", "BUILDING SECTION"],
    "4": ["ENLARGED", "LARGE SCALE", "RISER"],
    "5": ["DETAIL", "TYPICAL", "CONNECTION", "ASSEMBLY"],
    "6": ["SCHEDULE", "LEGEND", "DIAGRAM", "ABBREVIATION", "KEY"],
    "7": [],
    "8": [],
    "9": ["3D", "ISOMETRIC", "PERSPECTIVE", "RENDERING", "AXONOMETRIC"]
}

from data_model import *

# --- Revit Utility Functions ---

def get_sheet_collection_name(sheet):
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

# --- Generation Logic ---

def generate_suffixes(rows, cols):
    total = rows * cols
    if total <= 1: return [""]
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    suffixes = []
    for i in range(total):
        if i < 26: suffixes.append(alphabet[i])
        else: suffixes.append("A" + alphabet[i-26])
    return suffixes

def generate_discipline_sheets(disc_code, levels, active_series, rows, cols):
    targets = []
    suffixes = generate_suffixes(rows, cols)
    
    if "0" in active_series:
        targets.append({"num": "{}-001".format(disc_code), "name": "Cover Sheet / Index", "collection": "00. General", "type": "0"})
    
    if "1" in active_series:
        for idx, lvl in enumerate(levels):
            seq = (idx + 1) * 10
            base_num = "{}-1{:02d}".format(disc_code, seq)
            
            if len(suffixes) == 1:
                targets.append({"num": base_num, "name": "Floor Plan {}".format(lvl), "collection": "01. Plans", "type": "1"})
            else:
                targets.append({"num": base_num, "name": "Floor Plan {} - OVERALL".format(lvl), "collection": "01. Plans", "type": "1"})
                for s in suffixes:
                    targets.append({"num": "{}{}".format(base_num, s), "name": "Enlarged Plan {} Part {}".format(lvl, s), "collection": "01. Plans", "type": "1"})
                    
    if "2" in active_series:
        targets.append({"num": "{}-201".format(disc_code), "name": "Elevations", "collection": "02. Elevations", "type": "2"})
    if "3" in active_series:
        targets.append({"num": "{}-301".format(disc_code), "name": "Sections", "collection": "03. Sections", "type": "3"})
    if "5" in active_series:
        targets.append({"num": "{}-501".format(disc_code), "name": "Details", "collection": "05. Details", "type": "5"})
    if "6" in active_series:
        targets.append({"num": "{}-601".format(disc_code), "name": "Schedules", "collection": "06. Schedules", "type": "6"})
    if "7" in active_series:
        targets.append({"num": "{}-701".format(disc_code), "name": "Diagrams", "collection": "07. Diagrams", "type": "7"})
    if "9" in active_series:
        targets.append({"num": "{}-901".format(disc_code), "name": "3D Views", "collection": "09. 3D Views", "type": "9"})
        
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
    synonyms = SHEET_TYPE_SYNONYMS.get(target_type, [])
    has_synonym = False
    
    for syn in synonyms:
        if syn in live_name_upper:
            has_synonym = True
            score += 0.5 # 50% bonus for hitting a specific keyword
            break
            
    # Penalty: If it hit no synonyms for its type, but hits synonyms for a DIFFERENT type
    if not has_synonym:
        for t_type, syn_list in SHEET_TYPE_SYNONYMS.items():
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

class ManageSheetsWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'ui.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.apply_theme()
        
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_WinClose.Click += self.close_window
        
        # Grid/Tabs
        self.Btn_RunMatch.Click += self.run_fuzzy_match
        self.Btn_AddMissing.Click += self.add_missing
        self.Btn_Push.Click += self.sync_to_revit
        
        self.MainTabControl.SelectionChanged += self.on_tab_changed
        
        # Setup Lists Select All
        self.Btn_DiscAll.Click += lambda s,e: self.toggle_list(self.DisciplineNodes, True)
        self.Btn_DiscNone.Click += lambda s,e: self.toggle_list(self.DisciplineNodes, False)
        self.Btn_LevelAll.Click += lambda s,e: self.toggle_list(self.LevelNodes, True)
        self.Btn_LevelNone.Click += lambda s,e: self.toggle_list(self.LevelNodes, False)
        self.Btn_SeriesAll.Click += lambda s,e: self.toggle_list(self.SeriesNodes, True)
        self.Btn_SeriesNone.Click += lambda s,e: self.toggle_list(self.SeriesNodes, False)
        
        # Expand/Collapse Handlers
        self.Btn_ExpandNav.Click += lambda s,e: self.toggle_tree(self.NavRoot, True)
        self.Btn_CollapseNav.Click += lambda s,e: self.toggle_tree(self.NavRoot, False)
        self.Btn_ExpandEditor.Click += lambda s,e: self.toggle_editor(True)
        self.Btn_CollapseEditor.Click += lambda s,e: self.toggle_editor(False)
        
        # Real-time Generator binds
        self.Txt_GridRows.TextChanged += self.trigger_generation
        self.Txt_GridCols.TextChanged += self.trigger_generation

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
        
        self.ModifierNodes = ObservableCollection[SelectableNode]()
        self.List_Modifiers.ItemsSource = self.ModifierNodes
        self.DisciplineNodes = ObservableCollection[SelectableNode]()
        self.List_Disciplines.ItemsSource = self.DisciplineNodes
        
        # Internal Storage
        self.all_grid_nodes = []
        self.generated_targets = []
        
        self.load_revit_data()
        self.load_settings()
        self.generate_target_schema() # Initial run

    def drag_window(self, sender, e): self.DragMove()
    def close_window(self, sender, e): self.Close()
    
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

    def load_settings(self):
        self.Txt_GridRows.Text = str(cfg.get_option("grid_rows", 1))
        self.Txt_GridCols.Text = str(cfg.get_option("grid_cols", 1))
        
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

    def save_settings(self):
        try:
            cfg.grid_rows = int(self.Txt_GridRows.Text)
            cfg.grid_cols = int(self.Txt_GridCols.Text)
        except: pass
        
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
        
        script.save_config()
        
    def trigger_generation(self, sender, e):
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
        
        self.generated_targets = []
        for d in selected_discs:
            self.generated_targets.extend(generate_discipline_sheets(d, active_levels, active_series, r, c))
            
        self.TargetSchemaRoot.Clear()
        t_root = NavTreeNode("AIA Schema", "Root")
        self.TargetSchemaRoot.Add(t_root)
        
        d_map = {}
        c_map = {}
        
        for t in self.generated_targets:
            match = re.match(r"^([A-Z]+)[- ]?(\d+)", t["num"].upper())
            disc_code = match.group(1) if match else "Other"
            disc_dict = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors" }
            disc_name = "{} - {}".format(disc_code, disc_dict.get(disc_code, "Discipline")) if disc_code != "Other" else "Uncategorized"
            
            if disc_name not in d_map:
                dn = NavTreeNode(disc_name, "Discipline", tag=disc_name)
                d_map[disc_name] = dn
                t_root.Children.Add(dn)
                
            c_name = t["collection"]
            coll_key = (disc_name, c_name)
            
            if coll_key not in c_map:
                cn = NavTreeNode(c_name, "Collection", tag=disc_name)
                c_map[coll_key] = cn
                d_map[disc_name].Children.Add(cn)
                
            t_root.Count += 1
            d_map[disc_name].Count += 1
            c_map[coll_key].Count += 1
            
        if not self.generated_targets:
            self.Btn_RunMatch.IsEnabled = False
        else:
            self.Btn_RunMatch.IsEnabled = True

    def load_revit_data(self):
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        for lvl in sorted(levels, key=lambda l: l.Elevation):
            self.LevelNodes.Add(SelectableNode(lvl.Name, True, self.generate_target_schema))

        selected_ids = uidoc.Selection.GetElementIds()
        scope_ids = [i for i in selected_ids if isinstance(doc.GetElement(i), ViewSheet)]
        if not scope_ids: sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        else: sheets = [doc.GetElement(i) for i in scope_ids]
            
        root_node = NavTreeNode("All Sheets", "Root")
        self.NavRoot.Add(root_node)
        
        col_map = {}
        disc_map = {}
        for s in sheets:
            if s.IsPlaceholder: continue
            
            c_name = get_sheet_collection_name(s)
            if c_name not in col_map:
                col_node = NavTreeNode(c_name, "Collection")
                col_map[c_name] = col_node
                root_node.Children.Add(col_node)
            else:
                col_node = col_map[c_name]
                
            sh_row = SheetViewModel(s.Id, s.SheetNumber, s.Name, c_name, validation_callback=self.run_validation)
            self.all_grid_nodes.append(sh_row)
            
            disc_name = sh_row.DisciplineName
            disc_key = (c_name, disc_name)
            if disc_key not in disc_map:
                d_node = NavTreeNode(disc_name, "Discipline", tag=c_name)
                disc_map[disc_key] = d_node
                col_node.Children.Add(d_node)
            else:
                d_node = disc_map[disc_key]
                
            root_node.Count += 1
            col_node.Count += 1
            d_node.Count += 1
            
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
                            renames += 1
                    elif r.Action == "CREATE":
                        titleblocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
                        if titleblocks:
                            new_sheet = ViewSheet.Create(doc, titleblocks[0].Id)
                            new_sheet.SheetNumber = r.SheetNumber
                            new_sheet.Name = r.SheetName
                            assign_sheet_to_collection(doc, new_sheet, r.CollectionName)
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

if __name__ == '__main__':
    ManageSheetsWindow().ShowDialog()
