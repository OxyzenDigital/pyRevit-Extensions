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
from System.Windows.Media import SolidColorBrush, Color as WpfColor, Colors
from System.Windows import MessageBox, Visibility
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

# --- Utility Commands ---
class RelayCommand(ICommand):
    def __init__(self, action):
        self.action = action
    def add_CanExecuteChanged(self, handler): pass
    def remove_CanExecuteChanged(self, handler): pass
    def CanExecute(self, parameter): return True
    def Execute(self, parameter): self.action()

def generate_char_diff(original, current):
    if not original or not current or original == current:
        return False, original or "", "", "", ""
    orig_len = len(original)
    curr_len = len(current)
    min_len = min(orig_len, curr_len)
    
    prefix_len = 0
    while prefix_len < min_len and original[prefix_len] == current[prefix_len]:
        prefix_len += 1
        
    suffix_len = 0
    while suffix_len < min_len - prefix_len and original[orig_len - 1 - suffix_len] == current[curr_len - 1 - suffix_len]:
        suffix_len += 1
        
    prefix = original[:prefix_len]
    suffix = original[orig_len - suffix_len:] if suffix_len > 0 else ""
    old_mid = original[prefix_len:orig_len - suffix_len]
    new_mid = current[prefix_len:curr_len - suffix_len]
    
    return True, prefix, old_mid, new_mid, suffix

# --- ViewModels ---

class ViewModelBase(INotifyPropertyChanged):
    def __init__(self):
        self._property_changed_handlers = []
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    def OnPropertyChanged(self, property_name):
        args = PropertyChangedEventArgs(property_name)
        for handler in self._property_changed_handlers:
            handler(self, args)

class DiffNode(ViewModelBase):
    def __init__(self, original_val, action_callback=None):
        ViewModelBase.__init__(self)
        self._original_val = original_val
        self._proposed_val = original_val
        self._action_callback = action_callback
        
        self.HasDiff = False
        self.Prefix = original_val or ""
        self.Old = ""
        self.New = ""
        self.Suffix = ""
        
    @property
    def OriginalValue(self): return self._original_val
    @property
    def ProposedValue(self): return self._proposed_val
    @ProposedValue.setter
    def ProposedValue(self, val):
        self._proposed_val = val
        self.HasDiff, self.Prefix, self.Old, self.New, self.Suffix = generate_char_diff(self._original_val, val)
        self.OnPropertyChanged("ProposedValue")
        self.OnPropertyChanged("HasDiff")
        self.OnPropertyChanged("Prefix")
        self.OnPropertyChanged("Old")
        self.OnPropertyChanged("New")
        self.OnPropertyChanged("Suffix")
        if self._action_callback:
            self._action_callback()

class SelectableNode(ViewModelBase):
    def __init__(self, name, is_checked=True, callback=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self._is_checked = is_checked
        self.callback = callback
    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = val
        self.OnPropertyChanged("IsChecked")
        if self.callback: self.callback()

class DisciplineGroupNode(ViewModelBase):
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self.Name = name
        self.Children = []

class CollectionGroupNode(ViewModelBase):
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self.Name = name
        self.Children = []  # Holds DisciplineGroupNode instances

class EditableGridRowNode(ViewModelBase):
    def __init__(self, element_id, item_type, original_number, original_name, parent_sheet=None, is_template=False):
        ViewModelBase.__init__(self)
        self.ElementId = element_id
        self.ItemType = item_type # "Sheet" or "View"
        self.IsTemplate = is_template
        self.ParentSheet = parent_sheet
        self.TargetCollectionName = "00. General"
        self.Children = []
        
        self.IsView = (item_type == "View")
        self.ShowExpandToggle = (item_type == "Sheet")
        
        self._is_checked = False
        self._action = "MATCHED" if not is_template else "CREATE"
        self._is_editing = False
        self._is_expanded = True
        
        self.NumberDiffNode = DiffNode(original_number, self.update_action)
        self.NameDiffNode = DiffNode(original_name, self.update_action)
        
        self.EditCommand = RelayCommand(self.toggle_edit)
        self.DoneCommand = RelayCommand(self.toggle_edit)
        self.PurgeCommand = RelayCommand(self.mark_purge)
        self.ExpandCommand = RelayCommand(self.toggle_expand)

    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = val
        self.OnPropertyChanged("IsChecked")

    @property
    def IsEditing(self): return self._is_editing
    @IsEditing.setter
    def IsEditing(self, val):
        self._is_editing = val
        self.OnPropertyChanged("IsEditing")
        
    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, val):
        self._is_expanded = val
        self.OnPropertyChanged("IsExpanded")

    def toggle_edit(self, parameter=None): self.IsEditing = not self.IsEditing
    def mark_purge(self, parameter=None):
        self.Action = "PURGE"
        self.IsChecked = True
        
    def toggle_expand(self, parameter=None):
        self.IsExpanded = not self.IsExpanded

    @property
    def Action(self): return self._action
    @Action.setter
    def Action(self, val):
        self._action = val
        self.OnPropertyChanged("Action")
        self.OnPropertyChanged("ActionColorBrush")

    @property
    def CanPurge(self): return not self.IsTemplate and self.ItemType == "Sheet"

    @property
    def DisciplineName(self):
        match = re.match(r"^([A-Z]+)[- ]?(\d+)", self.NumberDiffNode.OriginalValue.upper())
        disc_code = match.group(1) if match else "Other"
        disc_map = {
            "A": "Architectural",
            "S": "Structural",
            "M": "Mechanical",
            "E": "Electrical",
            "P": "Plumbing",
            "C": "Civil",
            "L": "Landscape",
            "F": "Fire Protection",
            "G": "General",
            "I": "Interiors"
        }
        name = disc_map.get(disc_code, "Discipline")
        return "{} - {}".format(disc_code, name) if disc_code != "Other" else "Uncategorized"

    @property
    def ActionColorBrush(self):
        if self._action == "UPDATE": return SolidColorBrush(Colors.Orange)
        if self._action == "CREATE": return SolidColorBrush(Colors.Green)
        if self._action == "PURGE": return SolidColorBrush(Colors.Red)
        return SolidColorBrush(Colors.LightGray)

    def update_action(self):
        if self.IsTemplate: return
        if self.NumberDiffNode.HasDiff or self.NameDiffNode.HasDiff:
            if self._action != "UPDATE":
                self.Action = "UPDATE"
                self.IsChecked = True
        else:
            if self._action == "UPDATE":
                self.Action = "MATCHED"
                self.IsChecked = False

class NavTreeNode(ViewModelBase):
    def __init__(self, name, node_type, sheet_id=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self.NodeType = node_type
        self.SheetId = sheet_id
        self.Children = ObservableCollection[NavTreeNode]()
        self._is_expanded = True
        self._is_selected = False
        self._count = 0
        
    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, val):
        self._is_expanded = val
        self.OnPropertyChanged("IsExpanded")

    @property
    def IsSelected(self): return self._is_selected
    @IsSelected.setter
    def IsSelected(self, val):
        self._is_selected = val
        self.OnPropertyChanged("IsSelected")
        
    @property
    def Count(self): return self._count
    @Count.setter
    def Count(self, val):
        self._count = val
        self.OnPropertyChanged("Count")
        self.OnPropertyChanged("ShowCount")
        self.OnPropertyChanged("DisplayCount")
        
    @property
    def ShowCount(self): return self.Count > 0
    @property
    def DisplayCount(self):
        if self.NodeType == "Sheet": return " ({} Views)".format(self.Count)
        return " ({})".format(self.Count)
        
    @property
    def FontWeight(self): return "Bold" if self.NodeType in ["Root", "Collection"] else "Normal"

# --- Revit Utility Functions ---

def get_sheet_collection_name(sheet):
    if hasattr(ParameterTypeId, "SheetCollection"):
        param = sheet.GetParameter(ParameterTypeId.SheetCollection)
        if param and param.AsElementId() != ElementId.InvalidElementId:
            col_elem = doc.GetElement(param.AsElementId())
            if col_elem: return col_elem.Name
    else:
        param = sheet.LookupParameter("Sheet Collection")
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
            
        param = sheet.LookupParameter("Sheet Collection")
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

def generate_discipline_sheets(disc_code, levels, rows, cols):
    targets = []
    suffixes = generate_suffixes(rows, cols)
    
    # 0. General Cover
    targets.append({"num": "{}-001".format(disc_code), "name": "Cover Sheet / Index", "collection": "00. General", "type": "0"})
    
    # 1. Plans
    for idx, lvl in enumerate(levels):
        seq = (idx + 1) * 10
        base_num = "{}-1{:02d}".format(disc_code, seq)
        
        if len(suffixes) == 1:
            targets.append({"num": base_num, "name": "Floor Plan {}".format(lvl), "collection": "01. Plans", "type": "1"})
        else:
            targets.append({"num": base_num, "name": "Floor Plan {} - OVERALL".format(lvl), "collection": "01. Plans", "type": "1"})
            for s in suffixes:
                targets.append({"num": "{}{}".format(base_num, s), "name": "Enlarged Plan {} Part {}".format(lvl, s), "collection": "01. Plans", "type": "1"})
                
    # 2. Elevations
    targets.append({"num": "{}-201".format(disc_code), "name": "Elevations", "collection": "02. Elevations", "type": "2"})
    # 3. Sections
    targets.append({"num": "{}-301".format(disc_code), "name": "Sections", "collection": "03. Sections", "type": "3"})
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

# --- Main Window ---

class ManageSheetsWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'ui.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
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
        
        # Expand/Collapse Handlers
        self.Btn_ExpandNav.Click += lambda s,e: self.toggle_tree(self.NavRoot, True)
        self.Btn_CollapseNav.Click += lambda s,e: self.toggle_tree(self.NavRoot, False)
        self.Btn_ExpandEditor.Click += lambda s,e: self.toggle_editor(True)
        self.Btn_CollapseEditor.Click += lambda s,e: self.toggle_editor(False)
        
        # Real-time Generator binds
        self.Txt_GridRows.TextChanged += self.trigger_generation
        self.Txt_GridCols.TextChanged += self.trigger_generation

        # Data Models
        self.NavRoot = ObservableCollection[NavTreeNode]()
        self.NavTree.ItemsSource = self.NavRoot
        self.NavTree.SelectedItemChanged += self.on_tree_selection_changed
        
        self.TargetSchemaRoot = ObservableCollection[NavTreeNode]()
        self.TargetSchemaTree.ItemsSource = self.TargetSchemaRoot
        
        self.EditorItems = ObservableCollection[EditableGridRowNode]()
        self.EditorTree.ItemsSource = self.EditorItems
        
        self.LevelNodes = ObservableCollection[SelectableNode]()
        self.List_Levels.ItemsSource = self.LevelNodes
        
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
    
    def on_tab_changed(self, sender, e):
        if self.MainTabControl.SelectedIndex == 0:
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

    def load_settings(self):
        self.Txt_GridRows.Text = str(cfg.get_option("grid_rows", 1))
        self.Txt_GridCols.Text = str(cfg.get_option("grid_cols", 1))
        
        saved_discs = cfg.get_option("disciplines", ["A", "M", "E", "P"])
        for k, v in DISCIPLINE_CODES.items():
            is_chk = k in saved_discs
            self.DisciplineNodes.Add(SelectableNode("{} - {}".format(k, v), is_checked=is_chk, callback=self.generate_target_schema))

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
                
        active_levels = [lvl.Name for lvl in self.LevelNodes if lvl.IsChecked]
        
        self.generated_targets = []
        for d in selected_discs:
            self.generated_targets.extend(generate_discipline_sheets(d, active_levels, r, c))
            
        self.TargetSchemaRoot.Clear()
        t_root = NavTreeNode("AIA Schema", "Root")
        self.TargetSchemaRoot.Add(t_root)
        c_map = {}
        for t in self.generated_targets:
            c_name = t["collection"]
            if c_name not in c_map:
                cn = NavTreeNode(c_name, "Collection")
                c_map[c_name] = cn
                t_root.Children.Add(cn)
            sn = NavTreeNode("{} - {}".format(t["num"], t["name"]), "Sheet")
            c_map[c_name].Children.Add(sn)
            t_root.Count += 1
            c_map[c_name].Count += 1
            
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
        for s in sheets:
            if s.IsPlaceholder: continue
            
            c_name = get_sheet_collection_name(s)
            if c_name not in col_map:
                col_node = NavTreeNode(c_name, "Collection")
                col_map[c_name] = col_node
                root_node.Children.Add(col_node)
            else:
                col_node = col_map[c_name]
                
            sh_nav = NavTreeNode("{} - {}".format(s.SheetNumber, s.Name), "Sheet", s.Id)
            col_node.Children.Add(sh_nav)
            root_node.Count += 1
            col_node.Count += 1
            
            sh_row = EditableGridRowNode(s.Id, "Sheet", s.SheetNumber, s.Name)
            sh_row.TargetCollectionName = c_name
            self.all_grid_nodes.append(sh_row)
            
            views = 0
            for v_id in s.GetAllPlacedViews():
                v = doc.GetElement(v_id)
                if not v or v.ViewType in [ViewType.Schedule, ViewType.Legend, ViewType.PanelSchedule]: continue
                v_row = EditableGridRowNode(v.Id, "View", "", v.Name, parent_sheet=sh_row)
                sh_row.Children.append(v_row)
                views += 1
            sh_nav.Count = views # Populates the ({x} Views) suffix
            
    def on_tree_selection_changed(self, sender, e):
        node = self.NavTree.SelectedItem
        if not node: return
        
        valid_sheets = []
        if node.NodeType == "Root":
            self.Txt_GridTitle.Text = "All Sheets & Views"
            valid_sheets = self.all_grid_nodes
                
        elif node.NodeType == "Collection":
            self.Txt_GridTitle.Text = "Collection: " + node.Name
            valid_sheets = [s for s in self.all_grid_nodes if s.TargetCollectionName == node.Name]
                    
        elif node.NodeType == "Sheet":
            self.Txt_GridTitle.Text = "Sheet: " + node.Name
            valid_sheets = [s for s in self.all_grid_nodes if s.ElementId == node.SheetId]
            
        groups = {}
        for s in valid_sheets:
            disc = s.DisciplineName
            if disc not in groups:
                groups[disc] = DisciplineGroupNode(disc)
            groups[disc].Children.append(s)
            
        self.EditorItems.Clear()
        # Sort by discipline name (A, C, E, etc.)
        for d in sorted(groups.keys()):
            self.EditorItems.Add(groups[d])

    def run_fuzzy_match(self, sender, e):
        if not self.generated_targets: return
        
        mapped_count = 0
        for r in self.all_grid_nodes:
            if r.IsTemplate: continue
            
            live_num = r.NumberDiffNode.OriginalValue
            live_name = r.NameDiffNode.OriginalValue
            
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
                r.NameDiffNode.ProposedValue = best_match["name"]
                r.NumberDiffNode.ProposedValue = best_match["num"]
                r.TargetCollectionName = best_match["collection"]
                mapped_count += 1
                
        MessageBox.Show("Smart Fuzzy match complete!\nSuccessfully mapped {} live sheets to Target Schema.".format(mapped_count), "Match Results")
                    
    def add_missing(self, sender, e):
        if not self.generated_targets: return
        existing_numbers = set([r.NumberDiffNode.ProposedValue for r in self.all_grid_nodes])
                
        added = 0
        for t in self.generated_targets:
            if t["num"] not in existing_numbers:
                new_sh = EditableGridRowNode(ElementId.InvalidElementId, "Sheet", t["num"], t["name"], is_template=True)
                new_sh.TargetCollectionName = t["collection"]
                new_sh.IsChecked = True
                self.all_grid_nodes.append(new_sh)
                
                node = self.NavTree.SelectedItem
                if node and (node.NodeType == "Root" or (node.NodeType == "Collection" and node.Name == t["collection"])):
                    self.EditorItems.Add(new_sh)
                    
                existing_numbers.add(t["num"])
                added += 1
                
        if added > 0:
            MessageBox.Show("Added {} missing sheets from AIA Schema to the Grid.".format(added), "Info")

    def sync_to_revit(self, sender, e):
        renames, creates, purges = 0, 0, 0
        
        with Transaction(doc, "Manage Sheets Sync") as t:
            t.Start()
            for r in self.all_grid_nodes:
                if r.IsChecked:
                    if r.Action == "UPDATE" or r.Action == "MATCHED":
                        s_elem = doc.GetElement(r.ElementId)
                        if s_elem:
                            if r.NumberDiffNode.HasDiff: s_elem.SheetNumber = r.NumberDiffNode.ProposedValue
                            if r.NameDiffNode.HasDiff: s_elem.Name = r.NameDiffNode.ProposedValue
                            assign_sheet_to_collection(doc, s_elem, r.TargetCollectionName)
                            renames += 1
                    elif r.Action == "CREATE":
                        titleblocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
                        if titleblocks:
                            new_sheet = ViewSheet.Create(doc, titleblocks[0].Id)
                            new_sheet.SheetNumber = r.NumberDiffNode.ProposedValue
                            new_sheet.Name = r.NameDiffNode.ProposedValue
                            assign_sheet_to_collection(doc, new_sheet, r.TargetCollectionName)
                            creates += 1
                    elif r.Action == "PURGE":
                        doc.Delete(r.ElementId)
                        purges += 1
                        
                if r.Action != "PURGE":
                    for cv in r.Children:
                        if cv.IsChecked and cv.Action == "UPDATE" and cv.NameDiffNode.HasDiff:
                            v_elem = doc.GetElement(cv.ElementId)
                            if v_elem:
                                v_elem.Name = cv.NameDiffNode.ProposedValue
                                renames += 1
            t.Commit()
            
        MessageBox.Show("Push Completed!\nCreated: {}\nUpdated: {}\nPurged: {}".format(creates, renames, purges), "Success")
        self.Close()

if __name__ == '__main__':
    ManageSheetsWindow().ShowDialog()
