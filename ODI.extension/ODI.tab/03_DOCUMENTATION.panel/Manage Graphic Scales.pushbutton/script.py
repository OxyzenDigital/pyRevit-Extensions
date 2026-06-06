# -*- coding: utf-8 -*-
__title__ = "Manage\nGraphic Scales"
__doc__ = "Pairs and aligns graphic scales to viewports across sheets."
__author__ = "ODI"

import os
import clr
import json
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
import System
from System.Collections.Generic import List
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System import Double
from System.Windows.Media import SolidColorBrush, Color as WpfColor, Colors
from System.Windows.Input import Cursors, ICommand, Key

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, ViewSheetSet, 
    ViewType, FamilySymbol, Family, XYZ, SaveAsOptions, Viewport,
    ViewSheet, StorageType, ElementTransformUtils, BuiltInParameter,
    UnitUtils, SpecTypeId, UnitFormatUtils
)
from pyrevit import revit, forms, script, HOST_APP

doc = revit.doc
uidoc = revit.uidoc
DEFAULT_FAMILY_NAME = "Graphic Scale"
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

# --- Helpers ---
def get_id(element_id):
    """Revit 2024+ compatibility for ElementId."""
    if hasattr(element_id, "Value"):
        return element_id.Value
    return element_id.IntegerValue

def get_view_scale_string(scale_int):
    """Converts Revit scale integer to a readable string (e.g. 96 -> 1/8\" = 1'-0\")."""
    # For standard scales, 12 inches / scale_int gives the fraction.
    if scale_int == 0: return "Custom"
    if scale_int == 1: return "12\" = 1'-0\""
    if scale_int == 96: return "1/8\" = 1'-0\""
    if scale_int == 48: return "1/4\" = 1'-0\""
    if scale_int == 192: return "1/16\" = 1'-0\""
    if scale_int == 384: return "1/32\" = 1'-0\""
    if scale_int == 24: return "1/2\" = 1'-0\""
    if scale_int == 16: return "3/4\" = 1'-0\""
    if scale_int == 32: return "3/8\" = 1'-0\""
    if scale_int == 8: return "1 1/2\" = 1'-0\""
    if scale_int == 4: return "3\" = 1'-0\""
    return "1 : {}".format(scale_int)

def get_view_scale_val(view):
    """Retrieves the scale value from a custom 'Scale Value' parameter, or defaults to view.Scale."""
    p = view.LookupParameter("Scale Value")
    if p and p.HasValue:
        if p.StorageType == StorageType.Integer: return p.AsInteger()
        if p.StorageType == StorageType.Double: return int(p.AsDouble())
    return view.Scale

def get_linked_id(scale_inst):
    """Attempts to retrieve the linked Viewport ID from custom parameters or Comments."""
    for p_name in ["View Id", "Viewport Id"]:
        p = scale_inst.LookupParameter(p_name)
        if p and p.HasValue:
            if p.StorageType == StorageType.String: return p.AsString()
            elif p.StorageType == StorageType.Integer: return str(p.AsInteger())
            elif p.StorageType == StorageType.Double: return str(int(p.AsDouble()))
    p = scale_inst.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if p and p.HasValue: return p.AsString()
    return None

def set_linked_id(scale_inst, vp_id_str):
    """Attempts to save the linked Viewport ID to custom parameters or Comments."""
    for p_name in ["View Id", "Viewport Id"]:
        p = scale_inst.LookupParameter(p_name)
        if p and not p.IsReadOnly:
            if p.StorageType == StorageType.String: 
                p.Set(vp_id_str)
            elif p.StorageType == StorageType.Integer:
                try: p.Set(int(vp_id_str))
                except: pass
            elif p.StorageType == StorageType.Double:
                try: p.Set(float(vp_id_str))
                except: pass
            return
    p = scale_inst.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if p and not p.IsReadOnly:
        p.Set(vp_id_str)

class UnitHelper:
    @staticmethod
    def get_project_length_unit():
        return doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()

    @staticmethod
    def get_unit_symbol():
        try:
            tid = UnitHelper.get_project_length_unit().TypeId
            if "meters" in tid: return "m"
            if "centimeters" in tid: return "cm"
            if "millimeters" in tid: return "mm"
            if "feet" in tid: return "ft"
            if "inches" in tid: return "in"
        except: pass
        return "units"

    @staticmethod
    def to_internal(value_str):
        value_str = str(value_str).strip()
        units = doc.GetUnits()
        parsed_val = clr.Reference[Double]()
        if UnitFormatUtils.TryParse(units, SpecTypeId.Length, value_str, parsed_val):
            return parsed_val.Value
            
        # Fallback fraction and float parsing if native TryParse fails
        try:
            val = 0.0
            if "/" in value_str:
                parts = value_str.split()
                if len(parts) == 2: # e.g., "1 1/2"
                    whole = float(parts[0])
                    num, den = parts[1].split("/")
                    val = whole + (float(num) / float(den)) if whole >= 0 else whole - (float(num) / float(den))
                else: # e.g., "1/2"
                    num, den = value_str.split("/")
                    val = float(num) / float(den)
            else:
                clean_str = ''.join(c for c in value_str if c.isdigit() or c in '.-')
                if clean_str: val = float(clean_str)
                
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertToInternalUnits(val, unit_id)
        except:
            raise ValueError("Invalid unit format")

    @staticmethod
    def to_formatted_string(value_in_internal_units):
        try:
            val = float(value_in_internal_units)
            units = doc.GetUnits()
            return UnitFormatUtils.Format(units, SpecTypeId.Length, val, False)
        except: return "0.0"

# --- Data Models ---
class RelayCommand(ICommand):
    def __init__(self, execute, can_execute=None):
        self._execute = execute
        self._can_execute = can_execute
        self._events = []
    def add_CanExecuteChanged(self, value):
        self._events.append(value)
    def remove_CanExecuteChanged(self, value):
        self._events.remove(value)
    def Execute(self, parameter):
        self._execute(parameter)
    def CanExecute(self, parameter):
        return self._can_execute(parameter) if self._can_execute else True
    def RaiseCanExecuteChanged(self):
        for handler in self._events:
            handler(self, System.EventArgs.Empty)

class ViewModelBase(INotifyPropertyChanged):
    def __init__(self):
        self._events = []
    def add_PropertyChanged(self, value):
        self._events.append(value)
    def remove_PropertyChanged(self, value):
        self._events.remove(value)
    def OnPropertyChanged(self, name):
        for handler in self._events:
            handler(self, PropertyChangedEventArgs(name))

class NodeBase(ViewModelBase):
    def __init__(self, name, vm=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self.vm = vm
        self._is_checked = True
        self._is_expanded = True
        self._is_selected = False
        self.Children = []
        self.NodeType = "Item"
        self.FontWeight = "Normal"

    @property
    def IsSelected(self): return self._is_selected
    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("IsSelected")

    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, value):
        if self._is_checked != value:
            self._is_checked = value
            self.OnPropertyChanged("IsChecked")
            for child in self.Children:
                child.IsChecked = value
            # Cascade to views if this is a SheetNode
            if hasattr(self, "Views"):
                for v in self.Views:
                    v.IsChecked = value
            if self.vm: self.vm.refresh_commands()
        
    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, value):
        if self._is_expanded != value:
            self._is_expanded = value
            self.OnPropertyChanged("IsExpanded")
            if self.vm: self.vm.refresh_commands()

class ViewNode(NodeBase):
    def __init__(self, viewport, view, global_vm):
        NodeBase.__init__(self, view.Name, global_vm)
        self.Viewport = viewport
        self.View = view
        self.NodeType = "View"
        self.FontWeight = "Normal"
        self.ScaleVal = get_view_scale_val(view)
        self.ScaleText = get_view_scale_string(self.ScaleVal)
        self._status = "Missing"
        self.LinkedSymbol = None # Holds the actual Graphic Scale Element if found
        
        detail_param = viewport.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        self.DetailNumber = detail_param.AsString() if detail_param and detail_param.HasValue else "-"
        self.ViewTypeDisplay = str(view.ViewType)
        
        self._local_family = global_vm.SelectedFamily
        self._local_offset_x = global_vm.OffsetX
        self._local_offset_y = global_vm.OffsetY

    @property
    def Status(self): return self._status
    @Status.setter
    def Status(self, value):
        self._status = value
        self.OnPropertyChanged("Status")
        
    @property
    def LocalFamily(self): return self._local_family
    @LocalFamily.setter
    def LocalFamily(self, value):
        self._local_family = value
        self.OnPropertyChanged("LocalFamily")
        
    @property
    def LocalOffsetX(self): return self._local_offset_x
    @LocalOffsetX.setter
    def LocalOffsetX(self, value):
        self._local_offset_x = value
        self.OnPropertyChanged("LocalOffsetX")
        
    @property
    def LocalOffsetY(self): return self._local_offset_y
    @LocalOffsetY.setter
    def LocalOffsetY(self, value):
        self._local_offset_y = value
        self.OnPropertyChanged("LocalOffsetY")

class SheetNode(NodeBase):
    def __init__(self, sheet, vm=None):
        NodeBase.__init__(self, "{} - {}".format(sheet.SheetNumber, sheet.Name), vm)
        self.Sheet = sheet
        self.NodeType = "Sheet"
        self.FontWeight = "SemiBold"
        self.Views = [] # Holds ViewNodes, not added to Children so TreeView ignores them

class SheetSetNode(NodeBase):
    def __init__(self, name, vm=None):
        NodeBase.__init__(self, name, vm)
        self.NodeType = "SheetSet"
        self.FontWeight = "Bold"

class FamilyOption:
    def __init__(self, family, symbol):
        self.Family = family
        self.Symbol = symbol
        self.Name = family.Name

# --- ViewModel ---
class GraphicScaleViewModel(ViewModelBase):
    def __init__(self, window):
        ViewModelBase.__init__(self)
        self.window = window
        self._status_text = "Ready"
        self._current_scope = "active"
        self._sheets = []
        self._families = []
        self._selected_family = None
        self._offset_x = UnitHelper.to_formatted_string((1.0 / 4.0) / 12.0)
        self._offset_y = UnitHelper.to_formatted_string((17.0 / 32.0) / 12.0)
        self._current_views = []
        self.existing_scales_by_sheet = {}
        
        self.load_config()
        
        self.ApplyAllCommand = RelayCommand(self.apply_all, self.can_apply_all)
        self.ApplyCurrentCommand = RelayCommand(self.apply_current, self.can_apply_current)
        self.RemoveAllCommand = RelayCommand(self.remove_all, self.can_remove_all)
        self.RemoveCurrentCommand = RelayCommand(self.remove_current, self.can_remove_current)
        self.CancelCommand = RelayCommand(self.cancel_action)
        self.ScanActiveCommand = RelayCommand(self.scan_active)
        self.ScanProjectCommand = RelayCommand(self.scan_project)
        
        self.SelectAllCommand = RelayCommand(self.select_all, self.can_select_all)
        self.SelectNoneCommand = RelayCommand(self.select_none, self.can_select_none)
        self.ExpandAllCommand = RelayCommand(self.expand_all, self.can_expand_all)
        self.CollapseAllCommand = RelayCommand(self.collapse_all, self.can_collapse_all)
        
        self.load_data()
        
    def refresh_commands(self):
        """Forces the UI buttons to re-evaluate their enabled/disabled states."""
        self.SelectAllCommand.RaiseCanExecuteChanged()
        self.SelectNoneCommand.RaiseCanExecuteChanged()
        self.ExpandAllCommand.RaiseCanExecuteChanged()
        self.CollapseAllCommand.RaiseCanExecuteChanged()
        self.ApplyAllCommand.RaiseCanExecuteChanged()
        self.ApplyCurrentCommand.RaiseCanExecuteChanged()
        self.RemoveAllCommand.RaiseCanExecuteChanged()
        self.RemoveCurrentCommand.RaiseCanExecuteChanged()
        
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    if "OffsetX" in data: self._offset_x = data["OffsetX"]
                    if "OffsetY" in data: self._offset_y = data["OffsetY"]
            except: pass
            
    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump({"OffsetX": self.OffsetX, "OffsetY": self.OffsetY}, f)
        except: pass

    @property
    def Sheets(self): return self._sheets
    @Sheets.setter
    def Sheets(self, value):
        self._sheets = value
        self.OnPropertyChanged("Sheets")
        
    @property
    def CurrentViews(self): return self._current_views
    @CurrentViews.setter
    def CurrentViews(self, value):
        self._current_views = value
        self.OnPropertyChanged("CurrentViews")
        if hasattr(self, 'ApplyCurrentCommand'):
            self.ApplyCurrentCommand.RaiseCanExecuteChanged()
        if hasattr(self, 'RemoveCurrentCommand'):
            self.RemoveCurrentCommand.RaiseCanExecuteChanged()

    @property
    def Families(self): return self._families
    
    @property
    def SelectedFamily(self): return self._selected_family
    @SelectedFamily.setter
    def SelectedFamily(self, value):
        self._selected_family = value
        self.OnPropertyChanged("SelectedFamily")
        # Cascade global family change to all views
        for root in self.Sheets:
            for s_node in root.Children:
                for v_node in s_node.Views:
                    v_node.LocalFamily = value
                    
        if self._current_scope == "active":
            self.scan_active()
        else:
            self.scan_project()
        
    @property
    def OffsetX(self): return self._offset_x
    @OffsetX.setter
    def OffsetX(self, value):
        self._offset_x = value
        self.OnPropertyChanged("OffsetX")
        # Cascade global offset X change to all views
        for root in self.Sheets:
            for s_node in root.Children:
                for v_node in s_node.Views:
                    v_node.LocalOffsetX = value

    @property
    def OffsetY(self): return self._offset_y
    @OffsetY.setter
    def OffsetY(self, value):
        self._offset_y = value
        self.OnPropertyChanged("OffsetY")
        # Cascade global offset Y change to all views
        for root in self.Sheets:
            for s_node in root.Children:
                for v_node in s_node.Views:
                    v_node.LocalOffsetY = value

    @property
    def StatusText(self): return self._status_text
    @StatusText.setter
    def StatusText(self, value):
        self._status_text = value
        self.OnPropertyChanged("StatusText")

    def load_data(self):
        """Loads families and defaults to scanning active view."""
        self.load_families()
        
        if isinstance(doc.ActiveView, ViewSheet):
            self.scan_active()
        else:
            self.scan_project()

    def load_families(self):
        """Finds Annotation families with a 'View Scale' parameter, or loads from Resources."""
        fam_options = []
        
        # Find existing Generic Annotations with a "View Scale" parameter
        symbols = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).OfClass(FamilySymbol).ToElements()
        processed_fams = set()
        
        for sym in symbols:
            fam = sym.Family
            if fam.Id in processed_fams: continue
            
            has_scale_param = False
            for p in sym.Parameters:
                if p.Definition.Name.lower() in ["view scale", "scale value"]:
                    has_scale_param = True
                    break
            
            if has_scale_param or "scale" in fam.Name.lower():
                fam_options.append(FamilyOption(fam, sym))
                processed_fams.add(fam.Id)

        if not fam_options:
            res_dir = os.path.join(os.path.dirname(__file__), 'Resources')
            res_path = None
            if os.path.exists(res_dir):
                for f in os.listdir(res_dir):
                    if f.lower().endswith('.rfa') and 'graphic scale' in f.lower():
                        res_path = os.path.join(res_dir, f)
                        break
            if res_path and os.path.exists(res_path):
                try:
                    with Transaction(doc, "Load Graphic Scale Family") as t:
                        t.Start()
                        loaded_fam_ref = clr.Reference[Family]()
                        if doc.LoadFamily(res_path, loaded_fam_ref):
                            fam = loaded_fam_ref.Value
                            sym_id = list(fam.GetFamilySymbolIds())[0]
                            sym = doc.GetElement(sym_id)
                            if not sym.IsActive: sym.Activate()
                            fam_options.append(FamilyOption(fam, sym))
                        t.Commit()
                except Exception as e:
                    print("Error loading Graphic Scale family: {}".format(e))

        self._families = fam_options
        if fam_options:
            self._selected_family = fam_options[0]
            
    def can_select_all(self, param=None):
        if not self.Sheets: return False
        views = self.get_all_views_in_nodes(self.Sheets)
        if not views: return False
        return any(not v.IsChecked for v in views)
        
    def can_select_none(self, param=None):
        if not self.Sheets: return False
        views = self.get_all_views_in_nodes(self.Sheets)
        if not views: return False
        return any(v.IsChecked for v in views)
        
    def can_expand_all(self, param=None):
        if not self.Sheets: return False
        return self._check_expansion(self.Sheets, target=False)

    def can_collapse_all(self, param=None):
        if not self.Sheets: return False
        return self._check_expansion(self.Sheets, target=True)

    def _check_expansion(self, nodes, target):
        """Returns True if any node with children matches the target expansion state."""
        for node in nodes:
            if node.Children:
                if node.IsExpanded == target: return True
                if self._check_expansion(node.Children, target): return True
        return False

    def select_all(self, param=None):
        if not self.Sheets: return
        for node in self.Sheets:
            node.IsChecked = True
        self.CurrentViews = self.get_all_views_in_nodes(self.Sheets) # Force DataGrid refresh
        self.refresh_commands()
        
    def select_none(self, param=None):
        if not self.Sheets: return
        for node in self.Sheets:
            node.IsChecked = False
        self.CurrentViews = self.get_all_views_in_nodes(self.Sheets) # Force DataGrid refresh
        self.refresh_commands()

    def expand_all(self, param=None): 
        self._set_expansion(True)
        self.refresh_commands()
        
    def collapse_all(self, param=None): 
        self._set_expansion(False)
        self.refresh_commands()

    def _set_expansion(self, state):
        if not self.Sheets: return
        def recurse(node):
            node.IsExpanded = state
            for c in node.Children: recurse(c)
        for root in self.Sheets:
            recurse(root)
            
    def prefetch_project_data(self):
        """Pre-caches all graphic scales in the project to speed up tree generation."""
        fam_name_filter = self.SelectedFamily.Name if self.SelectedFamily else None
        self.existing_scales_by_sheet = {}
        if fam_name_filter:
            annos = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType().ToElements()
            for a in annos:
                try:
                    if a.Symbol.Family.Name == fam_name_filter:
                        sid = str(get_id(a.OwnerViewId))
                        if sid not in self.existing_scales_by_sheet:
                            self.existing_scales_by_sheet[sid] = []
                        self.existing_scales_by_sheet[sid].append(a)
                except: pass

    def get_valid_viewports(self, sheet):
        """Gets Viewports that contain scalable views (Plans, Sections, Drafting, etc.)."""
        valid_types = [
            ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.EngineeringPlan, 
            ViewType.AreaPlan, ViewType.Section, ViewType.Elevation, 
            ViewType.DraftingView, ViewType.Detail
        ]
        
        vps = []
        for vp_id in sheet.GetAllViewports():
            vp = doc.GetElement(vp_id)
            view = doc.GetElement(vp.ViewId)
            if view.ViewType in valid_types:
                vps.append((vp, view))
        return vps

    def build_sheet_node(self, sheet):
        """Constructs a fully populated SheetNode mapped to existing symbols."""
        vps = self.get_valid_viewports(sheet)
        s_node = SheetNode(sheet, self)
        if not vps: return s_node
        
        sid = str(get_id(sheet.Id))
        existing_scales = self.existing_scales_by_sheet.get(sid, [])
        unclaimed_scales = list(existing_scales)
        
        for vp, view in vps:
            v_node = ViewNode(vp, view, self)
            vp_id_str = str(get_id(vp.Id))
            
            matched_symbol = None
            
            # 1. Match by exact Linked ID (Programmatically linked previously)
            for scale_inst in list(unclaimed_scales):
                if get_linked_id(scale_inst) == vp_id_str:
                    matched_symbol = scale_inst
                    unclaimed_scales.remove(scale_inst)
                    break
            
            # 2. Fuzzy Match for manually placed scales (No ID in comments)
            if not matched_symbol and unclaimed_scales:
                for scale_inst in list(unclaimed_scales):
                    if not get_linked_id(scale_inst):
                        matched_symbol = scale_inst
                        unclaimed_scales.remove(scale_inst)
                        break
            
            if matched_symbol:
                v_node.LinkedSymbol = matched_symbol
                
                # Update Local Family based on what it actually is
                sym_fam_name = matched_symbol.Symbol.Family.Name
                for fam_opt in self.Families:
                    if fam_opt.Name == sym_fam_name:
                        v_node.LocalFamily = fam_opt
                        break
                
                scale_param = matched_symbol.LookupParameter("Scale Value")
                if not scale_param:
                    scale_param = matched_symbol.LookupParameter("View Scale")
                if scale_param:
                    val = scale_param.AsInteger() if scale_param.StorageType == StorageType.Integer else int(scale_param.AsDouble())
                    if val == v_node.ScaleVal:
                        v_node.Status = "Match"
                    else:
                        v_node.Status = "Mismatch"
                else:
                    v_node.Status = "Mismatch" # Missing parameter
            else:
                v_node.Status = "Missing"
                
            s_node.Views.append(v_node)
            
        return s_node

    def scan_active(self, param=None):
        selected_name = None
        if self.window.systemTree.SelectedItem:
            selected_name = self.window.systemTree.SelectedItem.Name
            
        self._current_scope = "active"
        active_view = doc.ActiveView
        if isinstance(active_view, ViewSheet):
            self.window.Cursor = Cursors.Wait
            self.prefetch_project_data()
            
            root = SheetSetNode("Current View", self)
            s_node = self.build_sheet_node(active_view)
            if s_node.Views:
                root.Children.append(s_node)
                
            self.Sheets = [root]
            self._restore_selection(selected_name)
            self.StatusText = "Scanned active sheet."
            self.window.Cursor = Cursors.Arrow
            self.refresh_commands()
        else:
            forms.alert("Active view is not a Sheet. Scanning full project instead.")
            self.scan_project()

    def scan_project(self, param=None):
        selected_name = None
        if self.window.systemTree.SelectedItem:
            selected_name = self.window.systemTree.SelectedItem.Name
            
        self._current_scope = "project"
        self.window.Cursor = Cursors.Wait
        self.StatusText = "Scanning project..."
        self.prefetch_project_data()
        
        tree_nodes = []
        
        # All Sheets Node
        all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        all_sheets = sorted(all_sheets, key=lambda s: s.SheetNumber)
        
        all_node = SheetSetNode("< All Sheets >", self)
        for sheet in all_sheets:
            s_node = self.build_sheet_node(sheet)
            if s_node.Views:
                all_node.Children.append(s_node)
        if all_node.Children:
            tree_nodes.append(all_node)
            
        # Print Sets Nodes
        sheet_sets = FilteredElementCollector(doc).OfClass(ViewSheetSet).ToElements()
        for sset in sorted(sheet_sets, key=lambda s: s.Name):
            sset_node = SheetSetNode(sset.Name, self)
            for sheet in sset.Views:
                if isinstance(sheet, ViewSheet):
                    s_node = self.build_sheet_node(sheet)
                    if s_node.Views:
                        sset_node.Children.append(s_node)
            if sset_node.Children:
                tree_nodes.append(sset_node)
                
        self.Sheets = tree_nodes
        self._restore_selection(selected_name)
        self.StatusText = "Scanned {} sheet sets.".format(len(tree_nodes)-1)
        self.window.Cursor = Cursors.Arrow
        self.refresh_commands()

    def _restore_selection(self, selected_name):
        """Restores the previous tree selection and filters the CurrentViews grid."""
        if not selected_name:
            self.CurrentViews = self.get_all_views_in_nodes(self.Sheets)
            return
            
        def find_and_select(nodes):
            for node in nodes:
                if node.Name == selected_name:
                    node.IsSelected = True
                    return node
                if hasattr(node, "Children") and node.Children:
                    res = find_and_select(node.Children)
                    if res: return res
            return None
            
        found_node = find_and_select(self.Sheets)
        if found_node:
            self.CurrentViews = self.get_all_views_in_nodes([found_node])
        else:
            self.CurrentViews = self.get_all_views_in_nodes(self.Sheets)

    def get_all_views_in_nodes(self, nodes):
        """Extracts all ViewNodes recursively for the DataGrid."""
        views = []
        for node in nodes:
            if hasattr(node, "Views") and node.Views:
                views.extend(node.Views)
            elif hasattr(node, "Children") and node.Children:
                views.extend(self.get_all_views_in_nodes(node.Children))
        return views

    def cancel_action(self, parameter):
        self.window.Close()
        
    def get_checked_views(self):
        """Flattens the TreeView items and deduplicates identical viewports across sets."""
        checked_views = {}
        if not self.Sheets: return checked_views
        for root in self.Sheets:
            for s_node in root.Children:
                for v_node in s_node.Views:
                    if v_node.IsChecked:
                        checked_views[get_id(v_node.Viewport.Id)] = v_node
        return list(checked_views.values())

    def can_apply_all(self, param=None):
        if not self.Sheets: return False
        return any(v.IsChecked for v in self.get_all_views_in_nodes(self.Sheets))

    def can_apply_current(self, param=None):
        if not self.CurrentViews: return False
        return any(v.IsChecked for v in self.CurrentViews)

    def can_remove_all(self, param=None):
        if not self.Sheets: return False
        return any(v.IsChecked and v.LinkedSymbol is not None for v in self.get_all_views_in_nodes(self.Sheets))

    def can_remove_current(self, param=None):
        if not self.CurrentViews: return False
        return any(v.IsChecked and v.LinkedSymbol is not None for v in self.CurrentViews)

    def apply_all(self, parameter):
        self._apply_scales_core(self.get_checked_views())

    def apply_current(self, parameter):
        checked_views = {}
        if self.CurrentViews:
            for v_node in self.CurrentViews:
                if v_node.IsChecked:
                    checked_views[get_id(v_node.Viewport.Id)] = v_node
        self._apply_scales_core(list(checked_views.values()))

    def remove_all(self, parameter):
        self._remove_scales_core(self.get_checked_views())

    def remove_current(self, parameter):
        checked_views = {}
        if self.CurrentViews:
            for v_node in self.CurrentViews:
                if v_node.IsChecked:
                    checked_views[get_id(v_node.Viewport.Id)] = v_node
        self._remove_scales_core(list(checked_views.values()))

    def _remove_scales_core(self, checked_views):
        if not checked_views:
            forms.alert("No views selected to update.")
            return

        self.window.Cursor = Cursors.Wait
        self.StatusText = "Removing Graphic Scales..."
        
        try:
            with Transaction(doc, "Remove Graphic Scales") as t:
                t.Start()
                deleted_count = 0
                touched_sheets = set(v_node.Viewport.SheetId for v_node in checked_views)
                
                for sheet_id in touched_sheets:
                    annos = FilteredElementCollector(doc, sheet_id).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType().ToElements()
                    for a in annos:
                        try:
                            # Delete ALL Graphic Scales found on checked sheets regardless of family
                            has_scale_param = False
                            for p in a.Symbol.Parameters:
                                if p.Definition.Name.lower() in ["view scale", "scale value"]:
                                    has_scale_param = True
                                    break
                            if has_scale_param or "scale" in a.Symbol.Family.Name.lower():
                                doc.Delete(a.Id)
                                deleted_count += 1
                        except: pass
                t.Commit()
                
            if self._current_scope == "active":
                self.scan_active()
            else:
                self.scan_project()
            self.StatusText = "Removed: {}.".format(deleted_count)
            forms.alert("Successfully removed {} graphic scales.".format(deleted_count), title="Complete")
            
        except Exception as e:
            self.StatusText = "Error occurred."
            import traceback
            print(traceback.format_exc())
            forms.alert("An error occurred:\n{}".format(e))
        finally:
            self.window.Cursor = Cursors.Arrow

    def _apply_scales_core(self, checked_views):
        if not checked_views:
            forms.alert("No views selected to resolve.")
            return

        self.window.Cursor = Cursors.Wait
        self.StatusText = "Resolving Graphic Scales..."
        
        try:
            with TransactionGroup(doc, "Manage Graphic Scales") as tg:
                tg.Start()
                
                t1 = Transaction(doc, "Create and Update Scales")
                t1.Start()
                
                # Activate all potential symbols
                for fam_opt in self.Families:
                    if not fam_opt.Symbol.IsActive:
                        fam_opt.Symbol.Activate()
                doc.Regenerate()
                
                placed_count = 0
                updated_count = 0
                
                for v_node in checked_views:
                    if not v_node.LocalFamily: continue
                    symbol = v_node.LocalFamily.Symbol
                    sheet = v_node.Viewport.SheetId
                    vp_id_str = str(get_id(v_node.Viewport.Id))
                    
                    inst = v_node.LinkedSymbol
                    
                    if not inst:
                        # Create new
                        inst = doc.Create.NewFamilyInstance(XYZ.Zero, symbol, doc.GetElement(sheet))
                        v_node.LinkedSymbol = inst
                        placed_count += 1
                        
                        # Link it
                        set_linked_id(inst, vp_id_str)
                    else:
                        # Change Symbol if the family type was overridden
                        if inst.Symbol.Id != symbol.Id:
                            inst.Symbol = symbol
                        updated_count += 1

                    # Update View Scale Parameter
                    scale_param = inst.LookupParameter("Scale Value")
                    if not scale_param:
                        scale_param = inst.LookupParameter("View Scale")
                    if scale_param and not scale_param.IsReadOnly:
                        if scale_param.StorageType == StorageType.Integer:
                            scale_param.Set(v_node.ScaleVal)
                        elif scale_param.StorageType == StorageType.Double:
                            scale_param.Set(float(v_node.ScaleVal))

                # Third pass (Moved up): Cleanup Ambiguous/Unchecked scales on modified sheets
                touched_sheets = set(v_node.Viewport.SheetId for v_node in checked_views)
                deleted_count = 0
                
                keep_ids = set(v_node.LinkedSymbol.Id for v_node in checked_views if v_node.LinkedSymbol)
                target_family_names = set(v_node.LocalFamily.Name for v_node in checked_views if v_node.LocalFamily)
                
                for sheet_id in touched_sheets:
                    annos = FilteredElementCollector(doc, sheet_id).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsNotElementType().ToElements()
                    for a in annos:
                        try:
                            fam_name = a.Symbol.Family.Name
                            # If it's a Graphic Scale family and it wasn't claimed by a viewport, delete it!
                            is_scale_fam = any(p.Definition.Name.lower() in ["view scale", "scale value"] for p in a.Symbol.Parameters) or "scale" in fam_name.lower()
                            
                            if is_scale_fam:
                                if a.Id not in keep_ids:
                                    doc.Delete(a.Id)
                                    deleted_count += 1
                        except: pass
                        
                t1.Commit()
                
                t2 = Transaction(doc, "Align Graphic Scales")
                t2.Start()
                doc.Regenerate() # Force geometry update for newly created symbols
                
                # Second pass: Align instances based on their actual resized bounding boxes
                for v_node in checked_views:
                    inst = v_node.LinkedSymbol
                    vp = v_node.Viewport
                    sheet = v_node.Viewport.SheetId
                    
                    user_offset_x = UnitHelper.to_internal(v_node.LocalOffsetX)
                    user_offset_y = UnitHelper.to_internal(v_node.LocalOffsetY)
                    
                    if not inst: continue
                    
                    # 1. Get Viewport Label Location
                    try:
                        vp_outline = vp.GetLabelOutline()
                        target_pt = XYZ(vp_outline.MaximumPoint.X, vp_outline.MinimumPoint.Y, 0)
                    except:
                        box = vp.GetBoxOutline()
                        target_pt = XYZ(box.MaximumPoint.X, box.MinimumPoint.Y, 0)
                        
                    # 2. Determine current insertion point of the graphic scale
                    loc = inst.Location
                    if hasattr(loc, "Point"):
                        current_pt = loc.Point
                        
                        # Use the 'Unit' parameter from the family for the leftward offset
                        unit_offset = 0.0
                        unit_param = inst.LookupParameter("Unit")
                        if unit_param and unit_param.HasValue:
                            if unit_param.StorageType == StorageType.Double:
                                unit_offset = unit_param.AsDouble()
                            elif unit_param.StorageType == StorageType.Integer:
                                unit_offset = float(unit_param.AsInteger())
                            
                        # Apply offsets (Subtract X to move left, Add Y to move up)
                        target_pt = XYZ(target_pt.X - unit_offset - user_offset_x, target_pt.Y + user_offset_y, 0)
                            
                        translation = target_pt - current_pt
                        if not translation.IsZeroLength():
                            ElementTransformUtils.MoveElement(doc, inst.Id, translation)

                t2.Commit()
                
                tg.Assimilate()
                
            if self._current_scope == "active":
                self.scan_active()
            else:
                self.scan_project()
            self.StatusText = "Placed: {}, Updated: {}, Cleaned: {}.".format(placed_count, updated_count, deleted_count)
            forms.alert("Success! Graphic scales processed.\n\nAdded: {}\nUpdated: {}\nCleaned: {}".format(placed_count, updated_count, deleted_count), title="Complete")
            
        except Exception as e:
            self.StatusText = "Error occurred."
            import traceback
            print(traceback.format_exc())
            forms.alert("An error occurred:\n{}".format(e))
        finally:
            self.window.Cursor = Cursors.Arrow

# --- Main UI Class ---
class GraphicScaleWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'ui.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_WinClose.Click += self.close_window
        
        self.systemTree.SelectedItemChanged += self.tree_selection_changed
        self.systemTree.MouseDoubleClick += self.tree_double_click
        self.sysDataGrid.MouseDoubleClick += self.grid_double_click
        self.sysDataGrid.SelectionChanged += self.grid_selection_changed
        
        self.Tb_OffsetX.LostFocus += self.format_textbox
        self.Tb_OffsetX.KeyDown += self.format_textbox
        self.Tb_OffsetY.LostFocus += self.format_textbox
        self.Tb_OffsetY.KeyDown += self.format_textbox
        
        self.apply_revit_theme()
        self.viewModel = GraphicScaleViewModel(self)
        self.DataContext = self.viewModel

        u_sym = UnitHelper.get_unit_symbol()
        self.Title += " [{}]".format(u_sym)

    def tree_selection_changed(self, sender, args):
        """Updates the DataGrid to only show views for the selected Tree item."""
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node:
                # Pass a list containing just the selected node to get all recursive views
                self.viewModel.CurrentViews = self.viewModel.get_all_views_in_nodes([selected_node])
                
                # Auto Zoom to Sheet if checked
                if self.Cb_AutoZoom.IsChecked and hasattr(selected_node, "Sheet"):
                    if doc.ActiveView.Id != selected_node.Sheet.Id:
                        revit.uidoc.ActiveView = selected_node.Sheet
                    elem_ids = List[ElementId]([selected_node.Sheet.Id])
                    revit.uidoc.ShowElements(elem_ids)
                    revit.uidoc.Selection.SetElementIds(elem_ids)
        except Exception:
            pass
            
    def grid_selection_changed(self, sender, args):
        """Auto-zooms to the selected viewport on the sheet when Auto Zoom is checked."""
        try:
            if not self.Cb_AutoZoom.IsChecked:
                return
            selected_items = self.sysDataGrid.SelectedItems
            if not selected_items: return
            
            elem_ids = List[ElementId]()
            sheet_id = None
            for item in selected_items:
                if hasattr(item, "LinkedSymbol") and item.LinkedSymbol:
                    elem_ids.Add(item.LinkedSymbol.Id)
                elif hasattr(item, "Viewport") and item.Viewport:
                    elem_ids.Add(item.Viewport.Id)
                if not sheet_id and hasattr(item, "Viewport") and item.Viewport:
                    sheet_id = item.Viewport.SheetId
                    
            if sheet_id and doc.ActiveView.Id != sheet_id:
                revit.uidoc.ActiveView = doc.GetElement(sheet_id)
                
            if elem_ids:
                revit.uidoc.ShowElements(elem_ids)
                revit.uidoc.Selection.SetElementIds(elem_ids)
        except Exception:
            pass

    def format_textbox(self, sender, args):
        """Hooks into standard project unit string conversion and sets border color."""
        if hasattr(args, "Key") and args.Key != Key.Enter:
            return
        try:
            if hasattr(sender, "Text"):
                val = UnitHelper.to_internal(sender.Text)
                sender.Text = UnitHelper.to_formatted_string(val)
                
                # Ensure View Model gets the updated formatted text
                if sender.Name == "Tb_OffsetX": self.viewModel.OffsetX = sender.Text
                if sender.Name == "Tb_OffsetY": self.viewModel.OffsetY = sender.Text
                
                if hasattr(sender, "BorderBrush"):
                    default_color = WpfColor.FromRgb(75, 85, 99) if getattr(self, "_is_dark", False) else WpfColor.FromRgb(209, 213, 219)
                    sender.BorderBrush = SolidColorBrush(default_color)
        except:
            if hasattr(sender, "BorderBrush"):
                sender.BorderBrush = SolidColorBrush(Colors.Red)

    def drag_window(self, sender, args):
        try: self.DragMove()
        except: pass

    def close_window(self, sender, args):
        self.Close()

    def tree_double_click(self, sender, args):
        """Zooms to the selected sheet in Revit on double-click."""
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node and hasattr(selected_node, "Sheet"):
                if doc.ActiveView.Id != selected_node.Sheet.Id:
                    revit.uidoc.ActiveView = selected_node.Sheet
                elem_ids = List[ElementId]([selected_node.Sheet.Id])
                revit.uidoc.ShowElements(elem_ids)
                revit.uidoc.Selection.SetElementIds(elem_ids)
        except Exception: pass

    def grid_double_click(self, sender, args):
        """Zooms to the selected viewport or graphic scale in Revit on double-click."""
        try:
            selected_items = self.sysDataGrid.SelectedItems
            if selected_items:
                elem_ids = List[ElementId]()
                sheet_id = None
                for item in selected_items:
                    if hasattr(item, "LinkedSymbol") and item.LinkedSymbol:
                        elem_ids.Add(item.LinkedSymbol.Id)
                    elif hasattr(item, "Viewport") and item.Viewport:
                        elem_ids.Add(item.Viewport.Id)
                    if not sheet_id and hasattr(item, "Viewport") and item.Viewport:
                        sheet_id = item.Viewport.SheetId
                        
                if sheet_id and doc.ActiveView.Id != sheet_id:
                    revit.uidoc.ActiveView = doc.GetElement(sheet_id)
                    
                if elem_ids:
                    revit.uidoc.ShowElements(elem_ids)
                    revit.uidoc.Selection.SetElementIds(elem_ids)
        except Exception: pass

    def apply_revit_theme(self):
        """Detects Revit theme and updates window resources if Dark."""
        is_dark = False
        try:
            if int(HOST_APP.version) >= 2024:
                from Autodesk.Revit.UI import UIThemeManager, UITheme
                if UIThemeManager.CurrentTheme == UITheme.Dark:
                    self._is_dark = is_dark = True
        except: pass
        
        if is_dark:
            res = self.Resources
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(45, 52, 64))
            res["ToolbarBrush"] = SolidColorBrush(WpfColor.FromRgb(31, 41, 55))
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(35, 41, 51))
            res["CardBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81))
            res["CardBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 100))
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
            res["TextLightBrush"] = SolidColorBrush(WpfColor.FromRgb(160, 165, 175))
            res["HeaderTextBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 100))
            res["ButtonBrush"] = SolidColorBrush(WpfColor.FromRgb(60, 68, 82))
            res["HoverBrush"] = SolidColorBrush(WpfColor.FromRgb(70, 85, 105))

if __name__ == '__main__':
    GraphicScaleWindow().ShowDialog()
