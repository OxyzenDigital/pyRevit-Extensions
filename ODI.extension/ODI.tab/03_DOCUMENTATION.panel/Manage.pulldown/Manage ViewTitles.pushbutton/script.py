# -*- coding: utf-8 -*-
__title__ = "Manage\nViews & Titles"
__doc__ = "Snaps viewports to a grid and forces view titles to an exact coordinate relative to the Title Block."
__author__ = "ODI"

import os
import clr
import json
import traceback
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
from System.Windows.Controls import ContextMenu, MenuItem

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction, ViewSheetSet, 
    ViewType, XYZ, Viewport, ViewSheet, ElementTransformUtils, BuiltInParameter, TransactionGroup,
    UnitUtils, SpecTypeId, UnitFormatUtils, ElementId, ElementType, ScheduleSheetInstance, BoundingBoxXYZ
)
from pyrevit import revit, forms, script, HOST_APP

doc = revit.doc
uidoc = revit.uidoc
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "manage_views_config.json")

# --- Helpers ---
def get_id(element_id):
    if hasattr(element_id, "Value"): return element_id.Value
    return element_id.IntegerValue

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
        try:
            val = float(value_str)
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertToInternalUnits(val, unit_id)
        except: raise ValueError("Invalid unit format")

    @staticmethod
    def parse_unit_to_internal(value_str):
        value_str = str(value_str).strip()
        # Force evaluation as project unit if no explicit unit is present
        test_str = value_str
        sym = UnitHelper.get_unit_symbol()
        if sym in ['in', 'ft'] and not any(c in test_str for c in ['"', "'", 'm', 'c', 'f']):
            test_str += '"' if sym == 'in' else "'"

        units = doc.GetUnits()
        parsed_val = clr.Reference[Double]()
        if UnitFormatUtils.TryParse(units, SpecTypeId.Length, test_str, parsed_val):
            return parsed_val.Value

        try:
            val = 0.0
            if "/" in value_str:
                parts = value_str.split()
                if len(parts) == 2:
                    val = float(parts[0]) + (float(parts[1].split("/")[0]) / float(parts[1].split("/")[1])) * (1 if float(parts[0]) >= 0 else -1)
                else:
                    val = float(value_str.split("/")[0]) / float(value_str.split("/")[1])
            else:
                clean_str = ''.join(c for c in value_str if c.isdigit() or c in '.-')
                if clean_str: val = float(clean_str)
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertToInternalUnits(val, unit_id)
        except: raise ValueError("Invalid unit format")

    @staticmethod
    def to_formatted_string(value_in_internal_units):
        try:
            val = float(value_in_internal_units)
            units = doc.GetUnits()
            return UnitFormatUtils.Format(units, SpecTypeId.Length, val, False)
        except: return "0.0"

    @staticmethod
    def to_formatted_string_with_symbol(value_in_internal_units):
        # Return a user-friendly formatted string; append unit symbol if format doesn't include it
        try:
            formatted = UnitHelper.to_formatted_string(value_in_internal_units)
            # If formatted string already contains unit markers, return as-is
            if any(u in formatted for u in ['m', 'mm', 'cm', 'ft', 'in', '"', "'"]):
                return formatted
            sym = UnitHelper.get_unit_symbol()
            return "{} {}".format(formatted, sym)
        except: return "0.0"

    @staticmethod
    def to_display_value(value_in_internal_units):
        try:
            val = float(value_in_internal_units)
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertFromInternalUnits(val, unit_id)
        except: return 0.0

    @staticmethod
    def is_metric():
        try:
            sym = UnitHelper.get_unit_symbol()
            return sym in ['m', 'cm', 'mm']
        except: return False

    @staticmethod
    def get_sheet_length_unit():
        # Use project length unit for sheet-related conversions as a sensible default
        return UnitHelper.get_project_length_unit()

# --- Data Models ---
class RelayCommand(ICommand):
    def __init__(self, execute, can_execute=None):
        self._execute = execute
        self._can_execute = can_execute
        self._events = []
    def add_CanExecuteChanged(self, value): self._events.append(value)
    def remove_CanExecuteChanged(self, value): self._events.remove(value)
    def Execute(self, parameter): self._execute(parameter)
    def CanExecute(self, parameter): return self._can_execute(parameter) if self._can_execute else True
    def RaiseCanExecuteChanged(self):
        for handler in self._events: handler(self, System.EventArgs.Empty)

class ViewModelBase(INotifyPropertyChanged):
    def __init__(self): self._events = []
    def add_PropertyChanged(self, value): self._events.append(value)
    def remove_PropertyChanged(self, value): self._events.remove(value)
    def OnPropertyChanged(self, name):
        for handler in self._events: handler(self, PropertyChangedEventArgs(name))

class NodeBase(ViewModelBase):
    def __init__(self, name, vm=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self.vm = vm
        self._is_checked = True
        self._is_expanded = True
        self._is_selected = False
        self.ParentNode = None
        self.Children = []
        self.NodeType = "Item"
        self.FontWeight = "Normal"

    def set_checked(self, value, cascade_down=True, cascade_up=True):
        if self._is_checked != value:
            self._is_checked = value
            self.OnPropertyChanged("IsChecked")
            if cascade_down:
                for child in self.Children: child.set_checked(value, True, False)
                if hasattr(self, "Views"):
                    for v in self.Views:
                        if value and not getattr(v, "HasTitleNumber", True):
                            continue
                        v.set_checked(value, True, False)
            if cascade_up and self.ParentNode: self.ParentNode.evaluate_checked_state()
            if self.vm: self.vm.refresh_commands()

    def evaluate_checked_state(self):
        all_states = []
        if hasattr(self, "Children") and self.Children:
            all_states.extend([c.IsChecked for c in self.Children if getattr(c, "HasTitleNumber", True)])
        if hasattr(self, "Views") and self.Views:
            all_states.extend([v.IsChecked for v in self.Views if getattr(v, "HasTitleNumber", True)])
        if not all_states: return
        new_state = True if all(s == True for s in all_states) else False if all(s == False for s in all_states) else None
        if self._is_checked != new_state:
            self._is_checked = new_state
            self.OnPropertyChanged("IsChecked")
            if self.ParentNode: self.ParentNode.evaluate_checked_state()
            if self.vm: self.vm.refresh_commands()

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
        if value is None: value = False
        self.set_checked(value, True, True)
        
    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, value):
        if self._is_expanded != value:
            self._is_expanded = value
            self.OnPropertyChanged("IsExpanded")

class ViewNode(NodeBase):
    def __init__(self, viewport, view, global_vm):
        NodeBase.__init__(self, view.Name, global_vm)
        self.Viewport = viewport
        self.View = view
        self.NodeType = "View"
        
        detail_param = viewport.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        detail_value = detail_param.AsString() if detail_param and detail_param.HasValue else ""
        detail_value = detail_value.strip() if detail_value else ""
        if detail_value and detail_value not in ["-", "---"] and any(ch.isdigit() for ch in detail_value):
            self._detail_number = detail_value
            self._has_title_number = True
            self._title_status = "-"
        else:
            self._detail_number = "-"
            self._has_title_number = False
            self._title_status = "No Title Number"
            self._is_checked = False
        self.ViewTypeDisplay = str(view.ViewType)

    @property
    def TitleStatus(self): return self._title_status
    @TitleStatus.setter
    def TitleStatus(self, value):
        self._title_status = value
        self.OnPropertyChanged("TitleStatus")

    @property
    def DetailNumber(self): return self._detail_number
    @DetailNumber.setter
    def DetailNumber(self, value):
        self._detail_number = value
        self.OnPropertyChanged("DetailNumber")

    @property
    def HasTitleNumber(self):
        return getattr(self, "_has_title_number", False)

    @property
    def IsEnabled(self):
        return self.HasTitleNumber

class SheetNode(NodeBase):
    def __init__(self, sheet, vm=None):
        NodeBase.__init__(self, "{} - {}".format(sheet.SheetNumber, sheet.Name), vm)
        self.Sheet = sheet
        self.NodeType = "Sheet"
        self.FontWeight = "SemiBold"
        self.Views = []

class SheetSetNode(NodeBase):
    def __init__(self, name, vm=None):
        NodeBase.__init__(self, name, vm)
        self.NodeType = "SheetSet"
        self.FontWeight = "Bold"

# --- ViewModel ---
class ManageViewsViewModel(ViewModelBase):
    def __init__(self, window):
        ViewModelBase.__init__(self)
        self.window = window
        self._status_text = "Ready"
        self._sheets = []
        self._current_views = []
        # Show sensible defaults in inches initially (will reformat to project units on first edit)
        self._grid_size = '1"'
        self._title_offset_x = "0' - 1 5/8\""
        self._title_offset_y = "0' - 0 27/32\""
        self._snap_to_grid = True
        
        self._anchor_options = ["Bottom-Left", "Bottom-Right", "Top-Left", "Top-Right"]
        self._anchor_corner = "Bottom-Left"
        self._right_margin = '10"'
        self._top_margin = '0"'
        self._bottom_margin = '0"'
        self._view_padding = '1"'
        self._run_view_cleanup = True
        self._sync_detail_numbers = True
        
        self.load_config()
        
        self.ArrangeTitlesCommand = RelayCommand(self.arrange_titles, self.can_act)
        self.AutoPackCommand = RelayCommand(self.auto_pack, self.can_act)
        self.ScanActiveCommand = RelayCommand(self.scan_active)
        self.ScanProjectCommand = RelayCommand(self.scan_project)
        self.CancelCommand = RelayCommand(self.cancel_action)
        
        self.SelectAllCommand = RelayCommand(self.select_all, self.can_select_all)
        self.SelectNoneCommand = RelayCommand(self.select_none, self.can_select_none)
        self.ExpandAllCommand = RelayCommand(self.expand_all)
        self.CollapseAllCommand = RelayCommand(self.collapse_all)
        
        if isinstance(doc.ActiveView, ViewSheet): self.scan_active()
        else: self.scan_project()
        
    def refresh_commands(self):
        self.SelectAllCommand.RaiseCanExecuteChanged()
        self.SelectNoneCommand.RaiseCanExecuteChanged()
        self.ArrangeTitlesCommand.RaiseCanExecuteChanged()
        self.AutoPackCommand.RaiseCanExecuteChanged()
        
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    def parse_to_unit(val_str):
                        try:
                            internal = UnitHelper.parse_unit_to_internal(val_str)
                            return UnitHelper.to_formatted_string_with_symbol(internal)
                        except ValueError:
                            return "0\""
                        
                    if "GridSize" in data: self._grid_size = parse_to_unit(data["GridSize"])
                    if "TitleOffsetX" in data: self._title_offset_x = parse_to_unit(data["TitleOffsetX"])
                    if "TitleOffsetY" in data: self._title_offset_y = parse_to_unit(data["TitleOffsetY"])
                    if "SnapToGrid" in data: self._snap_to_grid = data["SnapToGrid"]
                    if "AnchorCorner" in data: self._anchor_corner = data["AnchorCorner"]
                    if "RightMargin" in data: self._right_margin = parse_to_unit(data["RightMargin"])
                    if "TopMargin" in data: self._top_margin = parse_to_unit(data["TopMargin"])
                    if "BottomMargin" in data: self._bottom_margin = parse_to_unit(data["BottomMargin"])
                    if "ViewPadding" in data: self._view_padding = parse_to_unit(data["ViewPadding"])
                    if "RunViewCleanup" in data: self._run_view_cleanup = data["RunViewCleanup"]
                    if "SyncDetailNumbers" in data: self._sync_detail_numbers = data["SyncDetailNumbers"]
                    if "WinTop" in data and "WinLeft" in data:
                        self.window.WindowStartupLocation = 0
                        self.window.Top = data["WinTop"]
                        self.window.Left = data["WinLeft"]
                    if "WinWidth" in data: self.window.Width = data["WinWidth"]
                    if "WinHeight" in data: self.window.Height = data["WinHeight"]
            except: pass
            
    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump({
                    "GridSize": self.GridSize,
                    "TitleOffsetX": self.TitleOffsetX,
                    "TitleOffsetY": self.TitleOffsetY,
                    "SnapToGrid": self.SnapToGrid,
                    "AnchorCorner": self.AnchorCorner,
                    "RightMargin": self.RightMargin,
                    "TopMargin": self.TopMargin,
                    "BottomMargin": self.BottomMargin,
                    "ViewPadding": self.ViewPadding,
                    "RunViewCleanup": self.RunViewCleanup,
                    "SyncDetailNumbers": self.SyncDetailNumbers,
                    "WinTop": self.window.Top,
                    "WinLeft": self.window.Left,
                    "WinWidth": self.window.Width,
                    "WinHeight": self.window.Height
                }, f)
        except: pass

    @property
    def GridSize(self): return self._grid_size
    @GridSize.setter
    def GridSize(self, value):
        self._grid_size = value
        self.OnPropertyChanged("GridSize")

    @property
    def TitleOffsetX(self): return self._title_offset_x
    @TitleOffsetX.setter
    def TitleOffsetX(self, value):
        self._title_offset_x = value
        self.OnPropertyChanged("TitleOffsetX")

    @property
    def TitleOffsetY(self): return self._title_offset_y
    @TitleOffsetY.setter
    def TitleOffsetY(self, value):
        self._title_offset_y = value
        self.OnPropertyChanged("TitleOffsetY")
        
    @property
    def SnapToGrid(self): return self._snap_to_grid
    @SnapToGrid.setter
    def SnapToGrid(self, value):
        self._snap_to_grid = value
        self.OnPropertyChanged("SnapToGrid")
        
    @property
    def AnchorOptions(self): return self._anchor_options

    @property
    def AnchorCorner(self): return self._anchor_corner
    @AnchorCorner.setter
    def AnchorCorner(self, value):
        self._anchor_corner = value
        self.OnPropertyChanged("AnchorCorner")

    @property
    def RightMargin(self): return self._right_margin
    @RightMargin.setter
    def RightMargin(self, value):
        self._right_margin = value
        self.OnPropertyChanged("RightMargin")
        
    @property
    def TopMargin(self): return self._top_margin
    @TopMargin.setter
    def TopMargin(self, value):
        self._top_margin = value
        self.OnPropertyChanged("TopMargin")
        
    @property
    def BottomMargin(self): return self._bottom_margin
    @BottomMargin.setter
    def BottomMargin(self, value):
        self._bottom_margin = value
        self.OnPropertyChanged("BottomMargin")
        
    @property
    def ViewPadding(self): return self._view_padding
    @ViewPadding.setter
    def ViewPadding(self, value):
        self._view_padding = value
        self.OnPropertyChanged("ViewPadding")
        
    @property
    def RunViewCleanup(self): return self._run_view_cleanup
    @RunViewCleanup.setter
    def RunViewCleanup(self, value):
        self._run_view_cleanup = value
        self.OnPropertyChanged("RunViewCleanup")
        
    @property
    def SyncDetailNumbers(self): return self._sync_detail_numbers
    @SyncDetailNumbers.setter
    def SyncDetailNumbers(self, value):
        self._sync_detail_numbers = value
        self.OnPropertyChanged("SyncDetailNumbers")
        
    @property
    def UnitString(self):
        sym = UnitHelper.get_unit_symbol()
        if sym == "in": return "inches"
        if sym == "ft": return "feet"
        return sym

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

    @property
    def StatusText(self): return self._status_text
    @StatusText.setter
    def StatusText(self, value):
        self._status_text = value
        self.OnPropertyChanged("StatusText")
        
    def can_select_all(self, param=None): return any(not v.IsChecked for v in self.get_all_valid_views_in_nodes(self.Sheets)) if self.Sheets else False
    def can_select_none(self, param=None): return any(v.IsChecked for v in self.get_all_valid_views_in_nodes(self.Sheets)) if self.Sheets else False
    def can_act(self, param=None): return any(v.IsChecked for v in self.get_all_valid_views_in_nodes(self.Sheets)) if self.Sheets else False

    def select_all(self, param=None):
        if not self.Sheets: return
        for v in self.get_all_valid_views_in_nodes(self.Sheets): v.IsChecked = True
        self.CurrentViews = self.get_all_views_in_nodes(self.Sheets)
        self.refresh_commands()
        
    def select_none(self, param=None):
        if not self.Sheets: return
        for v in self.get_all_valid_views_in_nodes(self.Sheets): v.IsChecked = False
        self.CurrentViews = self.get_all_views_in_nodes(self.Sheets)
        self.refresh_commands()

    def expand_all(self, param=None): self._set_expansion(True)
    def collapse_all(self, param=None): self._set_expansion(False)

    def _set_expansion(self, state):
        if not self.Sheets: return
        def recurse(node):
            node.IsExpanded = state
            for c in node.Children: recurse(c)
        for root in self.Sheets: recurse(root)

    def get_valid_viewports(self, sheet):
        valid_types = [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.EngineeringPlan, ViewType.AreaPlan, ViewType.Section, ViewType.Elevation, ViewType.DraftingView, ViewType.Detail, ViewType.Legend]
        return [(doc.GetElement(vp_id), doc.GetElement(doc.GetElement(vp_id).ViewId)) for vp_id in sheet.GetAllViewports() if doc.GetElement(doc.GetElement(vp_id).ViewId).ViewType in valid_types]

    def get_view_title_types(self):
        title_types = []
        vp_types = FilteredElementCollector(doc).OfClass(ElementType).OfCategory(BuiltInCategory.OST_Viewports).ToElements()
        for vpt in vp_types:
            name_param = vpt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if name_param and "Oxyzen Standard" in name_param.AsString():
                title_fmt_param = vpt.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_TITLE_FORMAT)
                if title_fmt_param and title_fmt_param.AsElementId() != ElementId.InvalidElementId:
                    title_sym = doc.GetElement(title_fmt_param.AsElementId())
                    if title_sym:
                        len_param = title_sym.LookupParameter("Title Length")
                        if len_param and len_param.HasValue:
                            title_types.append({'type_id': vpt.Id, 'length': len_param.AsDouble()})
        return sorted(title_types, key=lambda x: x['length'])

    def get_best_title_type_id(self, view_width, title_types):
        if not title_types:
            return None
        best = title_types[-1]
        for ot in title_types:
            if ot['length'] >= view_width:
                best = ot
                break
        return best['type_id']

    def _get_title_height_for_type(self, vp_type_id):
        """Get title height from viewport type. Uses family bounding box, parameters, or safe default."""
        try:
            vpt = doc.GetElement(vp_type_id)
            if not vpt:
                return UnitUtils.ConvertToInternalUnits(0.375, UnitHelper.get_sheet_length_unit())
            
            # Get the title format from viewport type parameter
            title_fmt_param = vpt.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_TITLE_FORMAT)
            if not title_fmt_param or title_fmt_param.AsElementId() == ElementId.InvalidElementId:
                return UnitUtils.ConvertToInternalUnits(0.375, UnitHelper.get_sheet_length_unit())
            
            title_fam = doc.GetElement(title_fmt_param.AsElementId())
            if not title_fam:
                return UnitUtils.ConvertToInternalUnits(0.375, UnitHelper.get_sheet_length_unit())
            
            # Try to get the family definition
            fam_def = getattr(title_fam, 'Family', None)
            if fam_def:
                try:
                    # Try bounding box of family definition
                    bb = fam_def.get_BoundingBox(None)
                    if bb:
                        h = bb.Max.Y - bb.Min.Y
                        if h > 0.01:
                            return h
                except:
                    pass
            
            # Try bounding box of the family symbol
            try:
                bb = title_fam.get_BoundingBox(None)
                if bb:
                    h = bb.Max.Y - bb.Min.Y
                    if h > 0.01:
                        return h
            except:
                pass
            
            # Try common parameter names
            for pname in ("Title Height", "Height", "TITLE HEIGHT", "Box Height", "Size", "Width", "Depth"):
                try:
                    p = title_fam.LookupParameter(pname)
                    if p and p.HasValue:
                        val = p.AsDouble()
                        if val > 0.01:
                            return val
                except:
                    pass
            
        except Exception as ex:
            pass
        
        # Safe default: 3/8 inch (smaller to avoid excessive spacing)
        return UnitUtils.ConvertToInternalUnits(0.375, UnitHelper.get_sheet_length_unit())

    def build_sheet_node(self, sheet):
        vps = self.get_valid_viewports(sheet)
        s_node = SheetNode(sheet, self)
        for vp, view in vps:
            v_node = ViewNode(vp, view, self)
            v_node.ParentNode = s_node
            s_node.Views.append(v_node)
        valid_count = len([v for v in s_node.Views if getattr(v, "HasTitleNumber", False)])
        if valid_count: s_node.Name += " ({})".format(valid_count)
        return s_node

    def scan_active(self, param=None):
        active_view = doc.ActiveView
        if isinstance(active_view, ViewSheet):
            self.window.Cursor = Cursors.Wait
            try:
                root = SheetSetNode("Current View", self)
                s_node = self.build_sheet_node(active_view)
                if s_node.Views:
                    s_node.ParentNode = root
                    root.Children.append(s_node)
                self.Sheets = [root]
                self.CurrentViews = self.get_all_views_in_nodes(self.Sheets)
                valid_count = len(self.get_all_valid_views_in_nodes(self.Sheets))
                ignored_count = len(self.CurrentViews) - valid_count
                self.StatusText = "Scanned active sheet. {} title views, ignored {} without title number.".format(valid_count, ignored_count)
            finally:
                self.window.Cursor = Cursors.Arrow
                self.refresh_commands()
        else:
            forms.alert("Active view is not a Sheet. Scanning full project instead.")
            self.scan_project()

    def scan_project(self, param=None):
        self.window.Cursor = Cursors.Wait
        self.StatusText = "Scanning project..."
        try:
            tree_nodes = []
            all_sheets = sorted(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements(), key=lambda s: s.SheetNumber)
            all_node = SheetSetNode("< All Sheets >", self)
            for sheet in all_sheets:
                s_node = self.build_sheet_node(sheet)
                if s_node.Views:
                    s_node.ParentNode = all_node
                    all_node.Children.append(s_node)
            if all_node.Children: tree_nodes.append(all_node)
                
            sheet_sets = sorted(FilteredElementCollector(doc).OfClass(ViewSheetSet).ToElements(), key=lambda s: s.Name)
            for sset in sheet_sets:
                sset_node = SheetSetNode(sset.Name, self)
                for sheet in sset.Views:
                    if isinstance(sheet, ViewSheet):
                        s_node = self.build_sheet_node(sheet)
                        if s_node.Views:
                            s_node.ParentNode = sset_node
                            sset_node.Children.append(s_node)
                if sset_node.Children: tree_nodes.append(sset_node)
                    
            self.Sheets = tree_nodes
            self.CurrentViews = self.get_all_views_in_nodes(self.Sheets)
            valid_count = len(self.get_all_valid_views_in_nodes(self.Sheets))
            ignored_count = len(self.CurrentViews) - valid_count
            self.StatusText = "Scanned {} sheet sets: {} title views, ignored {} without title number.".format(len(tree_nodes)-1, valid_count, ignored_count)
        finally:
            self.window.Cursor = Cursors.Arrow
            self.refresh_commands()

    def get_all_views_in_nodes(self, nodes):
        views = []
        for node in nodes:
            if hasattr(node, "Views") and node.Views: views.extend(node.Views)
            elif hasattr(node, "Children") and node.Children: views.extend(self.get_all_views_in_nodes(node.Children))
        return views

    def get_all_valid_views_in_nodes(self, nodes):
        views = []
        for node in nodes:
            if hasattr(node, "Views") and node.Views:
                views.extend([v for v in node.Views if getattr(v, "HasTitleNumber", False)])
            elif hasattr(node, "Children") and node.Children:
                views.extend(self.get_all_valid_views_in_nodes(node.Children))
        return views

    def cancel_action(self, parameter): self.window.Close()
        
    def get_checked_views(self):
        return [v_node for root in self.Sheets for s_node in root.Children for v_node in s_node.Views if v_node.IsChecked and getattr(v_node, "HasTitleNumber", False)]

    def arrange_titles(self, parameter):
        checked_views = self.get_checked_views()
        self.window.Cursor = Cursors.Wait
        self.StatusText = "Arranging View Titles..."
        try:
            vp_types = FilteredElementCollector(doc).OfClass(ElementType).OfCategory(BuiltInCategory.OST_Viewports).ToElements()
            oxyzen_types = []
            for vpt in vp_types:
                name_param = vpt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if name_param and "Oxyzen Standard" in name_param.AsString():
                    title_fmt_param = vpt.get_Parameter(BuiltInParameter.VIEWPORT_ATTR_TITLE_FORMAT)
                    if title_fmt_param and title_fmt_param.AsElementId() != ElementId.InvalidElementId:
                        title_sym = doc.GetElement(title_fmt_param.AsElementId())
                        len_param = title_sym.LookupParameter("Title Length")
                        if len_param: oxyzen_types.append({'type_id': vpt.Id, 'length': len_param.AsDouble()})
            oxyzen_types = sorted(oxyzen_types, key=lambda x: x['length'])
            
            with Transaction(doc, "Arrange View Titles") as t:
                t.Start()
                count = 0
                for v_node in checked_views:
                    vp = v_node.Viewport
                    if vp: vp.LabelOffset = XYZ.Zero
                doc.Regenerate()

                # STEP 1: Calculate the label offset from the FIRST view only
                label_offset_x = 0.0
                label_offset_y = 0.0
                first_view = True
                
                for v_node in checked_views:
                    vp = v_node.Viewport
                    if not vp: continue
                    
                    if first_view:
                        # Only on first view: calculate offset based on anchor settings
                        first_view = False
                        
                        sheet = doc.GetElement(vp.SheetId)
                        tbs = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
                        tb_min_x, tb_min_y, tb_max_x, tb_max_y = 0.0, 0.0, 3.0, 2.0
                        if tbs:
                            tb_box = tbs[0].get_BoundingBox(sheet)
                            if tb_box: tb_min_x, tb_min_y, tb_max_x, tb_max_y = tb_box.Min.X, tb_box.Min.Y, tb_box.Max.X, tb_box.Max.Y
                        
                        try: tb_offset_x = UnitHelper.parse_unit_to_internal(self.TitleOffsetX)
                        except: tb_offset_x = 1.5 / 12.0
                        try: tb_offset_y = UnitHelper.parse_unit_to_internal(self.TitleOffsetY)
                        except: tb_offset_y = 0.5 / 12.0
                        
                        box = vp.GetBoxOutline()
                        min_x, min_y, max_x, max_y = box.MinimumPoint.X, box.MinimumPoint.Y, box.MaximumPoint.X, box.MaximumPoint.Y
                        
                        anchor = self.AnchorCorner
                        try:
                            if 'Left' in anchor:
                                x_anchor = tb_min_x + tb_offset_x
                                x_ref = min_x
                            else:
                                x_anchor = tb_max_x - tb_offset_x
                                x_ref = max_x

                            if 'Bottom' in anchor:
                                y_anchor = tb_min_y + tb_offset_y
                                y_ref = min_y
                            else:
                                y_anchor = tb_max_y - tb_offset_y
                                y_ref = max_y

                            label_offset_x = x_anchor - x_ref
                            label_offset_y = y_anchor - y_ref
                            print("[Arrange] First view offset: X={:.4f}, Y={:.4f}".format(label_offset_x, label_offset_y))
                        except Exception as e:
                            print("[Arrange] Error calculating offset from first view: {}".format(e))
                            label_offset_x = 0.0
                            label_offset_y = 0.0

                # STEP 2: Apply the same offset to ALL views
                for v_node in checked_views:
                    vp = v_node.Viewport
                    if not vp: continue
                    
                    sheet = doc.GetElement(vp.SheetId)
                    tbs = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
                    tb_min_x, tb_min_y, tb_max_x, tb_max_y = 0.0, 0.0, 3.0, 2.0
                    if tbs:
                        tb_box = tbs[0].get_BoundingBox(sheet)
                        if tb_box: tb_min_x, tb_min_y, tb_max_x, tb_max_y = tb_box.Min.X, tb_box.Min.Y, tb_box.Max.X, tb_box.Max.Y
                        
                    safe_width, safe_height = (tb_max_x - tb_min_x) - (12.5 / 12.0), (tb_max_y - tb_min_y) - (1.0 / 12.0)
                    box = vp.GetBoxOutline()
                    min_x, min_y, max_x, max_y = box.MinimumPoint.X, box.MinimumPoint.Y, box.MaximumPoint.X, box.MaximumPoint.Y
                    
                    try:
                        snap_grid = UnitHelper.parse_unit_to_internal(self.GridSize)
                    except:
                        snap_grid = 1.0 / 12.0

                    if self.SnapToGrid and not vp.Pinned and snap_grid > 0.01:
                        rel_x, rel_y = min_x - tb_min_x, min_y - tb_min_y
                        target_rel_x, target_rel_y = round(rel_x / snap_grid) * snap_grid, round(rel_y / snap_grid) * snap_grid
                        dx, dy = target_rel_x - rel_x, target_rel_y - rel_y
                        if abs(dx) > 0.0001 or abs(dy) > 0.0001:
                            center = vp.GetBoxCenter()
                            vp.SetBoxCenter(XYZ(center.X + dx, center.Y + dy, center.Z))
                            min_x += dx; min_y += dy; max_x += dx; max_y += dy

                    has_collision = False
                    schedules = FilteredElementCollector(doc, sheet.Id).OfClass(ScheduleSheetInstance).ToElements()
                    for sch in schedules:
                        sch_box = sch.get_BoundingBox(sheet)
                        if sch_box and not (max_x < sch_box.Min.X or min_x > sch_box.Max.X or max_y < sch_box.Min.Y or min_y > sch_box.Max.Y):
                            has_collision = True; break
                    if not has_collision:
                        for other_vpid in sheet.GetAllViewports():
                            if other_vpid != vp.Id:
                                other_box = doc.GetElement(other_vpid).GetBoxOutline()
                                if not (max_x < other_box.MinimumPoint.X or min_x > other_box.MaximumPoint.X or max_y < other_box.MinimumPoint.Y or min_y > other_box.MaximumPoint.Y):
                                    has_collision = True; break

                    if has_collision:
                        status = "Collision"
                    elif (max_x - min_x) > safe_width or (max_y - min_y) > safe_height:
                        status = "Oversized"
                        try:
                            max_title_len = oxyzen_types[-1]['length'] if oxyzen_types else None
                        except Exception:
                            max_title_len = None
                        try:
                            print("[ManageViews] Oversized viewport {} on sheet {}: box {:.3f}x{:.3f}, safe {:.3f}x{:.3f}, max_title_len={}".format(get_id(vp.Id) if vp else 'N/A', get_id(sheet.Id) if sheet else 'N/A', (max_x - min_x), (max_y - min_y), safe_width, safe_height, max_title_len))
                        except Exception:
                            pass
                    else:
                        status = "Arranged"
                    v_node.TitleStatus = status
                    
                    if oxyzen_types:
                        best_type_id = oxyzen_types[-1]['type_id']
                        for ot in oxyzen_types:
                            if ot['length'] >= (max_x - min_x):
                                best_type_id = ot['type_id']
                                break
                        if vp.GetTypeId() != best_type_id: vp.ChangeTypeId(best_type_id)
                    
                    # Apply ZERO offset: let titles follow their viewports naturally
                    vp.LabelOffset = XYZ.Zero
                    count += 1
                    
                t.Commit()
            self.StatusText = "Arranged titles for {} views.".format(count)
        except Exception as e:
            self.StatusText = "Error occurred."
            forms.alert("An error occurred:\n{}".format(e))
        finally: self.window.Cursor = Cursors.Arrow

    def auto_pack(self, parameter):
        checked_views = self.get_checked_views()
        self.window.Cursor = Cursors.Wait
        self.StatusText = "Auto-Packing Sheets..."
        try:
            views_by_sheet = {}
            for v_node in checked_views:
                # Implicitly ignore Legends during Auto-Pack layout
                if v_node.View.ViewType == ViewType.Legend: continue
                sheet_id = v_node.Viewport.SheetId
                views_by_sheet.setdefault(sheet_id, []).append(v_node)
                
            count = 0
            
            is_m = UnitHelper.is_metric()
            try: r_margin = UnitHelper.parse_unit_to_internal(self.RightMargin)
            except: r_margin = UnitUtils.ConvertToInternalUnits(250.0 if is_m else 10.0, UnitHelper.get_sheet_length_unit()) 
            try: t_margin = UnitHelper.parse_unit_to_internal(self.TopMargin)
            except: t_margin = 0.0
            try: b_margin = UnitHelper.parse_unit_to_internal(self.BottomMargin)
            except: b_margin = 0.0
            try: padding = UnitHelper.parse_unit_to_internal(self.ViewPadding)
            except: padding = UnitUtils.ConvertToInternalUnits(25.0 if is_m else 1.0, UnitHelper.get_sheet_length_unit()) 
            
            try: tb_offset_x = UnitHelper.parse_unit_to_internal(self.TitleOffsetX)
            except: tb_offset_x = UnitUtils.ConvertToInternalUnits(38.0 if is_m else 1.5, UnitHelper.get_sheet_length_unit()) 
            try: tb_offset_y = UnitHelper.parse_unit_to_internal(self.TitleOffsetY)
            except: tb_offset_y = UnitUtils.ConvertToInternalUnits(12.0 if is_m else 0.5, UnitHelper.get_sheet_length_unit()) 
            
            anchor = self.AnchorCorner

            # --- Pre-flight Simulation: ensure packing fits within title block bounds ---
            title_types = self.get_view_title_types()
            sheet_unit = UnitHelper.get_sheet_length_unit()
            min_spacing = UnitUtils.ConvertToInternalUnits(0.5, sheet_unit)
            title_height_est = UnitUtils.ConvertToInternalUnits(0.5, sheet_unit)
            row_gap = min_spacing + title_height_est + min_spacing
            horiz_gap = min_spacing

            # simulate per sheet to catch oversize or overflow before starting transactions
            for sheet_id, v_nodes in views_by_sheet.items():
                sheet = doc.GetElement(sheet_id)
                tbs = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
                tb_min_x, tb_min_y, tb_max_x, tb_max_y = 0.0, 0.0, 3.0, 2.0
                if tbs:
                    tb_box = tbs[0].get_BoundingBox(sheet)
                    if tb_box: tb_min_x, tb_min_y, tb_max_x, tb_max_y = tb_box.Min.X, tb_box.Min.Y, tb_box.Max.X, tb_box.Max.Y

                usable_min_x = tb_min_x + tb_offset_x
                usable_max_x = tb_max_x - r_margin - tb_offset_x
                usable_min_y = tb_min_y + b_margin + tb_offset_y
                usable_max_y = tb_max_y - t_margin - tb_offset_y

                usable_width = usable_max_x - usable_min_x
                usable_height = usable_max_y - usable_min_y

                # build vp data (preserve original order)
                sim_vp = []
                for v_node in v_nodes:
                    vp = v_node.Viewport
                    box = vp.GetBoxOutline()
                    w = box.MaximumPoint.X - box.MinimumPoint.X
                    h = box.MaximumPoint.Y - box.MinimumPoint.Y
                    title_h = self._get_title_height_for_type(vp.GetTypeId())
                    sim_vp.append({'w': w, 'h': h, 'title_h': title_h, 'id': get_id(vp.Id)})

                # quick oversize check
                for item in sim_vp:
                    if item['w'] > usable_width - 1e-6:
                        msg = "Auto-Pack warning: viewport {} on sheet {} width ({:.3f}) exceeds usable width ({:.3f}).".format(item['id'], get_id(sheet.Id), item['w'], usable_width)
                        print(msg)

                # Simulate row packing into bins using usable_width and horiz_gap
                rows = []
                rows_title = []
                curr_row_width = 0.0
                curr_row_max_h = 0.0
                curr_row_title_max = 0.0
                first_in_row = True
                for item in sim_vp:
                    w = item['w']
                    h = item['h']
                    t_h = item.get('title_h', title_height_est)
                    if first_in_row:
                        curr_row_width = w
                        curr_row_max_h = h
                        curr_row_title_max = t_h
                        first_in_row = False
                    else:
                        # can we fit in current row?
                        if curr_row_width + horiz_gap + w <= usable_width + 1e-6:
                            curr_row_width += horiz_gap + w
                            if h > curr_row_max_h: curr_row_max_h = h
                            if t_h > curr_row_title_max: curr_row_title_max = t_h
                        else:
                            rows.append(curr_row_max_h)
                            rows_title.append(curr_row_title_max)
                            curr_row_width = w
                            curr_row_max_h = h
                            curr_row_title_max = t_h
                if not first_in_row:
                    rows.append(curr_row_max_h)
                    rows_title.append(curr_row_title_max)

                # compute required vertical space: sum of view row heights + per-row title gaps (min_spacing + title_h + min_spacing)
                num_rows = len(rows)
                required_vertical = 0.0
                for i in range(num_rows):
                    required_vertical += rows[i]
                    # add gap below row except after last row? The title baseline adds spacing between rows; include for all but last
                    if i < num_rows - 1:
                        required_vertical += (min_spacing + rows_title[i] + min_spacing)

                if required_vertical > usable_height + 1e-6:
                    msg = "Auto-Pack warning: sheet {} required vertical space ({:.3f}) exceeds usable height ({:.3f}).".format(get_id(sheet.Id), required_vertical, usable_height)
                    print(msg)

            # start group now that pre-flight passed
            with TransactionGroup(doc, "Auto-Pack Sheet") as tg:
                tg.Start()
                
                # --- PHASE 1: Pre-Process (Scramble numbers & Clean View Extents) ---
                if self.RunViewCleanup or self.SyncDetailNumbers:
                    with Transaction(doc, "Pre-Process Views") as t1:
                        t1.Start()
                        for sheet_id, v_nodes in views_by_sheet.items():
                            for i, v_node in enumerate(v_nodes):
                                vp = v_node.Viewport
                                view = v_node.View
                                
                                if self.SyncDetailNumbers:
                                    p = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                    if p and not p.IsReadOnly: p.Set("TMP_{}_{}".format(get_id(vp.Id), i))
                                
                                if self.RunViewCleanup:
                                    cats_to_hide = [BuiltInCategory.OST_VolumeOfInterest, BuiltInCategory.OST_CLines]
                                    for cat in cats_to_hide:
                                        cat_id = ElementId(cat)
                                        if view.CanCategoryBeHidden(cat_id) and not view.GetCategoryHidden(cat_id):
                                            view.SetCategoryHidden(cat_id, True)
                                            
                                    viewers = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Viewers).WhereElementIsNotElementType().ToElements()
                                    hide_ids = List[ElementId]()
                                    for v in viewers:
                                        vp_param = v.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER)
                                        if not vp_param or not vp_param.HasValue or vp_param.AsString() == "---":
                                            hide_ids.Add(v.Id)
                                    if hide_ids.Count > 0: view.HideElements(hide_ids)
                        doc.Regenerate() # Recalculate physical bounding boxes after hiding junk
                        t1.Commit()
                
                # --- PHASE 2: Sizing, Sorting, and Shelf Packing ---
                with Transaction(doc, "Pack & Arrange Views") as t2:
                    t2.Start()
                    title_types = self.get_view_title_types()
                    sheet_unit = UnitHelper.get_sheet_length_unit()
                    min_spacing = UnitUtils.ConvertToInternalUnits(0.5, sheet_unit)
                    title_height_est = UnitUtils.ConvertToInternalUnits(0.5, sheet_unit)
                    row_gap = min_spacing + title_height_est + min_spacing
                    horiz_gap = min_spacing
                    
                    for sheet_id, v_nodes in views_by_sheet.items():
                        sheet = doc.GetElement(sheet_id)
                        tbs = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
                        
                        tb_min_x, tb_min_y, tb_max_x, tb_max_y = 0.0, 0.0, 3.0, 2.0
                        if tbs:
                            tb_box = tbs[0].get_BoundingBox(sheet)
                            if tb_box: tb_min_x, tb_min_y, tb_max_x, tb_max_y = tb_box.Min.X, tb_box.Min.Y, tb_box.Max.X, tb_box.Max.Y
                                
                        usable_min_x = tb_min_x + tb_offset_x
                        usable_max_x = tb_max_x - r_margin - tb_offset_x
                        usable_min_y = tb_min_y + b_margin + tb_offset_y
                        usable_max_y = tb_max_y - t_margin - tb_offset_y
                        
                        vp_data = []
                        for v_node in v_nodes:
                            vp = v_node.Viewport
                            box = vp.GetBoxOutline()
                            w = box.MaximumPoint.X - box.MinimumPoint.X
                            h = box.MaximumPoint.Y - box.MinimumPoint.Y
                            title_h = self._get_title_height_for_type(vp.GetTypeId())
                            vp_data.append({'v_node': v_node, 'vp': vp, 'w': w, 'h': h, 'title_h': title_h, 'area': w * h, 'id': get_id(vp.Id)})
                            
                        vp_data = sorted(vp_data, key=lambda x: (-x['area'], x['id']))
                        
                        dx = 1 if 'Left' in anchor else -1
                        dy = 1 if 'Bottom' in anchor else -1
                        
                        curr_x = usable_min_x if dx == 1 else usable_max_x
                        curr_y = usable_min_y if dy == 1 else usable_max_y
                        row_max_h = 0
                        # Track cumulative row bottom Y position (not row_index * gap)
                        row_bottom = usable_min_y if dy == 1 else usable_max_y
                        
                        taken_numbers = set()
                        if self.SyncDetailNumbers:
                            all_vp_ids = sheet.GetAllViewports()
                            checked_vp_ids = set(x['id'] for x in vp_data)
                            for vp_id in all_vp_ids:
                                if get_id(vp_id) not in checked_vp_ids:
                                    p = doc.GetElement(vp_id).get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                    if p and p.HasValue: taken_numbers.add(p.AsString())
                        
                        detail_num = 1
                        first_in_row = True
                        prev_width = 0.0
                        prev_center_x = None
                        prev_row_max_h = 0.0
                        prev_row_title_h = 0.0
                        
                        for data in vp_data:
                            w, h, vp, v_node = data['w'], data['h'], data['vp'], data['v_node']
                            # ensure we have title height for this item (may change after ChangeTypeId)
                            title_h = data.get('title_h', title_height_est)
                            
                            # Pick the closest title format for this view width
                            best_type_id = self.get_best_title_type_id(w, title_types)
                            if best_type_id and vp.GetTypeId() != best_type_id:
                                vp.ChangeTypeId(best_type_id)
                                box = vp.GetBoxOutline()
                                w = box.MaximumPoint.X - box.MinimumPoint.X
                                h = box.MaximumPoint.Y - box.MinimumPoint.Y
                                # recompute title height after type change
                                title_h = self._get_title_height_for_type(vp.GetTypeId())
                                data['title_h'] = title_h
                                v_node.TitleStatus = "Matched"

                            # Decide placement relative to previous view extents; row_bottom tracks cumulative Y
                            if first_in_row:
                                # First item in first row
                                if dx == 1:
                                    curr_x = usable_min_x + (w / 2.0)
                                else:
                                    curr_x = usable_max_x - (w / 2.0)
                                if dy == 1:
                                    curr_y = row_bottom + (h / 2.0)
                                else:
                                    curr_y = row_bottom - (h / 2.0)
                                first_in_row = False
                                prev_center_x = curr_x
                            else:
                                # compute next center x from previous view right/left extent + horiz_gap
                                if dx == 1:
                                    prev_right = prev_center_x + (prev_width / 2.0)
                                    next_center_x = prev_right + horiz_gap + (w / 2.0)
                                else:
                                    prev_left = prev_center_x - (prev_width / 2.0)
                                    next_center_x = prev_left - horiz_gap - (w / 2.0)

                                # Check wrap to next row using usable bounds
                                if (dx == 1 and next_center_x + (w / 2.0) > usable_max_x) or (dx == -1 and next_center_x - (w / 2.0) < usable_min_x):
                                    # Move to next row: accumulate from previous row bottom + height + title gap (0.5" + title_h + 0.5")
                                    row_gap_computed = min_spacing + prev_row_title_h + min_spacing
                                    if dy == 1:
                                        # Next row bottom = current row bottom + current row max height + row gap
                                        row_bottom = row_bottom + prev_row_max_h + row_gap_computed
                                        curr_y = row_bottom + (h / 2.0)
                                    else:
                                        row_bottom = row_bottom - prev_row_max_h - row_gap_computed
                                        curr_y = row_bottom - (h / 2.0)

                                    # reset X to row start
                                    if dx == 1:
                                        curr_x = usable_min_x + (w / 2.0)
                                    else:
                                        curr_x = usable_max_x - (w / 2.0)
                                    # reset row tracking for new row (heights reset, title will be set on next placement)
                                    row_max_h = 0
                                    prev_row_max_h = 0.0
                                    prev_row_title_h = 0.0
                                    prev_center_x = curr_x
                                else:
                                    curr_x = next_center_x
                                    prev_center_x = curr_x

                            # Place Viewport Center
                            c_x = curr_x
                            c_y = curr_y
                            vp.SetBoxCenter(XYZ(c_x, c_y, 0))
                            prev_width = w
                            prev_center_x = c_x
                            # Track this row's max viewport height and title height for next row computation
                            prev_row_max_h = max(prev_row_max_h, h)
                            prev_row_title_h = max(prev_row_title_h, title_h)

                            # Apply Auto-Numbering
                            if self.SyncDetailNumbers:
                                p = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                if p and not p.IsReadOnly:
                                    while str(detail_num) in taken_numbers: detail_num += 1
                                    p.Set(str(detail_num))
                                    taken_numbers.add(str(detail_num))
                                    v_node.DetailNumber = str(detail_num)
                                    detail_num += 1
                                    
                                    # Lock View Title to Standard Origin
                                    new_min_x, new_min_y = c_x - (w / 2.0), c_y - (h / 2.0)
                                    new_max_x, new_max_y = c_x + (w / 2.0), c_y + (h / 2.0)
                                    anchor = self.AnchorCorner
                                    try:
                                        if 'Left' in anchor:
                                            x_anchor = tb_min_x + tb_offset_x
                                            x_ref = new_min_x
                                        else:
                                            x_anchor = tb_max_x - tb_offset_x
                                            x_ref = new_max_x

                                        if 'Bottom' in anchor:
                                            y_anchor = tb_min_y + tb_offset_y
                                            y_ref = new_min_y
                                        else:
                                            y_anchor = tb_max_y - tb_offset_y
                                            y_ref = new_max_y

                                        vp.LabelOffset = XYZ(x_anchor - x_ref, y_anchor - y_ref, 0)
                                    except Exception:
                                        vp.LabelOffset = XYZ((tb_min_x + tb_offset_x) - new_min_x, (tb_min_y + tb_offset_y) - new_min_y, 0)
                            
                            # Advance Matrix
                            count += 1
                            # advance X by using prev_center_x for next iteration (columns handled above)
                            row_max_h = max(row_max_h, h)
                            v_node.TitleStatus = "Packed"
                            
                    t2.Commit()
                tg.Assimilate()
            self.StatusText = "Auto-packed {} views.".format(count)
        except Exception as e:
            self.StatusText = "Error occurred."
            forms.alert("An error occurred:\n{}".format(traceback.format_exc()))
        finally: self.window.Cursor = Cursors.Arrow

# --- Main UI Class ---
class ManageViewsWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, os.path.join(os.path.dirname(__file__), 'ui.xaml'))
        
        self.HeaderDrag.MouseLeftButtonDown += lambda s, a: self.DragMove()
        self.Btn_WinClose.Click += lambda s, a: self.Close()
        self.Closing += lambda s, a: self.viewModel.save_config()
        self.systemTree.SelectedItemChanged += self.tree_selection_changed
        self.systemTree.MouseDoubleClick += self.tree_double_click
        self.sysDataGrid.MouseDoubleClick += self.grid_double_click
        self.sysDataGrid.SelectionChanged += self.grid_selection_changed
        
        for tb in [self.Tb_GridSize, self.Tb_TitleOffsetX, self.Tb_TitleOffsetY,
                   self.Tb_RightMargin, self.Tb_TopMargin, self.Tb_BottomMargin, self.Tb_ViewPadding]:
            tb.LostFocus += self.format_textbox
            tb.KeyDown += self.format_textbox
        
        self.apply_revit_theme()
        self.viewModel = ManageViewsViewModel(self)
        self.DataContext = self.viewModel
        self.Title += " [{}]".format(UnitHelper.get_unit_symbol())

    def tree_selection_changed(self, sender, args):
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node:
                self.viewModel.CurrentViews = self.viewModel.get_all_views_in_nodes([selected_node])
        except: pass
            
    def grid_selection_changed(self, sender, args):
        try:
            selected_items = list(self.sysDataGrid.SelectedItems)
            if not selected_items: return
            elem_ids = List[ElementId]()
            for item in selected_items:
                if hasattr(item, "Viewport") and item.Viewport: elem_ids.Add(item.Viewport.Id)
            if elem_ids and self.Cb_AutoZoom.IsChecked:
                revit.uidoc.ShowElements(elem_ids)
                revit.uidoc.Selection.SetElementIds(elem_ids)
        except: pass

    def format_textbox(self, sender, args):
        if hasattr(args, "Key") and args.Key != Key.Enter: return
        try:
            if hasattr(sender, "Text"):
                val_internal = UnitHelper.parse_unit_to_internal(sender.Text)
                formatted_text = UnitHelper.to_formatted_string_with_symbol(val_internal)
                sender.Text = formatted_text
                
                if sender.Name == "Tb_GridSize": self.viewModel.GridSize = formatted_text
                elif sender.Name == "Tb_TitleOffsetX": self.viewModel.TitleOffsetX = formatted_text
                elif sender.Name == "Tb_TitleOffsetY": self.viewModel.TitleOffsetY = formatted_text
                elif sender.Name == "Tb_RightMargin": self.viewModel.RightMargin = formatted_text
                elif sender.Name == "Tb_TopMargin": self.viewModel.TopMargin = formatted_text
                elif sender.Name == "Tb_BottomMargin": self.viewModel.BottomMargin = formatted_text
                elif sender.Name == "Tb_ViewPadding": self.viewModel.ViewPadding = formatted_text

                if hasattr(sender, "BorderBrush"): sender.BorderBrush = SolidColorBrush(WpfColor.FromRgb(75, 85, 99) if getattr(self, "_is_dark", False) else WpfColor.FromRgb(209, 213, 219))
        except:
            if hasattr(sender, "BorderBrush"): sender.BorderBrush = SolidColorBrush(Colors.Red)

    def tree_double_click(self, sender, args):
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node and hasattr(selected_node, "Sheet"):
                elem_ids = List[ElementId]([selected_node.Sheet.Id])
                revit.uidoc.ShowElements(elem_ids)
                revit.uidoc.Selection.SetElementIds(elem_ids)
        except: pass

    def grid_double_click(self, sender, args):
        try:
            selected_items = list(self.sysDataGrid.SelectedItems)
            if selected_items:
                elem_ids = List[ElementId]()
                for item in selected_items:
                    if hasattr(item, "Viewport") and item.Viewport: elem_ids.Add(item.Viewport.Id)
                if elem_ids:
                    revit.uidoc.ShowElements(elem_ids)
                    revit.uidoc.Selection.SetElementIds(elem_ids)
        except: pass

    def apply_revit_theme(self):
        self._is_dark = False
        try:
            if int(HOST_APP.version) >= 2024:
                from Autodesk.Revit.UI import UIThemeManager, UITheme
                if UIThemeManager.CurrentTheme == UITheme.Dark: self._is_dark = True
        except: pass
        if self._is_dark:
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
            res["SelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(30, 58, 138))
            res["SelectionBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(59, 130, 246))
            res["SelectionTextBrush"] = SolidColorBrush(Colors.White)

if __name__ == '__main__':
    ManageViewsWindow().ShowDialog()