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
    UnitUtils, SpecTypeId, UnitFormatUtils, ElementId, ElementType, ScheduleSheetInstance, BoundingBoxXYZ,
    Line
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
        try:
            from Autodesk.Revit.DB import UnitTypeId
            if UnitHelper.is_metric():
                return UnitTypeId.Millimeters
            return UnitTypeId.Inches
        except:
            from Autodesk.Revit.DB import DisplayUnitType
            if UnitHelper.is_metric():
                return DisplayUnitType.DUT_MILLIMETERS
            return DisplayUnitType.DUT_FRAC_INCHES

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
        is_3d_or_rendering = view.ViewType in [ViewType.ThreeD, ViewType.Rendering, ViewType.Walkthrough]
        if (detail_value and detail_value not in ["-", "---"] and any(ch.isdigit() for ch in detail_value)) or is_3d_or_rendering:
            self._detail_number = detail_value if (detail_value and detail_value.strip()) else "-"
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
        self.DrawCrosshairCommand = RelayCommand(self.draw_crosshair)
        
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
        valid_types = [
            ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.EngineeringPlan, 
            ViewType.AreaPlan, ViewType.Section, ViewType.Elevation, 
            ViewType.DraftingView, ViewType.Detail, ViewType.Legend,
            ViewType.ThreeD, ViewType.Rendering, ViewType.Walkthrough
        ]
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
        
    def draw_crosshair(self, parameter=None):
        self.window.Cursor = Cursors.Wait
        self.StatusText = "Aligning sheet origins to absolute (0,0)..."
        
        # Get all sheets in the project
        all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
        
        table_data = []
        
        try:
            with Transaction(doc, "Align Title Blocks and Elements to Origin") as t:
                t.Start()
                
                adjusted_sheets_count = 0
                for sheet in sorted(all_sheets, key=lambda s: s.SheetNumber):
                    # Find title blocks
                    tbs = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
                    if not tbs:
                        table_data.append([sheet.SheetNumber, sheet.Name, "N/A", "0", "No Title Block found"])
                        continue
                        
                    tb = tbs[0]
                    tb_box = tb.get_BoundingBox(sheet)
                    if not tb_box:
                        table_data.append([sheet.SheetNumber, sheet.Name, "N/A", "0", "Title Block has no boundary"])
                        continue
                        
                    # Calculate translation vector to move bottom-left of TB to (0,0)
                    dx = 0.0 - tb_box.Min.X
                    dy = 0.0 - tb_box.Min.Y
                    
                    # Draw verification crosshair at absolute origin (0,0)
                    half_len = 1.0 / 12.0  # 1 inch
                    line_h = Line.CreateBound(XYZ(-half_len, 0, 0), XYZ(half_len, 0, 0))
                    line_v = Line.CreateBound(XYZ(0, -half_len, 0), XYZ(0, half_len, 0))
                    
                    try:
                        doc.Create.NewDetailCurve(sheet, line_h)
                        doc.Create.NewDetailCurve(sheet, line_v)
                        crosshair_status = "Yes"
                    except:
                        crosshair_status = "Error drawing"
                    
                    if abs(dx) < 0.0001 and abs(dy) < 0.0001:
                        table_data.append([sheet.SheetNumber, sheet.Name, "(0.00, 0.00)", "0", "Already aligned (Crosshair: {})".format(crosshair_status)])
                        continue
                        
                    translation_vector = XYZ(dx, dy, 0)
                    
                    # Collect all elements on sheet
                    element_ids = set()
                    
                    # 1. Viewports
                    for vp_id in sheet.GetAllViewports():
                        element_ids.add(vp_id)
                        
                    # 2. Visible elements owned by sheet
                    collector = FilteredElementCollector(doc, sheet.Id).WhereElementIsNotElementType()
                    for elem in collector:
                        if elem.Id == sheet.Id:
                            continue
                        if elem.Category is None:
                            continue
                        # Skip view references or camera references that cannot be moved
                        cat_id = get_id(elem.Category.Id)
                        if cat_id in [-2000279, -2000278]: 
                            continue
                        element_ids.add(elem.Id)
                        
                    # Unpin elements if needed
                    for eid in element_ids:
                        elem = doc.GetElement(eid)
                        if elem and elem.Pinned:
                            try:
                                elem.Pinned = False
                            except:
                                pass
                                
                    # Move elements
                    moved_count = 0
                    for eid in element_ids:
                        elem = doc.GetElement(eid)
                        if not elem:
                            continue
                        try:
                            if isinstance(elem, Viewport):
                                # Handle locked 3D views (e.g. perspective views)
                                view = doc.GetElement(elem.ViewId)
                                was_locked = False
                                if view and hasattr(view, "IsLocked") and view.IsLocked:
                                    try:
                                        view.IsLocked = False
                                        was_locked = True
                                    except:
                                        pass
                                
                                viewport_moved = False
                                try:
                                    # Try moving via SetBoxCenter first
                                    old_center = elem.GetBoxCenter()
                                    new_center = XYZ(old_center.X + dx, old_center.Y + dy, old_center.Z)
                                    elem.SetBoxCenter(new_center)
                                    viewport_moved = True
                                    moved_count += 1
                                except:
                                    pass
                                    
                                if not viewport_moved:
                                    try:
                                        # Fallback to ElementTransformUtils.MoveElement
                                        ElementTransformUtils.MoveElement(doc, eid, translation_vector)
                                        moved_count += 1
                                    except:
                                        pass
                                        
                                if was_locked:
                                    try:
                                        view.IsLocked = True
                                    except:
                                        pass
                            elif isinstance(elem, ScheduleSheetInstance):
                                # Move schedule sheet instance via Point property
                                old_point = elem.Point
                                new_point = XYZ(old_point.X + dx, old_point.Y + dy, old_point.Z)
                                elem.Point = new_point
                                moved_count += 1
                            else:
                                # Move all other elements (Title Blocks, annotations, text notes, lines, generic annotations, revision clouds, tags, images, etc.)
                                ElementTransformUtils.MoveElement(doc, eid, translation_vector)
                                moved_count += 1
                        except:
                            pass
                                
                    table_data.append([
                        sheet.SheetNumber, 
                        sheet.Name, 
                        "({:.3f}, {:.3f})".format(tb_box.Min.X, tb_box.Min.Y), 
                        str(moved_count), 
                        "Moved to (0,0) (Crosshair: {})".format(crosshair_status)
                    ])
                    adjusted_sheets_count += 1
                    
                t.Commit()
                
            self.StatusText = "Aligned {} sheets to absolute origin.".format(adjusted_sheets_count)
            
            # Print the summary as a beautiful table in pyRevit output window
            output = script.get_output()
            output.print_table(
                table_data=table_data,
                title="Universal Title Block Origin Alignment Summary",
                columns=["Sheet Number", "Sheet Name", "Original Bottom-Left", "Moved Elements", "Alignment Status"]
            )
            output.show()
            
        except Exception as e:
            self.StatusText = "Failed to align sheets."
            forms.alert("Error occurred during alignment:\n{}".format(traceback.format_exc()))
        finally:
            self.window.Cursor = Cursors.Arrow

    def get_viewport_geometry(self, vp, view):
        """Return a unified dict of RED+BLUE geometry data for a single viewport.
        
        This is the single source of truth for all collision detection and layout
        calculations. It captures both the viewport's drawn frame (RED from
        GetBoxOutline) and its rendered label (BLUE from GetLabelOutline) along
        with derived relationships between them.
        """
        try:
            box = vp.GetBoxOutline()
            red_min_x = box.MinimumPoint.X
            red_min_y = box.MinimumPoint.Y
            red_max_x = box.MaximumPoint.X
            red_max_y = box.MaximumPoint.Y
            red_w = red_max_x - red_min_x
            red_h = red_max_y - red_min_y
            
            label_box = vp.GetLabelOutline()
            blue_label_min_y = label_box.MinimumPoint.Y
            blue_label_max_y = label_box.MaximumPoint.Y
            blue_label_h = blue_label_max_y - blue_label_min_y
            blue_label_w = label_box.MaximumPoint.X - label_box.MinimumPoint.X
            
            geo = {
                'red_min_x': red_min_x,
                'red_min_y': red_min_y,
                'red_max_x': red_max_x,
                'red_max_y': red_max_y,
                'red_center': vp.GetBoxCenter(),
                'red_width': red_w,
                'red_height': red_h,
                'blue_label_min_y': blue_label_min_y,
                'blue_label_max_y': blue_label_max_y,
                'blue_label_height': blue_label_h,
                'blue_label_width': blue_label_w,
                'view_to_label_gap': red_max_y - blue_label_min_y,
                'rotation': vp.Rotation,
            }
        except:
            geo = {
                'red_min_x': 0, 'red_min_y': 0, 'red_max_x': 1, 'red_max_y': 1,
                'red_center': XYZ(0.5, 0.5, 0), 'red_width': 1, 'red_height': 1,
                'blue_label_min_y': 0, 'blue_label_max_y': 0, 'blue_label_height': 0,
                'blue_label_width': 1.0,
                'view_to_label_gap': 0, 'rotation': 0,
            }
        
        # Keep green clip region as optional fallback for type-matching edge cases
        try:
            outline = view.Outline
            geo['green_clip_width'] = outline.Max.U - outline.Min.U
            geo['green_clip_height'] = outline.Max.V - outline.Min.V
        except:
            geo['green_clip_width'] = None
            geo['green_clip_height'] = None
        
        return geo

    def get_pure_view_dimensions(self, vp, view):
        try:
            # view.Outline is already in paper space (Revit pre-divides by the
            # view scale internally) - do NOT divide by scale again here.
            outline = view.Outline
            w = outline.Max.U - outline.Min.U
            h = outline.Max.V - outline.Min.V
            from Autodesk.Revit.DB import ViewportRotation
            rot = vp.Rotation
            if rot == ViewportRotation.Clockwise or rot == ViewportRotation.Counterclockwise:
                return h, w
            return w, h
        except:
            box = vp.GetBoxOutline()
            return box.MaximumPoint.X - box.MinimumPoint.X, box.MaximumPoint.Y - box.MinimumPoint.Y

    def get_checked_views(self):
        seen_vp_ids = set()
        checked_nodes = []
        for root in self.Sheets:
            for s_node in root.Children:
                for v_node in s_node.Views:
                    if v_node.IsChecked and getattr(v_node, "HasTitleNumber", False):
                        vp = v_node.Viewport
                        if vp:
                            vp_id = get_id(vp.Id)
                            if vp_id not in seen_vp_ids:
                                seen_vp_ids.add(vp_id)
                                checked_nodes.append(v_node)
        return checked_nodes

    def _arrange_titles_logic(self, pending, min_spacing, premeasured_heights=None):
        doc.Regenerate()  # Force Revit to update element outlines after any moves
        
        # Read phase: measure all title heights first before making any modifications
        title_heights = premeasured_heights or {}
        for v_node in pending:
            vp = v_node.Viewport
            if not vp: continue
            if vp.Id in title_heights: continue
            
            geo = self.get_viewport_geometry(vp, v_node.View)
            title_h = geo['blue_label_height'] if geo['blue_label_height'] > 0.001 else self._get_title_height_for_type(vp.GetTypeId())
            title_heights[vp.Id] = title_h
            
        # Write phase: apply offsets
        count = 0
        for v_node in pending:
            vp = v_node.Viewport
            if not vp: continue
            
            title_h = title_heights.get(vp.Id, self._get_title_height_for_type(vp.GetTypeId()))
            
            # Label is always placed below the viewport
            shift_y = -(title_h + min_spacing)
            
            # Align label horizontally to the left (X=0.0) and set vertical spacing
            vp.LabelOffset = XYZ(0.0, shift_y, 0)
            count += 1
        return count

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
                
                try:
                    min_spacing = UnitHelper.parse_unit_to_internal(self.TitleOffsetY)
                except:
                    min_spacing = UnitUtils.ConvertToInternalUnits(0.5, UnitHelper.get_sheet_length_unit())

                # Pass A: status/collision check and best-fit type assignment.
                # Type must be finalized and regenerated before we can measure the
                # real rendered label height in Pass B.
                pending = []
                type_changed = False
                for v_node in checked_views:
                    vp = v_node.Viewport
                    view = v_node.View
                    if not vp: continue

                    sheet = doc.GetElement(vp.SheetId)
                    tbs = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
                    tb_min_x, tb_min_y, tb_max_x, tb_max_y = 0.0, 0.0, 3.0, 2.0
                    if tbs:
                        tb_box = tbs[0].get_BoundingBox(sheet)
                        if tb_box: tb_min_x, tb_min_y, tb_max_x, tb_max_y = tb_box.Min.X, tb_box.Min.Y, tb_box.Max.X, tb_box.Max.Y

                    safe_width, safe_height = (tb_max_x - tb_min_x) - (12.5 / 12.0), (tb_max_y - tb_min_y) - (1.0 / 12.0)

                    geo = self.get_viewport_geometry(vp, view)
                    pure_w = geo['red_width']
                    pure_h = geo['red_height']
                    pure_center = geo['red_center']

                    pure_min_x = geo['red_min_x']
                    pure_min_y = geo['red_min_y']
                    max_x = geo['red_max_x']
                    max_y = geo['red_max_y']

                    has_collision = False
                    schedules = FilteredElementCollector(doc, sheet.Id).OfClass(ScheduleSheetInstance).ToElements()

                    for sch in schedules:
                        sch_box = sch.get_BoundingBox(sheet)
                        if sch_box and not (max_x < sch_box.Min.X or pure_min_x > sch_box.Max.X or max_y < sch_box.Min.Y or pure_min_y > sch_box.Max.Y):
                            has_collision = True; break
                    if not has_collision:
                        for other_vpid in sheet.GetAllViewports():
                            if other_vpid != vp.Id:
                                other_vp = doc.GetElement(other_vpid)
                                other_view = doc.GetElement(other_vp.ViewId)
                                o_geo = self.get_viewport_geometry(other_vp, other_view)
                                o_min_x = o_geo['red_min_x']
                                o_max_x = o_geo['red_max_x']
                                o_min_y = o_geo['red_min_y']
                                o_max_y = o_geo['red_max_y']

                                if not (max_x < o_min_x or pure_min_x > o_max_x or max_y < o_min_y or pure_min_y > o_max_y):
                                    has_collision = True; break

                    if has_collision:
                        status = "Collision"
                    elif pure_w > safe_width or pure_h > safe_height:
                        status = "Oversized"
                    else:
                        status = "Arranged"
                    v_node.TitleStatus = status

                    if oxyzen_types:
                        best_type_id = oxyzen_types[-1]['type_id']
                        for ot in oxyzen_types:
                            if ot['length'] >= pure_w:
                                best_type_id = ot['type_id']
                                break
                        if vp.GetTypeId() != best_type_id:
                            vp.ChangeTypeId(best_type_id)
                            type_changed = True

                    pending.append(v_node)

                if type_changed:
                    doc.Regenerate()  # title geometry must reflect the new type before we measure it

                count = self._arrange_titles_logic(pending, min_spacing)

                t.Commit()
            self.StatusText = "Arranged titles for {} views.".format(count)
        except Exception as e:
            self.StatusText = "Error occurred."
            forms.alert("An error occurred:\\n{}".format(traceback.format_exc()))
        finally: self.window.Cursor = Cursors.Arrow

    # ======================================================================
    # NEW: Standalone Arrange-Labels method (label-only, zero type interference)
    # ======================================================================
    def _arrange_view_labels(self, v_nodes, sheet, anchor_corner="BottomRight", 
                              min_spacing=None, title_block_margin_top=0.0,
                              title_block_margin_bottom=0.0, view_extent_gap=0.0):
        """Arrange viewport labels relative to a title block's anchor corner.

        Rules:
          - Works on ALL eligible viewports on the sheet (filtered by applicable view types).
          - Eligible: ViewType has a view number / label (NOT Legend, NOT views with empty DetailNumber).
          - Uses title-block bounding box as the anchor reference.
          - First viewport is anchored to the chosen corner of the title block.
          - All subsequent labels flow from there using viewport extents + computed gaps.
          - ZERO type changes — only manipulates vp.LabelOffset.
        """
        if not v_nodes or not sheet:
            return 0

        t = Transaction(doc, "Arrange View Labels")
        count = 0
        try:
            t.Start()
            doc.Regenerate()
            
            # --- Title block bounding box (anchor reference) ----
            tb_box_min_x, tb_box_min_y, tb_box_max_x, tb_box_max_y = 0.0, 0.0, 3.0, 2.0
            tbs = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
            if tbs:
                tb_box = tbs[0].get_BoundingBox(sheet)
                if tb_box:
                    tb_box_min_x, tb_box_min_y, tb_box_max_x, tb_box_max_y = (
                        tb_box.Min.X, tb_box.Min.Y, tb_box.Max.X, tb_box.Max.Y
                    )

            # --- Determine usable margins from title block clear spaces ----
            is_m = UnitHelper.is_metric()
            try: right_margin = UnitHelper.parse_unit_to_internal(self.RightMargin)
            except: right_margin = UnitUtils.ConvertToInternalUnits(250.0 if is_m else 10.0, UnitHelper.get_sheet_length_unit())
            try: top_margin = UnitHelper.parse_unit_to_internal(self.TopMargin)
            except: top_margin = 0.0
            try: left_margin = UnitHelper.parse_unit_to_internal(self.LeftMargin)
            except: left_margin = right_margin

            sheet_unit = UnitHelper.get_sheet_length_unit()
            if min_spacing is None:
                min_spacing = UnitUtils.ConvertToInternalUnits(0.5, sheet_unit)

            # --- Anchor corner → start point ----
            if anchor_corner == "BottomRight":
                start_x, start_y = tb_box_max_x - right_margin - left_margin, tb_box_min_y + title_block_margin_bottom
            elif anchor_corner == "TopRight":
                start_x, start_y = tb_box_max_x - right_margin - left_margin, tb_box_max_y - title_block_margin_top
            elif anchor_corner == "BottomLeft":
                start_x, start_y = tb_box_min_x + left_margin, tb_box_min_y + title_block_margin_bottom
            elif anchor_corner == "TopLeft":
                start_x, start_y = tb_box_min_x + left_margin, tb_box_max_y - title_block_margin_top
            else:
                start_x, start_y = tb_box_max_x - right_margin - left_margin, tb_box_min_y + title_block_margin_bottom

            # --- Eligibility filter for Arrange Titles ----
            eligible_view_types = {
                ViewType.FloorPlan, ViewType.Section, ViewType.Elevation,
                ViewType.Detail, ViewType.CeilingPlan, ViewType.AreaPlan,
            }
            applicable_views = []
            for v_node in v_nodes:
                vp = v_node.Viewport
                view = v_node.View

                # Must be an applicable view type (exclude legends)
                if view.ViewType not in eligible_view_types:
                    continue

                # Must have a non-empty detail number / view label text
                detail_num_param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                detail_num = None
                if detail_num_param and detail_num_param.HasValue:
                    detail_num = detail_num_param.AsString()
                
                if not detail_num or not detail_num.strip():
                    continue

                # Must have a non-empty view title (optional but recommended)
                view_title_param = vp.get_Parameter(BuiltInParameter.VIEWPORT_TITLE_OVERWRITE_PARAM)
                view_title = None
                if view_title_param and view_title_param.HasValue:
                    view_title = view_title_param.AsString()
                
                # Also check the actual View.Name as fallback — if both are empty/space-only skip
                has_label = (detail_num and detail_num.strip()) or (view_title and view_title.strip()) or view.Name and view.Name.strip()
                if not has_label:
                    continue

                applicable_views.append(v_node)

            if not applicable_views:
                t.Commit()
                return 0

            # --- Layout engine: anchor-based flow from title block corner ----
            # Current label state per viewport tracked as (currentX, currentY)
            curr_x = start_x
            curr_y = start_y
            
            dx = 1 if 'Left' in anchor_corner else -1
            
            for v_node in applicable_views:
                vp = v_node.Viewport
                view = v_node.View
                
                try:
                    geo = self.get_viewport_geometry(vp, view)
                    vp_min_y = geo['red_min_y']
                    vp_max_y = geo['red_max_y']
                    vp_w = geo['red_width']
                    vp_h = geo['red_height']
                except:
                    continue

                try:
                    label_box = vp.GetLabelOutline()
                    real_h = label_box.MaximumPoint.Y - label_box.MinimumPoint.Y
                    title_h = real_h if real_h > 0.001 else self._get_title_height_for_type(vp.GetTypeId())
                except:
                    title_h = self._get_title_height_for_type(vp.GetTypeId())

                # Viewport extent (from GetBoxOutline) — the physical box including frame/border
                vp_min_x = geo['red_min_x']
                vp_max_x = geo['red_max_x']
                
                if anchor_corner in ("BottomRight", "TopRight"):
                    # Pack right-to-left
                    center_x = curr_x - (vp_w / 2.0)
                    
                    # Gap between rows: view extent bottom to previous label top
                    gap = title_h + min_spacing
                    next_min_y = vp_max_y + title_h + min_spacing
                elif anchor_corner in ("BottomLeft", "TopLeft"):
                    # Pack left-to-right (new feature for Arrange Titles)
                    center_x = curr_x + (vp_w / 2.0)
                    
                    gap = title_h + min_spacing
                    next_min_y = vp_max_y + title_h + min_spacing
                else:
                    center_x = curr_x - (vp_w / 2.0)
                    gap = title_h + min_spacing
                    next_min_y = vp_max_y + title_h + min_spacing

                # Calculate the proposed label offset to achieve the desired label position
                new_label_offset = XYZ(0, -(title_h + view_extent_gap), 0)
                
                try:
                    vp.LabelOffset = new_label_offset
                    count += 1
                except:
                    pass

                curr_x = curr_x + dx * (vp_w + min_spacing)

            t.Commit()
        except Exception as e:
            t.RollBack()
            self.StatusText = "Error during Arrange Titles."
            forms.alert("An error occurred:\\n{}".format(traceback.format_exc()))
        return count

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
            try:
                min_spacing = UnitHelper.parse_unit_to_internal(self.TitleOffsetY)
            except:
                min_spacing = UnitUtils.ConvertToInternalUnits(0.5, sheet_unit)
            title_height_est = min_spacing
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
                    view = v_node.View
                    w, h = self.get_pure_view_dimensions(vp, view)
                    title_h = self._get_title_height_for_type(vp.GetTypeId())
                    sim_vp.append({'w': w, 'h': h, 'title_h': title_h, 'id': get_id(vp.Id)})

                # quick oversize check
                for item in sim_vp:
                    if item['w'] > usable_width - 1e-6:
                        msg = "Auto-Pack warning: viewport {} on sheet {} width ({:.3f}) exceeds usable width ({:.3f}).".format(item['id'], get_id(sheet.Id), item['w'], usable_width)
                        print(msg)

                # Simulate row packing using 1/2" view to title logic
                rows = []
                curr_row_width = 0.0
                curr_row_max_h = 0.0
                first_in_row = True
                for item in sim_vp:
                    w = item['w']
                    h = item['h']
                    t_h = item.get('title_h', title_height_est)
                    block_h = h + min_spacing + t_h
                    
                    if first_in_row:
                        curr_row_width = w
                        curr_row_max_h = block_h
                        first_in_row = False
                    else:
                        if curr_row_width + padding + w <= usable_width + 1e-6:
                            curr_row_width += padding + w
                            curr_row_max_h = max(curr_row_max_h, block_h)
                        else:
                            rows.append(curr_row_max_h)
                            curr_row_width = w
                            curr_row_max_h = block_h
                if not first_in_row:
                    rows.append(curr_row_max_h)

                num_rows = len(rows)
                required_vertical = sum(rows) + max(0, num_rows - 1) * min_spacing

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
                    try:
                        min_spacing = UnitHelper.parse_unit_to_internal(self.TitleOffsetY)
                    except:
                        min_spacing = UnitUtils.ConvertToInternalUnits(0.5, sheet_unit)
                    title_height_est = min_spacing
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
                        
                        # Pass A: measure pure view size and apply the best-fit title type
                        # BEFORE measuring title height, since height can vary by type.
                        vp_data = []
                        type_changed = False
                        for v_node in v_nodes:
                            vp = v_node.Viewport
                            view = v_node.View
                            geo = self.get_viewport_geometry(vp, view)
                            w = geo['red_width']
                            h = geo['red_height']
                            best_type_id = self.get_best_title_type_id(w, title_types)
                            this_type_changed = bool(best_type_id) and vp.GetTypeId() != best_type_id
                            if this_type_changed:
                                vp.ChangeTypeId(best_type_id)
                                type_changed = True
                            vp_data.append({'v_node': v_node, 'vp': vp, 'w': w, 'h': h, 'best_type_id': best_type_id, 'type_changed': this_type_changed, 'area': w * h, 'id': get_id(vp.Id)})

                        vp_data = sorted(vp_data, key=lambda x: (-x['area'], x['id']))

                        # Assign FINAL detail numbers now, before measuring title height below.
                        # Phase 1 left every viewport on a long "TMP_<id>_<i>" placeholder to
                        # avoid numbering collisions; if we measured the label while that long
                        # placeholder was still showing, it can wrap to extra line(s) and make
                        # GetLabelOutline report a taller box than the real final number will need.
                        taken_numbers = set()
                        if self.SyncDetailNumbers:
                            all_vp_ids = sheet.GetAllViewports()
                            checked_vp_ids = set(x['id'] for x in vp_data)
                            for vp_id in all_vp_ids:
                                if get_id(vp_id) not in checked_vp_ids:
                                    p = doc.GetElement(vp_id).get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                    if p and p.HasValue: taken_numbers.add(p.AsString())

                            detail_num = 1
                            for data in vp_data:
                                vp = data['vp']
                                v_node = data['v_node']
                                p = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                if p and not p.IsReadOnly:
                                    while str(detail_num) in taken_numbers: detail_num += 1
                                    p.Set(str(detail_num))
                                    taken_numbers.add(str(detail_num))
                                    v_node.DetailNumber = str(detail_num)
                                    detail_num += 1

                        if type_changed or self.SyncDetailNumbers:
                            doc.Regenerate()  # title text/geometry must be final before we measure it

                        # Pass B: now that types/numbers/geometry are final, read the ACTUAL
                        # rendered label height via GetLabelOutline instead of guessing from
                        # family parameters - LabelOffset only translates the label, so its size
                        # here is valid regardless of what offset is applied afterward.
                        for data in vp_data:
                            vp = data['vp']
                            geo = self.get_viewport_geometry(vp, data['v_node'].View)
                            box = vp.GetBoxOutline()
                            data['center'] = geo['red_center']
                            data['red_min_x'] = box.MinimumPoint.X
                            data['red_min_y'] = box.MinimumPoint.Y
                            data['title_w'] = geo['blue_label_width']
                            try:
                                label_box = vp.GetLabelOutline()
                                real_h = label_box.MaximumPoint.Y - label_box.MinimumPoint.Y
                                data['title_h'] = real_h if real_h > 0.001 else self._get_title_height_for_type(vp.GetTypeId())
                            except:
                                data['title_h'] = self._get_title_height_for_type(vp.GetTypeId())

                        dx = 1 if 'Left' in anchor else -1
                        dy = 1 if 'Bottom' in anchor else -1

                        first_row = True
                        first_in_row = True
                        curr_left = usable_min_x
                        curr_right = usable_max_x
                        row_boundary_y = 0.0
                        
                        for data in vp_data:
                            w, h, vp, v_node = data['w'], data['h'], data['vp'], data['v_node']
                            title_h = data['title_h']
                            best_type_id = data['best_type_id']

                            if first_in_row:
                                if dx == 1:
                                    curr_left = usable_min_x
                                    current_x = curr_left
                                else:
                                    curr_right = usable_max_x
                                    current_x = curr_right
                                first_in_row = False
                                
                                # Initialize row vertical start
                                if first_row:
                                    if dy == 1:
                                        curr_row_view_start = tb_min_y + b_margin + tb_offset_y + title_h + min_spacing
                                    else:
                                        curr_row_view_start = tb_max_y - t_margin - tb_offset_y
                                    first_row = False
                                row_boundary_y = curr_row_view_start
                            else:
                                if dx == 1:
                                    curr_left = prev_right + padding
                                    if curr_left + w > usable_max_x + 1e-6: # Wrap to next row
                                        if dy == 1:
                                            curr_row_view_start = row_boundary_y + min_spacing + title_h + min_spacing
                                        else:
                                            curr_row_view_start = row_boundary_y - min_spacing
                                        curr_left = usable_min_x
                                        row_boundary_y = curr_row_view_start
                                        current_x = curr_left
                                    else:
                                        current_x = curr_left
                                else:
                                    curr_right = prev_left - padding
                                    if curr_right - w < usable_min_x - 1e-6: # Wrap to next row
                                        if dy == 1:
                                            curr_row_view_start = row_boundary_y + min_spacing + title_h + min_spacing
                                        else:
                                            curr_row_view_start = row_boundary_y - min_spacing
                                        curr_right = usable_max_x
                                        row_boundary_y = curr_row_view_start
                                        current_x = curr_right
                                    else:
                                        current_x = curr_right
                                        
                            # Calculate coordinates before moving
                            if dy == 1:
                                target_view_min_y = curr_row_view_start
                                target_view_max_y = target_view_min_y + h
                                target_title_y = target_view_min_y - min_spacing - title_h
                            else:
                                target_view_max_y = curr_row_view_start
                                target_view_min_y = target_view_max_y - h
                                target_title_y = target_view_min_y - min_spacing - title_h

                            if dx == 1:
                                target_view_min_x = current_x
                                target_title_x = current_x
                            else:
                                target_view_min_x = current_x - w
                                target_title_x = target_view_min_x
                                
                            tgt_min_x = target_view_min_x
                            target_min_y = target_view_min_y
                            v_max_y = target_view_max_y
                            t_y = target_title_y
                                
                            if dy == 1: 
                                row_boundary_y = max(row_boundary_y, v_max_y)
                            else: 
                                row_boundary_y = min(row_boundary_y, t_y)

                            if dx == 1: prev_right = current_x + w
                            else: prev_left = current_x - w
                            # Place Viewport directly using target View center
                            target_center_x = tgt_min_x + w / 2.0
                            target_center_y = target_min_y + h / 2.0
                            vp.SetBoxCenter(XYZ(target_center_x, target_center_y, 0))
                            
                            if data['type_changed']:
                                v_node.TitleStatus = "Matched"
                            v_node.TitleStatus = "Packed"
                            
                            count += 1

                        # Run the arrange titles logic to position all labels in one shot
                        pending_views = [data['v_node'] for data in vp_data]
                        premeasured = {data['vp'].Id: data['title_h'] for data in vp_data}
                        self._arrange_titles_logic(pending_views, min_spacing, premeasured_heights=premeasured)

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