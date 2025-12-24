# -*- coding: utf-8 -*-
"""
System Merger Tool

Description:
    A Modal WPF tool to visualize and merge disconnected pipe networks.
    Identifies "Islands" of connected pipes and merges them based on logic:
    - Vent Systems: Largest Volume wins (Master).
    - Pressure Systems: Network with Base Equipment wins.

Architecture:
    - Modal Window (ShowDialog): Ensures thread safety and prevents context crashes.
    - Config Persistence: Remembers window position/size.
    - Transactions: Explicitly handles all model changes.
"""

import os
import traceback
import math
import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.Generic import List
from System.Windows.Media import Colors, SolidColorBrush, Color as WpfColor
from System.Windows.Input import Cursors
from Autodesk.Revit.DB import (
    Transaction, BuiltInCategory, ElementId, FilteredElementCollector,
    OverrideGraphicSettings, Color, FillPatternElement, ElementTransformUtils, XYZ,
    BuiltInParameter, ElementMulticategoryFilter, Line
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, script, revit

# Try to import UIThemeManager (Revit 2024+)
try:
    from Autodesk.Revit.UI import UIThemeManager, UITheme
    HAS_THEME = True
except ImportError:
    HAS_THEME = False

__title__ = "System Merger"
__doc__ = "Modal tool to diagnose and merge disconnected pipe networks."
__context__ = "active-view-type: FloorPlan,CeilingPlan,EngineeringPlan,AreaPlan,Section,Elevation,ThreeD"

# Helper for Revit 2024+ compatibility
def get_id(element_id):
    if hasattr(element_id, "Value"):
        return element_id.Value
    return element_id.IntegerValue

# --- Data Model ---
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

class NodeBase(ViewModelBase):
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self.Name = name
        self._is_checked = False
        self._is_selected = False
        self.IsExpanded = True
        self.Children = []
        self.Type = "Item"
        self.Abbreviation = ""
        self.Length = "-"
        self.FixtureUnits = "-"
        self.Volume = "-"
        self.Count = "-"
        self.FontWeight = "Normal"
        self.AllElements = [] # Flat list of element IDs for highlighting
        self.NetworkColor = SolidColorBrush(Colors.Black)

    @property
    def IsChecked(self):
        return self._is_checked

    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value
        self.OnPropertyChanged("IsChecked")

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged("IsSelected")

class ClassificationNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self, name)
        self.FontWeight = "Bold"
        self.Type = "Classification"
        
    def aggregate_stats(self):
        """Sum up stats from children Types."""
        vol = sum(float(c.Volume) for c in self.Children if c.Volume)
        fu = sum(getattr(c, 'RawFixtureUnits', 0.0) for c in self.Children)
        length = sum(float(c.Length) for c in self.Children if c.Length)
        self.Volume = "{:.2f}".format(vol)
        self.FixtureUnits = "{:.1f}".format(fu)
        self.Length = "{:.2f}".format(length)
        self.Count = "{} Types".format(len(self.Children))

class TypeNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self, name)
        self.FontWeight = "Bold"
        self.Type = "System Type"

    def aggregate_stats(self):
        """Sum up stats from children systems."""
        vol = sum(float(c.Volume) for c in self.Children if c.Volume)
        fu = sum(getattr(c, 'RawFixtureUnits', 0.0) for c in self.Children)
        length = sum(float(c.Length) for c in self.Children if c.Length)
        self.RawFixtureUnits = fu # Store for parent aggregation
        self.Volume = "{:.2f}".format(vol)
        self.FixtureUnits = "{:.1f}".format(fu)
        self.Length = "{:.2f}".format(length)
        self.Count = "{} Systems".format(len(self.Children))

class ElementNode(NodeBase):
    """Represents a leaf node (Fixture/Equipment) in the tree."""
    def __init__(self, element):
        # Use Element Name (Family + Type usually) or Family Name
        # Format: "Family : Type" if possible, else Name
        name = element.Name
        if hasattr(element, "Symbol") and element.Symbol:
            name = "{} : {}".format(element.Symbol.FamilyName, element.Name)
        
        NodeBase.__init__(self, name)
        self.FontWeight = "Normal"
        self.AllElements = [element.Id]
        self.IsExpanded = False
        self.NetworkColor = SolidColorBrush(Colors.Gray)
        self.Type = element.Category.Name if element.Category else "Element"
        self.Count = "1"
        
        # Get FU for individual element if exists
        p_fu = element.get_Parameter(BuiltInParameter.RBS_PIPE_FIXTURE_UNITS_PARAM)
        if p_fu: self.FixtureUnits = "{:.1f}".format(p_fu.AsDouble())

        # Get Length and Volume for Pipes
        l_param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        d_param = element.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM)
        
        if l_param:
            len_val = l_param.AsDouble()
            self.Length = "{:.2f}".format(len_val)
            if d_param:
                r = d_param.AsDouble() / 2.0
                vol = math.pi * (r**2) * len_val
                self.Volume = "{:.3f}".format(vol)
        
        # Type column removed from UI, so we don't need to set self.Type explicitly for display
        # but we keep the logic clean.

class NetworkNode(NodeBase):
    def __init__(self, name, network_data, child_elements=None):
        NodeBase.__init__(self, name)
        self.FontWeight = "Normal"
        self.Type = "Network"
        self.Volume = "{:.2f}".format(network_data.volume)
        self.FixtureUnits = "{:.1f}".format(network_data.fixture_units)
        self.Length = "{:.2f}".format(network_data.length)
        self.Count = "{} Elem".format(network_data.count)
        self.AllElements = network_data.elements
        self.IsExpanded = True
        
        if child_elements:
            self.Children = [ElementNode(el) for el in child_elements]

class SystemNode(NodeBase):
    def __init__(self, name, sys_type, sys_abbr, sys_fu, networks, child_elements=None):
        NodeBase.__init__(self, name)
        self.Type = sys_type
        self.Abbreviation = sys_abbr
        self.FontWeight = "SemiBold"
        self.RawFixtureUnits = sys_fu
        
        # Aggregate
        vol = sum(n.volume for n in networks)
        length = sum(n.length for n in networks)
        count = sum(n.count for n in networks)
        # Note: sys_fu passed from MEPSystem parameter is usually total for system
        
        self.Volume = "{:.2f}".format(vol)
        self.Length = "{:.2f}".format(length)
        self.Count = "{} Networks".format(len(networks))
        
        # Highlight if system is split (more than 1 network)
        if len(networks) > 1:
            self.NetworkColor = SolidColorBrush(Colors.Red)
            self.FixtureUnits = "{:.1f} (Split)".format(sys_fu)
            
            # Create Network Nodes for split systems
            self.Children = []
            for i, net in enumerate(networks):
                # Filter child elements for this network
                net_children = []
                if child_elements:
                    net_ids_set = {get_id(eid) for eid in net.elements}
                    net_children = [el for el in child_elements if get_id(el.Id) in net_ids_set]
                
                self.Children.append(NetworkNode("Network {} (Island)".format(i+1), net, net_children))
        else:
            # Single System: Direct Children
            self.FixtureUnits = "{:.1f}".format(sys_fu)
            if child_elements:
                self.Children = [ElementNode(el) for el in child_elements]
        
        for n in networks:
            self.AllElements.extend(n.elements)

class NetworkData:
    """Raw data holder for processing."""
    def __init__(self, volume, length, fu, count, elements):
        self.volume = volume
        self.length = length
        self.fixture_units = fu
        self.count = count
        self.elements = elements

class ColorOption:
    def __init__(self, name, r, g, b):
        self.Name = name
        self.R = r
        self.G = g
        self.B = b
        self.Brush = SolidColorBrush(WpfColor.FromRgb(r, g, b))
        self.RevitColor = Color(r, g, b)

# --- Main Window Class ---
class SystemMergeWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'ui.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        # UI Event Bindings for Custom Title Bar
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_WinClose.Click += self.close_window
        self.Closing += self.window_closing
        
        self.Btn_SelectAll.Click += self.select_all_click
        self.Btn_Clear.Click += self.clear_list_click
        self.Btn_ExpandAll.Click += self.expand_all_click
        self.Btn_CollapseAll.Click += self.collapse_all_click
        self.Btn_ScanView.Click += self.scan_view_click
        self.Btn_Visualize.Click += self.visualize_click
        self.Btn_ClearVisuals.Click += self.reset_visuals_click
        self.Btn_Disconnect.Click += self.disconnect_click
        self.Btn_Rename.Click += self.rename_click
        
        # Handle TreeView Selection via ItemContainerStyle Binding
        # We no longer use SelectedItemChanged, but we can listen to property changes if needed.
        # However, for the logic, we can just iterate or bind commands. 
        # For simplicity in this hybrid approach, we will hook into the TreeView's SelectedItemChanged 
        # just to trigger the visualization logic, but rely on the ViewModel for state.
        self.systemTree.SelectedItemChanged += self.tree_selection_changed
        
        # Initial UI State: Disable actions until data is loaded
        self.Btn_SelectAll.IsEnabled = False
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Disconnect.IsEnabled = False
        self.Btn_Rename.IsEnabled = False
        
        # Default Header
        self.set_default_header()
        
        self.load_window_settings()
        self.doc = revit.doc
        self.uidoc = revit.uidoc

        self.last_highlighted_ids = []
        self.is_busy = False
        self.disabled_filters = [] # Track filters we disable to restore them later
        self.populate_colors()

        self.apply_revit_theme()

    def set_default_header(self):
        default_node = NodeBase("System Network Browser")
        default_node.Type = "Select an item to view details"
        self.RightPane.DataContext = default_node

    # --- UI Logic ---
    def drag_window(self, sender, args):
        self.DragMove()

    def close_window(self, sender, args):
        self.Close()

    def apply_revit_theme(self):
        """Detects Revit theme and updates window resources if Dark."""
        is_dark = False
        if HAS_THEME:
            try:
                if UIThemeManager.CurrentTheme == UITheme.Dark:
                    is_dark = True
            except: pass
        
        if is_dark:
            # Define Dark Theme Colors
            res = self.Resources
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(45, 45, 45))      # #2D2D2D
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(56, 56, 56))     # #383838
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))     # #F0F0F0
            res["TextLightBrush"] = SolidColorBrush(WpfColor.FromRgb(170, 170, 170))# #AAAAAA
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(80, 80, 80))      # #505050
            res["AccentBrush"] = SolidColorBrush(WpfColor.FromRgb(0, 96, 192))      # #0060C0
            res["SelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(64, 80, 96))   # #405060
            res["SelectionBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(80, 112, 144))
            res["HoverBrush"] = SolidColorBrush(WpfColor.FromRgb(58, 58, 58))       # #3A3A3A
            res["FooterBrush"] = SolidColorBrush(WpfColor.FromRgb(51, 51, 51))      # #333333
            res["AltRowBrush"] = SolidColorBrush(WpfColor.FromRgb(66, 66, 66))      # #424242
            # Keep CardBrush dark (#282a2f) as it fits well in Dark mode too

    # --- Persistence Logic ---
    def load_window_settings(self):
        """Restores window position and size from config."""
        cfg = script.get_config()
        self.Top = cfg.get_option('win_top', 200)
        self.Left = cfg.get_option('win_left', 200)
        self.Width = cfg.get_option('win_width', 800)
        self.Height = cfg.get_option('win_height', 450)

    def window_closing(self, sender, args):
        """Saves window position and size to config on close."""
        cfg = script.get_config()
        cfg.win_top = self.Top
        cfg.win_left = self.Left
        cfg.win_width = self.Width
        cfg.win_height = self.Height
        # Ensure we clean up view overrides when the window closes
        self.reset_selection_highlight()
        script.save_config()

    def populate_colors(self):
        """Generates a list of 50 distinct colors for the dropdown."""
        # Basic list of distinct colors
        base_colors = [
            ("Red", 255, 0, 0), ("Green", 0, 255, 0), ("Blue", 0, 0, 255),
            ("Yellow", 255, 255, 0), ("Cyan", 0, 255, 255), ("Magenta", 255, 0, 255),
            ("Orange", 255, 165, 0), ("Purple", 128, 0, 128), ("Lime", 50, 205, 50),
            ("Pink", 255, 192, 203), ("Teal", 0, 128, 128), ("Lavender", 230, 230, 250),
            ("Brown", 165, 42, 42), ("Beige", 245, 245, 220), ("Maroon", 128, 0, 0),
            ("Mint", 189, 252, 201), ("Olive", 128, 128, 0), ("Coral", 255, 127, 80),
            ("Navy", 0, 0, 128), ("Grey", 128, 128, 128), ("Gold", 255, 215, 0),
            ("Indigo", 75, 0, 130), ("Turquoise", 64, 224, 208), ("Violet", 238, 130, 238),
            ("Salmon", 250, 128, 114), ("Khaki", 240, 230, 140), ("Plum", 221, 160, 221)
        ]
        # Add more if needed or repeat with slight variations
        self.color_options = [ColorOption(n, r, g, b) for n, r, g, b in base_colors]
        self.Cmb_Colors.ItemsSource = self.color_options
        self.Cmb_Colors.SelectedIndex = 0

    # --- Network Logic ---
    def update_button_states(self):
        """Updates enable/disable state of action buttons based on checked items."""
        checked = self.get_checked_systems()
        has_checked = len(checked) > 0
        has_selection = self.systemTree.SelectedItem is not None
        
        self.Btn_Visualize.IsEnabled = has_checked or has_selection
        self.Btn_ClearVisuals.IsEnabled = has_checked or has_selection
        self.Btn_Disconnect.IsEnabled = has_checked or has_selection

    def expand_all_click(self, sender, args):
        self._set_expansion_state(True)

    def collapse_all_click(self, sender, args):
        self._set_expansion_state(False)

    def _set_expansion_state(self, is_expanded):
        if self.systemTree.ItemsSource:
            for node in self.systemTree.ItemsSource:
                self._recursive_expand(node, is_expanded)
            self.systemTree.Items.Refresh()

    def _recursive_expand(self, node, is_expanded):
        node.IsExpanded = is_expanded
        if hasattr(node, "Children"):
            for child in node.Children:
                self._recursive_expand(child, is_expanded)

    def select_all_click(self, sender, args):
        """Toggles between checking and unchecking all items in the tree."""
        if not self.systemTree.ItemsSource: return
        
        # Determine action based on current button text
        is_select_all = (self.Btn_SelectAll.Content == "Select All")
        target_state = True if is_select_all else False
        
        # Toggle Button Text
        self.Btn_SelectAll.Content = "Select None" if is_select_all else "Select All"

        for node in self.systemTree.ItemsSource:
            self._cascade_check(node, target_state)
        self.systemTree.Items.Refresh()
        self.update_button_states()

    def _cascade_check(self, node, state):
        """Recursively sets IsChecked state."""
        node.IsChecked = state
        if hasattr(node, "Children"):
            for child in node.Children:
                self._cascade_check(child, state)

    def clear_list_click(self, sender, args):
        """Clears the list and resets selection overrides."""
        if self.is_busy: return
        
        self.systemTree.ItemsSource = None
        self.reset_selection_highlight()
        self.statusLabel.Text = "List cleared."
        self.Btn_SelectAll.Content = "Select All"
        self.Btn_SelectAll.IsEnabled = False
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Disconnect.IsEnabled = False
        self.Btn_Rename.IsEnabled = False
        self.Tb_NewName.Text = ""
        self.set_default_header()

    def scan_view_click(self, sender, args):
        """Scans all pipe elements in the active view."""
        if self.is_busy: return
        self.is_busy = True
        self.Cursor = Cursors.Wait
        
        try:
            valid_cats = [
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_Sprinklers
            ]
            cat_filter = ElementMulticategoryFilter(List[BuiltInCategory](valid_cats))
            collector = FilteredElementCollector(self.doc, self.doc.ActiveView.Id).WherePasses(cat_filter)
            ids = [e.Id for e in collector]
            self.analyze_selection(ids)
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Error scanning view. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow

    def analyze_selection(self, element_ids):
        """Core logic to filter selection, run BFS, and populate Grid."""
        # Reset UI State
        self.systemTree.ItemsSource = None
        self.Btn_SelectAll.Content = "Select All"
        self.Btn_SelectAll.IsEnabled = False
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Disconnect.IsEnabled = False
        self.Btn_Rename.IsEnabled = False
        self.Tb_NewName.Text = ""
        self.set_default_header()

        try:
            if not element_ids:
                self.statusLabel.Text = "No elements to analyze."
                return
                
            # Filter for pipes, fittings, and accessories to ensure connectivity
            valid_cats = [
                int(BuiltInCategory.OST_PipeCurves),
                int(BuiltInCategory.OST_PipeFitting),
                int(BuiltInCategory.OST_PipeAccessory),
                int(BuiltInCategory.OST_PlumbingFixtures),
                int(BuiltInCategory.OST_MechanicalEquipment),
                int(BuiltInCategory.OST_Sprinklers)
            ]
            
            elements_to_process = []
            for eid in element_ids:
                el = self.doc.GetElement(eid)
                if el and el.Category and get_id(el.Category.Id) in valid_cats:
                    elements_to_process.append(el)
            
            if not elements_to_process:
                forms.alert("No valid Pipe, Fitting, or Accessory elements found in selection.")
                self.statusLabel.Text = "Selection contained no valid elements."
                return

            # 1. Group by System Identity First (Revit Logic)
            # Key: (Class, Type, Name) -> Data
            system_buckets = {} 

            # Helper to validate strings
            def is_valid(s):
                return s and s != "Undefined" and s != "Unassigned"

            for el in elements_to_process:
                sys_name = None
                sys_type = None
                sys_class = None
                sys_abbr = None
                mep_sys_obj = None

                # Try to get MEPSystem object directly
                if hasattr(el, "MEPSystem") and el.MEPSystem:
                    mep_sys_obj = el.MEPSystem
                    if is_valid(mep_sys_obj.Name): sys_name = mep_sys_obj.Name
                
                # Fallback to parameters if MEPSystem didn't give name
                if not is_valid(sys_name):
                    p = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
                    if p and is_valid(p.AsString()): sys_name = p.AsString()
                
                # Fallback to Connectors (Crucial for Fittings)
                if not is_valid(sys_name) or not mep_sys_obj:
                    mgr = None
                    if hasattr(el, "ConnectorManager"): mgr = el.ConnectorManager
                    elif hasattr(el, "MEPModel") and el.MEPModel: mgr = el.MEPModel.ConnectorManager
                    
                    if mgr:
                        for c in mgr.Connectors:
                            if c.MEPSystem and is_valid(c.MEPSystem.Name):
                                mep_sys_obj = c.MEPSystem
                                sys_name = mep_sys_obj.Name
                                break
                
                # If we have an MEPSystem object, get Type/Class/Abbr from it
                if mep_sys_obj:
                    type_id = mep_sys_obj.GetTypeId()
                    if type_id != ElementId.InvalidElementId:
                        t_elem = self.doc.GetElement(type_id)
                        if t_elem:
                            p_name = t_elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                            if p_name and is_valid(p_name.AsString()): sys_type = p_name.AsString()
                            else: sys_type = getattr(t_elem, "Name", "Unassigned")
                            
                            p_class = t_elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
                            if p_class and is_valid(p_class.AsString()): sys_class = p_class.AsString()
                            
                            p_abbr = t_elem.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)
                            if p_abbr and is_valid(p_abbr.AsString()): sys_abbr = p_abbr.AsString()
                
                # Final Fallbacks
                if not is_valid(sys_name): sys_name = "Unassigned"
                if not is_valid(sys_type): sys_type = "Unassigned"
                if not is_valid(sys_class): sys_class = "Unassigned"
                if not is_valid(sys_abbr): sys_abbr = ""

                key = (sys_class, sys_type, sys_name)
                if key not in system_buckets:
                    system_buckets[key] = {
                        'class': sys_class, 'type': sys_type, 'name': sys_name, 
                        'abbr': sys_abbr, 'mep_sys': mep_sys_obj, 'elements': []
                    }
                system_buckets[key]['elements'].append(el)

            # Build Tree Structure: Classification -> System
            # Group systems by Classification -> System Type
            hierarchy = {} # ClassName -> { TypeName -> [SystemNode] }
            
            # Categories to show as children in the tree
            child_cats = [
                int(BuiltInCategory.OST_PlumbingFixtures),
                int(BuiltInCategory.OST_MechanicalEquipment),
                int(BuiltInCategory.OST_Sprinklers)
            ]

            for key, data in system_buckets.items():
                cls = data['class']
                typ = data['type']
                name = data['name']
                abbr = data['abbr']
                
                # Calculate System-level FU
                sys_fu = 0.0
                if data['mep_sys']:
                    p_fu = data['mep_sys'].get_Parameter(BuiltInParameter.RBS_PIPE_FIXTURE_UNITS_PARAM)
                    if p_fu: sys_fu = p_fu.AsDouble()

                # Run BFS on elements WITHIN this system to find networks (islands)
                islands = self._bfs_traversal(data['elements'])
                
                networks = []
                for island_ids in islands:
                    total_vol = 0.0
                    total_len = 0.0
                    island_fu = 0.0
                    
                    for eid in island_ids:
                        el = self.doc.GetElement(eid)
                        # Volume
                        l_param = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
                        d_param = el.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM)
                        if l_param and d_param:
                            r = d_param.AsDouble() / 2.0
                            total_vol += math.pi * (r**2) * l_param.AsDouble()
                        
                        # Length
                        if l_param:
                            total_len += l_param.AsDouble()
                            
                        # Fixture Units (Sum from fixtures in this island)
                        p_fu_el = el.get_Parameter(BuiltInParameter.RBS_PIPE_FIXTURE_UNITS_PARAM)
                        if p_fu_el: island_fu += p_fu_el.AsDouble()

                    networks.append(NetworkData(round(total_vol, 3), round(total_len, 2), round(island_fu, 1), len(island_ids), island_ids))
                
                if cls not in hierarchy: hierarchy[cls] = {}
                if typ not in hierarchy[cls]: hierarchy[cls][typ] = []
                
                # Collect child elements for this system
                sys_children = []
                for el in data['elements']:
                    if el.Category and get_id(el.Category.Id) in child_cats:
                        sys_children.append(el)

                hierarchy[cls][typ].append(SystemNode(name, typ, abbr, sys_fu, networks, sys_children))
            
            # Create Root Nodes
            root_nodes = []
            for cls_name, type_dict in hierarchy.items():
                c_node = ClassificationNode(cls_name)
                
                type_nodes = []
                for typ_name, sys_nodes in type_dict.items():
                    t_node = TypeNode(typ_name)
                    t_node.Children = sorted(sys_nodes, key=lambda x: x.Name)
                    t_node.aggregate_stats()
                    type_nodes.append(t_node)
                
                c_node.Children = sorted(type_nodes, key=lambda x: x.Name)
                c_node.aggregate_stats()
                root_nodes.append(c_node)

            self.systemTree.ItemsSource = sorted(root_nodes, key=lambda x: x.Name)
            self.statusLabel.Text = "Found {} system classifications.".format(len(root_nodes))
            
            # Enable buttons if data exists
            has_data = len(root_nodes) > 0
            self.Btn_SelectAll.IsEnabled = has_data
            self.Btn_Visualize.IsEnabled = False # Wait for check
            self.Btn_ClearVisuals.IsEnabled = False # Wait for check
            self.Btn_Disconnect.IsEnabled = False # Wait for check
            self.Btn_Rename.IsEnabled = has_data
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Analysis Error. Check Output."

    def reset_selection_highlight(self):
        """Resets the temporary orange highlight on previously selected elements."""
        if self.last_highlighted_ids:
            try:
                with Transaction(self.doc, "Reset Highlight") as t:
                    t.Start()
                    for eid in self.last_highlighted_ids:
                        self.doc.ActiveView.SetElementOverrides(eid, OverrideGraphicSettings())
                    t.Commit()
            except Exception:
                pass
            self.last_highlighted_ids = []

    def on_checkbox_click(self, sender, args):
        """Manually syncs CheckBox state to DataContext."""
        # With MVVM, the binding is TwoWay, so self.IsChecked updates automatically.
        # We just need to handle the cascading logic.
        node = sender.DataContext
        if node:
            # Cascade to children (Classification -> Systems)
            if hasattr(node, "Children"):
                for child in node.Children:
                    child.IsChecked = node.IsChecked
            
            # Cascade Up (Child -> Parent)
            if self.systemTree.ItemsSource:
                # Need recursive check or simple 2-level check. 
                # Since we added TypeNode, it's Class -> Type -> System.
                # For simplicity, we just refresh the tree to let bindings update if we had full MVVM, 
                # but here we might miss the visual update on parents without explicit logic.
                # Given the complexity of 3-level recursion in this simple handler, 
                # we rely on the user checking the parent to select all, which is implemented.
                pass
            
            self.update_button_states()

    def tree_selection_changed(self, sender, args):
        """Syncs TreeView selection with Revit selection (Highlight)."""
        if self.is_busy: return
        
        self.reset_selection_highlight()
        self.update_button_states()
        
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node:
                selected_node.IsSelected = True # Ensure ViewModel is in sync

            if not selected_node: return
            
            # Populate Rename Box
            if isinstance(selected_node, SystemNode):
                self.Tb_NewName.Text = selected_node.Name
            else:
                self.Tb_NewName.Text = ""
            
            ids = set(selected_node.AllElements)
            # If it's a classification node, gather all children
            if isinstance(selected_node, (ClassificationNode, TypeNode)):
                # Recursively gather all elements
                ids = self._get_all_child_elements(selected_node)
            
            if ids:
                # Identify background elements (Rest of the Model)
                # We collect all elements in the active view to apply the dimming effect.
                view_id = self.doc.ActiveView.Id
                collector = FilteredElementCollector(self.doc, view_id).WhereElementIsNotElementType()
                all_view_ids = set(e.Id for e in collector)
                background_ids = all_view_ids - ids

                # Apply Bold Orange Highlight & Dim Background
                with Transaction(self.doc, "Highlight Selection") as t:
                    t.Start()
                    
                    # 1. Highlight Selected
                    ogs_sel = OverrideGraphicSettings()
                    ogs_sel.SetProjectionLineColor(Color(0, 128, 255)) # System Browser Blue
                    ogs_sel.SetProjectionLineWeight(12) # Extra Thick / Glow Effect
                    for eid in ids:
                        self.doc.ActiveView.SetElementOverrides(eid, ogs_sel)
                    
                    # 2. Dim Background (Halftone + Transparent)
                    if background_ids:
                        ogs_dim = OverrideGraphicSettings()
                        ogs_dim.SetHalftone(True)
                        ogs_dim.SetSurfaceTransparency(80) # 80% Transparent
                        for eid in background_ids:
                            self.doc.ActiveView.SetElementOverrides(eid, ogs_dim)
                            
                    t.Commit()
                
                self.last_highlighted_ids = list(ids.union(background_ids))
                self.uidoc.RefreshActiveView()
                
                elem_ids = List[ElementId](ids)
                
                # 1. Select in Revit
                self.uidoc.Selection.SetElementIds(elem_ids)
                
                # 2. Auto-Zoom if enabled
                if self.Cb_AutoZoom.IsChecked:
                    self.uidoc.ShowElements(elem_ids)
            
            # 3. Update Header Context
            self.RightPane.DataContext = selected_node

            # 4. Update DataGrid with Children
            if hasattr(selected_node, "Children") and selected_node.Children:
                self.sysDataGrid.ItemsSource = selected_node.Children
            else:
                self.sysDataGrid.ItemsSource = []
                    
        except Exception:
            pass # Prevent crash if selection fails

    def _get_all_child_elements(self, node):
        """Recursively gets all element IDs from a node and its children."""
        ids = set(node.AllElements)
        if hasattr(node, "Children"):
            for child in node.Children:
                ids.update(self._get_all_child_elements(child))
        return ids

    def get_checked_systems(self):
        """Helper to find all checked SystemNodes or NetworkNodes."""
        checked = []
        if self.systemTree.ItemsSource:
            for class_node in self.systemTree.ItemsSource:
                for type_node in class_node.Children:
                    for sys_node in type_node.Children:
                        # Check if system is split (has NetworkNode children)
                        is_split = False
                        if sys_node.Children and isinstance(sys_node.Children[0], NetworkNode):
                            is_split = True
                        
                        if is_split:
                            # If split, check individual networks so they get colored separately
                            for net_node in sys_node.Children:
                                if net_node.IsChecked:
                                    checked.append(net_node)
                        else:
                            # If not split, check the system itself
                            if sys_node.IsChecked:
                                checked.append(sys_node)
        return checked

    def _generate_dynamic_color(self, index):
        """Generates a distinct color using Golden Ratio for overflow."""
        # Use Golden Ratio Conjugate to spread hues evenly
        golden_ratio = 0.618033988749895
        h = (index * golden_ratio) % 1.0
        s = 0.85 # High saturation for visibility
        v = 0.95 # High value for brightness
        
        # HSV to RGB conversion
        i = int(h * 6)
        f = h * 6 - i
        p = v * (1 - s)
        q = v * (1 - f * s)
        t = v * (1 - (1 - f) * s)
        
        r, g, b = 0, 0, 0
        if i % 6 == 0: r, g, b = v, t, p
        elif i % 6 == 1: r, g, b = q, v, p
        elif i % 6 == 2: r, g, b = p, v, t
        elif i % 6 == 3: r, g, b = p, q, v
        elif i % 6 == 4: r, g, b = t, p, v
        elif i % 6 == 5: r, g, b = v, p, q
        
        return Color(int(r * 255), int(g * 255), int(b * 255))

    def visualize_click(self, sender, args):
        """Applies Neon Color Overrides to visualize islands."""
        if self.is_busy: return
        
        checked_systems = self.get_checked_systems()
        if not checked_systems:
            forms.alert("Please check at least one system to colorize.")
            return
        
        selected_color_opt = self.Cmb_Colors.SelectedItem
        if not selected_color_opt:
            forms.alert("Please select a color from the dropdown.")
            return

        self.is_busy = True
        self.Cursor = Cursors.Wait

        # Determine if we should cycle colors (Multi-selection)
        use_cycling = len(checked_systems) > 1
        start_idx = self.Cmb_Colors.SelectedIndex
        if start_idx < 0: start_idx = 0

        try:
            # Check for View Filters that might mask colors
            view = self.doc.ActiveView
            filters = view.GetFilters()
            if filters:
                visible_filters = [f for f in filters if view.GetFilterVisibility(f)]
                if visible_filters:
                    td = TaskDialog("View Filters Detected")
                    td.MainInstruction = "Active View Filters might mask the tool's colors."
                    td.MainContent = "Do you want to temporarily hide these filters in this view?"
                    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                    if td.Show() == TaskDialogResult.Yes:
                        with Transaction(self.doc, "Disable Filters") as t:
                            t.Start()
                            for fid in visible_filters:
                                view.SetFilterVisibility(fid, False)
                                if fid not in self.disabled_filters:
                                    self.disabled_filters.append(fid)
                            t.Commit()

            with Transaction(self.doc, "Visualize Networks") as t:
                t.Start()
                
                # Find the actual Solid Fill pattern (FirstElement() might return a hatch pattern)
                solid_pat = None
                patterns = FilteredElementCollector(self.doc).OfClass(FillPatternElement)
                for p in patterns:
                    if p.GetFillPattern().IsSolidFill:
                        solid_pat = p
                        break
                
                for i, sys_item in enumerate(checked_systems):
                    # Cycle colors if multiple systems selected, otherwise use selected
                    if use_cycling:
                        idx = start_idx + i
                        if idx < len(self.color_options):
                            revit_color = self.color_options[idx].RevitColor
                        else:
                            # Generate on the fly if we run out of presets
                            revit_color = self._generate_dynamic_color(idx)
                    else:
                        revit_color = selected_color_opt.RevitColor

                    ogs = OverrideGraphicSettings()
                    if solid_pat:
                        ogs.SetSurfaceForegroundPatternId(solid_pat.Id)
                        ogs.SetSurfaceForegroundPatternColor(revit_color)
                    
                    for eid in sys_item.AllElements:
                        self.doc.ActiveView.SetElementOverrides(eid, ogs)
                t.Commit()
                self.uidoc.RefreshActiveView()
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Visualization Error. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow

    def reset_visuals_click(self, sender, args):
        """Clears graphic overrides for the listed elements."""
        if self.is_busy: return
        
        checked_systems = self.get_checked_systems()
        # Allow reset if we have disabled filters, even if no systems are checked
        if not checked_systems and not self.disabled_filters:
            forms.alert("Please check systems to reset colors.")
            return

        self.is_busy = True
        self.Cursor = Cursors.Wait

        try:
            with Transaction(self.doc, "Reset Visuals") as t:
                t.Start()
                for sys_item in checked_systems:
                    for eid in sys_item.AllElements:
                        self.doc.ActiveView.SetElementOverrides(eid, OverrideGraphicSettings())
                
                # Restore View Filters if we disabled them
                if self.disabled_filters:
                    view = self.doc.ActiveView
                    for fid in self.disabled_filters:
                        if view.IsFilterApplied(fid):
                            view.SetFilterVisibility(fid, True)
                    self.disabled_filters = [] # Clear list after restoring

                t.Commit()
                self.uidoc.RefreshActiveView()
                self.statusLabel.Text = "Visual overrides reset."
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Reset Error. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow

    def rename_click(self, sender, args):
        """Renames the Revit System associated with the selected network."""
        if self.is_busy: return
        
        selected = self.systemTree.SelectedItem
        if not selected or not isinstance(selected, SystemNode):
            forms.alert("Please select a specific System (not Classification) to rename.")
            return
        
        new_name = self.Tb_NewName.Text
        if not new_name:
            return
            
        self.is_busy = True
        self.Cursor = Cursors.Wait

        # Find a system element to rename within the network
        system_elem = None
        for eid in selected.AllElements:
            el = self.doc.GetElement(eid)
            if el and hasattr(el, "MEPSystem") and el.MEPSystem:
                system_elem = el.MEPSystem
                break
        
        if system_elem:
            try:
                with Transaction(self.doc, "Rename System") as t:
                    t.Start()
                    system_elem.Name = new_name
                    t.Commit()
                self.statusLabel.Text = "Renamed system to '{}'".format(new_name)
                selected.Name = new_name
                self.systemTree.Items.Refresh()
            except Exception as e:
                err = traceback.format_exc()
                print(err)
                self.statusLabel.Text = "Rename Error. Check Output."
        else:
            forms.alert("Could not find a valid Revit System to rename in this network.")
        
        self.is_busy = False
        self.Cursor = Cursors.Arrow

    def disconnect_click(self, sender, args):
        """Disconnects selected systems from their base fixtures/equipment."""
        if self.is_busy: return
        
        # 1. Get items to process (Checked OR Selected)
        checked = self.get_checked_systems()
        items_to_process = checked
        
        # Fallback to selected item if nothing checked
        if not items_to_process:
            selected = self.systemTree.SelectedItem
            if selected:
                items_to_process = [selected]

        if not items_to_process:
            forms.alert("Check or select at least 1 item to disconnect.")
            return

        self.is_busy = True
        self.Cursor = Cursors.Wait
        
        try:
            disconnect_count = 0
            with Transaction(self.doc, "Disconnect & Gap Systems") as t:
                t.Start()
                
                terminal_cats = {
                    int(BuiltInCategory.OST_PlumbingFixtures),
                    int(BuiltInCategory.OST_MechanicalEquipment),
                    int(BuiltInCategory.OST_Sprinklers)
                }

                # Collect all unique element IDs to process
                all_eids = set()
                for node in items_to_process:
                    if hasattr(node, "AllElements"):
                        all_eids.update(node.AllElements)
                
                processed_pipes = set()

                for eid in all_eids:
                    el = self.doc.GetElement(eid)
                    if not el: continue
                    
                    cat_id = get_id(el.Category.Id)
                    
                    # Case A: Element is a Pipe
                    if cat_id == int(BuiltInCategory.OST_PipeCurves):
                        pid = get_id(eid)
                        if pid in processed_pipes: continue
                        if self._disconnect_and_gap_pipe(el, terminal_cats):
                            disconnect_count += 1
                        processed_pipes.add(pid)
                        
                    # Case B: Element is a Terminal (Fixture/Equip)
                    elif cat_id in terminal_cats:
                        mgr = None
                        if hasattr(el, "ConnectorManager"): mgr = el.ConnectorManager
                        elif hasattr(el, "MEPModel") and el.MEPModel: mgr = el.MEPModel.ConnectorManager
                        
                        if mgr:
                            for c in mgr.Connectors:
                                if c.IsConnected:
                                    for ref in c.AllRefs:
                                        pipe = ref.Owner
                                        if pipe and get_id(pipe.Category.Id) == int(BuiltInCategory.OST_PipeCurves):
                                            pid = get_id(pipe.Id)
                                            if pid in processed_pipes: continue
                                            
                                            if self._disconnect_and_gap_pipe(pipe, terminal_cats):
                                                disconnect_count += 1
                                            processed_pipes.add(pid)
                t.Commit()
            
            self.statusLabel.Text = "Disconnected {} connections.".format(disconnect_count)
            
            self.uidoc.RefreshActiveView()
            # Unlock busy state so scan can run
            self.is_busy = False
            # Repopulate the list to show changes
            self.scan_view_click(sender, args)
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Disconnect failed. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow

    def _disconnect_and_gap_pipe(self, pipe, terminal_cats):
        """Disconnects pipe from terminals and shortens it to create a gap."""
        disconnected = False
        mgr = pipe.ConnectorManager
        if not mgr: return False
        
        for c in mgr.Connectors:
            if c.IsConnected:
                for ref in c.AllRefs:
                    ref_owner = ref.Owner
                    if ref_owner and ref_owner.Category and get_id(ref_owner.Category.Id) in terminal_cats:
                        try:
                            # 1. Disconnect
                            c.DisconnectFrom(ref)
                            disconnected = True
                            
                            # 2. Create Gap (Shorten Pipe)
                            curve = pipe.Location.Curve
                            if isinstance(curve, Line):
                                p0 = curve.GetEndPoint(0)
                                p1 = curve.GetEndPoint(1)
                                con_origin = c.Origin
                                
                                # Determine which end to move
                                dist0 = p0.DistanceTo(con_origin)
                                dist1 = p1.DistanceTo(con_origin)
                                
                                gap = 0.1 # 0.1 ft gap (~1.2 inches)
                                direction = (p1 - p0).Normalize()
                                
                                new_p0, new_p1 = p0, p1
                                if dist0 < dist1:
                                    new_p0 = p0 + direction * gap
                                else:
                                    new_p1 = p1 - direction * gap
                                
                                # Ensure valid length
                                if new_p0.DistanceTo(new_p1) > 0.01:
                                    new_curve = Line.CreateBound(new_p0, new_p1)
                                    pipe.Location.Curve = new_curve
                        except Exception:
                            pass
        return disconnected

    def _bfs_traversal(self, elements):
        """Standard BFS to group connected elements."""
        el_dict = {get_id(e.Id): e for e in elements}
        unvisited = set(el_dict.keys())
        islands = []

        while unvisited:
            start_id = next(iter(unvisited))
            queue = [start_id]
            unvisited.remove(start_id)
            island = []

            while queue:
                curr_id = queue.pop(0)
                if curr_id in el_dict:
                    elem = el_dict[curr_id]
                    island.append(elem.Id)
                    
                    # Get neighbors via connectors
                    neighbors = self._get_connected_ids(elem)
                    
                    for n_id in neighbors:
                        if n_id in unvisited and n_id in el_dict:
                            unvisited.remove(n_id)
                            queue.append(n_id)
            
            islands.append(island)
        return islands

    def _get_connected_ids(self, element):
        ids = []
        mgr = None
        if hasattr(element, "ConnectorManager"):
            mgr = element.ConnectorManager
        elif hasattr(element, "MEPModel") and element.MEPModel:
            mgr = element.MEPModel.ConnectorManager
        
        if mgr:
            for c in mgr.Connectors:
                if c.IsConnected:
                    for ref in c.AllRefs:
                        if ref.Owner.Id != element.Id:
                            ids.append(get_id(ref.Owner.Id))
        return ids

if __name__ == '__main__':
    # 1. Modal Execution (Blocking)
    # This ensures the script stays on the main thread and prevents context crashes.
    SystemMergeWindow().ShowDialog()