# -*- coding: utf-8 -*-
"""
Quantity & Measures

Description:
    A Modal WPF tool to scan visible elements in the active view and aggregate 
    measurable quantities (Area, Volume, Length, etc.) by Category and Type.
"""

import os
import traceback
import math
import csv
import datetime
import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
from System.Collections.Generic import List
from System.Windows.Media import SolidColorBrush, Color as WpfColor
from System.Windows.Input import Cursors, Key
from Autodesk.Revit.DB import (
    Transaction, BuiltInCategory, ElementId, FilteredElementCollector,
    OverrideGraphicSettings, Color, FillPatternElement, ElementTransformUtils, XYZ,
    BuiltInParameter, ElementMulticategoryFilter, Line, StorageType, TemporaryViewMode
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, script, revit

# Import modularized components
from revit_utils import get_id, get_display_val_and_label, is_dark_theme, MEASURABLE_NAMES
from data_model import NodeBase, MeasurementNode, CategoryNode, FamilyTypeNode, InstanceItem, ColorOption
from calculators_walls import WallCMUCalculator
from settings_logic import SettingsWindow

__title__ = "Quantity & Measures"
__version__ = "0.3"
__doc__ = """A Modal WPF tool to visualize, quantify, and estimate materials for visible elements.
Features:
- Quantify: Aggregate Length, Area, Volume by Category/Type.
- Calculate: Estimate material requirements (e.g., CMU/Brick count) based on configurable settings.
- Visualize: Colorize or Isolate elements for visual verification.
- Export: Export quantified data to CSV."""
__context__ = "active-view-type: FloorPlan,CeilingPlan,EngineeringPlan,AreaPlan,Section,Elevation,ThreeD"

# --- Main Window Class ---
class SystemNetworkWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'UI.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        # UI Event Bindings for Custom Title Bar
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_WinClose.Click += self.close_window
        self.Closing += self.window_closing
        
        self.Btn_SelectAll.Click += self.select_all_click
        self.Btn_Clear.Click += self.clear_list_click
        self.Btn_Export.Click += self.export_click
        self.Btn_ExpandAll.Click += self.expand_all_click
        self.Btn_CollapseAll.Click += self.collapse_all_click
        self.Btn_ScanView.Click += self.scan_view_click
        self.Btn_Visualize.Click += self.visualize_click
        self.Btn_ClearVisuals.Click += self.reset_visuals_click
        self.sysDataGrid.SelectionChanged += self.grid_selection_changed
        self.Btn_Settings.Click += self.settings_click
        self.Btn_Isolate.Click += self.isolate_click
        
        # Handle TreeView Selection via ItemContainerStyle Binding
        # We no longer use SelectedItemChanged, but we can listen to property changes if needed.
        # However, for the logic, we can just iterate or bind commands. 
        # For simplicity in this hybrid approach, we will hook into the TreeView's SelectedItemChanged 
        # just to trigger the visualization logic, but rely on the ViewModel for state.
        self.systemTree.SelectedItemChanged += self.tree_selection_changed
        
        # Initial UI State: Disable actions until data is loaded
        self.Btn_SelectAll.IsEnabled = False
        self.Btn_ExpandAll.IsEnabled = False
        self.Btn_CollapseAll.IsEnabled = False
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Export.IsEnabled = False
        self.Btn_Isolate.IsEnabled = False
        
        # Default Header
        self.set_default_header()
        
        self.load_window_settings()
        self.doc = revit.doc
        self.uidoc = revit.uidoc

        self.last_highlighted_ids = []
        self.last_grid_selected_ids = []
        self.is_busy = False
        self.disabled_filters = [] # Track filters we disable to restore them later
        self.populate_colors()

        # Initialize Calculators
        self.wall_calculator = WallCMUCalculator()
        self.calc_settings = {"Walls": self.wall_calculator.default_setting}

        self.apply_revit_theme()

        # Check for pre-selection
        sel_ids = self.uidoc.Selection.GetElementIds()
        if sel_ids.Count > 0:
            self.analyze_selection(sel_ids)

    def set_default_header(self):
        default_node = NodeBase("Quantity & Measures")
        default_node.Type = "Scan view or select elements to begin."
        self.RightPane.DataContext = default_node

    # --- UI Logic ---
    def drag_window(self, sender, args):
        self.DragMove()

    def close_window(self, sender, args):
        self.Close()

    def apply_revit_theme(self):
        """Detects Revit theme and updates window resources if Dark."""
        if is_dark_theme():
            # Define Modern Dark Theme Colors (Slate/Blue Palette)
            res = self.Resources
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(31, 41, 55))      # #1F2937 (Gray-800)
            res["ToolbarBrush"] = SolidColorBrush(WpfColor.FromRgb(31, 41, 55))     # #1F2937 (Gray-800)
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(17, 24, 39))     # #111827 (Gray-900)
            res["ButtonBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81))      # #374151 (Gray-700)
            res["FooterBrush"] = SolidColorBrush(WpfColor.FromRgb(17, 24, 39))      # #111827 (Gray-900)
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(249, 250, 251))     # #F9FAFB (Gray-50)
            res["TextLightBrush"] = SolidColorBrush(WpfColor.FromRgb(156, 163, 175))# #9CA3AF (Gray-400)
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 99))      # #4B5563 (Gray-600)
            res["AccentBrush"] = SolidColorBrush(WpfColor.FromRgb(59, 130, 246))    # #3B82F6 (Blue-500)
            res["SelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(30, 58, 138))  # #1E3A8A (Blue-900)
            res["SelectionBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(59, 130, 246)) # Blue-500
            res["HoverBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81))       # #374151 (Gray-700)
            res["AltRowBrush"] = SolidColorBrush(WpfColor.FromRgb(31, 41, 55))      # #1F2937 (Gray-800)
            
            # Dashboard Specifics (Dark Card on Dark Background)
            res["CardBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81))        # #374151 (Gray-700 - Elevated)
            res["CardBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 99))  # #4B5563
            res["CardTextBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255)) # White
            res["CardSubTextBrush"] = SolidColorBrush(WpfColor.FromRgb(209, 213, 219)) # #D1D5DB (Gray-300)
            res["CardLabelBrush"] = SolidColorBrush(WpfColor.FromRgb(156, 163, 175))   # #9CA3AF (Gray-400)
            res["CardValueBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))   # White
            res["CardAccentBrush"] = SolidColorBrush(WpfColor.FromRgb(96, 165, 250))   # #60A5FA (Blue-400)

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
        self.Btn_Export.IsEnabled = has_checked or has_selection
        self.Btn_Isolate.Content = "Isolate"

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
        self.Btn_ExpandAll.IsEnabled = False
        self.Btn_CollapseAll.IsEnabled = False
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Export.IsEnabled = False
        self.Btn_Isolate.IsEnabled = False
        self.Btn_Isolate.Content = "Isolate"
        self.set_default_header()

    def scan_view_click(self, sender, args):
        """Scans all pipe elements in the active view."""
        if self.is_busy: return
        self.is_busy = True
        self.Cursor = Cursors.Wait
        
        try:
            self.uidoc.Selection.SetElementIds(List[ElementId]())
            self.analyze_view()
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Error scanning view. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow
            
    def analyze_view(self):
        """Analyzes all visible elements in the active view."""
        collector = FilteredElementCollector(self.doc, self.doc.ActiveView.Id).WhereElementIsNotElementType()
        elements = collector.ToElements()
        self.process_elements(elements)

    def analyze_selection(self, element_ids):
        """Analyzes a specific set of elements."""
        elements = []
        for eid in element_ids:
            el = self.doc.GetElement(eid)
            if el: elements.append(el)
        self.process_elements(elements)

    def process_elements(self, elements):
        """Core logic to aggregate quantities from a list of elements."""
        # Reset UI State
        self.systemTree.ItemsSource = None
        self.Btn_SelectAll.Content = "Select All"
        self.Btn_SelectAll.IsEnabled = False
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Export.IsEnabled = False
        self.Btn_Isolate.IsEnabled = False
        self.Btn_Isolate.Content = "Isolate"
        self.set_default_header()

        try:
            # Structure: { ParamName: { CategoryName: { TypeName: [InstanceItem] } } }
            tree_data = {}
            
            for el in elements:
                if not el.Category: continue
                
                # Iterate Parameters
                for p in el.Parameters:
                    if p.StorageType == StorageType.Double:
                        p_name = p.Definition.Name
                        
                        # Filter by Whitelist (Case Insensitive)
                        # We check if any whitelisted name is contained in the param name
                        # e.g. "Area" matches "Host Area", "Area", "Paint Area"
                        # But user wants specific top level items.
                        # Let's match exact names or very close ones to avoid noise.
                        # Actually, let's just check if the name is in our set.
                        
                        # Clean name logic?
                        # Let's just use the parameter name as the key.
                        # But filter:
                        is_measurable = False
                        for m_name in MEASURABLE_NAMES:
                            if m_name.lower() == p_name.lower():
                                is_measurable = True
                                break
                        
                        if not is_measurable: continue
                        
                        val, unit_label = get_display_val_and_label(p, self.doc)
                        if abs(val) < 0.0001: continue # Skip zero values
                        
                        cat_name = el.Category.Name
                        type_name = el.Name

                        # Run Calculation if applicable
                        calc_val = "-"
                        if cat_name == "Walls":
                            calc_val = self.wall_calculator.calculate(el, self.calc_settings["Walls"])
                        
                        if p_name not in tree_data: tree_data[p_name] = {}
                        if cat_name not in tree_data[p_name]: tree_data[p_name][cat_name] = {}
                        if type_name not in tree_data[p_name][cat_name]: tree_data[p_name][cat_name][type_name] = []
                        
                        tree_data[p_name][cat_name][type_name].append(InstanceItem(el, val, unit_label, calc_val))

            # 2. Build Tree Nodes
            root_nodes = []
            
            for p_name, cat_dict in sorted(tree_data.items()):
                m_node = MeasurementNode(p_name)
                total_val = 0.0
                total_count = 0
                
                for cat_name, type_dict in sorted(cat_dict.items()):
                    c_node = CategoryNode(cat_name)
                    c_val = 0.0
                    c_count = 0
                    
                    for type_name, instances in sorted(type_dict.items()):
                        t_node = FamilyTypeNode(type_name)
                        t_val = sum(i.Value for i in instances)
                        t_count = len(instances)
                        
                        t_node.Value = t_val
                        t_node.Count = t_count
                        t_node.AllElements = [i.Id for i in instances]
                        # Store instances for DataGrid
                        t_node.Instances = instances 
                        if instances:
                            t_node.UnitLabel = instances[0].UnitLabel
                        
                        c_node.Children.append(t_node)
                        c_val += t_val
                        c_count += t_count
                        c_node.AllElements.extend(t_node.AllElements)
                    
                    c_node.Value = c_val
                    c_node.Count = c_count
                    if c_node.Children:
                        c_node.UnitLabel = c_node.Children[0].UnitLabel
                    m_node.Children.append(c_node)
                    total_val += c_val
                    total_count += c_count
                    m_node.AllElements.extend(c_node.AllElements)
                
                m_node.Value = total_val
                m_node.Count = total_count
                if m_node.Children:
                    m_node.UnitLabel = m_node.Children[0].UnitLabel
                root_nodes.append(m_node)

            self.systemTree.ItemsSource = root_nodes
            self.statusLabel.Text = "Found {} measurable parameters.".format(len(root_nodes))
            
            # Enable buttons if data exists
            has_data = len(root_nodes) > 0
            self.Btn_SelectAll.IsEnabled = has_data
            self.Btn_ExpandAll.IsEnabled = has_data
            self.Btn_CollapseAll.IsEnabled = has_data
            self.update_button_states()
            self.Btn_Isolate.IsEnabled = has_data
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
        node = sender.DataContext
        if node:
            # Use recursive cascade check to ensure all children (including leaf nodes) are updated
            self._cascade_check(node, node.IsChecked)
            self.update_button_states()

    def tree_selection_changed(self, sender, args):
        """Syncs TreeView selection with Revit selection (Highlight)."""
        if self.is_busy: return
        
        self.reset_selection_highlight()
        self.last_grid_selected_ids = [] # Reset grid selection tracking on tree change
        self.update_button_states()
        
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node:
                selected_node.IsSelected = True # Ensure ViewModel is in sync

            if not selected_node: return
            
            # Update Header Context immediately to ensure UI updates even if highlighting fails
            self.RightPane.DataContext = selected_node
            
            ids = set(selected_node.AllElements)
            if isinstance(selected_node, (MeasurementNode, CategoryNode)):
                # Recursively gather all elements
                ids = self._get_all_child_elements(selected_node)
            
            if ids:
                # Identify background elements (Rest of the Model)
                # We collect all elements in the active view to apply the dimming effect.
                view_id = self.doc.ActiveView.Id
                collector = FilteredElementCollector(self.doc, view_id).WhereElementIsNotElementType()
                
                # Convert collector IDs to integers/longs for set operations
                all_view_ids = set(get_id(e.Id) for e in collector)
                background_ids = all_view_ids - ids

                # Prepare ElementId lists for Revit API calls
                ids_elem = []
                for i in ids:
                    try: ids_elem.append(ElementId(i))
                    except: pass
                    
                bg_elem = []
                for i in background_ids:
                    try: bg_elem.append(ElementId(i))
                    except: pass

                # Apply Bold Orange Highlight & Dim Background
                with Transaction(self.doc, "Highlight Selection") as t:
                    t.Start()
                    
                    # 1. Highlight Selected
                    ogs_sel = OverrideGraphicSettings()
                    ogs_sel.SetProjectionLineColor(Color(0, 128, 255)) # System Browser Blue
                    ogs_sel.SetProjectionLineWeight(12) # Extra Thick / Glow Effect
                    for eid in ids_elem:
                        self.doc.ActiveView.SetElementOverrides(eid, ogs_sel)
                    
                    # 2. Dim Background (Halftone + Transparent)
                    if bg_elem:
                        ogs_dim = OverrideGraphicSettings()
                        ogs_dim.SetHalftone(True)
                        ogs_dim.SetSurfaceTransparency(80) # 80% Transparent
                        for eid in bg_elem:
                            self.doc.ActiveView.SetElementOverrides(eid, ogs_dim)
                            
                    t.Commit()
                
                # Store ElementIds for reset
                self.last_highlighted_ids = ids_elem + bg_elem
                self.uidoc.RefreshActiveView()
                
                elem_ids = List[ElementId](ids_elem)
                
                # 1. Select in Revit
                if elem_ids:
                    self.uidoc.Selection.SetElementIds(elem_ids)
                
                # 2. Auto-Zoom if enabled
                if self.Cb_AutoZoom.IsChecked and elem_ids:
                    self.uidoc.ShowElements(elem_ids)
            
        except Exception:
            pass # Prevent crash if selection fails

    def grid_selection_changed(self, sender, args):
        """Handles selection in the DataGrid (Instances) for Zoom and Highlight."""
        if self.is_busy: return
        
        try:
            selected_items = self.sysDataGrid.SelectedItems
            current_ids = []
            if selected_items:
                for item in selected_items:
                    if hasattr(item, "Id") and item.Id:
                        current_ids.append(item.Id)
            
            # Determine changes (Integers)
            cur_set = set(current_ids)
            last_set = set(self.last_grid_selected_ids)
            
            to_orange = cur_set - last_set # Newly selected -> Orange
            to_blue = last_set - cur_set   # Deselected -> Revert to Blue
            
            if not to_orange and not to_blue:
                return

            with Transaction(self.doc, "Highlight Instance") as t:
                t.Start()
                
                # Revert to Blue (Type Highlight)
                if to_blue:
                    ogs_blue = OverrideGraphicSettings()
                    ogs_blue.SetProjectionLineColor(Color(0, 128, 255))
                    ogs_blue.SetProjectionLineWeight(12)
                    for i in to_blue:
                        try: self.doc.ActiveView.SetElementOverrides(ElementId(i), ogs_blue)
                        except: pass

                # Apply Orange (Instance Highlight)
                if to_orange:
                    ogs_orange = OverrideGraphicSettings()
                    ogs_orange.SetProjectionLineColor(Color(255, 128, 0))
                    ogs_orange.SetProjectionLineWeight(14)
                    for i in to_orange:
                        try: self.doc.ActiveView.SetElementOverrides(ElementId(i), ogs_orange)
                        except: pass
                
                t.Commit()
                self.uidoc.RefreshActiveView()
            
            self.last_grid_selected_ids = list(cur_set)
            
            # Auto-Zoom & Select in Revit
            if self.Cb_AutoZoom.IsChecked and current_ids:
                elem_ids = List[ElementId]([ElementId(i) for i in current_ids])
                self.uidoc.Selection.SetElementIds(elem_ids)
                self.uidoc.ShowElements(elem_ids)

        except Exception:
            pass

    def _get_all_child_elements(self, node):
        """Recursively gets all element IDs from a node and its children."""
        ids = set(node.AllElements)
        if hasattr(node, "Children"):
            for child in node.Children:
                ids.update(self._get_all_child_elements(child))
        return ids

    def get_checked_systems(self):
        """Helper to find all checked Nodes."""
        checked = []
        if self.systemTree.ItemsSource:
            for m_node in self.systemTree.ItemsSource:
                for c_node in m_node.Children:
                    for t_node in c_node.Children:
                        if t_node.IsChecked:
                            checked.append(t_node)
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

    def export_click(self, sender, args):
        """Exports the current tree data to a CSV file."""
        if not self.systemTree.ItemsSource:
            return

        # Check if any items are checked or selected
        checked_items = self.get_checked_systems()
        has_checked = len(checked_items) > 0
        selected_node = self.systemTree.SelectedItem
        
        if not has_checked and not selected_node:
            forms.alert("Please check or select items to export.")
            return

        # Generate Filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = "Quantities"
        
        if has_checked:
            # Analyze checked items to determine common ancestry
            measurements = set()
            categories = set()
            types = set()
            checked_set = set(checked_items)
            
            for m_node in self.systemTree.ItemsSource:
                for c_node in m_node.Children:
                    for t_node in c_node.Children:
                        if t_node in checked_set:
                            measurements.add(m_node.Name)
                            categories.add(c_node.Name)
                            types.add(t_node.Name)
            
            if len(measurements) == 1:
                m_name = list(measurements)[0]
                if len(categories) == 1:
                    c_name = list(categories)[0]
                    if len(types) == 1:
                        t_name = list(types)[0]
                        prefix = "{}-{}-{}".format(m_name, c_name, t_name)
                    else:
                        prefix = "{}-{}".format(m_name, c_name)
                else:
                    prefix = m_name
            else:
                prefix = "Selected_Quantities"
        elif selected_node:
            # Construct hierarchical name (e.g. Width-Doors)
            path_names = []
            found = False
            for m in self.systemTree.ItemsSource:
                if m == selected_node:
                    path_names = [m.Name]
                    found = True
                    break
                for c in m.Children:
                    if c == selected_node:
                        path_names = [m.Name, c.Name]
                        found = True
                        break
                    for t in c.Children:
                        if t == selected_node:
                            path_names = [m.Name, c.Name, t.Name]
                            found = True
                            break
                    if found: break
                if found: break
            
            if path_names:
                prefix = "-".join(path_names)
            else:
                prefix = selected_node.Name
            
        # Sanitize
        safe_prefix = "".join([c for c in prefix if c.isalnum() or c in (' ', '_', '-')]).strip()
        default_name = "{}_{}.csv".format(safe_prefix, timestamp)

        # Prompt for file save
        dest_file = forms.save_file(file_ext='csv', default_name=default_name)
        if not dest_file:
            return

        try:
            with open(dest_file, 'wb') as f:
                writer = csv.writer(f)
                # Header
                writer.writerow(['Measurement', 'Category', 'Type', 'Count', 'Value', 'Unit'])
                
                # Iterate Tree (Preserving Hierarchy Sequence)
                for m_node in self.systemTree.ItemsSource:
                    measurement_name = m_node.Name.encode('utf-8')
                    # Check if parent is selected (implies all children exported if nothing checked)
                    m_selected = (m_node == selected_node) and not has_checked
                    
                    for c_node in m_node.Children:
                        category_name = c_node.Name.encode('utf-8')
                        c_selected = (c_node == selected_node) and not has_checked
                        
                        for t_node in c_node.Children:
                            # Export if:
                            # 1. It is explicitly checked
                            # 2. Nothing is checked AND (Parent is selected OR Itself is selected)
                            t_selected = (t_node == selected_node) and not has_checked
                            
                            if t_node.IsChecked or m_selected or c_selected or t_selected:
                                type_name = t_node.Name.encode('utf-8')
                                count = str(t_node.Count)
                                val = "{:.2f}".format(t_node.Value)
                                unit = t_node.UnitLabel.encode('utf-8')
                                
                                writer.writerow([measurement_name, category_name, type_name, count, val, unit])
            
            self.statusLabel.Text = "Export successful: {}".format(os.path.basename(dest_file))
            os.startfile(dest_file)
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Export failed. Check Output."

    def settings_click(self, sender, args):
        """Opens a dialog to configure calculation settings."""
        win = SettingsWindow()
        win.ShowDialog()
        
        # After window closes, we can trigger a recalculation if needed.
        # Currently, the calculators need to be updated to read from the new JSON.
        # For now, we just notify the user.
        self.statusLabel.Text = "Settings saved."

    def recalculate_all(self):
        """Iterates through existing tree and updates calculated values."""
        if not self.systemTree.ItemsSource: return
        
        for m_node in self.systemTree.ItemsSource:
            for c_node in m_node.Children:
                if c_node.Name == "Walls":
                    for t_node in c_node.Children:
                        for instance in t_node.Instances:
                            # Recalculate
                            new_val = self.wall_calculator.calculate(instance.Element, self.calc_settings["Walls"])
                            instance.CalculatedValue = new_val

    def isolate_click(self, sender, args):
        """Isolates checked, selected, or all listed elements in the active view."""
        if self.is_busy: return

        # Toggle Logic
        if self.Btn_Isolate.Content == "Unisolate":
            try:
                with Transaction(self.doc, "Reset Isolate") as t:
                    t.Start()
                    self.doc.ActiveView.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
                    t.Commit()
                self.Btn_Isolate.Content = "Isolate"
            except Exception as e:
                print("Unisolate Error: {}".format(e))
            return
        
        ids_to_isolate = set()
        
        # 1. Checked Items
        checked = self.get_checked_systems()
        if checked:
            for node in checked:
                ids_to_isolate.update(node.AllElements)
        
        # 2. Selected Item (if nothing checked)
        elif self.systemTree.SelectedItem:
            node = self.systemTree.SelectedItem
            if hasattr(node, "AllElements"):
                ids_to_isolate.update(node.AllElements)
                
        # 3. All Items (if nothing checked or selected)
        elif self.systemTree.ItemsSource:
            for m_node in self.systemTree.ItemsSource:
                ids_to_isolate.update(m_node.AllElements)
                
        if not ids_to_isolate:
            forms.alert("No elements found to isolate.")
            return
            
        try:
            elem_ids = List[ElementId]()
            for i in ids_to_isolate:
                try: elem_ids.Add(ElementId(i))
                except: pass

            with Transaction(self.doc, "Isolate Elements") as t:
                t.Start()
                self.doc.ActiveView.IsolateElementsTemporary(elem_ids)
                t.Commit()
            self.Btn_Isolate.Content = "Unisolate"
        except Exception as e:
            print("Isolate Error: {}".format(e))

if __name__ == '__main__':
    SystemNetworkWindow().ShowDialog()
