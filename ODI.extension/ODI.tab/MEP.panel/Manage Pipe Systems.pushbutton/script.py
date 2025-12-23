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
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
from System.Collections.Generic import List
from System.Windows.Media import Colors, SolidColorBrush, Color as WpfColor
from System.Windows.Input import Cursors
from Autodesk.Revit.DB import (
    Transaction, BuiltInCategory, ElementId, FilteredElementCollector,
    OverrideGraphicSettings, Color, FillPatternElement, ElementTransformUtils, XYZ,
    BuiltInParameter, ElementMulticategoryFilter
)
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, script, revit

__title__ = "System Merger"
__doc__ = "Modal tool to diagnose and merge disconnected pipe networks."

# Helper for Revit 2024+ compatibility
def get_id(element_id):
    if hasattr(element_id, "Value"):
        return element_id.Value
    return element_id.IntegerValue

# --- Data Model ---
class NodeBase:
    def __init__(self, name):
        self.Name = name
        self.IsChecked = False
        self.IsExpanded = True
        self.Children = []
        self.Type = ""
        self.Abbreviation = ""
        self.Length = ""
        self.Volume = ""
        self.Count = ""
        self.FontWeight = "Normal"
        self.AllElements = [] # Flat list of element IDs for highlighting

class ClassificationNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self, name)
        self.FontWeight = "Bold"

class SystemNode(NodeBase):
    def __init__(self, name, sys_type, sys_abbr, networks):
        NodeBase.__init__(self, name)
        self.Type = sys_type
        self.Abbreviation = sys_abbr
        self.FontWeight = "SemiBold"
        
        # Aggregate
        vol = sum(n.volume for n in networks)
        length = sum(n.length for n in networks)
        count = sum(n.count for n in networks)
        
        self.Volume = "{:.2f}".format(vol)
        self.Length = "{:.2f}".format(length)
        self.Count = "{} Networks".format(len(networks))
        
        for n in networks:
            self.AllElements.extend(n.elements)

class NetworkData:
    """Raw data holder for processing."""
    def __init__(self, volume, length, count, elements):
        self.volume = volume
        self.length = length
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
        self.systemTree.SelectedItemChanged += self.tree_selection_changed
        self.Closing += self.window_closing
        
        self.Btn_Clear.Click += self.clear_list_click
        self.Btn_ScanView.Click += self.scan_view_click
        self.Btn_Visualize.Click += self.visualize_click
        self.Btn_ClearVisuals.Click += self.reset_visuals_click
        self.Btn_Fix.Click += self.fix_network_click
        self.Btn_Rename.Click += self.rename_click
        
        # Initial UI State: Disable actions until data is loaded
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Fix.IsEnabled = False
        self.Btn_Rename.IsEnabled = False
        
        self.load_window_settings()
        self.doc = revit.doc
        self.uidoc = revit.uidoc

        self.last_highlighted_ids = []
        self.is_busy = False
        self.populate_colors()

    # --- UI Logic ---
    def drag_window(self, sender, args):
        self.DragMove()

    def close_window(self, sender, args):
        self.Close()

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
    def clear_list_click(self, sender, args):
        """Clears the list and resets selection overrides."""
        if self.is_busy: return
        
        self.systemTree.ItemsSource = None
        self.reset_selection_highlight()
        self.statusLabel.Text = "List cleared."
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Fix.IsEnabled = False
        self.Btn_Rename.IsEnabled = False
        self.Tb_NewName.Text = ""

    def scan_view_click(self, sender, args):
        """Scans all pipe elements in the active view."""
        if self.is_busy: return
        self.is_busy = True
        self.Cursor = Cursors.Wait
        
        try:
            valid_cats = [
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory
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
        self.Btn_Visualize.IsEnabled = False
        self.Btn_ClearVisuals.IsEnabled = False
        self.Btn_Fix.IsEnabled = False
        self.Btn_Rename.IsEnabled = False
        self.Tb_NewName.Text = ""

        try:
            if not element_ids:
                self.statusLabel.Text = "No elements to analyze."
                return
                
            # Filter for pipes, fittings, and accessories to ensure connectivity
            valid_cats = [
                int(BuiltInCategory.OST_PipeCurves),
                int(BuiltInCategory.OST_PipeFitting),
                int(BuiltInCategory.OST_PipeAccessory)
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

            # BFS to find islands
            islands = self._bfs_traversal(elements_to_process)
            
            # Process Islands and Group by System
            system_map = {} # Name -> List[NetworkData]
            system_types = {} # Name -> Type
            system_classes = {} # Name -> Classification
            system_abbrs = {} # Name -> Abbreviation

            for island_ids in islands:
                # Calculate Volume (Vent Logic)
                total_vol = 0.0
                total_len = 0.0
                sys_name = None
                sys_type = None
                sys_class = None
                sys_abbr = None

                for eid in island_ids:
                    el = self.doc.GetElement(eid)
                    # Volume
                    l_param = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
                    d_param = el.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM)
                    if l_param and d_param:
                        r = d_param.AsDouble() / 2.0
                        total_vol += math.pi * (r**2) * l_param.AsDouble()
                    
                    # Length (for pipes)
                    if l_param:
                        total_len += l_param.AsDouble()
                    
                    # System Info - Robust Search
                    # We keep looking until we find valid non-empty strings
                    if sys_name is None:
                        if hasattr(el, "MEPSystem") and el.MEPSystem and el.MEPSystem.Name:
                            sys_name = el.MEPSystem.Name
                        else:
                            p = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
                            if p and p.AsString(): sys_name = p.AsString()
                    
                    if sys_type is None:
                        # Try via MEPSystem Type first
                        if hasattr(el, "MEPSystem") and el.MEPSystem:
                            type_id = el.MEPSystem.GetTypeId()
                            if type_id != ElementId.InvalidElementId:
                                t_elem = self.doc.GetElement(type_id)
                                if t_elem:
                                    # Fix for AttributeError: Name
                                    p_name = t_elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                                    if p_name: sys_type = p_name.AsString()
                                    if not sys_type: sys_type = getattr(t_elem, "Name", None)
                        
                        # Fallback to parameter
                        if sys_type is None:
                            p = el.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
                            if p and p.AsValueString(): sys_type = p.AsValueString()

                    if sys_class is None:
                        p = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
                        if p and p.AsString(): sys_class = p.AsString()
                    
                    if sys_abbr is None:
                        p = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)
                        if p and p.AsString(): sys_abbr = p.AsString()
                
                # Defaults if still missing
                if sys_name is None: sys_name = "Unassigned"
                if sys_type is None: sys_type = "Unassigned"
                if sys_class is None: sys_class = "Unassigned"
                if sys_abbr is None: sys_abbr = ""
                
                if sys_name not in system_map:
                    system_map[sys_name] = []
                    system_types[sys_name] = sys_type
                    system_classes[sys_name] = sys_class
                    system_abbrs[sys_name] = sys_abbr
                
                system_map[sys_name].append(NetworkData(round(total_vol, 3), round(total_len, 2), len(island_ids), island_ids))

            # Build Tree Structure: Classification -> System
            # Group systems by classification
            class_map = {} # ClassName -> List[SystemNode]
            
            for name, networks in system_map.items():
                cls = system_classes[name]
                if cls not in class_map:
                    class_map[cls] = []
                class_map[cls].append(SystemNode(name, system_types[name], system_abbrs[name], networks))
            
            # Create Root Nodes
            root_nodes = []
            for cls_name, sys_nodes in class_map.items():
                c_node = ClassificationNode(cls_name)
                c_node.Children = sorted(sys_nodes, key=lambda x: x.Name)
                root_nodes.append(c_node)

            self.systemTree.ItemsSource = sorted(root_nodes, key=lambda x: x.Name)
            self.statusLabel.Text = "Found {} systems.".format(len(system_map))
            
            # Enable buttons if data exists
            has_data = len(system_map) > 0
            self.Btn_Visualize.IsEnabled = has_data
            self.Btn_ClearVisuals.IsEnabled = has_data
            self.Btn_Fix.IsEnabled = has_data
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
        node = sender.DataContext
        if node:
            node.IsChecked = sender.IsChecked
            
            # Cascade to children (Classification -> Systems)
            if hasattr(node, "Children"):
                for child in node.Children:
                    child.IsChecked = node.IsChecked
            
            # Cascade Up (Child -> Parent)
            if self.systemTree.ItemsSource:
                for root in self.systemTree.ItemsSource:
                    if hasattr(root, "Children") and node in root.Children:
                        # Update parent based on all children
                        root.IsChecked = all(c.IsChecked for c in root.Children)
                        break
            
            # Refresh tree to update UI for children
            self.systemTree.Items.Refresh()

    def tree_selection_changed(self, sender, args):
        """Syncs TreeView selection with Revit selection (Highlight)."""
        self.reset_selection_highlight()
        
        try:
            selected_node = self.systemTree.SelectedItem
            if not selected_node: return
            
            # Populate Rename Box
            if isinstance(selected_node, SystemNode):
                self.Tb_NewName.Text = selected_node.Name
            else:
                self.Tb_NewName.Text = ""
            
            ids = set(selected_node.AllElements)
            # If it's a classification node, gather all children
            if isinstance(selected_node, ClassificationNode):
                ids = {eid for child in selected_node.Children for eid in child.AllElements}
            
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
                    ogs = OverrideGraphicSettings()
                    ogs.SetProjectionLineColor(Color(255, 165, 0)) # Orange
                    ogs.SetProjectionLineWeight(12) # Extra Thick / Glow Effect
                    for eid in ids:
                        self.doc.ActiveView.SetElementOverrides(eid, ogs)
                    
                    # 2. Dim Background (Halftone + Transparent)
                    if background_ids:
                        ogs_dim = OverrideGraphicSettings()
                        ogs_dim.SetHalftone(True)
                        ogs_dim.SetSurfaceTransparency(60) # 60% Transparent
                        for eid in background_ids:
                            self.doc.ActiveView.SetElementOverrides(eid, ogs_dim)
                            
                    t.Commit()
                
                self.last_highlighted_ids = list(ids.union(background_ids))
                self.uidoc.RefreshActiveView()
                
                # Prepare list for Zoom/Select
                elem_ids = List[ElementId](ids)

                # Auto-Zoom if enabled
                if self.Cb_AutoZoom.IsChecked:
                    self.uidoc.ShowElements(elem_ids)
                
                # Auto-Select if enabled
                if self.Cb_AutoSelect.IsChecked:
                    self.uidoc.Selection.SetElementIds(elem_ids)
                    
        except Exception:
            pass # Prevent crash if selection fails

    def get_checked_systems(self):
        """Helper to find all checked SystemNodes."""
        checked = []
        if self.systemTree.ItemsSource:
            for class_node in self.systemTree.ItemsSource:
                for sys_node in class_node.Children:
                    if sys_node.IsChecked:
                        checked.append(sys_node)
        return checked

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
            with Transaction(self.doc, "Visualize Networks") as t:
                t.Start()
                
                solid_pat = FilteredElementCollector(self.doc).OfClass(FillPatternElement).FirstElement() 
                
                for i, sys_item in enumerate(checked_systems):
                    # Cycle colors if multiple systems selected, otherwise use selected
                    if use_cycling:
                        c_opt = self.color_options[(start_idx + i) % len(self.color_options)]
                        revit_color = c_opt.RevitColor
                    else:
                        revit_color = selected_color_opt.RevitColor

                    ogs = OverrideGraphicSettings()
                    ogs.SetProjectionLineColor(revit_color)
                    if solid_pat:
                        ogs.SetSurfaceForegroundPatternId(solid_pat.Id)
                        ogs.SetSurfaceForegroundPatternColor(revit_color)
                    
                    for eid in sys_item.AllElements:
                        self.doc.ActiveView.SetElementOverrides(eid, ogs)
                t.Commit()
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
        if not checked_systems:
            forms.alert("Please check systems to reset.")
            return

        self.is_busy = True
        self.Cursor = Cursors.Wait

        try:
            with Transaction(self.doc, "Reset Visuals") as t:
                t.Start()
                for sys_item in checked_systems:
                    for eid in sys_item.AllElements:
                        self.doc.ActiveView.SetElementOverrides(eid, OverrideGraphicSettings())
                t.Commit()
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

    def fix_network_click(self, sender, args):
        """Fixes selected networks by moving the 'Loser' slightly to trigger Revit auto-connect."""
        if self.is_busy: return
        
        checked = self.get_checked_systems()
        if len(checked) < 2:
            forms.alert("Check at least 2 systems to fix/connect.")
            return

        self.is_busy = True
        self.Cursor = Cursors.Wait

        # Determine Winner (Logic: Has Equipment > Largest Volume)
        # Sort: Equipment (True=1, False=0) Desc, Volume Desc
        checked.sort(key=lambda x: float(x.Volume), reverse=True)
        
        winner = checked[0]
        losers = checked[1:]

        self.Hide()
        try:
            with Transaction(self.doc, "Fix Network Connections") as t:
                t.Start()
                for loser in losers:
                    # The Fix: Jiggle the first UNPINNED element of the loser network
                    # This forces Revit to re-evaluate connectivity if they are geometrically close
                    jiggle_id = None
                    for eid in loser.AllElements:
                        el = self.doc.GetElement(eid)
                        if el and not el.Pinned:
                            jiggle_id = eid
                            break
                    
                    if jiggle_id:
                        try:
                            ElementTransformUtils.MoveElement(self.doc, jiggle_id, XYZ(0.1, 0, 0))
                            ElementTransformUtils.MoveElement(self.doc, jiggle_id, XYZ(-0.1, 0, 0))
                        except Exception:
                            pass # Skip if movement fails (e.g. constraints)
                t.Commit()
            
            self.statusLabel.Text = "Fixed {} networks into {}.".format(len(losers), winner.Name)
            
            # Repopulate the list to show changes
            self.scan_view_click(sender, args)
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Fix failed. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow
            self.Show()

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
                island.append(ElementId(curr_id))
                
                curr_el = self.doc.GetElement(ElementId(curr_id))
                if not curr_el: continue

                # Get neighbors via connectors
                neighbors = self._get_connected_ids(curr_el)
                
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