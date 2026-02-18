# -*- coding: utf-8 -*-
__title__ = "Quantity & Measures"
__version__ = "1.0"
__doc__ = """A Modal WPF tool to visualize, quantify, and estimate materials for visible elements.
Features:
- Quantify: Aggregate Length, Area, Volume by Category/Type.
- Calculate: Estimate material requirements (e.g., CMU/Brick count) based on configurable settings.
- Visualize: Colorize or Isolate elements for visual verification.
- Export: Export quantified data to CSV."""
__author__ = "ODI"
__context__ = "doc-project"

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
from System.Collections.Generic import List, HashSet
from System.Windows.Media import SolidColorBrush, Color as WpfColor, Colors
from System.Windows.Input import Cursors, Key, Keyboard, ModifierKeys
from System.Windows import Clipboard
from System.Windows.Controls import ContextMenu, MenuItem, ScrollViewer, ScrollBarVisibility
from System.Windows.Markup import XamlReader
from Autodesk.Revit.DB import (
    Transaction, BuiltInCategory, ElementId, FilteredElementCollector,
    OverrideGraphicSettings, Color, FillPatternElement, ElementTransformUtils, XYZ,
    BuiltInParameter, ElementMulticategoryFilter, Line, StorageType, TemporaryViewMode, ParameterFilterElement
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, script, revit

# Import modularized components
from revit_utils import get_id, get_display_val_and_label, is_dark_theme, MEASURABLE_NAMES
from data_model import NodeBase, MeasurementNode, CategoryNode, FamilyTypeNode, InstanceItem, ColorOption, ViewModelBase
from calculators_walls import WallCMUCalculator
from settings_logic import SettingsWindow

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
        self.Btn_ScanView.ToolTip = "Click to Scan Active View."
        self.systemTree.MouseDoubleClick += self.tree_double_click
        self.sysDataGrid.MouseDoubleClick += self.grid_double_click
        self.setup_context_menu()
        self.setup_grid_context_menu()
        
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
        self.element_color_map = {} # Cache for persistent colors: {int_id: RevitColor}
        self.instance_item_map = {} # {element_id: [InstanceItem, ...]}
        self.solid_pattern_id = None
        self.is_busy = False
        self.disabled_filters = [] # Track filters we disable to restore them later
        self.populate_colors()

        # Initialize Calculators
        self.wall_calculator = WallCMUCalculator()
        self.calc_settings = {"Walls": self.wall_calculator.default_setting}

        self.apply_revit_theme()
        
        # Inject DataTemplate for Color ComboBox to show preview
        try:
            cmb_template = """
            <DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
                <StackPanel Orientation="Horizontal">
                    <Border Width="12" Height="12" Background="{Binding Brush}" BorderBrush="Gray" BorderThickness="1" Margin="0,0,8,0" VerticalAlignment="Center"/>
                    <TextBlock Text="{Binding Name}" VerticalAlignment="Center"/>
                </StackPanel>
            </DataTemplate>"""
            self.Cmb_Colors.ItemTemplate = XamlReader.Parse(cmb_template)
        except: pass
        
        # Disable Horizontal Scroll to force TreeViewItems to stretch to viewport width
        self.systemTree.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Disabled)
        
        # Inject DataTemplate for TreeView to show Color Circle
        try:
            tree_template = """
            <HierarchicalDataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
                                      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                                      ItemsSource="{Binding Children}">
                <Border Background="{DynamicResource CardBrush}" 
                        BorderBrush="{DynamicResource CardBorderBrush}" 
                        BorderThickness="1" CornerRadius="4" Margin="0,2,4,2" Padding="6">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        
                        <!-- Color Circle: Hidden when null to maintain alignment -->
                        <Border Grid.Column="0" Width="12" Height="12" CornerRadius="6" Margin="0,2,8,0" 
                                BorderBrush="#888888" BorderThickness="1" VerticalAlignment="Top">
                            <Border.Style>
                                <Style TargetType="Border">
                                    <Setter Property="Background" Value="{Binding AssignedColorBrush}"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding AssignedColorBrush}" Value="{x:Null}">
                                            <Setter Property="Visibility" Value="Hidden"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </Border.Style>
                        </Border>
                        
                        <!-- Content Card -->
                        <StackPanel Grid.Column="1">
                            <TextBlock Text="{Binding Name}" FontWeight="{Binding FontWeight}" 
                                       Foreground="{DynamicResource TextBrush}" 
                                       TextWrapping="Wrap" Margin="0,0,0,2"/>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="{Binding Count, StringFormat='Qty: {0}'}" 
                                           Foreground="{DynamicResource TextLightBrush}" FontSize="10" Margin="0,0,10,0"/>
                                <TextBlock Text="{Binding DisplayValue}" 
                                           Foreground="{DynamicResource TextLightBrush}" FontSize="10"/>
                            </StackPanel>
                        </StackPanel>
                    </Grid>
                </Border>
            </HierarchicalDataTemplate>
            """
            self.systemTree.ItemTemplate = XamlReader.Parse(tree_template)
            
            # Inject ItemContainerStyle to stretch TreeViewItems (Full Width Cards)
            # We use a custom ControlTemplate to ensure the header column is Width="*"
            style_xaml = """
            <Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
                   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
                   TargetType="TreeViewItem">
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                <Setter Property="IsExpanded" Value="{Binding IsExpanded, Mode=TwoWay}"/>
                <Setter Property="IsSelected" Value="{Binding IsSelected, Mode=TwoWay}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="TreeViewItem">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" MinWidth="19"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition/>
                                </Grid.RowDefinitions>
                                <ToggleButton x:Name="Expander" ClickMode="Press" IsChecked="{Binding IsExpanded, RelativeSource={RelativeSource TemplatedParent}}" VerticalAlignment="Center">
                                    <ToggleButton.Template>
                                        <ControlTemplate TargetType="ToggleButton">
                                            <Border Background="Transparent" Height="16" Padding="5" Width="16">
                                                <Path x:Name="ExpandPath" Data="M0,0 L0,6 L6,0 z" Fill="{DynamicResource TextLightBrush}" Stroke="{DynamicResource TextLightBrush}">
                                                    <Path.RenderTransform>
                                                        <RotateTransform Angle="135" CenterX="3" CenterY="3"/>
                                                    </Path.RenderTransform>
                                                </Path>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsChecked" Value="True">
                                                    <Setter Property="RenderTransform" TargetName="ExpandPath"><Setter.Value><RotateTransform Angle="180" CenterX="3" CenterY="3"/></Setter.Value></Setter>
                                                    <Setter Property="Fill" TargetName="ExpandPath" Value="{DynamicResource AccentBrush}"/>
                                                    <Setter Property="Stroke" TargetName="ExpandPath" Value="{DynamicResource AccentBrush}"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </ToggleButton.Template>
                                </ToggleButton>
                                <Border x:Name="Bd" Grid.Column="1" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="true">
                                    <ContentPresenter x:Name="PART_Header" ContentSource="Header" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
                                </Border>
                                <ItemsPresenter x:Name="ItemsHost" Grid.Column="1" Grid.Row="1"/>
                            </Grid>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsExpanded" Value="False"><Setter Property="Visibility" TargetName="ItemsHost" Value="Collapsed"/></Trigger>
                                <Trigger Property="IsSelected" Value="True"><Setter Property="Background" TargetName="Bd" Value="{DynamicResource SelectionBrush}"/><Setter Property="BorderBrush" TargetName="Bd" Value="{DynamicResource SelectionBorderBrush}"/></Trigger>
                                <MultiTrigger><MultiTrigger.Conditions><Condition Property="IsSelected" Value="True"/><Condition Property="IsSelectionActive" Value="False"/></MultiTrigger.Conditions><Setter Property="Background" TargetName="Bd" Value="{DynamicResource InactiveSelectionBrush}"/><Setter Property="BorderBrush" TargetName="Bd" Value="Transparent"/></MultiTrigger>
                                <Trigger Property="IsEnabled" Value="False"><Setter Property="Foreground" Value="{DynamicResource TextLightBrush}"/></Trigger>
                                <Trigger Property="HasItems" Value="False"><Setter Property="Visibility" TargetName="Expander" Value="Hidden"/></Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            """
            self.systemTree.ItemContainerStyle = XamlReader.Parse(style_xaml)
            
            # Inject DataGrid Template Column for Color
            # We clear existing columns to ensure order and binding, assuming standard columns
            if self.sysDataGrid.Columns.Count > 0:
                col_template = """
                <DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
                    <Border Width="10" Height="10" CornerRadius="5" HorizontalAlignment="Center" VerticalAlignment="Center"
                            Background="{Binding AssignedColorBrush}" BorderBrush="#888888" BorderThickness="1">
                         <Border.Style>
                            <Style TargetType="Border">
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding AssignedColorBrush}" Value="{x:Null}">
                                        <Setter Property="Visibility" Value="Hidden"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </Border.Style>
                    </Border>
                </DataTemplate>
                """
                # Create a TemplateColumn in code is verbose, but we can try to insert it if we can parse the column
                # Alternatively, we rely on the TreeView update which is the primary request "Tree view and datagrid".
                # Since DataGrid columns are often auto-generated or hardcoded in XAML, modifying them via Python without XAML access is risky.
                # However, we can try to add a column at index 0.
                
                # Let's try to construct a DataGridTemplateColumn using XamlReader
                col_xaml = """
                <DataGridTemplateColumn xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Header="Color">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Border Width="10" Height="10" CornerRadius="5" HorizontalAlignment="Center" VerticalAlignment="Center"
                                    Background="{Binding AssignedColorBrush}" BorderBrush="#888888" BorderThickness="1"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                """
                col = XamlReader.Parse(col_xaml)
                self.sysDataGrid.Columns.Insert(0, col)

        except Exception as e:
            print("Error injecting templates: {}".format(e))

        # Check for pre-selection
        sel_ids = self.uidoc.Selection.GetElementIds()
        if sel_ids.Count > 0:
            self.analyze_selection(sel_ids)

    def set_default_header(self):
        default_node = NodeBase("Quantity & Measures")
        default_node.Type = "Scan view or select elements to begin."
        self.RightPane.DataContext = default_node

    def get_solid_pattern_id(self):
        """Lazy loads the Solid Fill Pattern ID."""
        if self.solid_pattern_id: return self.solid_pattern_id
        patterns = FilteredElementCollector(self.doc).OfClass(FillPatternElement)
        for p in patterns:
            if p.GetFillPattern().IsSolidFill:
                self.solid_pattern_id = p.Id
                return p.Id
        return None

    # --- UI Logic ---
    def drag_window(self, sender, args):
        self.DragMove()

    def close_window(self, sender, args):
        self.Close()

    def apply_revit_theme(self):
        """Detects Revit theme and updates window resources."""
        res = self.Resources
        
        # Default (Light Theme) Card Styles
        res["CardBrush"] = SolidColorBrush(Colors.White)
        res["CardBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(220, 220, 220))
        res["SelectionTextBrush"] = SolidColorBrush(Colors.Black)
        res["InactiveSelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(224, 224, 224))
        
        if is_dark_theme():
            # Define Modern Dark Theme Colors (Slate/Blue Palette)
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
            res["SelectionTextBrush"] = SolidColorBrush(Colors.White)
            res["InactiveSelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81)) # Gray-700
            
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
        can_clear = has_checked or has_selection or self.element_color_map or self.disabled_filters
        
        self.Btn_Visualize.IsEnabled = has_checked or has_selection
        self.Btn_ClearVisuals.IsEnabled = can_clear
        self.Btn_Export.IsEnabled = has_checked
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
        self.Btn_Export.IsEnabled = False
        self.Btn_Isolate.IsEnabled = False
        self.Btn_Isolate.Content = "Isolate"
        self.set_default_header()
        self.update_button_states()

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
        self.statusLabel.Text = "Scanning Active View..."
        view = self.doc.ActiveView
        collector = FilteredElementCollector(self.doc, view.Id).WhereElementIsNotElementType()
        
        # Enhanced Visibility Check: Explicitly check Category visibility
        # Optimization: Cache category visibility to reduce API calls
        cat_hidden_cache = {}
        
        elements = []
        for e in collector:
            if e.IsHidden(view): continue
            if e.Category:
                cid_val = get_id(e.Category.Id)
                if cid_val not in cat_hidden_cache:
                    is_hidden = False
                    if view.CanCategoryBeHidden(e.Category.Id) and view.GetCategoryHidden(e.Category.Id):
                        is_hidden = True
                    cat_hidden_cache[cid_val] = is_hidden
                
                if cat_hidden_cache[cid_val]:
                    continue
            elements.append(e)
            
        self.process_elements(elements, "Active View")

    def analyze_selection(self, element_ids):
        """Analyzes a specific set of elements."""
        elements = []
        for eid in element_ids:
            el = self.doc.GetElement(eid)
            if el: elements.append(el)
        self.process_elements(elements, "Selection")

    def get_color_from_ogs(self, ogs):
        """Extracts the highest priority color from OverrideGraphicSettings."""
        if not ogs: return None
        
        # Check properties safely (Revit 2019+ vs older)
        # Priority: Surface BG > Cut BG > Surface FG > Cut FG
        props = [
            "SurfaceBackgroundPatternColor",
            "CutBackgroundPatternColor",
            "SurfaceForegroundPatternColor",
            "CutForegroundPatternColor",
            "ProjectionFillColor", # Fallback for older Revit
            "CutFillColor"         # Fallback for older Revit
        ]
        
        for p in props:
            try:
                if hasattr(ogs, p):
                    c = getattr(ogs, p)
                    if c.IsValid: return c
            except: pass
            
        return None

    def get_element_color(self, element, active_view, filter_data, cat_color_cache):
        """Resolves the effective Revit Color for an element (Element > Filter > Category)."""
        # 1. Element Override
        ogs = active_view.GetElementOverrides(element.Id)
        # This function should NOT cache the OGS, as it might be a temporary highlight.
        
        c = self.get_color_from_ogs(ogs)
        
        # 2. Filter Override (if no element override)
        if not c:
            for f_filter, f_color in filter_data:
                if f_filter.PassesFilter(element):
                    c = f_color
                    break # Top priority filter wins
        
        # 3. Category Override (if no filter override)
        if not c and element.Category:
            cid_val = get_id(element.Category.Id)
            if cid_val not in cat_color_cache:
                ogs_cat = active_view.GetCategoryOverrides(element.Category.Id)
                cat_color_cache[cid_val] = self.get_color_from_ogs(ogs_cat)
            c = cat_color_cache[cid_val]
        
        if c:
            return c
        return None

    def process_elements(self, elements, source_name="Selection"):
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
            
            # Optimization: Pre-compute lowercase set for O(1) lookup
            measurable_names_lower = {n.lower() for n in MEASURABLE_NAMES}
            
            active_view = self.doc.ActiveView
            self.element_color_map = {} # Reset map on new scan
            self.instance_item_map = {} # Reset map on new scan
            
            # Pre-fetch Filter Colors (Optimization)
            filter_data = [] # List of (ElementFilter, Color)
            for fid in active_view.GetFilters():
                if active_view.GetFilterVisibility(fid):
                    ogs = active_view.GetFilterOverrides(fid)
                    c = self.get_color_from_ogs(ogs)
                    if c:
                        f_elem = self.doc.GetElement(fid)
                        if isinstance(f_elem, ParameterFilterElement):
                            el_filter = f_elem.GetElementFilter()
                            if el_filter:
                                filter_data.append((el_filter, c))
            
            cat_color_cache = {} # CategoryId -> Color

            for el in elements:
                eid_int = get_id(el.Id)
                # Cache common properties
                if el.Category:
                    cat_name = el.Category.Name
                    cat_id_val = get_id(el.Category.Id)
                else:
                    cat_name = "Uncategorized"
                    cat_id_val = -1

                # Extract Color Override
                assigned_color = None
                try:
                    assigned_color = self.get_element_color(el, active_view, filter_data, cat_color_cache)
                    # Populate the persistent color map only on the initial scan
                    if assigned_color:
                        self.element_color_map[eid_int] = active_view.GetElementOverrides(el.Id)
                except: pass

                try:
                    type_name = el.Name
                except AttributeError:
                    type_name = "Unnamed Element"
                
                # 1. Always add to "Count" (Ensures Tags, Lines, etc. are listed)
                if "Count" not in tree_data: tree_data["Count"] = {}
                c_dict = tree_data["Count"]
                if cat_name not in c_dict: c_dict[cat_name] = {}
                t_dict = c_dict[cat_name]
                if type_name not in t_dict: t_dict[type_name] = []
                item_count = InstanceItem(el, 1.0, "ea", "-")
                
                # Map the instance item for direct updates later
                if eid_int not in self.instance_item_map: self.instance_item_map[eid_int] = []
                self.instance_item_map[eid_int].append(item_count)
                
                if assigned_color: item_count.RevitColor = assigned_color
                t_dict[type_name].append(item_count)
                
                # Iterate Parameters
                for p in el.Parameters:
                    if p.StorageType == StorageType.Double and p.HasValue:
                        p_name = p.Definition.Name
                        
                        # Optimization: Fast Lookup
                        if p_name.lower() not in measurable_names_lower:
                            continue
                        
                        val, unit_label = get_display_val_and_label(p, self.doc)
                        if abs(val) < 0.0001: continue # Skip zero values
                        
                        # Run Calculation if applicable
                        calc_val = "-"
                        if cat_name == "Walls":
                            calc_val = self.wall_calculator.calculate(el, self.calc_settings["Walls"])
                        
                        # Optimized Dictionary Access
                        if p_name not in tree_data: tree_data[p_name] = {}
                        cat_dict = tree_data[p_name]
                        
                        if cat_name not in cat_dict: cat_dict[cat_name] = {}
                        type_dict = cat_dict[cat_name]
                        
                        if type_name not in type_dict: type_dict[type_name] = []
                        item_param = InstanceItem(el, val, unit_label, calc_val)
                        
                        # Map the instance item for direct updates later
                        if eid_int not in self.instance_item_map: self.instance_item_map[eid_int] = []
                        self.instance_item_map[eid_int].append(item_param)
                        
                        if assigned_color: item_param.RevitColor = assigned_color
                        type_dict[type_name].append(item_param)

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
                        
                        # Aggregate Color for Type Node (Mixed = Black)
                        t_node.RevitColor = self.aggregate_colors([i.RevitColor for i in instances])

                        c_node.Children.append(t_node)
                        c_val += t_val
                        c_count += t_count
                        c_node.AllElements.extend(t_node.AllElements)
                    
                    c_node.Value = c_val
                    c_node.Count = c_count
                    
                    # Aggregate Color for Category Node
                    c_node.RevitColor = self.aggregate_colors([c.RevitColor for c in c_node.Children])
                    
                    if c_node.Children:
                        c_node.UnitLabel = c_node.Children[0].UnitLabel
                    m_node.Children.append(c_node)
                    total_val += c_val
                    total_count += c_count
                    m_node.AllElements.extend(c_node.AllElements)
                
                m_node.Value = total_val
                m_node.Count = total_count
                
                # Aggregate Color for Measurement Node
                m_node.RevitColor = self.aggregate_colors([c.RevitColor for c in m_node.Children])
                
                if m_node.Children:
                    m_node.UnitLabel = m_node.Children[0].UnitLabel
                root_nodes.append(m_node)

            self.systemTree.ItemsSource = root_nodes
            self.statusLabel.Text = "Found {} measurable parameters in {}.".format(len(root_nodes), source_name)
            
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

    def aggregate_colors(self, colors):
        """
        Aggregates a list of Revit Colors.
        - All None -> None
        - All Same Color -> Color
        - Mixed (Colors vs None, or Color A vs Color B) -> Black
        """
        if not colors: return None
        
        unique_colors = set()
        has_none = False
        
        for c in colors:
            if c is None:
                has_none = True
            else:
                # Store color as string or tuple to be hashable/comparable
                unique_colors.add((c.Red, c.Green, c.Blue))
        
        # Case 1: All None
        if not unique_colors:
            return None
            
        # Case 2: Mixed (Multiple colors OR Color + None)
        if len(unique_colors) > 1 or has_none:
            return Color(0, 0, 0) # Black
            
        # Case 3: Single Color
        r, g, b = list(unique_colors)[0]
        return Color(r, g, b)

    def tree_double_click(self, sender, args):
        """Zooms to the selected element(s) in the tree on double-click."""
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node:
                ids = set(selected_node.AllElements)
                if isinstance(selected_node, (MeasurementNode, CategoryNode)):
                    ids = self._get_all_child_elements(selected_node)
                
                if ids:
                    elem_ids = List[ElementId]([ElementId(i) for i in ids])
                    self.uidoc.ShowElements(elem_ids)
        except: pass

    def grid_double_click(self, sender, args):
        """Zooms to the selected element(s) in the grid on double-click."""
        try:
            selected_items = self.sysDataGrid.SelectedItems
            if selected_items:
                ids = [item.Id for item in selected_items if hasattr(item, "Id")]
                if ids:
                    elem_ids = List[ElementId]([ElementId(i) for i in ids])
                    self.uidoc.ShowElements(elem_ids)
        except: pass

    def setup_context_menu(self):
        """Adds a context menu to the TreeView to copy values."""
        ctx_menu = ContextMenu()
        item_copy = MenuItem()
        item_copy.Header = "Copy Value to Clipboard"
        item_copy.Click += self.copy_clipboard_click
        ctx_menu.Items.Add(item_copy)
        self.systemTree.ContextMenu = ctx_menu

    def copy_clipboard_click(self, sender, args):
        """Copies the current value (or selected sub-total) to clipboard."""
        try:
            node = self.systemTree.SelectedItem
            if node and hasattr(node, "Value"):
                val = node.Value
                if hasattr(node, "SelectedValue") and node.SelectedValue > 0.0001:
                    val = node.SelectedValue
                txt = "{:.2f}".format(val)
                Clipboard.SetText(txt)
                self.statusLabel.Text = "Copied '{}' to clipboard.".format(txt)
        except: pass

    def setup_grid_context_menu(self):
        """Adds a context menu to the DataGrid to copy values."""
        ctx_menu = ContextMenu()
        item_copy = MenuItem()
        item_copy.Header = "Copy Value to Clipboard"
        item_copy.Click += self.copy_grid_clipboard_click
        ctx_menu.Items.Add(item_copy)
        self.sysDataGrid.ContextMenu = ctx_menu

    def copy_grid_clipboard_click(self, sender, args):
        """Copies the sum of selected items in the grid to clipboard."""
        try:
            selected_items = self.sysDataGrid.SelectedItems
            if selected_items:
                total_val = sum(item.Value for item in selected_items if hasattr(item, "Value"))
                txt = "{:.2f}".format(total_val)
                Clipboard.SetText(txt)
                self.statusLabel.Text = "Copied '{}' to clipboard.".format(txt)
        except: pass

    def reset_selection_highlight(self):
        """Resets the temporary orange highlight on previously selected elements."""
        if self.last_highlighted_ids:
            try:
                with Transaction(self.doc, "Reset Highlight") as t:
                    t.Start()
                    solid_pid = self.get_solid_pattern_id()
                    
                    for eid in self.last_highlighted_ids:
                        eid_int = get_id(eid)
                        if eid_int in self.element_color_map:
                            # Restore Persistent Color instead of clearing
                            ogs = self.element_color_map[eid_int]
                            self.doc.ActiveView.SetElementOverrides(eid, ogs)
                        else:
                            # Reset to Default
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
                # Reset Selected Totals on tree change
                if isinstance(selected_node, NodeBase):
                    selected_node.SelectedCount = 0
                    selected_node.SelectedValue = 0.0

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

                # Apply Dim Background (Halftone + Transparent) only
                with Transaction(self.doc, "Highlight Selection") as t:
                    t.Start()
                    
                    # Dim Background (Transparency Only, Preserve Colors)
                    if bg_elem:
                        default_dim_ogs = OverrideGraphicSettings()
                        default_dim_ogs.SetSurfaceTransparency(60) # 60% Transparent
                        
                        for eid in bg_elem:
                            eid_int = get_id(eid)
                            if eid_int in self.element_color_map:
                                # Preserve existing overrides (e.g. Colorize colors)
                                base_ogs = self.element_color_map[eid_int]
                                ogs_dim = OverrideGraphicSettings(base_ogs)
                                ogs_dim.SetSurfaceTransparency(60)
                                self.doc.ActiveView.SetElementOverrides(eid, ogs_dim)
                            else:
                                # Apply default dimming
                                self.doc.ActiveView.SetElementOverrides(eid, default_dim_ogs)

                    t.Commit()
                
                # Store ElementIds for reset
                self.last_highlighted_ids = bg_elem
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
            
            # Sync Revit Selection (Cyan)
            elem_ids = List[ElementId]()
            if current_ids:
                elem_ids = List[ElementId]([ElementId(i) for i in current_ids])
            
            self.uidoc.Selection.SetElementIds(elem_ids)
            
            # Auto-Zoom & Select in Revit
            if self.Cb_AutoZoom.IsChecked and current_ids:
                self.uidoc.ShowElements(elem_ids)

            # Update Selected Totals in Dashboard
            s_count = 0
            s_val = 0.0
            if selected_items:
                s_count = len(selected_items)
                for item in selected_items:
                    if hasattr(item, "Value"):
                        s_val += item.Value
            
            node = self.systemTree.SelectedItem
            if node and isinstance(node, NodeBase):
                node.SelectedCount = s_count
                node.SelectedValue = s_val

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

    def get_target_elements_for_visuals(self):
        """Determines which elements to target based on UI selection hierarchy."""
        ids = set()
        
        # 1. Grid Selection (Specific Instances)
        if self.sysDataGrid.SelectedItems and self.sysDataGrid.SelectedItems.Count > 0:
            for item in self.sysDataGrid.SelectedItems:
                if hasattr(item, "Id"):
                    ids.add(item.Id)
            return ids # If grid has selection, prioritize it exclusively
            
        # 2. Tree Selection (Category/Type/Measurement)
        if self.systemTree.SelectedItem:
            node = self.systemTree.SelectedItem
            if hasattr(node, "AllElements"):
                if isinstance(node, (MeasurementNode, CategoryNode)):
                    ids.update(self._get_all_child_elements(node))
                else:
                    ids.update(node.AllElements)
            return ids # If tree has selection, prioritize it
            
        # 3. Checked Items (Batch)
        checked = self.get_checked_systems()
        for node in checked:
            ids.update(node.AllElements)
            
        return ids
        
    def get_target_objects_for_visuals(self):
        """Returns the actual ViewModel objects (Nodes/Instances) targeted for visualization."""
        objects = []
        
        # 1. Grid Selection
        if self.sysDataGrid.SelectedItems and self.sysDataGrid.SelectedItems.Count > 0:
            return list(self.sysDataGrid.SelectedItems)
            
        # 2. Tree Selection
        if self.systemTree.SelectedItem:
            return [self.systemTree.SelectedItem]
            
        # 3. Checked Items
        checked = self.get_checked_systems()
        if checked: return checked
        return []

    def visualize_click(self, sender, args):
        """Applies Color Overrides to selected or checked elements."""
        if self.is_busy: return
        
        ids_to_color = self.get_target_elements_for_visuals()
        if not ids_to_color:
            forms.alert("Please select or check elements to colorize.")
            return
        
        selected_color_opt = self.Cmb_Colors.SelectedItem
        if not selected_color_opt:
            forms.alert("Please select a color from the dropdown.")
            return

        self.is_busy = True
        self.Cursor = Cursors.Wait

        revit_color = selected_color_opt.RevitColor
        wpf_brush = selected_color_opt.Brush

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

            with Transaction(self.doc, "Colorize Elements") as t:
                t.Start()
                
                solid_pid = self.get_solid_pattern_id()
                
                ogs = OverrideGraphicSettings()
                if solid_pid:
                    ogs.SetSurfaceBackgroundPatternId(solid_pid)
                    ogs.SetSurfaceBackgroundPatternColor(revit_color)
                    # Also apply to Cut Pattern for consistency in sections/plans
                    ogs.SetCutBackgroundPatternId(solid_pid)
                    ogs.SetCutBackgroundPatternColor(revit_color)
                
                for eid_int in ids_to_color:
                    self.element_color_map[eid_int] = ogs # Update persistent map
                    try:
                        eid = ElementId(eid_int)
                        self.doc.ActiveView.SetElementOverrides(eid, ogs)
                    except: pass

                t.Commit()
                
                # Update ViewModel directly instead of re-scanning the view
                for eid_int in ids_to_color:
                    if eid_int in self.instance_item_map:
                        for item in self.instance_item_map[eid_int]:
                            item.RevitColor = revit_color
                
                # Re-aggregate colors up the tree
                self.refresh_tree_colors()
                
                # Remove colorized elements from the temporary highlight list
                # so they aren't reset when selection changes or window closes.
                self.last_highlighted_ids = [
                    eid for eid in self.last_highlighted_ids 
                    if get_id(eid) not in ids_to_color
                ]
                
                # Clear selection so the color overrides are clearly visible (not masked by selection highlight)
                self.uidoc.Selection.SetElementIds(List[ElementId]())
                self.uidoc.RefreshActiveView()
                self.uidoc.UpdateAllOpenViews()
                self.statusLabel.Text = "Colorized {} elements.".format(len(ids_to_color))
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Visualization Error. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow
            self.update_button_states()

    def refresh_tree_colors(self):
        """Re-calculates aggregated colors for the entire tree."""
        if not self.systemTree.ItemsSource: return
        
        for m_node in self.systemTree.ItemsSource:
            m_brushes = []
            for c_node in m_node.Children:
                c_brushes = []
                for t_node in c_node.Children:
                    # Aggregate Instances -> Type
                    inst_colors = [i.RevitColor for i in t_node.Instances]
                    t_node.RevitColor = self.aggregate_colors(inst_colors)
                    c_brushes.append(t_node.RevitColor)
                
                # Aggregate Types -> Category
                c_node.RevitColor = self.aggregate_colors(c_brushes)
                m_brushes.append(c_node.RevitColor)
            
            # Aggregate Categories -> Measurement
            m_node.RevitColor = self.aggregate_colors(m_brushes)

    def reset_visuals_click(self, sender, args):
        """Clears graphic overrides for the listed elements."""
        if self.is_busy: return
        
        ids_to_clear = self.get_target_elements_for_visuals()
        
        # If nothing selected, ask to clear ALL colors applied by the tool
        if not ids_to_clear and self.element_color_map:
            td = TaskDialog("Reset Visuals")
            td.MainInstruction = "No elements selected."
            td.MainContent = "Do you want to clear ALL colors applied by this tool in the current view?"
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
            if td.Show() == TaskDialogResult.Yes:
                ids_to_clear.update(self.element_color_map.keys())
        
        # Allow reset if we have disabled filters, even if no systems are checked
        if not ids_to_clear and not self.disabled_filters:
            forms.alert("No elements selected or colors to clear.")
            return

        self.is_busy = True
        self.Cursor = Cursors.Wait

        try:
            with Transaction(self.doc, "Reset Visuals") as t:
                t.Start()
                for eid_int in ids_to_clear:
                    if eid_int in self.element_color_map:
                        del self.element_color_map[eid_int] # Remove from map
                    try:
                        eid = ElementId(eid_int)
                        self.doc.ActiveView.SetElementOverrides(eid, OverrideGraphicSettings())
                    except: pass
                
                # Restore View Filters if we disabled them
                if self.disabled_filters:
                    view = self.doc.ActiveView
                    for fid in self.disabled_filters:
                        if view.IsFilterApplied(fid):
                            view.SetFilterVisibility(fid, True)
                    self.disabled_filters = [] # Clear list after restoring

                t.Commit()
                
                # Update ViewModel directly
                for eid_int in ids_to_clear:
                    if eid_int in self.instance_item_map:
                        for item in self.instance_item_map[eid_int]:
                            item.RevitColor = None

                # Re-aggregate colors up the tree
                self.refresh_tree_colors()
                self.uidoc.RefreshActiveView()
                self.uidoc.UpdateAllOpenViews()
                self.statusLabel.Text = "Visual overrides reset."
        except Exception as e:
            err = traceback.format_exc()
            print(err)
            self.statusLabel.Text = "Reset Error. Check Output."
        finally:
            self.is_busy = False
            self.Cursor = Cursors.Arrow
            self.update_button_states()

    def export_click(self, sender, args):
        """Exports the current tree data to a CSV file."""
        if not self.systemTree.ItemsSource:
            return

        # Check if any items are checked
        checked_items = self.get_checked_systems()
        has_checked = len(checked_items) > 0
        
        if not has_checked:
            forms.alert("Please check items to export.")
            return

        # Generate Filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = "Quantities"
        
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
                    
                    for c_node in m_node.Children:
                        category_name = c_node.Name.encode('utf-8')
                        
                        for t_node in c_node.Children:
                            # Export only if explicitly checked
                            if t_node.IsChecked:
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
        
        # Reload settings and recalculate
        self.wall_calculator.load_settings()
        self.calc_settings["Walls"] = self.wall_calculator.default_setting
        self.recalculate_all()
        self.statusLabel.Text = "Settings saved and recalculated."

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
