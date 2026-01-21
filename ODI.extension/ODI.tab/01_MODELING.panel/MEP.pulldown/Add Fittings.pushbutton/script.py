# -*- coding: utf-8 -*-
"""
Add Fittings Tool
Description: Adds a selected pipe fitting to the nearest open end of a selected pipe.
Version: 1.2
"""

__title__ = "Add Fittings"
__version__ = "1.2"
__context__ = "active-view-type: FloorPlan,CeilingPlan,EngineeringPlan,AreaPlan,Section,Elevation,ThreeD"

import sys
import clr
import math

# --- ASSEMBLIES ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
import System.Windows
from System.Windows import SystemParameters
from System.Windows.Media import BrushConverter
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Input import ICommand
from System.Collections.Generic import List
import System.Drawing
from System.Windows.Interop import Imaging
from System.Windows import Int32Rect
from System.Windows.Media.Imaging import BitmapSizeOptions

# --- IMPORTS ---
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    Transaction, ElementId, BuiltInCategory, FilteredElementCollector,
    FamilySymbol, XYZ, ConnectorProfileType, ElementTransformUtils,
    BuiltInParameter, PartType
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, revit, script, HOST_APP

doc = revit.doc
uidoc = revit.uidoc

# --- Logging Setup ---
log_buffer = []

def log_section(title):
    log_buffer.append("\n### {}".format(title))

def log_item(key, value):
    log_buffer.append("- **{}:** {}".format(key, value))

def log_point(name, point):
    log_buffer.append("- **{}:** ({:.4f}, {:.4f}, {:.4f})".format(name, point.X, point.Y, point.Z))

def show_log():
    if not log_buffer: return
    output = script.get_output()
    output.close_others()
    for msg in log_buffer:
        output.print_md(msg)

# ==========================================
# 1. HELPERS
# ==========================================

def get_id_value(element_id):
    """
    Safe retrieval of Integer value from ElementId for Revit 2024+ compatibility.
    Revit 2023 and older use .IntegerValue (int).
    Revit 2024 and newer use .Value (long).
    """
    if hasattr(element_id, "IntegerValue"):
        return element_id.IntegerValue
    return element_id.Value

class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        if not e.Category:
            return False
        cat_id = e.Category.Id
        val = get_id_value(cat_id)
        return val == int(BuiltInCategory.OST_PipeCurves)
    def AllowReference(self, r, p):
        return False

def get_image_source(element):
    """Extracts the preview image from a Revit element and converts to WPF ImageSource."""
    try:
        # Get larger preview (128x128) for better quality
        bitmap = element.GetPreviewImage(System.Drawing.Size(128, 128))
        if bitmap:
            return Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                System.IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions()
            )
    except:
        pass
    return None
    
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
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self.Name = name
        self.Image = None
        self.Children = []
        self._is_expanded = False
        self.IsSelected = False
        self.FontWeight = "Normal"
    
    @property
    def IsExpanded(self):
        return self._is_expanded
    
    @IsExpanded.setter
    def IsExpanded(self, value):
        self._is_expanded = value
        self.OnPropertyChanged("IsExpanded")

class TypeNode(NodeBase):
    def __init__(self, symbol, name):
        NodeBase.__init__(self, name)
        self.Symbol = symbol
        self.Image = get_image_source(symbol)

class GroupNode(NodeBase):
    def __init__(self, name, children, font_weight="SemiBold"):
        NodeBase.__init__(self, name)
        self.Children = children
        self.FontWeight = font_weight
        self.Image = None

def get_safe_name(element, is_family=False):
    try:
        name = element.Family.Name if is_family else element.Name
        if name: return name
    except:
        pass
    
    # Fallback to Parameters if property fails
    try:
        p_id = BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM if is_family else BuiltInParameter.SYMBOL_NAME_PARAM
        p = element.get_Parameter(p_id)
        if p and p.HasValue:
            return p.AsString()
    except:
        pass
        
    return "Unknown"

def get_part_type_name(part_type):
    """Maps Revit PartType enum to a plural display string."""
    # Use string representation to be safe with IronPython Enum handling
    s_type = str(part_type)
    # Handle potential fully qualified names or prefixes
    if "." in s_type:
        s_type = s_type.split(".")[-1]
    mapping = {
        "Elbow": "Elbows",
        "Tee": "Tees",
        "Cross": "Crosses",
        "Transition": "Transitions",
        "Union": "Unions",
        "Cap": "Caps",
        "Coupling": "Couplings",
        "ValveBreaks": "Valves",
        "ValveNormal": "Valves",
        "PipeFlange": "Flanges",
        "Wye": "Wyes",
        "LateralTee": "Lateral Tees",
        "Undefined": "Undefined"
    }
    return mapping.get(s_type, s_type + "s")

def get_grouped_pipe_fittings(doc):
    """Returns a list of GroupNode objects (PartType -> Family -> Type)."""
    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeFitting)
        .OfClass(FamilySymbol)
    )
    
    # Structure: PartType -> FamilyName -> List[TypeNode]
    grouped = {}
    
    for symbol in collector:
        try:
            # Get Part Type
            try:
                # Retrieve Part Type via BuiltInParameter for robustness
                fam = symbol.Family
                param = fam.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
                if param and param.HasValue:
                    # Use System.Enum.ToObject for safer casting in IronPython
                    p_type = System.Enum.ToObject(PartType, param.AsInteger())
                    pt_name = get_part_type_name(p_type)
                else:
                    pt_name = get_part_type_name(fam.PartType)
            except:
                pt_name = "Other"

            fam_name = get_safe_name(symbol, is_family=True)
            sym_name = get_safe_name(symbol, is_family=False)
            
            if pt_name not in grouped:
                grouped[pt_name] = {}
            if fam_name not in grouped[pt_name]:
                grouped[pt_name][fam_name] = []
            
            grouped[pt_name][fam_name].append(TypeNode(symbol, sym_name))
        except:
            continue
    
    # Build Tree Nodes
    root_nodes = []
    for pt_name in sorted(grouped.keys()):
        fam_nodes = []
        for fam_name in sorted(grouped[pt_name].keys()):
            type_nodes = sorted(grouped[pt_name][fam_name], key=lambda x: x.Name)
            fam_nodes.append(GroupNode(fam_name, type_nodes, font_weight="Normal"))
        
        root_nodes.append(GroupNode(pt_name, fam_nodes, font_weight="Bold"))
        
    return root_nodes

def get_open_connectors(pipe):
    """Returns a list of connectors that are NOT connected."""
    open_connectors = []
    try:
        connectors = pipe.ConnectorManager.Connectors
        for conn in connectors:
            if not conn.IsConnected:
                open_connectors.append(conn)
    except:
        pass
    return open_connectors

def get_closest_connector(connectors, point):
    """Returns the connector closest to the given point."""
    if not connectors: return None
    best_conn = None
    min_dist = float('inf')
    for conn in connectors:
        dist = conn.Origin.DistanceTo(point)
        if dist < min_dist:
            min_dist = dist
            best_conn = conn
    return best_conn

def get_uiview():
    """Returns the UIView for the active document."""
    for uv in uidoc.GetOpenUIViews():
        if uv.ViewId == doc.ActiveView.Id:
            return uv
    return None

# ==========================================
# 2. UI CLASS
# ==========================================

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

class AddFittingViewModel(ViewModelBase):
    def __init__(self, family_nodes, selection_data, view):
        ViewModelBase.__init__(self)
        self.family_nodes = family_nodes
        self.selection_data = selection_data
        self.view = view
        
        self._selected_fitting = None
        self._rotation_increments = ["15", "30", "45", "90", "180"]
        self._selected_increment = "90"
        self._last_placed_ids = []
        self._status_message = "Ready"
        self._is_fitting_added = False
        self._initial_zoom_corners = None
        self._capture_initial_zoom()
        
        # Commands
        self.MainActionCommand = RelayCommand(self.main_action, self.can_main_action)
        self.CloseCommand = RelayCommand(self.close_window)
        self.ExpandAllCommand = RelayCommand(self.expand_all)
        self.CollapseAllCommand = RelayCommand(self.collapse_all)

    @property
    def FamilyNodes(self): return self.family_nodes

    @property
    def SelectedFitting(self): return self._selected_fitting
    @SelectedFitting.setter
    def SelectedFitting(self, value):
        self._selected_fitting = value
        self.OnPropertyChanged("SelectedFitting")
        self.MainActionCommand.RaiseCanExecuteChanged()

    @property
    def RotationIncrements(self): return self._rotation_increments

    @property
    def SelectedIncrement(self): return self._selected_increment
    @SelectedIncrement.setter
    def SelectedIncrement(self, value):
        self._selected_increment = value
        self.OnPropertyChanged("SelectedIncrement")

    @property
    def StatusMessage(self): return self._status_message
    @StatusMessage.setter
    def StatusMessage(self, value):
        self._status_message = value
        self.OnPropertyChanged("StatusMessage")

    @property
    def IsSelectionEnabled(self): return not self._is_fitting_added

    @property
    def IsAdjustmentEnabled(self): return self._is_fitting_added

    @property
    def MainButtonText(self):
        return "Rotate" if self._is_fitting_added else "Add Fitting"

    def main_action(self, parameter):
        if self._is_fitting_added:
            self.rotate_fitting(parameter)
        else:
            self.add_fitting(parameter)

    def can_main_action(self, parameter):
        if self._is_fitting_added:
            return self.can_rotate(parameter)
        else:
            return self.can_add_fitting(parameter)

    def _capture_initial_zoom(self):
        uv = get_uiview()
        if uv:
            self._initial_zoom_corners = uv.GetZoomCorners()

    def _restore_initial_zoom(self):
        if self._initial_zoom_corners:
            uv = get_uiview()
            if uv:
                uv.ZoomAndCenterRectangle(self._initial_zoom_corners[0], self._initial_zoom_corners[1])

    def can_add_fitting(self, parameter):
        return self._selected_fitting is not None

    def can_rotate(self, parameter):
        return len(self._last_placed_ids) > 0

    def expand_all(self, parameter):
        for node in self.family_nodes:
            self._set_expansion(node, True)

    def collapse_all(self, parameter):
        for node in self.family_nodes:
            self._set_expansion(node, False)

    def _set_expansion(self, node, expanded):
        node.IsExpanded = expanded
        for child in node.Children:
            # Recursively expand/collapse if children are also GroupNodes
            self._set_expansion(child, expanded)

    def add_fitting(self, parameter):
        if not self._selected_fitting: return
        
        # Activate Symbol if needed
        if not self._selected_fitting.Symbol.IsActive:
            t_act = Transaction(doc, "Activate Symbol")
            t_act.Start()
            self._selected_fitting.Symbol.Activate()
            doc.Regenerate()
            t_act.Commit()

        try:
            new_ids = perform_add_fitting(self.selection_data, self._selected_fitting.Symbol)
            if new_ids:
                self._last_placed_ids = new_ids
                self.StatusMessage = "Added {} fitting(s). Ready to adjust.".format(len(new_ids))
                # Zoom to the new fittings immediately
                uidoc.ShowElements(List[ElementId](new_ids))
                
                # Switch State
                self._is_fitting_added = True
                self.OnPropertyChanged("IsSelectionEnabled")
                self.OnPropertyChanged("IsAdjustmentEnabled")
                self.OnPropertyChanged("MainButtonText")
                self.MainActionCommand.RaiseCanExecuteChanged()
            else:
                self.StatusMessage = "No fittings added."
        except Exception as e:
            forms.alert("Error adding fitting: {}".format(e))
            self.StatusMessage = "Error occurred."

    def rotate_fitting(self, parameter):
        if not self._last_placed_ids: return
        
        try:
            inc = float(self._selected_increment)
        except:
            inc = 90.0
        angle_rad = inc * (math.pi / 180.0)
        
        t = Transaction(doc, "Rotate Fitting")
        t.Start()
        
        for eid in self._last_placed_ids:
            el = doc.GetElement(eid)
            if not el: continue
            
            # Find axis based on connection to pipe
            axis = None
            mep_model = el.MEPModel
            if mep_model:
                conns = mep_model.ConnectorManager.Connectors
                for c in conns:
                    if c.IsConnected:
                        for ref in c.AllRefs:
                            # Check if connected to a Pipe Curve
                            if ref.Owner.Id != el.Id and get_id_value(ref.Owner.Category.Id) == int(BuiltInCategory.OST_PipeCurves):
                                center = c.Origin
                                axis_dir = c.CoordinateSystem.BasisZ
                                axis = DB.Line.CreateBound(center, center + axis_dir)
                                break
                    if axis: break
            
            if axis:
                ElementTransformUtils.RotateElement(doc, eid, axis, angle_rad)
        
        t.Commit()
        
        uidoc.ShowElements(List[ElementId](self._last_placed_ids))
        self.StatusMessage = "Rotated {}° and Zoomed.".format(inc)

    def close_window(self, parameter):
        self._restore_initial_zoom()
        self.view.Close()

class AddFittingWindow(forms.WPFWindow):
    def __init__(self, family_nodes, selection_data):
        """
        selection_data: dict containing:
            'pipes': list of Pipe Elements
            'ref_point': XYZ (optional, for single selection)
        """
        forms.WPFWindow.__init__(self, 'ui.xaml')
        
        # Initialize ViewModel
        self.viewModel = AddFittingViewModel(family_nodes, selection_data, self)
        self.DataContext = self.viewModel
            
        # UI State Logic
        self.has_point = selection_data.get('ref_point') is not None
        
        if self.has_point:
             self.Tb_SelectionInfo.Text = "Mode: Smart Insert (Closest End)"
        else:
             self.Tb_SelectionInfo.Text = "Mode: Auto (Connect All Open Ends)"

        # Hide explicit placement options for simplicity as per request
        self.Rb_NearestEnd.Visibility = System.Windows.Visibility.Collapsed
        self.Rb_SelectionPoint.Visibility = System.Windows.Visibility.Collapsed
        
        # Position Window (30% from Top-Left)
        self.set_window_position()

        self.bind_events()
        self.apply_revit_theme()

    def on_tree_selection_changed(self, sender, args):
        """Handle TreeView selection manually since binding SelectedItem is complex."""
        selected_item = self.FittingTree.SelectedItem
        if isinstance(selected_item, TypeNode):
            self.viewModel.SelectedFitting = selected_item
        else:
            self.viewModel.SelectedFitting = None

    def set_window_position(self):
        try:
            sw = SystemParameters.PrimaryScreenWidth
            sh = SystemParameters.PrimaryScreenHeight
            self.Left = sw * 0.3
            self.Top = sh * 0.3
        except:
            pass
            
    def apply_revit_theme(self):
        # Simple detection for Dark Theme (Revit 2024+)
        is_dark = False
        try:
            # Check if running in Revit 2024 or later and if UITheme is Dark
            if int(HOST_APP.version) >= 2024:
                from Autodesk.Revit.UI import UIThemeManager, UITheme
                if UIThemeManager.CurrentTheme == UITheme.Dark:
                    is_dark = True
        except:
            pass
            
        if is_dark:
            # Dark Slate Theme (Matching System Browser)
            self.set_resource_color("WindowBrush", "#282a2f")
            self.set_resource_color("ControlBrush", "#374151")
            self.set_resource_color("TextBrush", "#FFFFFF")
            self.set_resource_color("TextLightBrush", "#9CA3AF")
            self.set_resource_color("BorderBrush", "#4B5563")
            self.set_resource_color("ButtonBrush", "#374151")
            self.set_resource_color("HoverBrush", "#4B5563")
            self.set_resource_color("SelectionBrush", "#4B5563")
            self.set_resource_color("SelectionBorderBrush", "#60A5FA")
            self.set_resource_color("AccentBrush", "#60A5FA")

    def set_resource_color(self, key, hex_color):
        try:
            converter = BrushConverter()
            brush = converter.ConvertFromString(hex_color)
            self.Resources[key] = brush
        except:
            pass

    def bind_events(self):
        # Only View-specific events (Drag) remain in code-behind
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.FittingTree.SelectedItemChanged += self.on_tree_selection_changed

    def drag_window(self, sender, args):
        try: self.DragMove()
        except: pass

# ==========================================
# 3. LOGIC
# ==========================================

def perform_add_fitting(data, symbol):
    pipes = data['pipes']
    ref_point = data.get('ref_point')

    log_section("Execution")
    log_item("Fitting", get_safe_name(symbol))
    log_item("Pipe Count", len(pipes))
    
    t = Transaction(doc, "Add Pipe Fitting")
    t.Start()
    
    count_success = 0
    created_ids = []
    
    try:
        for pipe in pipes:
            connector_to_connect = None
            
            # Get Pipe Diameter
            pipe_diam = 0.0
            p_diam_param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            if p_diam_param:
                pipe_diam = p_diam_param.AsDouble()
            
            # 1. Get ONLY open connectors
            open_conns = get_open_connectors(pipe)
            
            if not open_conns:
                log_item("Pipe {}".format(pipe.Id), "Skipped: No open connectors")
                continue
                
            # 2. Determine which end to use
            if ref_point:
                connector_to_connect = get_closest_connector(open_conns, ref_point)
            else:
                if open_conns:
                    connector_to_connect = open_conns[0]
            
            if connector_to_connect:
                # Initial placement at connector origin
                location_to_place = connector_to_connect.Origin
                
                try:
                    # A. Place Instance
                    instance = doc.Create.NewFamilyInstance(
                        location_to_place, 
                        symbol, 
                        DB.Structure.StructuralType.NonStructural
                    )
                    
                    if instance:
                        # B. Match Diameter
                        if pipe_diam > 0:
                            for param_name in ["Nominal Diameter", "Diameter", "Size", "Nominal Radius"]:
                                p = instance.LookupParameter(param_name)
                                if p and not p.IsReadOnly:
                                    try:
                                        if "Radius" in param_name:
                                            p.Set(pipe_diam / 2.0)
                                        else:
                                            p.Set(pipe_diam)
                                        break 
                                    except:
                                        pass
                        
                        # C. Regenerate to update geometry/connectors after sizing
                        doc.Regenerate()

                        # D. Find Matching Connector on Fitting
                        fit_conns = instance.MEPModel.ConnectorManager.Connectors
                        c_fit = get_closest_connector(fit_conns, connector_to_connect.Origin)
                        
                        if c_fit:
                            # E. Precise Alignment (Move)
                            # Calculate offset from current fitting connector to target pipe connector
                            translation = connector_to_connect.Origin - c_fit.Origin
                            if not translation.IsZeroLength():
                                ElementTransformUtils.MoveElement(doc, instance.Id, translation)
                                doc.Regenerate() # Update positions
                            
                            # F. Precise Alignment (Rotate)
                            # Target direction is opposite to pipe connector direction
                            pipe_dir = connector_to_connect.CoordinateSystem.BasisZ
                            fit_dir = c_fit.CoordinateSystem.BasisZ
                            target_dir = -pipe_dir # Opposed
                            
                            # Angle calculation
                            # Using Cross Product to find axis of rotation
                            cross = fit_dir.CrossProduct(target_dir)
                            angle = fit_dir.AngleTo(target_dir)
                            
                            if angle > 0.001 and not cross.IsZeroLength():
                                axis = DB.Line.CreateBound(c_fit.Origin, c_fit.Origin + cross)
                                ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)
                                doc.Regenerate()
                            elif angle > 0.001 and cross.IsZeroLength():
                                # 180 degree flip (Ambiguous axis)
                                # Pick an arbitrary perpendicular axis
                                # If fit_dir is Z, use X. Else use Z.
                                perp = XYZ.BasisX if abs(fit_dir.DotProduct(XYZ.BasisZ)) > 0.9 else XYZ.BasisZ
                                axis = DB.Line.CreateBound(c_fit.Origin, c_fit.Origin + perp)
                                ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)
                                doc.Regenerate()

                            # G. Connect
                            try:
                                c_fit.ConnectTo(connector_to_connect)
                            except Exception as e_conn:
                                log_item("Connection Warning", str(e_conn))
                                pass
                                
                    count_success += 1
                    created_ids.append(instance.Id)
                    log_item("Pipe {}".format(pipe.Id), "Success")
                except Exception as e:
                    log_item("Pipe {}".format(pipe.Id), "Failed: {}".format(e))

        t.Commit()
        
    except Exception as e:
        t.RollBack()
        log_item("Critical Error", str(e))

    log_section("Summary")
    log_item("Total Added", count_success)
    
    show_log()
    return created_ids

# ==========================================
# 4. MAIN
# ==========================================

if __name__ == '__main__':
    # 1. Initial Selection Check
    selection_ids = uidoc.Selection.GetElementIds()
    
    pipes = []
    ref_point = None
    
    # Process Pre-selection
    if selection_ids:
        for eid in selection_ids:
            el = doc.GetElement(eid)
            if el and el.Category:
                cat_id = el.Category.Id
                val = get_id_value(cat_id)
                if val == int(BuiltInCategory.OST_PipeCurves):
                    pipes.append(el)
    
    # 2. Interactive Selection if Empty
    if not pipes:
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, PipeSelectionFilter(), "Select a Pipe")
            if ref:
                el = doc.GetElement(ref)
                pipes.append(el)
                ref_point = ref.GlobalPoint
        except:
            # User cancelled
            sys.exit(0)

    if not pipes:
        forms.alert("No pipes selected.")
        sys.exit(0)

    # Check if any selected pipe has open ends
    has_valid_pipe = False
    for p in pipes:
        if get_open_connectors(p):
            has_valid_pipe = True
            break
            
    if not has_valid_pipe:
        forms.alert("Selected pipe(s) do not have any open ends.")
        sys.exit(0)

    # 3. Gather Data
    data = {
        'pipes': pipes,
        'ref_point': ref_point
    }
    
    # 4. Load Fittings
    family_nodes = get_grouped_pipe_fittings(doc)
    
    if not family_nodes:
        forms.alert("No Pipe Fitting families found in project.")
        sys.exit(0)
        
    # 5. Launch UI
    win = AddFittingWindow(family_nodes, data)
    win.ShowDialog()
