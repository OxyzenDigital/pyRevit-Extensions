# -*- coding: utf-8 -*-
"""
Add Fittings Tool
Description: Adds a selected pipe fitting to the nearest open end of a selected pipe.
Version: 1.0
"""

__title__ = "Add Fittings"
__version__ = "1.0"
__context__ = "active-view-type: FloorPlan,CeilingPlan,EngineeringPlan,AreaPlan,Section,Elevation,ThreeD"

import sys
import clr
import math

# --- ASSEMBLIES ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')

import System.Windows
from System.Windows import SystemParameters

# --- IMPORTS ---
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    Transaction, ElementId, BuiltInCategory, FilteredElementCollector,
    FamilySymbol, Structure, XYZ, ConnectorProfileType, ElementTransformUtils,
    BuiltInParameter
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc

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

class FittingOption(object):
    """Wrapper for FamilySymbol to provide a nice display name."""
    def __init__(self, symbol):
        self.Symbol = symbol
        
        # Robust Name Retrieval
        self.FamilyName = self._get_safe_name(symbol, is_family=True)
        self.SymbolName = self._get_safe_name(symbol, is_family=False)
        
    def _get_safe_name(self, symbol, is_family=False):
        try:
            if is_family:
                return symbol.Family.Name
            else:
                return symbol.Name
        except:
            # Fallback to Parameters
            try:
                param_id = BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM if is_family else BuiltInParameter.SYMBOL_NAME_PARAM
                p = symbol.get_Parameter(param_id)
                if p and p.HasValue:
                    return p.AsString()
            except:
                pass
            return "Unknown"

    @property
    def DisplayName(self):
        return "{} : {}".format(self.FamilyName, self.SymbolName)

def get_all_pipe_fittings(doc):
    """Returns a sorted list of FittingOption objects for ALL Pipe Fittings."""
    collector = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PipeFitting)
        .OfClass(FamilySymbol)
    )
    
    fittings = []
    
    for symbol in collector:
        fittings.append(FittingOption(symbol))
            
    # Sort by Family Name then Symbol Name
    return sorted(fittings, key=lambda x: (x.FamilyName, x.SymbolName))

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

# ==========================================
# 2. UI CLASS
# ==========================================

class AddFittingWindow(forms.WPFWindow):
    def __init__(self, fittings, selection_data):
        """
        selection_data: dict containing:
            'pipes': list of Pipe Elements
            'ref_point': XYZ (optional, for single selection)
        """
        forms.WPFWindow.__init__(self, 'ui.xaml')
        
        self.fittings = fittings
        self.selection_data = selection_data
        
        # Populate Fittings
        self.Cmb_Fittings.ItemsSource = self.fittings
        if self.fittings:
            self.Cmb_Fittings.SelectedIndex = 0
            
        # UI State Logic
        self.has_point = self.selection_data.get('ref_point') is not None
        
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

    def set_window_position(self):
        try:
            sw = SystemParameters.PrimaryScreenWidth
            sh = SystemParameters.PrimaryScreenHeight
            self.Left = sw * 0.3
            self.Top = sh * 0.3
        except:
            pass

    def bind_events(self):
        self.Btn_Add.Click += self.on_add_click
        self.Btn_Close.Click += lambda s, e: self.Close()
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window

    def drag_window(self, sender, args):
        try: self.DragMove()
        except: pass

    def on_add_click(self, sender, args):
        selected_option = self.Cmb_Fittings.SelectedItem
        if not selected_option:
            self.Lb_Status.Text = "Please select a fitting type."
            return

        selected_fitting = selected_option.Symbol

        # Check if Symbol is active
        if not selected_fitting.IsActive:
            t_act = Transaction(doc, "Activate Symbol")
            t_act.Start()
            selected_fitting.Activate()
            doc.Regenerate()
            t_act.Commit()

        self.Close()
        
        # Run Logic
        try:
            perform_add_fitting(self.selection_data, selected_fitting)
        except Exception as e:
            forms.alert("Error adding fitting: {}".format(e))

# ==========================================
# 3. LOGIC
# ==========================================

def perform_add_fitting(data, symbol):
    pipes = data['pipes']
    ref_point = data.get('ref_point')
    
    t = Transaction(doc, "Add Pipe Fitting")
    t.Start()
    
    count = 0
    
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
                        Structure.StructuralType.NonStructural
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
                                print("Connection warning: {}".format(e_conn))
                                pass
                                
                    count += 1
                except Exception as e:
                    print("Failed on pipe {}: {}".format(pipe.Id, e))

        t.Commit()
        
    except Exception as e:
        t.RollBack()
        forms.alert("Critical Error: {}".format(e))

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

    # 3. Gather Data
    data = {
        'pipes': pipes,
        'ref_point': ref_point
    }
    
    # 4. Load Fittings
    fittings = get_all_pipe_fittings(doc)
    
    if not fittings:
        forms.alert("No Pipe Fitting families found in project.")
        sys.exit(0)
        
    # 5. Launch UI
    win = AddFittingWindow(fittings, data)
    win.ShowDialog()
