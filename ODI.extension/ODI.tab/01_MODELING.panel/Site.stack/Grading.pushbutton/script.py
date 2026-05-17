# -*- coding: utf-8 -*- 
__title__ = "Grading Assistant"
__doc__ = "Advanced Toposolid grading tool with sculpting, edging, and auto-triangulation features."
__author__ = "Oxyzen Digital"
__context__ = "doc-project"

import sys
import os
import json
import math
import clr
import traceback

# --- ASSEMBLIES ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

# --- IMPORTS ---
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    XYZ, Transaction, TransactionGroup, ElementId, BuiltInParameter, BuiltInCategory,
    ReferenceIntersector, FindReferenceTarget, Options, Solid, ViewType, View3D, Edge,
    ElementTransformUtils, FamilyInstance, CurveElement, UnitUtils, SpecTypeId, Line,
    FilteredElementCollector, Family, Toposolid, FilledRegion, IntersectionResultArray, SetComparisonResult, CurveLoop,
    Transform, UnitFormatUtils
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB.ExtensibleStorage import SchemaBuilder, Schema, Entity, FieldBuilder, AccessLevel
from System import Guid, Double
from pyrevit import forms, revit, script

# Try to import UIThemeManager (Revit 2024+)
try:
    from Autodesk.Revit.UI import UIThemeManager, UITheme
    HAS_THEME = True
except ImportError:
    HAS_THEME = False
from System.Windows.Media import Colors, SolidColorBrush, Color as WpfColor
from System.Windows.Input import Key
from System.Windows import Visibility

doc = revit.doc
uidoc = revit.uidoc

# ==========================================
# 0. UNIT HELPER
# ==========================================
class UnitHelper:
    @staticmethod
    def get_project_length_unit():
        # returns ForgeTypeId (UnitTypeId)
        return doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()

    @staticmethod
    def get_unit_symbol():
        try:
            # Try to get symbol from FormatOptions
            opts = doc.GetUnits().GetFormatOptions(SpecTypeId.Length)
            if opts.UseDefault:
                # Need to look up default symbol for this unit, hard to do easily in API without label utils
                # Fallback to simple mapping or labelutils
                pass
            
            # Simple fallback based on TypeId string for common cases
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
        units = doc.GetUnits()
        parsed_val = clr.Reference[Double]()
        # Try Revit's native parser first (handles 5' 6", 1500mm, etc.)
        if UnitFormatUtils.TryParse(units, SpecTypeId.Length, str(value_str).strip(), parsed_val):
            return parsed_val.Value
            
        # Fallback to float parsing if pure number given incorrectly
        try:
            val = float(value_str)
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertToInternalUnits(val, unit_id)
        except:
            raise ValueError("Invalid unit format")

    @staticmethod
    def from_internal(value_in_internal_units):
        try:
            val = float(value_in_internal_units)
            unit_id = UnitHelper.get_project_length_unit()
            return UnitUtils.ConvertFromInternalUnits(val, unit_id)
        except: return 0.0

    @staticmethod
    def to_formatted_string(value_in_internal_units):
        try:
            val = float(value_in_internal_units)
            units = doc.GetUnits()
            return UnitFormatUtils.Format(units, SpecTypeId.Length, val, False)
        except: return "0.0"

# ==========================================
# 1. SETTINGS & LOGGING
# ==========================================
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "grading_settings.json")
RECIPE_SCHEMA_GUID = Guid("A4B9C8D2-1234-4567-8901-ABCDEF123456")

# Tolerances
MIN_DIST_TOLERANCE = 0.25 
BOUNDARY_TOLERANCE = 0.1 

class BatchLogger(object):
    """Accumulates messages to display in a single dialog."""
    def __init__(self):
        self._errors = []
        self._infos = []
    
    def error(self, msg, detail=None):
        self._errors.append(str(msg))
        if detail:
            self._errors.append("Details: " + str(detail))
    
    def info(self, msg):
        self._infos.append(str(msg))

    def show(self, title="Grading Report"):
        if not self._errors and not self._infos:
            return

        out = script.get_output()
        if self._errors:
            out.print_html('<strong>--- ERRORS ---</strong>')
            for e in self._errors:
                out.print_html('<div style="color:red;">{}</div>'.format(e))
            out.print_html('<br>')
        
        if self._infos:
            out.print_html('<strong>--- INFO ---</strong>')
            for i in self._infos:
                out.print_html('<div style="color:gray;">{}</div>'.format(i))

def get_id_val(element):
    if not element: return -1
    if not hasattr(element, "Id"): return -1
    try: return element.Id.Value
    except AttributeError: return element.Id.IntegerValue

def get_element_label(element):
    if not element: return "[None]"
    eid = get_id_val(element)
    try:
        name = ""
        if isinstance(element, Toposolid):
            try:
                type_elem = doc.GetElement(element.GetTypeId())
                name = "Toposolid: " + type_elem.Name
            except: name = "Toposolid"
        elif hasattr(element, "Symbol") and element.Symbol:
            name = "{} - {}".format(element.Symbol.FamilyName, element.Name)
        elif hasattr(element, "LineStyle") and element.LineStyle:
            name = "Line - " + element.LineStyle.Name
        elif element.Category:
            name = element.Category.Name
        elif hasattr(element, "Name") and element.Name:
                name += " - " + element.Name
        elif hasattr(element, "Name") and element.Name:
            name = element.Name
        
        if name:
            if len(name) > 35:
                name = name[:32] + "..."
            return "{} ({})".format(name, eid)
    except: pass
    return "ID {}".format(eid)

class GradingRecipe:
    @staticmethod
    def get_schema():
        schema = Schema.Lookup(RECIPE_SCHEMA_GUID)
        if not schema:
            builder = SchemaBuilder(RECIPE_SCHEMA_GUID)
            builder.SetReadAccessLevel(AccessLevel.Public)
            builder.SetWriteAccessLevel(AccessLevel.Public)
            builder.SetSchemaName("OxyzenGradingRecipe")
            builder.AddSimpleField("JsonData", str) 
            schema = builder.Finish()
        return schema

    @staticmethod
    def save_recipe(element, data_dict):
        if not element: return
        try:
            schema = GradingRecipe.get_schema()
            entity = Entity(schema)
            entity.Set("JsonData", json.dumps(data_dict))
            element.SetEntity(entity)
        except: pass

    @staticmethod
    def read_recipe(element):
        if not element: return None
        try:
            schema = GradingRecipe.get_schema()
            entity = element.GetEntity(schema)
            if entity.IsValid():
                return json.loads(entity.Get("JsonData", str))
        except: pass
        return None

def load_settings_from_disk():
    defaults = {
        "width": "6.0", "falloff": "10.0", "grid": "3.0", "slope": "2.0", "mode": "stakes",
        "outlier_tol": "1.0",
        "win_top": "100", "win_left": "100"
    }
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                data = json.load(f)
                if data: defaults.update(data)
    except: pass
    return defaults

def save_state_to_disk(state):
    data = {
        "width": state.width, 
        "falloff": state.falloff, 
        "grid": state.grid, 
        "slope": state.slope_val, 
        "outlier_tol": getattr(state, "outlier_tol", "1.0"),
        "mode": state.mode,
        "apply_offset": state.apply_offset,
        "offset_val": state.offset_val,
        "apply_plan_offset": getattr(state, 'apply_plan_offset', False),
        "plan_offset_val": getattr(state, 'plan_offset_val', "0.0"),
        "plan_offset_dir": getattr(state, 'plan_offset_dir', "Both"),
        "square_ends": state.square_ends,
        "draw_split_lines": getattr(state, 'draw_split_lines', False),
        "win_top": str(state.win_top),
        "win_left": str(state.win_left)
    }
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(data, f)
    except: pass

# ==========================================
# 2. STATE & UI
# ==========================================
class GradingState(object):
    def __init__(self):
        sets = load_settings_from_disk()
        
        self.width = sets.get("width", "6.0")
        self.falloff = sets.get("falloff", "10.0")
        self.grid = sets.get("grid", "3.0")
        self.outlier_tol = sets.get("outlier_tol", "1.0")
        self.slope_val = sets.get("slope", "2.0")
        self.mode = sets.get("mode", "stakes")
        self.square_ends = sets.get("square_ends", False)
        self.apply_offset = sets.get("apply_offset", False)
        self.offset_val = sets.get("offset_val", "0.0")
        self.apply_plan_offset = sets.get("apply_plan_offset", False)
        self.plan_offset_val = sets.get("plan_offset_val", "0.0")
        self.plan_offset_dir = sets.get("plan_offset_dir", "Both")
        self.draw_split_lines = sets.get("draw_split_lines", False)
        self.reset_mode = False
        
        self.win_top = float(sets.get("win_top", "100"))
        self.win_left = float(sets.get("win_left", "100")),
        
        self.start_stake = None
        self.end_stake = None
        self.grading_line = None
        
        self.next_action = None

    @property
    def ready(self):
        # Basic ready check: must have start stake and line.
        # End stake is needed if mode is NOT slope.
        has_start = self.start_stake is not None
        has_line = self.grading_line is not None
        has_end = self.end_stake is not None
        
        if self.mode == "slope":
            return has_start and has_line
        else:
            return has_start and has_end and has_line

class GradingWindow(forms.WPFWindow):
    def __init__(self, state):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.state = state
        self.next_action = None
        
        # Restore Position
        try:
            if self.state.win_top > 0: self.Top = self.state.win_top
            if self.state.win_left > 0: self.Left = self.state.win_left
        except: pass
        
        self.apply_revit_theme()
        self.bind_ui()
        self.setup_events()

    _STATUS_COLORS = {
        "idle":      ((100, 100, 100), (120, 120, 120)),
        "info":      ((100, 100, 100), (120, 120, 120)),
        "success":   (( 16, 124,  16), ( 16, 124,  16)),
        "error":     ((197,  15,  31), (197,  15,  31)),
        "selecting": ((255, 140,   0), (180, 100,   0)),
        "busy":      ((  0, 120, 215), (  0, 100, 200)),
    }

    def set_status(self, msg, level="info"):
        try:
            dot_rgb, txt_rgb = self._STATUS_COLORS.get(level, self._STATUS_COLORS["info"])
            dot_brush = SolidColorBrush(WpfColor.FromRgb(*dot_rgb))
            txt_brush = SolidColorBrush(WpfColor.FromRgb(*txt_rgb))

            self.Lb_Status.Content    = msg
            self.Lb_Status.Foreground = txt_brush

            self.Dot_Status.Fill = dot_brush
            self.Tb_StatusHeader.Text = msg if len(msg) <= 46 else msg[:43] + "..."
        except: pass

    def refresh_ui(self):
        # 1. Fill TextBoxes (Internal -> Display)
        try:
            self.Tb_Width.Text = UnitHelper.to_formatted_string(self.state.width)
            self.Tb_Falloff.Text = UnitHelper.to_formatted_string(self.state.falloff)
            self.Tb_Grid.Text = UnitHelper.to_formatted_string(self.state.grid)
            if hasattr(self, "Tb_OutlierTol"):
                self.Tb_OutlierTol.Text = UnitHelper.to_formatted_string(self.state.outlier_tol)
        except:
            self.Tb_Width.Text = "6.0"
            self.Tb_Falloff.Text = "10.0"
            self.Tb_Grid.Text = "3.0"
            if hasattr(self, "Tb_OutlierTol"):
                self.Tb_OutlierTol.Text = "1.0"
        self.Tb_Slope.Text = str(self.state.slope_val)
        
        # 2. Mode
        if self.state.mode == "slope":
            self.Rb_UseSlope.IsChecked = True
            self.Tb_Slope.IsEnabled = True
        else:
            self.Rb_MatchStakes.IsChecked = True
            self.Tb_Slope.IsEnabled = False

        # 3. Selection indicator rows — dot colour + label text + row background
        dot_set   = SolidColorBrush(WpfColor.FromRgb( 16, 124,  16))
        dot_unset = SolidColorBrush(WpfColor.FromRgb(170, 170, 170))
        bg_set    = SolidColorBrush(WpfColor.FromRgb(240, 255, 240))
        bg_unset  = SolidColorBrush(WpfColor.FromRgb(248, 248, 248))

        if getattr(self, "_is_dark", False):
            dot_set   = SolidColorBrush(WpfColor.FromRgb(100, 220, 100))
            dot_unset = SolidColorBrush(WpfColor.FromRgb( 90,  90,  90))
            bg_set    = SolidColorBrush(WpfColor.FromRgb( 35,  55,  35))
            bg_unset  = SolidColorBrush(WpfColor.FromRgb( 45,  52,  64))

        if self.state.start_stake:
            self.Dot_StartStake.Fill    = dot_set
            self.Row_StartStake.Background = bg_set
            self.Lb_StartStake.Text     = get_element_label(self.state.start_stake)
            self.Lb_StartStake.Foreground = self.FindResource("TextBrush")
        else:
            self.Dot_StartStake.Fill    = dot_unset
            self.Row_StartStake.Background = bg_unset
            self.Lb_StartStake.Text     = "[None selected]"
            self.Lb_StartStake.Foreground = self.FindResource("TextLightBrush")

        if self.state.end_stake:
            self.Dot_EndStake.Fill      = dot_set
            self.Row_EndStake.Background = bg_set
            self.Lb_EndStake.Text       = get_element_label(self.state.end_stake)
            self.Lb_EndStake.Foreground = self.FindResource("TextBrush")
        else:
            self.Dot_EndStake.Fill      = dot_unset
            self.Row_EndStake.Background = bg_unset
            self.Lb_EndStake.Text       = "[None selected]"
            self.Lb_EndStake.Foreground = self.FindResource("TextLightBrush")

        if self.state.grading_line:
            self.Dot_Line.Fill          = dot_set
            self.Row_Line.Background    = bg_set
            self.Lb_Line.Text           = get_element_label(self.state.grading_line)
            self.Lb_Line.Foreground = self.FindResource("TextBrush")
        else:
            self.Dot_Line.Fill          = dot_unset
            self.Row_Line.Background    = bg_unset
            self.Lb_Line.Text           = "[None selected]"
            self.Lb_Line.Foreground = self.FindResource("TextLightBrush")
            
        # 4. Calculated Slope
        slope_info = "Slope: -"
        if self.state.start_stake and self.state.end_stake:
            try:
                if self.state.start_stake.IsValidObject and self.state.end_stake.IsValidObject:
                    p1 = self.state.start_stake.Location.Point
                    p2 = self.state.end_stake.Location.Point
                    
                    d_xy = flatten(p1).DistanceTo(flatten(p2))
                    d_z = p2.Z - p1.Z
                    
                    if d_xy > 0.01:
                        s_pct = (d_z / d_xy) * 100.0
                        d_z_disp = UnitHelper.to_formatted_string(d_z)
                        slope_info = "Slope: {:.2f}% (ΔZ: {})".format(s_pct, d_z_disp)
                    else:
                        slope_info = "Slope: Vertical"
            except: pass
        if hasattr(self, "Lb_CalculatedSlope"):
            self.Lb_CalculatedSlope.Text = slope_info
            
        # 5. Enable/Disable Swap
        self.Btn_Swap.IsEnabled = (self.state.start_stake is not None and self.state.end_stake is not None)
        
        # 6. Options
        self.Cb_ResetPoints.IsChecked = self.state.reset_mode
        if hasattr(self, "Cb_SquareEnds"):
            self.Cb_SquareEnds.IsChecked = self.state.square_ends
        if hasattr(self, "Cb_ApplyOffset"):
            self.Cb_ApplyOffset.IsChecked = self.state.apply_offset
            try: self.Tb_Offset.Text = UnitHelper.to_formatted_string(self.state.offset_val)
            except: self.Tb_Offset.Text = "0.0"
            
        if hasattr(self, "Cb_PlanOffset"):
            self.Cb_PlanOffset.IsChecked = getattr(self.state, "apply_plan_offset", False)
            try:
                self.Tb_PlanOffset.Text = UnitHelper.to_formatted_string(getattr(self.state, "plan_offset_val", "0.0"))
            except:
                self.Tb_PlanOffset.Text = "0.0"
                
            dir_val = getattr(self.state, "plan_offset_dir", "Both")
            for i in range(self.Cmb_PlanOffsetDir.Items.Count):
                item = self.Cmb_PlanOffsetDir.Items[i]
                content = getattr(item, "Content", None)
                if content == dir_val:
                    self.Cmb_PlanOffsetDir.SelectedIndex = i
                    break

        if hasattr(self, "Cb_Triangulate"):
            self.Cb_Triangulate.IsChecked = getattr(self.state, "draw_split_lines", False)

        # Re-run validation so buttons are correct
        self.validate_ui()

    def bind_ui(self):
        u_sym = UnitHelper.get_unit_symbol()
        self.Title += " [{}]".format(u_sym)
        for name in ("Lbl_Width_Unit", "Lbl_Falloff_Unit", "Lbl_Grid_Unit", "Lbl_Offset_Unit", "Lbl_PlanOffset_Unit", "Lbl_OutlierTol_Unit"):
            try: 
                lbl = getattr(self, name)
                lbl.Visibility = Visibility.Collapsed
            except: pass
        self.refresh_ui()

    def apply_revit_theme(self):
        """Detects Revit theme and updates window resources if Dark."""
        self._is_dark = False
        if HAS_THEME:
            try:
                self._is_dark = (UIThemeManager.CurrentTheme == UITheme.Dark)
            except: pass
        
        if self._is_dark:
            res = self.Resources
            res["WindowBrush"]        = SolidColorBrush(WpfColor.FromRgb( 45,  52,  64))
            res["ControlBrush"]       = SolidColorBrush(WpfColor.FromRgb( 35,  41,  51))
            res["TextBrush"]          = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
            res["TextLightBrush"]     = SolidColorBrush(WpfColor.FromRgb(160, 165, 175))
            res["AccentBrush"]        = SolidColorBrush(WpfColor.FromRgb(  0, 120, 215))
            res["HeaderTextBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            res["HeaderSubTextBrush"] = SolidColorBrush(WpfColor.FromRgb(180, 210, 240))
            res["BorderBrush"]        = SolidColorBrush(WpfColor.FromRgb( 75,  85, 100))
            res["StatusReadyBrush"]   = SolidColorBrush(WpfColor.FromRgb( 90, 210,  90))
            res["StatusErrorBrush"]   = SolidColorBrush(WpfColor.FromRgb(240, 100, 100))
            res["ButtonBrush"]        = SolidColorBrush(WpfColor.FromRgb( 60,  68,  82))
            res["HoverBrush"]         = SolidColorBrush(WpfColor.FromRgb( 70,  85, 105))
            res["PressedBrush"]       = SolidColorBrush(WpfColor.FromRgb(  0,  80, 160))
            res["RowUnsetBrush"]      = SolidColorBrush(WpfColor.FromRgb( 45,  52,  64))
            res["RowSetBrush"]        = SolidColorBrush(WpfColor.FromRgb( 30,  55,  35))
            res["DotUnsetBrush"]      = SolidColorBrush(WpfColor.FromRgb( 90,  90,  90))
            res["DotSetBrush"]        = SolidColorBrush(WpfColor.FromRgb( 90, 200,  90))

    def setup_events(self):
        self.Btn_SelectStakes.Click += self.a_stakes
        self.Btn_SelectLine.Click += self.a_line
        self.Btn_Swap.Click += self.a_swap
        self.Btn_Run.Click += self.a_run
        self.Btn_Edging.Click += self.a_edge
        self.Btn_Stitch.Click += self.a_stitch
        try: self.Btn_SmoothRegion.Click += self.a_smooth
        except AttributeError: pass
        self.Btn_ReadRecipe.Click += self.a_load
        try: self.Btn_LinePoints.Click += self.a_line_points
        except AttributeError: pass
        self.Rb_MatchStakes.Checked += self.mode_changed
        self.Rb_UseSlope.Checked += self.mode_changed
        
        # Custom Window Events
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_WinClose.Click += self._on_close_btn
        
        # Validation Events
        tbs = [self.Tb_Width, self.Tb_Falloff, self.Tb_Grid]
        if hasattr(self, "Tb_OutlierTol"): tbs.append(self.Tb_OutlierTol)
        if hasattr(self, "Tb_Offset"): tbs.append(self.Tb_Offset)
        if hasattr(self, "Tb_PlanOffset"): tbs.append(self.Tb_PlanOffset)
        
        for tb in tbs:
            tb.LostFocus += self.format_textbox
            tb.KeyDown += self.format_textbox

    def drag_window(self, sender, args):
        try: self.DragMove()
        except: pass

    def _on_close_btn(self, sender, args):
        try:
            self.state.win_top  = self.Top
            self.state.win_left = self.Left
            save_state_to_disk(self.state)
        except: pass
        self.Close()

    def format_textbox(self, sender, args):
        # If triggered by a key press, only process if it is the Enter key
        if hasattr(args, "Key") and args.Key != Key.Enter:
            return
            
        try:
            if hasattr(sender, "Text"):
                # Parse to internal units (interprets spaces like '50 6')
                val = UnitHelper.to_internal(sender.Text)
                # Format back to display string (e.g. 50' - 6")
                sender.Text = UnitHelper.to_formatted_string(val)
                if hasattr(sender, "BorderBrush"):
                    default_color = WpfColor.FromRgb(75, 85, 100) if getattr(self, "_is_dark", False) else WpfColor.FromRgb(171, 173, 179)
                    sender.BorderBrush = SolidColorBrush(default_color)
        except:
            if hasattr(sender, "BorderBrush"):
                sender.BorderBrush = SolidColorBrush(Colors.Red)
        self.validate_ui()

    def validate_ui(self, sender=None, args=None):
        try:
            w = UnitHelper.to_internal(self.Tb_Width.Text)
            f = UnitHelper.to_internal(self.Tb_Falloff.Text)
            g = UnitHelper.to_internal(self.Tb_Grid.Text)
            if hasattr(self, "Tb_OutlierTol"):
                _ = UnitHelper.to_internal(self.Tb_OutlierTol.Text)
            if hasattr(self, "Tb_Offset"):
                _ = UnitHelper.to_internal(self.Tb_Offset.Text)
            if hasattr(self, "Tb_PlanOffset"):
                _ = UnitHelper.to_internal(self.Tb_PlanOffset.Text)
            
            msg = None
            
            if w <= 0: msg = "Width must be > 0"
            elif f < 0: msg = "Falloff cannot be negative"
            elif g < 0.1: msg = "Grid must be >= 0.1 ft"
            elif g > w: msg = "Grid > Width (Path skipped!)"
            elif f > 0 and g > f: msg = "Grid > Falloff (Jagged!)"
            
            if msg:
                self.set_status(msg, "error")
                self.Btn_Run.IsEnabled = False
                self.Btn_Edging.IsEnabled = False
                return False

            if self.state.ready:
                self.set_status("Ready — click Sculpt Terrain or Run Edging.", "success")
                self.Btn_Run.IsEnabled = True
                self.Btn_Edging.IsEnabled = True
            else:
                self.set_status("Select required inputs (Stakes + Line).", "idle")
                self.Btn_Run.IsEnabled = False
                self.Btn_Edging.IsEnabled = False
            return True
                
        except:
            self.set_status("Invalid Number Format", "error")
            self.Btn_Run.IsEnabled = False
            self.Btn_Edging.IsEnabled = False
            return False

    def set_selection(self, elements):
        try:
            ids = List[ElementId]()
            for e in elements:
                if e and e.IsValidObject: ids.Add(e.Id)
            if ids.Count > 0:
                uidoc.Selection.SetElementIds(ids)
                uidoc.RefreshActiveView()
        except: pass

    def h_stakes_on(self, s, a): self.set_selection([self.state.start_stake, self.state.end_stake])
    def h_line_on(self, s, a): self.set_selection([self.state.grading_line])
    def h_off(self, s, a):
        uidoc.Selection.SetElementIds(List[ElementId]())
        uidoc.RefreshActiveView()

    def mode_changed(self, sender, args):
        self.state.mode = "slope" if self.Rb_UseSlope.IsChecked else "stakes"
        self.Tb_Slope.IsEnabled = (self.state.mode == "slope")
        self.Btn_Run.IsEnabled = self.state.ready
        self.Btn_Edging.IsEnabled = self.state.ready

    def update_state_from_ui(self):
        """Pushes UI values to Memory (converting Display -> Internal)."""
        try: self.state.width = str(UnitHelper.to_internal(self.Tb_Width.Text))
        except: pass
        try: self.state.falloff = str(UnitHelper.to_internal(self.Tb_Falloff.Text))
        except: pass
        try: self.state.grid = str(UnitHelper.to_internal(self.Tb_Grid.Text))
        except: pass
        if hasattr(self, "Tb_OutlierTol"):
            try: self.state.outlier_tol = str(UnitHelper.to_internal(self.Tb_OutlierTol.Text))
            except: pass
        self.state.slope_val = self.Tb_Slope.Text
        self.state.reset_mode = self.Cb_ResetPoints.IsChecked
        if hasattr(self, "Cb_SquareEnds"):
            self.state.square_ends = self.Cb_SquareEnds.IsChecked
        if hasattr(self, "Cb_ApplyOffset"):
            self.state.apply_offset = self.Cb_ApplyOffset.IsChecked
            try: self.state.offset_val = str(UnitHelper.to_internal(self.Tb_Offset.Text))
            except: pass
        if hasattr(self, "Cb_PlanOffset"):
            self.state.apply_plan_offset = self.Cb_PlanOffset.IsChecked
            try: self.state.plan_offset_val = str(UnitHelper.to_internal(self.Tb_PlanOffset.Text))
            except: pass
            if self.Cmb_PlanOffsetDir.SelectedItem:
                self.state.plan_offset_dir = getattr(self.Cmb_PlanOffsetDir.SelectedItem, "Content", "Both")
            
        if hasattr(self, "Cb_Triangulate"):
            self.state.draw_split_lines = self.Cb_Triangulate.IsChecked

    def _raise(self, action):
        self.update_state_from_ui()
        self.next_action = action
        self.Close()

    def a_stakes(self, s, a): self._raise("select_stakes")
    def a_line(self, s, a): self._raise("select_line")
    def a_swap(self, s, a): self._raise("swap")
    def a_run(self, s, a): self._raise("sculpt")
    def a_edge(self, s, a): self._raise("edge")
    def a_stitch(self, s, a): self._raise("stitch")
    def a_smooth(self, s, a): self._raise("smooth")
    def a_line_points(self, s, a): self._raise("line_points")
    def a_load(self, s, a): self._raise("load_recipe")

# ==========================================
# 3. HELPERS
# ==========================================
class UniversalFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, r, p): return True

def flatten(pt): return XYZ(pt.X, pt.Y, 0)
def lerp(a, b, t): return a + t * (b - a)

def project_2d(curve, pt):
    """
    Projects a 3D point onto a curve in the XY plane.
    Returns: (closest_3d_pt, normalized_param, distance_xy)
    """
    if isinstance(curve, Line):
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        p0_flat = flatten(p0)
        p1_flat = flatten(p1)
        v_line = p1_flat - p0_flat
        len_sq = v_line.X**2 + v_line.Y**2
        
        if len_sq < 0.0001:
            d = flatten(pt).DistanceTo(p0_flat)
            return p0, 0.0, d
            
        v_pt = flatten(pt) - p0_flat
        t = (v_pt.X * v_line.X + v_pt.Y * v_line.Y) / len_sq
        t_bounded = max(0.0, min(1.0, t))
        
        pt_3d = p0 + (p1 - p0) * t_bounded
        d = flatten(pt).DistanceTo(flatten(pt_3d))
        return pt_3d, t_bounded, d
    else:
        best_pt = None; best_dist = float('inf'); best_t = 0.0
        p_min, p_max = curve.GetEndParameter(0), curve.GetEndParameter(1)
        steps = 50
        for i in range(steps + 1):
            t_norm = i / float(steps)
            t_raw = p_min + t_norm * (p_max - p_min)
            p_3d = curve.Evaluate(t_raw, False)
            d = flatten(pt).DistanceTo(flatten(p_3d))
            if d < best_dist:
                best_dist = d; best_pt = p_3d; best_t = t_norm
        return best_pt, best_t, best_dist

def get_toposolid_max_z(toposolid):
    bb = toposolid.get_BoundingBox(None)
    if bb: return bb.Max.Z
    return 1000.0 

def is_point_on_solid(intersector, pt, start_z):
    if not intersector: return True 
    origin = XYZ(pt.X, pt.Y, start_z + 10.0) 
    try:
        if intersector.FindNearest(origin, XYZ(0, 0, -1)): return True
    except: pass
    return False

def is_too_close(candidate_pt, occupied_points, tolerance=MIN_DIST_TOLERANCE):
    cand_flat = flatten(candidate_pt)
    for existing in occupied_points:
        if cand_flat.DistanceTo(flatten(existing)) < tolerance: return True
    return False

def get_surface_info(intersector, pt, start_z):
    """Returns (Z_value, ElementId) of the surface hit at pt (x,y)."""
    if not intersector: return (None, None)
    origin = XYZ(pt.X, pt.Y, start_z + 50.0) # Start high
    try:
        # FindNearest finds the first hit
        context = intersector.FindNearest(origin, XYZ(0, 0, -1))
        if context: 
            ref = context.GetReference()
            return (ref.GlobalPoint.Z, ref.ElementId)
    except: pass
    return (None, None)

def get_subdivision_offset(doc, elem_id):
    """Checks if the element is a subdivision and returns its height parameter."""
    if not elem_id or elem_id == ElementId.InvalidElementId: return 0.0
    try:
        el = doc.GetElement(elem_id)
        # Check if it has HostTopoId (Revit 2024+)
        if hasattr(el, "HostTopoId") and el.HostTopoId != ElementId.InvalidElementId:
            p = el.get_Parameter(BuiltInParameter.TOPOSOLID_SUBDIV_HEIGHT)
            if p: return p.AsDouble()
    except: pass
    return 0.0

def get_boundary_curves(toposolid):
    opt = Options(); opt.ComputeReferences = True
    geom = toposolid.get_Geometry(opt)
    curves = []
    if not geom: return []
    for obj in geom:
        if isinstance(obj, Solid):
            for edge in obj.Edges:
                c = edge.AsCurve()
                if c: curves.append(c)
    return curves

def get_line_ends(curve):
    return curve.GetEndPoint(0), curve.GetEndPoint(1)

def resolve_toposolid_host(doc, element):
    """
    Checks if the selected Toposolid is a Subdivision.
    If so, returns the Host Toposolid (the one with the points).
    """
    try:
        # Revit 2024+ Property for Toposolid Subdivision Host
        if hasattr(element, "HostTopoId"):
            hid = element.HostTopoId
            if hid and hid != ElementId.InvalidElementId:
                host = doc.GetElement(hid)
                if host: return host
    except: pass
    return element

def validate_input(state, log):
    """Failsafe check for basic inputs and parameter logic before processing."""
    # 1. Check Elements
    if not state.start_stake:
        log.error("Start Stake is missing.")
        return False
    if not state.grading_line:
        log.error("Grading Line is missing.")
        return False
    
    try:
        if not state.start_stake.IsValidObject:
            log.error("Start Stake element is no longer valid.")
            return False
        if not state.grading_line.IsValidObject:
            log.error("Grading Line element is no longer valid.")
            return False
    except:
        log.error("Invalid element reference.")
        return False

    # 2. Check Parameters (Logic)
    try:
        w = float(state.width)
        f = float(state.falloff)
        g = float(state.grid)
        
        # A. Basic Bounds
        if w <= 0:
            log.error("Path Width must be greater than 0.")
            return False
        if f < 0:
            log.error("Falloff cannot be negative.")
            return False
        if g <= 0.01:
            log.error("Grid resolution must be greater than 0.")
            return False
            
        # B. Performance Safety
        if g < 0.1: # 0.1 ft ~= 30mm
            g_str = UnitHelper.to_formatted_string(g)
            log.error("Grid resolution is dangerously small ({}).".format(g_str), 
                      "This will likely freeze or crash Revit. Please use a value >= 0.25 ft.")
            return False
            
        # C. Geometric Logic
        # If the grid is bigger than the width, we might skip the whole path!
        if g > w:
            g_str = UnitHelper.to_formatted_string(g)
            w_str = UnitHelper.to_formatted_string(w)
            log.error("Grid resolution ({}) is larger than Path Width ({}).".format(g_str, w_str), 
                      "The grading points might jump over your path completely. Reduce Grid or increase Width.")
            return False
            
        # If grid is bigger than falloff, smoothing will look jagged or do nothing
        if f > 0 and g > f:
            g_str = UnitHelper.to_formatted_string(g)
            f_str = UnitHelper.to_formatted_string(f)
            log.error("Grid resolution ({}) is larger than Falloff ({}).".format(g_str, f_str),
                      "Smoothing will be ineffective. Reduce Grid size.")
            # We allow this but warn, or strictly block if desired. 
            # Blocking is safer for "useless" results.
            return False
            
    except ValueError:
        log.error("One or more parameters (Width, Falloff, Grid) are not valid numbers.")
        return False

    return True

def calculate_and_adjust_stakes(state, log):
    """Calculates Z levels and adjusts end stake if needed."""
    if not state.grading_line or not isinstance(state.grading_line, CurveElement):
        log.error("Invalid Grading Line selected.")
        raise Exception("Invalid Line")

    curve = state.grading_line.GeometryCurve
    l_start, l_end = get_line_ends(curve)
    
    # Failsafe: Stake might not have a LocationPoint if user picked wrong family
    if not hasattr(state.start_stake, "Location") or not hasattr(state.start_stake.Location, "Point"):
        log.error("Start Stake does not have a valid location point.")
        raise Exception("Invalid Stake")

    u_start_pt = state.start_stake.Location.Point
    z_start = u_start_pt.Z
    dist_start = u_start_pt.DistanceTo(l_start)
    dist_end = u_start_pt.DistanceTo(l_end)
    is_flipped = dist_end < dist_start 
    length = curve.Length
    z_end = 0.0
    
    if state.mode == "slope":
        try:
            pct = float(state.slope_val) / 100.0
            z_end = z_start + (length * pct)
            if state.end_stake:
                t_move = Transaction(doc, "Adjust Stake Height")
                t_move.Start()
                try:
                    current_pt = state.end_stake.Location.Point
                    diff_z = z_end - current_pt.Z
                    if abs(diff_z) > 0.001:
                        vec = XYZ(0, 0, diff_z)
                        ElementTransformUtils.MoveElement(doc, state.end_stake.Id, vec)
                except:
                    # Fallback for families constrained to host
                    try:
                        p = state.end_stake.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                        if p and not p.IsReadOnly: p.Set(z_end)
                    except: pass
                t_move.Commit()
        except Exception as e:
            log.error("Failed to adjust End Stake slope.", e)
            z_end = z_start
    else:
        if not state.end_stake or not hasattr(state.end_stake.Location, "Point"):
             log.error("End Stake is missing or invalid for 'Match Stakes' mode.")
             raise Exception("Invalid End Stake")
        z_end = state.end_stake.Location.Point.Z
        
    return (z_end, z_start) if is_flipped else (z_start, z_end)

# ==========================================
# 5. EXECUTION
# ==========================================
def perform_load_recipe(state):
    log = BatchLogger()
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        if not ref: return
        data = GradingRecipe.read_recipe(doc.GetElement(ref))
        if data:
            state.width = str(data.get("width", "6.0"))
            state.falloff = str(data.get("falloff", "10.0"))
            state.grid = str(data.get("grid", "3.0"))
            state.outlier_tol = str(data.get("outlier_tol", "1.0"))
            state.slope_val = str(data.get("slope", "2.0"))
            state.mode = str(data.get("mode", "stakes"))
            state.apply_offset = data.get("apply_offset", False)
            state.offset_val = str(data.get("offset_val", "0.0"))
            state.apply_plan_offset = data.get("apply_plan_offset", False)
            state.plan_offset_val = str(data.get("plan_offset_val", "0.0"))
            state.plan_offset_dir = data.get("plan_offset_dir", "Both")
            state.square_ends = data.get("square_ends", False)
            state.draw_split_lines = data.get("draw_split_lines", False)
            log.info("Recipe loaded successfully.")
        else:
            log.info("No grading recipe found on this element.")
    except Exception as e:
        log.error("Failed to load recipe.", e)
    finally:
        log.show()

def get_chain_of_edges(toposolid, start_edge):
    """
    Finds the boundary loop.
    Priority 1: Toposolid Sketch (Base Curves).
    Priority 2: Geometry Edge Graph (Filtered).
    """
    
    # --- STRATEGY 1: SKETCH (Base Curves) ---
    try:
        sketch = None
        # 1. Try SketchId property
        if hasattr(toposolid, "SketchId") and toposolid.SketchId != ElementId.InvalidElementId:
            sketch = doc.GetElement(toposolid.SketchId)
        
        # 2. Try Dependent Elements if property fails
        if not sketch:
            ids = toposolid.GetDependentElements(UniversalFilter()) # Re-use generic filter or None
            for eid in ids:
                el = doc.GetElement(eid)
                if el and "Sketch" in el.GetType().Name:
                    sketch = el
                    break
        
        if sketch and sketch.Profile:
            # Found Sketch. Find the loop matching selected edge.
            
            # Midpoint of selected edge for proximity check
            sel_curve = start_edge.AsCurve()
            mid_pt = sel_curve.Evaluate(0.5, True)
            mid_flat = flatten(mid_pt)
            
            best_loop = None
            min_dist = 1.0 # Tolerance ft
            
            # Iterate CurveArrArray (Profile is list of loops)
            for curve_array in sketch.Profile:
                # Check if this loop is "close" to our selection
                is_match = False
                for i in range(curve_array.Size):
                    sc = curve_array.get_Item(i)
                    
                    # Project selected midpoint to sketch curve (2D check)
                    # We flatten the sketch curve ends to check distance
                    sp0 = flatten(sc.GetEndPoint(0))
                    sp1 = flatten(sc.GetEndPoint(1))
                    
                    # Simple distance to segment check
                    # Or assume sketch curve is planar Z-flat, just ignore Z
                    # We can use our 'flatten' to create a new Line for check? 
                    # No, generic curve might be arc.
                    # Let's project 'mid_flat' onto 'sc' ignoring Z?
                    # Hard to do generically without creating geometry.
                    
                    # Quick check: Is mid_flat close to endpoints?
                    if mid_flat.DistanceTo(sp0) < 5.0 or mid_flat.DistanceTo(sp1) < 5.0:
                        # Closer check
                        pass

                    # Robust 2D Project:
                    # Create flat version of sc?
                    # Since sc is likely flat, just setting Z to 0 might work.
                    try:
                        # Only checking distance to "unbounded" curve might give false positives
                        # But Profile curves are bounded.
                        # Let's just check endpoints for now as a heuristic 
                        # OR check if mid_pt projects onto it.
                        res = sc.Project(mid_pt)
                        if res:
                            # Distance in XY plane
                            p_res = res.XYZPoint
                            d_xy = flatten(p_res).DistanceTo(mid_flat)
                            if d_xy < min_dist:
                                is_match = True
                                break
                    except: pass
                
                if is_match:
                    best_loop = []
                    for i in range(curve_array.Size):
                        best_loop.append(curve_array.get_Item(i))
                    return best_loop

    except Exception as e:
        # Log failure silently or to debug, fallback to method 2
        # print("Sketch lookup failed: " + str(e))
        pass

    # --- STRATEGY 2: GEOMETRY GRAPH (Fallback) ---
    opt = Options()
    opt.ComputeReferences = True
    geom = toposolid.get_Geometry(opt)
    
    all_edges = []
    if geom:
        for obj in geom:
            if isinstance(obj, Solid):
                for e in obj.Edges:
                    all_edges.append(e)
    
    # 1. Build Adjacency Graph (Endpoint -> List of Edges)
    # Key: (X, Y, Z) rounded tuple
    # Value: List of Edge objects
    adj_map = {}
    
    def pt_key(pt):
        return (round(pt.X, 4), round(pt.Y, 4), round(pt.Z, 4))
    
    # Helper to check if an edge is internal (connects two top faces)
    def is_boundary_candidate(edge):
        try:
            # Get the two faces sharing this edge
            f0 = edge.GetFace(0)
            f1 = edge.GetFace(1)
            
            if not f0 or not f1: return True # Keep if we can't determine (safe fallback) 
            
            def is_up(face):
                # Using face.ComputeNormal at UV center is robust for Planar/Ruled faces
                bbox = face.GetBoundingBox()
                center = (bbox.Min + bbox.Max) / 2.0
                try:
                    if hasattr(face, "FaceNormal"): return face.FaceNormal.Z > 0.1
                    res = face.Project(center)
                    if res:
                        norm = face.ComputeNormal(res.UVPoint)
                        return norm.Z > 0.1
                except: pass
                return False 
                
            up0 = is_up(f0)
            up1 = is_up(f1)
            
            if up0 and up1: return False
            return True
        except:
            return True 
            
    start_id = start_edge.Id
    
    filtered_count = 0
    for e in all_edges:
        c = e.AsCurve()
        if not c: continue
        
        # Optimization: Always include the user-selected edge without checking
        if e.Id != start_id:
            if not is_boundary_candidate(e):
                continue
        
        filtered_count += 1
        
        p0 = pt_key(c.GetEndPoint(0))
        p1 = pt_key(c.GetEndPoint(1))
        
        if p0 not in adj_map: adj_map[p0] = []
        if p1 not in adj_map: adj_map[p1] = []
        
        adj_map[p0].append(e)
        adj_map[p1].append(e)
        
    # 2. Traverse Graph using BFS/DFS
    visited_ids = set()
    chain = []
    stack = [start_edge]
    visited_ids.add(start_id)
    
    while stack:
        current_edge = stack.pop(0) # BFS
        chain.append(current_edge)
        
        c = current_edge.AsCurve()
        p0 = pt_key(c.GetEndPoint(0))
        p1 = pt_key(c.GetEndPoint(1))
        
        # Check neighbors at both ends
        for p in [p0, p1]:
            neighbors = adj_map.get(p, [])
            for n_edge in neighbors:
                nid = n_edge.Id
                if nid not in visited_ids:
                    visited_ids.add(nid)
                    stack.append(n_edge)

    return [e.AsCurve() for e in chain] 

def perform_manual_stitch(state):
    log = BatchLogger()
    log.info("--- MANUAL STITCH STARTED ---")
    
    try:
        # Use Grid size as the search tolerance/radius
        snap_dist = float(state.grid)
        log.info("Snap Tolerance (Grid): {:.3f} ft".format(snap_dist))
        
        if snap_dist <= 0: snap_dist = 1.0

        g_plan_off = 0.0
        if getattr(state, "apply_plan_offset", False):
            try:
                g_plan_off = float(getattr(state, "plan_offset_val", "0.0"))
                if g_plan_off != 0.0:
                    log.info("Applied Plan Offset: {}".format(UnitHelper.to_formatted_string(g_plan_off)))
            except: pass

        try:
            ref_edge = uidoc.Selection.PickObject(ObjectType.Edge, "Select Boundary Edge to Stitch")
        except: 
            return # Cancelled
            
        edge_elem = doc.GetElement(ref_edge)
        edge_geom = edge_elem.GetGeometryObjectFromReference(ref_edge)
        
        if not isinstance(edge_geom, Edge): 
            log.error("Selected object is not a valid Edge.")
            log.show(); return

        toposolid = resolve_toposolid_host(doc, edge_elem)
        if toposolid.Id != edge_elem.Id:
            log.info("Switching to Host Toposolid ID: {}".format(toposolid.Id))
        else:
            log.info("Selected Toposolid ID: {}".format(toposolid.Id))
        
        # NEW: Get the full chain of connected edges
        log.info("Tracing connected boundary edges...")
        boundary_curves = get_chain_of_edges(toposolid, edge_geom)
        log.info("Identified {} connected boundary segments.".format(len(boundary_curves)))
        
        if not boundary_curves:
            log.error("Could not find any boundary edges. Check geometry.")
            log.show(); return

        tg = TransactionGroup(doc, "Stitch Edge")
        tg.Start()
        
        t = Transaction(doc, "Stitch")
        t.Start()
        
        try:
            editor = toposolid.GetSlabShapeEditor()
            editor.Enable()
            
            existing_coords = set()
            candidates = []
            
            # Cache existing vertices
            for v in editor.SlabShapeVertices:
                pos = v.Position
                existing_coords.add((round(pos.X, 4), round(pos.Y, 4)))
                candidates.append(v)
            
            log.info("Internal Vertices to Check: {}".format(len(candidates)))
            
            points_to_add = []
            
            # 2. Find Candidates: Internal Points -> Project -> Closest Edge in Chain
            check_count = 0
            match_count = 0
            
            for v in candidates:
                check_count += 1
                best_proj = None
                best_dist = 9999.0
                
                # Iterate all segments in the boundary chain to find the closest one
                # Optimization: Could use a spatial index, but N_edges is usually small (<100)
                for b_curve in boundary_curves:
                    try:
                        proj_res = b_curve.Project(v.Position)
                        if not proj_res: continue
                        
                        p_target = proj_res.XYZPoint
                        v_flat = flatten(v.Position)
                        p_target_flat = flatten(p_target)
                        dist = v_flat.DistanceTo(p_target_flat)
                        
                        if dist < best_dist:
                            best_dist = dist
                            best_proj = p_target
                    except: pass
                
                # Logic: If close enough (but not ON the edge), propose new point
                if best_proj and (0.005 < best_dist < snap_dist):
                    
                    offset_pt = best_proj
                    if g_plan_off != 0.0:
                        v_flat = flatten(v.Position)
                        p_target_flat = flatten(best_proj)
                        vec_to_v = v_flat - p_target_flat
                        if not vec_to_v.IsAlmostEqualTo(XYZ.Zero):
                            dir_outward = vec_to_v.Normalize().Negate()
                            offset_pt = p_target_flat + (dir_outward * g_plan_off)
                            
                    new_pt = XYZ(offset_pt.X, offset_pt.Y, v.Position.Z)
                    
                    key = (round(new_pt.X, 4), round(new_pt.Y, 4))
                    if key not in existing_coords:
                        points_to_add.append(new_pt)
                        existing_coords.add(key)
                        match_count += 1
            
            log.info("Candidates Found within Tolerance: {}".format(match_count))
            
            # 3. Add Points
            added_count = 0
            for p in points_to_add:
                try:
                    editor.AddPoint(p)
                    added_count += 1
                except Exception as e:
                    pass
            
            log.info("Successfully Added Points: {}".format(added_count))

            t.Commit()
            tg.Assimilate()
            
        except Exception as e:
            t.RollBack()
            tg.RollBack()
            log.error("Stitch transaction failed.", "{}\n{}".format(e, traceback.format_exc()))
            
    except Exception as e:
        log.error("Critical error in stitch tool.", "{}\n{}".format(e, traceback.format_exc()))
    finally:
        log.show()
def perform_swap(state):
    if state.start_stake and state.end_stake:
        state.start_stake, state.end_stake = state.end_stake, state.start_stake

def perform_sculpt(state):
    log = BatchLogger()
    if not validate_input(state, log):
        log.show()
        return

    # Safety Check
    if doc.IsModifiable:
        log.error("Another transaction is currently active.", "Please finish or cancel the active command before running this tool.")
        log.show()
        return

    try:
        # 1. Parse Inputs & Log
        w_int = float(state.width)
        f_int = float(state.falloff)
        g_int = float(state.grid)
        
        log.info("--- SCULPT STARTED ---")
        log.info("Parameters: Width={}, Falloff={}, Grid={}".format(
            UnitHelper.to_formatted_string(w_int), 
            UnitHelper.to_formatted_string(f_int), 
            UnitHelper.to_formatted_string(g_int)
        ))
        
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid to Grade")
            raw_elem = doc.GetElement(ref)
            toposolid = resolve_toposolid_host(doc, raw_elem)

            if toposolid.Id != raw_elem.Id:
                log.info("Selected Element is a Subdivision. Targeting Host Toposolid ID: {}".format(toposolid.Id))
            else:
                log.info("Selected Toposolid ID: {}".format(toposolid.Id))

        except:
            return # Cancelled
            
        if not toposolid:
            log.error("Invalid Toposolid selection.")
            log.show(); return

    except Exception as e:
        log.error("Invalid parameter inputs.", e)
        log.show(); return

    tg = TransactionGroup(doc, "Sculpt Terrain")
    tg.Start()
    
    try:
        # Save Recipe
        rec = {
            "width": state.width, 
            "falloff": state.falloff, 
            "grid": state.grid,
            "outlier_tol": getattr(state, "outlier_tol", "1.0"),
            "slope": state.slope_val,
            "apply_offset": state.apply_offset,
            "offset_val": state.offset_val,
            "apply_plan_offset": getattr(state, 'apply_plan_offset', False),
            "plan_offset_val": getattr(state, 'plan_offset_val', "0.0"),
            "plan_offset_dir": getattr(state, 'plan_offset_dir', "Both"),
            "mode": state.mode,
            "draw_split_lines": getattr(state, "draw_split_lines", False)
        }
        
        t_rec = Transaction(doc, "Save Recipe")
        t_rec.Start()
        try:
            GradingRecipe.save_recipe(toposolid, rec)
            t_rec.Commit()
        except: t_rec.RollBack()

        # Calculation
        z_s, z_e = calculate_and_adjust_stakes(state, log)
        if state.apply_offset:
            try:
                g_z_off = float(state.offset_val)
                z_s += g_z_off
                z_e += g_z_off
                log.info("Applied Z-Offset: {}".format(UnitHelper.to_formatted_string(g_z_off)))
            except: pass
        log.info("Stake Elevations: Start={:.2f}, End={:.2f}".format(z_s, z_e))
        
        # Initialize Intersector for Raycasting Context
        # We need this even if 2D, but for Z checks we need 3D view usually.
        # Ensure we look for ALL Toposolids (Host + Subdivisions)
        intersector = None
        ray_start_z = get_toposolid_max_z(toposolid) + 100.0
        
        view3d = doc.ActiveView if doc.ActiveView.ViewType == ViewType.ThreeD else None
        if not view3d:
            cols = FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements()
            for v in cols:
                if not v.IsTemplate and v.Name == "{3D}":
                    view3d = v
                    break
            if not view3d and cols:
                for v in cols:
                    if not v.IsTemplate:
                        view3d = v
                        break
                        
        if view3d:
            col = FilteredElementCollector(doc).OfClass(Toposolid).WhereElementIsNotElementType()
            ids = List[ElementId]([e.Id for e in col])
            intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, view3d)
        else:
            log.info("Warning: No 3D view found. Raycasting disabled. Using path elevation for new points.")

        curve = state.grading_line.GeometryCurve
        log.info("Guide Line Length: {:.2f}".format(curve.Length))

        core_rad = w_int / 2.0
        total_rad = core_rad + f_int
        
        # Pre-calc for Square Ends
        check_square_ends = False
        sq_start_pt = None
        sq_start_tan = None
        sq_end_pt = None
        sq_end_tan = None

        if state.square_ends:
            try:
                sq_start_pt = curve.GetEndPoint(0)
                sq_end_pt = curve.GetEndPoint(1)
                
                if isinstance(curve, Line):
                    vec = (sq_end_pt - sq_start_pt)
                    vec_xy = XYZ(vec.X, vec.Y, 0)
                    if not vec_xy.IsZeroLength():
                        tan = vec_xy.Normalize()
                        sq_start_tan = tan
                        sq_end_tan = tan
                        check_square_ends = True
                else:
                    t0 = curve.GetEndParameter(0)
                    tan0 = curve.ComputeDerivatives(t0, False).BasisX
                    tan0_xy = XYZ(tan0.X, tan0.Y, 0)
                    
                    t1 = curve.GetEndParameter(1)
                    tan1 = curve.ComputeDerivatives(t1, False).BasisX
                    tan1_xy = XYZ(tan1.X, tan1.Y, 0)

                    if not tan0_xy.IsZeroLength() and not tan1_xy.IsZeroLength():
                        sq_start_tan = tan0_xy.Normalize()
                        sq_end_tan = tan1_xy.Normalize()
                        check_square_ends = True
                
                if check_square_ends:
                    log.info("Square Ends Active.")
            except Exception as e:
                log.error("Square Ends Calc Failed", e)

        def is_outside_bounds(pt):
            if not check_square_ends: return False
            v_s = XYZ(pt.X - sq_start_pt.X, pt.Y - sq_start_pt.Y, 0)
            if v_s.DotProduct(sq_start_tan) < -0.001: return True
            v_e = XYZ(pt.X - sq_end_pt.X, pt.Y - sq_end_pt.Y, 0)
            if v_e.DotProduct(sq_end_tan) > 0.001: return True
            return False

        # Phase 0: Reset (Optional)
        if state.reset_mode:
            log.info("--- PHASE 0: RESET POINTS ---")
            t0 = Transaction(doc, "Reset Points")
            t0.Start()
            try:
                editor = toposolid.GetSlabShapeEditor()
                editor.Enable()
                to_delete = []
                for v in editor.SlabShapeVertices:
                    p_3d, t_norm, d_xy = project_2d(curve, v.Position)
                    if d_xy < (total_rad + g_int):
                        to_delete.append(v)
                
                if to_delete:
                    count_del = 0
                    for v_del in to_delete:
                        try:
                            editor.DeletePoint(v_del)
                            count_del += 1
                        except: pass
                    log.info("Removed {} old points to clear resolution.".format(count_del))
                else:
                    log.info("No points found within grading zone to remove.")
                t0.Commit()
            except Exception as e:
                t0.RollBack()
                log.error("Failed to reset points.", e)

        # Phase 1: Densify
        log.info("--- PHASE 1: DENSIFY ---")
        t1 = Transaction(doc, "Densify")
        t1.Start()
        try:
            editor = toposolid.GetSlabShapeEditor()
            editor.Enable()
            occupied_points = [v.Position for v in editor.SlabShapeVertices]
            log.info("Initial Vertex Count: {}".format(len(occupied_points)))
            
            bb = state.grading_line.get_BoundingBox(None)
            start_x = math.floor((bb.Min.X - total_rad - g_int) / g_int) * g_int
            end_x   = math.ceil((bb.Max.X + total_rad + g_int) / g_int) * g_int
            start_y = math.floor((bb.Min.Y - total_rad - g_int) / g_int) * g_int
            end_y   = math.ceil((bb.Max.Y + total_rad + g_int) / g_int) * g_int
            
            log.info("Grid Search Bounds: X[{:.1f}, {:.1f}] Y[{:.1f}, {:.1f}]".format(start_x, end_x, start_y, end_y))
            
            grid_pts = []
            candidates_checked = 0
            x = start_x
            while x <= end_x:
                y = start_y
                while y <= end_y:
                    candidates_checked += 1
                    t_pt = XYZ(x, y, 0)
                    if is_outside_bounds(t_pt): y += g_int; continue

                    p_3d, t_norm, d_xy = project_2d(curve, t_pt)
                    if d_xy < (total_rad + g_int):
                        if is_point_on_solid(intersector, t_pt, ray_start_z):
                            rz, hit_id = get_surface_info(intersector, t_pt, ray_start_z)
                            off = 0.0
                            if rz is not None: 
                                off = get_subdivision_offset(doc, hit_id)
                            else:
                                rz = z_s + t_norm * (z_e - z_s) # Fallback elevation
                            grid_pts.append(XYZ(x, y, rz - off))
                    y += g_int
                x += g_int
                
            log.info("Grid Points Found on Solid: {} (out of {} checked)".format(len(grid_pts), candidates_checked))
            
            added_count = 0
            for p in grid_pts:
                if not is_too_close(p, occupied_points):
                    try: 
                        editor.AddPoint(p)
                        occupied_points.append(p)
                        added_count += 1
                    except: pass
            t1.Commit()
            log.info("Points Added: {}".format(added_count))
        except:
            t1.RollBack()
            raise

        # Phase 2: Sculpt
        log.info("--- PHASE 2: SCULPT ---")
        t2 = Transaction(doc, "Sculpt")
        t2.Start()
        modified_count = 0
        try:
            editor = toposolid.GetSlabShapeEditor()
            editor.Enable()
            updates = []
            current_verts = [v for v in editor.SlabShapeVertices]
            log.info("Total Vertices to Process: {}".format(len(current_verts)))
            sample_log = []
            
            for v in current_verts:
                if is_outside_bounds(v.Position): continue
                
                p_3d, t_norm, d_xy = project_2d(curve, v.Position)
                if d_xy > total_rad: continue
                
                rz, hit_id = get_surface_info(intersector, v.Position, ray_start_z)
                off = 0.0
                if rz is not None:
                     off = get_subdivision_offset(doc, hit_id)
                
                # Current State
                current_base_z = v.Position.Z
                current_top_z = current_base_z + off
                
                target_top_z = z_s + t_norm * (z_e - z_s)
                
                new_base_z = current_base_z
                is_core = d_xy <= core_rad

                if is_core: 
                    new_base_z = target_top_z - off
                else:
                    t_val = (d_xy - core_rad) / f_int
                    t_val = 1.0 if t_val > 1.0 else t_val
                    smooth_t = t_val * t_val * (3 - 2 * t_val)
                    # Blend TOP surfaces
                    desired_top_z = lerp(target_top_z, current_top_z, smooth_t)
                    # Convert back to BASE
                    new_base_z = desired_top_z - off
                
                if abs(new_base_z - v.Position.Z) > 0.005:
                    updates.append(XYZ(v.Position.X, v.Position.Y, new_base_z))
                    if len(sample_log) < 3:
                        sample_log.append("Pt ({:.1f}, {:.1f}): Z {:.2f} -> {:.2f} (Off={:.2f})".format(
                            v.Position.X, v.Position.Y, v.Position.Z, new_base_z, off
                        ))
            
            modified_count = len(updates)
            log.info("Points Identified for Modification: {}".format(modified_count))
            if sample_log:
                log.info("Sample Changes:\n" + "\n".join(sample_log))
            
            for p in updates:
                try: editor.AddPoint(p)
                except: pass
            
            t2.Commit()
        except:
            t2.RollBack()
            raise

        # Phase 3: Triangulate (Split Lines)
        if getattr(state, 'draw_split_lines', False):
            log.info("--- PHASE 3: TRIANGULATE ---")
            t3 = Transaction(doc, "Triangulate Path")
            t3.Start()
            split_lines_count = 0
            try:
                editor = toposolid.GetSlabShapeEditor() 
                editor.Enable()
                
                all_verts = [v for v in editor.SlabShapeVertices]
                search_tol = g_int * 0.25
                step_len = g_int
                if step_len < 0.1: step_len = 1.0
                
                l_len = curve.Length
                current_len = 0.0
                
                while current_len <= l_len:
                    norm_param = current_len / l_len
                    if norm_param > 1.0: norm_param = 1.0
                    
                    center_pt = curve.Evaluate(norm_param, True)
                    deriv = curve.ComputeDerivatives(norm_param, True)
                    tangent = deriv.BasisX.Normalize()
                    normal = tangent.CrossProduct(XYZ.BasisZ) 
                    
                    p_left = center_pt + (normal * core_rad)
                    p_right = center_pt - (normal * core_rad) 
                    
                    v_left = None
                    v_right = None
                    
                    min_d_l = search_tol
                    min_d_r = search_tol
                    
                    p_l_flat = flatten(p_left)
                    p_r_flat = flatten(p_right)
                    
                    for v in all_verts:
                        v_flat = flatten(v.Position)
                        d_l = v_flat.DistanceTo(p_l_flat)
                        if d_l < min_d_l:
                            min_d_l = d_l
                            v_left = v
                        d_r = v_flat.DistanceTo(p_r_flat)
                        if d_r < min_d_r:
                            min_d_r = d_r
                            v_right = v
                    
                    if v_left and v_right and not v_left.Position.IsAlmostEqualTo(v_right.Position):
                        dist_real = v_left.Position.DistanceTo(v_right.Position)
                        if abs(dist_real - w_int) < (w_int * 0.5):
                            try:
                                editor.DrawSplitLine(v_left, v_right)
                                split_lines_count += 1
                            except: pass
                    
                    current_len += step_len

                t3.Commit()
                log.info("Triangulation Complete. Split Lines Added: {}".format(split_lines_count))
            except:
                t3.RollBack()
                raise

        tg.Assimilate()
        log.info("Sculpt Transaction Committed.")
        log.info("Sculpt Complete.\nPoints Added: {}\nPoints Adjusted: {}".format(added_count, modified_count))
        
        state.reset_mode = False

    except Exception as e:
        tg.RollBack()
        log.error("Sculpting Failed", "{}\n{}".format(e, traceback.format_exc()))
    
    finally:
        log.show()

def perform_edging(state):
    log = BatchLogger()
    if not validate_input(state, log):
        log.show(); return

    try:
        w = float(state.width); g = float(state.grid)
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid for Edging")
            raw_elem = doc.GetElement(ref)
            toposolid = resolve_toposolid_host(doc, raw_elem)
            if toposolid.Id != raw_elem.Id:
                log.info("Switching to Host Toposolid ID: {}".format(toposolid.Id))
        except: return
    except Exception as e:
        log.error("Invalid inputs for edging.", e)
        log.show(); return

    tg = TransactionGroup(doc, "Edging")
    tg.Start()
    
    try:
        z_s, z_e = calculate_and_adjust_stakes(state, log)
        if state.apply_offset:
            try:
                g_z_off = float(state.offset_val)
                z_s += g_z_off
                z_e += g_z_off
                log.info("Applied Z-Offset: {}".format(UnitHelper.to_formatted_string(g_z_off)))
            except: pass
            
        g_plan_off = 0.0
        g_plan_dir = "Both"
        if getattr(state, "apply_plan_offset", False):
            try:
                g_plan_off = float(getattr(state, "plan_offset_val", "0.0"))
                g_plan_dir = getattr(state, "plan_offset_dir", "Both")
                if g_plan_off != 0.0:
                    log.info("Applied Plan Offset: {} ({})".format(UnitHelper.to_formatted_string(g_plan_off), g_plan_dir))
            except: pass
        
        # Initialize Intersector for Raycasting Context
        intersector = None
        ray_start_z = get_toposolid_max_z(toposolid) + 100.0
        
        view3d = doc.ActiveView if doc.ActiveView.ViewType == ViewType.ThreeD else None
        if not view3d:
            cols = FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements()
            for v in cols:
                if not v.IsTemplate and v.Name == "{3D}":
                    view3d = v
                    break
            if not view3d and cols:
                for v in cols:
                    if not v.IsTemplate:
                        view3d = v
                        break
                        
        if view3d:
            col = FilteredElementCollector(doc).OfClass(Toposolid).WhereElementIsNotElementType()
            ids = List[ElementId]([e.Id for e in col])
            intersector = ReferenceIntersector(ids, FindReferenceTarget.Element, view3d)
        else:
            log.info("Warning: No 3D view found. Raycasting disabled. Using path elevation for new points.")
        
        edge_offset = w / 2.0
        edge_res = g * 0.5 
        curve = state.grading_line.GeometryCurve
        editor = toposolid.GetSlabShapeEditor()
        editor.Enable()
        
        to_move = [] 
        to_add = []
        
        all_verts = [v for v in editor.SlabShapeVertices]
        
        # 1. Snap existing nearby points to exact edge
        for v in all_verts: 
            p_3d, t_norm, d_xy = project_2d(curve, v.Position)
            vec_to_pt = flatten(v.Position) - flatten(p_3d)
            if vec_to_pt.IsAlmostEqualTo(XYZ.Zero): continue
            
            try:
                deriv = curve.ComputeDerivatives(t_norm, True)
                tangent = deriv.BasisX.Normalize()
                normal = XYZ.BasisZ.CrossProduct(tangent).Normalize() # Left normal
                side_sign = 1.0 if vec_to_pt.DotProduct(normal) >= 0 else -1.0
            except:
                side_sign = 1.0
                
            active_offset = edge_offset
            if g_plan_off != 0.0:
                if g_plan_dir == "Both": active_offset += g_plan_off
                elif g_plan_dir == "Center": active_offset += (g_plan_off / 2.0)
                elif g_plan_dir == "Left" and side_sign > 0: active_offset += g_plan_off
                elif g_plan_dir == "Right" and side_sign < 0: active_offset += g_plan_off

            if abs(d_xy - active_offset) < 1.0:
                vec = vec_to_pt.Normalize()
                exact_xy = flatten(p_3d) + (vec * active_offset)
                road_z = z_s + t_norm * (z_e - z_s)
                
                t_check = XYZ(exact_xy.X, exact_xy.Y, 0)
                rz, hit_id = get_surface_info(intersector, t_check, ray_start_z)
                off = 0.0
                if rz is not None: off = get_subdivision_offset(doc, hit_id)
                
                to_move.append((v, XYZ(exact_xy.X, exact_xy.Y, road_z - off)))
        
        # 2. Add new points along the exact edge
        step_t = edge_res / curve.Length
        step_t = 0.05 if step_t > 0.05 else step_t
        t_val = 0.0
        
        while t_val <= 1.001:
            eval_t = max(0.0, min(1.0, t_val))
            center_pt = curve.Evaluate(eval_t, True)
            try:
                tangent = curve.ComputeDerivatives(eval_t, True).BasisX.Normalize()
                normal = XYZ.BasisZ.CrossProduct(tangent).Normalize()
            except:
                normal = XYZ(0, 1, 0)
            
            road_z = z_s + eval_t * (z_e - z_s)
            
            for side in [1.0, -1.0]:
                active_offset = edge_offset
                if g_plan_off != 0.0:
                    if g_plan_dir == "Both": active_offset += g_plan_off
                    elif g_plan_dir == "Center": active_offset += (g_plan_off / 2.0)
                    elif g_plan_dir == "Left" and side > 0: active_offset += g_plan_off
                    elif g_plan_dir == "Right" and side < 0: active_offset += g_plan_off
                    
                offset_vec = normal * (side * active_offset)
                final_pt = center_pt + offset_vec
                
                # Raycast check for validity + offset
                rz, hit_id = get_surface_info(intersector, final_pt, ray_start_z)
                
                if rz is not None:
                    # It hit something valid (Toposolid or Subdiv)
                    off = get_subdivision_offset(doc, hit_id)
                    to_add.append(XYZ(final_pt.X, final_pt.Y, road_z - off))
            
            t_val += step_t
            
        t = Transaction(doc, "Apply Edging")
        t.Start()
        
        occupied_points = [v.Position for v in all_verts]
        
        # Apply moves
        for item in to_move:
            try: 
                editor.DeletePoint(item[0])
                editor.AddPoint(item[1])
                occupied_points.append(item[1])
            except: pass
            
        # Apply adds
        for pt in to_add:
            if not is_too_close(pt, occupied_points):
                try: 
                    editor.AddPoint(pt)
                    occupied_points.append(pt)
                except: pass 
                
        t.Commit()
        tg.Assimilate()
        log.info("Edging Complete.\nSnapped: {}\nAdded: {}".format(len(to_move), len(to_add)))
        
    except Exception as e:
        tg.RollBack()
        log.error("Edging failed.", e)
    finally:
        log.show()

def perform_smooth_region(state):
    log = BatchLogger()
    log.info("--- SMOOTH REGION STARTED ---")
    
    try:
        # 1. Pick Elements
        ref_topo = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid to Smooth")
        raw_topo = doc.GetElement(ref_topo)
        toposolid = resolve_toposolid_host(doc, raw_topo)
        if not isinstance(toposolid, Toposolid):
            log.error("First selection is not a Toposolid.")
            log.show(); return
        
        ref_reg = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Filled Region defining the boundary")
        region_elem = doc.GetElement(ref_reg)
        if not isinstance(region_elem, FilledRegion):
            log.error("Second selection is not a Filled Region.")
            log.show(); return
            
        # 2. Calculate Dynamic Resolution based on View Zoom
        active_view_id = doc.ActiveView.Id
        view_width = 50.0
        for uv in uidoc.GetOpenUIViews():
            if uv.ViewId == active_view_id:
                corners = uv.GetZoomCorners()
                view_width = abs(corners[1].X - corners[0].X)
                break
        
        # Divide screen width by 30 to get a nice smoothing resolution, clamped to a safe minimum of 0.25ft (3 inches)
        dynamic_res = max(0.25, view_width / 30.0) 
        log.info("Dynamic Grid Resolution: {} (Based on current zoom level)".format(UnitHelper.to_formatted_string(dynamic_res)))
        
        # 3. Extract Region Geometry
        all_curves = []
        for loop in region_elem.GetBoundaries():
            for c in loop: all_curves.append(c)
        
        min_x = min([c.Evaluate(i/10.0, True).X for c in all_curves for i in range(11)])
        max_x = max([c.Evaluate(i/10.0, True).X for c in all_curves for i in range(11)])
        min_y = min([c.Evaluate(i/10.0, True).Y for c in all_curves for i in range(11)])
        max_y = max([c.Evaluate(i/10.0, True).Y for c in all_curves for i in range(11)])
        
        def is_inside(pt):
            if pt.X < min_x or pt.X > max_x or pt.Y < min_y or pt.Y > max_y: return False
            ints = 0
            ray_y = pt.Y + 0.000137
            for c in all_curves:
                c_z = c.GetEndPoint(0).Z
                ray = Line.CreateBound(XYZ(pt.X, ray_y, c_z), XYZ(max_x + 100.0, ray_y, c_z))
                res_arr = clr.Reference[IntersectionResultArray]()
                res = c.Intersect(ray, res_arr)
                if res == SetComparisonResult.Overlap and res_arr.Value is not None:
                    ints += res_arr.Value.Size
            return ints % 2 != 0

        # Raycast Context
        intersector = None
        ray_start_z = get_toposolid_max_z(toposolid) + 100.0
        view3d = doc.ActiveView if doc.ActiveView.ViewType == ViewType.ThreeD else None
        if not view3d:
            for v in FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements():
                if not v.IsTemplate: view3d = v; break
        if view3d:
            intersector = ReferenceIntersector(toposolid.Id, FindReferenceTarget.Element, view3d)
        
        # 4. Laplacian Smoothing Process
        tg = TransactionGroup(doc, "Smooth Region")
        tg.Start()
        t = Transaction(doc, "Smooth Points")
        t.Start()
        
        editor = toposolid.GetSlabShapeEditor()
        editor.Enable()
        
        class Node:
            def __init__(self, x, y, z, v_ref=None):
                self.x = x; self.y = y; self.z = z
                self.v_ref = v_ref
                self.next_z = z
        
        nodes = []
        interior_nodes = []
        
        off = get_subdivision_offset(doc, raw_topo.Id) if toposolid.Id != raw_topo.Id else 0.0

        # A. Classify existing points
        for v in editor.SlabShapeVertices:
            n = Node(v.Position.X, v.Position.Y, v.Position.Z, v)
            if is_inside(v.Position): interior_nodes.append(n)
            nodes.append(n)
            
        # A.1 Outlier Detection
        try: outlier_tol_val = float(state.outlier_tol)
        except: outlier_tol_val = 1.0
        
        outlier_nodes = []
        outlier_radius = dynamic_res * 2.0
        for n in interior_nodes:
            neighbors = [nb for nb in nodes if nb != n and math.hypot(n.x - nb.x, n.y - nb.y) < outlier_radius]
            if len(neighbors) >= 3:
                mean_z = sum(nb.z for nb in neighbors) / len(neighbors)
                variance = sum((nb.z - mean_z) ** 2 for nb in neighbors) / len(neighbors)
                std_dev = math.sqrt(max(0.0, variance))
                
                # Flag as outlier if it deviates significantly (e.g. > 2 std devs AND > user threshold)
                if abs(n.z - mean_z) > max(2.0 * std_dev, outlier_tol_val):
                    outlier_nodes.append(n)
                    
        if outlier_nodes:
            log.info("Removing {} outlier points before smoothing.".format(len(outlier_nodes)))
            for out_n in outlier_nodes:
                if out_n in interior_nodes: interior_nodes.remove(out_n)
                if out_n in nodes: nodes.remove(out_n)
                if out_n.v_ref:
                    try: editor.DeletePoint(out_n.v_ref)
                    except: pass
        
        # B. Densify (Add missing grid points)
        x = math.floor(min_x / dynamic_res) * dynamic_res
        while x <= max_x:
            y = math.floor(min_y / dynamic_res) * dynamic_res
            while y <= max_y:
                pt = XYZ(x, y, 0)
                if is_inside(pt):
                    # Ensure not too close to existing vertices
                    if not any(math.hypot(n.x - x, n.y - y) < dynamic_res * 0.4 for n in interior_nodes):
                        rz = 0.0
                        if intersector:
                            z_hit, _ = get_surface_info(intersector, pt, ray_start_z)
                            if z_hit is not None: rz = z_hit - off
                        n = Node(x, y, rz)
                        interior_nodes.append(n)
                        nodes.append(n)
                y += dynamic_res
            x += dynamic_res
        
        # C. Identify boundaries and Smooth
        smoothing_radius = dynamic_res * 1.5
        boundary_nodes = [n for n in nodes if n not in interior_nodes and any(math.hypot(n.x - i.x, n.y - i.y) < smoothing_radius for i in interior_nodes)]
        active_nodes = interior_nodes + boundary_nodes
        
        for _ in range(5): # 5 Iterations of Laplacian averaging
            for n in interior_nodes:
                neighbors = [a for a in active_nodes if a != n and math.hypot(n.x - a.x, n.y - a.y) < smoothing_radius]
                if neighbors:
                    n.next_z = sum(nb.z for nb in neighbors) / len(neighbors)
            for n in interior_nodes:
                n.z = n.next_z
        
        # D. Apply to Toposolid
        add_count, mod_count = 0, 0
        for n in interior_nodes:
            final_pt = XYZ(n.x, n.y, n.z)
            if n.v_ref:
                if abs(n.v_ref.Position.Z - n.z) > 0.005:
                    try:
                        editor.DeletePoint(n.v_ref)
                        editor.AddPoint(final_pt)
                        mod_count += 1
                    except: pass
            else:
                try:
                    editor.AddPoint(final_pt)
                    add_count += 1
                except: pass
                
        t.Commit()
        tg.Assimilate()
        log.info("Smooth Region Complete.\nAdded {} new points, Smoothed {} existing points.".format(add_count, mod_count))
        
    except Exception as e:
        if 'tg' in locals() and tg.HasStarted(): tg.RollBack()
        log.error("Smooth Region Failed", "{}\n{}".format(e, traceback.format_exc()))
    finally:
        log.show()

def perform_add_points_along_line(state):
    log = BatchLogger()
    log.info("--- GRADE WITH ELEMENTS STARTED ---")
    
    g_int = float(state.grid)
    if g_int <= 0.1: g_int = 1.0 # Fallback safety
    
    g_plan_off = 0.0
    g_plan_dir = "Both"
    if getattr(state, "apply_plan_offset", False):
        try:
            g_plan_off = float(getattr(state, "plan_offset_val", "0.0"))
            g_plan_dir = getattr(state, "plan_offset_dir", "Both")
            if g_plan_off != 0.0:
                log.info("Applied Plan Offset: {} ({})".format(UnitHelper.to_formatted_string(g_plan_off), g_plan_dir))
        except: pass
    
    toposolid = None
    raw_elem = None
    target_curves = []
    target_internal_pts = []
    
    def get_curves_from_elem(elem, g_int):
        crvs = []
        internal_pts = []
        region_info = None
        if isinstance(elem, CurveElement):
            c = elem.GeometryCurve
            if isinstance(c, Line) and flatten(c.GetEndPoint(0)).DistanceTo(flatten(c.GetEndPoint(1))) < 0.01:
                log.error("Selected line is purely vertical.", "Toposolids cannot have vertically stacked points at the exact same XY coordinates. Please select a sloped or horizontal line.")
                return None, None, None
                
            crvs.append((c, None))
            
        elif isinstance(elem, FilledRegion):
            view = doc.GetElement(elem.OwnerViewId)
            z_val = 0.0
            if view and hasattr(view, "GenLevel") and view.GenLevel:
                z_val = view.GenLevel.Elevation
                
            all_curves = []
            for loop in elem.GetBoundaries():
                for c in loop:
                    crvs.append((c, z_val))
                    all_curves.append(c)
                    
            if all_curves:
                min_x, max_x = float('inf'), float('-inf')
                min_y, max_y = float('inf'), float('-inf')
                for c in all_curves:
                    for i in range(11):
                        pt = c.Evaluate(i / 10.0, True)
                        min_x = min(min_x, pt.X)
                        max_x = max(max_x, pt.X)
                        min_y = min(min_y, pt.Y)
                        max_y = max(max_y, pt.Y)
                        
                start_x = math.floor(min_x / g_int) * g_int
                end_x   = math.ceil(max_x / g_int) * g_int
                start_y = math.floor(min_y / g_int) * g_int
                end_y   = math.ceil(max_y / g_int) * g_int
                
                
                y = start_y
                while y <= end_y:
                    x = start_x
                    while x <= end_x:
                        intersections = 0
                        ray_y = y + 0.000137 # Infinitesimal offset avoids collinear/vertex anomalies
                        for c in all_curves:
                            c_z = c.GetEndPoint(0).Z
                            ray_line_z = Line.CreateBound(XYZ(x, ray_y, c_z), XYZ(max_x + 100.0, ray_y, c_z))
                            
                            res_array = clr.Reference[IntersectionResultArray]()
                            res = c.Intersect(ray_line_z, res_array)
                            if res == SetComparisonResult.Overlap and res_array.Value is not None:
                                intersections += res_array.Value.Size
                        
                        if intersections % 2 != 0:
                            internal_pts.append(XYZ(x, y, z_val))
                            
                        x += g_int
                    y += g_int
                
                region_info = (all_curves, z_val, max_x, min_x, min_y, max_y)
        return crvs, internal_pts, region_info

    try:
        ref_first = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid, Model Line, or Filled Region")
        elem_first = doc.GetElement(ref_first)
        
        # Visual Clue: Highlight first selection while waiting
        try: uidoc.Selection.SetElementIds(List[ElementId]([ref_first.ElementId]))
        except: pass
        
        crvs, int_pts, r_info = get_curves_from_elem(elem_first, g_int)
        target_region_info = None
        
        if crvs is not None and len(crvs) > 0:
            target_curves = crvs
            target_internal_pts = int_pts
            target_region_info = r_info
            ref_second = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Path selected. Now select Toposolid to modify")
            raw_elem = doc.GetElement(ref_second)
            toposolid = resolve_toposolid_host(doc, raw_elem)
            if not isinstance(toposolid, Toposolid):
                log.error("Second selection is not a Toposolid.")
                log.show(); return
                
        elif isinstance(elem_first, Toposolid):
            raw_elem = elem_first
            toposolid = resolve_toposolid_host(doc, raw_elem)
            
            ref_second = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Toposolid selected. Now select Line or Filled Region")
            elem_second = doc.GetElement(ref_second)
            crvs, int_pts, r_info = get_curves_from_elem(elem_second, g_int)
            if crvs is not None and len(crvs) > 0:
                target_curves = crvs
                target_internal_pts = int_pts
                target_region_info = r_info
            else:
                log.error("Second selection is not a valid line or filled region.")
                log.show(); return
        else:
            log.error("Selected element is neither a Toposolid, Model Line, nor Filled Region.")
            log.show(); return
            
    except: return # Cancelled
    
    g_int = float(state.grid)
    if g_int <= 0.1: g_int = 1.0 # Fallback safety
    
    g_z_off = 0.0
    if state.apply_offset:
        try:
            g_z_off = float(state.offset_val)
            log.info("Applied Z-Offset: {}".format(UnitHelper.to_formatted_string(g_z_off)))
        except: pass
    
    tg = TransactionGroup(doc, "Points Along Line")
    tg.Start()
    
    t = Transaction(doc, "Add Points")
    t.Start()
    
    try:
        editor = toposolid.GetSlabShapeEditor()
        editor.Enable()
        
        # --- PHASE 0: RESET POINTS ---
        if state.reset_mode:
            log.info("Reset Mode: Removing existing points in target area.")
            to_delete = []
            for v in editor.SlabShapeVertices:
                pos = v.Position
                should_delete = False
                
                if target_region_info:
                    all_curves, reg_z, max_x, min_x, min_y, max_y = target_region_info
                    if min_x <= pos.X <= max_x and min_y <= pos.Y <= max_y:
                        intersections = 0
                        ray_y = pos.Y + 0.000137
                        for c in all_curves:
                            c_z = c.GetEndPoint(0).Z
                            ray_line_z = Line.CreateBound(XYZ(pos.X, ray_y, c_z), XYZ(max_x + 100.0, ray_y, c_z))
                            res_array = clr.Reference[IntersectionResultArray]()
                            res = c.Intersect(ray_line_z, res_array)
                            if res == SetComparisonResult.Overlap and res_array.Value is not None:
                                intersections += res_array.Value.Size
                        if intersections % 2 != 0:
                            should_delete = True
                else:
                    for curve, _ in target_curves:
                        p_3d, t_norm, d_xy = project_2d(curve, pos)
                        if d_xy < (g_int * 0.75):
                            should_delete = True
                            break
                        if g_plan_off != 0.0 and g_plan_dir != "Center":
                            if abs(d_xy - g_plan_off) < (g_int * 0.75):
                                should_delete = True
                                break
                            
                if should_delete:
                    to_delete.append(v)
                    
            if to_delete:
                del_count = 0
                for v in to_delete:
                    try:
                        editor.DeletePoint(v)
                        del_count += 1
                    except: pass
                log.info("Removed {} existing points.".format(del_count))

        pts_to_add = []
        if g_z_off != 0.0:
            pts_to_add.extend([XYZ(p.X, p.Y, p.Z + g_z_off) for p in target_internal_pts])
        else:
            pts_to_add.extend(target_internal_pts)
        
        def add_curve_pts(curve, param, override_z):
            pt = curve.Evaluate(param, True)
            try:
                deriv = curve.ComputeDerivatives(param, True)
                tangent = deriv.BasisX.Normalize()
                if tangent.IsAlmostEqualTo(XYZ.Zero):
                    normal = XYZ(0, 1, 0)
                else:
                    normal = XYZ.BasisZ.CrossProduct(tangent).Normalize()
            except:
                normal = XYZ(0, 1, 0)
                
            base_z = override_z if override_z is not None else pt.Z
            base_pt = XYZ(pt.X, pt.Y, base_z + g_z_off)
            
            res = [base_pt]
            # Only apply planar offsets if g_plan_off is not zero
            if g_plan_off != 0.0:
                offset_vec = normal.Multiply(g_plan_off)
                if g_plan_dir == "Center":
                    # Offset by half the value on each side, centered on the line
                    offset_half_vec = normal.Multiply(g_plan_off / 2.0)
                    res.append(base_pt.Add(offset_half_vec))  # Left
                    res.append(base_pt.Subtract(offset_half_vec)) # Right
                elif g_plan_dir == "Left":
                    res.append(base_pt.Add(offset_vec))
                elif g_plan_dir == "Right":
                    res.append(base_pt.Subtract(offset_vec))
                elif g_plan_dir == "Both":
                    # Offset by the full value on each side
                    res.append(base_pt.Add(offset_vec))  # Left
                    res.append(base_pt.Subtract(offset_vec)) # Right
            return res
        for curve, override_z in target_curves:
            length = curve.Length
            step_len = g_int
            if step_len < 0.1: step_len = 1.0
            
            current_len = 0.0
            while current_len <= length + 0.001:
                param = current_len / length
                if param > 1.0: param = 1.0
                pts_to_add.extend(add_curve_pts(curve, param, override_z))
                current_len += step_len
                
            # Ensure the exact end point is included if it was missed
            if abs((current_len - step_len) - length) > 0.01:
                pts_to_add.extend(add_curve_pts(curve, 1.0, override_z))
                
        # Account for Subdivision thickness offset so points sit correctly
        off = 0.0
        if toposolid.Id != raw_elem.Id:
            off = get_subdivision_offset(doc, raw_elem.Id)
            
        # 1. Resolve self-intersections or vertical lines within the line's own points
        # We keep the latest evaluated elevation for any given XY coordinate
        unique_pts = []
        for pt in pts_to_add:
            conflict_idx = None
            cand_flat = flatten(pt)
            for i, u_pt in enumerate(unique_pts):
                if flatten(u_pt).DistanceTo(cand_flat) < MIN_DIST_TOLERANCE:
                    conflict_idx = i
                    break
            if conflict_idx is not None:
                unique_pts[conflict_idx] = pt # Override with latest height
            else:
                unique_pts.append(pt)
        
        added = 0
        modified = 0
        all_verts = [v for v in editor.SlabShapeVertices]
        
        # 2. Add or Update points on the Toposolid
        for pt in unique_pts:
            adjusted_pt = XYZ(pt.X, pt.Y, pt.Z - off)
            
            cand_flat = flatten(adjusted_pt)
            matched_vert = None
            for v in all_verts:
                if flatten(v.Position).DistanceTo(cand_flat) < MIN_DIST_TOLERANCE:
                    matched_vert = v
                    break
                    
            if matched_vert:
                if abs(matched_vert.Position.Z - adjusted_pt.Z) > 0.005:
                    try:
                        editor.DeletePoint(matched_vert)
                        editor.AddPoint(adjusted_pt)
                        modified += 1
                    except: pass
            else:
                try:
                    editor.AddPoint(adjusted_pt)
                    added += 1
                except: pass
                
        # 3. Enforce flattening for all existing points inside Filled Region
        if target_region_info and not state.reset_mode:
            all_curves, reg_z, max_x, min_x, min_y, max_y = target_region_info
            target_z = reg_z - off + g_z_off
            
            for v in editor.SlabShapeVertices:
                pos = v.Position
                # Fast bounding box check
                if pos.X < min_x or pos.X > max_x or pos.Y < min_y or pos.Y > max_y:
                    continue
                    
                if abs(pos.Z - target_z) > 0.005:
                    intersections = 0
                    ray_y = pos.Y + 0.000137
                    ray_end_x = max_x + 100.0
                    
                    for c in all_curves:
                        c_z = c.GetEndPoint(0).Z
                        ray_line_z = Line.CreateBound(XYZ(pos.X, ray_y, c_z), XYZ(ray_end_x, ray_y, c_z))
                        res_array = clr.Reference[IntersectionResultArray]()
                        res = c.Intersect(ray_line_z, res_array)
                        if res == SetComparisonResult.Overlap and res_array.Value is not None:
                            intersections += res_array.Value.Size
                            
                    if intersections % 2 != 0:
                        try:
                            editor.DeletePoint(v)
                            editor.AddPoint(XYZ(pos.X, pos.Y, target_z))
                            modified += 1
                        except: pass
                        
        t.Commit()
        tg.Assimilate()
        log.info("Successfully graded line: Added {} new points, Modified {} existing points.".format(added, modified))
    except Exception as e:
        t.RollBack()
        tg.RollBack()
        log.error("Failed to add points.", e)
    finally:
        log.show()

# NOTE: Future enhancement for triangulation control
# The Revit API allows modifying triangulation via SlabShapeEditor.DrawSplitLine(v1, v2).
# To improve grading quality, we could implement a pass that connects points
# perpendicular to the guide curve (Left Point <-> Right Point) using Split Lines.
# This forces the triangulation to align with the path flow, avoiding diagonal artifacts.
# WARNING: API-created Split Lines are "User Drawn" and appear as visible "Folding Lines".
# They cannot be individually hidden via API. Users must hide the "Folding Lines"
# subcategory in Visibility/Graphics to make them invisible.

def ensure_grade_stake_family():
    """Checks for ODI-GradeStake family and loads it if missing."""
    log = BatchLogger()
    fam_name = "ODI-GradeStake"
    
    # Check if already loaded
    collector = FilteredElementCollector(doc).OfClass(Family)
    if any(f.Name == fam_name for f in collector):
        return

    log.info("Family '{}' missing. Attempting to load...".format(fam_name))

    # Construct path: .../Site.stack/resources/ODI-GradeStake.rfa
    # __file__ is inside .../Site.stack/Grading.pushbutton/script.py
    # We use abspath to ensure we don't rely on the Current Working Directory (CWD)
    script_path = os.path.abspath(__file__)
    stack_dir = os.path.dirname(os.path.dirname(script_path))
    fam_path = os.path.join(stack_dir, "_Resources", "ODI-GradeStake.rfa")
    log.info("Target Family Path: {}".format(fam_path))

    if os.path.exists(fam_path):
        t = Transaction(doc, "Load Grading Stake Family")
        t.Start()
        try:
            if doc.LoadFamily(fam_path):
                t.Commit()
                log.info("Family loaded successfully.")
            else:
                t.RollBack()
                log.error("Revit LoadFamily returned False.")
        except Exception as e:
            t.RollBack()
            log.error("Error loading family.", e)
    else:
        log.error("Family file not found at path.")

    log.show()

# ==========================================
# 6. LOOP
# ==========================================
if __name__ == '__main__':
    ensure_grade_stake_family()
    state = GradingState()
    while True:
        win = GradingWindow(state)
        win.ShowDialog()
        
        # Capture Position
        try:
            state.win_top = win.Top
            state.win_left = win.Left
        except: pass
        
        action = win.next_action
        
        if not action:
            save_state_to_disk(state)
            break 
            
        elif action == "select_stakes":
            try:
                ref_start = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Start Stake")
                state.start_stake = doc.GetElement(ref_start)
                
                # Visual clue: Highlight the start stake while waiting
                try: uidoc.Selection.SetElementIds(List[ElementId]([ref_start.ElementId]))
                except: pass
                
                ref_end = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Start Stake selected. Now select End Stake")
                state.end_stake = doc.GetElement(ref_end)
                
                save_state_to_disk(state)
            except: pass # Cancelled selection
            
        elif action == "select_line":
            try: 
                state.grading_line = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Guide Line"))
                save_state_to_disk(state)
            except: pass # Cancelled selection
            
        elif action == "swap": 
            perform_swap(state)
            save_state_to_disk(state)
            
        elif action == "sculpt": 
            perform_sculpt(state)
            save_state_to_disk(state)
            
        elif action == "edge": 
            perform_edging(state)
            save_state_to_disk(state)
            
        elif action == "stitch": 
            perform_manual_stitch(state)
            
        elif action == "smooth":
            perform_smooth_region(state)
            
        elif action == "line_points":
            perform_add_points_along_line(state)
            save_state_to_disk(state)
            
        elif action == "load_recipe": 
            perform_load_recipe(state)
            save_state_to_disk(state)