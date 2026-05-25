# -*- coding: utf-8 -*- 
__title__ = "Grading Assistant"
__doc__ = "Advanced Toposolid grading tool with sculpting, edging, and auto-triangulation features."
__author__ = "Oxyzen Digital"
__context__ = "doc-project"

import os
import json
import math
import clr
import traceback
import threading

# --- ASSEMBLIES ---
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

# --- IMPORTS ---
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    XYZ, Transaction, TransactionGroup, ElementId, BuiltInParameter,
    ReferenceIntersector, FindReferenceTarget, Options, Solid, ViewType, View3D, Edge,
    ElementTransformUtils, CurveElement, UnitUtils, SpecTypeId, Line,
    FilteredElementCollector, Family, Toposolid, FilledRegion, IntersectionResultArray, SetComparisonResult,
    UnitFormatUtils, IFailuresPreprocessor, FailureProcessingResult,
    FailureSeverity, TransactionStatus
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB.ExtensibleStorage import SchemaBuilder, Schema, Entity, AccessLevel
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

class SilentErrorPreprocessor(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        failures = failuresAccessor.GetFailureMessages()
        if failures.Count == 0:
            return FailureProcessingResult.Continue
            
        has_error = False
        for failure in failures:
            severity = failure.GetSeverity()
            if severity == FailureSeverity.Error or severity == FailureSeverity.DocumentCorruption:
                has_error = True
            elif severity == FailureSeverity.Warning:
                try:
                    failuresAccessor.DeleteWarning(failure)
                except: pass
                
        if has_error:
            return FailureProcessingResult.ProceedWithRollBack
            
        return FailureProcessingResult.Continue

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
            if hasattr(element, "Name") and element.Name:
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
        "outlier_tol": "1.0", "point_dist_tol": "0.25",
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
        "point_dist_tol": getattr(state, "point_dist_tol", "0.25"),
        "mode": state.mode,
        "apply_offset": state.apply_offset,
        "offset_val": state.offset_val,
        "apply_plan_offset": getattr(state, 'apply_plan_offset', False),
        "plan_offset_val": getattr(state, 'plan_offset_val', "0.0"),
        "plan_offset_dir": getattr(state, 'plan_offset_dir', "Both"),
        "square_ends": state.square_ends,
        "sharp_smoothing": getattr(state, 'sharp_smoothing', False),
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
        self.point_dist_tol = sets.get("point_dist_tol", "0.25")
        self.slope_val = sets.get("slope", "2.0")
        self.mode = sets.get("mode", "stakes")
        self.square_ends = sets.get("square_ends", False)
        self.sharp_smoothing = sets.get("sharp_smoothing", False)
        self.apply_offset = sets.get("apply_offset", False)
        self.offset_val = sets.get("offset_val", "0.0")
        self.apply_plan_offset = sets.get("apply_plan_offset", False)
        self.plan_offset_val = sets.get("plan_offset_val", "0.0")
        self.plan_offset_dir = sets.get("plan_offset_dir", "Both")
        self.draw_split_lines = sets.get("draw_split_lines", False)
        self.reset_mode = False
        
        try:
            self.win_top = float(sets.get("win_top", "100"))
        except ValueError:
            self.win_top = 100.0
            
        try:
            self.win_left = float(sets.get("win_left", "100"))
        except ValueError:
            self.win_left = 100.0
        
        self.start_stake = None
        self.end_stake = None
        self.grading_lines = []
        
        self.next_action = None

    @property
    def ready(self):
        # Basic ready check: must have start stake and line.
        # End stake is needed if mode is NOT slope.
        has_start = self.start_stake is not None
        has_line = len(self.grading_lines) > 0
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
            if hasattr(self, "Tb_PointTol"):
                self.Tb_PointTol.Text = UnitHelper.to_formatted_string(getattr(self.state, "point_dist_tol", "0.25"))
        except:
            self.Tb_Width.Text = "6.0"
            self.Tb_Falloff.Text = "10.0"
            self.Tb_Grid.Text = "3.0"
            if hasattr(self, "Tb_OutlierTol"):
                self.Tb_OutlierTol.Text = "1.0"
            if hasattr(self, "Tb_PointTol"):
                self.Tb_PointTol.Text = "0.25"
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

        if self.state.grading_lines:
            self.Dot_Line.Fill          = dot_set
            self.Row_Line.Background    = bg_set
            if len(self.state.grading_lines) == 1:
                self.Lb_Line.Text       = get_element_label(self.state.grading_lines[0])
            else:
                self.Lb_Line.Text       = "{} curves selected".format(len(self.state.grading_lines))
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
                    
                    d_xy = dist_2d(p1, p2)
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
        if hasattr(self, "Cb_SharpSmoothing"):
            self.Cb_SharpSmoothing.IsChecked = getattr(self.state, "sharp_smoothing", False)
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
        for name in ("Lbl_Width_Unit", "Lbl_Falloff_Unit", "Lbl_Grid_Unit", "Lbl_Offset_Unit", "Lbl_PlanOffset_Unit", "Lbl_OutlierTol_Unit", "Lbl_PointTol_Unit"):
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
        try: self.Btn_SmoothPath.Click += self.a_smooth_path
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
        if hasattr(self, "Tb_PointTol"): tbs.append(self.Tb_PointTol)
        
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
            if hasattr(self, "Tb_PointTol"):
                _ = UnitHelper.to_internal(self.Tb_PointTol.Text)
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
    def h_line_on(self, s, a): self.set_selection(self.state.grading_lines)
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
        if hasattr(self, "Tb_PointTol"):
            try: self.state.point_dist_tol = str(UnitHelper.to_internal(self.Tb_PointTol.Text))
            except: pass
        self.state.slope_val = self.Tb_Slope.Text
        self.state.reset_mode = self.Cb_ResetPoints.IsChecked
        if hasattr(self, "Cb_SquareEnds"):
            self.state.square_ends = self.Cb_SquareEnds.IsChecked
        if hasattr(self, "Cb_SharpSmoothing"):
            self.state.sharp_smoothing = self.Cb_SharpSmoothing.IsChecked
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
    def a_smooth_path(self, s, a): self._raise("smooth_path")
    def a_line_points(self, s, a): self._raise("line_points")
    def a_load(self, s, a): self._raise("load_recipe")

# ==========================================
# 3. HELPERS
# ==========================================
class UniversalFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, r, p): return True

def dist_2d(p1, p2): return math.hypot(p1.X - p2.X, p1.Y - p2.Y)
def lerp(a, b, t): return a + t * (b - a)

def flatten(pt):
    """Flattens a 3D point to the XY plane."""
    return XYZ(pt.X, pt.Y, 0)

def calculate_miter_point(p_prev, p_curr, p_next, offset):
    v1 = p_curr - p_prev
    v2 = p_next - p_curr
    t1 = XYZ(v1.X, v1.Y, 0)
    t2 = XYZ(v2.X, v2.Y, 0)
    if t1.IsZeroLength() or t2.IsZeroLength(): return XYZ(p_curr.X, p_curr.Y, p_curr.Z)
    t1 = t1.Normalize()
    t2 = t2.Normalize()
    if t1.IsAlmostEqualTo(t2) or t1.IsAlmostEqualTo(t2.Negate()):
        n1 = XYZ(-t1.Y, t1.X, 0)
        return XYZ(p_curr.X, p_curr.Y, p_curr.Z) + n1 * offset
    t_m = (t1 + t2).Normalize()
    n_m = XYZ(-t_m.Y, t_m.X, 0)
    n1 = XYZ(-t1.Y, t1.X, 0)
    dot = n_m.DotProduct(n1)
    if abs(dot) < 0.15: return XYZ(p_curr.X, p_curr.Y, p_curr.Z) + n1 * offset 
    miter_len = offset / dot
    return XYZ(p_curr.X, p_curr.Y, p_curr.Z) + n_m * miter_len

def generate_offset_loop(pts, offset):
    if len(pts) < 3: return pts
    offset_pts = []
    is_closed = pts[0].DistanceTo(pts[-1]) < 0.01
    n = len(pts) - 1 if is_closed else len(pts)
    for i in range(n):
        offset_pts.append(calculate_miter_point(pts[i - 1], pts[i], pts[(i + 1) % n], offset))
    if is_closed: offset_pts.append(offset_pts[0])
    return offset_pts

def generate_offset_polyline(pts, offset):
    if len(pts) < 2: return pts
    if pts[0].DistanceTo(pts[-1]) < 0.01: return generate_offset_loop(pts, offset)
    offset_pts = []
    v_start = pts[1] - pts[0]
    t_start = XYZ(v_start.X, v_start.Y, 0)
    if not t_start.IsZeroLength(): t_start = t_start.Normalize()
    n_start = XYZ(-t_start.Y, t_start.X, 0)
    offset_pts.append(pts[0] + n_start * offset)
    for i in range(1, len(pts) - 1):
        offset_pts.append(calculate_miter_point(pts[i-1], pts[i], pts[i+1], offset))
    v_end = pts[-1] - pts[-2]
    t_end = XYZ(v_end.X, v_end.Y, 0)
    if not t_end.IsZeroLength(): t_end = t_end.Normalize()
    n_end = XYZ(-t_end.Y, t_end.X, 0)
    offset_pts.append(pts[-1] + n_end * offset)
    return offset_pts

class CurveChain(object):
    def __init__(self, curve_elements):
        self.curves = []
        self.Length = 0.0
        if not curve_elements: return
        raw_curves = []
        for e in curve_elements:
            if isinstance(e, CurveElement): raw_curves.append(e.GeometryCurve)
            elif hasattr(e, "GetEndPoint"): raw_curves.append(e)
        if not raw_curves: return
        connected = [False] * len(raw_curves)
        chain = []
        def is_shared(pt, curves, skip_idx):
            for i, c in enumerate(curves):
                if i == skip_idx: continue
                if c.GetEndPoint(0).DistanceTo(pt) < 0.01 or c.GetEndPoint(1).DistanceTo(pt) < 0.01: return True
            return False
        start_idx = 0
        reverse_first = False
        for i, c in enumerate(raw_curves):
            if not is_shared(c.GetEndPoint(0), raw_curves, i):
                start_idx = i; reverse_first = False; break
            if not is_shared(c.GetEndPoint(1), raw_curves, i):
                start_idx = i; reverse_first = True; break
        curr_curve = raw_curves[start_idx]
        if reverse_first: curr_curve = curr_curve.CreateReversed()
        chain.append(curr_curve)
        connected[start_idx] = True
        curr_end = curr_curve.GetEndPoint(1)
        while len(chain) < len(raw_curves):
            found = False
            for i, c in enumerate(raw_curves):
                if connected[i]: continue
                if c.GetEndPoint(0).DistanceTo(curr_end) < 0.01:
                    chain.append(c); curr_end = c.GetEndPoint(1); connected[i] = True; found = True; break
                elif c.GetEndPoint(1).DistanceTo(curr_end) < 0.01:
                    rev_c = c.CreateReversed(); chain.append(rev_c); curr_end = rev_c.GetEndPoint(1)
                    connected[i] = True; found = True; break
            if not found: break
        self.curves = chain
        self.Length = sum(c.Length for c in self.curves)
        
    def get_tessellated_points(self):
        pts = []
        for c in self.curves:
            tess = c.Tessellate()
            if not pts: pts.extend(tess)
            else:
                for pt in tess[1:]:
                    if pts[-1].DistanceTo(pt) > 0.001: pts.append(pt)
        return pts
        
    def Evaluate(self, param, normalized=True):
        target_len = param * self.Length if normalized else param
        if target_len <= 0: return self.curves[0].GetEndPoint(0)
        if target_len >= self.Length: return self.curves[-1].GetEndPoint(1)
        curr = 0.0
        for c in self.curves:
            if curr + c.Length >= target_len - 0.0001:
                return c.Evaluate((target_len - curr) / c.Length, True)
            curr += c.Length
        return self.curves[-1].GetEndPoint(1)
        
    def ComputeDerivatives(self, param, normalized=True):
        target_len = param * self.Length if normalized else param
        if target_len <= 0: return self.curves[0].ComputeDerivatives(0.0, True)
        if target_len >= self.Length: return self.curves[-1].ComputeDerivatives(1.0, True)
        curr = 0.0
        for c in self.curves:
            if curr + c.Length >= target_len - 0.0001:
                return c.ComputeDerivatives((target_len - curr) / c.Length, True)
            curr += c.Length
        return self.curves[-1].ComputeDerivatives(1.0, True)

def setup_toposolid_intersector(doc, toposolid, log):
    """Sets up a 3D raycasting context to accurately find Toposolid elevations."""
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
                if not v.IsTemplate: view3d = v; break
    
    intersector = None
    if view3d:
        intersector = ReferenceIntersector(toposolid.Id, FindReferenceTarget.Element, view3d)
    else:
        log.info("Warning: No 3D view found. Raycasting disabled.")
    return intersector, ray_start_z

# Cache for curve discretizations to avoid IronPython/C# interop overhead
_curve_eval_cache = {}

def project_2d(curve, pt):
    """
    Projects a 3D point onto a curve in the XY plane.
    Returns: (closest_3d_pt, normalized_param, distance_xy)
    """
    pt_x, pt_y = pt.X, pt.Y

    if isinstance(curve, Line):
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        v_x = p1.X - p0.X
        v_y = p1.Y - p0.Y
        len_sq = v_x**2 + v_y**2
        
        if len_sq < 0.0001:
            d = math.hypot(pt_x - p0.X, pt_y - p0.Y)
            return p0, 0.0, d
            
        t = ((pt_x - p0.X) * v_x + (pt_y - p0.Y) * v_y) / len_sq
        t_bounded = max(0.0, min(1.0, t))
        
        pt_3d = p0 + (p1 - p0) * t_bounded
        d = math.hypot(pt_x - pt_3d.X, pt_y - pt_3d.Y)
        return pt_3d, t_bounded, d
    else:
        p_min, p_max = curve.GetEndParameter(0), curve.GetEndParameter(1)
        
        # Create a stable geometric signature to avoid wrapper hash volatility
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        cache_key = (
            round(p0.X, 4), round(p0.Y, 4), round(p0.Z, 4),
            round(p1.X, 4), round(p1.Y, 4), round(p1.Z, 4),
            round(curve.Length, 4)
        )
        if cache_key not in _curve_eval_cache:
            # High resolution steps for linear interpolation
            steps = int(max(100, curve.Length * 10.0))
            eval_pts = []
            for i in range(steps + 1):
                p = curve.Evaluate(p_min + (i / float(steps)) * (p_max - p_min), False)
                eval_pts.append((p.X, p.Y, p.Z))
            _curve_eval_cache[cache_key] = (steps, eval_pts)
            
        steps, eval_pts = _curve_eval_cache[cache_key]
        
        best_dist_sq = float('inf')
        best_px, best_py, best_pz = 0.0, 0.0, 0.0
        best_t_norm = 0.0
        
        # Pure Python segment projection to eliminate curve.Evaluate in the loop
        for i in range(steps):
            ax, ay, az = eval_pts[i]
            bx, by, bz = eval_pts[i+1]
            
            vx, vy = bx - ax, by - ay
            len_sq = vx**2 + vy**2
            
            if len_sq < 0.000001:
                t_seg = 0.0
                px, py, pz = ax, ay, az
            else:
                t_seg = max(0.0, min(1.0, ((pt_x - ax) * vx + (pt_y - ay) * vy) / len_sq))
                px = ax + vx * t_seg
                py = ay + vy * t_seg
                pz = az + (bz - az) * t_seg
                
            d_sq = (pt_x - px)**2 + (pt_y - py)**2
            if d_sq < best_dist_sq:
                best_dist_sq = d_sq
                best_px, best_py, best_pz = px, py, pz
                best_t_norm = (i + t_seg) / float(steps)
                
        return XYZ(best_px, best_py, best_pz), best_t_norm, math.sqrt(best_dist_sq)

def project_2d_chain(chain, pt):
    best_pt, best_d, best_t_global = None, float('inf'), 0.0
    if chain.Length == 0: return pt, 0.0, 0.0
    curr_len = 0.0
    for c in chain.curves:
        p_3d, t_local, d_xy = project_2d(c, pt)
        if d_xy < best_d:
            best_d = d_xy
            best_pt = p_3d
            best_t_global = (curr_len + t_local * c.Length) / chain.Length
        curr_len += c.Length
    return best_pt, best_t_global, best_d

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
    cx, cy = candidate_pt.X, candidate_pt.Y
    tol_sq = tolerance * tolerance
    for existing in occupied_points:
        if (cx - existing.X)**2 + (cy - existing.Y)**2 < tol_sq: return True
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

_subdivision_offset_cache = {}

def get_subdivision_offset(doc, elem_id):
    """Checks if the element is a subdivision and returns its height parameter."""
    if not elem_id or elem_id == ElementId.InvalidElementId: return 0.0
    
    try:
        key = elem_id.Value
    except AttributeError:
        key = elem_id.IntegerValue
        
    if key in _subdivision_offset_cache:
        return _subdivision_offset_cache[key]
        
    try:
        el = doc.GetElement(elem_id)
        # Check if it has HostTopoId (Revit 2024+)
        if hasattr(el, "HostTopoId") and el.HostTopoId != ElementId.InvalidElementId:
            p = el.get_Parameter(BuiltInParameter.TOPOSOLID_SUBDIV_HEIGHT)
            if p: 
                val = p.AsDouble()
                _subdivision_offset_cache[key] = val
                return val
    except: pass
    
    _subdivision_offset_cache[key] = 0.0
    return 0.0

class VirtualVertexTracker(object):
    """
    Maintains a purely Python-side virtual dictionary of all active Toposolid vertices 
    to eliminate slow C# interop API calls during dense loops.
    """
    def __init__(self, doc, toposolid):
        self.editor = toposolid.GetSlabShapeEditor()
        
        # Force editor enable within a tiny transaction to prevent Revit modification errors
        # If a transaction is already open, just enable it directly to avoid crashing.
        if not doc.IsModifiable:
            t_init = Transaction(doc, "Init Editor")
            t_init.Start()
            self.editor.Enable()
            t_init.Commit()
        else:
            self.editor.Enable()
        
        self.active_verts = {}
        # Robust iteration to gracefully skip corrupted enumerators or vertices
        try:
            for v in self.editor.SlabShapeVertices:
                try:
                    pos = v.Position
                    self.active_verts[(round(pos.X, 3), round(pos.Y, 3))] = (v, pos)
                except Exception:
                    pass # Gracefully skip corrupted vertex
        except Exception:
            pass # Gracefully handle full enumerator failure
            
    def delete(self, v_obj, pos_obj):
        try:
            self.editor.DeletePoint(v_obj)
            self.active_verts.pop((round(pos_obj.X, 3), round(pos_obj.Y, 3)), None)
            return True
        except Exception: return False
            
    def add(self, pos_obj):
        try:
            new_v = self.editor.AddPoint(pos_obj)
            if new_v: self.active_verts[(round(pos_obj.X, 3), round(pos_obj.Y, 3))] = (new_v, pos_obj)
            return True
        except Exception: return False
            
    def get_all(self):
        """Returns a snapshot list of (vertex, position) tuples."""
        return list(self.active_verts.values())

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
    if not state.grading_lines:
        log.error("Guide Line is missing.")
        return False
    
    try:
        if not state.start_stake.IsValidObject:
            log.error("Start Stake element is no longer valid.")
            return False
        for gl in state.grading_lines:
            if not gl.IsValidObject:
                log.error("Guide Line element is no longer valid.")
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
    if not state.grading_lines:
        log.error("Invalid Guide Lines selected.")
        raise Exception("Invalid Chain")

    chain = CurveChain(state.grading_lines)
    if chain.Length == 0:
        log.error("Guide lines could not be formed into a continuous chain.")
        raise Exception("Invalid Chain")
    l_start = chain.curves[0].GetEndPoint(0)
    l_end = chain.curves[-1].GetEndPoint(1)
    
    # Failsafe: Stake might not have a LocationPoint if user picked wrong family
    if not hasattr(state.start_stake, "Location") or not hasattr(state.start_stake.Location, "Point"):
        log.error("Start Stake does not have a valid location point.")
        raise Exception("Invalid Stake")

    u_start_pt = state.start_stake.Location.Point
    z_start = u_start_pt.Z
    dist_start = u_start_pt.DistanceTo(l_start)
    dist_end = u_start_pt.DistanceTo(l_end)
    is_flipped = dist_end < dist_start 
    length = chain.Length
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
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid")
        if not ref: return
    except:
        return # Cancelled
        
    log = BatchLogger()
    try:
        data = GradingRecipe.read_recipe(doc.GetElement(ref))
        if data:
            state.width = str(data.get("width", "6.0"))
            state.falloff = str(data.get("falloff", "10.0"))
            state.grid = str(data.get("grid", "3.0"))
            state.outlier_tol = str(data.get("outlier_tol", "1.0"))
            state.point_dist_tol = str(data.get("point_dist_tol", "0.25"))
            state.slope_val = str(data.get("slope", "2.0"))
            state.mode = str(data.get("mode", "stakes"))
            state.apply_offset = data.get("apply_offset", False)
            state.offset_val = str(data.get("offset_val", "0.0"))
            state.apply_plan_offset = data.get("apply_plan_offset", False)
            state.plan_offset_val = str(data.get("plan_offset_val", "0.0"))
            state.plan_offset_dir = data.get("plan_offset_dir", "Both")
            state.square_ends = data.get("square_ends", False)
            state.sharp_smoothing = data.get("sharp_smoothing", False)
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
            
            best_loop = None
            min_dist = 1.0 # Tolerance ft
            
            # Iterate CurveArrArray (Profile is list of loops)
            for curve_array in sketch.Profile:
                # Check if this loop is "close" to our selection
                is_match = False
                for i in range(curve_array.Size):
                    sc = curve_array.get_Item(i)
                    
                    # Project selected midpoint to sketch curve (2D check)
                    sp0 = sc.GetEndPoint(0)
                    sp1 = sc.GetEndPoint(1)
                    
                    # Simple distance to segment check
                    # Or assume sketch curve is planar Z-flat, just ignore Z
                    # We can use our 'flatten' to create a new Line for check? 
                    # No, generic curve might be arc.
                    # Let's project 'mid_flat' onto 'sc' ignoring Z?
                    # Hard to do generically without creating geometry.
                    
                    # Quick check: Is mid_flat close to endpoints?
                    if dist_2d(mid_pt, sp0) < 5.0 or dist_2d(mid_pt, sp1) < 5.0:
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
                            d_xy = dist_2d(p_res, mid_pt)
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
        fail_opts = t.GetFailureHandlingOptions()
        fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
        t.SetFailureHandlingOptions(fail_opts)
        
        try:
            tracker = VirtualVertexTracker(doc, toposolid)
            
            existing_coords = set()
            candidates = []
            
            # Cache existing vertices
            for v, pos in tracker.get_all():
                existing_coords.add((round(pos.X, 4), round(pos.Y, 4)))
                candidates.append((v, pos))
            
            log.info("Internal Vertices to Check: {}".format(len(candidates)))
            
            points_to_add = []
            
            # 2. Find Candidates: Internal Points -> Project -> Closest Edge in Chain
            check_count = 0
            match_count = 0
            
            for v, pos in candidates:
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
                        dist = dist_2d(v.Position, p_target)
                        
                        if dist < best_dist:
                            best_dist = dist
                            best_proj = p_target
                    except: pass
                
                # Logic: If close enough (but not ON the edge), propose new point
                if best_proj and (0.005 < best_dist < snap_dist):
                    
                    offset_pt = best_proj
                    if g_plan_off != 0.0:
                        vec_to_v = XYZ(v.Position.X - best_proj.X, v.Position.Y - best_proj.Y, 0)
                        if not vec_to_v.IsAlmostEqualTo(XYZ.Zero):
                            dir_outward = vec_to_v.Normalize().Negate()
                            offset_pt = best_proj + (dir_outward * g_plan_off)
                            
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
                if tracker.add(p): added_count += 1
            
            log.info("Successfully Added Points: {}".format(added_count))

            status = t.Commit()
            if status == TransactionStatus.Committed:
                tg.Assimilate()
            else:
                if tg.HasStarted(): tg.RollBack()
                log.error("Revit rejected the changes.", "The Toposolid geometry may have become invalid.")
            
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
            "point_dist_tol": getattr(state, "point_dist_tol", "0.25"),
            "slope": state.slope_val,
            "apply_offset": state.apply_offset,
            "offset_val": state.offset_val,
            "apply_plan_offset": getattr(state, 'apply_plan_offset', False),
            "plan_offset_val": getattr(state, 'plan_offset_val', "0.0"),
            "plan_offset_dir": getattr(state, 'plan_offset_dir', "Both"),
            "mode": state.mode,
            "sharp_smoothing": getattr(state, "sharp_smoothing", False),
            "draw_split_lines": getattr(state, "draw_split_lines", False)
        }
        
        t_rec = Transaction(doc, "Save Recipe")
        t_rec.Start()
        try:
            GradingRecipe.save_recipe(toposolid, rec)
            t_rec.Commit()
        except: t_rec.RollBack()
        
        tracker = VirtualVertexTracker(doc, toposolid)

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

        chain = CurveChain(state.grading_lines)
        log.info("Guide Chain Length: {:.2f}".format(chain.Length))
        pts = chain.get_tessellated_points()

        core_rad = w_int / 2.0
        total_rad = core_rad + f_int
        
        # Pre-calc for Square Ends
        check_square_ends = False
        sq_ends_info = []
        if state.square_ends:
            try:
                sq_ends_info.append({'pt': pts[0], 'tan': XYZ(pts[0].X - pts[1].X, pts[0].Y - pts[1].Y, 0).Normalize()})
                sq_ends_info.append({'pt': pts[-1], 'tan': XYZ(pts[-1].X - pts[-2].X, pts[-1].Y - pts[-2].Y, 0).Normalize()})
                check_square_ends = True

            except Exception as e:
                log.error("Square Ends Calc Failed", e)

        def is_outside_bounds(pt):
            if not check_square_ends: return False
            for ep in sq_ends_info:
                if XYZ(pt.X - ep['pt'].X, pt.Y - ep['pt'].Y, 0).DotProduct(ep['tan']) > 0.001: return True
            return False

        # Phase 0: Reset (Optional)
        if state.reset_mode:
            log.info("--- PHASE 0: RESET POINTS ---")
            t0 = Transaction(doc, "Reset Points")
            t0.Start()
            fail_opts0 = t0.GetFailureHandlingOptions()
            fail_opts0.SetFailuresPreprocessor(SilentErrorPreprocessor())
            t0.SetFailureHandlingOptions(fail_opts0)
            try:
                to_delete = []
                for v, pos in tracker.get_all():
                    p_3d, t_norm, d_xy = project_2d_chain(chain, pos)
                    if d_xy < (total_rad + g_int):
                        to_delete.append((v, pos))
                
                if to_delete:
                    count_del = 0
                    for v_del, pos_del in to_delete:
                        if tracker.delete(v_del, pos_del):
                            count_del += 1
                    log.info("Removed {} old points to clear resolution.".format(count_del))
                else:
                    log.info("No points found within grading zone to remove.")
                status0 = t0.Commit()
                if status0 != TransactionStatus.Committed:
                    log.error("Reset Points transaction was rejected by Revit.")
            except Exception as e:
                t0.RollBack()
                log.error("Failed to reset points.", e)
        # Generate Station & Offset Ribbon Grid
        offsets_set = set([0.0, core_rad, -core_rad, total_rad, -total_rad])
        off = g_int
        while off < core_rad: offsets_set.add(off); offsets_set.add(-off); off += g_int
        if f_int > 0:
            off = core_rad + g_int
            while off < total_rad: offsets_set.add(off); offsets_set.add(-off); off += g_int
        offsets = sorted(list(offsets_set))
        
        ribbon_pts = []
        for off_val in offsets:
            off_line = generate_offset_polyline(pts, off_val)
            for i in range(len(off_line) - 1):
                p_start, p_end = off_line[i], off_line[i+1]
                vec = p_end - p_start
                seg_len = XYZ(vec.X, vec.Y, 0).GetLength()
                if seg_len < 0.001: continue
                steps = int(math.ceil(seg_len / g_int))
                for s in range(steps + 1):
                    t = (s * g_int) / seg_len if (s * g_int) < seg_len else 1.0
                    ribbon_pts.append(p_start + vec * t)
                    
        if not state.square_ends:
            for ep in [pts[0], pts[-1]]:
                rx_s, rx_e = math.floor((ep.X - total_rad)/g_int)*g_int, math.ceil((ep.X + total_rad)/g_int)*g_int
                ry_s, ry_e = math.floor((ep.Y - total_rad)/g_int)*g_int, math.ceil((ep.Y + total_rad)/g_int)*g_int
                x = rx_s
                while x <= rx_e:
                    y = ry_s
                    while y <= ry_e:
                        pt = XYZ(x, y, 0)
                        if dist_2d(pt, ep) <= total_rad: ribbon_pts.append(pt)
                        y += g_int
                    x += g_int

        # Phase 1: Densify
        log.info("--- PHASE 1: DENSIFY ---")
        t1 = Transaction(doc, "Densify")
        t1.Start()
        fail_opts1 = t1.GetFailureHandlingOptions()
        fail_opts1.SetFailuresPreprocessor(SilentErrorPreprocessor())
        t1.SetFailureHandlingOptions(fail_opts1)
        try:
            log.info("Initial Vertex Count: {}".format(len(tracker.get_all())))
            occupied_points = [pos for v, pos in tracker.get_all()]
            
            grid_pts = []
            for t_pt in ribbon_pts:
                p_3d, t_norm, d_xy = project_2d_chain(chain, t_pt)
                if d_xy < (total_rad + 0.01):
                    if is_point_on_solid(intersector, t_pt, ray_start_z):
                        rz, hit_id = get_surface_info(intersector, t_pt, ray_start_z)
                        off_val = get_subdivision_offset(doc, hit_id) if rz is not None else 0.0
                        if rz is None: rz = z_s + t_norm * (z_e - z_s)
                        grid_pts.append(XYZ(t_pt.X, t_pt.Y, rz - off_val))
                
            # Deduplicate Ribbon Points
            cell_size = g_int * 0.5
            pts_grid = {}
            unique_ribbon = []
            for pt in grid_pts:
                cx, cy = int(math.floor(pt.X / cell_size)), int(math.floor(pt.Y / cell_size))
                conflict = False
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        for idx in pts_grid.get((cx + dx, cy + dy), []):
                            if dist_2d(unique_ribbon[idx], pt) < MIN_DIST_TOLERANCE:
                                conflict = True; break
                        if conflict: break
                    if conflict: break
                if not conflict:
                    unique_ribbon.append(pt)
                    pts_grid.setdefault((cx, cy), []).append(len(unique_ribbon)-1)
                    
            log.info("Unique Ribbon Grid Points: {}".format(len(unique_ribbon)))
            added_count = 0
            for p in unique_ribbon:
                if not is_too_close(p, occupied_points, MIN_DIST_TOLERANCE):
                    if tracker.add(p):
                        occupied_points.append(p)
                        added_count += 1
            status1 = t1.Commit()
            if status1 != TransactionStatus.Committed:
                raise Exception("Densify transaction was rejected by Revit.")
            log.info("Points Added: {}".format(added_count))
        except:
            t1.RollBack()
            raise

        # Phase 2: Sculpt
        log.info("--- PHASE 2: SCULPT ---")
        t2 = Transaction(doc, "Sculpt")
        t2.Start()
        fail_opts2 = t2.GetFailureHandlingOptions()
        fail_opts2.SetFailuresPreprocessor(SilentErrorPreprocessor())
        t2.SetFailureHandlingOptions(fail_opts2)
        modified_count = 0
        try:
            updates = []
            current_verts = tracker.get_all()
            log.info("Total Vertices to Process: {}".format(len(current_verts)))
            sample_log = []

            for v, pos in current_verts:
                if is_outside_bounds(pos): continue
                
                p_3d, t_norm, d_xy = project_2d_chain(chain, pos)
                if d_xy > total_rad: continue
                
                rz, hit_id = get_surface_info(intersector, pos, ray_start_z)
                off = 0.0
                if rz is not None:
                     off = get_subdivision_offset(doc, hit_id)
                
                # Current State
                current_base_z = pos.Z
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
                
                if abs(new_base_z - pos.Z) > 0.005:
                    updates.append(XYZ(pos.X, pos.Y, new_base_z))
                    if len(sample_log) < 3:
                        sample_log.append("Pt ({:.1f}, {:.1f}): Z {:.2f} -> {:.2f} (Off={:.2f})".format(
                            pos.X, pos.Y, pos.Z, new_base_z, off
                        ))
            
            modified_count = len(updates)
            log.info("Points Identified for Modification: {}".format(modified_count))
            if sample_log:
                log.info("Sample Changes:\n" + "\n".join(sample_log))
            
            for p in updates:
                tracker.add(p)
            
            status2 = t2.Commit()
            if status2 != TransactionStatus.Committed:
                raise Exception("Sculpt transaction was rejected by Revit (Toposolid may be too thin).")
        except:
            t2.RollBack()
            raise

        # Phase 3: Triangulate (Split Lines)
        if getattr(state, 'draw_split_lines', False):
            log.info("--- PHASE 3: TRIANGULATE ---")
            t3 = Transaction(doc, "Triangulate Path")
            t3.Start()
            fail_opts3 = t3.GetFailureHandlingOptions()
            fail_opts3.SetFailuresPreprocessor(SilentErrorPreprocessor())
            t3.SetFailureHandlingOptions(fail_opts3)
            split_lines_count = 0
            try:
                all_verts = tracker.get_all()
                search_tol = g_int * 0.25
                step_len = g_int
                if step_len < 0.1: step_len = 1.0
                
                l_len = chain.Length
                current_len = 0.0
                
                while current_len <= l_len:
                    norm_param = current_len / l_len
                    if norm_param > 1.0: norm_param = 1.0
                    
                    center_pt = chain.Evaluate(norm_param, True)
                    deriv = chain.ComputeDerivatives(norm_param, True)
                    tangent = deriv.BasisX.Normalize()
                    normal = tangent.CrossProduct(XYZ.BasisZ) 
                    
                    p_left = center_pt + (normal * core_rad)
                    p_right = center_pt - (normal * core_rad) 
                    
                    v_left = None; pos_left = None
                    v_right = None; pos_right = None
                    
                    min_d_l = search_tol
                    min_d_r = search_tol
                    
                    for v, pos in all_verts:
                        d_l = dist_2d(pos, p_left)
                        if d_l < min_d_l:
                            min_d_l = d_l
                            v_left = v
                            pos_left = pos
                        d_r = dist_2d(pos, p_right)
                        if d_r < min_d_r:
                            min_d_r = d_r
                            v_right = v
                            pos_right = pos
                    
                    if v_left and v_right:
                        try:
                            if not pos_left.IsAlmostEqualTo(pos_right):
                                dist_real = pos_left.DistanceTo(pos_right)
                                if abs(dist_real - w_int) < (w_int * 0.5):
                                    tracker.editor.DrawSplitLine(v_left, v_right)
                                    split_lines_count += 1
                        except: pass
                    
                    current_len += step_len

                status3 = t3.Commit()
                if status3 == TransactionStatus.Committed:
                    log.info("Triangulation Complete. Split Lines Added: {}".format(split_lines_count))
                else:
                    log.error("Triangulation transaction was rejected by Revit.")
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
        chain = CurveChain(state.grading_lines)
        pts = chain.get_tessellated_points()
        
        tracker = VirtualVertexTracker(doc, toposolid)
        
        to_move = [] 
        to_add = []
        
        # 1. Snap existing nearby points to exact edge
        for v, pos in tracker.get_all(): 
            p_3d, t_norm, d_xy = project_2d_chain(chain, pos)
            vec_to_pt = XYZ(pos.X - p_3d.X, pos.Y - p_3d.Y, 0)
            if vec_to_pt.IsAlmostEqualTo(XYZ.Zero): continue
            
            try:
                deriv = chain.ComputeDerivatives(t_norm, True)
                tangent = deriv.BasisX.Normalize()
                normal = XYZ.BasisZ.CrossProduct(tangent).Normalize() # Left normal
                side_sign = 1.0 if vec_to_pt.DotProduct(normal) >= 0 else -1.0
            except:
                side_sign = 1.0
                
            active_offset = edge_offset
            if g_plan_off != 0.0:
                if g_plan_dir == "Both": active_offset += g_plan_off
                elif g_plan_dir == "Center": active_offset += (g_plan_off / 2.0)
                elif g_plan_dir == "Inside" and side_sign > 0: active_offset += g_plan_off
                elif g_plan_dir == "Outside" and side_sign < 0: active_offset += g_plan_off

            if abs(d_xy - active_offset) < 1.0:
                vec = vec_to_pt.Normalize()
                exact_x = p_3d.X + vec.X * active_offset
                exact_y = p_3d.Y + vec.Y * active_offset
                road_z = z_s + t_norm * (z_e - z_s)
                
                t_check = XYZ(exact_x, exact_y, 0)
                rz, hit_id = get_surface_info(intersector, t_check, ray_start_z)
                off = 0.0
                if rz is not None: off = get_subdivision_offset(doc, hit_id)
                
                to_move.append((v, pos, XYZ(exact_x, exact_y, road_z - off)))
        
        # 2. Add new points along the exact edge
        offset_lines = []
        for side in [1.0, -1.0]:
            active_offset = edge_offset
            if g_plan_off != 0.0:
                if g_plan_dir == "Both": active_offset += g_plan_off
                elif g_plan_dir == "Center": active_offset += (g_plan_off / 2.0)
                elif g_plan_dir == "Inside" and side > 0: active_offset += g_plan_off
                elif g_plan_dir == "Outside" and side < 0: active_offset += g_plan_off
            offset_lines.append(generate_offset_polyline(pts, side * active_offset))
            
        for off_line in offset_lines:
            for i in range(len(off_line) - 1):
                p_start, p_end = off_line[i], off_line[i+1]
                vec = p_end - p_start
                seg_len = XYZ(vec.X, vec.Y, 0).GetLength()
                if seg_len < 0.001: continue
                steps = int(math.ceil(seg_len / edge_res))
                for s in range(steps + 1):
                    t_param = (s * edge_res) / seg_len if (s * edge_res) < seg_len else 1.0
                    hp = p_start + vec * t_param
                    
                    _, t_norm, _ = project_2d_chain(chain, hp)
                    road_z = z_s + t_norm * (z_e - z_s)
                    
                    rz, hit_id = get_surface_info(intersector, hp, ray_start_z)
                    if rz is not None:
                        off_val = get_subdivision_offset(doc, hit_id)
                        to_add.append(XYZ(hp.X, hp.Y, road_z - off_val))
            
        t = Transaction(doc, "Apply Edging")
        t.Start()
        fail_opts = t.GetFailureHandlingOptions()
        fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
        t.SetFailureHandlingOptions(fail_opts)
        
        occupied_points = [pos for v, pos in tracker.get_all()]
        
        # Apply moves
        for v_obj, old_pos, new_pt in to_move:
            if tracker.delete(v_obj, old_pos):
                if tracker.add(new_pt):
                    occupied_points.append(new_pt)
            
        # Apply adds
        for pt in to_add:
            if not is_too_close(pt, occupied_points, MIN_DIST_TOLERANCE):
                if tracker.add(pt):
                    occupied_points.append(pt)
        
        status = t.Commit()
        if status == TransactionStatus.Committed:
            tg.Assimilate()
            log.info("Edging Complete.\nSnapped: {}\nAdded: {}".format(len(to_move), len(to_add)))
        else:
            if tg.HasStarted(): tg.RollBack()
            log.error("Revit rejected the changes.", "The Toposolid geometry may have become invalid (e.g., too thin).")
        
    except Exception as e:
        tg.RollBack()
        log.error("Edging failed.", e)
    finally:
        log.show()

class SmoothingNode:
    def __init__(self, x, y, z, v_ref=None):
        self.x = x
        self.y = y
        self.z = z
        self.orig_z = z
        self.v_ref = v_ref
        self.next_z = z

def apply_laplacian_smoothing(nodes, interior_nodes, smoothing_radius, iterations=5, z_preservation_weight=0.0):
    """
    A highly-optimized, multi-threaded Laplacian smoothing engine.
    Applies inverse-distance weighting to pull points towards the average of their neighbors.
    """
    # Identify boundaries and active subset
    boundary_nodes = [n for n in nodes if n not in interior_nodes and any(math.hypot(n.x - i.x, n.y - i.y) < smoothing_radius for i in interior_nodes)]
    active_nodes = interior_nodes + boundary_nodes
    
    # Pre-compute static neighbors and inverse-distance weights (Multi-threaded)
    node_neighbors = {}
    nn_lock = threading.Lock()
    nn_errors = []
    
    def process_nn(chunk):
        try:
            local_nn = {}
            for n in chunk:
                nbs = []
                for a in active_nodes:
                    if a != n:
                        dist = math.hypot(n.x - a.x, n.y - a.y)
                        if dist < smoothing_radius:
                            nbs.append((a, 1.0 / max(dist, 0.001)))
                local_nn[n] = nbs
            with nn_lock:
                node_neighbors.update(local_nn)
        except Exception as e:
            with nn_lock:
                nn_errors.append(e)
            
    if not interior_nodes: return
    chunk_size = max(1, int(len(interior_nodes) / 8.0))
    chunks = [interior_nodes[i:i+chunk_size] for i in range(0, len(interior_nodes), chunk_size)]
    threads = []
    for c in chunks:
        th = threading.Thread(target=process_nn, args=(c,))
        th.start()
        threads.append(th)
    for th in threads: th.join()
    
    if nn_errors:
        raise Exception("Multi-threading execution failed in Laplacian smoothing: {}".format(nn_errors[0]))
    
    # Iterative Laplacian Smoothing
    for _ in range(iterations): 
        for n in interior_nodes:
            neighbors = node_neighbors.get(n, [])
            if neighbors:
                total_weight = sum(w for a, w in neighbors) + z_preservation_weight
                n.next_z = (sum(a.z * w for a, w in neighbors) + (n.orig_z * z_preservation_weight)) / total_weight
        for n in interior_nodes:
            n.z = n.next_z

def perform_smooth_region(state):
    try:
        ref_topo = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid to Smooth")
        raw_topo = doc.GetElement(ref_topo)
        toposolid = resolve_toposolid_host(doc, raw_topo)
        if not isinstance(toposolid, Toposolid):
            forms.alert("First selection must be a Toposolid.", exitscript=True)
    except:
        return # Cancelled
        
    log = BatchLogger()
    
    # Raycast Context (setup once)
    intersector = None
    ray_start_z = get_toposolid_max_z(toposolid) + 100.0
    view3d = doc.ActiveView if doc.ActiveView.ViewType == ViewType.ThreeD else None
    if not view3d:
        for v in FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements():
            if not v.IsTemplate: view3d = v; break
    if view3d:
        intersector = ReferenceIntersector(toposolid.Id, FindReferenceTarget.Element, view3d)

    # --- LOOP START ---
    while True:
        tg = None
        try:
            # Pick one or more Filled Regions
            refs = uidoc.Selection.PickObjects(ObjectType.Element, UniversalFilter(), "Select Filled Regions (Finish when done, ESC to exit)")
            region_elems = [doc.GetElement(r) for r in refs if isinstance(doc.GetElement(r), FilledRegion)]
        except OperationCanceledException:
            break # User pressed ESC to exit loop
        except Exception as e:
            log.error("Selection failed", str(e))
            break

        if not region_elems:
            log.error("No valid Filled Regions selected.")
            continue
            
        log.info("--- SMOOTH REGION STARTED ---")
        
        try:
            # 1. Parameters
            g_int = float(state.grid)
            dynamic_res = max(0.25, g_int) 
            log.info("Grid Resolution: {}".format(UnitHelper.to_formatted_string(dynamic_res)))
            
            # 2. Extract Region Geometries
            region_data = []
            combined_min_x, combined_max_x = float('inf'), float('-inf')
            combined_min_y, combined_max_y = float('inf'), float('-inf')

            for region_elem in region_elems:
                curves = []
                for loop in region_elem.GetBoundaries():
                    for c in loop: curves.append(c)
                
                if not curves: continue

                min_x = min([c.Evaluate(i/10.0, True).X for c in curves for i in range(11)])
                max_x = max([c.Evaluate(i/10.0, True).X for c in curves for i in range(11)])
                min_y = min([c.Evaluate(i/10.0, True).Y for c in curves for i in range(11)])
                max_y = max([c.Evaluate(i/10.0, True).Y for c in curves for i in range(11)])
                
                region_data.append({'curves': curves, 'min_x': min_x, 'max_x': max_x, 'min_y': min_y, 'max_y': max_y})
                
                combined_min_x = min(combined_min_x, min_x)
                combined_max_x = max(combined_max_x, max_x)
                combined_min_y = min(combined_min_y, min_y)
                combined_max_y = max(combined_max_y, max_y)

            def is_inside(x, y):
                for r_data in region_data:
                    if x < r_data['min_x'] or x > r_data['max_x'] or y < r_data['min_y'] or y > r_data['max_y']:
                        continue
                    
                    ints = 0
                    ray_y = y + 0.000137
                    for c in r_data['curves']:
                        c_z = c.GetEndPoint(0).Z
                        ray = Line.CreateBound(XYZ(x, ray_y, c_z), XYZ(r_data['max_x'] + 100.0, ray_y, c_z))
                        res_arr = clr.Reference[IntersectionResultArray]()
                        res = c.Intersect(ray, res_arr)
                        if res == SetComparisonResult.Overlap and res_arr.Value is not None:
                            ints += res_arr.Value.Size
                    if ints % 2 != 0:
                        return True # It's inside at least one region
                return False

            # 3. Laplacian Smoothing Process
            tg = TransactionGroup(doc, "Smooth Region")
            tg.Start()
            
            tracker = VirtualVertexTracker(doc, toposolid)
            
            t = Transaction(doc, "Smooth Points")
            t.Start()
            fail_opts = t.GetFailureHandlingOptions()
            fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
            t.SetFailureHandlingOptions(fail_opts)
            
            nodes = []
            interior_nodes = []
            
            off = get_subdivision_offset(doc, raw_topo.Id) if toposolid.Id != raw_topo.Id else 0.0

            # A. Classify existing points
            for v, pos in tracker.get_all():
                n = SmoothingNode(pos.X, pos.Y, pos.Z, v)
                if is_inside(pos.X, pos.Y): interior_nodes.append(n)
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
                    if abs(n.z - mean_z) > max(2.0 * std_dev, outlier_tol_val):
                        outlier_nodes.append(n)
                        
            if outlier_nodes:
                log.info("Removing {} outlier points before smoothing.".format(len(outlier_nodes)))
                for out_n in outlier_nodes:
                    if out_n in interior_nodes: interior_nodes.remove(out_n)
                    if out_n in nodes: nodes.remove(out_n)
                    if out_n.v_ref:
                        tracker.delete(out_n.v_ref, XYZ(out_n.x, out_n.y, out_n.orig_z))
            
            # B. Densify (Add missing grid points)
            start_x = math.floor(combined_min_x / dynamic_res) * dynamic_res
            start_y = math.floor(combined_min_y / dynamic_res) * dynamic_res
            
            y_vals = []
            y_temp = start_y
            while y_temp <= combined_max_y:
                y_vals.append(y_temp)
                y_temp += dynamic_res
                
            inside_pts = []
            thread_errors = []
            bag_lock = threading.Lock()
            
            def process_row(y):
                row_pts = []
                x = start_x
                while x <= combined_max_x:
                    if is_inside(x, y):
                        row_pts.append((x, y))
                    x += dynamic_res
                with bag_lock:
                    inside_pts.extend(row_pts)
                    
            def chunk_process(chunk):
                try:
                    for y_val in chunk:
                        process_row(y_val)
                except Exception as e:
                    with bag_lock:
                        thread_errors.append(e)
                    
            chunk_size = max(1, int(len(y_vals) / 8.0))
            chunks = [y_vals[i:i+chunk_size] for i in range(0, len(y_vals), chunk_size)]
            threads = []
            for c in chunks:
                th = threading.Thread(target=chunk_process, args=(c,))
                th.start()
                threads.append(th)
            for th in threads: th.join()
            
            if thread_errors:
                raise Exception("Multi-threading execution failed: {}".format(thread_errors[0]))

            # Process the identified inside points on the Main Thread safely
            for px, py in inside_pts:
                if not any(math.hypot(n.x - px, n.y - py) < dynamic_res * 0.4 for n in interior_nodes):
                    pt = XYZ(px, py, 0)
                    rz = None
                    if intersector:
                        z_hit, _ = get_surface_info(intersector, pt, ray_start_z)
                        if z_hit is not None: rz = z_hit - off
                    
                    if rz is None:
                        if nodes:
                            nearest = min(nodes, key=lambda nd: math.hypot(nd.x - px, nd.y - py))
                            rz = nearest.z
                        else:
                            rz = 0.0
                            
                    n = SmoothingNode(px, py, rz)
                    interior_nodes.append(n)
                    nodes.append(n)
            
            # B.2 Densify Boundary anchors
            smoothing_radius = dynamic_res * 1.5
            potential_boundary_pts = set()
            for n in interior_nodes:
                for dx in [-dynamic_res, 0, dynamic_res]:
                    for dy in [-dynamic_res, 0, dynamic_res]:
                        if dx == 0 and dy == 0: continue
                        bx = round((n.x + dx) / dynamic_res) * dynamic_res
                        by = round((n.y + dy) / dynamic_res) * dynamic_res
                        potential_boundary_pts.add((bx, by))
            
            for bx, by in potential_boundary_pts:
                if not any(math.hypot(n.x - bx, n.y - by) < dynamic_res * 0.4 for n in nodes):
                    if not is_inside(bx, by):
                        pt = XYZ(bx, by, 0)
                        rz = None
                        if intersector:
                            z_hit, _ = get_surface_info(intersector, pt, ray_start_z)
                            if z_hit is not None: rz = z_hit - off
                        
                        if rz is None:
                            nearest = min(nodes, key=lambda nd: math.hypot(nd.x - bx, nd.y - by))
                            rz = nearest.z
                            
                        n = SmoothingNode(bx, by, rz)
                        nodes.append(n)

            # C. Identify boundaries and Smooth
            boundary_nodes = [n for n in nodes if n not in interior_nodes and any(math.hypot(n.x - i.x, n.y - i.y) < smoothing_radius for i in interior_nodes)]
            active_nodes = interior_nodes + boundary_nodes
            
            # Pre-compute static neighbors and inverse-distance weights
            node_neighbors = {}
            for n in interior_nodes:
                nbs = []
                for a in active_nodes:
                    if a != n:
                        dist = math.hypot(n.x - a.x, n.y - a.y)
                        if dist < smoothing_radius:
                            nbs.append((a, 1.0 / max(dist, 0.001)))
                node_neighbors[n] = nbs
            
            # Z-preservation weight (0.0 = full smoothing, higher values = more retention of original shape)
            z_preservation_weight = 0.0
            
            for _ in range(5):
                for n in interior_nodes:
                    neighbors = node_neighbors[n]
                    if neighbors:
                        total_weight = sum(w for a, w in neighbors) + z_preservation_weight
                        n.next_z = (sum(a.z * w for a, w in neighbors) + (n.orig_z * z_preservation_weight)) / total_weight
                for n in interior_nodes:
                    n.z = n.next_z
            
            # D. Apply to Toposolid
            add_count, mod_count = 0, 0
            for n in interior_nodes:
                final_pt = XYZ(n.x, n.y, n.z)
                if n.v_ref:
                    if abs(n.orig_z - n.z) > 0.005:
                        if tracker.delete(n.v_ref, XYZ(n.x, n.y, n.orig_z)):
                            if tracker.add(final_pt):
                                mod_count += 1
                else:
                    if tracker.add(final_pt):
                        add_count += 1
                    
            status = t.Commit()
            if status == TransactionStatus.Committed:
                tg.Assimilate()
                log.info("Smooth Region Complete.\nAdded {} new points, Smoothed {} existing points.\n".format(add_count, mod_count))
            else:
                if tg.HasStarted(): tg.RollBack()
                log.error("Revit rejected the changes.", "The Toposolid geometry may have become invalid (e.g., too thin).")
            
        except Exception as loop_e:
            if 't' in locals() and t.HasStarted(): t.RollBack()
            if tg is not None and tg.HasStarted(): tg.RollBack()
            log.error("Smooth Region Iteration Failed", "{}\n{}".format(loop_e, traceback.format_exc()))
            break # Prevent infinite selection loop on failure

    log.show()

def perform_smooth_path(state):
    log = BatchLogger()
    try:
        ref_topo = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid to Smooth")
        raw_topo = doc.GetElement(ref_topo)
        ref_lines = uidoc.Selection.PickObjects(ObjectType.Element, UniversalFilter(), "Select Guide Lines (Tab for chain, Finish when done)")
        line_elems = [doc.GetElement(ref) for ref in ref_lines]
    except OperationCanceledException:
        return # Cancelled
    except Exception as e:
        log.error("Selection failed", str(e))
        log.show()
        return
    
    try:
        # 1. Elements
        toposolid = resolve_toposolid_host(doc, raw_topo)
        if not isinstance(toposolid, Toposolid):
            log.error("First selection is not a Toposolid.")
            log.show(); return
            
        curves = []
        for elem in line_elems:
            if isinstance(elem, CurveElement):
                curves.append(elem.GeometryCurve)
                
        if not curves:
            log.error("No valid Lines/Curves selected.")
            log.show(); return
        
        # 2. Parameters
        w_int = float(state.width)
        f_int = float(state.falloff)
        g_int = float(state.grid)
        core_rad = w_int / 2.0
        total_rad = core_rad + f_int
        
        if getattr(state, "sharp_smoothing", False):
            total_rad = core_rad
            
        if getattr(state, "apply_plan_offset", False):
            try: total_rad += abs(float(getattr(state, "plan_offset_val", "0.0")))
            except: pass
        
        dynamic_res = max(0.25, g_int) 
        log.info("Smoothing Width: {} (Core: {}, Falloff: {})".format(
            UnitHelper.to_formatted_string(w_int if getattr(state, "sharp_smoothing", False) else w_int + 2*f_int),
            UnitHelper.to_formatted_string(w_int),
            UnitHelper.to_formatted_string(0.0 if getattr(state, "sharp_smoothing", False) else f_int)
        ))
        
        # Square Ends logic for chains
        check_square_ends = False
        sq_ends_info = []
        if state.square_ends:
            try:
                end_map = []
                for c in curves:
                    p0 = c.GetEndPoint(0)
                    p1 = c.GetEndPoint(1)
                    if isinstance(c, Line):
                        vec = (p1 - p0)
                        tan_start_outward = vec.Normalize().Negate()
                        tan_end_outward = vec.Normalize()
                    else:
                        t0 = c.GetEndParameter(0)
                        t1 = c.GetEndParameter(1)
                        tan_start_outward = c.ComputeDerivatives(t0, False).BasisX.Normalize().Negate()
                        tan_end_outward = c.ComputeDerivatives(t1, False).BasisX.Normalize()
                        
                    end_map.append({'pt': p0, 'tan': tan_start_outward})
                    end_map.append({'pt': p1, 'tan': tan_end_outward})
                    
                # Find open ends
                open_ends = []
                for i, e1 in enumerate(end_map):
                    matched = False
                    for j, e2 in enumerate(end_map):
                        if i != j and e1['pt'].DistanceTo(e2['pt']) < 0.01:
                            matched = True
                            break
                    if not matched:
                        open_ends.append(e1)
                        
                if open_ends:
                    check_square_ends = True
                    for ep in open_ends:
                        tan_xy = XYZ(ep['tan'].X, ep['tan'].Y, 0)
                        if not tan_xy.IsZeroLength():
                            sq_ends_info.append({'pt': ep['pt'], 'tan': tan_xy.Normalize()})
            except Exception as e: pass

        def is_outside_bounds(pt):
            if not check_square_ends: return False
            for ep in sq_ends_info:
                v = XYZ(pt.X - ep['pt'].X, pt.Y - ep['pt'].Y, 0)
                if v.DotProduct(ep['tan']) > 0.001: 
                    return True
            return False

        def is_inside(pt):
            if is_outside_bounds(pt): return False
            min_d = float('inf')
            for c in curves:
                p_3d, t_norm, d_xy = project_2d(c, pt)
                if d_xy < min_d: min_d = d_xy
            return min_d <= total_rad
            
        # 3. Raycast Context
        intersector = None
        ray_start_z = get_toposolid_max_z(toposolid) + 100.0
        view3d = doc.ActiveView if doc.ActiveView.ViewType == ViewType.ThreeD else None
        if not view3d:
            for v in FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements():
                if not v.IsTemplate: view3d = v; break
        if view3d:
            intersector = ReferenceIntersector(toposolid.Id, FindReferenceTarget.Element, view3d)
        
        # 4. Laplacian Smoothing Process
        tg = TransactionGroup(doc, "Smooth Path")
        tg.Start()
        
        tracker = VirtualVertexTracker(doc, toposolid)
        
        t = Transaction(doc, "Smooth Points")
        t.Start()
        fail_opts = t.GetFailureHandlingOptions()
        fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
        t.SetFailureHandlingOptions(fail_opts)
        
        nodes = []
        interior_nodes = []
        off = get_subdivision_offset(doc, raw_topo.Id) if toposolid.Id != raw_topo.Id else 0.0

        # A. Classify existing points
        for v, pos in tracker.get_all():
            n = SmoothingNode(pos.X, pos.Y, pos.Z, v)
            if is_inside(pos): interior_nodes.append(n)
            nodes.append(n)
            
        # A.1 Outlier Detection
        try: outlier_tol_val = float(state.outlier_tol)
        except: outlier_tol_val = 1.0
        
        outlier_nodes = []
        outlier_radius = dynamic_res * 2.0
        outlier_lock = threading.Lock()
        outlier_errors = []
        
        def process_outliers(chunk):
            try:
                local_outliers = []
                for n in chunk:
                    neighbors = [nb for nb in nodes if nb != n and math.hypot(n.x - nb.x, n.y - nb.y) < outlier_radius]
                    if len(neighbors) >= 3:
                        mean_z = sum(nb.z for nb in neighbors) / len(neighbors)
                        variance = sum((nb.z - mean_z) ** 2 for nb in neighbors) / len(neighbors)
                        std_dev = math.sqrt(max(0.0, variance))
                        if abs(n.z - mean_z) > max(2.0 * std_dev, outlier_tol_val):
                            local_outliers.append(n)
                with outlier_lock:
                    outlier_nodes.extend(local_outliers)
            except Exception as e:
                with outlier_lock:
                    outlier_errors.append(e)
                
        chunk_size = max(1, int(len(interior_nodes) / 8.0))
        chunks = [interior_nodes[i:i+chunk_size] for i in range(0, len(interior_nodes), chunk_size)]
        threads = []
        for c in chunks:
            th = threading.Thread(target=process_outliers, args=(c,))
            th.start()
            threads.append(th)
        for th in threads: th.join()
                    
        if outlier_errors:
            raise Exception("Multi-threading execution failed in outliers: {}".format(outlier_errors[0]))

        if outlier_nodes:
            log.info("Removing {} outlier points before smoothing.".format(len(outlier_nodes)))
            for out_n in outlier_nodes:
                if out_n in interior_nodes: interior_nodes.remove(out_n)
                if out_n in nodes: nodes.remove(out_n)
                if out_n.v_ref:
                    tracker.delete(out_n.v_ref, XYZ(out_n.x, out_n.y, out_n.orig_z))
        
        # B. Densify (Add missing grid points)
        min_x = min_y = float('inf')
        max_x = max_y = float('-inf')
        for elem in line_elems:
            if isinstance(elem, CurveElement):
                bb = elem.get_BoundingBox(None)
                if bb:
                    min_x = min(min_x, bb.Min.X)
                    min_y = min(min_y, bb.Min.Y)
                    max_x = max(max_x, bb.Max.X)
                    max_y = max(max_y, bb.Max.Y)
                    
        start_x = math.floor((min_x - total_rad - g_int) / g_int) * g_int
        end_x   = math.ceil((max_x + total_rad + g_int) / g_int) * g_int
        start_y = math.floor((min_y - total_rad - g_int) / g_int) * g_int
        end_y   = math.ceil((max_y + total_rad + g_int) / g_int) * g_int
        
        x = start_x
        while x <= end_x:
            y = start_y
            while y <= end_y:
                pt = XYZ(x, y, 0)
                if is_inside(pt):
                    if not any(math.hypot(n.x - x, n.y - y) < dynamic_res * 0.4 for n in interior_nodes):
                        rz = None
                        if intersector:
                            z_hit, _ = get_surface_info(intersector, pt, ray_start_z)
                            if z_hit is not None: rz = z_hit - off
                        
                        if rz is None:
                            if nodes:
                                nearest = min(nodes, key=lambda nd: math.hypot(nd.x - x, nd.y - y))
                                rz = nearest.z
                            else:
                                rz = 0.0
                                
                        n = SmoothingNode(x, y, rz)
                        interior_nodes.append(n)
                        nodes.append(n)
                y += dynamic_res
            x += dynamic_res
        
        # B.2 Densify Boundary anchors
        smoothing_radius = dynamic_res * 1.5
        potential_boundary_pts = set()
        for n in interior_nodes:
            for dx in [-dynamic_res, 0, dynamic_res]:
                for dy in [-dynamic_res, 0, dynamic_res]:
                    if dx == 0 and dy == 0: continue
                    bx = round((n.x + dx) / dynamic_res) * dynamic_res
                    by = round((n.y + dy) / dynamic_res) * dynamic_res
                    potential_boundary_pts.add((bx, by))
        
        for bx, by in potential_boundary_pts:
            if not any(math.hypot(n.x - bx, n.y - by) < dynamic_res * 0.4 for n in nodes):
                pt = XYZ(bx, by, 0)
                if not is_inside(pt):
                    rz = None
                    if intersector:
                        z_hit, _ = get_surface_info(intersector, pt, ray_start_z)
                        if z_hit is not None: rz = z_hit - off
                    
                    if rz is None:
                        nearest = min(nodes, key=lambda nd: math.hypot(nd.x - bx, nd.y - by))
                        rz = nearest.z
                        
                    n = SmoothingNode(bx, by, rz)
                    nodes.append(n)

        # C. Identify boundaries and Smooth
        boundary_nodes = [n for n in nodes if n not in interior_nodes and any(math.hypot(n.x - i.x, n.y - i.y) < smoothing_radius for i in interior_nodes)]
        active_nodes = interior_nodes + boundary_nodes
        
        # Pre-compute static neighbors and inverse-distance weights
        node_neighbors = {}
        nn_lock = threading.Lock()
        nn_errors = []
        
        def process_nn(chunk):
            try:
                local_nn = {}
                for n in chunk:
                    nbs = []
                    for a in active_nodes:
                        if a != n:
                            dist = math.hypot(n.x - a.x, n.y - a.y)
                            if dist < smoothing_radius:
                                nbs.append((a, 1.0 / max(dist, 0.001)))
                    local_nn[n] = nbs
                with nn_lock:
                    node_neighbors.update(local_nn)
            except Exception as e:
                with nn_lock:
                    nn_errors.append(e)
                
        chunk_size = max(1, int(len(interior_nodes) / 8.0))
        chunks = [interior_nodes[i:i+chunk_size] for i in range(0, len(interior_nodes), chunk_size)]
        threads = []
        for c in chunks:
            th = threading.Thread(target=process_nn, args=(c,))
            th.start()
            threads.append(th)
        for th in threads: th.join()
        
        if nn_errors:
            raise Exception("Multi-threading execution failed in nearest neighbors: {}".format(nn_errors[0]))
        
        # Z-preservation weight (0.0 = full smoothing, higher values = more retention of original shape)
        z_preservation_weight = 0.0
        
        for _ in range(5): 
            for n in interior_nodes:
                neighbors = node_neighbors[n]
                if neighbors:
                    total_weight = sum(w for a, w in neighbors) + z_preservation_weight
                    n.next_z = (sum(a.z * w for a, w in neighbors) + (n.orig_z * z_preservation_weight)) / total_weight
            for n in interior_nodes:
                n.z = n.next_z
        
        # D. Apply to Toposolid
        add_count, mod_count = 0, 0
        for n in interior_nodes:
            final_pt = XYZ(n.x, n.y, n.z)
            if n.v_ref:
                if abs(n.orig_z - n.z) > 0.005:
                    if tracker.delete(n.v_ref, XYZ(n.x, n.y, n.orig_z)):
                        if tracker.add(final_pt):
                            mod_count += 1
            else:
                if tracker.add(final_pt):
                    add_count += 1
                
        status = t.Commit()
        if status == TransactionStatus.Committed:
            tg.Assimilate()
            log.info("Smooth Path Complete.\nAdded {} new points, Smoothed {} existing points.".format(add_count, mod_count))
        else:
            if tg.HasStarted(): tg.RollBack()
            log.error("Revit rejected the changes.", "The Toposolid geometry may have become invalid (e.g., too thin).")
        
    except Exception as e:
        if 'tg' in locals() and tg.HasStarted(): tg.RollBack()
        log.error("Smooth Path Failed", "{}\n{}".format(e, traceback.format_exc()))
    finally:
        log.show()

def get_region_z(px, py, region_infos):
    if not region_infos: return None
    for r_info in region_infos:
        tess_segs, reg_z, max_x, min_x, min_y, max_y = r_info
        if min_x <= px <= max_x and min_y <= py <= max_y:
            intersections = 0
            for (ax, ay), (bx, by) in tess_segs:
                if ((ay > py) != (by > py)):
                    intersect_x = (bx - ax) * (py - ay) / (by - ay) + ax
                    if px < intersect_x: intersections += 1
            if intersections % 2 != 0: return reg_z
    return None

def is_in_graded_region(px, py, target_region_infos, target_curves, p_off, p_dir):
    if not target_region_infos: return False
    is_in_orig = get_region_z(px, py, target_region_infos) is not None
    if p_off == 0.0 or p_dir == "Center": return is_in_orig
    
    min_d = float('inf')
    pt = XYZ(px, py, 0)
    for c, override_z in target_curves:
        if override_z is not None:
            _, _, d_xy = project_2d(c, pt)
            if d_xy < min_d: min_d = d_xy
            
    if p_dir == "Outside" or p_dir == "Both": return is_in_orig or (min_d <= p_off - 0.01)
    elif p_dir == "Inside": return is_in_orig and (min_d >= p_off + 0.01)
    return is_in_orig

def apply_triangulation_halo(state, doc, tracker, target_curves, target_region_infos, 
                             intersector, ray_start_z, off, g_int, g_plan_off, 
                             g_plan_dir, g_z_off, pts_grid, unique_pts, cell_size):
    eliminated_count = 0
    halo_added = 0
    halo_offset_inner = float(getattr(state, "halo_inner", "0.25"))
    halo_offset_outer = g_int # Automatically bounds to the outer grid resolution
    halo_search_rad = max(2.0, g_int * 1.5)
    halo_pts = []
    
    # Tolerance buffer to reject wildly deviating raycasts
    try: raycast_tol = float(getattr(state, "outlier_tol", "1.0")) * 3.0
    except: raycast_tol = 3.0

    # Build a coarser spatial hash specifically for the halo search
    halo_cell_size = halo_search_rad
    halo_verts_grid = {}
    for v, pos in tracker.get_all():
        try:
            hcx = int(math.floor(pos.X / halo_cell_size))
            hcy = int(math.floor(pos.Y / halo_cell_size))
            halo_verts_grid.setdefault((hcx, hcy), []).append(pos)
        except: pass
    
    # 1. Generate Dual-Ring Halo Points using Continuous Offset Geometry
    chains = []
    current_chain = []
    current_z = None
    
    for curve, override_z in target_curves:
        if not current_chain:
            current_chain.append(curve)
            current_z = override_z
        else:
            last_curve = current_chain[-1]
            if last_curve.GetEndPoint(1).DistanceTo(curve.GetEndPoint(0)) < 0.01 and override_z == current_z:
                current_chain.append(curve)
            else:
                chains.append((current_chain, current_z))
                current_chain = [curve]
                current_z = override_z
    if current_chain:
        chains.append((current_chain, current_z))

    for chain, override_z in chains:
        is_region = (override_z is not None and target_region_infos)
        
        pts = []
        for curve in chain:
            tess = curve.Tessellate()
            if not pts: pts.append(tess[0])
            for pt in tess[1:]:
                if pts[-1].DistanceTo(pt) > 0.001:
                    pts.append(pt)
        
        offset_lines = []
        active_offsets = [0.0]
        if g_plan_off != 0.0:
            if g_plan_dir == "Inside": active_offsets = [g_plan_off]
            elif g_plan_dir == "Outside": active_offsets = [-g_plan_off]
            elif g_plan_dir == "Both": active_offsets = [g_plan_off, -g_plan_off]
            elif g_plan_dir == "Center": active_offsets = [g_plan_off / 2.0, -g_plan_off / 2.0]
            
        for a_off in active_offsets:
            offset_lines.append(generate_offset_polyline(pts, a_off + halo_offset_inner))
            offset_lines.append(generate_offset_polyline(pts, a_off - halo_offset_inner))
            offset_lines.append(generate_offset_polyline(pts, a_off + halo_offset_outer))
            offset_lines.append(generate_offset_polyline(pts, a_off - halo_offset_outer))
        
        valid_halos = []
        for off_line in offset_lines:
            for i in range(len(off_line) - 1):
                p_start = off_line[i]
                p_end = off_line[i+1]
                vec = p_end - p_start
                
                vec_xy = XYZ(vec.X, vec.Y, 0)
                length = vec_xy.GetLength()
                if length < 0.001: continue
                
                steps = int(math.ceil(length / g_int))
                for s in range(steps + 1):
                    t = (s * g_int) / length
                    if t > 1.0: t = 1.0
                    hp = p_start + vec * t
                    
                    if is_region and is_in_graded_region(hp.X, hp.Y, target_region_infos, target_curves, g_plan_off, g_plan_dir):
                        continue
                        
                    valid_halos.append(hp)
        
        for hp in valid_halos:
            hcx_halo = int(math.floor(hp.X / halo_cell_size))
            hcy_halo = int(math.floor(hp.Y / halo_cell_size))
            
            nearby_z = []
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for pos in halo_verts_grid.get((hcx_halo + dx, hcy_halo + dy), []):
                        if dist_2d(pos, hp) < halo_search_rad:
                            if is_region and get_region_z(pos.X, pos.Y, target_region_infos) is not None:
                                continue
                            nearby_z.append(pos.Z)
                            
            avg_z = sum(nearby_z) / len(nearby_z) if nearby_z else (override_z if override_z is not None else hp.Z) + g_z_off - off
            
            hz = None
            if intersector:
                z_hit, hit_id = get_surface_info(intersector, hp, ray_start_z)
                if z_hit is not None:
                    ray_z = z_hit - get_subdivision_offset(doc, hit_id)
                    if abs(ray_z - avg_z) <= raycast_tol:
                        hz = ray_z
                        
            if hz is None:
                hz = avg_z
                
            halo_pts.append(XYZ(hp.X, hp.Y, hz))
            
    # 2. Eliminate existing points strictly within the optimized boundary zone
    for v, pos in tracker.get_all():
            is_region = False
            if target_region_infos:
                if is_in_graded_region(pos.X, pos.Y, target_region_infos, target_curves, g_plan_off, g_plan_dir):
                    is_region = True
            if is_region: continue
            
            conflict_with_new = False
            pcx = int(math.floor(pos.X / cell_size))
            pcy = int(math.floor(pos.Y / cell_size))
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for idx in pts_grid.get((pcx + dx, pcy + dy), []):
                        if dist_2d(unique_pts[idx], pos) < MIN_DIST_TOLERANCE:
                            conflict_with_new = True
                            break
                    if conflict_with_new: break
                if conflict_with_new: break
                
            if conflict_with_new: continue
            
            min_d = float('inf')
            for curve, override_z in target_curves:
                _, _, d_xy = project_2d(curve, pos)
                
                if g_plan_off != 0.0 and g_plan_dir != "Center" and override_z is None:
                    vec_to_pt = XYZ(pos.X - p_3d.X, pos.Y - p_3d.Y, 0)
                    try:
                        t_fwd = curve.ComputeDerivatives(t_norm, True).BasisX.Normalize()
                        normal = XYZ(0,1,0) if t_fwd.IsAlmostEqualTo(XYZ.Zero) else XYZ.BasisZ.CrossProduct(t_fwd).Normalize()
                        side_sign = 1.0 if vec_to_pt.DotProduct(normal) >= 0 else -1.0
                    except: side_sign = 1.0
                    if g_plan_dir == "Inside": d_xy = abs((side_sign * d_xy) - g_plan_off)
                    elif g_plan_dir == "Outside": d_xy = abs((side_sign * d_xy) + g_plan_off)
                    elif g_plan_dir == "Both": d_xy = min(abs(d_xy - g_plan_off), abs(d_xy + g_plan_off))
                elif g_plan_off != 0.0 and override_z is not None:
                    if g_plan_dir == "Outside" or g_plan_dir == "Both": d_xy = max(0.0, d_xy - g_plan_off)
                    elif g_plan_dir == "Inside": d_xy = d_xy + g_plan_off
                    
                if d_xy < min_d: min_d = d_xy
                
            if min_d <= halo_offset_outer + 0.1:
                if tracker.delete(v, pos): eliminated_count += 1

    # 3. Snapshot valid existing points post-deletion to avoid false conflicts
    valid_verts_grid = {}
    for v, pos in tracker.get_all():
        try:
            hcx = int(math.floor(pos.X / cell_size))
            hcy = int(math.floor(pos.Y / cell_size))
            valid_verts_grid.setdefault((hcx, hcy), []).append(pos)
        except: pass
    
    # 4. Add Halo Points
    for hp in halo_pts:
        cx = int(math.floor(hp.X / cell_size))
        cy = int(math.floor(hp.Y / cell_size))
        
        conflict = False
        # Check against the core grading points
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for idx in pts_grid.get((cx + dx, cy + dy), []):
                    if dist_2d(unique_pts[idx], hp) < MIN_DIST_TOLERANCE:
                        conflict = True
                        break
                if conflict: break
                
        # Check against the remaining valid ground points
        if not conflict:
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for pos in valid_verts_grid.get((cx + dx, cy + dy), []):
                        if dist_2d(pos, hp) < MIN_DIST_TOLERANCE:
                            conflict = True
                            break
                    if conflict: break
                    
        if not conflict:
            if tracker.add(hp):
                unique_pts.append(hp)
                pts_grid.setdefault((cx, cy), []).append(len(unique_pts) - 1)
                halo_added += 1

    return eliminated_count, halo_added

def perform_add_points_along_line(state):
    log = BatchLogger()
    
    g_int = float(state.grid)
    if g_int <= 0.1: g_int = 1.0 # Fallback safety
    
    g_plan_off = 0.0
    g_plan_dir = "Both"
    if getattr(state, "apply_plan_offset", False):
        try:
            g_plan_off = float(getattr(state, "plan_offset_val", "0.0"))
            g_plan_dir = getattr(state, "plan_offset_dir", "Both")
        except: pass
    
    def get_curves_from_elem(elem, g_int, p_off=0.0, p_dir="Both"):
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
                tessellated_segments = []
                for c in all_curves:
                    pts = c.Tessellate()
                    for i in range(len(pts) - 1):
                        p1, p2 = pts[i], pts[i+1]
                        tessellated_segments.append(((p1.X, p1.Y), (p2.X, p2.Y)))
                        min_x = min(min_x, p1.X, p2.X)
                        max_x = max(max_x, p1.X, p2.X)
                        min_y = min(min_y, p1.Y, p2.Y)
                        max_y = max(max_y, p1.Y, p2.Y)
                
                search_pad = p_off + g_int if (p_off > 0 and p_dir in ["Outside", "Both"]) else g_int
                start_x = math.floor((min_x - search_pad) / g_int) * g_int
                end_x   = math.ceil((max_x + search_pad) / g_int) * g_int
                start_y = math.floor((min_y - search_pad) / g_int) * g_int
                end_y   = math.ceil((max_y + search_pad) / g_int) * g_int
                
                y_vals = []
                y_temp = start_y
                while y_temp <= end_y:
                    y_vals.append(y_temp)
                    y_temp += g_int
                    
                bag = []
                bag_lock = threading.Lock()
                thread_errors = []
                
                def process_row(y):
                    row_pts = []
                    x = start_x
                    while x <= end_x:
                        intersections = 0
                        for (ax, ay), (bx, by) in tessellated_segments:
                            if ((ay > y) != (by > y)):
                                intersect_x = (bx - ax) * (y - ay) / (by - ay) + ax
                                if x < intersect_x: intersections += 1
                                
                        is_in_orig = (intersections % 2 != 0)
                        if p_off != 0.0 and p_dir != "Center":
                            pt_3d = XYZ(x, y, z_val)
                            min_d = float('inf')
                            for c in all_curves:
                                _, _, d_xy = project_2d(c, pt_3d)
                                if d_xy < min_d: min_d = d_xy
                                
                            if p_dir == "Outside" or p_dir == "Both":
                                if is_in_orig or min_d <= p_off + 0.01: row_pts.append((x, y))
                            elif p_dir == "Inside":
                                if is_in_orig and min_d >= p_off + 0.01: row_pts.append((x, y))
                        else:
                            if is_in_orig:
                                row_pts.append((x, y))
                                
                        x += g_int
                    with bag_lock:
                        bag.extend(row_pts)
                        
                def chunk_process(chunk):
                    try:
                        for y_val in chunk:
                            process_row(y_val)
                    except Exception as e:
                        with bag_lock:
                            thread_errors.append(e)
                        
                chunk_size = max(1, int(len(y_vals) / 8.0))
                chunks = [y_vals[i:i+chunk_size] for i in range(0, len(y_vals), chunk_size)]
                threads = []
                for c in chunks:
                    th = threading.Thread(target=chunk_process, args=(c,))
                    th.start()
                    threads.append(th)
                for th in threads: th.join()
                
                if thread_errors:
                    raise Exception("Multi-threading execution failed in add points: {}".format(thread_errors[0]))
                
                # Instantiate Revit API XYZ objects safely back on the Main Thread
                for px, py in bag:
                    internal_pts.append(XYZ(px, py, z_val))
                    
                region_info = (tessellated_segments, z_val, max_x, min_x, min_y, max_y)
        return crvs, internal_pts, region_info

    try:
        ref_topo = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Toposolid to modify")
        raw_elem = doc.GetElement(ref_topo)
        toposolid = resolve_toposolid_host(doc, raw_elem)
        if not isinstance(toposolid, Toposolid):
            log.error("First selection is not a Toposolid.")
            log.show(); return
            try: 
                uidoc.Selection.SetElementIds(List[ElementId]([ref_topo.ElementId]))
            except: 
                pass
    except: 
        return # Cancelled
    
    while True:
        tg = None
        try:
            ref_elems = uidoc.Selection.PickObjects(ObjectType.Element, UniversalFilter(), "Select Model Lines or Filled Regions (Tab for chain, Finish when done, ESC to exit)")
            target_elems = [doc.GetElement(ref) for ref in ref_elems]
        except: 
            break
        
        log.info("--- GRADE WITH ELEMENTS STARTED ---")
        if getattr(state, "apply_plan_offset", False) and g_plan_off != 0.0:
            log.info("Applied Plan Offset: {} ({})".format(UnitHelper.to_formatted_string(g_plan_off), g_plan_dir))

        target_curves = []
        target_internal_pts = []
        target_region_infos = []
        
        for elem in target_elems:
            crvs, int_pts, r_info = get_curves_from_elem(elem, g_int, g_plan_off, g_plan_dir)
            if crvs: target_curves.extend(crvs)
            if int_pts: target_internal_pts.extend(int_pts)
            if r_info: target_region_infos.append(r_info)
            
        if not target_curves:
            log.error("No valid lines or filled regions selected.")
            continue
            
        # Build connection map for corners
        endpoints_map = {}
        for idx, (curve, _) in enumerate(target_curves):
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            k0 = (round(p0.X, 3), round(p0.Y, 3), round(p0.Z, 3))
            k1 = (round(p1.X, 3), round(p1.Y, 3), round(p1.Z, 3))
            endpoints_map.setdefault(k0, []).append((idx, 0))
            endpoints_map.setdefault(k1, []).append((idx, 1))

        def get_tangent(curve, param):
            try:
                deriv = curve.ComputeDerivatives(param, True)
                return deriv.BasisX.Normalize()
            except:
                return XYZ(1, 0, 0)

        def intersect_2d(p1, d1, p2, d2):
            det = d1.X * d2.Y - d1.Y * d2.X
            if abs(det) < 1e-6: return None
            dx = p2.X - p1.X
            dy = p2.Y - p1.Y
            t_val = (dx * d2.Y - dy * d2.X) / det
            return XYZ(p1.X + d1.X * t_val, p1.Y + d1.Y * t_val, p1.Z)
            
        g_z_off = 0.0
        if state.apply_offset:
            try:
                g_z_off = float(state.offset_val)
                log.info("Applied Z-Offset: {}".format(UnitHelper.to_formatted_string(g_z_off)))
            except: pass
        
        intersector, ray_start_z = setup_toposolid_intersector(doc, toposolid, log)

        tg = TransactionGroup(doc, "Points Along Line")
        tg.Start()
        
        try:
            tracker = VirtualVertexTracker(doc, toposolid)
            
            t = Transaction(doc, "Add Points")
            t.Start()
            fail_opts = t.GetFailureHandlingOptions()
            fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
            t.SetFailureHandlingOptions(fail_opts)
            
            if state.reset_mode:
                log.info("Reset Mode: Removing existing points in target area.")
                to_delete = []
                for v, pos in tracker.get_all():
                    should_delete = False
                    if is_in_graded_region(pos.X, pos.Y, target_region_infos, target_curves, g_plan_off, g_plan_dir):
                        should_delete = True
                        
                    if not should_delete:
                        for curve, override_z in target_curves:
                            if override_z is not None: continue
                            
                            p_3d, t_norm, d_xy = project_2d(curve, pos)
                            if g_plan_off != 0.0 and g_plan_dir != "Center":
                                vec_to_pt = XYZ(pos.X - p_3d.X, pos.Y - p_3d.Y, 0)
                                try:
                                    t_fwd = curve.ComputeDerivatives(t_norm, True).BasisX.Normalize()
                                    normal = XYZ(0,1,0) if t_fwd.IsAlmostEqualTo(XYZ.Zero) else XYZ.BasisZ.CrossProduct(t_fwd).Normalize()
                                    side_sign = 1.0 if vec_to_pt.DotProduct(normal) >= 0 else -1.0
                                except: side_sign = 1.0
                                
                                active_offset = g_plan_off if g_plan_dir == "Inside" else -g_plan_off
                                if g_plan_dir == "Both":
                                    if abs(d_xy - g_plan_off) < (g_int * 0.75) or abs(d_xy + g_plan_off) < (g_int * 0.75):
                                        should_delete = True
                                        break
                                else:
                                    dist_to_shifted = abs((side_sign * d_xy) - active_offset)
                                    if dist_to_shifted < (g_int * 0.75):
                                        should_delete = True
                                        break
                            else:
                                if d_xy < (g_int * 0.75):
                                    should_delete = True
                                    break
                    if should_delete:
                        to_delete.append((v, pos))
                if to_delete:
                    del_count = 0
                    for v, pos in to_delete:
                        if tracker.delete(v, pos):
                            del_count += 1
                    log.info("Removed {} existing points.".format(del_count))

            pts_to_add = []
            if g_z_off != 0.0:
                pts_to_add.extend([XYZ(p.X, p.Y, p.Z + g_z_off) for p in target_internal_pts])
            else:
                pts_to_add.extend(target_internal_pts)
            
            intersection_cache = {}
            def get_cached_intersection(c_idx, end, off_val, override_z):
                key = (c_idx, end, off_val)
                if key in intersection_cache:
                    return intersection_cache[key]
                
                curve, _ = target_curves[c_idx]
                pt = curve.GetEndPoint(end)
                k = (round(pt.X, 3), round(pt.Y, 3), round(pt.Z, 3))
                connections = endpoints_map.get(k, [])
                unique_connections = []
                seen_idx = set()
                for conn in connections:
                    if conn[0] not in seen_idx:
                        seen_idx.add(conn[0])
                        unique_connections.append(conn)
                        
                if len(unique_connections) == 2:
                    idx1, end1 = unique_connections[0]
                    idx2, end2 = unique_connections[1]
                    c1, _ = target_curves[idx1]
                    c2, _ = target_curves[idx2]
                    t1_fwd = get_tangent(c1, 0.0 if end1 == 0 else 1.0)
                    t2_fwd = get_tangent(c2, 0.0 if end2 == 0 else 1.0)
                    n1 = XYZ(0, 1, 0) if t1_fwd.IsAlmostEqualTo(XYZ.Zero) else XYZ.BasisZ.CrossProduct(t1_fwd).Normalize()
                    n2 = XYZ(0, 1, 0) if t2_fwd.IsAlmostEqualTo(XYZ.Zero) else XYZ.BasisZ.CrossProduct(t2_fwd).Normalize()
                    base_z = override_z if override_z is not None else pt.Z
                    base_pt = XYZ(pt.X, pt.Y, base_z + g_z_off)
                    p1 = base_pt + n1.Multiply(off_val)
                    p2 = base_pt + n2.Multiply(off_val)
                    I = intersect_2d(p1, t1_fwd, p2, t2_fwd)
                    if I:
                        res_I = XYZ(I.X, I.Y, base_pt.Z)
                        intersection_cache[key] = res_I
                        return res_I
                intersection_cache[key] = None
                return None

            def add_curve_pts(curve_idx, param, override_z):
                curve, _ = target_curves[curve_idx]
                pt = curve.Evaluate(param, True)
                t_fwd = get_tangent(curve, param)
                if t_fwd.IsAlmostEqualTo(XYZ.Zero):
                    normal = XYZ(0, 1, 0)
                else:
                    normal = XYZ.BasisZ.CrossProduct(t_fwd).Normalize()
                    
                base_z = override_z if override_z is not None else pt.Z
                base_pt = XYZ(pt.X, pt.Y, base_z + g_z_off)
                res = [base_pt]
                
                if g_plan_off != 0.0:
                    offsets = []
                    if g_plan_dir == "Center": offsets = [g_plan_off / 2.0, -g_plan_off / 2.0]
                    elif g_plan_dir == "Inside": offsets = [g_plan_off]
                    elif g_plan_dir == "Outside": offsets = [-g_plan_off]
                    elif g_plan_dir == "Both": offsets = [g_plan_off, -g_plan_off]
                    
                    is_endpoint = (param == 0.0 or param == 1.0)
                    k = (round(pt.X, 3), round(pt.Y, 3), round(pt.Z, 3))
                    connections = endpoints_map.get(k, []) if is_endpoint else []
                    
                    unique_connections = []
                    seen_idx = set()
                    for conn in connections:
                        if conn[0] not in seen_idx:
                            seen_idx.add(conn[0])
                            unique_connections.append(conn)
                            
                    if len(unique_connections) == 2:
                        idx1, end1 = unique_connections[0]
                        idx2, end2 = unique_connections[1]
                        
                        c1, _ = target_curves[idx1]
                        c2, _ = target_curves[idx2]
                        
                        t1_fwd = get_tangent(c1, 0.0 if end1 == 0 else 1.0)
                        t2_fwd = get_tangent(c2, 0.0 if end2 == 0 else 1.0)
                        
                        n1 = XYZ(0, 1, 0) if t1_fwd.IsAlmostEqualTo(XYZ.Zero) else XYZ.BasisZ.CrossProduct(t1_fwd).Normalize()
                        n2 = XYZ(0, 1, 0) if t2_fwd.IsAlmostEqualTo(XYZ.Zero) else XYZ.BasisZ.CrossProduct(t2_fwd).Normalize()
                        
                        for off_val in offsets:
                            p1 = base_pt + n1.Multiply(off_val)
                            p2 = base_pt + n2.Multiply(off_val)
                            I = intersect_2d(p1, t1_fwd, p2, t2_fwd)
                            
                            if I:
                                I = XYZ(I.X, I.Y, base_pt.Z)
                                if dist_2d(I, base_pt) > abs(off_val) * 10.0:
                                    res.append(p1)
                                    res.append(p2)
                                else:
                                    t_val = (I.X - p1.X) * t1_fwd.X + (I.Y - p1.Y) * t1_fwd.Y
                                    s_val = (I.X - p2.X) * t2_fwd.X + (I.Y - p2.Y) * t2_fwd.Y
                                    t_inside = -t_val if end1 == 1 else t_val
                                    s_inside = -s_val if end2 == 1 else s_val
                                    
                                    if t_inside > -1e-4 and s_inside > -1e-4:
                                        # INNER CORNER: 1 shared intersection point
                                        res.append(I)
                                    else:
                                        # OUTER CORNER: 2 parallel offset points
                                        res.append(p1)
                                        res.append(p2)
                            else:
                                # PARALLEL / NO INTERSECTION
                                res.append(p1)
                                res.append(p2)
                    else:
                        for off_val in offsets:
                            P_off = base_pt + normal.Multiply(off_val)
                            valid = True
                            
                            # Cull points that cross the inner miter line at the start
                            I_start = get_cached_intersection(curve_idx, 0, off_val, override_z)
                            if I_start:
                                t0_fwd = get_tangent(curve, 0.0)
                                if (P_off.X - I_start.X) * t0_fwd.X + (P_off.Y - I_start.Y) * t0_fwd.Y < -1e-4:
                                    valid = False
                                    
                            # Cull points that cross the inner miter line at the end
                            if valid:
                                I_end = get_cached_intersection(curve_idx, 1, off_val, override_z)
                                if I_end:
                                    t1_fwd = get_tangent(curve, 1.0)
                                    if (P_off.X - I_end.X) * (-t1_fwd.X) + (P_off.Y - I_end.Y) * (-t1_fwd.Y) < -1e-4:
                                        valid = False
                                        
                            if valid:
                                res.append(P_off)
                return res
            for curve_idx, (curve, override_z) in enumerate(target_curves):
                length = curve.Length
                step_len = g_int
                if step_len < 0.1: step_len = 1.0
                current_len = 0.0
                while current_len <= length + 0.001:
                    param = current_len / length
                    if param > 1.0: param = 1.0
                    pts_to_add.extend(add_curve_pts(curve_idx, param, override_z))
                    current_len += step_len
                if abs((current_len - step_len) - length) > 0.01:
                    pts_to_add.extend(add_curve_pts(curve_idx, 1.0, override_z))
            
            off = 0.0
            if toposolid.Id != raw_elem.Id:
                off = get_subdivision_offset(doc, raw_elem.Id)
            
            try: pt_tol = float(getattr(state, "point_dist_tol", "0.25"))
            except: pt_tol = 0.25

            # Spatial Hashing for fast pure Python deduplication
            cell_size = max(pt_tol, 0.25)
            pts_grid = {}
            unique_pts = []
            
            for pt in pts_to_add:
                cx = int(math.floor(pt.X / cell_size))
                cy = int(math.floor(pt.Y / cell_size))
                conflict_idx = None
                
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        for idx in pts_grid.get((cx + dx, cy + dy), []):
                            if dist_2d(unique_pts[idx], pt) < pt_tol:
                                conflict_idx = idx
                                break
                        if conflict_idx is not None: break
                        
                if conflict_idx is not None:
                    unique_pts[conflict_idx] = pt
                else:
                    pts_grid.setdefault((cx, cy), []).append(len(unique_pts))
                    unique_pts.append(pt)
            
            added = 0
            modified = 0
            
            # Spatial Hash for existing vertices to avoid O(N*M) lookups
            verts_grid = {}
            for v, pos in tracker.get_all():
                try:
                    cx = int(math.floor(pos.X / cell_size))
                    cy = int(math.floor(pos.Y / cell_size))
                    verts_grid.setdefault((cx, cy), []).append((v, pos))
                except: continue
            
            for pt in unique_pts:
                adjusted_pt = XYZ(pt.X, pt.Y, pt.Z - off)
                matched_vert = None
                matched_pos = None
                
                cx = int(math.floor(adjusted_pt.X / cell_size))
                cy = int(math.floor(adjusted_pt.Y / cell_size))
                
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        for v, pos in verts_grid.get((cx + dx, cy + dy), []):
                            if dist_2d(pos, adjusted_pt) < pt_tol:
                                matched_vert = v
                                matched_pos = pos
                                break
                        if matched_vert: break
                        
                if matched_vert:
                    if abs(matched_pos.Z - adjusted_pt.Z) > 0.005:
                        if tracker.delete(matched_vert, matched_pos):
                            if tracker.add(adjusted_pt): modified += 1
                else:
                    if tracker.add(adjusted_pt): added += 1
            
            if target_region_infos and not state.reset_mode:
                for v, pos in tracker.get_all():
                    reg_z = get_region_z(pos.X, pos.Y, target_region_infos)
                    if reg_z is not None:
                        target_z = reg_z - off + g_z_off
                        if abs(pos.Z - target_z) > 0.005:
                            if tracker.delete(v, pos):
                                if tracker.add(XYZ(pos.X, pos.Y, target_z)): modified += 1
                            
            # --- ADAPTIVE DUAL-RING TRIANGULATION HALO ---
            enable_halo = getattr(state, "enable_halo", True)
            eliminated_count = 0
            halo_added = 0
            
            if enable_halo:
                halo_offset_inner = float(getattr(state, "halo_inner", "0.25"))
                halo_offset_outer = g_int # Automatically bounds to the outer grid resolution
                halo_search_rad = max(2.0, g_int * 1.5)
                halo_pts = []
                
                # Tolerance buffer to reject wildly deviating raycasts
                try: raycast_tol = float(getattr(state, "outlier_tol", "1.0")) * 3.0
                except: raycast_tol = 3.0

                # Build a coarser spatial hash specifically for the halo search
                halo_cell_size = halo_search_rad
                halo_verts_grid = {}
                for v, pos in tracker.get_all():
                    try:
                        hcx = int(math.floor(pos.X / halo_cell_size))
                        hcy = int(math.floor(pos.Y / halo_cell_size))
                        halo_verts_grid.setdefault((hcx, hcy), []).append(pos)
                    except: pass
                
                # 1. Generate Dual-Ring Halo Points
                for curve, override_z in target_curves:
                    length = curve.Length
                    step_len = g_int
                    if step_len < 0.1: step_len = 1.0
                    current_len = 0.0
                    
                    while current_len <= length + 0.001:
                        param = current_len / length
                        if param > 1.0: param = 1.0
                        pt = curve.Evaluate(param, True)
                        
                        try:
                            deriv = curve.ComputeDerivatives(param, True)
                            tangent = deriv.BasisX.Normalize()
                            if tangent.IsAlmostEqualTo(XYZ.Zero): normal = XYZ(0, 1, 0)
                            else: normal = XYZ.BasisZ.CrossProduct(tangent).Normalize()
                        except:
                            normal = XYZ(0, 1, 0)
                            
                        h1_in = pt + normal * halo_offset_inner
                        h2_in = pt - normal * halo_offset_inner
                        h1_out = pt + normal * halo_offset_outer
                        h2_out = pt - normal * halo_offset_outer
                        
                        valid_halos = []
                        is_region = (override_z is not None and target_region_infos)
                        
                        if is_region:
                            if get_region_z(h1_in.X, h1_in.Y, target_region_infos) is None: valid_halos.append(h1_in)
                            if get_region_z(h2_in.X, h2_in.Y, target_region_infos) is None: valid_halos.append(h2_in)
                            if get_region_z(h1_out.X, h1_out.Y, target_region_infos) is None: valid_halos.append(h1_out)
                            if get_region_z(h2_out.X, h2_out.Y, target_region_infos) is None: valid_halos.append(h2_out)
                        else:
                            if g_plan_off == 0.0 or g_plan_dir == "Both":
                                valid_halos.extend([h1_in, h2_in, h1_out, h2_out])
                            elif g_plan_dir == "Center":
                                valid_halos.extend([h1_in, h2_in, h1_out, h2_out])
                            elif g_plan_dir == "Inside":
                                valid_halos.extend([h2_in, h2_out])
                            elif g_plan_dir == "Outside":
                                valid_halos.extend([h1_in, h1_out])
                                
                        for hp in valid_halos:
                            # 1. Calculate fast spatial average as a reliable baseline
                            hcx_halo = int(math.floor(hp.X / halo_cell_size))
                            hcy_halo = int(math.floor(hp.Y / halo_cell_size))
                            
                            nearby_z = []
                            for dx in (-1, 0, 1):
                                for dy in (-1, 0, 1):
                                    for pos in halo_verts_grid.get((hcx_halo + dx, hcy_halo + dy), []):
                                        if dist_2d(pos, hp) < halo_search_rad:
                                            if override_z is not None and is_in_graded_region(pos.X, pos.Y, target_region_infos, target_curves, g_plan_off, g_plan_dir):
                                                continue
                                            nearby_z.append(pos.Z)
                                            
                            avg_z = sum(nearby_z) / len(nearby_z) if nearby_z else (override_z if override_z is not None else pt.Z) + g_z_off - off
                            
                            hz = None
                            
                            # 2. Attempt Exact Surface Raycast
                            if intersector:
                                z_hit, hit_id = get_surface_info(intersector, hp, ray_start_z)
                                if z_hit is not None:
                                    ray_z = z_hit - get_subdivision_offset(doc, hit_id)
                                    # 3. Tolerance Check: Reject wildly deviating raycasts
                                    if abs(ray_z - avg_z) <= raycast_tol:
                                        hz = ray_z
                            
                            # 4. Fallback to spatial average
                            if hz is None:
                                hz = avg_z
                                
                            halo_pts.append(XYZ(hp.X, hp.Y, hz))
                        
                        current_len += step_len
                        
                # 2. Eliminate existing points strictly within the optimized boundary zone
                for v, pos in tracker.get_all():
                        is_region = False
                        if target_region_infos:
                            if is_in_graded_region(pos.X, pos.Y, target_region_infos, target_curves, g_plan_off, g_plan_dir):
                                is_region = True
                        if is_region: continue
                        
                        conflict_with_new = False
                        pcx = int(math.floor(pos.X / cell_size))
                        pcy = int(math.floor(pos.Y / cell_size))
                        for dx in (-1, 0, 1):
                            for dy in (-1, 0, 1):
                                for idx in pts_grid.get((pcx + dx, pcy + dy), []):
                                    if dist_2d(unique_pts[idx], pos) < MIN_DIST_TOLERANCE:
                                        conflict_with_new = True
                                        break
                                if conflict_with_new: break
                            if conflict_with_new: break
                            
                        if conflict_with_new: continue
                        
                        min_d = float('inf')
                        for curve, _ in target_curves:
                            _, _, d_xy = project_2d(curve, pos)
                            if d_xy < min_d: min_d = d_xy
                            
                        if min_d <= halo_offset_outer + 0.1:
                            if tracker.delete(v, pos): eliminated_count += 1

                # 3. Snapshot valid existing points post-deletion to avoid false conflicts
                valid_verts_grid = {}
                for v, pos in tracker.get_all():
                    try:
                        hcx = int(math.floor(pos.X / cell_size))
                        hcy = int(math.floor(pos.Y / cell_size))
                        valid_verts_grid.setdefault((hcx, hcy), []).append(pos)
                    except: pass
                
                # 4. Add Halo Points
                for hp in halo_pts:
                    cx = int(math.floor(hp.X / cell_size))
                    cy = int(math.floor(hp.Y / cell_size))
                    
                    conflict = False
                    # Check against the core grading points
                    for dx in (-1, 0, 1):
                        for dy in (-1, 0, 1):
                            for idx in pts_grid.get((cx + dx, cy + dy), []):
                                if dist_2d(unique_pts[idx], hp) < MIN_DIST_TOLERANCE:
                                    conflict = True
                                    break
                            if conflict: break
                            
                    # Check against the remaining valid ground points
                    if not conflict:
                        for dx in (-1, 0, 1):
                            for dy in (-1, 0, 1):
                                for pos in valid_verts_grid.get((cx + dx, cy + dy), []):
                                    if dist_2d(pos, hp) < MIN_DIST_TOLERANCE:
                                        conflict = True
                                        break
                                if conflict: break
                                
                    if not conflict:
                        if tracker.add(hp):
                            unique_pts.append(hp)
                            pts_grid.setdefault((cx, cy), []).append(len(unique_pts) - 1)
                            halo_added += 1
            
            status = t.Commit()
            if status == TransactionStatus.Committed:
                tg.Assimilate()
                log.info("Successfully graded elements: Added {} new points, Modified {} existing points.".format(added, modified))
                if enable_halo:
                    log.info("Triangulation Anchors: Cleared {} boundary points and added {} halo points for clean blending.\n".format(eliminated_count, halo_added))
            else:
                if tg.HasStarted(): tg.RollBack()
                log.error("Revit rejected the changes.", "The Toposolid geometry may have become invalid (e.g., too thin).")
        except Exception as loop_e:
            if 't' in locals() and t.HasStarted(): t.RollBack()
            if tg is not None and tg.HasStarted(): tg.RollBack()
            log.error("Failed to add points.", "{}\n{}".format(loop_e, traceback.format_exc()))
            
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
        except: 
            pass
        
        action = win.next_action
        
        if not action:
            save_state_to_disk(state)
            break 
            
        elif action == "select_stakes":
            try:
                ref_start = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Select Start Stake")
                state.start_stake = doc.GetElement(ref_start)
                
                # Visual clue: Highlight the start stake while waiting
                try: 
                    uidoc.Selection.SetElementIds(List[ElementId]([ref_start.ElementId]))
                except: 
                    pass
                
                ref_end = uidoc.Selection.PickObject(ObjectType.Element, UniversalFilter(), "Start Stake selected. Now select End Stake")
                state.end_stake = doc.GetElement(ref_end)
                
                save_state_to_disk(state)
            except: 
                pass # Cancelled selection
            
        elif action == "select_line":
            try: 
                refs = uidoc.Selection.PickObjects(ObjectType.Element, UniversalFilter(), "Select Guide Lines (Tab for chain, Finish when done)")
                state.grading_lines = [doc.GetElement(r) for r in refs]
                save_state_to_disk(state)
            except: 
                pass # Cancelled selection
            
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
            
        elif action == "smooth_path":
            perform_smooth_path(state)
            
        elif action == "line_points":
            perform_add_points_along_line(state)
            save_state_to_disk(state)
            
        elif action == "load_recipe": 
            perform_load_recipe(state)
            save_state_to_disk(state)
                
        # Clear the evaluation cache to free memory before reopening the UI
        _curve_eval_cache.clear()
        _subdivision_offset_cache.clear()