# -*- coding: utf-8 -*-
"""
Join Pipes Extension
"""
__context__ = "active-view-type: FloorPlan,CeilingPlan,EngineeringPlan,AreaPlan,Section,Elevation,ThreeD"

import sys
import os
import clr
import json # Added
import traceback # Added for error details

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

from Autodesk.Revit.UI import UIThemeManager, UITheme
from Autodesk.Revit.Exceptions import OperationCanceledException # Import specific exception
from System.Windows.Media import SolidColorBrush, Color as WpfColor

try:
    from Autodesk.Revit.UI import UIThemeManager, UITheme
    HAS_THEME = True
except ImportError:
    HAS_THEME = False

from pyrevit import forms, revit, script

# Custom Modules
import data_model
import logic
import revit_service

doc = revit.doc
uidoc = revit.uidoc

SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "settings.json")

def get_id_val(obj):
    """Safe retrieval of ElementId integer value for Revit 2024+ compatibility."""
    if hasattr(obj, "Value"): return obj.Value # Revit 2024+
    if hasattr(obj, "IntegerValue"): return obj.IntegerValue # Revit <2024
    # If obj is actually an Element, get its Id first
    if hasattr(obj, "Id"):
        eid = obj.Id
        if hasattr(eid, "Value"): return eid.Value
        if hasattr(eid, "IntegerValue"): return eid.IntegerValue
    return -1

# --- LOGGER ---
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

    def show(self, title="Join Pipes Report"):
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

# --- UI CLASS ---
class SettingsWindow(forms.WPFWindow):
    def __init__(self, state):
        forms.WPFWindow.__init__(self, 'settings.xaml')
        self.state = state
        self.apply_revit_theme()
        self.bind_ui()
        self.Btn_Save.Click += self.save_settings
        self.Btn_Cancel.Click += self.close_window

    def apply_revit_theme(self):
        if HAS_THEME and UIThemeManager.CurrentTheme == UITheme.Dark:
            res = self.Resources
            # Dark Theme Palette (Matching Main Window)
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(30, 30, 30))
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(45, 48, 55))
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
            res["TextSubBrush"] = SolidColorBrush(WpfColor.FromRgb(180, 180, 180))
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(70, 70, 70))
            # Button Text in Dark Mode is White on Dark Background
            # Button Background is ControlBrush or dedicated ButtonBrush?
            # In settings.xaml, Button uses ControlBrush.
            
            # Accent stays #0080FF
            pass

    def bind_ui(self):
        self.Cb_Rolling.IsChecked = self.state.allow_rolling
        self.Cb_Vertical.IsChecked = self.state.allow_vertical

    def save_settings(self, sender, args):
        self.state.allow_rolling = self.Cb_Rolling.IsChecked
        self.state.allow_vertical = self.Cb_Vertical.IsChecked
        
        # Save to JSON
        try:
            data = {
                "allow_rolling": self.state.allow_rolling,
                "allow_vertical": self.state.allow_vertical
            }
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(data, f)
        except Exception as e:
            print("Failed to save settings: " + str(e))
            
        self.Close()

    def close_window(self, sender, args):
        self.Close()

class JoinPipesWindow(forms.WPFWindow):
    def __init__(self, state):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.state = state
        
        # Restore Position
        try:
            if self.state.win_top > 0: self.Top = self.state.win_top
            if self.state.win_left > 0: self.Left = self.state.win_left
        except: pass
        
        self.apply_revit_theme()
        self.bind_ui()
        self.setup_events()

    def apply_revit_theme(self):
        if HAS_THEME and UIThemeManager.CurrentTheme == UITheme.Dark:
            res = self.Resources
            # Dark Theme Palette
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(30, 30, 30))
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(45, 48, 55)) # Replaces White for Solution Card
            res["ButtonBrush"] = SolidColorBrush(WpfColor.FromRgb(50, 50, 50))
            
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
            res["TextLightBrush"] = SolidColorBrush(WpfColor.FromRgb(180, 180, 180))
            
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(70, 70, 70))
            
            # Note: CardBrush (#282a2f) is already dark, works well on #1E1E1E window.
            # CardTextBrush is already White.
            # We just need to make sure the "ControlBrush" used for the Solution card flips to dark.
            
            # Accent can stay #0080FF or get slightly brighter if needed.

    def bind_ui(self):
        # 1. Labels
        # Format Source Description
        if hasattr(self.state, 'source_data') and self.state.source_data:
            d = self.state.source_data
            desc = "{}\nSize: {}\nSys: {}\nSlope: {}".format(
                d.get('type_name', 'Unknown'),
                d.get('size_str', '-'),
                d.get('system_name', '-'),
                d.get('slope_str', '-')
            )
            self.Lb_Source.Text = desc
        else:
            self.Lb_Source.Text = self.state.source_desc

        # Format Target Description
        if hasattr(self.state, 'target_data') and self.state.target_data:
            d = self.state.target_data
            desc = "{}\nSize: {}\nSys: {}\nSlope: {}".format(
                d.get('type_name', 'Unknown'),
                d.get('size_str', '-'),
                d.get('system_name', '-'),
                d.get('slope_str', '-')
            )
            self.Lb_Target.Text = desc
        else:
            self.Lb_Target.Text = self.state.target_desc
        
        # 3. Solution Info
        if self.state.solutions and self.state.selected_solution_index >= 0:
            sol = self.state.current_solution
            self.Lb_SolName.Text = sol.name.upper()
            self.Lb_SolDesc.Text = sol.description
            self.Lb_SolCount.Text = "{} / {}".format(self.state.selected_solution_index + 1, len(self.state.solutions))
            
            self.Btn_Prev.IsEnabled = (self.state.selected_solution_index > 0)
            self.Btn_Next.IsEnabled = (self.state.selected_solution_index < len(self.state.solutions) - 1)
            self.Btn_Commit.IsEnabled = sol.is_valid
        else:
            self.Lb_SolName.Text = "NO SOLUTION"
            self.Lb_SolDesc.Text = "Select elements to calculate."
            self.Lb_SolCount.Text = "- / -"
            self.Btn_Prev.IsEnabled = False
            self.Btn_Next.IsEnabled = False
            self.Btn_Commit.IsEnabled = False

        self.Lb_Status.Content = self.state.status_message

    def setup_events(self):
        # Header / Window
        self.Btn_WinClose.Click += self.close_window
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_Settings.Click += self.act_settings
        
        # Main Actions
        self.Btn_Select.Click += self.act_select
        self.Btn_Swap.Click += self.act_swap
        self.Btn_Commit.Click += self.act_commit
        
        # Navigation
        self.Btn_Next.Click += self.act_next
        self.Btn_Prev.Click += self.act_prev

    # Event Handlers
    def close_window(self, s, a): self.Close()
    def drag_window(self, s, a): self.DragMove()
    
    def act_settings(self, s, a):
        # Open Settings Modal
        sw = SettingsWindow(self.state)
        sw.Owner = self
        sw.ShowDialog()
        # Refresh UI if needed (though settings mostly affect calculation which happens on select/swap)

    def act_select(self, s, a):
        self.state.next_action = "select"
        self.Close()

    def act_swap(self, s, a):
        self.state.next_action = "swap"
        self.Close()
        
    def act_commit(self, s, a):
        self.state.next_action = "commit"
        self.Close()
        
    def act_next(self, s, a):
        self.state.next_action = "next"
        self.Close()
        
    def act_prev(self, s, a):
        self.state.next_action = "prev"
        self.Close()

# --- MAIN LOOP ---

def main():
    # 1. Initialize State
    app_state = data_model.AppState()
    
    # Load Settings
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                data = json.load(f)
                app_state.allow_rolling = data.get("allow_rolling", True)
                app_state.allow_vertical = data.get("allow_vertical", True)
        except: pass

    srv = revit_service.RevitService(doc, uidoc)
    solver = logic.Solver(app_state)
    log = BatchLogger() # Instantiate Logger
    
    # 2. Loop
    try:
        while True:
            # Show Window
            win = JoinPipesWindow(app_state)
            win.ShowDialog()
            
            # Capture Window Position
            try:
                app_state.win_top = win.Top
                app_state.win_left = win.Left
            except: pass
            
            # Get Action
            action = app_state.next_action
            app_state.next_action = None # Reset
            
            if not action:
                srv.clear_preview(app_state) # Cleanup on close
                srv.highlight_elements([]) # Clear selection
                break # User closed window
                
            # Execute Action
            try:
                if action == "select":
                    srv.clear_preview(app_state) # Clear old preview
                    try:
                        el1 = srv.pick_element("Select Source Pipe")
                        if el1:
                            app_state.source_id = el1.Id
                            app_state.source_desc = "{} (ID: {})".format(el1.Name, get_id_val(el1.Id))
                            
                            el2 = srv.pick_element("Select Target Pipe")
                            if el2:
                                app_state.target_id = el2.Id
                                app_state.target_desc = "{} (ID: {})".format(el2.Name, get_id_val(el2.Id))
                                
                                # Run Calculation
                                data1 = srv.get_element_data(app_state.source_id)
                                data2 = srv.get_element_data(app_state.target_id)
                                
                                # Store Data for UI
                                app_state.source_data = data1
                                app_state.target_data = data2
                                
                                solutions = solver.calculate_solutions(data1, data2)
                                app_state.solutions = solutions
                                app_state.selected_solution_index = 0 if solutions else -1
                                app_state.status_message = "Found {} solutions.".format(len(solutions))
                                
                                # Highlight selected elements
                                srv.highlight_elements([app_state.source_id, app_state.target_id])

                                # Optional: Auto-preview first solution
                                if solutions:
                                    srv.visualize_solution(solutions[0], app_state)
                            else:
                                app_state.status_message = "Target selection cancelled."
                        else:
                            app_state.status_message = "Source selection cancelled."
                    
                    except OperationCanceledException:
                        app_state.status_message = "Selection cancelled by user."
                    except Exception as e:
                        log.error("Selection Error", traceback.format_exc())
                        app_state.status_message = "Error during selection."

                elif action == "swap":
                    if app_state.source_id and app_state.target_id:
                        # Swap Ids
                        app_state.source_id, app_state.target_id = app_state.target_id, app_state.source_id
                        app_state.source_desc, app_state.target_desc = app_state.target_desc, app_state.source_desc
                        
                        # Swap Data if exists
                        if hasattr(app_state, 'source_data') and hasattr(app_state, 'target_data'):
                             app_state.source_data, app_state.target_data = app_state.target_data, app_state.source_data
                        
                        # Re-calc
                        data1 = srv.get_element_data(app_state.source_id)
                        data2 = srv.get_element_data(app_state.target_id)
                        
                        # Ensure state data is fresh/correct (optional but safe)
                        app_state.source_data = data1
                        app_state.target_data = data2
                        
                        solutions = solver.calculate_solutions(data1, data2)
                        app_state.solutions = solutions
                        app_state.selected_solution_index = 0 if solutions else -1
                        app_state.status_message = "Swapped. Found {} solutions.".format(len(solutions))
                        
                        # Highlight selected elements
                        srv.highlight_elements([app_state.source_id, app_state.target_id])
                        
                        if solutions:
                             srv.visualize_solution(solutions[0], app_state)
                    else:
                        app_state.status_message = "Nothing to swap."

                elif action == "next":
                    if app_state.selected_solution_index < len(app_state.solutions) - 1:
                        app_state.selected_solution_index += 1
                        srv.visualize_solution(app_state.current_solution, app_state)

                elif action == "prev":
                    if app_state.selected_solution_index > 0:
                        app_state.selected_solution_index -= 1
                        srv.visualize_solution(app_state.current_solution, app_state)

                elif action == "commit":
                    srv.clear_preview(app_state) # Clear before real commit
                    srv.highlight_elements([]) # Clear selection
                    if app_state.current_solution:
                        success = srv.commit_solution(app_state.current_solution)
                        if success:
                            app_state.status_message = "Join completed successfully."
                        else:
                            app_state.status_message = "Error: Failed to join pipes."

            except Exception as e:
                log.error("Unexpected Error", traceback.format_exc())
                app_state.status_message = "Unexpected error occurred."

    except Exception as e:
        log.error("Critical Loop Error", traceback.format_exc())
        
    # Show log at the end if errors occurred
    log.show()

if __name__ == '__main__':
    main()