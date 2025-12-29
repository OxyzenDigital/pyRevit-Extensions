# -*- coding: utf-8 -*-
"""
Join Pipes Extension
"""
__context__ = "active-view-type: FloorPlan,CeilingPlan,EngineeringPlan,AreaPlan,Section,Elevation,ThreeD"

import sys
import os
import clr
import traceback # Added for error details

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

from Autodesk.Revit.UI import UIThemeManager, UITheme
from Autodesk.Revit.Exceptions import OperationCanceledException # Import specific exception
from System.Windows.Media import SolidColorBrush, Color as WpfColor

from pyrevit import forms, revit, script

# Custom Modules
import data_model
import logic
import revit_service

doc = revit.doc
uidoc = revit.uidoc

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
class JoinPipesWindow(forms.WPFWindow):
    def __init__(self, state):
        forms.WPFWindow.__init__(self, 'ui.xaml')
        self.state = state
        
        # Restore Position
        try:
            if self.state.win_top > 0: self.Top = self.state.win_top
            if self.state.win_left > 0: self.Left = self.state.win_left
        except: pass
        
        self.bind_ui()
        self.setup_events()
        self.apply_revit_theme()

    def bind_ui(self):
        # 1. Labels
        self.Lb_Source.Text = self.state.source_desc
        self.Lb_Target.Text = self.state.target_desc
        
        # 2. Settings
        self.Cb_Rolling.IsChecked = self.state.allow_rolling
        self.Cb_Vertical.IsChecked = self.state.allow_vertical
        
        # 3. Solution Info
        if self.state.solutions and self.state.selected_solution_index >= 0:
            sol = self.state.current_solution
            self.Lb_SolName.Text = "Solution: " + sol.name
            self.Lb_SolDesc.Text = sol.description
            self.Lb_SolCount.Text = "{} / {}".format(self.state.selected_solution_index + 1, len(self.state.solutions))
            
            self.Btn_Prev.IsEnabled = (self.state.selected_solution_index > 0)
            self.Btn_Next.IsEnabled = (self.state.selected_solution_index < len(self.state.solutions) - 1)
            self.Btn_Commit.IsEnabled = True
        else:
            self.Lb_SolName.Text = "No Solutions"
            self.Lb_SolDesc.Text = "Select pipes to calculate."
            self.Lb_SolCount.Text = "- / -"
            self.Btn_Prev.IsEnabled = False
            self.Btn_Next.IsEnabled = False
            self.Btn_Commit.IsEnabled = False

        self.Lb_Status.Content = self.state.status_message

    def apply_revit_theme(self):
        # Mimic Grading Tool Theme
        try:
            is_dark = (UIThemeManager.CurrentTheme == UITheme.Dark)
        except: is_dark = False
        
        if is_dark:
            res = self.Resources
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(59, 68, 83))
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(40, 46, 56))
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(245, 245, 245))
            res["TextLightBrush"] = SolidColorBrush(WpfColor.FromRgb(170, 175, 185))
            res["AccentBrush"] = SolidColorBrush(WpfColor.FromRgb(0, 120, 215))
            res["HeaderTextBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(85, 95, 110))
            res["ButtonBrush"] = SolidColorBrush(WpfColor.FromRgb(70, 80, 95))

    def setup_events(self):
        # Header / Window
        self.Btn_WinClose.Click += self.close_window
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        
        # Main Actions
        self.Btn_Select.Click += self.act_select
        self.Btn_Swap.Click += self.act_swap
        self.Btn_Commit.Click += self.act_commit
        
        # Navigation
        self.Btn_Next.Click += self.act_next
        self.Btn_Prev.Click += self.act_prev
        
        # Settings
        self.Cb_Rolling.Checked += self.settings_changed
        self.Cb_Rolling.Unchecked += self.settings_changed
        self.Cb_Vertical.Checked += self.settings_changed
        self.Cb_Vertical.Unchecked += self.settings_changed

    # Event Handlers
    def close_window(self, s, a): self.Close()
    def drag_window(self, s, a): self.DragMove()
    
    def update_state(self):
        self.state.allow_rolling = self.Cb_Rolling.IsChecked
        self.state.allow_vertical = self.Cb_Vertical.IsChecked

    def act_select(self, s, a):
        self.update_state()
        self.state.next_action = "select"
        self.Close()

    def act_swap(self, s, a):
        self.update_state()
        self.state.next_action = "swap"
        self.Close()
        
    def act_commit(self, s, a):
        self.update_state()
        self.state.next_action = "commit"
        self.Close()
        
    def act_next(self, s, a):
        self.update_state()
        self.state.next_action = "next"
        self.Close()
        
    def act_prev(self, s, a):
        self.update_state()
        self.state.next_action = "prev"
        self.Close()

    def settings_changed(self, s, a):
        self.update_state()
        pass

# --- MAIN LOOP ---

def main():
    # 1. Initialize State
    app_state = data_model.AppState()
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
                                
                                solutions = solver.calculate_solutions(data1, data2)
                                app_state.solutions = solutions
                                app_state.selected_solution_index = 0 if solutions else -1
                                app_state.status_message = "Found {} solutions.".format(len(solutions))
                                
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
                        # Re-calc
                        data1 = srv.get_element_data(app_state.source_id)
                        data2 = srv.get_element_data(app_state.target_id)
                        solutions = solver.calculate_solutions(data1, data2)
                        app_state.solutions = solutions
                        app_state.selected_solution_index = 0 if solutions else -1
                        app_state.status_message = "Swapped. Found {} solutions.".format(len(solutions))
                        
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