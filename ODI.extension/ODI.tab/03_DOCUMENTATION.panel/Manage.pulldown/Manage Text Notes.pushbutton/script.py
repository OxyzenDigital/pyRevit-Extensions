# -*- coding: utf-8 -*-
__title__ = "Manage Text Notes"
__author__ = "Oxyzen Digital Inc"

import os
from pyrevit import forms, revit, HOST_APP
from System.Collections.Generic import List
from System.Windows.Media import SolidColorBrush, Colors
from System.Windows.Media import Color as WpfColor
from view_model import ManageTextNotesViewModel

doc = revit.doc
uidoc = revit.uidoc

class ManageTextNotesWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        self.vm = ManageTextNotesViewModel()
        self.DataContext = self.vm
        self.apply_revit_theme()

        # Bind header/window actions
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_WinClose.Click += self.close_window

        # Main Toolbar
        self.Btn_ScanActive.Click += self.scan_active
        self.Btn_ScanProject.Click += self.scan_project

        # Tree Footer
        self.Btn_ExpandAll.Click += self.expand_all
        self.Btn_CollapseAll.Click += self.collapse_all
        self.Btn_SelectAll.Click += self.select_all
        self.Btn_SelectNone.Click += self.select_none
        # Notes Grid Toolbar
        self.Btn_ExportNew.Click += self.export_new
        self.Btn_AppendProject.Click += self.append_project

        # Master DB Editor Toolbar
        self.Btn_ClearCanvas.Click += self.clear_canvas
        self.Btn_PickFile.Click += self.pick_file
        self.Btn_SaveFile.Click += self.save_file
        self.Btn_SaveAsFile.Click += self.save_as_file
        self.Btn_AddNote.Click += self.add_note
        
        self.Btn_MasterExpandAll.Click += self.master_expand_all
        self.Btn_MasterCollapseAll.Click += self.master_collapse_all

        # Tree Events
        self.systemTree.SelectedItemChanged += self.tree_selection_changed
        self.systemTree.MouseDoubleClick += self.tree_double_click
        self.Tree_Master.MouseDoubleClick += self.tree_master_double_click

        # Grid Events
        self.Grid_Notes.MouseDoubleClick += self.grid_double_click
        self.Grid_Notes.SelectionChanged += self.grid_selection_changed

        self.setup_grid_context_menu()
        self.setup_tree_context_menu()
        self.setup_master_tree_context_menu()
        
        self.vm.ActionCallback = self.update_action_buttons
        self.update_action_buttons()

    def apply_revit_theme(self):
        """Detects Revit theme and updates window resources if Dark."""
        is_dark = False
        try:
            if int(HOST_APP.version) >= 2024:
                from Autodesk.Revit.UI import UIThemeManager, UITheme
                if UIThemeManager.CurrentTheme == UITheme.Dark:
                    self._is_dark = is_dark = True
        except: pass
        
        if is_dark:
            res = self.Resources
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(45, 52, 64))
            res["ToolbarBrush"] = SolidColorBrush(WpfColor.FromRgb(31, 41, 55))
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(35, 41, 51))
            res["CardBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81))
            res["CardBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 100))
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
            res["TextLightBrush"] = SolidColorBrush(WpfColor.FromRgb(160, 165, 175))
            res["HeaderTextBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 100))
            res["ButtonBrush"] = SolidColorBrush(WpfColor.FromRgb(60, 68, 82))
            res["HoverBrush"] = SolidColorBrush(WpfColor.FromRgb(70, 85, 105))
            res["SelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(30, 58, 138))
            res["SelectionBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(59, 130, 246))
            res["SelectionTextBrush"] = SolidColorBrush(Colors.White)
            res["CheckedBrush"] = SolidColorBrush(WpfColor.FromRgb(30, 58, 138))
            res["CheckedTextBrush"] = SolidColorBrush(Colors.White)
            res["TracedBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 100))
            res["TracedBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(156, 163, 175))
            res["TracedTextBrush"] = SolidColorBrush(Colors.White)

    # --- UI Events ---

    def drag_window(self, sender, args):
        try: self.DragMove()
        except: pass

    def close_window(self, sender, args):
        self.Close()

    def export_new(self, sender, args):
        checked_notes = [n for n in self.vm.CurrentNotes if n.IsChecked]
        selected_notes = list(dict.fromkeys(checked_notes + list(self.Grid_Notes.SelectedItems)))
        if not selected_notes:
            selected_notes = self.vm.get_notes_from_checked_sheets()
        if not selected_notes:
            forms.alert("No project notes selected in the center pane, and no sheets are checked.")
            return
        self.vm.export_selected_to_new_file(selected_notes)
        self.clear_grid_selection()

    def BtnAppendSingleNote_Click(self, sender, args):
        try:
            note_item = sender.DataContext
            if note_item:
                self.vm.append_project_notes_to_master([note_item])
                self.clear_grid_selection()
        except Exception: pass

    def clear_canvas(self, sender, args):
        try:
            if len(self.vm.MasterNotes) > 0:
                res = forms.alert("This will clear all notes from the canvas.\n\nContinue?", yes=True, no=True)
                if not res: return
            self.vm.MasterNotes = []
            self.vm.filter_master_notes()
            self.vm.StatusText = "Canvas cleared."
            self.update_action_buttons()
        except Exception: pass

    def pick_file(self, sender, args):
        self.vm.pick_file()

    def save_file(self, sender, args):
        self.vm.save_master_file()

    def save_as_file(self, sender, args):
        self.vm.save_master_file(save_as=True)

    def add_note(self, sender, args):
        self.vm.add_new_master_note()

    def master_expand_all(self, sender, args):
        self.vm.expand_all_master(True)

    def master_collapse_all(self, sender, args):
        self.vm.expand_all_master(False)

    def Tree_Master_SelectedItemChanged(self, sender, args):
        self.vm.SelectedMasterNote = args.NewValue
        self.update_action_buttons()
    def BtnReplaceNotes_Click(self, sender, args):
        try:
            master_note = sender.DataContext
            if not master_note: return
            self.vm.SelectedMasterNote = master_note
            
            checked_notes = [n for n in self.vm.CurrentNotes if n.IsChecked]
            selected_notes = list(dict.fromkeys(checked_notes + list(self.Grid_Notes.SelectedItems)))
            if not selected_notes: return
            
            if len(selected_notes) > 1:
                res = forms.alert("You are about to replace {} notes with this Master Note.\n\nContinue?".format(len(selected_notes)), yes=True, no=True)
                if not res: return
                
            self.vm.replace_selected_project_notes(selected_notes)
            self.clear_grid_selection()
        except Exception as e:
            import traceback
            forms.alert(traceback.format_exc())

    def BtnEditNote_Click(self, sender, args):
        try:
            master_note = sender.DataContext
            if not master_note: return
            self.begin_editing_master_note(master_note)
        except Exception: pass

    def BtnAddNoteBelow_Click(self, sender, args):
        try:
            master_note = sender.DataContext
            if not master_note: return
            self.vm.add_note_below(master_note)
        except: pass

    def BtnDeleteNote_Click(self, sender, args):
        try:
            master_note = sender.DataContext
            if not master_note: return
            self.vm.delete_specific_master_note(master_note)
        except: pass

    def BtnDoneEdit_Click(self, sender, args):
        try:
            note = sender.DataContext
            note.IsEditing = False
        except Exception: pass

    def BtnCaseUpper_Click(self, sender, args):
        try:
            note = sender.DataContext
            if note and note.Text: note.Text = note.Text.upper()
        except Exception: pass

    def BtnCaseLower_Click(self, sender, args):
        try:
            note = sender.DataContext
            if note and note.Text: note.Text = note.Text.lower()
        except Exception: pass

    def BtnCaseTitle_Click(self, sender, args):
        try:
            note = sender.DataContext
            if note and note.Text: note.Text = note.Text.title()
        except Exception: pass

    def scan_project(self, sender, args):
        self.vm.scan_project()
        
    def scan_active(self, sender, args):
        self.vm.scan_active()
        self.update_action_buttons()

    def expand_all(self, sender, args):
        self.vm.expand_all_sheets(True)

    def collapse_all(self, sender, args):
        self.vm.expand_all_sheets(False)

    def select_all(self, sender, args):
        self.vm.select_all_sheets(True)

    # --- UI Events ---

    def begin_editing_master_note(self, master_note):
        try:
            self.vm.SelectedMasterNote = master_note
            master_note.IsEditing = True
            
            self.clear_grid_selection()
            
            match_text = master_note.Text
            found_match = False
            for n in self.vm.CurrentNotes:
                if n.Text == match_text:
                    n.IsChecked = True
                    found_match = True
            if found_match:
                self.Grid_Notes.Items.Refresh()
                self.update_action_buttons()
        except Exception: pass

    def clear_grid_selection(self):
        try:
            for n in self.vm.CurrentNotes:
                n.IsChecked = False
            self.Grid_Notes.SelectedItems.Clear()
            self.Grid_Notes.Items.Refresh()
            self.update_action_buttons()
        except Exception: pass

    def update_action_buttons(self):
        checked_notes = [n for n in self.vm.CurrentNotes if n.IsChecked]
        has_grid_target = len(self.Grid_Notes.SelectedItems) > 0 or len(checked_notes) > 0
        
        has_master_note = self.vm.SelectedMasterNote is not None
        
        has_checked_sheets = False
        def check_sheets(nodes):
            for n in nodes:
                if getattr(n, 'NodeType', '') == "Sheet" and n.IsChecked:
                    return True
                if hasattr(n, 'Children') and n.Children:
                    if check_sheets(n.Children): return True
            return False
        if self.vm.Sheets:
            has_checked_sheets = check_sheets(self.vm.Sheets)
            
        self.Btn_ExportNew.IsEnabled = has_grid_target or has_checked_sheets
        self.Btn_AppendProject.IsEnabled = has_grid_target or has_checked_sheets
        # We also need to update the VM's property for RelativeSource bindings
        self.vm.CanReplaceSelected = has_grid_target
        self.vm.CanReplaceAll = has_checked_sheets
        
        self.Btn_SaveFile.IsEnabled = len(self.vm.MasterNotes) > 0
        self.Btn_SaveAsFile.IsEnabled = len(self.vm.MasterNotes) > 0

    def tree_selection_changed(self, sender, args):
        self.vm.update_selected_sheet(args.NewValue)
        self.update_action_buttons()

    def grid_selection_changed(self, sender, args):
        self.update_action_buttons()

    def select_none(self, sender, args):
        self.vm.select_all_sheets(False)

    # --- Actions ---

    def append_project(self, sender, args):
        checked_notes = [n for n in self.vm.CurrentNotes if n.IsChecked]
        selected_notes = list(dict.fromkeys(checked_notes + list(self.Grid_Notes.SelectedItems)))
        if not selected_notes:
            selected_notes = self.vm.get_notes_from_checked_sheets()
        if not selected_notes: return
        self.vm.append_project_notes_to_master(selected_notes)
        self.clear_grid_selection()

    def replace_selected(self, sender, args):
        try:
            master_note = self.vm.SelectedMasterNote
            if not master_note:
                forms.alert("Please select a Master Note from the Right Pane first.")
                return
                
            checked_notes = [n for n in self.vm.CurrentNotes if n.IsChecked]
            selected_notes = list(dict.fromkeys(checked_notes + list(self.Grid_Notes.SelectedItems)))
            if not selected_notes: return
            
            if len(selected_notes) > 1:
                res = forms.alert("You are about to replace {} notes with this Master Note.\n\nContinue?".format(len(selected_notes)), yes=True, no=True)
                if not res: return
                
            self.vm.replace_selected_project_notes(selected_notes)
            self.clear_grid_selection()
        except Exception as e:
            import traceback
            forms.alert(traceback.format_exc())

    # --- Navigation ---

    def tree_selection_changed(self, sender, args):
        try:
            selected_node = self.systemTree.SelectedItem
            self.vm.update_selected_sheet(selected_node)
        except Exception: pass

    def tree_double_click(self, sender, args):
        try:
            selected_node = self.systemTree.SelectedItem
            if selected_node and hasattr(selected_node, "Sheet"):
                if doc.ActiveView.Id != selected_node.Sheet.Id:
                    uidoc.ActiveView = selected_node.Sheet
                elem_ids = List[revit.DB.ElementId]([selected_node.Sheet.Id])
                uidoc.ShowElements(elem_ids)
                uidoc.Selection.SetElementIds(elem_ids)
        except Exception: pass

    def grid_selection_changed(self, sender, args):
        try:
            if getattr(self.Cb_AutoZoom, "IsChecked", False):
                selected_note = self.Grid_Notes.SelectedItem
                if selected_note:
                    self.vm.zoom_to_note(selected_note)
        except Exception: pass

    def grid_double_click(self, sender, args):
        try:
            selected_note = self.Grid_Notes.SelectedItem
            if selected_note:
                self.ctx_select_parent(sender, args)
                self.vm.zoom_to_note(selected_note)
        except Exception: pass

    def tree_master_double_click(self, sender, args):
        try:
            selected_node = self.Tree_Master.SelectedItem
            if selected_node and selected_node.NodeType == "MasterNote":
                self.begin_editing_master_note(selected_node)
        except Exception: pass

    # --- Context Menus ---

    def setup_tree_context_menu(self):
        from System.Windows.Controls import ContextMenu, MenuItem
        ctx_menu = ContextMenu()
        
        item_zoom = MenuItem()
        item_zoom.Header = "Zoom to Sheet"
        item_zoom.Click += self.ctx_tree_zoom
        ctx_menu.Items.Add(item_zoom)
        
        item_isolate = MenuItem()
        item_isolate.Header = "Isolate Sheet (Uncheck Others)"
        item_isolate.Click += self.ctx_tree_isolate
        ctx_menu.Items.Add(item_isolate)
        
        self.systemTree.ContextMenu = ctx_menu

    def ctx_tree_zoom(self, sender, args):
        self.tree_double_click(sender, args)

    def ctx_tree_isolate(self, sender, args):
        try:
            selected_node = self.systemTree.SelectedItem
            if not selected_node: return
            
            for node in self.vm.Sheets:
                node.IsChecked = False
                
            selected_node.IsChecked = True
        except Exception: pass

    def setup_master_tree_context_menu(self):
        from System.Windows.Controls import ContextMenu, MenuItem
        ctx_menu = ContextMenu()
        
        item_find = MenuItem()
        item_find.Header = "Locate Identical Notes in Views Pane"
        item_find.Click += self.ctx_master_locate_in_views
        ctx_menu.Items.Add(item_find)
        
        self.Tree_Master.ContextMenu = ctx_menu

    def ctx_master_locate_in_views(self, sender, args):
        try:
            selected_node = self.Tree_Master.SelectedItem
            if not selected_node or selected_node.NodeType != "MasterNote": return
            
            target_text = selected_node.Text
            
            self.clear_grid_selection()
            
            matching_count = 0
            for n in self.vm.CurrentNotes:
                if n.Text == target_text:
                    n.IsChecked = True
                    matching_count += 1
            
            self.Grid_Notes.Items.Refresh()
            self.update_action_buttons()
            if matching_count > 0:
                self.vm.StatusText = "Selected {} identical notes in the current view.".format(matching_count)
            else:
                self.vm.StatusText = "No identical notes found in the currently displayed view."
        except Exception: pass

    def setup_grid_context_menu(self):
        from System.Windows.Controls import ContextMenu, MenuItem
        ctx_menu = ContextMenu()
        
        item_locate_canvas = MenuItem()
        item_locate_canvas.Header = "Locate in Canvas"
        item_locate_canvas.Click += self.ctx_grid_locate_in_canvas
        ctx_menu.Items.Add(item_locate_canvas)
        
        item_find = MenuItem()
        item_find.Header = "Isolate Identical Notes in Project"
        item_find.Click += self.ctx_select_identical
        ctx_menu.Items.Add(item_find)
        
        item_unisolate = MenuItem()
        item_unisolate.Header = "Unisolate (Restore View)"
        item_unisolate.Click += self.ctx_unisolate
        ctx_menu.Items.Add(item_unisolate)
        
        item_apply = MenuItem()
        item_apply.Header = "Replace with Selected Note from Master"
        item_apply.Click += self.replace_selected
        ctx_menu.Items.Add(item_apply)
        
        item_parent = MenuItem()
        item_parent.Header = "Locate Parent Sheet in Tree"
        item_parent.Click += self.ctx_select_parent
        ctx_menu.Items.Add(item_parent)
        
        item_highlight = MenuItem()
        item_highlight.Header = "Isolate Parent Sheet Notes"
        item_highlight.Click += self.ctx_highlight_children
        ctx_menu.Items.Add(item_highlight)
        
        self.Grid_Notes.ContextMenu = ctx_menu

    def ctx_grid_locate_in_canvas(self, sender, args):
        try:
            selected_items = list(self.Grid_Notes.SelectedItems)
            if not selected_items: return
            target_text = selected_items[0].Text
            
            match = next((m for m in self.vm.MasterNotes if m.Text == target_text), None)
            if match:
                self.vm.expand_all_master(True)
                for m in self.vm.MasterNotes: m.IsSelected = False
                match.IsSelected = True
                self.vm.SelectedMasterNote = match
                self.vm.StatusText = "Located note in Canvas."
            else:
                res = forms.alert("Note not found in Canvas. Add it now?", yes=True, no=True)
                if res:
                    self.vm.append_project_notes_to_master([selected_items[0]])
        except Exception: pass

    def ctx_select_identical(self, sender, args):
        try:
            selected_items = list(self.Grid_Notes.SelectedItems)
            if not selected_items: return
            target_text = selected_items[0].Text
            
            all_notes = self.vm.get_all_notes_in_nodes(self.vm.Sheets)
            matching_notes = [n for n in all_notes if n.Text == target_text]
            
            # Uncheck all sheets first to clear UI clutter
            for sheet in self.vm.Sheets:
                sheet.IsChecked = False
                
            # Populate CurrentNotes with ONLY these matching notes
            self.vm.CurrentNotes = matching_notes
            self.Grid_Notes.Items.Refresh()
            self.update_action_buttons()
            self.vm.StatusText = "Isolated {} identical notes across the project.".format(len(matching_notes))
        except Exception: pass

    def ctx_unisolate(self, sender, args):
        try:
            selected_node = self.systemTree.SelectedItem
            self.vm.update_selected_sheet(selected_node)
            self.update_action_buttons()
            self.vm.StatusText = "Restored notes view."
        except Exception: pass

    def ctx_select_parent(self, sender, args):
        try:
            selected_items = list(self.Grid_Notes.SelectedItems)
            if not selected_items: return
            
            item = selected_items[0]
            if hasattr(item, "ParentNode") and item.ParentNode:
                # Clear existing traces
                for node in self.vm.Sheets:
                    node.IsTraced = False
                    if hasattr(node, "Children") and node.Children:
                        for child in node.Children:
                            child.IsTraced = False
                            
                curr_node = getattr(item.ParentNode, "ParentNode", None)
                while curr_node:
                    curr_node.IsExpanded = True
                    curr_node = getattr(curr_node, "ParentNode", None)
                
                item.ParentNode.IsTraced = True
                
                try:
                    parent = getattr(item.ParentNode, "ParentNode", None)
                    if parent:
                        p_container = self.systemTree.ItemContainerGenerator.ContainerFromItem(parent)
                        if p_container:
                            p_container.IsExpanded = True
                            p_container.UpdateLayout()
                            c_container = p_container.ItemContainerGenerator.ContainerFromItem(item.ParentNode)
                            if c_container:
                                c_container.BringIntoView()
                    else:
                        container = self.systemTree.ItemContainerGenerator.ContainerFromItem(item.ParentNode)
                        if container:
                            container.BringIntoView()
                except: pass
        except Exception: pass

    def ctx_highlight_children(self, sender, args):
        try:
            selected_items = list(self.Grid_Notes.SelectedItems)
            if not selected_items: return
            
            target_parent = getattr(selected_items[0], "ParentNode", None)
            if not target_parent: return
            
            self.ctx_select_parent(sender, args)
            
            for node in self.vm.Sheets:
                node.IsChecked = False
                
            target_parent.IsChecked = True
            
            self.Grid_Notes.SelectAll()
        except Exception: pass


def main():
    xaml_file = os.path.join(os.path.dirname(__file__), "UI.xaml")
    window = ManageTextNotesWindow(xaml_file)
    window.show_dialog()

if __name__ == '__main__':
    main()
