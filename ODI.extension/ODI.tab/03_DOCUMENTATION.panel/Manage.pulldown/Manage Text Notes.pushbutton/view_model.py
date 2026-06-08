# -*- coding: utf-8 -*-
import os
import io
import json
import uuid
from pyrevit import forms, revit, script, DB
from System.Collections.Generic import List

from data_model import MasterNote, SheetItem, NoteItem, SheetSetNode

doc = revit.doc
uidoc = revit.uidoc
cfg = script.get_config()

class ManageTextNotesViewModel(forms.Reactive):
    def __init__(self):
        forms.Reactive.__init__(self)
        
        self._sheets = []
        self._current_notes = []
        
        self._master_notes = []
        self._filtered_master_notes = []
        self._master_tree = []
        self._selected_master_note = None
        self._master_search_text = ""
        
        self._status_text = "Ready."
        self._file_path = "Unsaved Canvas"
        self._zoom_scale = 1.5
        
        # UI Button States
        self._can_tree_select_all = False
        self._can_tree_select_none = False
        self._can_tree_expand_all = False
        self._can_tree_collapse_all = False
        
        self._can_grid_select_all = False
        self._can_grid_select_none = False
        
        self._can_master_expand_all = True
        self._can_master_collapse_all = True
        
        self._can_replace_selected = False
        self._can_replace_all = False
        
        self.ActionCallback = None
        
        self.initialize_data()

    @property
    def Sheets(self): return self._sheets
    @Sheets.setter
    def Sheets(self, value):
        self._sheets = value
        self.OnPropertyChanged('Sheets')

    @property
    def CurrentNotes(self): return self._current_notes
    @CurrentNotes.setter
    def CurrentNotes(self, value):
        self._current_notes = value
        self.OnPropertyChanged('CurrentNotes')

    @property
    def MasterNotes(self): return self._master_notes
    @MasterNotes.setter
    def MasterNotes(self, value):
        self._master_notes = value
        self.OnPropertyChanged('MasterNotes')
        self.filter_master_notes()

    @property
    def FilteredMasterNotes(self): return self._filtered_master_notes
    @FilteredMasterNotes.setter
    def FilteredMasterNotes(self, value):
        self._filtered_master_notes = value
        self.OnPropertyChanged('FilteredMasterNotes')
        self.build_master_tree()

    @property
    def MasterTree(self): return self._master_tree
    @MasterTree.setter
    def MasterTree(self, value):
        self._master_tree = value
        self.OnPropertyChanged('MasterTree')

    @property
    def SelectedMasterNote(self): return self._selected_master_note
    @SelectedMasterNote.setter
    def SelectedMasterNote(self, value):
        self._selected_master_note = value
        self.OnPropertyChanged('SelectedMasterNote')

    @property
    def MasterSearchText(self): return self._master_search_text
    @MasterSearchText.setter
    def MasterSearchText(self, value):
        self._master_search_text = value
        self.OnPropertyChanged('MasterSearchText')
        self.filter_master_notes()

    @property
    def StatusText(self): return self._status_text
    @StatusText.setter
    def StatusText(self, value):
        self._status_text = value
        self.OnPropertyChanged('StatusText')

    @property
    def FilePath(self): return self._file_path
    @FilePath.setter
    def FilePath(self, value):
        self._file_path = value
        self.OnPropertyChanged('FilePath')

    @property
    def ZoomScale(self): return self._zoom_scale
    @ZoomScale.setter
    def ZoomScale(self, value):
        self._zoom_scale = value
        self.OnPropertyChanged('ZoomScale')

    # --- Button States ---
    @property
    def CanTreeSelectAll(self): return self._can_tree_select_all
    @CanTreeSelectAll.setter
    def CanTreeSelectAll(self, value):
        self._can_tree_select_all = value
        self.OnPropertyChanged('CanTreeSelectAll')

    @property
    def CanTreeSelectNone(self): return self._can_tree_select_none
    @CanTreeSelectNone.setter
    def CanTreeSelectNone(self, value):
        self._can_tree_select_none = value
        self.OnPropertyChanged('CanTreeSelectNone')

    @property
    def CanTreeExpandAll(self): return self._can_tree_expand_all
    @CanTreeExpandAll.setter
    def CanTreeExpandAll(self, value):
        self._can_tree_expand_all = value
        self.OnPropertyChanged('CanTreeExpandAll')

    @property
    def CanTreeCollapseAll(self): return self._can_tree_collapse_all
    @CanTreeCollapseAll.setter
    def CanTreeCollapseAll(self, value):
        self._can_tree_collapse_all = value
        self.OnPropertyChanged('CanTreeCollapseAll')

    @property
    def CanGridSelectAll(self): return self._can_grid_select_all
    @CanGridSelectAll.setter
    def CanGridSelectAll(self, value):
        self._can_grid_select_all = value
        self.OnPropertyChanged('CanGridSelectAll')

    @property
    def CanGridSelectNone(self): return self._can_grid_select_none
    @CanGridSelectNone.setter
    def CanGridSelectNone(self, value):
        self._can_grid_select_none = value
        self.OnPropertyChanged('CanGridSelectNone')

    @property
    def CanMasterExpandAll(self): return self._can_master_expand_all
    @CanMasterExpandAll.setter
    def CanMasterExpandAll(self, value):
        self._can_master_expand_all = value
        self.OnPropertyChanged('CanMasterExpandAll')

    @property
    def CanMasterCollapseAll(self): return self._can_master_collapse_all
    @CanMasterCollapseAll.setter
    def CanMasterCollapseAll(self, value):
        self._can_master_collapse_all = value
        self.OnPropertyChanged('CanMasterCollapseAll')

    @property
    def CanReplaceSelected(self): return self._can_replace_selected
    @CanReplaceSelected.setter
    def CanReplaceSelected(self, value):
        self._can_replace_selected = value
        self.OnPropertyChanged('CanReplaceSelected')

    @property
    def CanReplaceAll(self): return self._can_replace_all
    @CanReplaceAll.setter
    def CanReplaceAll(self, value):
        self._can_replace_all = value
        self.OnPropertyChanged('CanReplaceAll')

    def evaluate_states(self):
        # Tree Check States
        has_unchecked = False
        has_checked = False
        
        # Tree Expand States
        has_collapsed = False
        has_expanded = False
        
        def traverse_tree(nodes):
            for n in nodes:
                # Need `nonlocal` equivalent in Py2, so use list trick
                state_flags[0] = state_flags[0] or (n.IsChecked != True)
                state_flags[1] = state_flags[1] or (n.IsChecked == True)
                state_flags[2] = state_flags[2] or (n.IsExpanded == False and hasattr(n, 'Children') and n.Children)
                state_flags[3] = state_flags[3] or (n.IsExpanded == True and hasattr(n, 'Children') and n.Children)
                
                if hasattr(n, 'Children') and n.Children:
                    traverse_tree(n.Children)
                    
        if self.Sheets:
            state_flags = [False, False, False, False]
            traverse_tree(self.Sheets)
            self.CanTreeSelectAll = state_flags[0]
            self.CanTreeSelectNone = state_flags[1]
            self.CanTreeExpandAll = state_flags[2]
            self.CanTreeCollapseAll = state_flags[3]
        else:
            self.CanTreeSelectAll = False
            self.CanTreeSelectNone = False
            self.CanTreeExpandAll = False
            self.CanTreeCollapseAll = False
            
        # Grid Check States
        self.CanGridSelectAll = False
        self.CanGridSelectNone = False

        if self.ActionCallback:
            self.ActionCallback()

    def inject_callbacks(self, nodes):
        for n in nodes:
            n.StateCallback = self.evaluate_states
            if hasattr(n, 'Children') and n.Children:
                self.inject_callbacks(n.Children)

    def build_master_tree(self):
        from data_model import ProjectNode
        projects = {}
        for note in self.FilteredMasterNotes:
            p_name = note.Project if note.Project else "Uncategorized"
            if p_name not in projects:
                projects[p_name] = ProjectNode(p_name)
            
            projects[p_name].Children.append(note)
            note.ParentNode = projects[p_name]
            
        self.MasterTree = sorted(projects.values(), key=lambda p: p.Name)

    # --- Methods ---

    def initialize_data(self):
        self.scan_active()

    def load_master_file(self, filepath):
        self.FilePath = filepath
        cfg.master_notes_file = filepath
        script.save_config()
        
        loaded_notes = []
        try:
            if filepath.lower().endswith('.json'):
                with io.open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for item in data:
                        loaded_notes.append(MasterNote(
                            key=item.get("Key", ""),
                            text=item.get("Text", ""),
                            parent_key=item.get("ParentKey", ""),
                            project=item.get("Project", "")
                        ))
            else:
                # Fallback for old txt format
                with io.open(filepath, 'r', encoding='utf-8') as f:
                    for line in f:
                        parts = line.strip().split('\t')
                        if len(parts) >= 2:
                            key = parts[0]
                            text = parts[1]
                            parent = parts[2] if len(parts) > 2 else ""
                            if text:
                                loaded_notes.append(MasterNote(key, text, parent, ""))
                                
            self.MasterNotes = loaded_notes
            self.StatusText = "Loaded {} master notes.".format(len(self.MasterNotes))
        except Exception as e:
            self.StatusText = "Error loading file: {}".format(str(e))



    def export_selected_to_new_file(self, selected_notes):
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
        default_name = 'ExportedNotes_{}.json'.format(timestamp)
        filepath = forms.save_file(file_ext='json', default_name=default_name, title="Export Selected Notes")
        if not filepath:
            return
        
        project_name = doc.Title
        data = []
        for item in selected_notes:
            data.append({
                "Key": str(uuid.uuid4())[:8],
                "Text": item.Text,
                "ParentKey": "",
                "Project": project_name
            })
            
        try:
            with io.open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            self.StatusText = "Exported {} notes to {}.".format(len(selected_notes), os.path.basename(filepath))
            forms.alert("Successfully exported to:\n" + filepath)
        except Exception as e:
            forms.alert("Error exporting file: {}".format(str(e)))

    def save_master_file(self, save_as=False):
        filepath = self.FilePath
        if save_as or not self.FilePath or self.FilePath == "No file selected...":
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
            default_name = 'NotesFile_{}.json'.format(timestamp)
            filepath = forms.save_file(file_ext='json', default_name=default_name, title="Save Notes File")
            if not filepath:
                return
            self.FilePath = filepath
            cfg.master_notes_file = filepath
            script.save_config()
            
        # If the file path is a .txt, prompt to save as JSON instead
        filepath = self.FilePath
        if filepath.lower().endswith('.txt'):
            filepath = filepath[:-4] + ".json"
            self.FilePath = filepath
            cfg.master_notes_file = filepath
            script.save_config()

        try:
            data = []
            for n in self.MasterNotes:
                data.append({
                    "Key": n.Key,
                    "Text": n.Text,
                    "ParentKey": n.ParentKey,
                    "Project": n.Project
                })
            with io.open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            self.StatusText = "Successfully saved {} notes to {}.".format(len(self.MasterNotes), os.path.basename(filepath))
        except Exception as e:
            forms.alert("Error saving file: {}".format(str(e)))

    def pick_file(self):
        filepath = forms.pick_file(file_ext='json', title="Select Notes File")
        if not filepath:
            filepath = forms.pick_file(file_ext='txt', title="Select Legacy Notes File")
        if filepath:
            self.load_master_file(filepath)

    def filter_master_notes(self):
        if not self.MasterSearchText:
            self.FilteredMasterNotes = self.MasterNotes
            return
            
        search_lower = self.MasterSearchText.lower()
        filtered = []
        for note in self.MasterNotes:
            if search_lower in note.Text.lower() or search_lower in note.Project.lower() or search_lower in note.Key.lower():
                filtered.append(note)
        self.FilteredMasterNotes = filtered

    def add_new_master_note(self):
        new_key = str(uuid.uuid4())[:8]
        new_note = MasterNote(key=new_key, text="New Note", project=doc.Title)
        new_note.IsEditing = True
        self.MasterNotes.insert(0, new_note)
        self.MasterNotes = list(self.MasterNotes)
        self.filter_master_notes()
        self.expand_all_master(True)
        self.SelectedMasterNote = new_note
        self.StatusText = "Added new note."

    def add_note_below(self, master_note):
        try:
            idx = self.MasterNotes.index(master_note)
            new_key = str(uuid.uuid4())[:8]
            new_note = MasterNote(key=new_key, text="New Note", project=doc.Title)
            new_note.IsEditing = True
            self.MasterNotes.insert(idx + 1, new_note)
            self.MasterNotes = list(self.MasterNotes)
            self.filter_master_notes()
            self.expand_all_master(True)
            self.SelectedMasterNote = new_note
            self.StatusText = "Added new note."
        except Exception as e:
            forms.alert("Could not add note: " + str(e))

    def delete_specific_master_note(self, master_note):
        res = forms.alert("Are you sure you want to delete this note from the Notes File?", yes=True, no=True)
        if res:
            try:
                self.MasterNotes.remove(master_note)
                self.MasterNotes = list(self.MasterNotes)
                if self.SelectedMasterNote == master_note:
                    self.SelectedMasterNote = None
                self.filter_master_notes()
                self.expand_all_master(True)
                self.StatusText = "Deleted note."
            except Exception as e:
                forms.alert("Could not delete note: " + str(e))

    def append_project_notes_to_master(self, note_items):
        count = 0
        project_name = doc.Title
        for item in note_items:
            # Check if text already exists
            exists = any(m.Text == item.Text for m in self.MasterNotes)
            if not exists:
                new_key = str(uuid.uuid4())[:8]
                new_note = MasterNote(key=new_key, text=item.Text, project=project_name)
                self.MasterNotes.append(new_note)
                count += 1
        
        if count > 0:
            self.MasterNotes = list(self.MasterNotes)
            self.filter_master_notes()
            self.expand_all_master(True)
            self.StatusText = "Appended {} new notes to Master Database.".format(count)
        else:
            self.StatusText = "No new unique notes appended."

    def get_notes_from_checked_sheets(self):
        notes = []
        def traverse(nodes):
            for n in nodes:
                if getattr(n, 'NodeType', '') == "Sheet" and n.IsChecked:
                    notes.extend(n.Notes)
                if hasattr(n, 'Children') and n.Children:
                    traverse(n.Children)
        traverse(self.Sheets)
        return notes

    # --- Revit Data Gathering ---
    def _get_text_notes_for_sheet(self, sheet):
        notes = []
        
        # Notes directly on the sheet
        sheet_notes = (DB.FilteredElementCollector(doc, sheet.Id)
                 .OfClass(DB.TextNote)
                 .ToElements())
        for n in sheet_notes:
            notes.append(NoteItem(n, viewport=None, sheet=sheet))

        # Notes inside viewports on the sheet
        for vp_id in sheet.GetAllViewports():
            vp = doc.GetElement(vp_id)
            view = doc.GetElement(vp.ViewId)
            
            vp_notes = (DB.FilteredElementCollector(doc, view.Id)
                        .OfClass(DB.TextNote)
                        .ToElements())
            for n in vp_notes:
                notes.append(NoteItem(n, viewport=vp, sheet=sheet))
                
        return notes

    def build_sheet_node(self, sheet):
        notes = self._get_text_notes_for_sheet(sheet)
        if notes:
            return SheetItem(sheet, notes)
        return None

    def scan_active(self):
        active_view = doc.ActiveView
        if isinstance(active_view, DB.ViewSheet):
            self.StatusText = "Scanning active sheet..."
            root = SheetSetNode("Current View")
            s_node = self.build_sheet_node(active_view)
            if s_node:
                s_node.ParentNode = root
                root.Children.append(s_node)
            self.Sheets = [root]
            self.inject_callbacks(self.Sheets)
            self.evaluate_states()
            self.StatusText = "Scanned active sheet."
        else:
            forms.alert("The active view is not a Sheet.\n\nFalling back to scanning the entire project for sheets...", warn_icon=True)
            self.scan_project()

    def scan_project(self):
        self.StatusText = "Scanning project for sheets with Text Notes..."
        tree_nodes = []
        
        # All Sheets Node
        all_sheets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()
        all_sheets = sorted(all_sheets, key=lambda s: s.SheetNumber)
        
        all_node = SheetSetNode("< All Sheets >")
        for sheet in all_sheets:
            s_node = self.build_sheet_node(sheet)
            if s_node:
                s_node.ParentNode = all_node
                all_node.Children.append(s_node)
        if all_node.Children:
            tree_nodes.append(all_node)
            
        # Print Sets Nodes
        sheet_sets = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet).ToElements()
        for sset in sorted(sheet_sets, key=lambda s: s.Name):
            sset_node = SheetSetNode(sset.Name)
            for sheet in sset.Views:
                if isinstance(sheet, DB.ViewSheet):
                    s_node = self.build_sheet_node(sheet)
                    if s_node:
                        s_node.ParentNode = sset_node
                        sset_node.Children.append(s_node)
            if sset_node.Children:
                tree_nodes.append(sset_node)
                
        self.Sheets = tree_nodes
        self.inject_callbacks(tree_nodes)
        self.evaluate_states()
        self.StatusText = "Scanned {} sheet sets.".format(len(tree_nodes)-1)

    def get_all_notes_in_nodes(self, nodes):
        notes = []
        for node in nodes:
            if hasattr(node, "Notes") and node.Notes:
                notes.extend(node.Notes)
            elif hasattr(node, "Children") and node.Children:
                notes.extend(self.get_all_notes_in_nodes(node.Children))
        return notes

    def update_selected_sheet(self, selected_node):
        if not selected_node:
            self.CurrentNotes = self.get_all_notes_in_nodes(self.Sheets)
            return

        if selected_node.NodeType == "SheetSet":
            self.CurrentNotes = self.get_all_notes_in_nodes([selected_node])
        else:
            self.CurrentNotes = selected_node.Notes
        self.evaluate_states()

    def select_all_sheets(self, state=True):
        def set_checked(nodes, st):
            for n in nodes:
                n.IsChecked = st
                if hasattr(n, "Children") and n.Children:
                    set_checked(n.Children, st)
        set_checked(self.Sheets, state)
        self.evaluate_states()

    def expand_all_sheets(self, state=True):
        def set_expanded(nodes, st):
            for n in nodes:
                n.IsExpanded = st
                if hasattr(n, "Children") and n.Children:
                    set_expanded(n.Children, st)
        set_expanded(self.Sheets, state)
        self.evaluate_states()

    def expand_all_master(self, state=True):
        def set_expanded(nodes, st):
            for n in nodes:
                n.IsExpanded = st
                if hasattr(n, "Children") and n.Children:
                    set_expanded(n.Children, st)
        if self.MasterTree:
            set_expanded(self.MasterTree, state)

    def select_all_grid(self, state=True):
        pass

    def zoom_to_note(self, note_item):
        if not note_item: return
        
        sheet = note_item.Sheet
        viewport = note_item.Viewport
        
        try:
            if doc.ActiveView.Id != sheet.Id:
                uidoc.ActiveView = sheet
                
            if viewport:
                elem = doc.GetElement(viewport.Id)
                elem_ids = List[DB.ElementId]([viewport.Id])
            else:
                elem = doc.GetElement(note_item.Id)
                elem_ids = List[DB.ElementId]([note_item.Id])
                
            uidoc.Selection.SetElementIds(elem_ids)

            # BoundingBox Zoom Math
            bbox = elem.get_BoundingBox(sheet)
            if bbox:
                min_pt = bbox.Min
                max_pt = bbox.Max
                
                center = (min_pt + max_pt) / 2.0
                width = max_pt.X - min_pt.X
                height = max_pt.Y - min_pt.Y
                
                scale = self.ZoomScale
                new_width = width * scale
                new_height = height * scale
                
                if new_width < 0.1: new_width = 1.0
                if new_height < 0.1: new_height = 1.0
                
                new_min = DB.XYZ(center.X - new_width/2.0, center.Y - new_height/2.0, min_pt.Z)
                new_max = DB.XYZ(center.X + new_width/2.0, center.Y + new_height/2.0, max_pt.Z)
                
                for uiview in uidoc.GetOpenUIViews():
                    if uiview.ViewId == sheet.Id:
                        uiview.ZoomAndCenterRectangle(new_min, new_max)
                        break
            else:
                uidoc.ShowElements(elem_ids)
                
        except Exception as e:
            forms.alert("Could not zoom to note: {}".format(str(e)))

    def update_note(self, note_item, master_note):
        # Kept for single-updates if needed elsewhere, but using transaction outside.
        if not note_item or not master_note:
            return False
            
        note_elem = note_item.Element
        new_text = master_note.Text

        if note_elem.Text == new_text:
            return False

        try:
            note_elem.Text = new_text
            note_item.Text = new_text.replace('\r', ' ').replace('\n', ' ')
            return True
        except Exception as e:
            forms.alert("Error updating note: {}".format(str(e)))
            return False

    def replace_selected_project_notes(self, selected_notes):
        if not self.SelectedMasterNote:
            forms.alert("Please select a Master Note from the Right Pane.")
            return
            
        count = 0
        try:
            with revit.Transaction("Batch Update Text Notes"):
                for note_item in selected_notes:
                    if self.update_note(note_item, self.SelectedMasterNote):
                        count += 1
        except Exception as e:
            import traceback
            forms.alert("Transaction Error:\n" + traceback.format_exc())
                
        self.StatusText = "Successfully replaced {} selected notes.".format(count)

    def replace_all_checked_sheets(self):
        if not self.SelectedMasterNote:
            forms.alert("Please select a Master Note from the Right Pane.")
            return

        checked_sheets = []
        def get_checked(nodes):
            for n in nodes:
                if n.NodeType == "Sheet" and n.IsChecked:
                    checked_sheets.append(n)
                if hasattr(n, "Children") and n.Children:
                    get_checked(n.Children)
                    
        get_checked(self.Sheets)
        
        unique_sheets = {s.Id: s for s in checked_sheets}.values()

        if not unique_sheets:
            forms.alert("No sheets are checked in the Tree View.")
            return

        count = 0
        with revit.TransactionGroup("Batch Replace Text Notes"):
            for sheet in unique_sheets:
                for note in sheet.Notes:
                    if self.update_note(note, self.SelectedMasterNote):
                        count += 1
                        
        self.StatusText = "Successfully replaced {} notes across {} sheets.".format(count, len(unique_sheets))
