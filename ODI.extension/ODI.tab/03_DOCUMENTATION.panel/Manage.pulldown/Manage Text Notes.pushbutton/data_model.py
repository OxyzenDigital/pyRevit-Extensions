# -*- coding: utf-8 -*-
from pyrevit import forms

class NodeBase(forms.Reactive):
    def __init__(self):
        forms.Reactive.__init__(self)
        self._is_checked = False
        self._is_selected = False
        self._is_expanded = False
        self._is_traced = False

    @property
    def IsTraced(self):
        return self._is_traced

    @IsTraced.setter
    def IsTraced(self, value):
        self._is_traced = value
        self.OnPropertyChanged('IsTraced')

    @property
    def IsChecked(self):
        return self._is_checked

    @IsChecked.setter
    def IsChecked(self, value):
        if self._is_checked == value: return
        self._is_checked = value
        self.OnPropertyChanged('IsChecked')
        
        # Propagate down
        if value is not None:
            if hasattr(self, 'Children') and self.Children:
                for child in self.Children:
                    child.set_checked_without_parent_update(value)
            if hasattr(self, 'Notes') and self.Notes:
                for note in self.Notes:
                    note.set_checked_without_parent_update(value)
                
        # Propagate up
        if hasattr(self, 'ParentNode') and self.ParentNode:
            self.ParentNode.update_check_state_from_children()
            
        if hasattr(self, 'StateCallback') and self.StateCallback:
            self.StateCallback()

    def set_checked_without_parent_update(self, value):
        if self._is_checked == value: return
        self._is_checked = value
        self.OnPropertyChanged('IsChecked')
        if hasattr(self, 'Children') and self.Children:
            for child in self.Children:
                child.set_checked_without_parent_update(value)
        if hasattr(self, 'Notes') and self.Notes:
            for note in self.Notes:
                note.set_checked_without_parent_update(value)

    def update_check_state_from_children(self):
        children = getattr(self, 'Children', []) or []
        notes = getattr(self, 'Notes', []) or []
        all_items = list(children) + list(notes)
        
        if not all_items: return
        
        all_checked = all(c.IsChecked == True for c in all_items)
        all_unchecked = all(c.IsChecked == False for c in all_items)
        
        if all_checked:
            new_val = True
        elif all_unchecked:
            new_val = False
        else:
            new_val = None
            
        if self._is_checked == new_val: return
        self._is_checked = new_val
        self.OnPropertyChanged('IsChecked')
        
        if hasattr(self, 'ParentNode') and self.ParentNode:
            self.ParentNode.update_check_state_from_children()
            
        if hasattr(self, 'StateCallback') and self.StateCallback:
            self.StateCallback()
        
    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged('IsSelected')
        
    @property
    def IsExpanded(self):
        return self._is_expanded

    @IsExpanded.setter
    def IsExpanded(self, value):
        self._is_expanded = value
        self.OnPropertyChanged('IsExpanded')
        if hasattr(self, 'StateCallback') and self.StateCallback:
            self.StateCallback()

class ProjectNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self)
        self.NodeType = "Project"
        self.Name = name
        self.Children = []
        self.ParentNode = None

class MasterNote(NodeBase):
    def __init__(self, key="", text="", parent_key="", project=""):
        NodeBase.__init__(self)
        self.NodeType = "MasterNote"
        self.Key = key
        self._text = text
        self.ParentKey = parent_key
        self.Project = project
        self._is_editing = False

    @property
    def IsEditing(self):
        return self._is_editing

    @IsEditing.setter
    def IsEditing(self, value):
        self._is_editing = value
        self.OnPropertyChanged('IsEditing')

    @property
    def Text(self):
        return self._text

    @Text.setter
    def Text(self, value):
        self._text = value
        self.OnPropertyChanged('Text')

class NoteItem(NodeBase):
    def __init__(self, element, viewport=None, sheet=None):
        NodeBase.__init__(self)
        self.Element = element
        self.Viewport = viewport
        self.Sheet = sheet
        self.Id = element.Id
        self._text = element.Text.replace('\r', ' ').replace('\n', ' ')
        self.ParentNode = None
        
        # Determine the view name
        if viewport:
            from pyrevit import revit
            view = revit.doc.GetElement(viewport.ViewId)
            self.ViewName = view.Name if view else "Unknown View"
        else:
            self.ViewName = sheet.Name if sheet else "Unknown Sheet"

    @property
    def FormattedViewName(self):
        name = self.ViewName
        if name and len(name) > 15:
            return name[:6] + "..." + name[-6:]
        return name

    @property
    def Text(self):
        return self._text

    @Text.setter
    def Text(self, value):
        self._text = value
        self.OnPropertyChanged('Text')

class SheetItem(NodeBase):
    def __init__(self, element, notes):
        NodeBase.__init__(self)
        self.NodeType = "Sheet"
        self.Element = element
        self.Id = element.Id
        self.Sheet = element
        self.Name = "{} - {}".format(element.SheetNumber, element.Name)
        self.Notes = notes
        self.NoteCount = len(notes)
        self.ParentNode = None

        # Link parent back to child notes
        for n in self.Notes:
            n.ParentNode = self

    # Return empty list to support HierarchicalDataTemplate
    @property
    def Children(self):
        return []

class SheetSetNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self)
        self.NodeType = "SheetSet"
        self.Name = name
        self.Children = []
        self.ParentNode = None
