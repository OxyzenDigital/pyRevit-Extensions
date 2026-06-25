# -*- coding: utf-8 -*-
import re
import difflib
import clr

clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Input import ICommand
from Autodesk.Revit.DB import ElementId

# --- Utility Commands ---
class RelayCommand(ICommand):
    def __init__(self, action):
        self.action = action
    def add_CanExecuteChanged(self, handler): pass
    def remove_CanExecuteChanged(self, handler): pass
    def CanExecute(self, parameter): return True
    def Execute(self, parameter): self.action()

def generate_char_diff(original, current):
    if not original or not current or original == current:
        return False, original or "", "", "", ""
    orig_len = len(original)
    curr_len = len(current)
    min_len = min(orig_len, curr_len)
    
    prefix_len = 0
    while prefix_len < min_len and original[prefix_len] == current[prefix_len]:
        prefix_len += 1
        
    suffix_len = 0
    while suffix_len < min_len - prefix_len and original[orig_len - 1 - suffix_len] == current[curr_len - 1 - suffix_len]:
        suffix_len += 1
        
    prefix = original[:prefix_len]
    suffix = original[orig_len - suffix_len:] if suffix_len > 0 else ""
    old_mid = original[prefix_len:orig_len - suffix_len]
    new_mid = current[prefix_len:curr_len - suffix_len]
    
    return True, prefix, old_mid, new_mid, suffix

# --- ViewModels ---

class ViewModelBase(INotifyPropertyChanged):
    def __init__(self):
        self._property_changed_handlers = []
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    def OnPropertyChanged(self, property_name):
        args = PropertyChangedEventArgs(property_name)
        for handler in self._property_changed_handlers:
            handler(self, args)

class DiffNode(ViewModelBase):
    def __init__(self, original_val, action_callback=None):
        ViewModelBase.__init__(self)
        self._original_val = original_val
        self._proposed_val = original_val
        self._action_callback = action_callback
        
        self.HasDiff = False
        self.Prefix = original_val or ""
        self.Old = ""
        self.New = ""
        self.Suffix = ""
        
    @property
    def OriginalValue(self): return self._original_val
    @property
    def ProposedValue(self): return self._proposed_val
    @ProposedValue.setter
    def ProposedValue(self, val):
        self._proposed_val = val
        self.HasDiff, self.Prefix, self.Old, self.New, self.Suffix = generate_char_diff(self._original_val, val)
        self.OnPropertyChanged("ProposedValue")
        self.OnPropertyChanged("HasDiff")
        self.OnPropertyChanged("Prefix")
        self.OnPropertyChanged("Old")
        self.OnPropertyChanged("New")
        self.OnPropertyChanged("Suffix")
        if self._action_callback:
            self._action_callback()

class SelectableNode(ViewModelBase):
    def __init__(self, name, is_checked=True, callback=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self._is_checked = is_checked
        self.callback = callback
    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = val
        self.OnPropertyChanged("IsChecked")
        if self.callback: self.callback()

class DisciplineGroupNode(ViewModelBase):
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self.Name = name
        self.Children = []

class CollectionGroupNode(ViewModelBase):
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self.Name = name
        self.Children = []  # Holds DisciplineGroupNode instances

class ViewViewModel(ViewModelBase):
    def __init__(self, view_id, name, view_type="FloorPlan", scale="1/8\" = 1'-0\"", is_new=False):
        ViewModelBase.__init__(self)
        self.ViewId = view_id
        self._name = name
        self._view_type = view_type
        self._scale = scale
        self._is_new = is_new
        
    @property
    def Name(self): return self._name
    @Name.setter
    def Name(self, val):
        self._name = val
        self.OnPropertyChanged("Name")
        
    @property
    def ViewType(self): return self._view_type
    @ViewType.setter
    def ViewType(self, val):
        self._view_type = val
        self.OnPropertyChanged("ViewType")

    @property
    def Scale(self): return self._scale
    @Scale.setter
    def Scale(self, val):
        self._scale = val
        self.OnPropertyChanged("Scale")
        
class SheetViewModel(ViewModelBase):
    def __init__(self, element_id, number, name, collection_name, is_template=False, validation_callback=None):
        ViewModelBase.__init__(self)
        self.validation_callback = validation_callback
        self.ElementId = element_id
        self.IsTemplate = is_template
        self._sheet_number = number
        self._sheet_name = name
        self._collection_name = collection_name
        self.OriginalNumber = number
        self.OriginalName = name
        
        self.Views = ObservableCollection[ViewViewModel]()
        self.AvailableNames = ObservableCollection[str]()
        
        self._is_checked = False
        self._is_name_unique = True
        self._action = "MATCHED" if not is_template else "CREATE"
        
        self.PurgeCommand = RelayCommand(self.mark_purge)
        self.AddViewCommand = RelayCommand(self.add_view)
        self.UndoCommand = RelayCommand(self.undo_changes)

    @property
    def SheetNumber(self): return self._sheet_number
    @SheetNumber.setter
    def SheetNumber(self, val):
        if val is None: val = ""
        self._sheet_number = str(val)
        self.OnPropertyChanged("SheetNumber")
        self.update_action()

    @property
    def SheetName(self): return self._sheet_name
    @SheetName.setter
    def SheetName(self, val):
        if val is None: val = ""
        self._sheet_name = str(val)
        self.OnPropertyChanged("SheetName")
        self.update_action()

    @property
    def CollectionName(self): return self._collection_name
    @CollectionName.setter
    def CollectionName(self, val):
        self._collection_name = val
        self.OnPropertyChanged("CollectionName")

    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = val
        self.OnPropertyChanged("IsChecked")
        
    @property
    def IsNameUnique(self): return self._is_name_unique
    @IsNameUnique.setter
    def IsNameUnique(self, val):
        self._is_name_unique = val
        self.OnPropertyChanged("IsNameUnique")

    @property
    def Action(self): return self._action
    @Action.setter
    def Action(self, val):
        self._action = val
        self.OnPropertyChanged("Action")

    @property
    def DisciplineName(self):
        match = re.match(r"^([A-Z]+)[- ]?(\d+)", self.SheetNumber.upper())
        disc_code = match.group(1) if match else "Other"
        disc_map = { "A": "Architectural", "S": "Structural", "M": "Mechanical", "E": "Electrical", "P": "Plumbing", "C": "Civil", "L": "Landscape", "F": "Fire Protection", "G": "General", "I": "Interiors" }
        name = disc_map.get(disc_code, "Discipline")
        return "{} - {}".format(disc_code, name) if disc_code != "Other" else "Uncategorized"

    def mark_purge(self, parameter=None):
        self.Action = "PURGE"
        self.IsChecked = True
        
    def undo_changes(self, parameter=None):
        self.SheetNumber = self.OriginalNumber
        self.SheetName = self.OriginalName
        
        # Remove newly added views
        new_views = [v for v in self.Views if getattr(v, '_is_new', False)]
        for nv in new_views:
            self.Views.Remove(nv)
            
        if self.IsTemplate:
            self.Action = "CREATE"
            self.IsChecked = True
        else:
            self.Action = "MATCHED"
            self.IsChecked = False
            
        if self.validation_callback: self.validation_callback()
        
    def add_view(self, parameter=None):
        self.Views.Add(ViewViewModel(ElementId.InvalidElementId, "New View", is_new=True))

    def update_action(self):
        if self.IsTemplate: 
            if self.validation_callback: self.validation_callback()
            return
        if self.SheetNumber != self.OriginalNumber or self.SheetName != self.OriginalName:
            self.Action = "UPDATE"
            self.IsChecked = True
        else:
            self.Action = "MATCHED"
            self.IsChecked = False
        if self.validation_callback: self.validation_callback()

class NavTreeNode(ViewModelBase):
    def __init__(self, name, node_type, tag=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self.NodeType = node_type
        self.Tag = tag
        self.callback = None
        self.Children = ObservableCollection[NavTreeNode]()
        self._is_expanded = True
        self._is_selected = False
        self._is_checked = False
        self._count = 0
        
    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, val):
        self._is_expanded = val
        self.OnPropertyChanged("IsExpanded")

    @property
    def IsSelected(self): return self._is_selected
    @IsSelected.setter
    def IsSelected(self, val):
        self._is_selected = val
        self.OnPropertyChanged("IsSelected")

    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = val
        self.OnPropertyChanged("IsChecked")
        if self.callback: self.callback()
        
    @property
    def Count(self): return self._count
    @Count.setter
    def Count(self, val):
        self._count = val
        self.OnPropertyChanged("Count")
        self.OnPropertyChanged("ShowCount")
        self.OnPropertyChanged("DisplayCount")
        
    @property
    def ShowCount(self): return self.Count > 0
    @property
    def DisplayCount(self):
        if self.NodeType == "Sheet": return " ({} Views)".format(self.Count)
        return " ({})".format(self.Count)
        
    @property
    def FontWeight(self): return "Bold" if self.NodeType in ["Root", "Collection"] else "Normal"
