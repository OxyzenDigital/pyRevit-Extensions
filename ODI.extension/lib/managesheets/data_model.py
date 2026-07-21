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
    def __init__(self, action, can_execute=None):
        self.action = action
        self.can_execute = can_execute
        self._can_execute_handlers = []
        
    def add_CanExecuteChanged(self, handler):
        self._can_execute_handlers.append(handler)
        
    def remove_CanExecuteChanged(self, handler):
        if handler in self._can_execute_handlers:
            self._can_execute_handlers.remove(handler)
            
    def CanExecute(self, parameter):
        if self.can_execute:
            return self.can_execute()
        return True
        
    def Execute(self, parameter):
        if self.CanExecute(parameter):
            self.action()

    def RaiseCanExecuteChanged(self):
        import System
        for handler in self._can_execute_handlers:
            handler(self, System.EventArgs.Empty)

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
    def __init__(self, name, is_checked=True, callback=None, display_name=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self.DisplayName = display_name if display_name else name
        self._is_checked = is_checked
        self._is_enabled = True
        self.callback = callback
        
    @property
    def IsEnabled(self): return self._is_enabled
    @IsEnabled.setter
    def IsEnabled(self, val):
        self._is_enabled = val
        self.OnPropertyChanged("IsEnabled")
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

class AssignableViewNode(ViewModelBase):
    def __init__(self, view_id, name, is_placed=False, sheet_number=""):
        ViewModelBase.__init__(self)
        self.Id = view_id
        self.Name = name
        self.IsPlaced = is_placed
        self.SheetNumber = sheet_number
        self.StatusColor = "#9CA3AF" if is_placed else "#3B82F6"
        self.DisplayName = "{} (on Sheet {})".format(name, sheet_number) if is_placed and sheet_number else name

class ViewViewModel(ViewModelBase):
    def __init__(self, view_id, name, view_type="FloorPlan", scale="1/8\" = 1'-0\"", is_new=False, level_name="", parent_sheet=None, view_number=""):
        ViewModelBase.__init__(self)
        self.ViewId = view_id
        self._name = name
        self._view_type = view_type
        self._scale = scale
        self._is_new = is_new
        v_num_str = str(view_number).strip() if view_number is not None else ""
        if v_num_str.lower() in ["nan", "none", "null"]: v_num_str = ""
        self._view_number = v_num_str
        self._assignment_mode = "Existing" if not is_new else "New"
        self._source_view_id = ElementId.InvalidElementId
        self._validation_warning = ""
        self._level_name = level_name
        self.ParentSheet = parent_sheet
        self.OriginalName = name
        self.OriginalViewNumber = v_num_str
        
        self.AvailableLevels = AVAILABLE_LEVELS
        self.AvailableViewFamilyTypes = AVAILABLE_VIEW_FAMILY_TYPES
        
        # Attempt to infer level if new view
        if self._is_new and not self._level_name and self.ParentSheet:
            matched_lvl = None
            search_strings = [self._name.upper()]
            if self.ParentSheet.SheetName:
                search_strings.append(self.ParentSheet.SheetName.upper())
                
            for s_str in search_strings:
                for lvl in self.AvailableLevels:
                    if lvl.upper() in s_str:
                        matched_lvl = lvl
                        break
                if matched_lvl: break
                
            if not matched_lvl and self.ParentSheet.SheetNumber:
                import re
                m = re.search(r'\d+', self.ParentSheet.SheetNumber)
                if m:
                    num_str = m.group()
                    for lvl in self.AvailableLevels:
                        if num_str in lvl:
                            matched_lvl = lvl
                            break
                            
            if matched_lvl:
                self._level_name = matched_lvl
                        
        if not self._view_type and self.AvailableViewFamilyTypes:
            self._view_type = self.AvailableViewFamilyTypes[0]
            
        self.AvailableScales = [
            "12\" = 1'-0\"", "6\" = 1'-0\"", "3\" = 1'-0\"", "1 1/2\" = 1'-0\"",
            "1\" = 1'-0\"", "3/4\" = 1'-0\"", "1/2\" = 1'-0\"", "3/8\" = 1'-0\"",
            "1/4\" = 1'-0\"", "3/16\" = 1'-0\"", "1/8\" = 1'-0\"", "1\" = 10'-0\"",
            "3/32\" = 1'-0\"", "1/16\" = 1'-0\"", "1\" = 20'-0\"", "3/64\" = 1'-0\"",
            "1\" = 30'-0\"", "1/32\" = 1'-0\"", "1\" = 40'-0\"", "1\" = 50'-0\"",
            "1\" = 60'-0\"", "1/64\" = 1'-0\"", "1\" = 80'-0\"", "1\" = 100'-0\"",
            "1\" = 160'-0\"", "1\" = 200'-0\"", "1\" = 300'-0\"", "1\" = 400'-0\""
        ]
        
        from pyrevit.forms import Reactive
        self.DeleteCommand = RelayCommand(self.delete_view)
        self.MoveUpCommand = RelayCommand(self.move_up)
        self.MoveDownCommand = RelayCommand(self.move_down)
        
    def delete_view(self, param=None):
        if self.ParentSheet:
            self.ParentSheet.Views.Remove(self)
            self.ParentSheet.reassign_view_numbers()
            
    def move_up(self, param=None):
        if self.ParentSheet:
            self.ParentSheet.move_view_up(self)
            
    def move_down(self, param=None):
        if self.ParentSheet:
            self.ParentSheet.move_view_down(self)
        
    def trigger_validation(self):
        if self.ParentSheet and hasattr(self.ParentSheet, 'validation_callback') and self.ParentSheet.validation_callback:
            self.ParentSheet.validation_callback()

    @property
    def AssignmentMode(self): return self._assignment_mode
    @AssignmentMode.setter
    def AssignmentMode(self, val):
        self._assignment_mode = val
        self._is_new = (val == "New")
        self.OnPropertyChanged("AssignmentMode")
        self.OnPropertyChanged("IsNew")
        self.OnPropertyChanged("IsExisting")
        self.trigger_validation()
        
    @property
    def IsNew(self): return self._is_new
    
    @property
    def IsExisting(self): return not self._is_new

    @property
    def LevelName(self): return self._level_name
    @LevelName.setter
    def LevelName(self, val):
        self._level_name = val
        self.OnPropertyChanged("LevelName")
        self.check_semantic_mismatch()
        self.trigger_validation()
        
    def check_semantic_mismatch(self):
        if not self.ParentSheet or not self.ParentSheet.SheetName:
            if self.ValidationWarning == "Warning: View Level does not match Sheet Intent":
                self.ValidationWarning = ""
            return
        
        sheet_intent = self.ParentSheet.SheetName.lower()
        if self._level_name and self._level_name.lower() not in sheet_intent and "level" in sheet_intent:
            self.ValidationWarning = "Warning: View Level does not match Sheet Intent"
        elif self.ValidationWarning == "Warning: View Level does not match Sheet Intent":
            self.ValidationWarning = ""

    @property
    def ValidationWarning(self): return self._validation_warning
    @ValidationWarning.setter
    def ValidationWarning(self, val):
        self._validation_warning = val
        self.OnPropertyChanged("ValidationWarning")
        
    @property
    def Name(self): return self._name
    @Name.setter
    def Name(self, val):
        self._name = val
        self.OnPropertyChanged("Name")
        self.trigger_validation()
        
    @property
    def ViewNumber(self): return self._view_number
    @ViewNumber.setter
    def ViewNumber(self, val):
        v_num_str = str(val).strip() if val is not None else ""
        if v_num_str.lower() in ["nan", "none", "null"]: v_num_str = ""
        self._view_number = v_num_str
        self.OnPropertyChanged("ViewNumber")
        self.trigger_validation()
        
    @property
    def ViewType(self): return self._view_type
    @ViewType.setter
    def ViewType(self, val):
        self._view_type = val
        self.OnPropertyChanged("ViewType")
        self.OnPropertyChanged("PlanType")
        self.OnPropertyChanged("IsDetailNumberEditable")
        self.trigger_validation()

    @property
    def IsDetailNumberEditable(self):
        return self._view_type not in ["Legend", "Schedule", "PanelSchedule"]
        
    @property
    def PlanType(self): return self._view_type
    @PlanType.setter
    def PlanType(self, val):
        self.ViewType = val

    @property
    def Scale(self): return self._scale
    @Scale.setter
    def Scale(self, val):
        self._scale = val
        self.OnPropertyChanged("Scale")
        self.trigger_validation()
        
    @property
    def SourceViewId(self): return self._source_view_id
    @SourceViewId.setter
    def SourceViewId(self, val):
        self._source_view_id = val
        self.OnPropertyChanged("SourceViewId")
        self.OnPropertyChanged("IsCreateNewMode")

    @property
    def IsNewView(self): return self._is_new

    @property
    def IsCreateNewMode(self): 
        return self._is_new and self._source_view_id == ElementId.InvalidElementId
class SheetViewModel(ViewModelBase):
    def __init__(self, element_id, number, name, collection_name, discipline_name="Unknown", content_group_name="Uncategorized", series_name="Unknown", is_template=False, validation_callback=None, number_changed_callback=None, move_up_callback=None, move_down_callback=None, context_callback=None):
        ViewModelBase.__init__(self)
        self.IsDirty = False
        self.context_callback = context_callback
        self.validation_callback = validation_callback
        self.number_changed_callback = number_changed_callback
        self.move_up_callback = move_up_callback
        self.move_down_callback = move_down_callback
        self.ElementId = element_id
        self.IsTemplate = is_template
        self._sheet_number = number
        self._sheet_name = name
        self._collection_name = collection_name
        self._discipline_name = discipline_name
        self._content_group_name = content_group_name
        self.SheetSeries = series_name
        self.OriginalNumber = number
        self.OriginalName = name
        self.OriginalCollectionName = collection_name
        
        self.NumberDiff = DiffNode(number, action_callback=self.update_action)
        self.NameDiff = DiffNode(name, action_callback=self.update_action)
        
        self.Views = ObservableCollection[ViewViewModel]()
        self.AvailableNames = ObservableCollection[str]()
        
        self._is_checked = False
        self._is_name_unique = True
        self._validation_warning = ""
        self._validation_brush = "Transparent"
        self._action = "MATCHED" if not is_template else "CREATE"
        self._is_expanded = False
        
        self.PurgeCommand = RelayCommand(self.mark_purge)
        self.AddViewCommand = RelayCommand(self.add_view)
        self.UndoCommand = RelayCommand(self.undo_changes)
        self.MoveUpCommand = RelayCommand(self.on_move_up)
        self.MoveDownCommand = RelayCommand(self.on_move_down)
        
        self.FormatUpperCommand = RelayCommand(lambda: self.format_view_casing("UPPER"))
        self.FormatLowerCommand = RelayCommand(lambda: self.format_view_casing("lower"))
        self.FormatTitleCommand = RelayCommand(lambda: self.format_view_casing("Title"))
        self.SyncNamesCommand = RelayCommand(self.sync_view_names)
        
        self.populate_available_names()
        
    def format_view_casing(self, case_type):
        for v in self.Views:
            if case_type == "UPPER":
                v.Name = v.Name.upper()
            elif case_type == "lower":
                v.Name = v.Name.lower()
            elif case_type == "Title":
                v.Name = v.Name.title()
                
    def sync_view_names(self):
        intent = self.SheetName
        if intent.endswith("s") and not intent.endswith("ss"):
            intent = intent[:-1]
        for v in self.Views:
            if intent.lower() not in v.Name.lower():
                v.Name = v.Name + " - " + intent
        
    def on_move_up(self, parameter=None):
        if self.move_up_callback:
            self.move_up_callback(self)
            
    def on_move_down(self, parameter=None):
        if self.move_down_callback:
            self.move_down_callback(self)
            
    def move_view_up(self, view_vm):
        idx = self.Views.IndexOf(view_vm)
        if idx > 0:
            self.Views.Move(idx, idx - 1)
            self.reassign_view_numbers()
            
    def move_view_down(self, view_vm):
        idx = self.Views.IndexOf(view_vm)
        if idx >= 0 and idx < self.Views.Count - 1:
            self.Views.Move(idx, idx + 1)
            self.reassign_view_numbers()
            
    def reassign_view_numbers(self):
        num = 1
        for v in self.Views:
            if v.IsDetailNumberEditable:
                v.ViewNumber = str(num)
                num += 1
        
    def populate_available_names(self):
        import classification
        self.AvailableNames.Clear()
        
        base_name = self.SheetName
        if base_name.endswith("s") and not base_name.endswith("ss"):
            base_name = base_name[:-1]
            
        self.AvailableNames.Add(base_name)
        
        schemes = classification.NAMING_SCHEMES
        for scheme_list in schemes.values():
            for mod in scheme_list:
                self.AvailableNames.Add("{} - {}".format(base_name, mod))

    @property
    def SheetNumber(self): return self._sheet_number
    @SheetNumber.setter
    def SheetNumber(self, val):
        if val is None: val = ""
        if self._sheet_number == str(val): return
        
        old_val = self._sheet_number
        self._sheet_number = str(val)
        self.NumberDiff.ProposedValue = self._sheet_number
        self.IsDirty = True
        self.OnPropertyChanged("SheetNumber")
        self.reclassify()
        self.update_action()
        if hasattr(self, 'number_changed_callback') and self.number_changed_callback:
            self.number_changed_callback(self, old_val, self._sheet_number)

    @property
    def SheetName(self): return self._sheet_name
    @SheetName.setter
    def SheetName(self, val):
        if val is None: val = ""
        self._sheet_name = str(val)
        self.NameDiff.ProposedValue = self._sheet_name
        self.IsDirty = True
        self.OnPropertyChanged("SheetName")
        self.reclassify()
        self.update_action()

    @property
    def CollectionName(self): return self._collection_name
    @CollectionName.setter
    def CollectionName(self, val):
        self._collection_name = val
        self.IsDirty = True
        self.OnPropertyChanged("CollectionName")
        self.OnPropertyChanged("CollectionAndSeries")
        if self.validation_callback: self.validation_callback()

    @property
    def CollectionAndSeries(self):
        c = self.CollectionName
        s = getattr(self, "SheetSeries", "Unknown")
        return "{} - {}".format(c, s)

    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = val
        self.IsDirty = True
        self.OnPropertyChanged("IsChecked")
        if hasattr(self, 'validation_callback') and self.validation_callback:
            self.validation_callback()
        
    @property
    def IsNameUnique(self): return self._is_name_unique
    @IsNameUnique.setter
    def IsNameUnique(self, val):
        self._is_name_unique = val
        self.OnPropertyChanged("IsNameUnique")

    @property
    def ValidationWarning(self): return self._validation_warning
    @ValidationWarning.setter
    def ValidationWarning(self, val):
        self._validation_warning = val
        self.OnPropertyChanged("ValidationWarning")

    @property
    def ValidationBrush(self): return self._validation_brush
    @ValidationBrush.setter
    def ValidationBrush(self, val):
        self._validation_brush = val
        self.OnPropertyChanged("ValidationBrush")

    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, val):
        self._is_expanded = val
        self.OnPropertyChanged("IsExpanded")

    @property
    def Action(self): return self._action
    @Action.setter
    def Action(self, val):
        self._action = val
        self.OnPropertyChanged("Action")
        self.OnPropertyChanged("ActionIcon")
        self.OnPropertyChanged("ActionBrush")
        
    @property
    def ActionIcon(self):
        if self._action == "MATCHED": return u"✔"
        elif self._action in ["CREATE", "MISSING"]: return u"➕"
        elif "RENAME" in self._action: return u"✎"
        elif self._action in ["UNRECONCILED", "EXTRA"]: return u"⚠"
        elif self._action == "PURGE": return u"✖"
        return u"·"
        
    @property
    def ActionBrush(self):
        if self._action == "MATCHED": return "#10B981"
        elif self._action in ["CREATE", "MISSING"]: return "#3B82F6"
        elif "RENAME" in self._action: return "#F59E0B"
        elif self._action in ["UNRECONCILED", "EXTRA"]: return "#EF4444"
        elif self._action == "PURGE": return "#6B7280"
        return "#9CA3AF"

    @property
    def DisciplineName(self): return self._discipline_name
    @DisciplineName.setter
    def DisciplineName(self, val):
        self._discipline_name = val
        self.OnPropertyChanged("DisciplineName")
        if self.validation_callback: self.validation_callback()

    @property
    def ContentGroupName(self): return self._content_group_name
    @ContentGroupName.setter
    def ContentGroupName(self, val):
        self._content_group_name = val
        self.OnPropertyChanged("ContentGroupName")
        if self.validation_callback: self.validation_callback()

    def mark_purge(self, parameter=None):
        self.Action = "PURGE"
        self.IsChecked = True
        
    @property
    def ToggleText(self):
        return "Redo" if getattr(self, '_is_reverted_to_original', False) else "Undo"
        
    def undo_changes(self, parameter=None): # Keeping name as undo_changes for existing bindings if any, but acting as toggle
        if not hasattr(self, '_is_reverted_to_original'):
            self._is_reverted_to_original = False
            self.ProposedNumber = self.SheetNumber
            self.ProposedName = self.SheetName
            
        if self._is_reverted_to_original:
            self.SheetNumber = getattr(self, 'ProposedNumber', self.SheetNumber)
            self.SheetName = getattr(self, 'ProposedName', self.SheetName)
            self._is_reverted_to_original = False
        else:
            self.ProposedNumber = self.SheetNumber
            self.ProposedName = self.SheetName
            self.SheetNumber = self.OriginalNumber
            self.SheetName = self.OriginalName
            self._is_reverted_to_original = True
            
        self.OnPropertyChanged("ToggleText")
        if self.validation_callback: self.validation_callback()
        
    def add_view(self, param=None):
        from Autodesk.Revit.DB import ElementId
        new_v = ViewViewModel(ElementId.InvalidElementId, "New View", is_new=True, parent_sheet=self)
        self.Views.Add(new_v)
        self.reassign_view_numbers()
        if hasattr(self, 'validation_callback') and self.validation_callback:
            self.validation_callback()

    def update_action(self):
        if getattr(self, '_is_purged', False): return
        
        has_changes = (self.SheetName != self.OriginalName) or (self.SheetNumber != self.OriginalNumber)
        # Determine RENAME status
        if has_changes:
            if self.SheetName != self.OriginalName and self.SheetNumber != self.OriginalNumber:
                self.MatchStatus = "RENAME_BOTH"
            elif self.SheetName != self.OriginalName:
                self.MatchStatus = "RENAME_NAME"
            else:
                self.MatchStatus = "RENAME_NUMBER"
        else:
            self.MatchStatus = "MATCHED"
            
        # Check if views have changes
        views_changed = False
        for v in self.Views:
            if v.IsNew or getattr(v, 'Name', '') != getattr(v, 'OriginalName', '') or getattr(v, 'ViewNumber', '') != getattr(v, 'OriginalViewNumber', ''):
                views_changed = True
                break
            
        if self.IsTemplate:
            self.Action = "CREATE"
        elif self.MatchStatus != "MATCHED" or views_changed:
            self.Action = "UPDATE"
        else:
            self.Action = "MATCHED"
            
        # View validation
        nums = set()
        for v in self.Views:
            if v.ViewNumber and v.IsDetailNumberEditable:
                if v.ViewNumber in nums:
                    if not v.ValidationWarning:
                        v.ValidationWarning = "Duplicate Detail Number"
                else:
                    nums.add(v.ViewNumber)
                    if v.ValidationWarning == "Duplicate Detail Number":
                        v.ValidationWarning = ""
            v.check_semantic_mismatch()
            
        if hasattr(self, 'validation_callback') and self.validation_callback:
            self.validation_callback()

    def reclassify(self):
        try:
            import classification
            import re
            
            series_map = {
                "0": "General", "1": "Plans", "2": "Elevations", "3": "Sections",
                "4": "Large Scale Views", "5": "Details", "6": "Schedules", "7": "Diagrams", "8": "User Defined", "9": "ThreeD"
            }
            
            series_num = "0"
            disc_code = "Unknown"
            
            match = re.match(r"^([A-Z]+)[- ]?(\d)(\d\d)([A-Za-z]?)(.*)", self._sheet_number)
            if match:
                disc_code = match.group(1)
                series_num = match.group(2)
                
            c_info = classification.classify_sheet(self._sheet_number, self._sheet_name)
            
            if c_info:
                if c_info.get("drawingTypeCode", "99") != "99":
                    dtc = c_info["drawingTypeCode"]
                    if len(dtc) > 0:
                        series_num = dtc[0]
                if c_info.get("contentGroup") and c_info.get("contentGroup") != "Uncategorized":
                    self.ContentGroupName = c_info["contentGroup"]
                if c_info.get("discipline") and c_info.get("discipline") != "Unknown":
                    self.DisciplineName = c_info["discipline"]
            
            series_name_str = series_map.get(series_num, series_num)
            new_series = "0{}. {}".format(series_num, series_name_str)
            
            if getattr(self, "SheetSeries", "") != new_series:
                self.SheetSeries = new_series
                self.OnPropertyChanged("SheetSeries")
                self.OnPropertyChanged("CollectionAndSeries")
        except:
            pass

    @property
    def IsUnknownSeries(self):
        return getattr(self, "SheetSeries", "Unknown") == "Unknown"
        
    @property
    def AvailableSeriesOptions(self):
        if not hasattr(self, "_available_series_options"):
            self._available_series_options = [
                "00. General",
                "01. Plans",
                "02. Elevations",
                "03. Sections",
                "04. Large Scale Views",
                "05. Details",
                "06. Schedules",
                "07. Diagrams",
                "08. User Defined",
                "09. ThreeD"
            ]
        return self._available_series_options
        
    @property
    def OverrideSeriesText(self):
        val = getattr(self, "_override_series_text", None)
        if val is not None:
            return val
            
        current_series = getattr(self, "SheetSeries", "Unknown")
        if current_series == "Unknown":
            return None
            
        for opt in self.AvailableSeriesOptions:
            parts = opt.split('.')
            if len(parts) > 0 and parts[0].isdigit():
                if str(int(parts[0])) == str(current_series):
                    return opt
        return None
        
    @OverrideSeriesText.setter
    def OverrideSeriesText(self, value):
        self._override_series_text = value
        if not value: return
        # Extract the series code from the string (e.g. "01. Plans" -> "1")
        parts = value.split('.')
        if len(parts) > 0 and parts[0].isdigit():
            code_str = str(int(parts[0]))
            self._auto_resolve_series(code_str)
        self.OnPropertyChanged("OverrideSeriesText")
        
    def _auto_resolve_series(self, series_code):
        import re
        import classification
        
        # 1. Grab Discipline
        disc_code = "A"
        if getattr(self, 'DisciplineName', "Unknown") != "Unknown":
            # If DisciplineName is set, we could map it back to code, but regex on number is safer
            pass
            
        match = re.match(r"^([A-Z]+)[- ]?(\d)", self._sheet_number)
        if match: 
            disc_code = match.group(1)
        else:
            # Fallback if no valid prefix, look at other sheets
            all_sheets = []
            if hasattr(self, 'context_callback') and self.context_callback:
                ctx = self.context_callback()
                all_sheets = ctx.get("all_sheets", [])
                
            for s in all_sheets:
                m = re.match(r"^([A-Z]+)[- ]?(\d)", getattr(s, "SheetNumber", ""))
                if m:
                    disc_code = m.group(1)
                    break

        append_sequence = False
        all_sheets = []
        if hasattr(self, 'context_callback') and self.context_callback:
            ctx = self.context_callback()
            append_sequence = ctx.get("append_sequence", False)
            all_sheets = ctx.get("all_sheets", [])
            
        # Determine global pad format by sniffing all sheets
        pad_fmt = "{:02d}"
        if ctx and ctx.get("is_100_based", False):
            pad_fmt = "{:03d}"
        else:
            for s in all_sheets:
                n = getattr(s, "SheetNumber", "")
                # If any sheet has 3 digits after the series digit, it's 100-based
                if re.match(r"^([A-Z]+)[- ]?\d(\d\d\d)", n):
                    pad_fmt = "{:03d}"
                    break

        # 2. Find next available number
        prefix = "{}{}".format(disc_code, series_code)
        used_nums = []
        for s in all_sheets:
            n = getattr(s, "SheetNumber", "")
            if pad_fmt == "{:03d}":
                m = re.match(r"^([A-Z]+)[- ]?(\d)(\d\d\d)", n)
            else:
                m = re.match(r"^([A-Z]+)[- ]?(\d)(\d\d)", n)
                
            if m and m.group(1) == disc_code and m.group(2) == series_code:
                used_nums.append(int(m.group(3)))
                
        next_num = 1
        if used_nums:
            if append_sequence:
                next_num = max(used_nums) + 1
            else:
                for i in range(1, 999):
                    if i not in used_nums:
                        next_num = i
                        break
                        
        # Support A-101 or A101 based on existing format
        has_dash = "-" in getattr(self, "SheetNumber", "")
        dash_str = "-" if has_dash else ""
        proposed_number = "{}{}{}{}".format(disc_code, dash_str, series_code, pad_fmt.format(next_num))
        
        # 3. Find closest name via classification engine
        proposed_name = self.SheetName
        if len(self.Views) > 0:
            proposed_name = self.Views[0].Name
            
        c_info = classification.classify_sheet(proposed_number, proposed_name)
        if c_info and c_info.get("originalName") != proposed_name:
            # Ensure the matched name actually belongs to the chosen series code if possible
            if c_info.get("drawingTypeCode", "99")[0] == series_code:
                proposed_name = c_info["originalName"]
        
        self.SheetName = proposed_name
        self.SheetNumber = proposed_number

class NavTreeNode(ViewModelBase):
    def __init__(self, name, node_type, tag=None, parent=None):
        ViewModelBase.__init__(self)
        self.Name = name
        self.NodeType = node_type
        self.Tag = tag
        self.Parent = parent
        self.callback = None
        self.Children = ObservableCollection[NavTreeNode]()
        self._is_expanded = True
        self._is_selected = False
        self._is_checked = False
        self._is_target_included = True
        self._is_visible = True
        self._is_enabled = True
        self._count = 0
        self.target_callback = None
        self._is_updating = False
        self._grid_overrides = None
        self._is_segmentable = False
        self._is_active_context = False
        
        # Grid Configuration command
        self.ConfigureGridsCommand = RelayCommand(self.on_configure_grids)
        
    @property
    def GridOverrides(self):
        return self._grid_overrides
    @GridOverrides.setter
    def GridOverrides(self, val):
        self._grid_overrides = val
        self.OnPropertyChanged("GridOverrides")
        self.OnPropertyChanged("HasGridOverrides")
        
    @property
    def HasGridOverrides(self):
        return self._grid_overrides is not None
        
    def on_configure_grids(self):
        if hasattr(self, "grid_config_callback") and self.grid_config_callback:
            self.grid_config_callback(self)
            
    @property
    def HasGridConfig(self):
        return self.NodeType == "Modifier"
        
    @property
    def IsActiveContext(self):
        return self._is_active_context
        
    @IsActiveContext.setter
    def IsActiveContext(self, value):
        self._is_active_context = value
        self.OnPropertyChanged("IsActiveContext")
        self.OnPropertyChanged("FontWeight")

    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, val):
        self._is_expanded = val
        self.OnPropertyChanged("IsExpanded")

    @property
    def IsThreeState(self):
        return len(self.Children) > 0


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
        if self._is_checked == val: return
        self._is_checked = val
        self.OnPropertyChanged("IsChecked")
        
        if not self._is_updating:
            self._is_updating = True
            
            # Cascade down: If this node is checked/unchecked by the user, force all enabled children to match.
            # (If val is None, we don't force children to None, the user can't click to set None anyway in 2-state mode, 
            # but if they can in 3-state, we set children to False)
            if val is not None:
                for child in self.Children:
                    if child.IsEnabled:
                        child._set_checked_from_parent(val)
            elif val is None and len(self.Children) > 0:
                # If set to indeterminate manually (rare), we clear children
                for child in self.Children:
                    if child.IsEnabled:
                        child._set_checked_from_parent(False)
                        
            # Cascade up: Tell parent to re-evaluate its state
            if self.Parent:
                self.Parent._evaluate_checked_state()
                
            self._is_updating = False

        if self.callback: self.callback()

    def _set_checked_from_parent(self, val):
        """Helper to set checked state from a parent without triggering upward cascade"""
        if self._is_checked == val: return
        self._is_checked = val
        self.OnPropertyChanged("IsChecked")
        
        self._is_updating = True
        for child in self.Children:
            if child.IsEnabled:
                child._set_checked_from_parent(val)
        self._is_updating = False
        
        if self.callback: self.callback()
        
    def _evaluate_checked_state(self):
        """Helper to evaluate this node's state based on its children's states"""
        if not self.Children: return
        
        has_checked = False
        has_unchecked = False
        has_indeterminate = False
        
        # Only evaluate based on enabled children (or all if none are enabled?)
        enabled_children = [c for c in self.Children if c.IsEnabled]
        if not enabled_children:
            enabled_children = self.Children # fallback if all disabled
            
        for child in enabled_children:
            if child.IsChecked is True:
                has_checked = True
            elif child.IsChecked is False:
                has_unchecked = True
            else:
                has_indeterminate = True
                
        new_val = False
        if has_checked and not has_unchecked and not has_indeterminate:
            new_val = True
        elif has_unchecked and not has_checked and not has_indeterminate:
            new_val = False
        elif has_checked or has_indeterminate:
            new_val = None
            
        if self._is_checked != new_val:
            self._is_checked = new_val
            self.OnPropertyChanged("IsChecked")
            if self.Parent:
                self.Parent._evaluate_checked_state()
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
    def FontWeight(self):
        if self.NodeType == "Root": return "Bold"
        if self.NodeType == "Collection": return "Bold" if self.IsActiveContext else "Normal"
        return "Normal"
    
    @property
    def IsTargetIncluded(self): return self._is_target_included
    @IsTargetIncluded.setter
    def IsTargetIncluded(self, val):
        self._is_target_included = val
        self.OnPropertyChanged("IsTargetIncluded")
        self.OnPropertyChanged("TargetOpacity")
        if self.target_callback: self.target_callback(self)
        
    @property
    def IsVisible(self): return self._is_visible
    @IsVisible.setter
    def IsVisible(self, val):
        if self._is_visible == val: return
        self._is_visible = val
        self.OnPropertyChanged("IsVisible")
        
    @property
    def IsEnabled(self): return self._is_enabled
    @IsEnabled.setter
    def IsEnabled(self, val):
        if self._is_enabled == val: return
        self._is_enabled = val
        self.OnPropertyChanged("IsEnabled")
        if not val and self.IsChecked:
            self.IsChecked = False # auto-uncheck if disabled
        
    @property
    def TargetOpacity(self): return 1.0 if self.IsTargetIncluded else 0.5

class ManageSheetsViewModel(ViewModelBase):
    def __init__(self):
        ViewModelBase.__init__(self)
        self._editor_items = ObservableCollection[object]()
        self._nav_root = ObservableCollection[NavTreeNode]()
        
        self._is_push_enabled = False
        self._is_reviewer_ready = False
        self._has_valid_work = False

    @property
    def EditorItems(self): return self._editor_items
    
    @property
    def NavRoot(self): return self._nav_root

    @property
    def IsPushEnabled(self): return self._is_push_enabled
    @IsPushEnabled.setter
    def IsPushEnabled(self, value):
        self._is_push_enabled = value
        self.OnPropertyChanged("IsPushEnabled")
        
    @property
    def IsReviewerReady(self): return self._is_reviewer_ready
    @IsReviewerReady.setter
    def IsReviewerReady(self, value):
        self._is_reviewer_ready = value
        self.OnPropertyChanged("IsReviewerReady")
