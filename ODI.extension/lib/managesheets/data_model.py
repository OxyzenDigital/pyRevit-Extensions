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

AVAILABLE_LEVELS = []
AVAILABLE_VIEW_FAMILY_TYPES = []
AVAILABLE_EXISTING_VIEWS = []
AVAILABLE_ASSIGNABLE_VIEWS = []

SCALE_MAP_STR_TO_INT = {
    "12\" = 1'-0\"": 1,
    "6\" = 1'-0\"": 2,
    "3\" = 1'-0\"": 4,
    "1 1/2\" = 1'-0\"": 8,
    "1\" = 1'-0\"": 12,
    "3/4\" = 1'-0\"": 16,
    "1/2\" = 1'-0\"": 24,
    "3/8\" = 1'-0\"": 32,
    "1/4\" = 1'-0\"": 48,
    "3/16\" = 1'-0\"": 64,
    "1/8\" = 1'-0\"": 96,
    "1\" = 10'-0\"": 120,
    "3/32\" = 1'-0\"": 128,
    "1/16\" = 1'-0\"": 192,
    "1\" = 20'-0\"": 240,
    "3/64\" = 1'-0\"": 256,
    "1\" = 30'-0\"": 360,
    "1/32\" = 1'-0\"": 384,
    "1\" = 40'-0\"": 480,
    "1\" = 50'-0\"": 600,
    "1\" = 60'-0\"": 720,
    "1/64\" = 1'-0\"": 768,
    "1\" = 80'-0\"": 960,
    "1\" = 100'-0\"": 1200,
    "1\" = 160'-0\"": 1920,
    "1\" = 200'-0\"": 2400,
    "1\" = 300'-0\"": 3600,
    "1\" = 400'-0\"": 4800
}

SCALE_MAP_INT_TO_STR = dict((v, k) for k, v in SCALE_MAP_STR_TO_INT.items())

def format_revit_scale(scale_int):
    if not scale_int:
        return "1/8\" = 1'-0\""
    try:
        s_int = int(scale_int)
        if s_int in SCALE_MAP_INT_TO_STR:
            return SCALE_MAP_INT_TO_STR[s_int]
        return "1 : {}".format(s_int)
    except Exception:
        return str(scale_int)

def parse_revit_scale(scale_str):
    if not scale_str:
        return None
    s_clean = str(scale_str).strip()
    if s_clean in SCALE_MAP_STR_TO_INT:
        return SCALE_MAP_STR_TO_INT[s_clean]
    if s_clean.startswith("1 : "):
        try:
            return int(s_clean.replace("1 : ", "").strip())
        except Exception:
            return None
    return None

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
        
        self._has_diff = False
        self._prefix = original_val or ""
        self._old = ""
        self._new = ""
        self._suffix = ""
        
    @property
    def HasDiff(self): return self._has_diff
    @HasDiff.setter
    def HasDiff(self, val):
        self._has_diff = val
        self.OnPropertyChanged("HasDiff")

    @property
    def Prefix(self): return self._prefix
    @Prefix.setter
    def Prefix(self, val):
        self._prefix = val
        self.OnPropertyChanged("Prefix")

    @property
    def Old(self): return self._old
    @Old.setter
    def Old(self, val):
        self._old = val
        self.OnPropertyChanged("Old")

    @property
    def New(self): return self._new
    @New.setter
    def New(self, val):
        self._new = val
        self.OnPropertyChanged("New")

    @property
    def Suffix(self): return self._suffix
    @Suffix.setter
    def Suffix(self, val):
        self._suffix = val
        self.OnPropertyChanged("Suffix")

    @property
    def OriginalValue(self): return self._original_val
    @property
    def ProposedValue(self): return self._proposed_val
    @ProposedValue.setter
    def ProposedValue(self, val):
        self._proposed_val = val
        self._has_diff, self._prefix, self._old, self._new, self._suffix = generate_char_diff(self._original_val, val)
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
        self._name = name
        self._display_name = display_name if display_name else name
        self._is_checked = is_checked
        self._is_enabled = True
        self.callback = callback
        
    @property
    def Name(self): return self._name
    @Name.setter
    def Name(self, val):
        self._name = val
        self.OnPropertyChanged("Name")

    @property
    def DisplayName(self): return self._display_name
    @DisplayName.setter
    def DisplayName(self, val):
        self._display_name = val
        self.OnPropertyChanged("DisplayName")

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
        self._name = name
        self._children = []
    @property
    def Name(self): return self._name
    @property
    def Children(self): return self._children

class CollectionGroupNode(ViewModelBase):
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self._name = name
        self._children = []
    @property
    def Name(self): return self._name
    @property
    def Children(self): return self._children

class AssignableViewNode(ViewModelBase):
    def __init__(self, view_id, name, is_placed=False, sheet_number="", view_type="FloorPlan", level_name="", scale_int=96):
        ViewModelBase.__init__(self)
        self.Id = view_id
        self.ViewId = view_id
        self._name = name
        self.IsPlaced = is_placed
        self.SheetNumber = sheet_number
        self.ViewType = str(view_type)
        self.LevelName = level_name or ""
        self.ScaleInt = scale_int or 96
        self.ScaleStr = format_revit_scale(self.ScaleInt)
        self.StatusColor = "#9CA3AF" if is_placed else "#3B82F6"
        if is_placed and sheet_number:
            self._display_name = "{} [on {}]".format(name, sheet_number)
        else:
            lvl_tag = " ({})".format(level_name) if level_name else ""
            self._display_name = "{}{} [Unplaced]".format(name, lvl_tag)

    @property
    def Name(self): return self._name
    @property
    def ViewName(self): return self._name
    @property
    def DisplayName(self): return self._display_name

class ViewViewModel(ViewModelBase):
    def __init__(self, view_id, name, view_type="FloorPlan", scale="1/8\" = 1'-0\"", is_new=False, level_name="", parent_sheet=None, view_number="", has_graphic_scale=None, has_north_arrow=None):
        ViewModelBase.__init__(self)
        self.ViewId = view_id
        self._name = name
        self._view_type = str(view_type) if view_type else "FloorPlan"
        self._scale = scale if scale else "1/8\" = 1'-0\""
        self._is_new = is_new
        v_num_str = str(view_number).strip() if view_number is not None else ""
        if v_num_str.lower() in ["nan", "none", "null"]: v_num_str = ""
        self._view_number = v_num_str
        self._assignment_mode = "Existing" if not is_new else "New"
        self._source_view_id = ElementId.InvalidElementId
        self._validation_warning = ""
        self._scale_warning = ""
        self._level_name = level_name or ""
        self.ParentSheet = parent_sheet
        
        self._available_levels = AVAILABLE_LEVELS
        self._available_view_family_types = AVAILABLE_VIEW_FAMILY_TYPES
        self._available_existing_views = AVAILABLE_EXISTING_VIEWS
        self._selected_existing_view = None
        
        # Attempt to infer level if new view
        if self._is_new and not self._level_name and self.ParentSheet:
            self.infer_level_from_sheet()
                        
        if not self._view_type and self.AvailableViewFamilyTypes:
            self._view_type = self.AvailableViewFamilyTypes[0]
            
        plan_keywords = ["plan", "floor", "ceiling", "structural", "engineering", "area", "rcp", "framing"]
        is_plan = any(k in self._view_type.lower() for k in plan_keywords)
        is_non_scaled = any(k in self._view_type.lower() for k in ["legend", "schedule", "three", "3d", "rendering"])
        
        if has_graphic_scale is None:
            self._has_graphic_scale = not is_non_scaled
        else:
            self._has_graphic_scale = bool(has_graphic_scale)
            
        if has_north_arrow is None:
            self._has_north_arrow = is_plan
        else:
            self._has_north_arrow = bool(has_north_arrow)
            
        self.OriginalName = self._name
        self.OriginalViewNumber = self._view_number
        self.OriginalScale = self._scale
        self.OriginalHasGraphicScale = self._has_graphic_scale
        self.OriginalHasNorthArrow = self._has_north_arrow
        
        self._available_scales = [
            "12\" = 1'-0\"", "6\" = 1'-0\"", "3\" = 1'-0\"", "1 1/2\" = 1'-0\"",
            "1\" = 1'-0\"", "3/4\" = 1'-0\"", "1/2\" = 1'-0\"", "3/8\" = 1'-0\"",
            "1/4\" = 1'-0\"", "3/16\" = 1'-0\"", "1/8\" = 1'-0\"", "1\" = 10'-0\"",
            "3/32\" = 1'-0\"", "1/16\" = 1'-0\"", "1\" = 20'-0\"", "3/64\" = 1'-0\"",
            "1\" = 30'-0\"", "1/32\" = 1'-0\"", "1\" = 40'-0\"", "1\" = 50'-0\"",
            "1\" = 60'-0\"", "1/64\" = 1'-0\"", "1\" = 80'-0\"", "1\" = 100'-0\"",
            "1\" = 160'-0\"", "1\" = 200'-0\"", "1\" = 300'-0\"", "1\" = 400'-0\""
        ]
        if self._scale and self._scale not in self._available_scales:
            self._available_scales.insert(0, self._scale)
        
        self._delete_command = RelayCommand(self.delete_view)
        self._move_up_command = RelayCommand(self.move_up)
        self._move_down_command = RelayCommand(self.move_down)

    def infer_level_from_sheet(self):
        if not self.ParentSheet:
            return
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
            # Check for 01. Plans level digit in A-101, A-110, A-120, etc.
            m = re.match(r"^[A-Z]+[- ]?1(\d)\d", self.ParentSheet.SheetNumber.upper())
            if m:
                lvl_idx = int(m.group(1)) - 1
                if 0 <= lvl_idx < len(self.AvailableLevels):
                    matched_lvl = self.AvailableLevels[lvl_idx]
            if not matched_lvl:
                m2 = re.search(r'\d+', self.ParentSheet.SheetNumber)
                if m2:
                    num_str = m2.group()
                    for lvl in self.AvailableLevels:
                        if num_str in lvl:
                            matched_lvl = lvl
                            break
                        
        if matched_lvl:
            self._level_name = matched_lvl

    @property
    def AvailableScales(self): return self._available_scales
    @property
    def AvailableLevelsList(self): return self.AvailableLevels
    @property
    def AvailableViewFamilyTypesList(self): return self.AvailableViewFamilyTypes
    @property
    def AvailableExistingViewsList(self): return self.AvailableExistingViews

    @property
    def DeleteCommand(self): return self._delete_command
    @property
    def MoveUpCommand(self): return self._move_up_command
    @property
    def MoveDownCommand(self): return self._move_down_command
        
    def delete_view(self, param=None):
        if self.ParentSheet:
            self.ParentSheet.Views.Remove(self)
            self.ParentSheet.reassign_view_numbers()
            self.ParentSheet.update_action()
            
    def move_up(self, param=None):
        if self.ParentSheet:
            self.ParentSheet.move_view_up(self)
            
    def move_down(self, param=None):
        if self.ParentSheet:
            self.ParentSheet.move_view_down(self)
        
    def trigger_validation(self):
        if self.ParentSheet:
            self.ParentSheet.update_action()
            if hasattr(self.ParentSheet, "notify_reconciliation_state"):
                self.ParentSheet.notify_reconciliation_state()

    @property
    def AssignmentMode(self): return self._assignment_mode
    @AssignmentMode.setter
    def AssignmentMode(self, val):
        if not val or self._assignment_mode == val:
            return
        self._assignment_mode = val
        self._is_new = (val == "New")
        self.OnPropertyChanged("AssignmentMode")
        self.OnPropertyChanged("IsNew")
        self.OnPropertyChanged("IsExisting")
        self.OnPropertyChanged("ShowExistingViewPicker")
        self.trigger_validation()
        
    @property
    def IsNew(self): return self._is_new
    
    @property
    def IsExisting(self): return not self._is_new

    @property
    def ShowExistingViewPicker(self):
        return (not self._is_new) and (self.ViewId == ElementId.InvalidElementId)

    @property
    def SelectedExistingView(self): return self._selected_existing_view
    @SelectedExistingView.setter
    def SelectedExistingView(self, node):
        self._selected_existing_view = node
        if node and hasattr(node, "Id"):
            self.ViewId = node.Id
            self._name = node.Name
            self._view_type = node.ViewType
            self._level_name = node.LevelName
            self._scale = node.ScaleStr
            self.OnPropertyChanged("Name")
            self.OnPropertyChanged("ViewType")
            self.OnPropertyChanged("PlanType")
            self.OnPropertyChanged("LevelName")
            self.OnPropertyChanged("Scale")
            self.OnPropertyChanged("IsPlanView")
            self.OnPropertyChanged("ShowExistingViewPicker")
            self.check_semantic_mismatch()
            self.trigger_validation()
        self.OnPropertyChanged("SelectedExistingView")

    @property
    def AvailableLevels(self): return self._available_levels
    @AvailableLevels.setter
    def AvailableLevels(self, val):
        self._available_levels = val
        self.OnPropertyChanged("AvailableLevels")

    @property
    def AvailableViewFamilyTypes(self): return self._available_view_family_types
    @AvailableViewFamilyTypes.setter
    def AvailableViewFamilyTypes(self, val):
        self._available_view_family_types = val
        self.OnPropertyChanged("AvailableViewFamilyTypes")

    @property
    def AvailableExistingViews(self):
        return self._available_existing_views or AVAILABLE_EXISTING_VIEWS or AVAILABLE_ASSIGNABLE_VIEWS
    @AvailableExistingViews.setter
    def AvailableExistingViews(self, val):
        self._available_existing_views = val
        self.OnPropertyChanged("AvailableExistingViews")

    @property
    def LevelName(self): return self._level_name
    @LevelName.setter
    def LevelName(self, val):
        self._level_name = val or ""
        self.OnPropertyChanged("LevelName")
        self.check_semantic_mismatch()
        self.trigger_validation()
        
    def check_semantic_mismatch(self):
        warnings = []
        if self.ParentSheet and self.ParentSheet.SheetName:
            sheet_intent = self.ParentSheet.SheetName.lower()
            sheet_series = getattr(self.ParentSheet, "SheetSeries", "").lower()
            v_type_lower = (self._view_type or "").lower()
            
            # 1. Level mismatch check
            if self._level_name:
                lvl_low = self._level_name.lower()
                # Check if sheet explicitly mentions another level from AvailableLevels
                other_level_mentioned = None
                for known_lvl in self.AvailableLevels:
                    k_low = known_lvl.lower()
                    if k_low != lvl_low and k_low in sheet_intent:
                        other_level_mentioned = known_lvl
                        break
                if other_level_mentioned and lvl_low not in sheet_intent:
                    warnings.append("Wrong Level: View is on '{}' but Sheet is for '{}'".format(self._level_name, other_level_mentioned))
                elif "roof" in sheet_intent and "roof" not in lvl_low and ("level" in lvl_low or "floor" in lvl_low):
                    warnings.append("Wrong Level: '{}' placed on a Roof sheet".format(self._level_name))

            # 2. ViewType vs Sheet Intent / Series check
            if "ceiling" in sheet_intent or "rcp" in sheet_intent:
                if "floorplan" in v_type_lower or ("floor" in v_type_lower and "ceiling" not in v_type_lower):
                    warnings.append("Wrong View Type: Floor Plan placed on a Reflected Ceiling Plan sheet")
            elif ("floor plan" in sheet_intent or "dimension plan" in sheet_intent or "partition plan" in sheet_intent or "furniture plan" in sheet_intent) and "ceiling" not in sheet_intent:
                if "ceiling" in v_type_lower or "rcp" in v_type_lower:
                    warnings.append("Wrong View Type: Ceiling Plan placed on a Floor Plan sheet")
                    
            if "01. plans" in sheet_series or "plan" in sheet_intent:
                if "elevation" in v_type_lower or "section" in v_type_lower:
                    warnings.append("Wrong View Type: {} placed on a Plan sheet".format(self._view_type))
            elif "02. elevations" in sheet_series:
                if "plan" in v_type_lower or "section" in v_type_lower:
                    warnings.append("Wrong View Type: {} placed on an Elevations sheet".format(self._view_type))
            elif "03. sections" in sheet_series:
                if "plan" in v_type_lower or "elevation" in v_type_lower:
                    warnings.append("Wrong View Type: {} placed on a Sections sheet".format(self._view_type))

        if self._scale_warning:
            warnings.append(self._scale_warning)
            
        # Preserve duplicate detail number warning if set externally
        if self._validation_warning and "Duplicate" in self._validation_warning and not any("Duplicate" in w for w in warnings):
            warnings.insert(0, "Duplicate Detail Number on sheet.")
            
        self.ValidationWarning = " | ".join(warnings)

    @property
    def ScaleWarning(self): return self._scale_warning
    @ScaleWarning.setter
    def ScaleWarning(self, val):
        self._scale_warning = val or ""
        self.check_semantic_mismatch()

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
        self._view_type = val or ""
        self.OnPropertyChanged("ViewType")
        self.OnPropertyChanged("PlanType")
        self.OnPropertyChanged("IsDetailNumberEditable")
        self.OnPropertyChanged("IsPlanView")
        self.check_semantic_mismatch()
        self.trigger_validation()

    @property
    def IsPlanView(self):
        vt = (self._view_type or "").lower()
        return any(k in vt for k in ["plan", "floor", "ceiling", "structural", "engineering", "area", "rcp", "framing"])

    @property
    def HasGraphicScale(self): return self._has_graphic_scale
    @HasGraphicScale.setter
    def HasGraphicScale(self, val):
        self._has_graphic_scale = bool(val)
        self.OnPropertyChanged("HasGraphicScale")
        self.trigger_validation()

    @property
    def HasNorthArrow(self): return self._has_north_arrow
    @HasNorthArrow.setter
    def HasNorthArrow(self, val):
        self._has_north_arrow = bool(val)
        self.OnPropertyChanged("HasNorthArrow")
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

    def has_pending_changes(self):
        if self._is_new:
            return True
        if self._name != self.OriginalName:
            return True
        if self._view_number != self.OriginalViewNumber:
            return True
        if self._scale != self.OriginalScale:
            return True
        if self._has_graphic_scale != self.OriginalHasGraphicScale:
            return True
        if self._has_north_arrow != self.OriginalHasNorthArrow:
            return True
        return False

    def clone(self, new_parent_sheet):
        c = ViewViewModel(
            view_id=self.ViewId,
            name=self.OriginalName,
            view_type=self._view_type,
            scale=self.OriginalScale,
            is_new=self._is_new,
            level_name=self._level_name,
            parent_sheet=new_parent_sheet,
            view_number=self.OriginalViewNumber,
            has_graphic_scale=self.OriginalHasGraphicScale,
            has_north_arrow=self.OriginalHasNorthArrow
        )
        return c


class SheetViewModel(ViewModelBase):
    def __init__(self, element_id, number, name, collection_name, discipline_name="Unknown", content_group_name="Uncategorized", series_name="Unknown", is_template=False, validation_callback=None, number_changed_callback=None, move_up_callback=None, move_down_callback=None, context_callback=None, clone_sheet_callback=None, auto_populate_callback=None, select_sheet_callback=None, accept_suggestion_callback=None):
        ViewModelBase.__init__(self)
        self.IsDirty = False
        self.context_callback = context_callback
        self.validation_callback = validation_callback
        self.number_changed_callback = number_changed_callback
        self.move_up_callback = move_up_callback
        self.move_down_callback = move_down_callback
        self.clone_sheet_callback = clone_sheet_callback
        self.auto_populate_callback = auto_populate_callback
        self.select_sheet_callback = select_sheet_callback
        self.accept_suggestion_callback = accept_suggestion_callback
        self.ElementId = element_id
        self.IsTemplate = is_template
        self._sheet_number = number
        self._sheet_name = name
        self._collection_name = collection_name
        self._discipline_name = discipline_name
        self._content_group_name = content_group_name
        self._sheet_series = series_name if series_name and series_name != "Unknown" else "Unknown"
        self._original_number = number
        self._original_name = name
        self.OriginalCollectionName = collection_name
        
        self._number_diff = DiffNode(number, action_callback=self.update_action)
        self._name_diff = DiffNode(name, action_callback=self.update_action)
        
        self._views = ObservableCollection[ViewViewModel]()
        self._available_names = ObservableCollection[str]()
        
        self._is_checked = False
        self._is_name_unique = True
        self._validation_warning = ""
        self._validation_brush = "Transparent"
        self._action = "MATCHED" if not is_template else "CREATE"
        self._is_expanded = False
        self._is_selected = False
        self._is_user_modified = False
        self._suggested_number = ""
        self._suggested_name = ""
        self._suggested_series = ""
        
        self._purge_command = RelayCommand(self.mark_purge)
        self._add_view_command = RelayCommand(self.add_view)
        self._undo_command = RelayCommand(self.undo_changes)
        self._move_up_command = RelayCommand(self.on_move_up)
        self._move_down_command = RelayCommand(self.on_move_down)
        self._clone_sheet_command = RelayCommand(self.on_clone_sheet)
        self._auto_populate_views_command = RelayCommand(self.on_auto_populate_views)
        self._select_sheet_command = RelayCommand(self.on_select_sheet)
        self._accept_suggestion_command = RelayCommand(self.on_accept_suggestion)
        
        self._format_upper_command = RelayCommand(lambda: self.format_view_casing("UPPER"))
        self._format_lower_command = RelayCommand(lambda: self.format_view_casing("lower"))
        self._format_title_command = RelayCommand(lambda: self.format_view_casing("Title"))
        self._sync_names_command = RelayCommand(self.sync_view_names)
        self._use_view_name_command = RelayCommand(self.use_view_name_for_sheet)
        
        if self._sheet_series == "Unknown":
            self.reclassify()
        self.populate_available_names()

    def clone_for_sandbox(
        self,
        validation_callback=None,
        number_changed_callback=None,
        move_up_callback=None,
        move_down_callback=None,
        context_callback=None,
        clone_sheet_callback=None,
        auto_populate_callback=None,
        select_sheet_callback=None,
        accept_suggestion_callback=None,
        **kwargs
    ):
        cloned = SheetViewModel(
            element_id=self.ElementId,
            number=self.OriginalNumber,
            name=self.OriginalName,
            collection_name=self.OriginalCollectionName,
            discipline_name=self._discipline_name,
            content_group_name=self._content_group_name,
            series_name=self._sheet_series,
            is_template=self.IsTemplate,
            validation_callback=validation_callback or self.validation_callback,
            number_changed_callback=number_changed_callback or self.number_changed_callback,
            move_up_callback=move_up_callback or self.move_up_callback,
            move_down_callback=move_down_callback or self.move_down_callback,
            context_callback=context_callback or self.context_callback,
            clone_sheet_callback=clone_sheet_callback or self.clone_sheet_callback,
            auto_populate_callback=auto_populate_callback or self.auto_populate_callback,
            select_sheet_callback=select_sheet_callback or self.select_sheet_callback,
            accept_suggestion_callback=accept_suggestion_callback or self.accept_suggestion_callback
        )
        for v in self._views:
            cloned.Views.Add(v.clone(cloned))
        cloned.IsDirty = False
        cloned._is_user_modified = False
        return cloned

    def update_diffs(self):
        self._number_diff.ProposedValue = self._sheet_number
        self._name_diff.ProposedValue = self._sheet_name
        self.OnPropertyChanged("SheetNumber")
        self.OnPropertyChanged("SheetName")
        self.OnPropertyChanged("MatrixPillSummary")
        self.populate_available_names()

    @property
    def NumberDiff(self): return self._number_diff
    @property
    def NameDiff(self): return self._name_diff
    @property
    def Views(self): return self._views
    @property
    def AvailableNames(self): return self._available_names

    @property
    def PurgeCommand(self): return self._purge_command
    @property
    def AddViewCommand(self): return self._add_view_command
    @property
    def UndoCommand(self): return self._undo_command
    @property
    def MoveUpCommand(self): return self._move_up_command
    @property
    def MoveDownCommand(self): return self._move_down_command
    @property
    def CloneSheetCommand(self): return self._clone_sheet_command
    @property
    def AutoPopulateViewsCommand(self): return self._auto_populate_views_command
    @property
    def SelectSheetCommand(self): return self._select_sheet_command
    @property
    def AcceptSuggestionCommand(self): return self._accept_suggestion_command
    @property
    def FormatUpperCommand(self): return self._format_upper_command
    @property
    def FormatLowerCommand(self): return self._format_lower_command
    @property
    def FormatTitleCommand(self): return self._format_title_command
    @property
    def SyncNamesCommand(self): return self._sync_names_command
    @property
    def UseViewNameCommand(self): return self._use_view_name_command

    def _clean_view_name(self, raw_name):
        name_str = (raw_name or "").strip()
        if not name_str:
            return ""
        m = re.match(r"^(.*?)\s+-\s+(?:[A-Z0-9]{1,5}\s+-\s+)?\d+\s+-\s+[A-Z]{1,4}[- ]?\d+.*$", name_str)
        if m and m.group(1).strip():
            return m.group(1).strip()
        return name_str

    def notify_reconciliation_state(self):
        props = [
            "HasDiffFromOriginal", "DiffArrowText", "ChangeDiffSummary",
            "MatchedViewSummary", "HasPlacedView", "CanUseViewName",
            "CanSyncViewToSheet", "CanAcceptProposal", "IsAccepted",
            "CanRevertOrUndo", "AcceptButtonText", "StatusBadgeText",
            "Action", "ActionIcon", "ActionBrush", "DisciplineCode",
            "DisciplineShortLabel", "OriginalNameDisplay", "PlacedViewsBaselineText",
            "HasNoViews", "ViewsSummaryBadge", "MatrixPillSummary",
            "HasWarning", "WarningSummary", "HasSuggestion", "SuggestionShortText", "SuggestionBadge"
        ]
        for p in props:
            self.OnPropertyChanged(p)

    def use_view_name_for_sheet(self, parameter=None):
        if self._views.Count > 0:
            clean_vn = self._clean_view_name(self._views[0].Name)
            if clean_vn:
                self.SheetName = clean_vn
        elif self._views.Count == 0:
            self.on_auto_populate_views()
        self.notify_reconciliation_state()

    @property
    def DisciplineCode(self):
        d = (self._discipline_name or "").strip()
        if " - " in d:
            return d.split(" - ")[0].strip().upper()
        num = (self._sheet_number or "").strip()
        m = re.match(r"^([A-Za-z]+)", num)
        if m:
            return m.group(1).upper()
        return "G"

    @property
    def DisciplineShortLabel(self):
        disc_map = {
            "A": u"A · Arch",
            "S": u"S · Struct",
            "M": u"M · Mech",
            "E": u"E · Elec",
            "P": u"P · Plumb",
            "C": u"C · Civil",
            "L": u"L · Land",
            "F": u"F · Fire",
            "FP": u"FP · Fire",
            "G": u"G · Gen",
            "I": u"I · Int",
            "CS": u"CS · Cover"
        }
        code = self.DisciplineCode
        if code in disc_map:
            return disc_map[code]
        if code:
            return u"{} · Disc".format(code)
        return u"G · Gen"

    @property
    def OriginalNameDisplay(self):
        return self._original_name if not self.IsTemplate else u"— (New Sheet)"

    @property
    def PlacedViewsBaselineText(self):
        if self._views.Count == 0:
            return u"No placed views"
        v0 = self._views[0]
        v_name = self._clean_view_name(v0.OriginalName or v0.Name)
        v_scale = v0.OriginalScale or v0.Scale or ""
        base_str = u"{} ({})".format(v_name, v_scale) if v_scale else v_name
        if self._views.Count > 1:
            return base_str + u" +{}".format(self._views.Count - 1)
        return base_str

    @property
    def HasDiffFromOriginal(self):
        if self.IsTemplate or self._action == "PURGE":
            return True
        if (self._sheet_number != self._original_number) or (self._sheet_name != self._original_name):
            return True
        return any(v.has_pending_changes() for v in self._views)

    @property
    def DiffArrowText(self):
        return u"→" if self.HasDiffFromOriginal else u"="

    @property
    def ChangeDiffSummary(self):
        if self._action == "PURGE":
            return u"Marked for PURGE quarantine"
        elif self.IsTemplate:
            return u"New sheet: {} \"{}\"".format(self._sheet_number, self._sheet_name)
        elif not self.HasDiffFromOriginal:
            return u"No change from Revit"
        parts = []
        if self._sheet_number != self._original_number:
            parts.append(u"# {} → {}".format(self._original_number, self._sheet_number))
        if self._sheet_name != self._original_name:
            parts.append(u"Name: \"{}\" → \"{}\"".format(self._original_name, self._sheet_name))
        if any(v.has_pending_changes() for v in self._views):
            parts.append(u"View/scale updated")
        return u"  |  ".join(parts) if parts else u"No change from Revit"

    @property
    def MatchedViewSummary(self):
        if self._views.Count == 0:
            return u"0 Views — Click ⚡ Match View"
        v0 = self._views[0]
        v_cur = self._clean_view_name(v0.Name)
        v_orig = self._clean_view_name(v0.OriginalName)
        if v0.IsNew:
            res = u"Proposed New View: \"{}\" ({})".format(v_cur, v0.Scale)
        elif v_cur != v_orig:
            res = u"View Title: \"{}\" → \"{}\" ({})".format(v_orig, v_cur, v0.Scale)
        else:
            res = u"View: \"{}\" ({})".format(v_cur, v0.Scale)
        if self._views.Count > 1:
            res += u" +{} more".format(self._views.Count - 1)
        return res

    @property
    def HasPlacedView(self):
        return self._views.Count > 0

    @property
    def CanUseViewName(self):
        if self._views.Count == 0:
            return False
        clean_v0 = self._clean_view_name(self._views[0].Name).upper()
        sheet_up = (self._sheet_name or "").strip().upper()
        return clean_v0 != sheet_up

    @property
    def CanSyncViewToSheet(self):
        if self._views.Count == 0:
            return False
        v0_name = (self._views[0].Name or "").strip().upper()
        sheet_up = (self._sheet_name or "").strip().upper()
        return v0_name != sheet_up

    @property
    def CanAcceptProposal(self):
        return (not self._is_checked) and self.HasDiffFromOriginal and (self._action != "PURGE")

    @property
    def IsAccepted(self):
        return bool(self._is_checked) and (self.HasDiffFromOriginal or self.IsTemplate or self._action == "PURGE")

    @property
    def CanRevertOrUndo(self):
        return (not self.IsTemplate) and (self.HasDiffFromOriginal or self._is_checked)

    @property
    def AcceptButtonText(self):
        return u"✓ Accept Modified" if getattr(self, "_is_user_modified", False) else u"✓ Accept"

    @property
    def StatusBadgeText(self):
        if self._action == "PURGE":
            return u"PURGE"
        elif self.IsTemplate:
            return u"NEW"
        elif self._is_checked and self.HasDiffFromOriginal:
            return u"ACCEPTED"
        elif self.HasDiffFromOriginal:
            return u"MODIFIED" if getattr(self, "_is_user_modified", False) else u"PROPOSAL"
        else:
            return u"AS-IS"

    @property
    def IsSelected(self): return self._is_selected
    @IsSelected.setter
    def IsSelected(self, val):
        self._is_selected = bool(val)
        self.OnPropertyChanged("IsSelected")

    @property
    def OriginalNumber(self):
        return self._original_number

    @OriginalNumber.setter
    def OriginalNumber(self, val):
        self._original_number = val or ""
        self.OnPropertyChanged("OriginalNumber")
        self.OnPropertyChanged("OriginalNumberDisplay")

    @property
    def OriginalName(self):
        return self._original_name

    @OriginalName.setter
    def OriginalName(self, val):
        self._original_name = val or ""
        self.OnPropertyChanged("OriginalName")

    @property
    def OriginalNumberDisplay(self):
        return self._original_number if not self.IsTemplate else "NEW"

    @property
    def HasSuggestion(self):
        if self.IsTemplate:
            return False
        if self._suggested_number and self._suggested_number != self._sheet_number:
            return True
        if self._suggested_name and self._suggested_name != self._sheet_name:
            return True
        return False

    @property
    def SuggestionShortText(self):
        if self._suggested_number:
            return u"✨ {} · Accept".format(self._suggested_number)
        return u"✨ Standardize"

    @property
    def SuggestedNumber(self): return self._suggested_number
    @SuggestedNumber.setter
    def SuggestedNumber(self, val):
        self._suggested_number = val or ""
        self.OnPropertyChanged("SuggestedNumber")
        self.OnPropertyChanged("SuggestionBadge")
        self.OnPropertyChanged("HasSuggestion")
        self.OnPropertyChanged("SuggestionShortText")

    @property
    def SuggestedName(self): return self._suggested_name
    @SuggestedName.setter
    def SuggestedName(self, val):
        self._suggested_name = val or ""
        self.OnPropertyChanged("SuggestedName")
        self.OnPropertyChanged("SuggestionBadge")
        self.OnPropertyChanged("HasSuggestion")
        self.OnPropertyChanged("SuggestionShortText")

    @property
    def SuggestedSeries(self): return self._suggested_series
    @SuggestedSeries.setter
    def SuggestedSeries(self, val):
        self._suggested_series = val or ""
        self.OnPropertyChanged("SuggestedSeries")

    @property
    def SuggestionBadge(self):
        if self._suggested_number and self._suggested_name:
            return u"✨ Rule Book: {} · {}".format(self._suggested_number, self._suggested_name)
        return u"✨ Click Auto-Fix to standardize"

    @property
    def HasWarning(self):
        if self._validation_warning:
            return True
        for v in self._views:
            if v.ValidationWarning or v.ScaleWarning:
                return True
        return False

    @property
    def WarningSummary(self):
        msgs = []
        if self._validation_warning:
            msgs.append(self._validation_warning)
        for v in self._views:
            if v.ValidationWarning:
                msgs.append("{}: {}".format(v.Name, v.ValidationWarning))
            elif v.ScaleWarning:
                msgs.append("{}: {}".format(v.Name, v.ScaleWarning))
        return " | ".join(msgs)

    @property
    def MatrixPillSummary(self):
        cnt = self._views.Count
        if cnt == 0:
            return u"0 Views · Click ⚡ to populate"
        first_v = self._views[0]
        scale_str = first_v.Scale or ""
        icons = []
        if any(v.HasGraphicScale for v in self._views):
            icons.append(u"📏")
        if any(v.HasNorthArrow for v in self._views):
            icons.append(u"🧭")
        icon_suffix = (" " + "".join(icons)) if icons else ""
        if self.HasWarning:
            return u"⚠ {} View{} · {}{}".format(cnt, "s" if cnt != 1 else "", scale_str, icon_suffix)
        return u"{} View{} · {}{}".format(cnt, "s" if cnt != 1 else "", scale_str, icon_suffix)

    @property
    def HasNoViews(self):
        return self._views.Count == 0

    @property
    def ViewsSummaryBadge(self):
        cnt = self._views.Count
        if cnt == 0:
            return "0 Views"
        new_cnt = sum(1 for v in self._views if v.IsNew)
        if new_cnt > 0:
            return "{} View{} ({} New)".format(cnt, "s" if cnt != 1 else "", new_cnt)
        return "{} View{}".format(cnt, "s" if cnt != 1 else "")

    def on_select_sheet(self, parameter=None):
        if self.select_sheet_callback:
            self.select_sheet_callback(self)

    def on_accept_suggestion(self, parameter=None):
        if self.accept_suggestion_callback:
            self.accept_suggestion_callback(self)

    def on_clone_sheet(self, parameter=None):
        if self.clone_sheet_callback:
            self.clone_sheet_callback(self)

    def on_auto_populate_views(self, parameter=None):
        if self.auto_populate_callback:
            self.auto_populate_callback(self)
        
    def format_view_casing(self, case_type):
        for v in self.Views:
            if case_type == "UPPER":
                v.Name = v.Name.upper()
            elif case_type == "lower":
                v.Name = v.Name.lower()
            elif case_type == "Title":
                v.Name = v.Name.title()
                
    def sync_view_names(self):
        for v in self.Views:
            v.Name = self.SheetName
        self.update_action()
        self.notify_reconciliation_state()
        if self.validation_callback:
            self.validation_callback()
        
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
            self.update_action()
            
    def move_view_down(self, view_vm):
        idx = self.Views.IndexOf(view_vm)
        if idx >= 0 and idx < self.Views.Count - 1:
            self.Views.Move(idx, idx + 1)
            self.reassign_view_numbers()
            self.update_action()
            
    def reassign_view_numbers(self):
        num = 1
        for v in self.Views:
            if v.IsDetailNumberEditable:
                v.ViewNumber = str(num)
                num += 1
        self.OnPropertyChanged("HasNoViews")
        self.OnPropertyChanged("ViewsSummaryBadge")
        
    def populate_available_names(self):
        import classification
        self.AvailableNames.Clear()
        
        seen = set()
        def add_name(n):
            if not n:
                return
            n_str = u"{}".format(n).strip()
            if not n_str or n_str in seen:
                return
            seen.add(n_str)
            self.AvailableNames.Add(n_str)

        # 1) Suggested name
        if self._suggested_name:
            add_name(self._suggested_name)

        # 2) Clean names of views
        for v in self._views:
            clean_v = self._clean_view_name(v.OriginalName or v.Name)
            if clean_v:
                add_name(clean_v)

        # 3) Original sheet name
        if self._original_name:
            add_name(self._original_name)

        # 4) Standard NAMING_SCHEMES variants derived from base_name
        base_name = self.SheetName or self._original_name or ""
        if base_name.endswith("s") and not base_name.endswith("ss"):
            base_name = base_name[:-1]
        if base_name:
            add_name(base_name)
            schemes = classification.NAMING_SCHEMES
            for scheme_list in schemes.values():
                for mod in scheme_list:
                    add_name(u"{} - {}".format(base_name, mod))

    @property
    def SheetSeries(self): return self._sheet_series
    @SheetSeries.setter
    def SheetSeries(self, val):
        self._sheet_series = val or "Unknown"
        self.OnPropertyChanged("SheetSeries")
        self.OnPropertyChanged("CollectionAndSeries")
        self.OnPropertyChanged("IsUnknownSeries")
        self.OnPropertyChanged("OverrideSeriesText")

    @property
    def SheetNumber(self): return self._sheet_number
    @SheetNumber.setter
    def SheetNumber(self, val):
        if val is None: val = u""
        val_str = u"{}".format(val)
        if self._sheet_number == val_str: return
        
        old_val = self._sheet_number
        self._sheet_number = val_str
        self.NumberDiff.ProposedValue = self._sheet_number
        self.IsDirty = True
        self._is_user_modified = True
        self.OnPropertyChanged("SheetNumber")
        self.OnPropertyChanged("HasSuggestion")
        self.OnPropertyChanged("SuggestionShortText")
        self.reclassify()
        self.update_action()
        self.notify_reconciliation_state()
        if hasattr(self, 'number_changed_callback') and self.number_changed_callback:
            self.number_changed_callback(self, old_val, self._sheet_number)

    @property
    def SheetName(self): return self._sheet_name
    @SheetName.setter
    def SheetName(self, val):
        if val is None: val = u""
        val_str = u"{}".format(val)
        if self._sheet_name == val_str: return
        self._sheet_name = val_str
        self.NameDiff.ProposedValue = self._sheet_name
        self.IsDirty = True
        self._is_user_modified = True
        self.OnPropertyChanged("SheetName")
        self.OnPropertyChanged("HasSuggestion")
        self.OnPropertyChanged("SuggestionShortText")
        self.reclassify()
        self.update_action()
        self.notify_reconciliation_state()

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
        return "{}  ›  {}".format(c, s)

    @property
    def IsChecked(self): return self._is_checked
    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = val
        self.IsDirty = True
        self.OnPropertyChanged("IsChecked")
        self.notify_reconciliation_state()
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
        elif "RENAME" in self._action or self._action == "UPDATE": return u"✎"
        elif self._action in ["UNRECONCILED", "EXTRA"]: return u"⚠"
        elif self._action == "PURGE": return u"✖"
        return u"·"
        
    @property
    def ActionBrush(self):
        if self.HasWarning or self._action == "PURGE":
            return "#DC2626"
        elif self.IsAccepted or self.IsTemplate:
            return "#0078D7"
        elif self.CanAcceptProposal:
            return "#2563EB"
        else:
            return "#6B7280"

    @property
    def DisciplineName(self): return self._discipline_name
    @DisciplineName.setter
    def DisciplineName(self, val):
        self._discipline_name = val
        self.OnPropertyChanged("DisciplineName")
        self.OnPropertyChanged("DisciplineCode")
        self.OnPropertyChanged("DisciplineShortLabel")
        self.notify_reconciliation_state()
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
        
    def undo_changes(self, parameter=None):
        self._sheet_number = self.OriginalNumber
        self._sheet_name = self.OriginalName
        self._is_user_modified = False
        
        views_to_keep = []
        for v in self._views:
            if not v.IsNew:
                v.Name = v.OriginalName
                v.Scale = v.OriginalScale
                v.ViewNumber = v.OriginalViewNumber
                v.HasGraphicScale = v.OriginalHasGraphicScale
                v.HasNorthArrow = v.OriginalHasNorthArrow
                views_to_keep.append(v)
        self._views.Clear()
        for v in views_to_keep:
            self._views.Add(v)
            
        self._action = "MATCHED" if not self.IsTemplate else "CREATE"
        self._is_checked = False
        self.update_diffs()
        self.notify_reconciliation_state()
        if self.validation_callback:
            self.validation_callback()
        
    def add_view(self, param=None):
        from Autodesk.Revit.DB import ElementId
        default_name = self.SheetName if self.SheetName else "New View"
        new_v = ViewViewModel(ElementId.InvalidElementId, default_name, is_new=True, parent_sheet=self)
        self.Views.Add(new_v)
        self.reassign_view_numbers()
        self.IsExpanded = True
        self.update_action()

    def update_action(self):
        if getattr(self, '_is_purged', False) or self._action == "PURGE":
            return
        
        has_changes = (self.SheetName != self.OriginalName) or (self.SheetNumber != self.OriginalNumber)
        if has_changes:
            if self.SheetName != self.OriginalName and self.SheetNumber != self.OriginalNumber:
                self.MatchStatus = "RENAME_BOTH"
            elif self.SheetName != self.OriginalName:
                self.MatchStatus = "RENAME_NAME"
            else:
                self.MatchStatus = "RENAME_NUMBER"
        else:
            self.MatchStatus = "MATCHED"
            
        views_changed = False
        for v in self.Views:
            if v.has_pending_changes():
                views_changed = True
                break
            
        if self.IsTemplate:
            self.Action = "CREATE"
        elif self.MatchStatus != "MATCHED" or views_changed:
            self.Action = "UPDATE"
        else:
            self.Action = "MATCHED"
            
        if self.Action == "MATCHED" and not views_changed and not self.IsTemplate:
            self.IsChecked = False
            
        # View validation
        nums = set()
        for v in self.Views:
            if v.ViewNumber and v.IsDetailNumberEditable:
                if v.ViewNumber in nums:
                    v._validation_warning = "Duplicate Detail Number on sheet."
                else:
                    nums.add(v.ViewNumber)
                    if v._validation_warning == "Duplicate Detail Number on sheet.":
                        v._validation_warning = ""
            v.check_semantic_mismatch()
            
        self.notify_reconciliation_state()
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
            match = re.match(r"^([A-Z]+)[- ]?(\d)(\d\d)([A-Za-z]?)(.*)", self._sheet_number)
            if match:
                series_num = match.group(2)
            else:
                match_alt = re.match(r"^([A-Z]+)(\d)[- ]?(\d+)", self._sheet_number)
                if match_alt:
                    series_num = match_alt.group(2)
                
            c_info = classification.classify_sheet(self._sheet_number, self._sheet_name)
            
            if c_info:
                if c_info.get("drawingTypeCode", "99") != "99":
                    dtc = c_info["drawingTypeCode"]
                    if len(dtc) > 0:
                        series_num = dtc[0]
                if c_info.get("contentGroup") and c_info.get("contentGroup") != "Uncategorized":
                    self._content_group_name = c_info["contentGroup"]
                    self.OnPropertyChanged("ContentGroupName")
                if c_info.get("discipline") and c_info.get("discipline") != "Unknown":
                    self._discipline_name = c_info["discipline"]
                    self.OnPropertyChanged("DisciplineName")
            
            series_name_str = series_map.get(series_num, series_num)
            new_series = "0{}. {}".format(series_num, series_name_str)
            
            if getattr(self, "_sheet_series", "") != new_series:
                self.SheetSeries = new_series
        except Exception:
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
        if current_series == "Unknown" or not current_series:
            return None
            
        for opt in self.AvailableSeriesOptions:
            if opt == current_series:
                return opt
            parts = opt.split('.')
            curr_parts = str(current_series).split('.')
            if len(parts) > 0 and parts[0].isdigit() and len(curr_parts) > 0 and curr_parts[0].isdigit():
                if int(parts[0]) == int(curr_parts[0]):
                    return opt
        return None
        
    @OverrideSeriesText.setter
    def OverrideSeriesText(self, value):
        if not value or value == getattr(self, "_sheet_series", None):
            self._override_series_text = value
            return
        self._override_series_text = value
        parts = value.split('.')
        if len(parts) > 0 and parts[0].isdigit():
            code_str = str(int(parts[0]))
            self._auto_resolve_series(code_str)
        self.OnPropertyChanged("OverrideSeriesText")
        
    def _auto_resolve_series(self, series_code):
        import re
        import classification
        
        disc_code = "A"
        match = re.match(r"^([A-Z]+)[- ]?(\d)", self._sheet_number)
        if match: 
            disc_code = match.group(1)
        else:
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
        ctx = {}
        if hasattr(self, 'context_callback') and self.context_callback:
            ctx = self.context_callback() or {}
            append_sequence = ctx.get("append_sequence", False)
            all_sheets = ctx.get("all_sheets", [])
            
        pad_fmt = "{:02d}"
        if ctx and ctx.get("is_100_based", False):
            pad_fmt = "{:03d}"
        else:
            for s in all_sheets:
                n = getattr(s, "SheetNumber", "")
                if re.match(r"^([A-Z]+)[- ]?\d(\d\d\d)", n):
                    pad_fmt = "{:03d}"
                    break

        used_nums = []
        for s in all_sheets:
            if s == self: continue
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
                        
        has_dash = "-" in getattr(self, "SheetNumber", "-") or True
        dash_str = "-" if has_dash else ""
        proposed_number = "{}{}{}{}".format(disc_code, dash_str, series_code, pad_fmt.format(next_num))
        
        proposed_name = self.SheetName
        if len(self.Views) > 0 and self.Views[0].Name:
            proposed_name = self.Views[0].Name
            
        c_info = classification.classify_sheet(proposed_number, proposed_name)
        if c_info and c_info.get("originalName") != proposed_name:
            if c_info.get("drawingTypeCode", "99")[0] == series_code:
                proposed_name = c_info["originalName"]
        
        self.SheetName = proposed_name
        self.SheetNumber = proposed_number


class NavTreeNode(ViewModelBase):
    def __init__(self, name, node_type, tag=None, parent=None, is_expanded=True):
        ViewModelBase.__init__(self)
        self._name = name
        self._node_type = node_type
        self.Tag = tag
        self.Parent = parent
        self.callback = None
        self._children = ObservableCollection[NavTreeNode]()
        self._is_expanded = bool(is_expanded)
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
        
        self._configure_grids_command = RelayCommand(self.on_configure_grids)

    @property
    def Name(self): return self._name
    @Name.setter
    def Name(self, val):
        self._name = val
        self.OnPropertyChanged("Name")

    @property
    def NodeType(self): return self._node_type
    @property
    def Children(self): return self._children
    @property
    def ConfigureGridsCommand(self): return self._configure_grids_command

    @property
    def IsSegmentable(self): return self._is_segmentable
    @IsSegmentable.setter
    def IsSegmentable(self, val):
        self._is_segmentable = bool(val)
        self.OnPropertyChanged("IsSegmentable")
        
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
            if val is not None:
                for child in self.Children:
                    if child.IsEnabled:
                        child._set_checked_from_parent(val)
            elif val is None and len(self.Children) > 0:
                for child in self.Children:
                    if child.IsEnabled:
                        child._set_checked_from_parent(False)
                        
            if self.Parent:
                self.Parent._evaluate_checked_state()
                
            self._is_updating = False

        if self.callback: self.callback()

    def _set_checked_from_parent(self, val):
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
        if not self.Children: return
        
        has_checked = False
        has_unchecked = False
        has_indeterminate = False
        
        enabled_children = [c for c in self.Children if c.IsEnabled]
        if not enabled_children:
            enabled_children = self.Children
            
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
            self.IsChecked = False
        
    @property
    def TargetOpacity(self): return 1.0 if self.IsTargetIncluded else 0.5


class MatrixCellViewModel(ViewModelBase):
    """
    Represents a single [Row (Level/Scope), Column (AIA Series)] cell on the 2D Building x Series Matrix Board.
    """
    def __init__(self, row_key, col_key, placeholder_text="+ Add", can_split_sectors=False, show_quick_add=True, add_sheet_callback=None, split_sectors_callback=None):
        ViewModelBase.__init__(self)
        self.RowKey = row_key
        self.ColKey = col_key
        self._placeholder_text = placeholder_text
        self._can_split_sectors = bool(can_split_sectors)
        self._show_quick_add = bool(show_quick_add)
        self._sheets = ObservableCollection[object]()
        self.add_sheet_callback = add_sheet_callback
        self.split_sectors_callback = split_sectors_callback
        self._add_sheet_command = RelayCommand(self.on_add_sheet)
        self._split_sectors_command = RelayCommand(self.on_split_sectors)

    @property
    def Sheets(self): return self._sheets
    @property
    def HasSheets(self): return self._sheets.Count > 0
    @property
    def IsEmpty(self): return self._sheets.Count == 0
    @property
    def CellPlaceholderText(self): return self._placeholder_text
    @property
    def CanSplitSectors(self): return self._can_split_sectors
    @property
    def ShowQuickAdd(self): return self._show_quick_add
    @property
    def AddSheetCommand(self): return self._add_sheet_command
    @property
    def SplitSectorsCommand(self): return self._split_sectors_command

    def add_sheet_vm(self, sheet_vm):
        self._sheets.Add(sheet_vm)
        self.OnPropertyChanged("HasSheets")
        self.OnPropertyChanged("IsEmpty")

    def on_add_sheet(self, parameter=None):
        if self.add_sheet_callback:
            self.add_sheet_callback(self)

    def on_split_sectors(self, parameter=None):
        if self.split_sectors_callback:
            self.split_sectors_callback(self)


class MatrixRowViewModel(ViewModelBase):
    """
    Represents a horizontal Building Level or Project Scope row in the 2D Matrix Board.
    """
    def __init__(self, row_key, row_title, row_subtitle="", is_level_row=True, add_sheet_callback=None, split_sectors_callback=None):
        ViewModelBase.__init__(self)
        self.RowKey = row_key
        self._row_title = row_title
        self._row_subtitle = row_subtitle
        self._is_level_row = bool(is_level_row)

        self._cell_floor_plans = MatrixCellViewModel(row_key, "FLOOR_PLANS", "+ Floor Plan", can_split_sectors=is_level_row, show_quick_add=is_level_row, add_sheet_callback=add_sheet_callback, split_sectors_callback=split_sectors_callback)
        self._cell_ceiling_plans = MatrixCellViewModel(row_key, "CEILING_PLANS", "+ Ceiling Plan", can_split_sectors=is_level_row, show_quick_add=is_level_row, add_sheet_callback=add_sheet_callback, split_sectors_callback=split_sectors_callback)
        self._cell_enlarged_plans = MatrixCellViewModel(row_key, "ENLARGED_PLANS", "+ Enlarged / 2x2", can_split_sectors=True, show_quick_add=is_level_row, add_sheet_callback=add_sheet_callback, split_sectors_callback=split_sectors_callback)
        self._cell_general = MatrixCellViewModel(row_key, "GENERAL", "+ Cover / Gen", can_split_sectors=False, show_quick_add=(not is_level_row), add_sheet_callback=add_sheet_callback, split_sectors_callback=split_sectors_callback)
        self._cell_elevations = MatrixCellViewModel(row_key, "ELEVATIONS", "+ Elevation", can_split_sectors=False, show_quick_add=(not is_level_row), add_sheet_callback=add_sheet_callback, split_sectors_callback=split_sectors_callback)
        self._cell_sections = MatrixCellViewModel(row_key, "SECTIONS", "+ Section", can_split_sectors=False, show_quick_add=(not is_level_row), add_sheet_callback=add_sheet_callback, split_sectors_callback=split_sectors_callback)
        self._cell_details = MatrixCellViewModel(row_key, "DETAILS_SCHEDULES", "+ Detail / Sched", can_split_sectors=False, show_quick_add=(not is_level_row), add_sheet_callback=add_sheet_callback, split_sectors_callback=split_sectors_callback)

    @property
    def RowTitle(self): return self._row_title
    @property
    def RowSubtitle(self): return self._row_subtitle
    @property
    def IsLevelRow(self): return self._is_level_row

    @property
    def CellFloorPlans(self): return self._cell_floor_plans
    @property
    def CellCeilingPlans(self): return self._cell_ceiling_plans
    @property
    def CellEnlargedPlans(self): return self._cell_enlarged_plans
    @property
    def CellGeneral(self): return self._cell_general
    @property
    def CellElevations(self): return self._cell_elevations
    @property
    def CellSections(self): return self._cell_sections
    @property
    def CellDetails(self): return self._cell_details

    def all_cells(self):
        return [
            self._cell_floor_plans,
            self._cell_ceiling_plans,
            self._cell_enlarged_plans,
            self._cell_general,
            self._cell_elevations,
            self._cell_sections,
            self._cell_details,
        ]

    @property
    def RowBadge(self):
        total = sum(c.Sheets.Count for c in self.all_cells())
        return "{} Sheet{}".format(total, "s" if total != 1 else "")


class ManageSheetsViewModel(ViewModelBase):
    def __init__(self):
        ViewModelBase.__init__(self)
        self._editor_items = ObservableCollection[object]()
        self._nav_root = ObservableCollection[NavTreeNode]()
        self._matrix_rows = ObservableCollection[MatrixRowViewModel]()
        self._unreconciled_items = ObservableCollection[object]()
        self._selected_sheet = None
        self._matrix_summary_text = "Ready"
        
        self._is_push_enabled = False
        self._is_reviewer_ready = True
        self._has_valid_work = False
        self._is_cartoon_drawer_open = False

        self.fill_floor_plans_callback = None
        self.fill_ceiling_plans_callback = None
        self.auto_fix_messy_callback = None

        self._fill_floor_plans_command = RelayCommand(self.on_fill_floor_plans)
        self._fill_ceiling_plans_command = RelayCommand(self.on_fill_ceiling_plans)
        self._auto_fix_messy_sheets_command = RelayCommand(self.on_auto_fix_messy_sheets)

    def on_fill_floor_plans(self, parameter=None):
        if self.fill_floor_plans_callback:
            self.fill_floor_plans_callback()

    def on_fill_ceiling_plans(self, parameter=None):
        if self.fill_ceiling_plans_callback:
            self.fill_ceiling_plans_callback()

    def on_auto_fix_messy_sheets(self, parameter=None):
        if self.auto_fix_messy_callback:
            self.auto_fix_messy_callback()

    @property
    def FillFloorPlansCommand(self): return self._fill_floor_plans_command
    @property
    def FillCeilingPlansCommand(self): return self._fill_ceiling_plans_command
    @property
    def AutoFixMessySheetsCommand(self): return self._auto_fix_messy_sheets_command

    @property
    def EditorItems(self): return self._editor_items
    
    @property
    def NavRoot(self): return self._nav_root

    @property
    def MatrixRows(self): return self._matrix_rows

    @property
    def UnreconciledItems(self): return self._unreconciled_items

    @property
    def HasUnreconciledItems(self): return self._unreconciled_items.Count > 0

    @property
    def UnreconciledCountText(self):
        return "{} Messy / Unmatched Sheet{} Found".format(
            self._unreconciled_items.Count,
            "s" if self._unreconciled_items.Count != 1 else ""
        )

    @property
    def SelectedSheet(self): return self._selected_sheet
    @SelectedSheet.setter
    def SelectedSheet(self, value):
        if self._selected_sheet and self._selected_sheet != value:
            self._selected_sheet.IsSelected = False
        self._selected_sheet = value
        if self._selected_sheet:
            self._selected_sheet.IsSelected = True
        self.OnPropertyChanged("SelectedSheet")
        self.OnPropertyChanged("HasSelectedSheet")

    @property
    def HasSelectedSheet(self): return self._selected_sheet is not None

    @property
    def MatrixSummaryText(self): return self._matrix_summary_text
    @MatrixSummaryText.setter
    def MatrixSummaryText(self, val):
        self._matrix_summary_text = val or ""
        self.OnPropertyChanged("MatrixSummaryText")

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

    @property
    def IsCartoonDrawerOpen(self): return self._is_cartoon_drawer_open
    @IsCartoonDrawerOpen.setter
    def IsCartoonDrawerOpen(self, value):
        self._is_cartoon_drawer_open = bool(value)
        self.OnPropertyChanged("IsCartoonDrawerOpen")
        self.OnPropertyChanged("DrawerToggleText")

    @property
    def DrawerToggleText(self):
        return u"▲ Hide Cartoon Rules" if self._is_cartoon_drawer_open else u"⚙ Cartoon Rules"

