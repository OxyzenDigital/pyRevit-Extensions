# -*- coding: utf-8 -*-
import os

try:
    from pyrevit import forms
    ReactiveBase = forms.Reactive
except Exception:
    class ReactiveBase(object):
        def OnPropertyChanged(self, property_name):
            pass


class NamedRangeItem(ReactiveBase):
    def __init__(self, data):
        self._name = data.get('name', '')
        self._sheet = data.get('sheet', '')
        self._formula = data.get('formula', '')
        self._is_full_sheet = bool(data.get('is_full_sheet', False))

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, value):
        self._name = value
        self.OnPropertyChanged('Name')

    @property
    def Sheet(self):
        return self._sheet

    @Sheet.setter
    def Sheet(self, value):
        self._sheet = value
        self.OnPropertyChanged('Sheet')

    @property
    def Formula(self):
        return self._formula

    @Formula.setter
    def Formula(self, value):
        self._formula = value
        self.OnPropertyChanged('Formula')

    @property
    def IsFullSheet(self):
        return self._is_full_sheet

    @IsFullSheet.setter
    def IsFullSheet(self, value):
        self._is_full_sheet = bool(value)
        self.OnPropertyChanged('IsFullSheet')
        self.OnPropertyChanged('DisplayMeta')

    @property
    def DisplayMeta(self):
        if self._is_full_sheet:
            return "Worksheet Range"
        return "({})".format(self._sheet if self._sheet else "Global")

    # Backward compatibility properties
    @property
    def name(self):
        return self._name

    @property
    def sheet(self):
        return self._sheet if self._sheet else "Global"

    @property
    def formula(self):
        return self._formula


class LinkedTableItem(ReactiveBase):
    def __init__(self, family_name, source_path="", sheet_name="", range_name="", text_scale=80.0, family_id=None, hint_path=""):
        self._family_name = family_name
        self._source_path = source_path
        self._hint_path = hint_path or ""
        self._sheet_name = sheet_name
        self._range_name = range_name
        self._text_scale = text_scale
        self.FamilyId = family_id
        self._is_checked = False
        self._is_selected = False
        self.StateCallback = None

    @property
    def FamilyName(self):
        return self._family_name

    @FamilyName.setter
    def FamilyName(self, val):
        self._family_name = val
        self.OnPropertyChanged('FamilyName')
        self.OnPropertyChanged('LinkHintToolTip')

    @property
    def HintPath(self):
        return self._hint_path

    @HintPath.setter
    def HintPath(self, val):
        self._hint_path = val or ""
        self.OnPropertyChanged('HintPath')
        self.OnPropertyChanged('FileName')
        self.OnPropertyChanged('StatusText')
        self.OnPropertyChanged('LinkHintToolTip')

    @property
    def SourcePath(self):
        return self._source_path

    @SourcePath.setter
    def SourcePath(self, val):
        self._source_path = val or ""
        self.OnPropertyChanged('SourcePath')
        self.OnPropertyChanged('FileName')
        self.OnPropertyChanged('FileExists')
        self.OnPropertyChanged('StatusText')
        self.OnPropertyChanged('LinkHintToolTip')

    @property
    def FileName(self):
        if self._source_path:
            return os.path.basename(self._source_path)
        elif self._hint_path:
            return "~ " + os.path.basename(self._hint_path)
        return "Unlinked"

    @property
    def SheetName(self):
        return self._sheet_name

    @SheetName.setter
    def SheetName(self, val):
        self._sheet_name = val
        self.OnPropertyChanged('SheetName')
        self.OnPropertyChanged('DisplaySource')
        self.OnPropertyChanged('LinkHintToolTip')

    @property
    def RangeName(self):
        return self._range_name

    @RangeName.setter
    def RangeName(self, val):
        self._range_name = val
        self.OnPropertyChanged('RangeName')
        self.OnPropertyChanged('DisplaySource')
        self.OnPropertyChanged('LinkHintToolTip')

    @property
    def DisplaySource(self):
        return "{} | {}".format(self._sheet_name or "Sheet", self._range_name or "<Entire Worksheet>")

    @property
    def TextScale(self):
        return self._text_scale

    @TextScale.setter
    def TextScale(self, val):
        self._text_scale = val
        self.OnPropertyChanged('TextScale')

    @property
    def IsChecked(self):
        return self._is_checked

    @IsChecked.setter
    def IsChecked(self, val):
        self._is_checked = bool(val)
        self.OnPropertyChanged('IsChecked')
        if callable(self.StateCallback):
            self.StateCallback()

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, val):
        self._is_selected = bool(val)
        self.OnPropertyChanged('IsSelected')

    @property
    def FileExists(self):
        return bool(self._source_path and os.path.exists(self._source_path))

    @property
    def StatusText(self):
        if self.FileExists:
            return "Ready"
        elif self._source_path:
            return "Missing File"
        elif self._hint_path:
            return "Hint Only"
        return "No Link"

    @property
    def LinkHintToolTip(self):
        line1 = "Table: {}".format(self._family_name)
        if self._source_path:
            line2 = "Original Path: {}".format(self._source_path)
        elif self._hint_path:
            line2 = "Hint Path: {}".format(self._hint_path)
        else:
            line2 = "Original Path: Not recorded"
        line3 = "Sheet / Range: {}".format(self.DisplaySource)
        line4 = "Status: {}".format(self.StatusText)
        return "{}\n{}\n{}\n{}".format(line1, line2, line3, line4)


class ExcelImportViewModel(ReactiveBase):
    def __init__(self):
        self._file_path = ""
        self._sheets = []
        self._selected_sheet = None
        self._ranges = []
        self._selected_range = None
        self._table_name = ""
        self._text_scale = "80"
        self._place_on_view = True
        self._status_text = "Ready"
        self._linked_tables = []
        self._selected_linked_table = None
        self._linked_file_hint = "Select a table, choose a saved source, or browse for an Excel file."
        self._recent_files = []
        self._selected_recent_file = None
        self._can_import = False
        self._can_update_checked = False
        self._can_place_selected = False
        self._import_button_text = "Create & Place Table"
        self.existing_family_names = set()

    @property
    def FilePath(self):
        return self._file_path

    @FilePath.setter
    def FilePath(self, val):
        self._file_path = val or ""
        self.OnPropertyChanged('FilePath')
        self.evaluate_states()

    @property
    def LinkedFileHint(self):
        return self._linked_file_hint

    @LinkedFileHint.setter
    def LinkedFileHint(self, val):
        self._linked_file_hint = val or ""
        self.OnPropertyChanged('LinkedFileHint')

    @property
    def RecentFiles(self):
        return self._recent_files

    @RecentFiles.setter
    def RecentFiles(self, val):
        self._recent_files = list(val) if val else []
        self.OnPropertyChanged('RecentFiles')

    @property
    def SelectedRecentFile(self):
        return self._selected_recent_file

    @SelectedRecentFile.setter
    def SelectedRecentFile(self, val):
        self._selected_recent_file = val
        self.OnPropertyChanged('SelectedRecentFile')

    @property
    def Sheets(self):
        return self._sheets

    @Sheets.setter
    def Sheets(self, val):
        self._sheets = val
        self.OnPropertyChanged('Sheets')

    @property
    def SelectedSheet(self):
        return self._selected_sheet

    @SelectedSheet.setter
    def SelectedSheet(self, val):
        self._selected_sheet = val
        self.OnPropertyChanged('SelectedSheet')

    @property
    def Ranges(self):
        return self._ranges

    @Ranges.setter
    def Ranges(self, val):
        self._ranges = val
        self.OnPropertyChanged('Ranges')

    @property
    def SelectedRange(self):
        return self._selected_range

    @SelectedRange.setter
    def SelectedRange(self, val):
        self._selected_range = val
        self.OnPropertyChanged('SelectedRange')

    @property
    def TableName(self):
        return self._table_name

    @TableName.setter
    def TableName(self, val):
        self._table_name = val or ""
        self.OnPropertyChanged('TableName')
        self.evaluate_states()

    @property
    def TextScale(self):
        return self._text_scale

    @TextScale.setter
    def TextScale(self, val):
        self._text_scale = str(val) if val is not None else ""
        self.OnPropertyChanged('TextScale')

    @property
    def PlaceOnView(self):
        return self._place_on_view

    @PlaceOnView.setter
    def PlaceOnView(self, val):
        self._place_on_view = bool(val)
        self.OnPropertyChanged('PlaceOnView')
        self.evaluate_states()

    @property
    def StatusText(self):
        return self._status_text

    @StatusText.setter
    def StatusText(self, val):
        self._status_text = val
        self.OnPropertyChanged('StatusText')

    @property
    def LinkedTables(self):
        return self._linked_tables

    @LinkedTables.setter
    def LinkedTables(self, val):
        self._linked_tables = list(val) if val else []
        for t in self._linked_tables:
            t.StateCallback = self.evaluate_states
        self.OnPropertyChanged('LinkedTables')
        self.evaluate_states()

    @property
    def SelectedLinkedTable(self):
        return self._selected_linked_table

    @SelectedLinkedTable.setter
    def SelectedLinkedTable(self, val):
        self._selected_linked_table = val
        self.OnPropertyChanged('SelectedLinkedTable')
        self.evaluate_states()

    @property
    def CanImport(self):
        return self._can_import

    @CanImport.setter
    def CanImport(self, val):
        self._can_import = bool(val)
        self.OnPropertyChanged('CanImport')

    @property
    def CanUpdateChecked(self):
        return self._can_update_checked

    @CanUpdateChecked.setter
    def CanUpdateChecked(self, val):
        self._can_update_checked = bool(val)
        self.OnPropertyChanged('CanUpdateChecked')

    @property
    def CanPlaceSelected(self):
        return self._can_place_selected

    @CanPlaceSelected.setter
    def CanPlaceSelected(self, val):
        self._can_place_selected = bool(val)
        self.OnPropertyChanged('CanPlaceSelected')

    @property
    def ImportButtonText(self):
        return self._import_button_text

    @ImportButtonText.setter
    def ImportButtonText(self, val):
        self._import_button_text = val
        self.OnPropertyChanged('ImportButtonText')

    def evaluate_states(self):
        has_file = bool(self._file_path and os.path.exists(self._file_path))
        has_name = bool(self._table_name and self._table_name.strip())
        self.CanImport = has_file and has_name
        self.CanUpdateChecked = any(bool(t.IsChecked) for t in self._linked_tables)
        self.CanPlaceSelected = (self._selected_linked_table is not None)

        if self._table_name and self._table_name.strip() in self.existing_family_names:
            self.ImportButtonText = "Update Existing Table"
        else:
            self.ImportButtonText = "Create & Place Table" if self._place_on_view else "Create Table"

    def select_all_linked(self, state=True):
        for t in self._linked_tables:
            t.IsChecked = bool(state)
        self.evaluate_states()
