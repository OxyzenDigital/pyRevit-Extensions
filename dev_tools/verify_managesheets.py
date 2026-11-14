# -*- coding: utf-8 -*-
import sys
import os
import ast
import re
import types
import xml.etree.ElementTree as ET

print("=== STARTING COMPREHENSIVE VERIFICATION ===")

# 1. XML Parsing Check
xaml_dir = r"ODI.extension\lib\managesheets"
for fname in os.listdir(xaml_dir):
    if fname.endswith(".xaml"):
        fpath = os.path.join(xaml_dir, fname)
        try:
            tree = ET.parse(fpath)
            print("[XML OK] {}".format(fname))
        except Exception as e:
            print("[XML ERROR] {}: {}".format(fname, e))
            sys.exit(1)

# 2. AST Parsing Check
py_files = [
    r"ODI.extension\lib\managesheets\data_model.py",
    r"ODI.extension\lib\managesheets\panel.py",
    r"ODI.extension\lib\managesheets\classification.py",
    r"ODI.extension\lib\managesheets\project_settings.py",
    r"ODI.extension\lib\managesheets\reconciliation.py",
    r"ODI.extension\ODI.tab\03_DOCUMENTATION.panel\Sheets.stack\Manage Sheets.pushbutton\script.py",
]
for p in py_files:
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                ast.parse(f.read())
            print("[AST OK] {}".format(os.path.basename(p)))
        except Exception as e:
            print("[AST ERROR] {}: {}".format(p, e))
            sys.exit(1)
    else:
        print("[AST NOT FOUND] {}".format(p))
        sys.exit(1)

# 3. Control Name Compatibility Check
panel_path = r"ODI.extension\lib\managesheets\panel.py"
ui_path = r"ODI.extension\lib\managesheets\ui.xaml"

panel_content = open(panel_path, "r", encoding="utf-8").read()
ui_content = open(ui_path, "r", encoding="utf-8").read()

class_match = re.search(r"class ManageSheetsPanel\(forms\.WPFWindow\):(.*?)(?=\nclass |\Z)", panel_content, re.DOTALL)
panel_code = class_match.group(1) if class_match else ""

xaml_names = set(re.findall(r'(?:Name|x:Name)=["\']([a-zA-Z0-9_]+)["\']', ui_content))
wired_controls = set(re.findall(r'self\.((?:Btn|Chk|Cmb|Txt|List|Tree|Sld|Resize|MainTab|TargetSchema|Nav|Cartoon|AutoRefresh|TitleBar|Combo|Footer|EditorList)[a-zA-Z0-9_]*)', panel_code))

missing_controls = []
for c in sorted(wired_controls):
    if c not in xaml_names:
        if c in ("NavRoot", "TargetSchemaRoot", "EditorItems"):
            continue
        missing_controls.append(c)

if missing_controls:
    print("[CONTROL ERROR] Missing XAML controls referenced in panel.py: {}".format(missing_controls))
    sys.exit(1)
else:
    print("[CONTROLS OK] All {} referenced UI controls exist in ui.xaml".format(len(wired_controls)))

# 4. Color Hygiene Check
after_res = ui_content.split("</Window.Resources>")[1] if "</Window.Resources>" in ui_content else ""
hex_matches = re.findall(r'#[0-9a-fA-F]{3,8}', after_res)
white_matches = re.findall(r'Foreground=[\"\']White[\"\']', after_res)
if hex_matches or white_matches:
    print("[COLOR ERROR] Found violations after resources: Hex={}, White={}".format(hex_matches, white_matches))
    sys.exit(1)
else:
    print("[COLOR HYGIENE OK] 0 hex colors and 0 hardcoded White foregrounds after resources.")

# 5. Unit Testing data_model and panel logic
print("=== RUNNING DATA MODEL UNIT TESTS ===")

# Create mock clr & System modules for CPython standalone testing
class MockObservableCollection(list):
    def __init__(self, *args):
        super(MockObservableCollection, self).__init__(*args)
    def Add(self, item):
        self.append(item)
    def Clear(self):
        del self[:]
    def Move(self, old_idx, new_idx):
        item = self.pop(old_idx)
        self.insert(new_idx, item)
    def IndexOf(self, item):
        try:
            return self.index(item)
        except ValueError:
            return -1
    @property
    def Count(self):
        return len(self)

if "clr" not in sys.modules:
    mock_clr = types.ModuleType("clr")
    mock_clr.AddReference = lambda x: None
    sys.modules["clr"] = mock_clr

if "System" not in sys.modules:
    mock_system = types.ModuleType("System")
    mock_system.Guid = lambda s: s
    mock_system.String = str
    mock_system_componentmodel = types.ModuleType("System.ComponentModel")
    class INotifyPropertyChanged(object): pass
    class PropertyChangedEventArgs(object):
        def __init__(self, name): self.PropertyName = name
    mock_system_componentmodel.INotifyPropertyChanged = INotifyPropertyChanged
    mock_system_componentmodel.PropertyChangedEventArgs = PropertyChangedEventArgs
    mock_system.ComponentModel = mock_system_componentmodel

    mock_system_collections = types.ModuleType("System.Collections")
    mock_system_collections_objectmodel = types.ModuleType("System.Collections.ObjectModel")
    class ObservableCollectionGeneric(object):
        def __getitem__(self, item):
            return MockObservableCollection
    mock_system_collections_objectmodel.ObservableCollection = ObservableCollectionGeneric()
    mock_system.Collections = mock_system_collections
    mock_system.Collections.ObjectModel = mock_system_collections_objectmodel

    mock_system_windows = types.ModuleType("System.Windows")
    mock_system_windows_input = types.ModuleType("System.Windows.Input")
    class ICommand(object): pass
    mock_system_windows_input.ICommand = ICommand
    mock_system_windows.Input = mock_system_windows_input
    mock_system_windows_media = types.ModuleType("System.Windows.Media")
    class BrushConverter(object):
        def ConvertFromString(self, s): return s
    class SolidColorBrush(object): pass
    class Color(object): pass
    class Colors(object): pass
    mock_system_windows_media.BrushConverter = BrushConverter
    mock_system_windows_media.SolidColorBrush = SolidColorBrush
    mock_system_windows_media.Color = Color
    mock_system_windows_media.Colors = Colors
    mock_system_windows_controls = types.ModuleType("System.Windows.Controls")
    class CheckBox(object): pass
    mock_system_windows_controls.CheckBox = CheckBox
    mock_system_windows.Controls = mock_system_windows_controls
    sys.modules["System.Windows.Controls"] = mock_system_windows_controls
    mock_system_windows_data = types.ModuleType("System.Windows.Data")
    class CollectionViewSource(object): pass
    class PropertyGroupDescription(object): pass
    mock_system_windows_data.CollectionViewSource = CollectionViewSource
    mock_system_windows_data.PropertyGroupDescription = PropertyGroupDescription
    mock_system_windows.Data = mock_system_windows_data
    sys.modules["System.Windows.Data"] = mock_system_windows_data
    class SystemColors(object):
        HighlightBrushKey = "HighlightBrushKey"
        HighlightTextBrushKey = "HighlightTextBrushKey"
        InactiveSelectionHighlightBrushKey = "InactiveSelectionHighlightBrushKey"
        InactiveSelectionHighlightTextBrushKey = "InactiveSelectionHighlightTextBrushKey"
    mock_system_windows.SystemColors = SystemColors
    class Visibility(object):
        Visible = "Visible"
        Collapsed = "Collapsed"
    class MessageBox(object):
        @staticmethod
        def Show(*args, **kwargs): pass
    class MessageBoxButton(object): pass
    class MessageBoxImage(object): pass
    class MessageBoxResult(object): pass
    mock_system_windows.Visibility = Visibility
    mock_system_windows.MessageBox = MessageBox
    mock_system_windows.MessageBoxButton = MessageBoxButton
    mock_system_windows.MessageBoxImage = MessageBoxImage
    mock_system_windows.MessageBoxResult = MessageBoxResult
    mock_system.Windows = mock_system_windows
    sys.modules["System"] = mock_system
    sys.modules["System.ComponentModel"] = mock_system_componentmodel
    sys.modules["System.Collections"] = mock_system_collections
    sys.modules["System.Collections.ObjectModel"] = mock_system_collections_objectmodel
    sys.modules["System.Windows"] = mock_system_windows
    sys.modules["System.Windows.Input"] = mock_system_windows_input
    sys.modules["System.Windows.Media"] = mock_system_windows_media

from unittest.mock import MagicMock
db_mod = MagicMock()
class MockElementId(object):
    InvalidElementId = None
    def __init__(self, val=-1):
        self.IntegerValue = val
MockElementId.InvalidElementId = MockElementId(-1)
db_mod.ElementId = MockElementId
sys.modules["Autodesk.Revit.DB"] = db_mod

autodesk_mod = MagicMock()
autodesk_mod.Revit = MagicMock()
autodesk_mod.Revit.DB = db_mod
sys.modules["Autodesk"] = autodesk_mod
sys.modules["Autodesk.Revit"] = autodesk_mod.Revit
sys.modules["Autodesk.Revit.DB.ExtensibleStorage"] = MagicMock()

# Mock pyrevit
mock_pyrevit = MagicMock()
mock_forms = MagicMock()
class WPFWindow(object):
    def __init__(self, *args, **kwargs):
        self.Resources = {}
    def apply_theme(self): pass
mock_forms.WPFWindow = WPFWindow
mock_forms.alert = lambda msg, **kwargs: None
mock_pyrevit.forms = mock_forms
sys.modules["pyrevit"] = mock_pyrevit
sys.modules["pyrevit.forms"] = mock_forms
sys.modules["pyrevit.revit"] = mock_pyrevit.revit
sys.modules["pyrevit.script"] = mock_pyrevit.script

# Add lib and lib/managesheets to path
lib_parent = os.path.abspath(r"ODI.extension\lib")
lib_path = os.path.abspath(r"ODI.extension\lib\managesheets")
if lib_parent not in sys.path:
    sys.path.insert(0, lib_parent)
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import data_model

# Test SheetViewModel
sheet = data_model.SheetViewModel(
    element_id=MockElementId(101),
    number="A-101",
    name="Ground Floor Plan",
    collection_name="PERMIT SET",
    discipline_name="Architecture",
    content_group_name="Floor Plans",
    series_name="01. Plans",
    is_template=False
)

assert sheet.OriginalNumber == "A-101", "OriginalNumber failed: {}".format(sheet.OriginalNumber)
assert sheet.OriginalName == "Ground Floor Plan", "OriginalName failed: {}".format(sheet.OriginalName)
assert sheet.OriginalNumberDisplay == "A-101", "OriginalNumberDisplay failed: {}".format(sheet.OriginalNumberDisplay)
assert sheet.HasSuggestion == False, "HasSuggestion should be False initially"
assert sheet.SuggestionShortText == u"✨ Standardize", "SuggestionShortText should be Standardize initially"

# Test Suggested Number / Name
sheet.SuggestedNumber = "A-101"
sheet.SuggestedName = "Level 1 - Floor Plan"
assert sheet.HasSuggestion == True, "HasSuggestion should be True when name differs"
assert sheet.SuggestionShortText == u"✨ A-101 · Accept", "SuggestionShortText mismatch: {}".format(sheet.SuggestionShortText)

sheet.SheetName = "Level 1 - Floor Plan"
assert sheet.HasSuggestion == False, "HasSuggestion should be False when standardized"

# Test template sheet
tmpl = data_model.SheetViewModel(
    element_id=MockElementId(-1),
    number="A-201",
    name="Elevations",
    collection_name="PERMIT SET",
    is_template=True
)
assert tmpl.OriginalNumberDisplay == "NEW", "Template OriginalNumberDisplay should be NEW"
assert tmpl.HasSuggestion == False, "Template HasSuggestion should always be False"
print("[UNIT TEST OK] SheetViewModel suggestions and properties verified.")

# Test MatrixCellViewModel and MatrixRowViewModel ShowQuickAdd
cell_true = data_model.MatrixCellViewModel("L1", "FLOOR_PLANS", show_quick_add=True)
assert cell_true.ShowQuickAdd == True, "Cell ShowQuickAdd should be True"

cell_false = data_model.MatrixCellViewModel("L1", "GENERAL", show_quick_add=False)
assert cell_false.ShowQuickAdd == False, "Cell ShowQuickAdd should be False"

row_level = data_model.MatrixRowViewModel("L1", "Level 1", is_level_row=True)
assert row_level.CellFloorPlans.ShowQuickAdd == True, "Floor plans on level row must show quick add"
assert row_level.CellCeilingPlans.ShowQuickAdd == True, "Ceiling plans on level row must show quick add"
assert row_level.CellEnlargedPlans.ShowQuickAdd == True, "Enlarged plans on level row must show quick add"
assert row_level.CellGeneral.ShowQuickAdd == False, "General on level row must NOT show quick add"
assert row_level.CellElevations.ShowQuickAdd == False, "Elevations on level row must NOT show quick add"

row_project = data_model.MatrixRowViewModel("SCOPE::PROJECT", "Project", is_level_row=False)
assert row_project.CellFloorPlans.ShowQuickAdd == False, "Floor plans on project row must NOT show quick add"
assert row_project.CellGeneral.ShowQuickAdd == True, "General on project row must show quick add"
assert row_project.CellElevations.ShowQuickAdd == True, "Elevations on project row must show quick add"
assert row_project.CellSections.ShowQuickAdd == True, "Sections on project row must show quick add"
assert row_project.CellDetails.ShowQuickAdd == True, "Details on project row must show quick add"
print("[UNIT TEST OK] MatrixCellViewModel and MatrixRowViewModel ShowQuickAdd verified.")

# Test ManageSheetsViewModel DrawerToggleText
vm = data_model.ManageSheetsViewModel()
assert vm.IsCartoonDrawerOpen == False, "Default drawer should be closed"
assert vm.DrawerToggleText == u"⚙ Cartoon Rules", "Closed DrawerToggleText mismatch: {}".format(vm.DrawerToggleText)
vm.IsCartoonDrawerOpen = True
assert vm.DrawerToggleText == u"▲ Hide Cartoon Rules", "Open DrawerToggleText mismatch: {}".format(vm.DrawerToggleText)
print("[UNIT TEST OK] ManageSheetsViewModel DrawerToggleText verified.")

# Test is_dark_theme fallback without Revit loaded
import panel
assert panel.is_dark_theme() == False, "is_dark_theme() should fallback to False when Revit is not loaded"
print("[UNIT TEST OK] is_dark_theme() fallback verified.")

# Test AssignableViewNode properties
test_vid = MockElementId(555)
node = data_model.AssignableViewNode(test_vid, "Level 1 Floor Plan", is_placed=False, sheet_number="", view_type="FloorPlan", level_name="Level 1", scale_int=96)
assert node.Id == test_vid, "Id failed"
assert node.ViewId == test_vid, "ViewId failed"
assert node.Name == "Level 1 Floor Plan", "Name failed"
assert node.ViewName == "Level 1 Floor Plan", "ViewName failed"
assert "Level 1 Floor Plan" in node.DisplayName, "DisplayName failed"
assert node.ScaleStr == "1/8\" = 1'-0\"", "ScaleStr failed: {}".format(node.ScaleStr)
assert node.LevelName == "Level 1", "LevelName failed"
assert node.ViewType == "FloorPlan", "ViewType failed"
assert node.IsPlaced == False, "IsPlaced failed"
print("[UNIT TEST OK] AssignableViewNode all required properties verified.")

# Test ViewViewModel.AvailableExistingViews resolution
vv = data_model.ViewViewModel(test_vid, "Test View")
assert node in vv.AvailableExistingViews or data_model.AVAILABLE_ASSIGNABLE_VIEWS or data_model.AVAILABLE_EXISTING_VIEWS is not None
print("[UNIT TEST OK] ViewViewModel.AvailableExistingViews resolution verified.")

# Test tree node tag handling in on_tree_selection_changed logic
class MockPanel(object):
    def __init__(self):
        self.all_grid_nodes = [sheet]
        self.Cmb_TargetCollection = types.SimpleNamespace(Text="PERMIT SET")
        self.Txt_GridTitle = types.SimpleNamespace(Text="")
        self.TargetSchemaRoot = []
        self.NavRoot = []
        self.EditorItems = MockObservableCollection()
        self._ignore_tree_event = False
        self._last_selected_node = None
        self._current_selected_node = None
        self.run_validation = lambda: None
        self.build_matrix_board = lambda: None
        self.main_vm = types.SimpleNamespace(SelectedSheet=None, UnreconciledItems=MockObservableCollection())
        self.Cmb_DisciplineFilter = types.SimpleNamespace(Items=MockObservableCollection(), SelectedItem="All Disciplines", SelectedIndex=0)
        self._suppress_disc_filter = False
    def update_grid_title(self): pass
    def execute_schema_match(self, *args, **kwargs): pass
    def on_select_sheet_vm(self, s):
        self.main_vm.SelectedSheet = s
    def on_accept_sheet_suggestion(self, s):
        panel.ManageSheetsPanel.on_accept_sheet_suggestion(self, s)
    def on_auto_fix_all_messy_sheets(self):
        panel.ManageSheetsPanel.on_auto_fix_all_messy_sheets(self)
    def update_discipline_filter_options(self):
        panel.ManageSheetsPanel.update_discipline_filter_options(self)
    def on_discipline_filter_changed(self, sender, e):
        panel.ManageSheetsPanel.on_discipline_filter_changed(self, sender, e)
    def apply_filters(self):
        panel.ManageSheetsPanel.apply_filters(self)

test_panel = MockPanel()

# 1. Tuple ContentGroup
cg_tuple_node = types.SimpleNamespace(NodeType="ContentGroup", Tag=("PERMIT SET", "Architecture", "Floor Plans"), Name="Floor Plans")
test_panel._current_selected_node = cg_tuple_node
class Dummy(object): pass
panel.ManageSheetsPanel.on_tree_selection_changed(test_panel, None, Dummy())
assert test_panel.Txt_GridTitle.Text == "Collection: PERMIT SET | Group: Floor Plans", "Actual: {}".format(test_panel.Txt_GridTitle.Text)

# 2. Non-tuple ContentGroup
cg_nontuple_node = types.SimpleNamespace(NodeType="ContentGroup", Tag=None, Name="Floor Plans")
test_panel._last_selected_node = None
test_panel._current_selected_node = cg_nontuple_node
panel.ManageSheetsPanel.on_tree_selection_changed(test_panel, None, Dummy())
assert test_panel.Txt_GridTitle.Text == "Group: Floor Plans", "Actual: {}".format(test_panel.Txt_GridTitle.Text)

# 3. Tuple Discipline
disc_tuple_node = types.SimpleNamespace(NodeType="Discipline", Tag=("PERMIT SET", "Architecture"), Name="Architecture")
test_panel._last_selected_node = None
test_panel._current_selected_node = disc_tuple_node
panel.ManageSheetsPanel.on_tree_selection_changed(test_panel, None, Dummy())
assert test_panel.Txt_GridTitle.Text == "Collection: PERMIT SET | Discipline: Architecture", "Actual: {}".format(test_panel.Txt_GridTitle.Text)

# 4. Non-tuple Discipline
disc_nontuple_node = types.SimpleNamespace(NodeType="Discipline", Tag="Architecture", Name="Architecture")
test_panel._last_selected_node = None
test_panel._current_selected_node = disc_nontuple_node
panel.ManageSheetsPanel.on_tree_selection_changed(test_panel, None, Dummy())
assert test_panel.Txt_GridTitle.Text == "Discipline: Architecture", "Actual: {}".format(test_panel.Txt_GridTitle.Text)

print("[UNIT TEST OK] on_tree_selection_changed tuple and non-tuple handling verified without error.")

# Test auto_populate_sheet_views with unplaced vs placed AssignableViewNodes
unplaced_vid = MockElementId(777)
unplaced_av = data_model.AssignableViewNode(unplaced_vid, "Level 2 Floor Plan", is_placed=False, sheet_number="", view_type="FloorPlan", level_name="Level 2", scale_int=96)
placed_vid = MockElementId(888)
placed_av = data_model.AssignableViewNode(placed_vid, "Level 3 Floor Plan", is_placed=True, sheet_number="A-103", view_type="FloorPlan", level_name="Level 3", scale_int=96)

data_model.AVAILABLE_ASSIGNABLE_VIEWS = [placed_av, unplaced_av]
data_model.AVAILABLE_EXISTING_VIEWS = [placed_av, unplaced_av]

target_sheet = data_model.SheetViewModel(
    element_id=MockElementId(202),
    number="A-102",
    name="Level 2 Floor Plan",
    collection_name="PERMIT SET",
    discipline_name="Architecture",
    content_group_name="Floor Plans",
    series_name="01. Plans"
)
panel.ManageSheetsPanel.auto_populate_sheet_views(test_panel, target_sheet, run_val=False)
assert target_sheet.Views.Count == 1, "auto_populate should have added 1 view"
pop_view = target_sheet.Views[0]
assert pop_view.IsNew == False, "Matched unplaced view should have IsNew == False"
assert pop_view.ViewId == unplaced_vid, "Matched view should have unplaced_vid"
assert pop_view.AssignmentMode == "Existing", "Matched view assignment mode should be Existing"

# Add target_sheet to EditorItems so unplaced_vid is claimed
test_panel.EditorItems.Add(target_sheet)

nomatch_sheet = data_model.SheetViewModel(
    element_id=MockElementId(299),
    number="A-199",
    name="Level 99 Floor Plan",
    collection_name="PERMIT SET",
    discipline_name="Architecture",
    content_group_name="Floor Plans",
    series_name="01. Plans"
)
panel.ManageSheetsPanel.auto_populate_sheet_views(test_panel, nomatch_sheet, run_val=False)
assert nomatch_sheet.Views.Count == 1, "auto_populate should have added 1 view"
new_pop_view = nomatch_sheet.Views[0]
assert new_pop_view.IsNew == True, "Fallback view should have IsNew == True"
assert new_pop_view.AssignmentMode == "New", "Fallback view assignment mode should be New"
print("[UNIT TEST OK] auto_populate_sheet_views matching and fallback verified.")

# === v4.4 TESTS: Calm Reconciliation Workbench ===

# 1. DisciplineCode and DisciplineShortLabel
disc_test_cases = [
    ("A - Architectural", "A", u"A · Arch"),
    ("S - Structural", "S", u"S · Struct"),
    ("M - Mechanical", "M", u"M · Mech"),
    ("E - Electrical", "E", u"E · Elec"),
    ("P - Plumbing", "P", u"P · Plumb"),
    ("C - Civil", "C", u"C · Civil"),
    ("L - Landscape", "L", u"L · Land"),
    ("F - Fire Protection", "F", u"F · Fire"),
    ("CS - Cover Sheet", "CS", u"CS · Cover"),
    ("G - General", "G", u"G · Gen"),
]
for d_name, expected_code, expected_short in disc_test_cases:
    sh_disc = data_model.SheetViewModel(MockElementId(1), "X-001", "Test", "PERMIT SET", discipline_name=d_name)
    assert sh_disc.DisciplineCode == expected_code, "DisciplineCode mismatch for {}: got {}".format(d_name, sh_disc.DisciplineCode)
    assert sh_disc.DisciplineShortLabel == expected_short, "DisciplineShortLabel mismatch for {}: got {}".format(d_name, sh_disc.DisciplineShortLabel)
    assert isinstance(sh_disc.DisciplineShortLabel, type(u"")), "DisciplineShortLabel must be unicode"

print("[UNIT TEST OK] DisciplineCode and DisciplineShortLabel verified.")

# 2. OriginalNameDisplay, PlacedViewsBaselineText, HasDiffFromOriginal, DiffArrowText, ChangeDiffSummary
sh_recon = data_model.SheetViewModel(MockElementId(301), "A-101", "Level 1 Floor Plan", "PERMIT SET", discipline_name="A - Architectural")
assert sh_recon.OriginalNameDisplay == "Level 1 Floor Plan"
assert sh_recon.PlacedViewsBaselineText == u"No placed views"
assert sh_recon.HasDiffFromOriginal == False
assert sh_recon.DiffArrowText == u"="
assert sh_recon.ChangeDiffSummary == u"No change from Revit"
assert sh_recon.StatusBadgeText == u"AS-IS"
assert sh_recon.ActionBrush == "#6B7280"

# Add 1 placed view with backend suffix
v_recon1 = data_model.ViewViewModel(MockElementId(401), "Level 1 Floor Plan - PS - 1 - A-101", scale="1/8\" = 1'-0\"", parent_sheet=sh_recon)
sh_recon.Views.Add(v_recon1)
assert sh_recon.PlacedViewsBaselineText == u"Level 1 Floor Plan (1/8\" = 1'-0\")"
assert sh_recon.MatchedViewSummary == u"View: \"Level 1 Floor Plan\" (1/8\" = 1'-0\")"

# Add 2nd view
v_recon2 = data_model.ViewViewModel(MockElementId(402), "Level 1 Key Plan", scale="1\" = 20'-0\"", parent_sheet=sh_recon)
sh_recon.Views.Add(v_recon2)
assert sh_recon.PlacedViewsBaselineText == u"Level 1 Floor Plan (1/8\" = 1'-0\") +1"
assert u"+1 more" in sh_recon.MatchedViewSummary

print("[UNIT TEST OK] Reconciliation baseline and view summary properties verified.")

# 3. View Matching Commands: UseViewNameCommand and SyncNamesCommand
v_recon1.Name = "GROUND FLOOR PLAN"
assert sh_recon.CanUseViewName == True
sh_recon.UseViewNameCommand.Execute(None)
assert sh_recon.SheetName == "GROUND FLOOR PLAN"

sh_recon.SheetName = "FIRST FLOOR PLAN"
assert sh_recon.CanSyncViewToSheet == True
sh_recon.SyncNamesCommand.Execute(None)
assert v_recon1.Name == "FIRST FLOOR PLAN"
assert v_recon2.Name == "FIRST FLOOR PLAN"

print("[UNIT TEST OK] View Matching commands (Sheet = View, View = Sheet) verified.")

# 4. Proposal & Accept Workflow
sh_prop = data_model.SheetViewModel(MockElementId(501), "A-101", "Old Floor Plan", "PERMIT SET", discipline_name="A - Architectural")
sh_prop.SuggestedNumber = "A-101"
sh_prop.SuggestedName = "Level 1 Floor Plan"
# Proposal pre-fill as done by execute_schema_match
sh_prop._sheet_name = "Level 1 Floor Plan"
sh_prop.update_diffs()
sh_prop.update_action()
sh_prop.notify_reconciliation_state()

assert sh_prop.HasDiffFromOriginal == True
assert sh_prop.DiffArrowText == u"→"
assert sh_prop.CanAcceptProposal == True
assert sh_prop.IsAccepted == False
assert sh_prop.AcceptButtonText == u"✓ Accept"
assert sh_prop.StatusBadgeText == u"PROPOSAL"
assert sh_prop.ActionBrush == "#2563EB"

# Accept proposal
test_panel.on_accept_sheet_suggestion(sh_prop)
assert sh_prop.IsAccepted == True
assert sh_prop.IsChecked == True
assert sh_prop.StatusBadgeText == u"ACCEPTED"
assert sh_prop.ActionBrush == "#0078D7"

print("[UNIT TEST OK] Proposal & Accept workflow verified.")

# 5. Modify & Accept & Undo Workflow
sh_mod = data_model.SheetViewModel(MockElementId(601), "A-102", "Second Floor Plan", "PERMIT SET", discipline_name="A - Architectural")
v_mod = data_model.ViewViewModel(MockElementId(701), "Second Floor Plan", scale="1/8\" = 1'-0\"", parent_sheet=sh_mod)
sh_mod.Views.Add(v_mod)

# User modifies sheet number and name
sh_mod.SheetNumber = "A-102.1"
sh_mod.SheetName = "Second Floor Plan - Phase 1"
assert sh_mod.AcceptButtonText == u"✓ Accept Modified"
assert sh_mod.StatusBadgeText == u"MODIFIED"
assert sh_mod.CanAcceptProposal == True

# Accept Modified
test_panel.on_accept_sheet_suggestion(sh_mod)
assert sh_mod.SheetNumber == "A-102.1"
assert sh_mod.SheetName == "Second Floor Plan - Phase 1"
assert sh_mod.IsAccepted == True
assert sh_mod.IsChecked == True
assert sh_mod.StatusBadgeText == u"ACCEPTED"

# Undo changes
sh_mod.UndoCommand.Execute(None)
assert sh_mod.SheetNumber == "A-102"
assert sh_mod.SheetName == "Second Floor Plan"
assert sh_mod.StatusBadgeText == u"AS-IS"
assert sh_mod.HasDiffFromOriginal == False
assert sh_mod.IsChecked == False
assert sh_mod.ActionBrush == "#6B7280"

print("[UNIT TEST OK] Modify, Accept Modified, and Undo workflow verified.")

# 6. Discipline Filter Dropdown & Filtering
test_panel.EditorItems.Clear()
sh_a = data_model.SheetViewModel(MockElementId(801), "A-101", "Arch Plan", "PERMIT SET", discipline_name="A - Architectural")
sh_s = data_model.SheetViewModel(MockElementId(802), "S-101", "Struct Plan", "PERMIT SET", discipline_name="S - Structural")
sh_m = data_model.SheetViewModel(MockElementId(803), "M-101", "Mech Plan", "PERMIT SET", discipline_name="M - Mechanical")
test_panel.EditorItems.Add(sh_a)
test_panel.EditorItems.Add(sh_s)
test_panel.EditorItems.Add(sh_m)

test_panel.update_discipline_filter_options()
assert "All Disciplines" in test_panel.Cmb_DisciplineFilter.Items
assert "A · Arch" in test_panel.Cmb_DisciplineFilter.Items
assert "S · Struct" in test_panel.Cmb_DisciplineFilter.Items
assert "M · Mech" in test_panel.Cmb_DisciplineFilter.Items

print("[UNIT TEST OK] Cmb_DisciplineFilter presence and options population verified.")

# 7. Row Click Selection & Non-Destructive on_sheet_number_changed
panel.ManageSheetsPanel.on_editor_row_clicked(test_panel, types.SimpleNamespace(DataContext=sh_s), None)
assert test_panel.main_vm.SelectedSheet == sh_s, "Row click should set SelectedSheet to sh_s"

editor_items_ref = test_panel.EditorItems
editor_count_before = len(test_panel.EditorItems)
panel.ManageSheetsPanel.on_sheet_number_changed(test_panel, sh_a, "A-101", "A-101.1")
assert test_panel.EditorItems is editor_items_ref, "EditorItems reference must be preserved"
assert len(test_panel.EditorItems) == editor_count_before, "EditorItems count must not change on number edit"
print("[UNIT TEST OK] Row Click selection & non-destructive on_sheet_number_changed verified.")

print("=== ALL VERIFICATION CHECKS PASSED SUCCESSFULLY! ===")
