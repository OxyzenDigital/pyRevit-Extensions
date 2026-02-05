# -*- coding: utf-8 -*-
import os
import json
import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Input import ICommand

from pyrevit import forms, script
from System.Windows.Media import SolidColorBrush, Color as WpfColor
from data_model import ViewModelBase
from revit_utils import is_dark_theme

SETTINGS_FILE = os.path.join(os.path.dirname(__file__), 'calculation_settings.json')

class RelayCommand(ICommand):
    def __init__(self, action):
        self.action = action
        self.events = []
    def add_CanExecuteChanged(self, handler): self.events.append(handler)
    def remove_CanExecuteChanged(self, handler): self.events.remove(handler)
    def CanExecute(self, parameter): return True
    def Execute(self, parameter): self.action(parameter)

class CalculationItemVM(ViewModelBase):
    def __init__(self, data):
        ViewModelBase.__init__(self)
        self.ItemId = data.get("itemId", "")
        self.Label = data.get("label", "Unknown")
        self.Type = data.get("type", "input")
        self._value = data.get("value", 0)

    @property
    def Value(self):
        return self._value

    @Value.setter
    def Value(self, val):
        self._value = val
        self.OnPropertyChanged("Value")

    def to_dict(self):
        return {
            "itemId": self.ItemId,
            "label": self.Label,
            "type": self.Type,
            "value": self.Value
        }

class MaterialTypeVM(ViewModelBase):
    def __init__(self, data):
        ViewModelBase.__init__(self)
        self.MaterialId = data.get("materialId", "")
        self.Name = data.get("name", "Unnamed Type")
        self.Units = data.get("units", "")
        
        self.CalculationItems = ObservableCollection[CalculationItemVM]()
        for item in data.get("calculationItems", []):
            self.CalculationItems.Add(CalculationItemVM(item))
            
        self._is_selected = False

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, val):
        self._is_selected = val
        self.OnPropertyChanged("IsSelected")

    def to_dict(self):
        return {
            "materialId": self.MaterialId,
            "name": self.Name,
            "units": self.Units,
            "calculationItems": [i.to_dict() for i in self.CalculationItems]
        }

class MaterialGroupVM(ViewModelBase):
    def __init__(self, data):
        ViewModelBase.__init__(self)
        self.Name = data.get("name", "Unnamed Group")
        self._is_expanded = True
        self._is_selected = False
        
        self.Types = ObservableCollection[MaterialTypeVM]()
        for t in data.get("types", []):
            self.Types.Add(MaterialTypeVM(t))

    @property
    def IsExpanded(self): return self._is_expanded
    @IsExpanded.setter
    def IsExpanded(self, val): self._is_expanded = val; self.OnPropertyChanged("IsExpanded")

    @property
    def IsSelected(self): return self._is_selected
    @IsSelected.setter
    def IsSelected(self, val): self._is_selected = val; self.OnPropertyChanged("IsSelected")

    def to_dict(self):
        return { "name": self.Name, "types": [t.to_dict() for t in self.Types] }

class CategoryVM(ViewModelBase):
    def __init__(self, data):
        ViewModelBase.__init__(self)
        self.CategoryId = data.get("categoryId", "")
        self.Name = data.get("name", "Unnamed Category")
        self._is_expanded = True
        self._is_selected = False
        
        self.Groups = ObservableCollection[MaterialGroupVM]()
        for g in data.get("groups", []):
            self.Groups.Add(MaterialGroupVM(g))
            
    @property
    def IsExpanded(self):
        return self._is_expanded

    @IsExpanded.setter
    def IsExpanded(self, val):
        self._is_expanded = val
        self.OnPropertyChanged("IsExpanded")

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, val):
        self._is_selected = val
        self.OnPropertyChanged("IsSelected")

    def to_dict(self):
        return {
            "categoryId": self.CategoryId,
            "name": self.Name,
            "groups": [g.to_dict() for g in self.Groups]
        }

class SettingsViewModel(ViewModelBase):
    def __init__(self):
        ViewModelBase.__init__(self)
        self.Categories = ObservableCollection[CategoryVM]()
        self._selected_item = None # Can be Category, Group, or Type
        self.RemoveFieldCommand = RelayCommand(self.remove_calculation_item)
        self.load_data()

    def load_data(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    data = json.load(f)
                    for cat in data.get("categories", []):
                        self.Categories.Add(CategoryVM(cat))
            except Exception as e:
                print("Error loading settings: {}".format(e))
        
        # Initialize default category if empty
        if self.Categories.Count == 0:
            self.populate_defaults()

    @property
    def SelectedItem(self):
        return self._selected_item

    @SelectedItem.setter
    def SelectedItem(self, val):
        self._selected_item = val
        self.OnPropertyChanged("SelectedItem")
        self.OnPropertyChanged("SelectedType") # For UI binding

    @property
    def SelectedType(self):
        if isinstance(self._selected_item, MaterialTypeVM):
            return self._selected_item
        return None

    def populate_defaults(self):
        """Populates comprehensive default settings for various categories."""
        def item(id, label, val):
            return {"itemId": id, "label": label, "type": "input", "value": val}

        # --- Walls ---
        walls = CategoryVM({"name": "Walls", "groups": []})
        
        masonry = MaterialGroupVM({"name": "Masonry", "types": []})
        masonry.Types.Add(MaterialTypeVM({
            "name": "CMU 8x8x16", "units": "Block",
            "calculationItems": [
                item("face_area", "Face Area (SF)", 0.89),
                item("waste", "Waste Factor", 1.05),
                item("mortar", "Mortar (CF/Block)", 0.03)
            ]
        }))
        masonry.Types.Add(MaterialTypeVM({
            "name": "Brick Standard", "units": "Brick",
            "calculationItems": [
                item("per_sf", "Bricks per SF", 6.75),
                item("waste", "Waste Factor", 1.05),
                item("mortar", "Mortar (CF/Brick)", 0.01)
            ]
        }))
        walls.Groups.Add(masonry)

        framing = MaterialGroupVM({"name": "Framing", "types": []})
        framing.Types.Add(MaterialTypeVM({
            "name": "Metal Studs 16in OC", "units": "Stud",
            "calculationItems": [
                item("spacing", "Spacing (in)", 16.0),
                item("tracks", "Tracks (Count)", 2.0),
                item("waste", "Waste Factor", 1.10)
            ]
        }))
        framing.Types.Add(MaterialTypeVM({
            "name": "Wood Studs 16in OC", "units": "Stud",
            "calculationItems": [
                item("spacing", "Spacing (in)", 16.0),
                item("plates", "Plates (Count)", 3.0),
                item("waste", "Waste Factor", 1.10)
            ]
        }))
        walls.Groups.Add(framing)
        
        finishes = MaterialGroupVM({"name": "Finishes", "types": []})
        finishes.Types.Add(MaterialTypeVM({
            "name": "Paint (Interior)", "units": "Gallon",
            "calculationItems": [
                item("coverage", "Coverage (SF/Gal)", 350.0),
                item("coats", "Coats", 2.0)
            ]
        }))
        walls.Groups.Add(finishes)
        self.Categories.Add(walls)

        # --- Floors ---
        floors = CategoryVM({"name": "Floors", "groups": []})
        tiling = MaterialGroupVM({"name": "Tiling", "types": []})
        tiling.Types.Add(MaterialTypeVM({
            "name": "Ceramic 12x12", "units": "Tile",
            "calculationItems": [
                item("area", "Tile Area (SF)", 1.0),
                item("waste", "Waste Factor", 1.10),
                item("grout", "Grout Width (in)", 0.25)
            ]
        }))
        floors.Groups.Add(tiling)
        
        carpet = MaterialGroupVM({"name": "Carpet", "types": []})
        carpet.Types.Add(MaterialTypeVM({
            "name": "Broadloom", "units": "SY",
            "calculationItems": [
                item("waste", "Waste Factor", 1.15),
                item("glue", "Glue Coverage (SF/Gal)", 100.0)
            ]
        }))
        floors.Groups.Add(carpet)
        self.Categories.Add(floors)
        
        # --- Ceilings ---
        ceilings = CategoryVM({"name": "Ceilings", "groups": []})
        act = MaterialGroupVM({"name": "ACT", "types": []})
        act.Types.Add(MaterialTypeVM({
            "name": "2x2 Grid", "units": "Tile",
            "calculationItems": [item("tile_area", "Tile Area (SF)", 4.0), item("waste", "Waste Factor", 1.10)]
        }))
        ceilings.Groups.Add(act)
        self.Categories.Add(ceilings)

    def add_group(self):
        if isinstance(self.SelectedItem, CategoryVM):
            new_group = MaterialGroupVM({"name": "New Group", "types": []})
            self.SelectedItem.Groups.Add(new_group)
            new_group.IsSelected = True

    def add_type(self):
        target_group = None
        if isinstance(self.SelectedItem, MaterialGroupVM):
            target_group = self.SelectedItem
        elif isinstance(self.SelectedItem, MaterialTypeVM):
            # Find parent group? Hard without parent ref.
            # Iterate to find parent
            for cat in self.Categories:
                for grp in cat.Groups:
                    if self.SelectedItem in grp.Types:
                        target_group = grp
                        break
        
        if target_group:
            new_type = MaterialTypeVM({
                "name": "New Type", 
                "units": "Unit", 
                "calculationItems": [{"itemId": "factor", "label": "Factor", "type": "input", "value": 1.0}]
            })
            target_group.Types.Add(new_type)
            new_type.IsSelected = True

    def delete_item(self):
        sel = self.SelectedItem
        if not sel: return
        
        # Try to delete Group from Category
        for cat in self.Categories:
            if sel in cat.Groups:
                if forms.alert("Delete Group '{}'?".format(sel.Name), yes=True, no=True):
                    cat.Groups.Remove(sel)
                return
            # Try to delete Type from Group
            for grp in cat.Groups:
                if sel in grp.Types:
                    if forms.alert("Delete Type '{}'?".format(sel.Name), yes=True, no=True):
                        grp.Types.Remove(sel)
                    return

    def add_calculation_item(self):
        if self.SelectedType:
            name = forms.ask_for_string(prompt="Enter Field Name (e.g. 'Waste Factor'):", title="Add Field")
            if name:
                import re
                safe_id = re.sub(r'[^a-zA-Z0-9]', '_', name).lower()
                new_item = CalculationItemVM({
                    "itemId": safe_id,
                    "label": name,
                    "type": "input",
                    "value": 0.0
                })
                self.SelectedType.CalculationItems.Add(new_item)

    def remove_calculation_item(self, item):
        if self.SelectedType and item in self.SelectedType.CalculationItems:
            self.SelectedType.CalculationItems.Remove(item)

    def save_data(self):
        data = {
            "categories": [c.to_dict() for c in self.Categories]
        }
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(data, f, indent=2)
            return True
        except Exception as e:
            print("Error saving settings: {}".format(e))
            return False

class SettingsWindow(forms.WPFWindow):
    def __init__(self):
        xaml_file = os.path.join(os.path.dirname(__file__), 'Settings_ui.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.ViewModel = SettingsViewModel()
        self.DataContext = self.ViewModel
        
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_Close.Click += lambda s, a: self.Close()
        self.Btn_Cancel.Click += lambda s, a: self.Close()
        self.Btn_Save.Click += self.save_click
        
        self.Btn_AddGroup.Click += lambda s, a: self.ViewModel.add_group()
        self.Btn_AddType.Click += lambda s, a: self.ViewModel.add_type()
        self.Btn_Delete.Click += lambda s, a: self.ViewModel.delete_item()
        self.Btn_AddField.Click += lambda s, a: self.ViewModel.add_calculation_item()
        
        self.tvCategories.SelectedItemChanged += self.tree_selection_changed
        
        self.apply_revit_theme()

    def drag_window(self, sender, args):
        self.DragMove()

    def tree_selection_changed(self, sender, args):
        self.ViewModel.SelectedItem = self.tvCategories.SelectedItem

    def save_click(self, sender, args):
        if self.ViewModel.save_data():
            self.Close()

    def apply_revit_theme(self):
        if is_dark_theme():
            res = self.Resources
            res["WindowBrush"] = SolidColorBrush(WpfColor.FromRgb(31, 41, 55))
            res["ControlBrush"] = SolidColorBrush(WpfColor.FromRgb(17, 24, 39))
            res["TextBrush"] = SolidColorBrush(WpfColor.FromRgb(249, 250, 251))
            res["TextLightBrush"] = SolidColorBrush(WpfColor.FromRgb(156, 163, 175))
            res["BorderBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 99))
            res["ButtonBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81))
            res["HoverBrush"] = SolidColorBrush(WpfColor.FromRgb(75, 85, 99))
            res["FooterBrush"] = SolidColorBrush(WpfColor.FromRgb(17, 24, 39))
            res["AccentBrush"] = SolidColorBrush(WpfColor.FromRgb(59, 130, 246))
            res["SelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(30, 58, 138))  # #1E3A8A (Blue-900)
            res["SelectionBorderBrush"] = SolidColorBrush(WpfColor.FromRgb(59, 130, 246)) # #3B82F6 (Blue-500)
            res["SelectionTextBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            res["InactiveSelectionBrush"] = SolidColorBrush(WpfColor.FromRgb(55, 65, 81)) # #374151 (Gray-700)