# -*- coding: utf-8 -*-
import os
import json
import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

from pyrevit import forms, script
from System.Windows.Media import SolidColorBrush, Color as WpfColor
from data_model import ViewModelBase
from revit_utils import is_dark_theme

SETTINGS_FILE = os.path.join(os.path.dirname(__file__), 'calculation_settings.json')

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

class MaterialVM(ViewModelBase):
    def __init__(self, data):
        ViewModelBase.__init__(self)
        self.MaterialId = data.get("materialId", "")
        self.Name = data.get("name", "Unnamed Material")
        self.Units = data.get("units", "")
        
        self.CalculationItems = []
        for item in data.get("calculationItems", []):
            self.CalculationItems.append(CalculationItemVM(item))
            
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

class CategoryVM(ViewModelBase):
    def __init__(self, data):
        ViewModelBase.__init__(self)
        self.CategoryId = data.get("categoryId", "")
        self.Name = data.get("name", "Unnamed Category")
        
        self.Materials = []
        for mat in data.get("materials", []):
            self.Materials.append(MaterialVM(mat))

    def to_dict(self):
        return {
            "categoryId": self.CategoryId,
            "name": self.Name,
            "materials": [m.to_dict() for m in self.Materials]
        }

class SettingsViewModel(ViewModelBase):
    def __init__(self):
        ViewModelBase.__init__(self)
        self.Categories = []
        self._selected_material = None
        self.load_data()

    def load_data(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    data = json.load(f)
                    for cat in data.get("categories", []):
                        self.Categories.append(CategoryVM(cat))
            except Exception as e:
                print("Error loading settings: {}".format(e))

    @property
    def SelectedMaterial(self):
        return self._selected_material

    @SelectedMaterial.setter
    def SelectedMaterial(self, val):
        self._selected_material = val
        self.OnPropertyChanged("SelectedMaterial")

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
        xaml_file = os.path.join(os.path.dirname(__file__), 'settings_ui.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        
        self.ViewModel = SettingsViewModel()
        self.DataContext = self.ViewModel
        
        self.HeaderDrag.MouseLeftButtonDown += self.drag_window
        self.Btn_Close.Click += lambda s, a: self.Close()
        self.Btn_Cancel.Click += lambda s, a: self.Close()
        self.Btn_Save.Click += self.save_click
        self.tvCategories.SelectedItemChanged += self.tree_selection_changed
        
        self.apply_revit_theme()

    def drag_window(self, sender, args):
        self.DragMove()

    def tree_selection_changed(self, sender, args):
        selected = self.tvCategories.SelectedItem
        if isinstance(selected, MaterialVM):
            self.ViewModel.SelectedMaterial = selected
        else:
            self.ViewModel.SelectedMaterial = None

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