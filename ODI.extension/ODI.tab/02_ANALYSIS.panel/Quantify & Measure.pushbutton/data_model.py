# -*- coding: utf-8 -*-
import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Media import SolidColorBrush, Colors, Color as WpfColor
from Autodesk.Revit.DB import BuiltInParameter, Color
from revit_utils import get_id, is_dark_theme

def format_value(val):
    """Formats a float to 2 decimal places."""
    try:
        return "{:.2f}".format(val)
    except:
        return str(val)

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

class NodeBase(ViewModelBase):
    def __init__(self, name):
        ViewModelBase.__init__(self)
        self.Name = name
        self.Id = None
        self._is_checked = False
        self._is_selected = False
        self.IsExpanded = True
        self.Children = []
        self.Type = "Item"
        self.Count = 0
        self.Value = 0.0
        self.FontWeight = "Normal"
        self.UnitLabel = ""
        self.AllElements = [] # Flat list of element IDs for highlighting
        self.NetworkColor = SolidColorBrush(Colors.White if is_dark_theme() else Colors.Black)

    @property
    def IsChecked(self):
        return self._is_checked

    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value
        self.OnPropertyChanged("IsChecked")

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged("IsSelected")

    @property
    def DisplayValue(self):
        return "{} {}".format(format_value(self.Value), self.UnitLabel).strip()

    @property
    def GridRows(self):
        return self.Children

class MeasurementNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self, name)
        self.FontWeight = "Bold"
        self.Type = "Measurement"
        self.NetworkColor = SolidColorBrush(Colors.White if is_dark_theme() else Colors.Black)

class CategoryNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self, name)
        self.FontWeight = "SemiBold"
        self.Type = "Category"
        self.NetworkColor = SolidColorBrush(Colors.LightGray if is_dark_theme() else Colors.Gray)

class FamilyTypeNode(NodeBase):
    def __init__(self, name):
        NodeBase.__init__(self, name)
        self.FontWeight = "Normal"
        self.Type = "Type"
        self.NetworkColor = SolidColorBrush(Colors.LightGray if is_dark_theme() else Colors.Gray)
        self.Instances = []

    @property
    def GridRows(self):
        return self.Instances

class InstanceItem(ViewModelBase):
    """Represents a single row in the DataGrid when a Type is selected."""
    def __init__(self, element, value, unit_label, calculated_val="-"):
        ViewModelBase.__init__(self)
        self.Name = element.Name
        self.Id = get_id(element.Id)
        self.Value = value
        self.UnitLabel = unit_label
        self._calculated_value = calculated_val
        
        # Try to get Family Name for Type column
        fam_name = element.Category.Name if element.Category else "Element"
        p_fam = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
        if p_fam and p_fam.HasValue:
            fam_name = p_fam.AsValueString()
        self.Type = fam_name
        self.Count = 1
        self.Element = element
    
    @property
    def DisplayValue(self):
        return "{} {}".format(format_value(self.Value), self.UnitLabel).strip()
        
    @property
    def CalculatedValue(self):
        return self._calculated_value

    @CalculatedValue.setter
    def CalculatedValue(self, val):
        self._calculated_value = val
        self.OnPropertyChanged("CalculatedValue")

class ColorOption(ViewModelBase):
    def __init__(self, name, r, g, b):
        ViewModelBase.__init__(self)
        self.Name = name
        self.R = r
        self.G = g
        self.B = b
        self.Brush = SolidColorBrush(WpfColor.FromRgb(r, g, b))
        self.RevitColor = Color(r, g, b)

    def __repr__(self):
        return self.Name
        
    def ToString(self):
        return self.Name