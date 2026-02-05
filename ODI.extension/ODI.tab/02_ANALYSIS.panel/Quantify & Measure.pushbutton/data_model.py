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
        self._is_expanded = True
        self.Children = []
        self.Type = "Item"
        self._count = 0
        self._value = 0.0
        self._selected_count = 0
        self._selected_value = 0.0
        self.FontWeight = "Normal"
        self.UnitLabel = ""
        self.AllElements = [] # Flat list of element IDs for highlighting
        self.NetworkColor = SolidColorBrush(Colors.White if is_dark_theme() else Colors.Black)
        self._assigned_color_brush = None
        self._revit_color = None

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
    def IsExpanded(self):
        return self._is_expanded

    @IsExpanded.setter
    def IsExpanded(self, value):
        self._is_expanded = value
        self.OnPropertyChanged("IsExpanded")

    @property
    def Count(self):
        if self._selected_count > 0:
            return "{} (Sel: {})".format(self._count, self._selected_count)
        return self._count

    @Count.setter
    def Count(self, val):
        self._count = val
        self.OnPropertyChanged("Count")

    @property
    def Value(self):
        return self._value

    @Value.setter
    def Value(self, val):
        self._value = val
        self.OnPropertyChanged("Value")
        self.OnPropertyChanged("DisplayValue")

    @property
    def SelectedCount(self):
        return self._selected_count

    @SelectedCount.setter
    def SelectedCount(self, val):
        self._selected_count = val
        self.OnPropertyChanged("SelectedCount")
        self.OnPropertyChanged("Count")

    @property
    def SelectedValue(self):
        return self._selected_value

    @SelectedValue.setter
    def SelectedValue(self, val):
        self._selected_value = val
        self.OnPropertyChanged("SelectedValue")
        self.OnPropertyChanged("DisplayValue")

    @property
    def DisplayValue(self):
        base = "{} {}".format(format_value(self.Value), self.UnitLabel).strip()
        if self._selected_value > 0:
            sel = "{} {}".format(format_value(self._selected_value), self.UnitLabel).strip()
            return "{} (Sel: {})".format(base, sel)
        return base

    @property
    def GridRows(self):
        return self.Children

    @property
    def AssignedColorBrush(self):
        return self._assigned_color_brush

    @AssignedColorBrush.setter
    def AssignedColorBrush(self, brush):
        self._assigned_color_brush = brush
        self.OnPropertyChanged("AssignedColorBrush")

    @property
    def RevitColor(self):
        return self._revit_color

    @RevitColor.setter
    def RevitColor(self, val):
        self._revit_color = val
        if val:
            self._assigned_color_brush = SolidColorBrush(WpfColor.FromRgb(val.Red, val.Green, val.Blue))
        else:
            self._assigned_color_brush = None
        self.OnPropertyChanged("RevitColor")
        self.OnPropertyChanged("AssignedColorBrush")

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
        try:
            self.Name = element.Name
        except AttributeError:
            self.Name = "Unnamed Element"
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
        self._assigned_color_brush = None
        self._revit_color = None
    
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

    @property
    def AssignedColorBrush(self):
        return self._assigned_color_brush

    @AssignedColorBrush.setter
    def AssignedColorBrush(self, brush):
        self._assigned_color_brush = brush
        self.OnPropertyChanged("AssignedColorBrush")

    @property
    def RevitColor(self):
        return self._revit_color

    @RevitColor.setter
    def RevitColor(self, val):
        self._revit_color = val
        if val:
            self._assigned_color_brush = SolidColorBrush(WpfColor.FromRgb(val.Red, val.Green, val.Blue))
        else:
            self._assigned_color_brush = None
        self.OnPropertyChanged("RevitColor")
        self.OnPropertyChanged("AssignedColorBrush")

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