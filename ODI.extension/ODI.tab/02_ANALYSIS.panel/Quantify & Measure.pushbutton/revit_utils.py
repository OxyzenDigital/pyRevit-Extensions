# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import UnitUtils, LabelUtils

# Try to import UIThemeManager (Revit 2024+)
try:
    from Autodesk.Revit.UI import UIThemeManager, UITheme
    HAS_THEME = True
except ImportError:
    HAS_THEME = False

MEASURABLE_NAMES = {"Area", "Volume", "Length", "Perimeter", "Width", "Thickness", "Height", "Diameter", "Cut Length"}

def get_id(element_id):
    """Safe method to get integer ID for 2023/2024+ compatibility."""
    if hasattr(element_id, "Value"):
        return element_id.Value
    return element_id.IntegerValue

def is_dark_theme():
    """Checks if Revit is currently using the Dark Theme."""
    if HAS_THEME:
        try:
            if UIThemeManager.CurrentTheme == UITheme.Dark:
                return True
        except: pass
    return False

def get_display_val_and_label(param, doc):
    """Gets value converted to Project Units and extracts Unit Label."""
    internal_val = param.AsDouble()
    
    val = internal_val
    label = ""
    
    # Attempt to use Revit 2022+ ForgeTypeId / FormatOptions
    try:
        # 1. Get Spec (DataType)
        spec_id = None
        if hasattr(param.Definition, "GetDataType"): # 2022+
            spec_id = param.Definition.GetDataType()
        
        if spec_id and UnitUtils.IsMeasurableSpec(spec_id):
            # 2. Get Project Unit Settings for this Spec
            units = doc.GetUnits()
            format_opts = units.GetFormatOptions(spec_id)
            
            # 3. Convert Value
            unit_id = format_opts.GetUnitTypeId()
            val = UnitUtils.ConvertFromInternalUnits(internal_val, unit_id)
            
            # 4. Get Symbol
            symbol_id = format_opts.GetSymbolTypeId()
            if not symbol_id.Empty():
                label = LabelUtils.GetLabelForSymbol(symbol_id)
            else:
                # Fallback: Map common units if symbol is hidden in project settings
                u_label = LabelUtils.GetLabelForUnit(unit_id)
                if "Feet" in u_label: label = "ft"
                elif "Meters" in u_label: label = "m"
                elif "Inches" in u_label: label = "in"
                elif "Millimeters" in u_label: label = "mm"
                elif "Square Feet" in u_label: label = "SF"
                elif "Square Meters" in u_label: label = "m²"
                elif "Cubic Feet" in u_label: label = "CF"
                elif "Cubic Meters" in u_label: label = "m³"
                elif "Cubic Yards" in u_label: label = "CY"
                
    except Exception:
        # Fallback for older Revit versions or non-measurable specs
        val_str = param.AsValueString()
        if val_str:
            parts = val_str.strip().split(' ')
            if len(parts) > 1:
                candidate = parts[-1]
                if any(c.isalpha() for c in candidate) or any(s in candidate for s in ['°', '³', '²']):
                    label = candidate
                    
    return val, label