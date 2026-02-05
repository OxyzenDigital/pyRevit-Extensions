# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import BuiltInParameter

class WallCMUCalculator:
    def __init__(self):
        self.name = "Wall Material Calculator"
        self.target_category = "Walls"
        self.options = {
            "Standard CMU (8x8x16)": 1.125, # 1 / 0.89 sqft
            "Half-High CMU (4x8x16)": 2.25,
            "Jumbo CMU (8x12x16)": 1.125,
            "Modular Brick (4x2.6x8)": 6.75,
            "Utility Brick (4x4x12)": 3.0
        }
        self.default_setting = "Standard CMU (8x8x16)"

    def calculate(self, element, setting_key):
        """Calculates material count based on Wall Area."""
        # Get Area (Internal Units = Sq Ft)
        # Try BuiltInParameter first (Robust)
        p_area = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
        if not p_area:
            # Fallback to name lookup
            p_area = element.LookupParameter("Area")
        
        if p_area:
            area = p_area.AsDouble()
            factor = self.options.get(setting_key, 1.125)
            val = area * factor
            return "{:,.0f} units".format(val)
        return "-"