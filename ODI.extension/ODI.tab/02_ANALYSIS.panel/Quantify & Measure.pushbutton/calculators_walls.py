# -*- coding: utf-8 -*-
import os
import json
from Autodesk.Revit.DB import BuiltInParameter

DEFAULT_SETTINGS_FILE = os.path.join(os.path.dirname(__file__), 'calculation_settings.json')
USER_SETTINGS_DIR = os.path.join(os.environ.get('APPDATA', ''), 'OXYZEN_Digital', 'QuantifyAndMeasure')
SETTINGS_FILE = os.path.join(USER_SETTINGS_DIR, 'calculation_settings.json')

class WallCMUCalculator:
    def __init__(self):
        self.name = "Wall Material Calculator"
        self.target_category = "Walls"
        self.options = {}
        self.default_setting = ""
        self.load_settings()

    def load_settings(self):
        cfg_path = SETTINGS_FILE if os.path.exists(SETTINGS_FILE) else DEFAULT_SETTINGS_FILE
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, 'r') as f:
                    data = json.load(f)
                    for cat in data.get("categories", []):
                        if cat.get("name") == "Walls":
                            for grp in cat.get("groups", []):
                                grp_name = grp.get("name", "")
                                for typ in grp.get("types", []):
                                    name = "{} - {}".format(grp_name, typ.get("name", ""))
                                    # Find primary factor (first item or specific key)
                                    factor = 1.0
                                    for item in typ.get("calculationItems", []):
                                        if "per_sf" in item.get("itemId", ""):
                                            factor = float(item.get("value", 1.0))
                                            break
                                    self.options[name] = factor
            except: pass
        
        if self.options:
            self.default_setting = sorted(self.options.keys())[0]
        else:
            self.default_setting = "No Data"

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