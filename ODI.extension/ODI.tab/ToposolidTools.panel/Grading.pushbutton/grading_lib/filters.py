# grading_lib/filters.py
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.DB import BuiltInCategory, ModelLine, ModelCurve

class StakeFilter(ISelectionFilter):
    """
    Strictly allows only specific API Categories for Stakes.
    Uses BuiltInCategory constants for 100% accuracy.
    """
    def __init__(self):
        # We define the Allowed Categories using strictly BuiltInCategory
        self.allowed_categories = [
            int(BuiltInCategory.OST_Site),                # Standard Site families
            int(BuiltInCategory.OST_GenericModel),        # Common for custom families
            int(BuiltInCategory.OST_Entourage),           # People/Vehicles
            int(BuiltInCategory.OST_Planting),            # Trees/RPCs
            int(BuiltInCategory.OST_Hardscape),           # Civil 3D/Landscape items
            int(BuiltInCategory.OST_Furniture)            # Just in case
        ]

    def AllowElement(self, elem):
        # 1. Safety Check: Must have a category
        if not elem.Category: return False
        
        # 2. Strict ID Check against our Allowed List
        cat_id = elem.Category.Id.IntegerValue
        
        if cat_id in self.allowed_categories:
            return True
            
        return False

    def AllowReference(self, ref, point):
        return True


class LineFilter(ISelectionFilter):
    """Strictly allows Model Lines or Model Curves."""
    def AllowElement(self, elem):
        # Check if the element class is strictly a Model Curve/Line
        return isinstance(elem, (ModelLine, ModelCurve))

    def AllowReference(self, ref, point):
        return True


class TopoFilter(ISelectionFilter):
    """Allows ONLY Ground elements (Toposolids/Floors/Topo)."""
    def __init__(self):
        self.ground_categories = [
            int(BuiltInCategory.OST_Toposolids), # Revit 2024+
            int(BuiltInCategory.OST_Floors),     # Slabs
            int(BuiltInCategory.OST_Topography)  # Legacy
        ]

    def AllowElement(self, elem):
        if not elem.Category: return False
        
        cat_id = elem.Category.Id.IntegerValue
        if cat_id in self.ground_categories:
            return True
            
        return False
        
    def AllowReference(self, ref, point):
        return True