# -*- coding: utf-8 -*-
"""
Find Slopes Tool
Calculates the slope percentage between two selected objects.
"""
__title__ = "Find Slopes"
__author__ = "Oxyzen Digital"

from pyrevit import revit, forms, script
from pyrevit import DB
import math

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# --- Unit Helper ---
class UnitHelper:
    @staticmethod
    def get_project_length_unit():
        try:
            # Revit 2021+
            return doc.GetUnits().GetFormatOptions(DB.SpecTypeId.Length).GetUnitTypeId()
        except AttributeError:
            # Fallback for older Revit versions
            return doc.GetUnits().GetFormatOptions(DB.UnitType.UT_Length).DisplayUnits

    @staticmethod
    def from_internal(value_in_internal_units):
        try:
            val = float(value_in_internal_units)
            unit_id = UnitHelper.get_project_length_unit()
            return DB.UnitUtils.ConvertFromInternalUnits(val, unit_id)
        except: 
            return value_in_internal_units

# --- Geometry Helper ---
def get_element_location(element):
    """
    Robustly determines a single point representing the element's location.
    Prioritizes LocationPoint, then Curve Midpoint, then BoundingBox Center.
    """
    if not element: return None
    
    # 1. Try Location Point (Family Instances, etc.)
    if hasattr(element, "Location") and isinstance(element.Location, DB.LocationPoint):
        if element.Location.Point:
            return element.Location.Point
        
    # 2. Try Location Curve (Walls, Pipes, Ducts, etc.) -> Use Midpoint
    if hasattr(element, "Location") and isinstance(element.Location, DB.LocationCurve):
        curve = element.Location.Curve
        if curve:
            # Evaluate midpoint (0.5 normalized parameter)
            return curve.Evaluate(0.5, True)
            
    # 3. Fallback: Bounding Box Center (Floors, Roofs, Solids)
    bbox = element.get_BoundingBox(None)
    if bbox:
        return (bbox.Min + bbox.Max) / 2.0
        
    return None

def main():
    # 1. Get Selection
    selection_ids = uidoc.Selection.GetElementIds()
    
    # 2. Validate Selection Count
    if len(selection_ids) != 2:
        forms.alert("Please select exactly two objects to calculate the slope.", 
                    title="Selection Error", 
                    sub_msg="Current selection count: {}".format(len(selection_ids)))
        return

    # 3. Get Elements and Points
    e1 = doc.GetElement(selection_ids[0])
    e2 = doc.GetElement(selection_ids[1])
    
    p1 = get_element_location(e1)
    p2 = get_element_location(e2)
    
    if not p1 or not p2:
        forms.alert("Could not determine the 3D location for one or both selected objects.", 
                    title="Geometry Error")
        return

    # 4. Calculate Slope
    # Internal Units are Feet
    d_x = p2.X - p1.X
    d_y = p2.Y - p1.Y
    d_z = p2.Z - p1.Z
    
    dist_xy = math.sqrt(d_x**2 + d_y**2)
    
    # 5. Prepare Output
    output.close_others()
    output.resize(1000, 800)
    
    # Results
    disp_xy = UnitHelper.from_internal(dist_xy)
    disp_z = UnitHelper.from_internal(d_z)
    
    # Styling
    style_container = "font-family: 'Segoe UI', sans-serif; padding: 20px; background-color: #f9f9f9; border-radius: 8px; border: 1px solid #e0e0e0; width: 100%; box-sizing: border-box;"
    style_card = "background: white; padding: 20px; border-radius: 6px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center;"
    style_label = "font-size: 24px; text-transform: uppercase; letter-spacing: 1px; color: #888; margin-bottom: 5px;"
    style_big_val = "font-size: 112px; font-weight: 300; color: #0078D7; line-height: 1.0; margin-bottom: 10px;"
    style_sub_val = "font-size: 32px; color: #555; font-weight: 500;"
    style_grid = "display: flex; gap: 15px; margin-bottom: 15px;"
    style_small_card = "flex: 1; background: white; padding: 15px; border-radius: 6px; text-align: center; box-shadow: 0 1px 2px rgba(0,0,0,0.05);"
    style_footer = "font-size: 22px; color: #999; border-top: 1px solid #eee; padding-top: 10px; line-height: 1.6;"

    html_content = '<div style="{}">'.format(style_container)
    
    if dist_xy < 0.001:
        # Vertical Case
        html_content += '<div style="{}"><div style="{}">SLOPE</div><div style="{}">VERTICAL</div></div>'.format(style_card, style_label, style_big_val)
    else:
        slope_pct = (d_z / dist_xy) * 100.0
        angle_deg = math.degrees(math.atan2(d_z, dist_xy))
        
        ratio_str = "-"
        if abs(d_z) > 0.001:
            ratio = abs(dist_xy / d_z)
            ratio_str = "1 : {:.2f}".format(ratio)
            
        html_content += '<div style="{}">'.format(style_card)
        html_content += '<div style="{}">SLOPE PERCENTAGE</div>'.format(style_label)
        html_content += '<div style="{}">{:.2f}%</div>'.format(style_big_val, slope_pct)
        html_content += '<div style="{}">Angle: {:.2f}° &nbsp;&nbsp;<span style="color:#ddd">|</span>&nbsp;&nbsp; Ratio: {}</div>'.format(style_sub_val, angle_deg, ratio_str)
        html_content += '</div>'

    # Measurements Grid
    html_content += '<div style="{}">'.format(style_grid)
    html_content += '<div style="{}"><div style="{}">DISTANCE (XY)</div><div style="font-size: 40px; font-weight: 600; color: #333;">{:.2f}</div></div>'.format(style_small_card, style_label, disp_xy)
    html_content += '<div style="{}"><div style="{}">DELTA (Z)</div><div style="font-size: 40px; font-weight: 600; color: #333;">{:.2f}</div></div>'.format(style_small_card, style_label, disp_z)
    html_content += '</div>'
    
    # Footer
    html_content += '<div style="{}">'.format(style_footer)
    html_content += '<div><strong>Start:</strong> {} (ID: {})</div>'.format(e1.Name, e1.Id)
    html_content += '<div><strong>End:</strong> {} (ID: {})</div>'.format(e2.Name, e2.Id)
    html_content += '</div></div>'
    
    output.print_html(html_content)

if __name__ == '__main__':
    main()
