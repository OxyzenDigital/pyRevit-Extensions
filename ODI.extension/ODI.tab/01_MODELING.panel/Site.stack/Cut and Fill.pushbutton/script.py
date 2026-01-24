# -*- coding: utf-8 -*-
"""
Cut and Fill Tool
Calculates Cut and Fill volumes for Toposolids in the project.
"""
__title__ = "Cut and Fill"
__author__ = "Oxyzen Digital"

import json
from pyrevit import revit, forms, script
from pyrevit import DB
import datetime

doc = revit.doc
output = script.get_output()

# --- Constants ---
# Conversion factors
CF_TO_CY = 1.0 / 27.0
# Density assumption: 1.6 Metric Tonnes per Cubic Meter
# 1 CF = 0.0283168 m3
# Weight (Ton) = Vol (CF) * 0.0283168 * 1.6
CF_TO_MTON = 0.0283168 * 1.6 

# --- Truck & Soil Factors (Texas Standards) ---
# Standard Dump Truck (Tandem/Bobtail) approx 12-14 CY.
# Large Belly Dump/End Dump approx 18-24 CY.
# Using 14 CY as a conservative "Standard Truck".
TRUCK_CAP_CY = 14.0

# Expansion (Swell) Factor: Bank -> Loose
# Common Earth/Clay Swell ~25%
FACTOR_SWELL = 1.25

# Compaction Factor: Loose -> Compacted
# To get 1 CY Compacted, you need ~1.30 CY Loose.
# (Assuming ~20-25% shrinkage from Loose to Compacted)
FACTOR_COMPACTION_REQ = 1.30

class TopoData:
    def __init__(self, element):
        self.Id = element.Id
        self.Name = element.Name
        
        # Design Option
        self.DesignOption = "Main Model"
        if element.DesignOption:
            self.DesignOption = element.DesignOption.Name
            
        # Phase
        self.Phase = "Unknown"
        if element.CreatedPhaseId != DB.ElementId.InvalidElementId:
            phase = doc.GetElement(element.CreatedPhaseId)
            if phase:
                self.Phase = phase.Name

        # Volumes (Internal = CF)
        self.CutCF = 0.0
        self.FillCF = 0.0
        
        # Get Parameters
        # Check for both Toposolid (2024+) and Topography (Legacy) parameter names
        # safely using getattr to avoid AttributeErrors on different Revit versions.
        # Fallback to Display Names if BuiltInParameters fail or return 0.
        
        def get_volume_param(elem, bip_names, display_names):
            # 1. Try BuiltInParameters
            for name in bip_names:
                bip = getattr(DB.BuiltInParameter, name, None)
                if bip:
                    p = elem.get_Parameter(bip)
                    if p and p.HasValue:
                        val = p.AsDouble()
                        if abs(val) > 0.001: return val

            # 2. Try Display Names (LookupParameter)
            for name in display_names:
                p = elem.LookupParameter(name)
                if p and p.HasValue:
                    val = p.AsDouble()
                    if abs(val) > 0.001: return val
            
            return 0.0

        cut_bips = ["TOPOSOLID_CUT_VOLUME", "HOST_VOLUME_CUT_COMPUTED"]
        fill_bips = ["TOPOSOLID_FILL_VOLUME", "HOST_VOLUME_FILL_COMPUTED"]
        
        self.CutCF = get_volume_param(element, cut_bips, ["Cut", "Cut Volume"])
        self.FillCF = get_volume_param(element, fill_bips, ["Fill", "Fill Volume"])
            
        self.NetCF = self.CutCF - self.FillCF # Cut is positive removal, Fill is positive addition. Net = Cut - Fill? Or Net Volume change? Usually Net = Fill - Cut (import required) or Cut - Fill (export available). 
        # Let's denote Net as (Cut - Fill). Positive = Export, Negative = Import.
        
    @property
    def CutCY(self): return self.CutCF * CF_TO_CY
    @property
    def FillCY(self): return self.FillCF * CF_TO_CY
    @property
    def NetCY(self): return self.NetCF * CF_TO_CY
    
    @property
    def CutTon(self): return self.CutCF * CF_TO_MTON
    @property
    def FillTon(self): return self.FillCF * CF_TO_MTON
    @property
    def NetTon(self): return self.NetCF * CF_TO_MTON

    @property
    def CutTrucks(self):
        # Haul Off: Bank Volume * Swell / Truck Cap
        loose_cy = self.CutCY * FACTOR_SWELL
        return loose_cy / TRUCK_CAP_CY

    @property
    def FillTrucks(self):
        # Haul In: Compacted Volume * Requirement Factor / Truck Cap
        required_loose_cy = self.FillCY * FACTOR_COMPACTION_REQ
        return required_loose_cy / TRUCK_CAP_CY

def main():
    # 1. Collect Toposolids
    try:
        # Strict check for Toposolids (Revit 2024+)
        cat = DB.BuiltInCategory.OST_Toposolid
    except AttributeError:
        forms.alert("This tool requires Revit 2024+ and Toposolids.")
        return

    # Collect all Toposolids in the document (Main Model + Design Options)
    elements = DB.FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
        
    if not elements:
        forms.alert("No Toposolids found in the project.")
        return

    # 2. Process Data & Group by Option
    options_map = {}
    
    for el in elements:
        t_data = TopoData(el)
        
        # Filter: Exclude Existing/Ungraded conditions (Zero Cut/Fill)
        if t_data.CutCF < 0.01 and t_data.FillCF < 0.01:
            continue
        
        # Group by (Design Option + Phase)
        opt_name = t_data.DesignOption
        phase_name = t_data.Phase
        key = (opt_name, phase_name)
        
        if key not in options_map:
            options_map[key] = {
                "rows": [],
                "totals": {
                    "cut_cy": 0.0, "cut_cf": 0.0, "cut_ton": 0.0, "cut_trucks": 0.0,
                    "fill_cy": 0.0, "fill_cf": 0.0, "fill_ton": 0.0, "fill_trucks": 0.0
                    # Net calculated later
                }
            }
        
        options_map[key]["rows"].append(t_data)
        
        # Accumulate Totals
        t = options_map[key]["totals"]
        t["cut_cy"] += t_data.CutCY
        t["cut_cf"] += t_data.CutCF
        t["cut_ton"] += t_data.CutTon
        t["cut_trucks"] += t_data.CutTrucks
        
        t["fill_cy"] += t_data.FillCY
        t["fill_cf"] += t_data.FillCF
        t["fill_ton"] += t_data.FillTon
        t["fill_trucks"] += t_data.FillTrucks

    if not options_map:
         forms.alert("No Graded Toposolids found (elements with Cut/Fill > 0).")
         return

    # Calculate Nets
    for opt in options_map:
        t = options_map[opt]["totals"]
        t["net_cy"] = t["cut_cy"] - t["fill_cy"]
        t["net_cf"] = t["cut_cf"] - t["fill_cf"]
        t["net_ton"] = t["cut_ton"] - t["fill_ton"]
        t["net_trucks"] = t["cut_trucks"] - t["fill_trucks"]

    # Sort keys: Main Model first, then Option Name, then Phase
    sorted_keys = sorted(options_map.keys(), key=lambda x: ("" if x[0] == "Main Model" else x[0], x[1]))
    
    # 3. Render HTML
    output.close_others()
    output.resize(1100, 1200)
    
    # Load External CSS
    css_content = ""
    css_file = script.get_bundle_file('style.css')
    if css_file:
        with open(css_file, 'r') as f:
            css_content = f.read()
            output.add_style(css_content)
    
    # --- Block 1: Header ---
    html = ''
    
    # Embed CSS directly to ensure it persists during Printing
    if css_content:
        html += '<style>' + css_content + '</style>'

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    html += '<header>'
    html += '<div><h1>Earthwork Report</h1><div class="meta">Logistics Analysis</div></div>'
    html += '<div class="meta" style="text-align:right;">Generated: {}</div>'.format(now_str)
    html += '</header>'
    
    # Print Header & CSS
    output.print_html(html)
    
    # --- Block 2: Chart ---
    # Note: We print the title in a separate block to avoid unclosed div issues with chart.draw()
    output.print_html('<div class="analysis-section"><div class="option-title">Grading Trends</div></div>')
    
    # Use pyRevit's built-in Charting (Local Resources)
    chart = output.make_line_chart()
    chart.options.title = {'display': True, 'text': 'Net Truck Logistics', 'fontSize': 20}
    chart.options.scales = {'yAxes': [{'ticks': {'beginAtZero': True}}]}
    chart.options.tooltips = {'mode': 'index', 'intersect': False}
    chart.options.maintainAspectRatio = True
    chart.options.aspectRatio = 2
    chart.options.responsive = True
    
    # Inject dummy start/end points (0) to create a "Pile of Dirt" look
    # Multiline labels for Chart.js (List of Lists)
    chart.data.labels = [""] + [[k[0], "({})".format(k[1])] for k in sorted_keys] + [""]
    
    # Calculate Net Data
    net_vals = [options_map[k]["totals"]["net_trucks"] for k in sorted_keys]
    
    ds_export = chart.data.new_dataset('Net Export (Trucks Out)')
    ds_export.data = [0] + [max(0, v) for v in net_vals] + [0]
    ds_export.set_color(231, 76, 60, 0.5) # Red (Semi-transparent)
    ds_export.fill = True
    ds_export.tension = 0.4 # Smooth curve
    
    ds_import = chart.data.new_dataset('Net Import (Trucks In)')
    ds_import.data = [0] + [abs(min(0, v)) for v in net_vals] + [0]
    ds_import.set_color(39, 174, 96, 0.5) # Green (Semi-transparent)
    ds_import.fill = True
    ds_import.tension = 0.4 # Smooth curve

    chart.draw()
    
    # --- Block 3: Table & Footer ---
    # --- The Spreadsheet ---
    # Calculate max NET value for scaling visual bars (Visual Balance)
    max_net_abs = 0.0
    for k in sorted_keys:
        t = options_map[k]["totals"]
        max_net_abs = max(max_net_abs, abs(t["net_trucks"]))
    if max_net_abs == 0: max_net_abs = 1.0

    html = '<div class="analysis-section"><div class="option-title">Detailed Breakdown</div>'
    html += '<div class="table-container">'
    html += '<table><thead><tr>'
    html += '<th>Design Option</th><th>Phase</th>'
    html += '<th>Visual Balance</th>'
    html += '<th class="num">Cut (C.Y.)</th><th class="num">Fill (C.Y.)</th>'
    html += '<th class="num">Net (C.Y.)</th><th class="num">Net Trucks</th>'
    html += '</tr></thead><tbody>'
    
    for key in sorted_keys:
        opt, phase = key
        t = options_map[key]["totals"]
        net_t = t["net_trucks"]
        net_cy = t["net_cy"]
        
        c_style = "color:#7f8c8d;"
        if net_t > 0.1: c_style = "color:#c0392b; font-weight:bold;"
        elif net_t < -0.1: c_style = "color:#27ae60; font-weight:bold;"
        
        # Visual Bar Logic
        # Calculate percentage of the HALF width (50%)
        pct = (abs(net_t) / max_net_abs) * 50.0
        
        bar_html = '<div class="visual-bar-container"><div class="visual-center-line"></div>'
        if net_t > 0:
            bar_html += '<div class="bar-fill cut-bar" style="width:{:.1f}%;"></div>'.format(pct)
        elif net_t < 0:
            bar_html += '<div class="bar-fill fill-bar" style="width:{:.1f}%;"></div>'.format(pct)
        bar_html += '</div>'
        
        html += '<tr><td><strong>{}</strong></td><td>{}</td>'.format(opt, phase)
        html += '<td>{}</td>'.format(bar_html)
        html += '<td class="num">{:,.0f}</td><td class="num">{:,.0f}</td>'.format(t["cut_cy"], t["fill_cy"])
        html += '<td class="num" style="{}">{:,.0f}</td><td class="num" style="{}">{:,.1f}</td>'.format(c_style, abs(net_cy), c_style, abs(net_t))
        html += '</tr>'
        
    html += '</tbody></table></div>'
    html += '</div>' # End analysis-section
    html += '</div>'

    # Disclaimer
    html += '<div class="disclaimer">'
    html += '<strong>Assumptions & Notes:</strong><br>'
    html += '<ul>'
    html += '<li><strong>Truck Capacity:</strong> Standard 14 CY Dump Truck.</li>'
    html += '<li><strong>Cut (Export):</strong> Based on Bank Volume * 1.25 Swell Factor.</li>'
    html += '<li><strong>Fill (Import):</strong> Based on Compacted Volume * 1.30 Requirement Factor.</li>'
    html += '<li><strong>Cost Impact:</strong> "Export Required" implies disposal costs. "Import Required" implies material purchase + haul costs.</li>'
    html += '</ul></div>'
    
    html += '<script>window.scrollTo(0,0);</script>'
    output.print_html(html)

if __name__ == '__main__':
    main()