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
    output.resize(1000, 1200)
    
    # CSS
    # Construct a full HTML document to ensure IE Edge mode is strictly enforced
    html_start = """<!DOCTYPE html>
    <html>
    <head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <style>
        /* Reset & Base */
        * { box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif !important; 
            background-color: #f4f4f4; 
            padding: 40px; 
            color: #222; 
            font-size: 24px !important; 
            zoom: 100%; /* Force reset zoom level */
        }
        .container { max-width: 95%; margin: 0 auto; background: white; padding: 60px; box-shadow: 0 15px 30px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; }
        
        @media print {
            @page { size: letter portrait; margin: 0.5cm; }
            body { 
                background-color: white; 
                padding: 0; 
                font-size: 14pt; 
                zoom: 0.55; 
                -webkit-print-color-adjust: exact; 
                print-color-adjust: exact; 
            }
            .container { box-shadow: none; max-width: 100%; padding: 0; margin: 0; border-radius: 0; }
            .no-print, .action-footer { display: none !important; }
            .analysis-section { page-break-inside: avoid; }
            h1 { font-size: 36pt !important; }
            .btn { display: none !important; }
        }
        
        /* Headers */
        header { 
            background-color: #2c3e50; 
            padding: 50px 60px; 
            margin: -60px -60px 60px -60px; 
            border-bottom: 8px solid #34495e; 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
        }
        @media print { header { margin: 0 0 40px 0; padding: 40px; } }
        h1 { margin: 0; font-weight: 800; font-size: 48px !important; color: white !important; text-transform: uppercase; letter-spacing: 2px; line-height: 1.1; }
        .meta { font-size: 24px !important; color: #bdc3c7 !important; margin-top: 10px; font-weight: 400; }
        
        /* Buttons (Bottom Right) - Styled as Ghost Buttons */
        .action-footer { display: flex; justify-content: flex-end; align-items: center; gap: 20px; margin-top: 80px; padding-top: 40px; border-top: 2px solid #eee; }
        .btn { 
            background-color: white; border: 3px solid #2c3e50; padding: 12px 30px; cursor: pointer; 
            font-size: 20px !important; font-weight: 700; color: #2c3e50 !important;
            text-decoration: none !important; font-family: inherit; border-radius: 6px;
            display: inline-block; transition: all 0.2s;
        }
        .btn:hover { background-color: #2c3e50; color: white !important; }
        
        /* Analysis Card */
        .analysis-section { margin-bottom: 60px; }
        .option-title { 
            background-color: #34495e; 
            color: white; 
            padding: 15px 25px; 
            font-size: 32px !important; 
            font-weight: 700; 
            margin-bottom: 30px; 
            border-left: 10px solid #3498db; 
            border-radius: 4px;
        }
        
        /* Chart Area */
        .chart-wrapper { position: relative; width: 100%; margin-bottom: 40px; border: 1px solid #eee; padding: 30px; border-radius: 8px; }

        /* PyRevit Chart Overrides */
        /* Attempt to constrain the pyrevit generated chart containers if they are too large */
        div[id^="chart-"] { margin-bottom: 30px; }
        .pie-grid div[id^="chart-"] { max-width: 400px; margin: 0 auto; }

        /* Legend */
        .legend-container { display: flex; justify-content: flex-end; margin-bottom: 10px; }
        .legend { display: flex; gap: 20px; font-size: 16px !important; background: #f9f9f9; padding: 10px 15px; border-radius: 4px; border: 1px solid #eee; }
        .legend-item { display: flex; align-items: center; font-weight: 600; }
        .legend-color { width: 16px; height: 16px; margin-right: 8px; border-radius: 2px; }
        
        /* Visual Bars in Table - Absolute Positioning (Robust) */
        .visual-bar-container { 
            position: relative; 
            width: 100%; 
            min-width: 250px; 
            height: 40px; 
            background: #e0e0e0; 
            border-radius: 4px; 
            border: 1px solid #ccc;
            overflow: hidden; 
        }
        .visual-center-line {
            position: absolute;
            left: 50%;
            top: 0;
            bottom: 0;
            width: 2px;
            background: #bdc3c7;
            z-index: 5;
            transform: translateX(-50%);
        }
        .bar-fill { 
            position: absolute;
            top: 0;
            bottom: 0;
            height: 100%; 
            opacity: 0.9; 
            z-index: 10;
        }
        .cut-bar { 
            right: 50%; 
            background-color: #e74c3c; 
            border-top-left-radius: 4px; 
            border-bottom-left-radius: 4px; 
        }
        .fill-bar { 
            left: 50%; 
            background-color: #27ae60; 
            border-top-right-radius: 4px; 
            border-bottom-right-radius: 4px; 
        }
        
        /* Spreadsheet Table */
        .table-container { margin-top: 40px; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 28px !important; margin-top: 20px; }
        th { background: #fff; color: #555; font-weight: 700; text-align: left; padding: 30px 25px; border-bottom: 5px solid #333; text-transform: uppercase; font-size: 24px !important; letter-spacing: 1px; }
        td { padding: 30px 25px; border-bottom: 1px solid #eee; color: #333; }
        tr:hover td { background: #f8f9fa; }
        .num { text-align: right; font-family: 'Consolas', 'Monaco', monospace; font-weight: 600; font-size: 36px !important; }
        
        /* Disclaimer */
        .disclaimer { margin-top: 100px; padding-top: 30px; border-top: 2px solid #eee; color: #777; font-size: 16px !important; line-height: 1.8; }
        .disclaimer strong { color: #333; text-transform: uppercase; font-size: 14px !important; letter-spacing: 1px; }
        .disclaimer ul { padding-left: 20px; margin-top: 10px; }
        .disclaimer li { margin-bottom: 4px; }
    </style>
    </head>
    <body>
    """
    
    html = html_start + '<div class="container">'
    
    html += '<header>'
    html += '<div><h1>Earthwork Report</h1><div class="meta">Logistics Analysis</div></div>'
    html += '</header>'
    
    # Print Header & CSS
    output.print_html(html)
    
    # --- Comparative Summary (Dashboard) ---
    output.print_html('<div class="analysis-section"><div class="option-title">Grading Trends</div>')
    
    # Use pyRevit's built-in Charting (Local Resources)
    chart = output.make_line_chart()
    chart.options.title = {'display': True, 'text': 'Net Truck Logistics', 'fontSize': 20}
    chart.options.scales = {'yAxes': [{'ticks': {'beginAtZero': True}}]}
    chart.options.tooltips = {'mode': 'index', 'intersect': False}
    chart.options.maintainAspectRatio = True
    chart.options.aspectRatio = 2
    chart.options.responsive = True
    
    # Inject dummy start/end points (0) to create a "Pile of Dirt" look
    chart.data.labels = [""] + ["{} ({})".format(k[0], k[1]) for k in sorted_keys] + [""]
    
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

    output.print_html('<div class="chart-wrapper">')
    chart.draw()
    output.print_html('</div>')
    output.print_html('</div>') # End analysis-section
    
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
    
    # Action Footer
    html += '<div class="action-footer">'
    html += '<a href="#" onclick="try{window.print();}catch(e){alert(\'Press Ctrl+P to print.\');} return false;" class="btn no-print">Print Report</a>'
    html += '</div>'
    
    html += '</div>' # End Container
    html += '<script>window.scrollTo(0,0);</script>'
    html += '</body></html>'
    output.print_html(html)

if __name__ == '__main__':
    main()