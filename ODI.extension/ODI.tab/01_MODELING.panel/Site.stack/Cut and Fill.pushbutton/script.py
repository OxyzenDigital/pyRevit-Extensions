# -*- coding: utf-8 -*-
"""
Cut and Fill Tool
Calculates Cut and Fill volumes for Toposolids in the project.
"""
__title__ = "Cut and Fill"
__version__ = "2.1"
__author__ = "Oxyzen Digital"
__context__ = "doc-project"

import json
import os
import codecs
import datetime
from pyrevit import revit, forms, script
from pyrevit import DB

doc = revit.doc

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
        
        # Type Name
        self.TypeName = ""
        try:
            tid = element.GetTypeId()
            if tid != DB.ElementId.InvalidElementId:
                etype = doc.GetElement(tid)
                if etype: self.TypeName = etype.Name
        except: pass
        
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

        # Subdivision Check
        self.IsSubdivision = False
        if hasattr(element, "HostTopoId") and element.HostTopoId != DB.ElementId.InvalidElementId:
            self.IsSubdivision = True

        # Volumes (Internal = CF)
        self.CutCF = 0.0
        self.FillCF = 0.0
        self.TotalVolCF = 0.0
        
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
        
        # Total Volume (Geometric)
        p_vol = element.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED)
        if p_vol and p_vol.HasValue:
            self.TotalVolCF = p_vol.AsDouble()
            
        self.NetCF = self.CutCF - self.FillCF # Cut is positive removal, Fill is positive addition. Net = Cut - Fill? Or Net Volume change? Usually Net = Fill - Cut (import required) or Cut - Fill (export available). 
        # Let's denote Net as (Cut - Fill). Positive = Export, Negative = Import.
        
    @property
    def CutCY(self): return self.CutCF * CF_TO_CY
    @property
    def FillCY(self): return self.FillCF * CF_TO_CY
    @property
    def NetCY(self): return self.NetCF * CF_TO_CY
    @property
    def TotalVolCY(self): return self.TotalVolCF * CF_TO_CY
    
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
    all_data = []
    
    for el in elements:
        t_data = TopoData(el)
        all_data.append(t_data)
        
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
    
    # 3. Prepare HTML Report
    
    # --- Prepare CSS ---
    css_content = ""
    css_file = script.get_bundle_file('style.css')
    if css_file:
        # Robust read: try utf-8-sig first (handles BOM), then default
        try:
            with codecs.open(css_file, 'r', encoding='utf-8-sig') as f:
                css_content = f.read()
        except:
            with open(css_file, 'r') as f:
                css_content = f.read()
    
    # --- Prepare Header HTML ---
    html_header = ''
    
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    html_header += '<header>'
    html_header += '<div><h1>Earthwork Report</h1><div class="meta">Logistics Analysis</div></div>'
    html_header += '<div class="meta" style="text-align:right;">Generated: {}</div>'.format(now_str)
    html_header += '</header>'
    
    # Inject dummy start/end points (0) to create a "Pile of Dirt" look
    # Prepare Data for Chart.js
    labels_raw = [""] + ["{} ({})".format(k[0], k[1]) for k in sorted_keys] + [""]
    net_vals = [options_map[k]["totals"]["net_trucks"] for k in sorted_keys]
    data_export = [0] + [max(0, v) for v in net_vals] + [0]
    data_import = [0] + [abs(min(0, v)) for v in net_vals] + [0]

    # Standalone Chart Config (Chart.js JSON)
    chart_json = {
        "type": "line",
        "data": {
            "labels": labels_raw,
            "datasets": [
                {
                    "label": "Net Export (Trucks Out)",
                    "data": data_export,
                    "backgroundColor": "rgba(231, 76, 60, 0.5)",
                    "borderColor": "rgba(231, 76, 60, 1)",
                    "borderWidth": 1,
                    "fill": True,
                    "lineTension": 0.4
                },
                {
                    "label": "Net Import (Trucks In)",
                    "data": data_import,
                    "backgroundColor": "rgba(39, 174, 96, 0.5)",
                    "borderColor": "rgba(39, 174, 96, 1)",
                    "borderWidth": 1,
                    "fill": True,
                    "lineTension": 0.4
                }
            ]
        },
        "options": {
            "title": {"display": True, "text": "Net Truck Logistics", "fontSize": 20},
            "scales": {"yAxes": [{"ticks": {"beginAtZero": True}}]},
            "tooltips": {"mode": "index", "intersect": False},
            "maintainAspectRatio": True,
            "aspectRatio": 2,
            "responsive": True
        }
    }

    # --- Prepare Report Body HTML ---
    # --- The Spreadsheet ---
    # Calculate max NET value for scaling visual bars (Visual Balance)
    max_net_abs = 0.0
    for k in sorted_keys:
        t = options_map[k]["totals"]
        max_net_abs = max(max_net_abs, abs(t["net_trucks"]))
    if max_net_abs == 0: max_net_abs = 1.0

    html_body = '<div class="analysis-section"><div class="option-title">Detailed Breakdown</div>'
    
    # --- Settings Controls (Embedded JS) ---
    html_body += '<div style="background:#f8f9fa; padding:15px; border-bottom:1px solid #ddd; font-size:14px;">'
    
    # Row 1: Logistics
    html_body += '<div style="margin-bottom:10px; display:flex; align-items:center; gap:20px;">'
    html_body += '<strong>Logistics:</strong>'
    html_body += '<label title="Cubic Yards per Truck">Truck Cap (CY): <input type="number" id="truckCap" value="{}" step="0.5" style="width:60px; padding:4px; border:1px solid #ccc; border-radius:3px;"></label>'.format(TRUCK_CAP_CY)
    html_body += '<label title="Bank to Loose">Swell Factor: <input type="number" id="swellFactor" value="{}" step="0.05" style="width:60px; padding:4px; border:1px solid #ccc; border-radius:3px;"></label>'.format(FACTOR_SWELL)
    html_body += '<label title="Loose to Compacted">Compaction Req: <input type="number" id="compactFactor" value="{}" step="0.05" style="width:60px; padding:4px; border:1px solid #ccc; border-radius:3px;"></label>'.format(FACTOR_COMPACTION_REQ)
    html_body += '</div>'

    # Row 2: Cost
    html_body += '<div style="display:flex; align-items:center; gap:20px;">'
    html_body += '<strong>Est. Cost:</strong>'
    html_body += '<label>Unit Cost: $<input type="number" id="unitCost" value="200.00" style="width:80px; padding:4px; margin-left:5px; border:1px solid #ccc; border-radius:3px;"></label>'
    html_body += '<label>Basis: <select id="costBasis" style="padding:4px; margin-left:5px; border:1px solid #ccc; border-radius:3px;"><option value="truck">Per Truck Load</option><option value="cy">Per Cubic Yard</option></select></label>'
    html_body += '</div>'
    
    html_body += '</div>' # End settings div

    html_body += '<div class="table-container">'
    html_body += '<table id="breakdownTable"><thead><tr>'
    html_body += '<th>Design Option</th><th>Phase</th>'
    html_body += '<th>Visual Balance</th>'
    html_body += '<th class="num">Cut (C.Y.)</th><th class="num">Fill (C.Y.)</th>'
    html_body += '<th class="num">Net (C.Y.)</th><th class="num">Net Trucks</th>'
    html_body += '<th class="num">Est. Cost</th>'
    html_body += '</tr></thead><tbody>'
    
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
        
        bar_html = '<div class="visual-bar-container" style="background-color: #e0e0e0 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact;"><div class="visual-center-line"></div>'
        if net_t > 0:
            bar_html += '<div class="bar-fill cut-bar" style="width:{:.1f}%; background-color: #e74c3c !important; -webkit-print-color-adjust: exact; print-color-adjust: exact;"></div>'.format(pct)
        elif net_t < 0:
            bar_html += '<div class="bar-fill fill-bar" style="width:{:.1f}%; background-color: #27ae60 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact;"></div>'.format(pct)
        bar_html += '</div>'
        
        html_body += '<tr data-cut-cy="{:.2f}" data-fill-cy="{:.2f}" data-net-cy="{:.2f}"><td><strong>{}</strong></td><td>{}</td>'.format(t["cut_cy"], t["fill_cy"], net_cy, opt, phase)
        html_body += '<td>{}</td>'.format(bar_html)
        html_body += '<td class="num">{:,.0f}</td><td class="num">{:,.0f}</td>'.format(t["cut_cy"], t["fill_cy"])
        html_body += '<td class="num" style="{}">{:,.0f}</td><td class="num" style="{}">{:,.1f}</td>'.format(c_style, abs(net_cy), c_style, abs(net_t))
        html_body += '<td class="num cost-cell">$0.00</td>'
        html_body += '</tr>'
        
    html_body += '</tbody></table></div>'
    html_body += '</div>' # End analysis-section
    html_body += '</div>'

    # Disclaimer
    html_body += '<div class="disclaimer">'
    html_body += '<strong>Assumptions & Notes:</strong><br>'
    html_body += '<ul>'
    html_body += '<li><strong>Cut (Export):</strong> Based on Bank Volume * Swell Factor.</li>'
    html_body += '<li><strong>Fill (Import):</strong> Based on Compacted Volume * Compaction Factor.</li>'
    html_body += '<li id="calcDisclaimer"><strong>Logistics:</strong> Calculating...</li>'
    html_body += '</ul></div>'
    
    # --- Raw Data Log ---
    html_body += '<div class="analysis-section" style="margin-top: 40px;">'
    html_body += '<div class="option-title" style="background-color: #7f8c8d; border-left-color: #95a5a6;">Raw Data Log</div>'
    html_body += '<div class="table-container"><table><thead><tr>'
    html_body += '<th>Element ID</th><th>Type</th><th>Subdiv?</th>'
    html_body += '<th class="num">Cut (C.Y.)</th><th class="num">Fill (C.Y.)</th><th class="num">Net (C.Y.)</th>'
    html_body += '<th class="num">Total Vol (C.Y.)</th>'
    html_body += '</tr></thead><tbody>'

    for t in all_data:
        # Highlight Subdivisions
        row_style = "background-color: #fff3cd;" if t.IsSubdivision else ""
        eid_val = t.Id.IntegerValue if hasattr(t.Id, "IntegerValue") else t.Id.Value
        
        html_body += '<tr style="{}">'.format(row_style)
        html_body += '<td>{}</td>'.format(eid_val)
        html_body += '<td>{}</td>'.format(t.TypeName)
        html_body += '<td style="text-align:center;">{}</td>'.format("Yes" if t.IsSubdivision else "-")
        html_body += '<td class="num">{:,.1f}</td>'.format(t.CutCY)
        html_body += '<td class="num">{:,.1f}</td>'.format(t.FillCY)
        html_body += '<td class="num">{:,.1f}</td>'.format(t.NetCY)
        html_body += '<td class="num" style="font-weight:bold;">{:,.1f}</td>'.format(t.TotalVolCY)
        html_body += '</tr>'

    html_body += '</tbody></table></div></div>'

    # --- Footer Buttons ---
    html_body += '<div class="action-footer">'
    html_body += '<button class="btn" onclick="window.print()">Print Report</button>'
    html_body += '</div>'

    # --- Construct Full HTML Document ---
    full_html = "<!DOCTYPE html><html><head>"
    full_html += '<meta charset="utf-8">'
    full_html += '<title>Cut & Fill Report</title>'
    full_html += '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.9.4/Chart.min.js"></script>'
    if css_content:
        full_html += '<style>' + css_content + '</style>'
    full_html += '</head><body>'
    
    full_html += html_header
    
    # Chart Section
    full_html += '<div class="analysis-section"><div class="option-title">Grading Trends</div>'
    full_html += '<canvas id="logisticsChart"></canvas></div>'
    full_html += '<script>var ctx = document.getElementById("logisticsChart").getContext("2d"); var myChart = new Chart(ctx, {});</script>'.format(json.dumps(chart_json))
    
    full_html += html_body
    
    # --- Embedded JS for Cost Calculation ---
    full_html += '''
<script>
    (function() {
        // Inputs
        var costInput = document.getElementById('unitCost');
        var basisSelect = document.getElementById('costBasis');
        var truckCapInput = document.getElementById('truckCap');
        var swellInput = document.getElementById('swellFactor');
        var compactInput = document.getElementById('compactFactor');
        
        // Elements
        var table = document.getElementById('breakdownTable');
        var disclaimer = document.getElementById('calcDisclaimer');

        function updateCalculations() {
            // Read Values
            var cost = parseFloat(costInput.value) || 0;
            var basis = basisSelect.value;
            var cap = parseFloat(truckCapInput.value) || 14;
            var swell = parseFloat(swellInput.value) || 1.25;
            var compact = parseFloat(compactInput.value) || 1.30;

            var rows = table.querySelectorAll('tbody tr');
            var maxTrucks = 0;
            var rowData = [];
            
            // 1. Calculate Trucks & Find Max for Scaling
            rows.forEach(function(row) {
                var cutCy = parseFloat(row.getAttribute('data-cut-cy')) || 0;
                var fillCy = parseFloat(row.getAttribute('data-fill-cy')) || 0;
                var netCy = parseFloat(row.getAttribute('data-net-cy')) || 0;
                
                var cutTrucks = (cutCy * swell) / cap;
                var fillTrucks = (fillCy * compact) / cap;
                var netTrucks = cutTrucks - fillTrucks;
                
                if(Math.abs(netTrucks) > maxTrucks) maxTrucks = Math.abs(netTrucks);
                
                rowData.push({row: row, netTrucks: netTrucks, netCy: netCy});
            });
            
            if(maxTrucks === 0) maxTrucks = 1;

            // 2. Update DOM
            rowData.forEach(function(d) {
                // Update Visual Bar
                var pct = (Math.abs(d.netTrucks) / maxTrucks) * 50.0;
                var barHtml = '<div class="visual-center-line"></div>';
                if(d.netTrucks > 0) barHtml += '<div class="bar-fill cut-bar" style="width:' + pct.toFixed(1) + '%; background-color: #e74c3c !important; -webkit-print-color-adjust: exact; print-color-adjust: exact;"></div>';
                else if(d.netTrucks < 0) barHtml += '<div class="bar-fill fill-bar" style="width:' + pct.toFixed(1) + '%; background-color: #27ae60 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact;"></div>';
                d.row.cells[2].querySelector('.visual-bar-container').innerHTML = barHtml;

                // Update Net Trucks Text
                var tCell = d.row.cells[6];
                tCell.textContent = Math.abs(d.netTrucks).toFixed(1);
                tCell.style.color = (d.netTrucks > 0.1) ? '#c0392b' : (d.netTrucks < -0.1 ? '#27ae60' : '#7f8c8d');
                tCell.style.fontWeight = (Math.abs(d.netTrucks) > 0.1) ? 'bold' : 'normal';

                // Update Cost
                var qty = (basis === 'truck') ? Math.abs(d.netTrucks) : Math.abs(d.netCy);
                var total = qty * cost;
                d.row.querySelector('.cost-cell').textContent = '$' + total.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
            });
            
            var basisLabel = (basis === 'truck') ? 'Per Truck Load' : 'Per Cubic Yard';
            disclaimer.innerHTML = '<strong>Logistics:</strong> ' + cap + ' CY Trucks. Swell: ' + swell + '. Compact: ' + compact + '.<br><strong>Cost:</strong> $' + cost.toFixed(2) + ' ' + basisLabel + '.';
        }

        if(costInput && basisSelect) {
            [costInput, basisSelect, truckCapInput, swellInput, compactInput].forEach(function(el) {
                el.addEventListener('input', updateCalculations);
                el.addEventListener('change', updateCalculations);
            });
            updateCalculations();
        }
    })();
</script>
'''
    full_html += '</body></html>'

    # --- Save to Downloads & Open ---
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = "CutAndFill_Report_{}.html".format(datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S"))
    file_path = os.path.join(downloads_path, filename)
    
    try:
        with codecs.open(file_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        
        # Open in Default Browser
        os.startfile(file_path)
        
        # Notify User
        forms.alert("Report generated successfully!\n\nSaved to: {}\nOpened in default browser.".format(filename), title="Success")
        
    except Exception as e:
        forms.alert("Failed to save report: {}".format(e))

if __name__ == '__main__':
    main()