# -*- coding: utf-8 -*-
"""
Finds and selects duplicate elements based on Revit Warnings.

This script scans the document for 'Identical Instances' warnings and lists the failing elements.
"""

__title__ = 'Find Duplicates'
__author__ = 'Oxyzen Digital'
__version__ = '1.1'


from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, 
                               LocationPoint, LocationCurve, CategoryType, ElementId, BuiltInFailures)

from pyrevit import revit, script, forms
import codecs
from collections import defaultdict

doc = revit.doc
uidoc = revit.uidoc

# --- Functions ---
def get_id_value(element_id):
    if hasattr(element_id, "Value"):
        return element_id.Value
    return element_id.IntegerValue

def get_element_location_key(element):
    """
    Creates a robust string representation of the element's geometry.
    Prioritizes LocationPoint and LocationCurve, falls back to BoundingBox.
    """
    if not hasattr(element, "Location"):
        return None
        
    loc = element.Location
    
    # 1. Point-based (Columns, Furniture, etc.)
    if isinstance(loc, LocationPoint):
        pt = loc.Point
        # Include Rotation to distinguish rotated duplicates
        try:
            rot = loc.Rotation
        except Exception:
            rot = 0.0
        return "Pt({:.4f},{:.4f},{:.4f}|{:.4f})".format(pt.X, pt.Y, pt.Z, rot)

    # 2. Curve-based (Walls, Beams, Ducts, Pipes)
    if isinstance(loc, LocationCurve):
        curve = loc.Curve
        # Check endpoints
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return "Crv({:.4f},{:.4f},{:.4f}-{:.4f},{:.4f},{:.4f})".format(
                p0.X, p0.Y, p0.Z, p1.X, p1.Y, p1.Z
            )
        except: pass

    # 3. Fallback: BoundingBox (Floors, Roofs, Generic Models)
    # Use Min and Max to define exact spatial extent, not just center.
    bounding_box = element.get_BoundingBox(None)
    if bounding_box:
        min_pt = bounding_box.Min
        max_pt = bounding_box.Max
        return "Box({:.4f},{:.4f},{:.4f}-{:.4f},{:.4f},{:.4f})".format(
            min_pt.X, min_pt.Y, min_pt.Z, max_pt.X, max_pt.Y, max_pt.Z
        )

    # Return None if no location can be determined
    return None

# --- Main Script ---
def find_duplicates():
    """Main function to find and report duplicate elements."""
    print("Scanning Revit Warnings...")
    
    duplicate_groups = []
    all_warnings = doc.GetWarnings()
    target_guid = BuiltInFailures.OverlapFailures.DuplicateInstances.Guid
    
    for w in all_warnings:
        if w.GetFailureDefinitionId().Guid == target_guid:
            ids = w.GetFailingElements()
            group = []
            for eid in ids:
                el = doc.GetElement(eid)
                if el:
                    group.append(el)
            
            if len(group) > 1:
                duplicate_groups.append(group)
            
    # --- Reporting Results ---
    if duplicate_groups:
        output = script.get_output()
        output.close_others()
        output.set_title("Duplicate Elements Report")
        output.resize(1100, 1100)
        
        
        # Inject CSS from external file
        css_file = script.get_bundle_file('style.css')
        if css_file:
            try:
                with codecs.open(css_file, 'r', encoding='utf-8-sig') as f:
                    output.add_style(f.read())
            except:
                pass
        
        # Group by Category
        # category_name -> list of groups (where each group is a list of elements)
        by_category = defaultdict(list)
        total_duplicates = 0
        
        for group in duplicate_groups:
            first_el = group[0]
            cat_name = "Unknown Category"
            if first_el.Category and first_el.Category.Name:
                cat_name = first_el.Category.Name
            else:
                # Fallback for elements with missing category names (e.g. some curves)
                cat_name = first_el.GetType().Name.split('.')[-1]
            
            by_category[cat_name].append(group)
            total_duplicates += len(group)
        
        # --- Summary Section ---
        summary_html = """<div class="report-summary">
            <div class="summary-text">Found {count} duplicate elements.</div>
            <span class="tip-text"><strong>Tip:</strong> Click 'Select Duplicates' to select redundant elements in Revit, then press <strong>Delete</strong> to purge.</span>
            <span class="tip-text"><strong>Note:</strong> Switch to a 3D View to ensure all element types can be selected.</span>
        </div>""".format(count=total_duplicates)
        
        output.print_html(summary_html)
        
        for cat_name in sorted(by_category.keys()):
            groups = by_category[cat_name]
            
            # Calculate total IDs for this category for "Select All"
            all_ids_in_cat = [e.Id for grp in groups for e in grp]
            
            # Calculate Purgeable IDs (All except first in each group, sorted by ID)
            purgeable_ids = []
            for grp in groups:
                # Sort by ID to ensure we keep the oldest (lowest ID)
                s_grp = sorted(grp, key=lambda x: get_id_value(x.Id))
                for e in s_grp[1:]:
                    purgeable_ids.append(e.Id)
            
            # Select All Button
            btn_link = output.linkify(all_ids_in_cat, title="🔍 Select All")
            btn_html = '<span class="select-all-wrapper">{}</span>'.format(btn_link)
            
            # Select Purgeable Button
            purge_link = output.linkify(purgeable_ids, title="🔍 Select Duplicates")
            purge_html = '<span class="select-duplicates-wrapper">{}</span>'.format(purge_link)
            
            # Start Table HTML
            table_html = '<table style="margin:0;"><thead><tr>'
            table_html += '<th style="width:20%">Duplicate IDs</th><th style="width:15%">Family</th><th style="width:15%">Type</th>'
            table_html += '<th style="width:10%">Level</th><th style="width:10%">Workset</th><th style="width:15%">Location</th><th style="width:15%">Action</th>'
            table_html += '</tr></thead><tbody>'
            
            for group in groups:
                # Sort group by ID to ensure deterministic order (Oldest first)
                group = sorted(group, key=lambda x: get_id_value(x.Id))
                el = group[0] # Original
                fam = ""
                typ = ""
                try:
                    type_id = el.GetTypeId()
                    if type_id != ElementId.InvalidElementId:
                        etype = doc.GetElement(type_id)
                        if etype:
                            fam = etype.FamilyName if hasattr(etype, "FamilyName") else ""
                            typ = etype.Name if hasattr(etype, "Name") else ""
                except: pass
                
                # Level Info
                lvl = "-"
                if hasattr(el, "LevelId") and el.LevelId != ElementId.InvalidElementId:
                    l_el = doc.GetElement(el.LevelId)
                    if l_el: lvl = l_el.Name
                    
                # Workset Info
                workset_name = "-"
                if doc.IsWorkshared:
                    try:
                        ws_id = el.WorksetId
                        ws = doc.GetWorksetTable().GetWorkset(ws_id)
                        if ws:
                            workset_name = ws.Name
                    except:
                        pass
                
                # Create a visual group of ID links
                id_links = []
                for i, e in enumerate(group):
                    lnk = output.linkify(e.Id)
                    if i == 0:
                        # Highlight Original
                        id_links.append('<span class="id-original" title="Original (Oldest)">{}</span>'.format(lnk))
                    else:
                        id_links.append(lnk)
                        
                ids_html = '<div class="id-group">{}</div>'.format(" ".join(id_links))
                
                # Row Action
                row_dups = [e.Id for e in group[1:]]
                action_link = output.linkify(row_dups, title="Select Dups")
                action_html = '<span class="row-action">{}</span>'.format(action_link)
                
                # Append Row
                table_html += "<tr>"
                table_html += "<td>{}</td>".format(ids_html)
                table_html += "<td>{}</td>".format(fam)
                table_html += "<td>{}</td>".format(typ)
                table_html += "<td>{}</td>".format(lvl)
                table_html += "<td>{}</td>".format(workset_name)
                table_html += "<td>{}</td>".format(get_element_location_key(el))
                table_html += "<td>{}</td>".format(action_html)
                table_html += "</tr>"
            
            table_html += "</tbody></table>"
            
            # Wrap in Section Div (Card Style)
            full_block = '<div class="category-section">'
            full_block += '<div class="category-header">'
            full_block += '<div class="category-title">{cat} <span class="item-count">{count} items</span></div>'
            full_block += '<div>{}{}</div></div>'.format(btn_html, purge_html)
            full_block += '<div class="category-content">{table}</div>'
            full_block += '</div>'
            full_block = full_block.format(cat=cat_name, count=len(all_ids_in_cat), btn=btn_html, table=table_html)
            
            output.print_html(full_block)
            
        # Scroll to top
        output.print_html('<script>window.scrollTo(0,0);</script>')
            
    else:
        # If no duplicates are found
        print("No duplicates found. Model is clean. \u2705")
        forms.alert("No duplicate model elements found.", title="Clean")

# --- Run the script ---
if __name__ == '__main__':
    # A transaction is not needed for just selecting elements
    find_duplicates()