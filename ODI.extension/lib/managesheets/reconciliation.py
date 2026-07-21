# -*- coding: utf-8 -*-
"""
AIA Sheet Reconciliation & Sequencing Algorithm — Implementation Spec v1
Phase 1-6 Engine (Pure Python / Revit API Data Layer)
"""

from collections import OrderedDict
import re

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    ElementId, ViewSheet, View, ViewType,
    Level, BoundingBoxXYZ, XYZ, Transaction, TransactionGroup
)

class SheetRecord(object):
    def __init__(self, element_id, number, name, is_placeholder, sheet_collection):
        self.element_id = element_id
        self.number = number
        self.name = name
        self.is_placeholder = is_placeholder
        self.sheet_collection = sheet_collection or "Default"
        self.series_name = "Unknown"
        self.placed_view_ids = []

class ViewRecord(object):
    def __init__(self, element_id, name, view_type, level_name, scope_box_name, crop_active, crop_area_sqft, scale, on_sheet_id):
        self.element_id = element_id
        self.name = name
        self.view_type = view_type
        self.level_name = level_name
        self.scope_box_name = scope_box_name
        self.crop_active = crop_active
        self.crop_area_sqft = crop_area_sqft
        self.scale = scale
        self.on_sheet_id = on_sheet_id

class ScopeBoxRecord(object):
    def __init__(self, element_id, name, bbox):
        self.element_id = element_id
        self.name = name
        self.bbox = bbox

class LevelRecord(object):
    def __init__(self, element_id, name, elevation):
        self.element_id = element_id
        self.name = name
        self.elevation = elevation

class HarvestData(object):
    def __init__(self):
        self.sheets = []          # list of SheetRecord
        self.views = []           # list of ViewRecord
        self.scope_boxes = []     # list of ScopeBoxRecord
        self.levels = []          # list of LevelRecord, sorted by elevation
        self.grid_size = "none"   # "none", "2x2", "3x3", "4x4"
        self.warnings = []

def harvest_project(doc):
    """Phase 1: Harvest"""
    data = HarvestData()
    
    # 1. Harvest Levels
    lvl_collector = FilteredElementCollector(doc).OfClass(Level)
    for lvl in lvl_collector:
        data.levels.append(LevelRecord(lvl.Id, lvl.Name, lvl.Elevation))
    data.levels.sort(key=lambda x: x.elevation)
    
    # 2. Harvest Scope Boxes
    sb_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType()
    for sb in sb_collector:
        bbox = sb.get_BoundingBox(None)
        data.scope_boxes.append(ScopeBoxRecord(sb.Id, sb.Name, bbox))
        
    segment_count = sum(1 for sb in data.scope_boxes if "segment" in sb.name.lower() or "area" in sb.name.lower() or sb.name.startswith("Sector"))
    if segment_count == 4: data.grid_size = "2x2"
    elif segment_count == 9: data.grid_size = "3x3"
    elif segment_count == 16: data.grid_size = "4x4"
    elif segment_count > 0:
        data.warnings.append("GRID_IRREGULAR: {} segment boxes found. Expected 4, 9, or 16.".format(segment_count))
    
    # 3. Harvest Sheets
    sheet_collector = FilteredElementCollector(doc).OfClass(ViewSheet)
    for sh in sheet_collector:
        sheet_collection = "Undefined"
        try:
            from Autodesk.Revit.DB import ParameterTypeId
            if hasattr(ParameterTypeId, "SheetCollection"):
                param = sh.GetParameter(ParameterTypeId.SheetCollection)
                if param and param.AsElementId() != ElementId.InvalidElementId:
                    col_elem = doc.GetElement(param.AsElementId())
                    if col_elem: sheet_collection = col_elem.Name
            else:
                param = sh.LookupParameter(" Sheet Collection")
                if param and param.HasValue: sheet_collection = param.AsString()
        except:
            pass
            
        rec = SheetRecord(sh.Id, sh.SheetNumber, sh.Name, sh.IsPlaceholder, sheet_collection)
        
        # Capture the custom "Sheet Series" parameter to persist user overrides
        try:
            series_param = sh.LookupParameter("Sheet Series")
            if series_param and series_param.HasValue:
                rec.series_name = series_param.AsString()
        except:
            pass
            
        try:
            for v_id in sh.GetAllPlacedViews():
                rec.placed_view_ids.append(v_id)
        except:
            pass
        data.sheets.append(rec)
        
    # 4. Harvest Views
    view_collector = FilteredElementCollector(doc).OfClass(View)
    for v in view_collector:
        if v.IsTemplate: continue
        v_type = v.ViewType
        if v_type not in (ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.EngineeringPlan, ViewType.Elevation, ViewType.Section, ViewType.Detail, ViewType.ThreeD):
            continue
            
        on_sheet_id = None
        sheet_num_param = v.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER)
        if sheet_num_param and sheet_num_param.AsString() and sheet_num_param.AsString() != "---":
            for sh in data.sheets:
                if sh.number == sheet_num_param.AsString():
                    on_sheet_id = sh.element_id
                    break
        
        data.views.append(ViewRecord(
            element_id=v.Id,
            name=v.Name,
            view_type=v_type,
            level_name=v.GenLevel.Name if hasattr(v, "GenLevel") and v.GenLevel else None,
            scope_box_name=None,
            crop_active=v.CropBoxActive,
            crop_area_sqft=0.0,
            scale=v.Scale,
            on_sheet_id=on_sheet_id
        ))
        
    return data

class SlotRecord(object):
    def __init__(self, target_number, target_name, discipline, series, ordinal, suffix, canonical_name, collection, series_name="Unknown", cg="Unknown"):
        self.target_number = target_number
        self.target_name = target_name
        self.discipline = discipline
        self.series = series
        self.ordinal = ordinal
        self.suffix = suffix
        self.canonical_name = canonical_name
        self.collection = collection
        self.series_name = series_name
        self.cg = cg
        
        # Match variables
        self.status = "MISSING"
        self.existing_number = ""
        self.existing_name = ""
        self.sheet_element_id = ElementId.InvalidElementId
        self.match_score = 0.0
        
        # View resolution fields
        self.view_action = "NONE"
        self.view_element_id = ElementId.InvalidElementId
        self.view_recipe = None
        self.requires_user_decision = False

def generate_slots(generated_targets):
    """
    Phase 2: Schema Generation (in a vacuum)
    Returns a deterministic list of required AIA sheet objects ("slots") based on the targets from Project Setup.
    """
    slots = []
    
    for t in generated_targets:
        t_num = t["num"]
        t_name = t["name"]
        
        # parse discipline, series, ordinal, suffix
        match = re.match(r"^([A-Z]+)[- ]?(\d)(\d\d)([A-Za-z]?)(.*)", t_num)
        if match:
            disc, series, ordinal, suffix, rest = match.groups()
            suffix = suffix + rest
        else:
            match2 = re.match(r"^([A-Z]+)[- ]?(\d+)", t_num)
            disc = match2.group(1) if match2 else "Unknown"
            series = "0"
            ordinal = "00"
            suffix = ""
            
        slots.append(SlotRecord(
            target_number=t_num,
            target_name=t_name,
            discipline=disc,
            series=series,
            ordinal=ordinal,
            suffix=suffix,
            canonical_name=t_name,
            collection=t.get("collection", "Undefined"),
                series_name=t.get("series_name", "Unknown"),
            cg=t.get("cg", "Unknown")
        ))
            
    return slots
    
def normalize_name(name):
    if not name: return ""
    name = name.lower()
    name = re.sub(r'[-–—]', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    abbrevs = {"flr": "floor", "plt": "plant", "elev": "elevation", "dtl": "detail", "sched": "schedule", "enl": "enlarged", "rcp": "reflected ceiling plan", "demo": "demolition", "lvl": "level", "l": "level"}
    words = name.split()
    for i in range(len(words)):
        if words[i] in abbrevs:
            words[i] = abbrevs[words[i]]
    return " ".join(words)

def token_set_ratio(s1, s2):
    w1 = set(s1.split())
    w2 = set(s2.split())
    common = w1.intersection(w2)
    if not w1 and not w2: return 1.0
    if not w1 or not w2: return 0.0
    return (2.0 * len(common)) / (len(w1) + len(w2))

def match_sheets(harvested_data, slots):
    """
    Phase 3: Match
    Pure in-memory mapping function. No Revit transactions.
    """
    unmatched_sheets = list(harvested_data.sheets)
    unmatched_slots = list(slots)
    
    sheet_to_views = {}
    for v in harvested_data.views:
        if v.on_sheet_id:
            val = v.on_sheet_id.IntegerValue if hasattr(v.on_sheet_id, "IntegerValue") else v.on_sheet_id.Value
            if val not in sheet_to_views:
                sheet_to_views[val] = []
            sheet_to_views[val].append(v)
            
    def get_matches(slots_subset, sheets_subset):
        def remove_matched(s, sh):
            slots_subset.remove(s)
            sheets_subset.remove(sh)
            if s in unmatched_slots: unmatched_slots.remove(s)
            if sh in unmatched_sheets: unmatched_sheets.remove(sh)
            
        # Pass A: exact number + exact name
        for s in list(slots_subset):
            for sh in list(sheets_subset):
                if s.target_number == sh.number and s.target_name == sh.name:
                    s.status = "MATCHED"
                    s.existing_number, s.existing_name, s.sheet_element_id = sh.number, sh.name, sh.element_id
                    remove_matched(s, sh)
                    break
        # Pass B: exact number, name differs
        for s in list(slots_subset):
            for sh in list(sheets_subset):
                if s.target_number == sh.number:
                    s.status = "RENAME_NAME"
                    s.existing_number, s.existing_name, s.sheet_element_id = sh.number, sh.name, sh.element_id
                    remove_matched(s, sh)
                    break
        # Pass C: exact name, number differs
        for s in list(slots_subset):
            norm_slot = normalize_name(s.target_name)
            for sh in list(sheets_subset):
                if norm_slot == normalize_name(sh.name):
                    s.status = "RENAME_NUMBER"
                    s.existing_number, s.existing_name, s.sheet_element_id = sh.number, sh.name, sh.element_id
                    remove_matched(s, sh)
                    break
        # Pass D: Fuzzy match
        for s in list(slots_subset):
            best_sh, best_score = None, 0.0
            norm_slot = normalize_name(s.target_name)
            
            # Extract intent from target slot name
            target_level_hint = None
            lvl_match = re.search(r'level (\d+)|(\d+)(st|nd|rd|th) floor', norm_slot)
            if lvl_match:
                target_level_hint = lvl_match.group(1) or lvl_match.group(2)
                
            is_dim_plan = "dimension" in norm_slot
            is_demo_plan = "demolition" in norm_slot
            
            core_keywords = {"dimension", "demolition", "floor", "ceiling", "enlarged", "detail", "elevation", "section", "plan"}
            
            for sh in list(sheets_subset):
                sh_name_norm = normalize_name(sh.name)
                name_score = token_set_ratio(norm_slot, sh_name_norm)
                
                # Subset detection & Keyword matching
                w1 = set(norm_slot.split())
                w2 = set(sh_name_norm.split())
                
                core_w1 = w1.intersection(core_keywords)
                core_w2 = w2.intersection(core_keywords)
                if core_w1 and core_w1 == core_w2:
                    name_score += 0.2
                    
                if len(w2) >= 2 and w2.issubset(w1):
                    name_score += 0.3
                    
                num_aff = 0.0
                sh_num_norm = re.sub(r'[^A-Z0-9]', '', sh.number.upper())
                s_num_norm = re.sub(r'[^A-Z0-9]', '', s.target_number.upper())
                if sh_num_norm == s_num_norm:
                    num_aff += 1.0
                else:
                    if len(s_num_norm) > 0 and (sh_num_norm.startswith(s_num_norm) or s_num_norm.startswith(sh_num_norm)):
                        num_aff += 0.2
                score = 0.5 * name_score + 0.5 * num_aff
                
                # --- Semantic Boosting ---
                sh_val = sh.element_id.IntegerValue if hasattr(sh.element_id, "IntegerValue") else sh.element_id.Value
                views_on_sheet = sheet_to_views.get(sh_val, [])
                
                semantic_boost = 0.0
                for v in views_on_sheet:
                    # 1. Level Match
                    if target_level_hint and v.level_name:
                        v_lvl_norm = normalize_name(v.level_name)
                        v_lvl_match = re.search(r'level (\d+)|(\d+)(st|nd|rd|th) floor', v_lvl_norm)
                        if v_lvl_match:
                            v_hint = v_lvl_match.group(1) or v_lvl_match.group(2)
                            if v_hint == target_level_hint:
                                semantic_boost += 0.3
                                
                    # 2. Intent Match
                    v_name_norm = normalize_name(v.name)
                    if is_dim_plan and "dimension" in v_name_norm:
                        semantic_boost += 0.2
                    if is_demo_plan and "demolition" in v_name_norm:
                        semantic_boost += 0.2
                        
                    # Cap boost per view
                    if semantic_boost > 0.4:
                        semantic_boost = 0.4
                        break
                        
                score += semantic_boost
                
                if score > best_score:
                    best_score, best_sh = score, sh
            
            threshold = 0.60 if s.collection == getattr(best_sh, "sheet_collection", "") else 0.70
            
            if best_score >= threshold and best_sh:
                s.status = "RENAME_BOTH"
                s.match_score = best_score
                s.existing_number, s.existing_name, s.sheet_element_id = best_sh.number, best_sh.name, best_sh.element_id
                remove_matched(s, best_sh)

        # Pass E: Semantic View-Centric Match
        for s in list(slots_subset):
            best_sh, best_score = None, 0.0
            norm_slot = normalize_name(s.target_name)
            
            target_level_hint = None
            lvl_match = re.search(r'level (\d+)|(\d+)(st|nd|rd|th) floor', norm_slot)
            if lvl_match:
                target_level_hint = lvl_match.group(1) or lvl_match.group(2)
                
            is_dim_plan = "dimension" in norm_slot
            is_demo_plan = "demolition" in norm_slot
            
            for sh in list(sheets_subset):
                sh_val = sh.element_id.IntegerValue if hasattr(sh.element_id, "IntegerValue") else sh.element_id.Value
                views_on_sheet = sheet_to_views.get(sh_val, [])
                
                for v in views_on_sheet:
                    v_name_norm = normalize_name(v.name)
                    view_score = token_set_ratio(norm_slot, v_name_norm)
                    
                    # 1. Level Match
                    if target_level_hint and v.level_name:
                        v_lvl_norm = normalize_name(v.level_name)
                        v_lvl_match = re.search(r'level (\d+)|(\d+)(st|nd|rd|th) floor', v_lvl_norm)
                        if v_lvl_match:
                            v_hint = v_lvl_match.group(1) or v_lvl_match.group(2)
                            if v_hint == target_level_hint:
                                view_score += 0.5
                                
                    if target_level_hint and not v.level_name:
                        v_lvl_match = re.search(r'level (\d+)|(\d+)(st|nd|rd|th) floor', v_name_norm)
                        if v_lvl_match:
                            v_hint = v_lvl_match.group(1) or v_lvl_match.group(2)
                            if v_hint == target_level_hint:
                                view_score += 0.5
                                
                    # 2. Intent Match
                    if is_dim_plan and "dimension" in v_name_norm:
                        view_score += 0.5
                    if is_demo_plan and "demolition" in v_name_norm:
                        view_score += 0.5
                        
                    if view_score > best_score:
                        best_score = view_score
                        best_sh = sh
                        
            if best_score >= 0.70 and best_sh:
                s.status = "RENAME_BOTH"
                s.match_score = best_score
                s.existing_number, s.existing_name, s.sheet_element_id = best_sh.number, best_sh.name, best_sh.element_id
                remove_matched(s, best_sh)
    
    # Pass 1: Strict Same Collection Matching
    target_collections = set(s.collection for s in slots)
    
    for coll in target_collections:
        coll_slots = [s for s in unmatched_slots if s.collection == coll]
        coll_sheets = [sh for sh in unmatched_sheets if sh.sheet_collection == coll]
        get_matches(coll_slots, coll_sheets)

    # Note: Cross-Collection Migration Matching has been removed to strictly isolate Sheet Collections.

    # Pass 2: Auto-Adopt properly formatted but unanticipated sheets
    for sh in unmatched_sheets:
        m = re.match(r"^([A-Z]+)[- ]?(\d)(\d\d\d?)(.*)$", sh.number)
        if m:
            prefix = m.group(1)
            series_digit = m.group(2)
            
            ref_slot = None
            # Prioritize matching a slot in the same collection for precise ContentGroup metadata
            for s in slots:
                if s.collection == sh.sheet_collection:
                    sm = re.match(r"^([A-Z]+)[- ]?(\d)", s.target_number)
                    if sm and sm.group(1) == prefix and sm.group(2) == series_digit:
                        ref_slot = s
                        break
            # Fallback to any collection if needed to borrow discipline/series names
            if not ref_slot:
                for s in slots:
                    sm = re.match(r"^([A-Z]+)[- ]?(\d)", s.target_number)
                    if sm and sm.group(1) == prefix and sm.group(2) == series_digit:
                        ref_slot = s
                        break
                        
            if not ref_slot and getattr(sh, 'series_name', None) and sh.series_name != "Unknown":
                new_slot = SlotRecord(
                    target_number=sh.number,
                    target_name=sh.name,
                    discipline="Unknown",
                    series=prefix,
                    ordinal=int(m.group(3)),
                    suffix=m.group(4),
                    canonical_name=sh.name,
                    collection=sh.sheet_collection,
                    series_name=sh.series_name,
                    cg="Uncategorized"
                )
                new_slot.status = "MATCHED"
                new_slot.match_score = 1.0
                new_slot.existing_number = sh.number
                new_slot.existing_name = sh.name
                new_slot.sheet_element_id = sh.element_id
                slots.append(new_slot)
            elif ref_slot:
                new_slot = SlotRecord(
                    target_number=sh.number,
                    target_name=sh.name,
                    discipline=ref_slot.discipline,
                    series=ref_slot.series,
                    ordinal=int(m.group(3)),
                    suffix=m.group(4),
                    canonical_name=sh.name,
                    collection=sh.sheet_collection,
                    series_name=ref_slot.series_name,
                    cg=ref_slot.cg
                )
                new_slot.status = "MATCHED"
                new_slot.match_score = 1.0
                new_slot.existing_number = sh.number
                new_slot.existing_name = sh.name
                new_slot.sheet_element_id = sh.element_id
                slots.append(new_slot)

    def safe_id(eid):
        if eid == ElementId.InvalidElementId: return -1
        return eid.IntegerValue if hasattr(eid, "IntegerValue") else eid.Value
        
    matched_ids = set([safe_id(s.sheet_element_id) for s in slots if s.sheet_element_id != ElementId.InvalidElementId])
    
    # Any remaining unmapped slots are MISSING, unmapped sheets (within target collections) are EXTRA
    extra_sheets = [sh for sh in harvested_data.sheets 
                    if sh.sheet_collection in target_collections 
                    and safe_id(sh.element_id) not in matched_ids]
    
    return slots, extra_sheets

def resolve_renumbering_sequence(plan_rows):
    """
    Analyzes the plan_rows to find the safe order to renumber sheets without collisions.
    Returns:
      ordered_sequence: list of ops that can be executed safely in order.
      cycle_ops: list of ops that need a temporary rename to break cycles (swaps).
    """
    renumber_ops = [r for r in plan_rows if r['status'] in ("RENAME_NUMBER", "RENAME_BOTH") and r['sheet_element_id'] != -1]
    
    current_number_to_op = {}
    for r in plan_rows:
        if r.get('existing_number') and r.get('sheet_element_id') != -1:
            current_number_to_op[r['existing_number']] = r
            
    adj = {op['sheet_element_id']: [] for op in renumber_ops}
    in_degree = {op['sheet_element_id']: 0 for op in renumber_ops}
    op_by_id = {op['sheet_element_id']: op for op in renumber_ops}
    
    for op in renumber_ops:
        target_num = op['target_number']
        blocking_op = current_number_to_op.get(target_num)
        
        if blocking_op and blocking_op['sheet_element_id'] != op['sheet_element_id']:
            blocking_id = blocking_op['sheet_element_id']
            if blocking_id in op_by_id:
                # blocking_op must move BEFORE op
                adj[blocking_id].append(op['sheet_element_id'])
                in_degree[op['sheet_element_id']] += 1
                
    queue = [op_id for op_id in in_degree if in_degree[op_id] == 0]
    ordered_sequence = []
    
    while queue:
        curr_id = queue.pop(0)
        ordered_sequence.append(op_by_id[curr_id])
        for neighbor_id in adj[curr_id]:
            in_degree[neighbor_id] -= 1
            if in_degree[neighbor_id] == 0:
                queue.append(neighbor_id)
                
    cycle_ops = []
    if len(ordered_sequence) < len(renumber_ops):
        for op_id in in_degree:
            if in_degree[op_id] > 0:
                cycle_ops.append(op_by_id[op_id])
                
    return ordered_sequence, cycle_ops

def apply_plan(doc, plan_rows):
    """
    Phase 6: Three-Phase Application
    1. View Prep (Rename existing, Create new)
    2. Sheet Sequencing (Topological rename, create new)
    3. View Placement
    """
    with TransactionGroup(doc, "AIA Reconciliation") as tg:
        tg.Start()
        try:
            # --- PHASE 1: View Generation & Renaming ---
            with Transaction(doc, "Phase 1 - View Prep") as t1:
                t1.Start()
                # Placeholder for View Prep logic
                # E.g., rename existing views to [Desired] - [Collection] - [Detail] - [Sheet]
                # Create missing views based on Plan Type.
                t1.Commit()
                
            # --- PHASE 2: Sequence-Based Sheet Renaming & Creation ---
            with Transaction(doc, "Phase 2 - Sheet Sequencing") as t2:
                t2.Start()
                
                ordered_sequence, cycle_ops = resolve_renumbering_sequence(plan_rows)
                
                # 1. Break Cycles
                for op in cycle_ops:
                    sheet_id = ElementId(op['sheet_element_id'])
                    sheet = doc.GetElement(sheet_id)
                    if sheet and sheet.SheetNumber != op['target_number']:
                        sheet.SheetNumber = op['target_number'] + "_TEMP" 
                
                # 2. Execute Linear Sequence
                for op in ordered_sequence:
                    sheet_id = ElementId(op['sheet_element_id'])
                    sheet = doc.GetElement(sheet_id)
                    if sheet:
                        sheet.SheetNumber = op['target_number']
                        if op.get('target_name') and sheet.Name != op['target_name']:
                            sheet.Name = op['target_name']
                            
                # 3. Resolve Parked Cycles
                for op in cycle_ops:
                    sheet_id = ElementId(op['sheet_element_id'])
                    sheet = doc.GetElement(sheet_id)
                    if sheet:
                        sheet.SheetNumber = op['target_number']
                        if op.get('target_name') and sheet.Name != op['target_name']:
                            sheet.Name = op['target_name']

                # Apply names for other sheets
                rename_name_ops = [r for r in plan_rows if r['status'] in ("RENAME_NAME", "MATCH") and r['sheet_element_id'] != -1]
                for op in rename_name_ops:
                    sheet_id = ElementId(op['sheet_element_id'])
                    sheet = doc.GetElement(sheet_id)
                    if sheet and op.get('target_name') and sheet.Name != op['target_name']:
                        sheet.Name = op['target_name']
                        
                # Create missing sheets
                missing_ops = [r for r in plan_rows if r['status'] == "MISSING"]
                for op in missing_ops:
                    new_sheet = ViewSheet.CreatePlaceholder(doc)
                    new_sheet.SheetNumber = op['target_number']
                    new_sheet.Name = op['target_name']
                    try:
                        param = new_sheet.LookupParameter("Sheet Collection")
                        if param and op.get('collection') and op.get('collection') != "Undefined":
                            param.Set(op['collection'])
                    except:
                        pass
                
                t2.Commit()
                
            # --- PHASE 3: View Placement ---
            with Transaction(doc, "Phase 3 - View Placement") as t3:
                t3.Start()
                # Placeholder for View Placement logic
                # E.g., placing the prepped views onto the finalized sheets via Viewport.Create()
                t3.Commit()

            tg.Assimilate()
        except Exception as e:
            tg.RollBack()
            raise e

def generate_plan_json(harvested_data, matched_slots, extra_sheets):
    rows = []
    seq = 1
    for s in matched_slots:
        row = {
            "seq": seq,
            "target_number": s.target_number,
            "target_name": s.target_name,
            "status": s.status,
            "existing_number": s.existing_number,
            "existing_name": s.existing_name,
            "sheet_element_id": (s.sheet_element_id.IntegerValue if hasattr(s.sheet_element_id, "IntegerValue") else s.sheet_element_id.Value) if s.sheet_element_id != ElementId.InvalidElementId else -1,
            "view_action": s.view_action,
            "view_element_id": (s.view_element_id.IntegerValue if hasattr(s.view_element_id, "IntegerValue") else s.view_element_id.Value) if s.view_element_id != ElementId.InvalidElementId else -1,
            "view_recipe": s.view_recipe,
            "requires_user_decision": s.requires_user_decision,
            "match_score": s.match_score,
            "collection": s.collection,
            "discipline": s.discipline,
            "cg": s.cg,
            "series_name": s.series_name
        }
        rows.append(row)
        seq += 1
        
    for sh in extra_sheets:
        row = {
            "seq": seq,
            "target_number": "",
            "target_name": "",
            "status": "EXTRA",
            "existing_number": sh.number,
            "existing_name": sh.name,
            "sheet_element_id": (sh.element_id.IntegerValue if hasattr(sh.element_id, "IntegerValue") else sh.element_id.Value) if sh.element_id != ElementId.InvalidElementId else -1,
            "view_action": "NONE",
            "view_element_id": -1,
            "view_recipe": None,
            "requires_user_decision": False,
            "match_score": 0.0,
            "collection": sh.sheet_collection,
            "discipline": "Unknown",
            "cg": "Uncategorized",
            "series_name": getattr(sh, 'series_name', 'Unknown')
        }
        rows.append(row)
        seq += 1
        
    return {
        "schema_version": "1.0",
        "grid_size": harvested_data.grid_size,
        "rows": rows,
        "warnings": harvested_data.warnings
    }

def run_pipeline(doc, generated_targets, existing_sheet_ids=None):
    """
    Main entry point to run Phases 1-5
    """
    h_data = harvest_project(doc)
    
    if existing_sheet_ids is not None:
        h_data.sheets = [sh for sh in h_data.sheets if (sh.element_id.IntegerValue if hasattr(sh.element_id, "IntegerValue") else sh.element_id.Value) in existing_sheet_ids]
        
    slots = generate_slots(generated_targets)
    matched_slots, extra_sheets = match_sheets(h_data, slots)
    plan_json = generate_plan_json(h_data, matched_slots, extra_sheets)
    return plan_json

