import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\reconciliation.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace get_matches and collection processing
old_match = r'        # Pass D: Fuzzy match(.*?)extra_sheets = \[sh for sh in harvested_data.sheets'

new_match = r'''        # Pass D: Fuzzy match
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
                
                # Base word overlap
                name_score = token_set_ratio(norm_slot, sh_name_norm)
                
                # Subset detection (e.g. "dimension plan floor" is a subset of "dimension plan level 01 overall")
                w1 = set(norm_slot.split())
                w2 = set(sh_name_norm.split())
                
                # Core Keyword Matching
                core_w1 = w1.intersection(core_keywords)
                core_w2 = w2.intersection(core_keywords)
                if core_w1 and core_w1 == core_w2:
                    name_score += 0.2
                
                if len(w2) >= 2 and w2.issubset(w1):
                    name_score += 0.3 # Heavy boost if existing sheet name is fully contained in target
                
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
                slots_subset.remove(s)
                sheets_subset.remove(best_sh)
    
    # Pass 1: Strict Same Collection Matching
    collections = set(sh.sheet_collection for sh in unmatched_sheets)
    collections.add("Default")
    
    for coll in collections:
        coll_slots = [s for s in unmatched_slots if s.collection == coll]
        coll_sheets = [sh for sh in unmatched_sheets if sh.sheet_collection == coll]
        get_matches(coll_slots, coll_sheets)

    # Pass 2: Cross-Collection Migration Matching (For sheets moving from 'Default' to organized collections)
    get_matches(unmatched_slots, unmatched_sheets)

    # Any remaining unmapped slots are MISSING, unmapped sheets are EXTRA
    extra_sheets = [sh for sh in harvested_data.sheets'''

content = re.sub(old_match, new_match, content, flags=re.DOTALL)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
