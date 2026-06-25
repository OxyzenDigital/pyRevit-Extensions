# -*- coding: utf-8 -*-
import os
import json
import re

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "classification.json")

def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "DRAWING_TYPES": {},
        "CLASSIFICATION_DICT": {},
        "SHEET_TYPE_SYNONYMS": {},
        "NAMING_SCHEMES": {}
    }

config = load_config()
DRAWING_TYPES = config.get("DRAWING_TYPES", {})
CLASSIFICATION_DICT = config.get("CLASSIFICATION_DICT", {})
SHEET_TYPE_SYNONYMS = config.get("SHEET_TYPE_SYNONYMS", {})
NAMING_SCHEMES = config.get("NAMING_SCHEMES", {})

# Flattens the dict into a single lookup for fast sheet classification
# Key: lowercased sheet name
# Value: dict with keys: discipline, contentGroup, drawingType, drawingTypeCode
FLAT_CLASSIFICATION = {}
for disc, groups in CLASSIFICATION_DICT.items():
    for group, sheets in groups.items():
        for item in sheets:
            if len(item) >= 2:
                s_name = item[0]
                s_code = str(item[1])
                FLAT_CLASSIFICATION[s_name.lower()] = {
                    "discipline": disc,
                    "contentGroup": group,
                    "drawingType": DRAWING_TYPES.get(s_code, "Unknown"),
                    "drawingTypeCode": s_code,
                    "originalName": s_name
                }

def normalize_sheet_name(name):
    """Normalize sheet name for better fuzzy matching."""
    # Remove common prefixes like 'A-101 ' or 'A101 - '
    name = re.sub(r'^[A-Z]+\-?\d+[A-Z]?\s*[-|:]?\s*', '', name, flags=re.IGNORECASE)
    # Remove 'Overall', 'Enlarged', 'Level' to focus on core name if desired, or keep it.
    # We will just strip non-alphanumeric except spaces.
    name = re.sub(r'[^\w\s]', '', name)
    # Convert multiple spaces to single
    name = re.sub(r'\s+', ' ', name).strip().lower()
    return name

NORM_FLAT = {normalize_sheet_name(k): v for k, v in FLAT_CLASSIFICATION.items()}

def classify_sheet(sheet_name):
    lower_name = sheet_name.lower()
    norm_name = normalize_sheet_name(sheet_name)
    
    # 1. Exact match
    if lower_name in FLAT_CLASSIFICATION:
        return FLAT_CLASSIFICATION[lower_name]
        
    # 2. Normalized Exact match
    if norm_name in NORM_FLAT and norm_name:
        return NORM_FLAT[norm_name]
        
    # 3. Partial match (longest match wins)
    best_match = None
    best_len = 0
    
    # Try normalized partial match first
    for k, v in NORM_FLAT.items():
        if k and k in norm_name and len(k) > best_len:
            best_len = len(k)
            best_match = v
            
    if best_match:
        return best_match
        
    # Default fallback
    return {
        "discipline": "Unknown",
        "contentGroup": "Uncategorized",
        "drawingType": "Unknown",
        "drawingTypeCode": "99",
        "originalName": sheet_name
    }
