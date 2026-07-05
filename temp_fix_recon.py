import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\reconciliation.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace collection with series_name in SlotRecord init
content = content.replace('def __init__(self, target_number, target_name, discipline, series, ordinal, suffix, canonical_name, collection="Default", cg="Unknown"):', 
'def __init__(self, target_number, target_name, discipline, series, ordinal, suffix, canonical_name, collection="Default", series_name="Unknown", cg="Unknown"):')

content = content.replace('self.collection = collection', 'self.collection = collection\n        self.series_name = series_name')

content = content.replace('collection=t.get("collection", "Default")', 'collection=t.get("collection", "Default"),\n                series_name=t.get("series_name", "Unknown")')

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
