import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\reconciliation.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_s = '''            "collection": s.collection,
            "discipline": s.discipline,
            "cg": s.cg'''

new_s = '''            "collection": s.collection,
            "discipline": s.discipline,
            "cg": s.cg,
            "series_name": s.series_name'''

content = content.replace(old_s, new_s)

old_ex = '''            "collection": ex.sheet_collection,
            "discipline": "Unknown",
            "cg": "Unknown"'''

new_ex = '''            "collection": ex.sheet_collection,
            "discipline": "Unknown",
            "cg": "Unknown",
            "series_name": "Unknown"'''

content = content.replace(old_ex, new_ex)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
