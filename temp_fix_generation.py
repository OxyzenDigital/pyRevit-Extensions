import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace "collection": "01. Plans" with "series_name": "01. Plans"
content = re.sub(r'"collection": "01\. Plans"', r'"series_name": "01. Plans"', content)
# Replace "collection": coll_name with "series_name": coll_name
content = re.sub(r'"collection": coll_name', r'"series_name": coll_name', content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
