import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(
    'MessageBox.Show("Sync Complete!\\nRenamed: {}\\nCreated: {}\\nPurged: {}".format(renames, creates, purges), "Success")',
    'MessageBox.Show("Sync Complete!\\nProcessed/Renamed: {}\\nCreated: {}".format(renames, creates), "Success")'
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
