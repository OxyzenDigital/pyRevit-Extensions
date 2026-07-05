import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\data_model.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add sheet_series to init signature
content = content.replace(
    'def __init__(self, element_id, number, name, collection_name, discipline_name="Unknown", content_group_name="Uncategorized", is_template=False',
    'def __init__(self, element_id, number, name, collection_name, discipline_name="Unknown", content_group_name="Uncategorized", series_name="Unknown", is_template=False'
)

# Add self.SheetSeries assignment
old_assign = '''        self._collection_name = collection_name
        self._discipline_name = discipline_name
        self._content_group_name = content_group_name
        self.OriginalNumber = number'''
new_assign = '''        self._collection_name = collection_name
        self._discipline_name = discipline_name
        self._content_group_name = content_group_name
        self.SheetSeries = series_name
        self.OriginalNumber = number'''

content = content.replace(old_assign, new_assign)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
