import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_logic = '''        has_error = False
        for num, items in all_numbers.items():
            if len(items) > 1:
                has_error = True
                for i in items: i.IsNameUnique = False
            else:
                for i in items: i.IsNameUnique = True'''

new_logic = '''        has_error = False
        for num, items in all_numbers.items():
            if len(items) > 1:
                has_error = True
                for i in items:
                    i.IsNameUnique = False
                    # Generate conflict message showing what it collides with
                    conflicts = []
                    for other in items:
                        if other != i:
                            name = other.SheetName or "Unnamed"
                            coll = getattr(other, "CollectionName", "Default")
                            conflicts.append("'{}' ({})".format(name, coll))
                    i.ValidationWarning = "Duplicate number! Conflicts with: " + ", ".join(conflicts)
            else:
                for i in items:
                    i.IsNameUnique = True
                    i.ValidationWarning = ""'''

content = content.replace(old_logic, new_logic)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
