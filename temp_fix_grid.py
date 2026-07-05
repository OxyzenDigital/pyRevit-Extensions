import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Remove the Chk_AddMissing filter
old_missing = r'''            if row\["status"\] == "MISSING":
                if self\.Chk_AddMissing\.IsChecked != True: continue
                is_template = True'''
new_missing = r'''            if row["status"] == "MISSING":
                is_template = True'''
content = re.sub(old_missing, new_missing, content)

# Add update_grid_title at the end of execute_schema_match
old_end = r'''            else:
                    vm\._action = row\["status"\]
                self\.EditorItems\.Add\(vm\)'''
new_end = r'''            else:
                    vm._action = row["status"]
                self.EditorItems.Add(vm)
        
        self.update_grid_title()'''
content = re.sub(old_end, new_end, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
