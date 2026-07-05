import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_logic = '''        self.update_grid_title()
        self.execute_schema_match(valid_sheets, active_collection=c_name)'''

new_logic = '''        self.update_grid_title()
        
        # Highlight active collection in NavTree
        for root_node in self.NavRoot:
            if c_name is None:
                root_node.IsActiveContext = True
            else:
                root_node.IsActiveContext = False
            for c_node in root_node.Children:
                c_node.IsActiveContext = (c_node.Tag == c_name)
                
        self.execute_schema_match(valid_sheets, active_collection=c_name)'''

content = content.replace(old_logic, new_logic)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
