import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Update on_tree_selection_changed to pass c_name
old_tree = r'''        valid_sheets = \[\]
        if node\.NodeType == "Root":
            # If root is selected, pass empty list so ONLY the raw AIA schema is shown \(CREATE\)
            valid_sheets = \[\]
        elif node\.NodeType == "Collection":
            c_name = node\.Tag
            valid_sheets = \[s for s in self\.all_grid_nodes if getattr\(s, 'OriginalCollectionName', s\.CollectionName\) == c_name\]
        elif node\.NodeType == "Discipline":
            c_name, d_name = node\.Tag
            self\.Txt_GridTitle\.Text = "Discipline: " \+ d_name
            valid_sheets = \[s for s in self\.all_grid_nodes if getattr\(s, 'OriginalCollectionName', s\.CollectionName\) == c_name and s\.DisciplineName == d_name\]
        elif node\.NodeType == "ContentGroup":
            c_name, d_name, cg_name = node\.Tag
            self\.Txt_GridTitle\.Text = "Group: " \+ cg_name
            valid_sheets = \[s for s in self\.all_grid_nodes if getattr\(s, 'OriginalCollectionName', s\.CollectionName\) == c_name and s\.DisciplineName == d_name and s\.ContentGroupName == cg_name\]
            
        self\.update_grid_title\(\)
        
        self\.execute_schema_match\(valid_sheets\)'''

new_tree = r'''        valid_sheets = []
        c_name = "Default"
        if node.NodeType == "Root":
            # If root is selected, pass empty list so ONLY the raw AIA schema is shown (CREATE)
            valid_sheets = []
            self.Txt_GridTitle.Text = "All Generated Sheets"
        elif node.NodeType == "Collection":
            c_name = node.Tag
            self.Txt_GridTitle.Text = "Collection: " + c_name
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name]
        elif node.NodeType == "Discipline":
            c_name, d_name = node.Tag
            self.Txt_GridTitle.Text = "Collection: {} | Discipline: {}".format(c_name, d_name)
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name and s.DisciplineName == d_name]
        elif node.NodeType == "ContentGroup":
            c_name, d_name, cg_name = node.Tag
            self.Txt_GridTitle.Text = "Collection: {} | Group: {}".format(c_name, cg_name)
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name and s.DisciplineName == d_name and s.ContentGroupName == cg_name]
            
        self.update_grid_title()
        
        self.execute_schema_match(valid_sheets, active_collection=c_name)'''

content = re.sub(old_tree, new_tree, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
