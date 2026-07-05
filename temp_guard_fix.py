import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

def inject_guard(method_name, content_str):
    pattern = r'(    def ' + method_name + r'\(self, sender, e\):)\n'
    replacement = r'\1\n        if hasattr(e, "OriginalSource") and e.OriginalSource != sender: return\n'
    return re.sub(pattern, replacement, content_str)

content = inject_guard('on_scheme_selected', content)
content = inject_guard('on_disc_selected', content)
content = inject_guard('on_mod_selected', content)
content = inject_guard('on_tree_selection_changed', content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
