import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_tab_changed = r'''    def on_tab_changed\(self, sender, e\):
        if self\.MainTabControl\.SelectedIndex == 1:'''

new_tab_changed = r'''    def on_tab_changed(self, sender, e):
        # Ignore SelectionChanged events that bubble up from child controls (like DataGrid or ComboBox)
        if e.OriginalSource != self.MainTabControl:
            return
            
        if self.MainTabControl.SelectedIndex == 1:'''

content = re.sub(old_tab_changed, new_tab_changed, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
