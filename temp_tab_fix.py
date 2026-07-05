import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_tab_changed = r'''    def on_tab_changed\(self, sender, e\):
        if self\.MainTabControl\.SelectedIndex == 1:
            self\.FooterBorder\.Visibility = Visibility\.Visible
        else:
            self\.FooterBorder\.Visibility = Visibility\.Collapsed'''

new_tab_changed = r'''    def on_tab_changed(self, sender, e):
        if self.MainTabControl.SelectedIndex == 1:
            self.FooterBorder.Visibility = Visibility.Visible
            
            # Force a refresh of the grid based on the latest schema when switching tabs
            node = getattr(self, "_current_selected_node", None)
            if hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
                node = self.NavTree.SelectedItem
            elif self.NavRoot.Count > 0:
                node = self.NavRoot[0]
                
            if node:
                self._current_selected_node = node
                class DummyArgs: pass
                self.on_tree_selection_changed(self.NavTree, DummyArgs())
        else:
            self.FooterBorder.Visibility = Visibility.Collapsed'''

content = re.sub(old_tab_changed, new_tab_changed, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
