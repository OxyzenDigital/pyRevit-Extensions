import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\data_model.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add _is_active_context to NavTreeNode init
old_init = '''        self._is_segmentable = False'''
new_init = '''        self._is_segmentable = False
        self._is_active_context = False'''
content = content.replace(old_init, new_init)

# Add IsActiveContext property
prop = '''
    @property
    def IsActiveContext(self):
        return self._is_active_context
        
    @IsActiveContext.setter
    def IsActiveContext(self, value):
        self._is_active_context = value
        self.OnPropertyChanged("IsActiveContext")
        self.OnPropertyChanged("FontWeight")
'''
content = content.replace('''
    @property
    def IsSegmentable(self):''', prop + '''
    @property
    def IsSegmentable(self):''')

# Update FontWeight property
old_fw = '''    @property
    def FontWeight(self): return "Bold" if self.NodeType in ["Root", "Collection"] else "Normal"'''
new_fw = '''    @property
    def FontWeight(self):
        if self.NodeType == "Root": return "Bold"
        if self.NodeType == "Collection": return "Bold" if self.IsActiveContext else "Normal"
        return "Normal"'''
content = content.replace(old_fw, new_fw)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
