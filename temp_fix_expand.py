import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\data_model.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add _is_expanded to SheetViewModel
old_init = '''        self.IsChecked = False
        self._action = "MATCHED"'''
new_init = '''        self.IsChecked = False
        self._action = "MATCHED"
        self._is_expanded = False'''

content = content.replace(old_init, new_init)

# Add IsExpanded property
prop = '''
    @property
    def IsChecked(self): return self._is_checked
    
    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value
        self.OnPropertyChanged("IsChecked")
        
    @property
    def IsExpanded(self): return self._is_expanded
    
    @IsExpanded.setter
    def IsExpanded(self, value):
        self._is_expanded = value
        self.OnPropertyChanged("IsExpanded")
'''
content = content.replace('''
    @property
    def IsChecked(self): return self._is_checked
    
    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value
        self.OnPropertyChanged("IsChecked")''', prop)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
