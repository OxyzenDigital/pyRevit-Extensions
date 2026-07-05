import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\data_model.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add _is_expanded to SheetViewModel init
old_init = '''        self._is_name_unique = True
        self._validation_warning = ""
        self._action = "MATCHED"'''
new_init = '''        self._is_name_unique = True
        self._validation_warning = ""
        self._action = "MATCHED"
        self._is_expanded = False'''
content = content.replace(old_init, new_init)

# Add IsExpanded property
prop = '''
    @property
    def IsNameUnique(self): return self._is_name_unique
    
    @IsNameUnique.setter
    def IsNameUnique(self, value):
        self._is_name_unique = value
        self.OnPropertyChanged("IsNameUnique")
        
    @property
    def IsExpanded(self): return self._is_expanded
    
    @IsExpanded.setter
    def IsExpanded(self, value):
        self._is_expanded = value
        self.OnPropertyChanged("IsExpanded")
'''
content = content.replace('''
    @property
    def IsNameUnique(self): return self._is_name_unique
    
    @IsNameUnique.setter
    def IsNameUnique(self, value):
        self._is_name_unique = value
        self.OnPropertyChanged("IsNameUnique")''', prop)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
