import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\data_model.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add _is_segmentable to init
old_init = '''        self._is_updating = False
        self._grid_overrides = None'''
new_init = '''        self._is_updating = False
        self._grid_overrides = None
        self._is_segmentable = False'''

content = content.replace(old_init, new_init)

# Add IsSegmentable property
prop = '''
    @property
    def GridOverrides(self):
        return self._grid_overrides
        
    @GridOverrides.setter
    def GridOverrides(self, value):
        self._grid_overrides = value
        self.OnPropertyChanged("GridOverrides")
        self.OnPropertyChanged("HasGridOverrides")
        
    @property
    def IsSegmentable(self):
        return self._is_segmentable
        
    @IsSegmentable.setter
    def IsSegmentable(self, value):
        self._is_segmentable = value
        self.OnPropertyChanged("IsSegmentable")
'''
content = content.replace('''
    @property
    def GridOverrides(self):
        return self._grid_overrides
        
    @GridOverrides.setter
    def GridOverrides(self, value):
        self._grid_overrides = value
        self.OnPropertyChanged("GridOverrides")
        self.OnPropertyChanged("HasGridOverrides")''', prop)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
