import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\data_model.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace undo_changes with toggle_changes
old_undo = r'    def undo_changes\(self, parameter=None\):.*?if self\.validation_callback: self\.validation_callback\(\)'

new_toggle = r'''    @property
    def ToggleText(self):
        return "Redo" if getattr(self, '_is_reverted_to_original', False) else "Undo"
        
    def undo_changes(self, parameter=None): # Keeping name as undo_changes for existing bindings if any, but acting as toggle
        if not hasattr(self, '_is_reverted_to_original'):
            self._is_reverted_to_original = False
            self.ProposedNumber = self.SheetNumber
            self.ProposedName = self.SheetName
            
        if self._is_reverted_to_original:
            self.SheetNumber = getattr(self, 'ProposedNumber', self.SheetNumber)
            self.SheetName = getattr(self, 'ProposedName', self.SheetName)
            self._is_reverted_to_original = False
        else:
            self.ProposedNumber = self.SheetNumber
            self.ProposedName = self.SheetName
            self.SheetNumber = self.OriginalNumber
            self.SheetName = self.OriginalName
            self._is_reverted_to_original = True
            
        self.OnPropertyChanged("ToggleText")
        if self.validation_callback: self.validation_callback()'''

content = re.sub(old_undo, new_toggle, content, flags=re.DOTALL)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
