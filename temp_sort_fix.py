import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_func_pattern = r'    def on_sheet_number_changed\(self, sheet, old_val, new_val\):.*?(?=    def run_validation)'

new_func = r'''    def on_sheet_number_changed(self, sheet, old_val, new_val):
        if getattr(self, "_is_auto_sequencing", False):
            return
            
        self._is_auto_sequencing = True
        try:
            idx = -1
            for i, r in enumerate(self.EditorItems):
                if r == sheet:
                    idx = i
                    break
                    
            if idx != -1 and idx < len(self.EditorItems) - 1:
                match_new = re.search(r'(\d+)$', new_val)
                if match_new:
                    new_prefix = new_val[:match_new.start()]
                    num_str = match_new.group(1)
                    num_len = len(num_str)
                    current_num = int(num_str)
                    
                    for i in range(idx + 1, len(self.EditorItems)):
                        next_sheet = self.EditorItems[i]
                        # Only auto-sequence if it belongs to the exact same Discipline and Group
                        if next_sheet.DisciplineName == sheet.DisciplineName and next_sheet.ContentGroupName == sheet.ContentGroupName:
                            current_num += 1
                            next_val = "{}{:0{}d}".format(new_prefix, current_num, num_len)
                            next_sheet.SheetNumber = next_val
                            
            # Sort the EditorItems list by SheetNumber
            items_list = list(self.EditorItems)
            
            def get_sort_key(s):
                num = s.SheetNumber if s.SheetNumber else ""
                # Pad numbers so string sorting works correctly (e.g. A-010 before A-100)
                # We split the string by numbers and text
                parts = re.split(r'(\d+)', num)
                key = []
                for p in parts:
                    if p.isdigit():
                        key.append(p.zfill(10))
                    else:
                        key.append(p.lower())
                return key
                
            items_list.sort(key=get_sort_key)
            
            self.EditorItems.Clear()
            for item in items_list:
                self.EditorItems.Add(item)
                
        finally:
            self._is_auto_sequencing = False
            self.run_validation()

'''

content = re.sub(old_func_pattern, new_func, content, flags=re.DOTALL)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
