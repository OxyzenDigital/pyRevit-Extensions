import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_logic = '''            if idx != -1 and idx < len(self.EditorItems) - 1:
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
                            next_sheet.SheetNumber = next_val'''

new_logic = '''            if idx != -1 and idx < len(self.EditorItems) - 1:
                match_new = re.search(r'(\d+)$', new_val)
                match_old = re.search(r'(\d+)$', old_val)
                if match_new and match_old:
                    new_prefix = new_val[:match_new.start()]
                    new_num = int(match_new.group(1))
                    new_len = len(match_new.group(1))
                    
                    old_prefix = old_val[:match_old.start()]
                    old_num = int(match_old.group(1))
                    
                    delta = new_num - old_num
                    cascaded = False
                    
                    for i in range(idx + 1, len(self.EditorItems)):
                        next_sheet = self.EditorItems[i]
                        # Only auto-sequence if it belongs to the exact same Discipline and Group
                        if next_sheet.DisciplineName != sheet.DisciplineName or next_sheet.ContentGroupName != sheet.ContentGroupName:
                            continue
                            
                        ns_num_str = next_sheet.SheetNumber
                        if not ns_num_str: continue
                        
                        match_ns = re.search(r'(\d+)$', ns_num_str)
                        if not match_ns: continue
                            
                        ns_prefix = ns_num_str[:match_ns.start()]
                        ns_num = int(match_ns.group(1))
                        
                        if ns_prefix != old_prefix:
                            continue
                            
                        if not cascaded:
                            if new_prefix == old_prefix and new_num < ns_num:
                                break
                            cascaded = True
                            
                        proposed_num = ns_num + delta
                        next_sheet.SheetNumber = "{}{:0{}d}".format(new_prefix, proposed_num, new_len)'''

content = content.replace(old_logic, new_logic)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
