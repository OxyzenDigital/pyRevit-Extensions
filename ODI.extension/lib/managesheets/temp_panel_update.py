import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update load_revit_data to call generate_target_schema automatically
# and trigger root selection so the AIA schema populates the grid.
load_revit_data_end_pattern = r'(sh_row\.Views\.Add\(v_row\)\s+views \+= 1\s+)self\.run_validation\(\)'
new_load_revit_data_end = r'''\1self.run_validation()
        
        # Prepopulate AIA Schema
        self.generate_target_schema()
        # Automatically select the Root node to populate the initial grid
        self.NavTree.SelectedItemChanged -= self.on_tree_selection_changed
        self.NavTree.SelectedItemChanged += self.on_tree_selection_changed
        if self.NavRoot.Count > 0:
            self._fake_selection(self.NavRoot[0])
            
    def _fake_selection(self, node):
        class DummyArgs: pass
        self._current_selected_node = node
        self.on_tree_selection_changed(self.NavTree, DummyArgs())
'''
content = re.sub(load_revit_data_end_pattern, new_load_revit_data_end, content)


# 2. Rewrite on_tree_selection_changed and run_fuzzy_match into one.
on_tree_selection_changed_pattern = r'    def on_tree_selection_changed\(self, sender, e\):.*?(?=    def update_grid_title)'
new_on_tree_selection = r'''    def on_tree_selection_changed(self, sender, e):
        node = getattr(self, "_current_selected_node", None)
        if hasattr(self.NavTree, "SelectedItem") and self.NavTree.SelectedItem:
            node = self.NavTree.SelectedItem
            self._current_selected_node = node
            
        if not node: return
        
        valid_sheets = []
        if node.NodeType == "Root":
            # If root is selected, pass empty list so ONLY the raw AIA schema is shown (CREATE)
            valid_sheets = []
        elif node.NodeType == "Collection":
            c_name = node.Tag
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name]
        elif node.NodeType == "Discipline":
            c_name, d_name = node.Tag
            self.Txt_GridTitle.Text = "Discipline: " + d_name
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name and s.DisciplineName == d_name]
        elif node.NodeType == "ContentGroup":
            c_name, d_name, cg_name = node.Tag
            self.Txt_GridTitle.Text = "Group: " + cg_name
            valid_sheets = [s for s in self.all_grid_nodes if getattr(s, 'OriginalCollectionName', s.CollectionName) == c_name and s.DisciplineName == d_name and s.ContentGroupName == cg_name]
            
        self.update_grid_title()
        self.execute_schema_match(valid_sheets)

    def execute_schema_match(self, valid_sheets):
        if not getattr(self, "generated_targets", None): return
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        doc = uidoc.Document if uidoc else None
        if not doc: return
        
        try:
            import reconciliation
        except Exception:
            return
            
        valid_sheet_ids = []
        for s in valid_sheets:
            if hasattr(s.ElementId, "IntegerValue"):
                valid_sheet_ids.append(s.ElementId.IntegerValue)
            elif hasattr(s.ElementId, "Value"):
                valid_sheet_ids.append(s.ElementId.Value)
                
        plan = reconciliation.run_pipeline(doc, self.generated_targets, existing_sheet_ids=valid_sheet_ids)
        
        from Autodesk.Revit.DB import ElementId
        existing_nodes = {}
        for n in self.all_grid_nodes:
            if n.ElementId and n.ElementId != ElementId.InvalidElementId:
                try:
                    key = n.ElementId.IntegerValue if hasattr(n.ElementId, "IntegerValue") else n.ElementId.Value
                    existing_nodes[key] = n
                except Exception:
                    pass
                    
        self.EditorItems.Clear()
        
        for row in plan["rows"]:
            is_template = False
            if row["status"] == "MISSING":
                if self.Chk_AddMissing.IsChecked != True: continue
                is_template = True
            
            sh_id = row["sheet_element_id"]
            
            if sh_id in existing_nodes and sh_id != -1:
                vm = existing_nodes[sh_id]
                vm.CollectionName = row["collection"]
                vm.DisciplineName = row["discipline"]
                vm.ContentGroupName = row["cg"]
                if row["status"] != "UNRECONCILED" and row["status"] != "EXTRA":
                    vm.SheetNumber = row["target_number"]
                    vm.SheetName = row["target_name"]
                
                vm.move_up_callback = self.move_item_up
                vm.move_down_callback = self.move_item_down
                try:
                    vm.MoveUpCommand.RaiseCanExecuteChanged()
                    vm.MoveDownCommand.RaiseCanExecuteChanged()
                except:
                    pass
                
                vm.IsChecked = True
                vm._action = row["status"]
                self.EditorItems.Add(vm)
            else:
                real_id = ElementId(sh_id) if sh_id != -1 else ElementId.InvalidElementId
                vm = SheetViewModel(real_id, row["target_number"] if is_template else row["existing_number"], 
                                    row["target_name"] if is_template else row["existing_name"], 
                                    row["collection"], row["discipline"], row["cg"], 
                                    is_template=is_template, validation_callback=self.run_validation, 
                                    number_changed_callback=self.on_sheet_number_changed,
                                    move_up_callback=self.move_item_up, move_down_callback=self.move_item_down)
                vm.OriginalCollectionName = row["collection"]
                vm.IsChecked = True
                if row["status"] in ["CREATE", "MISSING"]:
                    vm._action = "CREATE"
                else:
                    vm._action = row["status"]
                self.EditorItems.Add(vm)

    def move_item_up(self, item):
        idx = self.EditorItems.IndexOf(item)
        if idx > 0:
            self.EditorItems.Move(idx, idx - 1)
            
    def move_item_down(self, item):
        idx = self.EditorItems.IndexOf(item)
        if idx >= 0 and idx < self.EditorItems.Count - 1:
            self.EditorItems.Move(idx, idx + 1)
            
'''
content = re.sub(on_tree_selection_changed_pattern, new_on_tree_selection, content, flags=re.DOTALL)

# 3. Remove the old run_fuzzy_match function completely.
run_fuzzy_match_pattern = r'    def run_fuzzy_match\(self, sender, e\):.*?(?=    def run_validation)'
content = re.sub(run_fuzzy_match_pattern, '', content, flags=re.DOTALL)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Panel updated!")
