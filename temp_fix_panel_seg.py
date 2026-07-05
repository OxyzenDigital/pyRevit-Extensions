import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\panel.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add IsSegmentable when creating m_node
old_node = r'''                        m_node = NavTreeNode\(m_name, "Modifier", tag=m_code, parent=cg_node\)
                        m_node\.callback = self\.generate_target_schema
                        m_node\.grid_config_callback = self\.open_grid_config
                        m_node\.GridOverrides = modifier_grid_overrides\.get\(m_name, None\)
                        # Bypass the IsChecked setter logic temporarily to avoid mass triggering during load
                        m_node\._is_checked = m_name in saved_modifiers
                        cg_node\.Children\.Add\(m_node\)'''

new_node = r'''                        m_node = NavTreeNode(m_name, "Modifier", tag=m_code, parent=cg_node)
                        m_node.callback = self.generate_target_schema
                        m_node.grid_config_callback = self.open_grid_config
                        m_node.GridOverrides = modifier_grid_overrides.get(m_name, None)
                        m_node.IsSegmentable = ("Plan" in cg) or ("Plan" in m_name)
                        # Bypass the IsChecked setter logic temporarily to avoid mass triggering during load
                        m_node._is_checked = m_name in saved_modifiers
                        cg_node.Children.Add(m_node)'''

content = re.sub(old_node, new_node, content)


old_custom = r'''            for mod in custom_modifiers:
                m_node = NavTreeNode\(mod, "Modifier", tag="9", parent=custom_node\)
                m_node\.callback = self\.generate_target_schema
                m_node\.grid_config_callback = self\.open_grid_config
                m_node\.GridOverrides = modifier_grid_overrides\.get\(mod, None\)
                m_node\._is_checked = True 
                custom_node\.Children\.Add\(m_node\)'''

new_custom = r'''            for mod in custom_modifiers:
                m_node = NavTreeNode(mod, "Modifier", tag="9", parent=custom_node)
                m_node.callback = self.generate_target_schema
                m_node.grid_config_callback = self.open_grid_config
                m_node.GridOverrides = modifier_grid_overrides.get(mod, None)
                m_node.IsSegmentable = True
                m_node._is_checked = True 
                custom_node.Children.Add(m_node)'''

content = re.sub(old_custom, new_custom, content)


# Also in the other load_modifiers_from_cfg block:
old_custom2 = r'''            for mod in custom_modifiers:
                m_node = NavTreeNode\(mod, "Modifier", tag="9", parent=custom_node\)
                m_node\.callback = self\.generate_target_schema
                m_node\.grid_config_callback = self\.open_grid_config
                m_node\.GridOverrides = modifier_grid_overrides\.get\(mod, None\)
                m_node\._is_checked = True # Active by definition if it's in this list
                custom_node\.Children\.Add\(m_node\)'''

new_custom2 = r'''            for mod in custom_modifiers:
                m_node = NavTreeNode(mod, "Modifier", tag="9", parent=custom_node)
                m_node.callback = self.generate_target_schema
                m_node.grid_config_callback = self.open_grid_config
                m_node.GridOverrides = modifier_grid_overrides.get(mod, None)
                m_node.IsSegmentable = True
                m_node._is_checked = True # Active by definition if it's in this list
                custom_node.Children.Add(m_node)'''

content = re.sub(old_custom2, new_custom2, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
