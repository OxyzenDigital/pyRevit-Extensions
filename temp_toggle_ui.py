import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = re.sub(r'<Button Content="Undo" Command="\{Binding UndoCommand\}"', r'<Button Content="{Binding ToggleText}" Command="{Binding UndoCommand}"', content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
