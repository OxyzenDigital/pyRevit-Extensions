import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace UpdateSourceTrigger=PropertyChanged for SheetNumber and SheetName with LostFocus
pattern = r'Text="\{Binding SheetNumber, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged\}"'
replacement = r'Text="{Binding SheetNumber, Mode=TwoWay, UpdateSourceTrigger=LostFocus}"'
content = re.sub(pattern, replacement, content)

pattern2 = r'Text="\{Binding SheetName, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged\}"'
replacement2 = r'Text="{Binding SheetName, Mode=TwoWay, UpdateSourceTrigger=LostFocus}"'
content = re.sub(pattern2, replacement2, content)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
