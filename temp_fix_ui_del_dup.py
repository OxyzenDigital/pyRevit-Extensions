import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Remove the second RowDetailsTemplate block
pattern = r'<!-- Nested DataGrid for Views on Sheet -->\s*<DataGrid\.RowDetailsTemplate>.*?</DataGrid\.RowDetailsTemplate>'
content = re.sub(pattern, '', content, flags=re.DOTALL)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
