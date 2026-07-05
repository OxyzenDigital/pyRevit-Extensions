import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Add Sheet Series column before Sheet Collection
new_col = '''                                    <DataGridTextColumn Header="Sheet Series" Binding="{Binding SheetSeries}" IsReadOnly="True" ElementStyle="{StaticResource DarkGridTextLightStyle}" Width="100"/>
                                    <DataGridTextColumn Header="Sheet Collection" Binding="{Binding CollectionName}"'''

content = content.replace('<DataGridTextColumn Header="Sheet Collection" Binding="{Binding CollectionName}"', new_col)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
