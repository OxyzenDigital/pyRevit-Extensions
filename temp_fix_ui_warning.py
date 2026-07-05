import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_trigger = '''                                                    <DataTrigger Binding="{Binding IsNameUnique}" Value="False">
                                                        <Setter Property="Background" Value="{DynamicResource ErrorBrush}"/>
                                                        <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                                                        <Setter Property="ToolTip" Value="Duplicate Sheet Number!"/>
                                                    </DataTrigger>'''

new_trigger = '''                                                    <DataTrigger Binding="{Binding IsNameUnique}" Value="False">
                                                        <Setter Property="Background" Value="{DynamicResource ErrorBrush}"/>
                                                        <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                                                        <Setter Property="ToolTip" Value="{Binding ValidationWarning}"/>
                                                    </DataTrigger>'''

content = content.replace(old_trigger, new_trigger)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
