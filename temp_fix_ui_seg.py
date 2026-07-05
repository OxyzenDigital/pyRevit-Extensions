import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace DataTrigger
old_trigger = '''                                                                        <Style.Triggers>
                                                                          <DataTrigger Binding="{Binding NodeType}" Value="Modifier">
                                                                              <Setter Property="Visibility" Value="Visible"/>
                                                                          </DataTrigger>
                                                                          <DataTrigger Binding="{Binding HasGridOverrides}" Value="True">
                                                                              <Setter Property="Foreground" Value="{DynamicResource AccentBrush}"/>
                                                                          </DataTrigger>
                                                                      </Style.Triggers>'''

new_trigger = '''                                                                        <Style.Triggers>
                                                                          <DataTrigger Binding="{Binding IsSegmentable}" Value="True">
                                                                              <Setter Property="Visibility" Value="Visible"/>
                                                                          </DataTrigger>
                                                                          <DataTrigger Binding="{Binding HasGridOverrides}" Value="True">
                                                                              <Setter Property="Foreground" Value="{DynamicResource AccentBrush}"/>
                                                                          </DataTrigger>
                                                                      </Style.Triggers>'''

content = content.replace(old_trigger, new_trigger)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
