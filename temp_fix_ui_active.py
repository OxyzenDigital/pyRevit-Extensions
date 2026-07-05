import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

old_tb = '''                                              <StackPanel Orientation="Horizontal" Margin="0,2">
                                                <TextBlock Text="{Binding Name}" FontWeight="{Binding FontWeight}" VerticalAlignment="Center" TextWrapping="Wrap"/>
                                                <TextBlock Text="{Binding DisplayCount}" Foreground="{DynamicResource TextLightBrush}" FontSize="10" Margin="5,0,0,0" VerticalAlignment="Center" Visibility="{Binding ShowCount, Converter={StaticResource BoolToVis}}"/>'''

new_tb = '''                                              <StackPanel Orientation="Horizontal" Margin="0,2">
                                                <TextBlock Text="{Binding Name}" FontWeight="{Binding FontWeight}" VerticalAlignment="Center" TextWrapping="Wrap">
                                                    <TextBlock.Style>
                                                        <Style TargetType="TextBlock">
                                                            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                                                            <Style.Triggers>
                                                                <DataTrigger Binding="{Binding IsActiveContext}" Value="True">
                                                                    <Setter Property="Foreground" Value="{DynamicResource AccentBrush}"/>
                                                                </DataTrigger>
                                                            </Style.Triggers>
                                                        </Style>
                                                    </TextBlock.Style>
                                                </TextBlock>
                                                <TextBlock Text="{Binding DisplayCount}" Foreground="{DynamicResource TextLightBrush}" FontSize="10" Margin="5,0,0,0" VerticalAlignment="Center" Visibility="{Binding ShowCount, Converter={StaticResource BoolToVis}}"/>'''

content = content.replace(old_tb, new_tb)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
