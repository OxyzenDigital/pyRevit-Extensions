import re

file_path = r'c:\Users\radhi\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\pyRevit-Extensions\ODI.extension\lib\managesheets\ui.xaml'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add DetailsVisibility setter to DataGridRow style
old_rowstyle = '''                                <DataGrid.RowStyle>
                                    <Style TargetType="{x:Type DataGridRow}">
                                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>'''

new_rowstyle = '''                                <DataGrid.RowStyle>
                                    <Style TargetType="{x:Type DataGridRow}">
                                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                                        <Setter Property="DetailsVisibility" Value="{Binding IsExpanded, Converter={StaticResource BoolToVis}}"/>'''

content = content.replace(old_rowstyle, new_rowstyle)

# 2. Add RowDetailsTemplate and Expand Column
old_columns = '''                                <DataGrid.Columns>
                                    <DataGridTemplateColumn Header="Sync" Width="40">'''

new_columns = '''                                <DataGrid.RowDetailsTemplate>
                                    <DataTemplate>
                                        <Border Margin="30,5,10,10" Background="{DynamicResource ControlBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" CornerRadius="3" Padding="5">
                                            <StackPanel>
                                                <TextBlock Text="Views Placed on Sheet" FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}" Margin="0,0,0,5"/>
                                                <DataGrid ItemsSource="{Binding Views}" AutoGenerateColumns="False" HeadersVisibility="Column" Background="Transparent" BorderThickness="0" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="{DynamicResource BorderBrush}">
                                                    <DataGrid.Columns>
                                                        <DataGridTextColumn Header="View Name" Binding="{Binding Name}" IsReadOnly="True" Width="*" Foreground="{DynamicResource TextBrush}"/>
                                                        <DataGridTextColumn Header="Type" Binding="{Binding ViewType}" IsReadOnly="True" Width="100" Foreground="{DynamicResource TextLightBrush}"/>
                                                        <DataGridTextColumn Header="Level" Binding="{Binding LevelName}" IsReadOnly="True" Width="100" Foreground="{DynamicResource TextLightBrush}"/>
                                                        <DataGridTextColumn Header="Scale" Binding="{Binding Scale}" IsReadOnly="True" Width="80" Foreground="{DynamicResource TextLightBrush}"/>
                                                    </DataGrid.Columns>
                                                </DataGrid>
                                            </StackPanel>
                                        </Border>
                                    </DataTemplate>
                                </DataGrid.RowDetailsTemplate>
                                <DataGrid.Columns>
                                    <DataGridTemplateColumn Width="30">
                                        <DataGridTemplateColumn.CellTemplate>
                                            <DataTemplate>
                                                <ToggleButton IsChecked="{Binding IsExpanded, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Background="Transparent" BorderThickness="0" Foreground="{DynamicResource TextLightBrush}">
                                                    <ToggleButton.Style>
                                                        <Style TargetType="ToggleButton">
                                                            <Setter Property="Content" Value="▶"/>
                                                            <Style.Triggers>
                                                                <Trigger Property="IsChecked" Value="True">
                                                                    <Setter Property="Content" Value="▼"/>
                                                                </Trigger>
                                                                <DataTrigger Binding="{Binding Views.Count}" Value="0">
                                                                    <Setter Property="Visibility" Value="Hidden"/>
                                                                </DataTrigger>
                                                            </Style.Triggers>
                                                        </Style>
                                                    </ToggleButton.Style>
                                                </ToggleButton>
                                            </DataTemplate>
                                        </DataGridTemplateColumn.CellTemplate>
                                    </DataGridTemplateColumn>
                                    <DataGridTemplateColumn Header="Sync" Width="40">'''

content = content.replace(old_columns, new_columns)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
