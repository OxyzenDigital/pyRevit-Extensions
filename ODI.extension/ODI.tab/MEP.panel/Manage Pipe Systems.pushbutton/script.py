# -*- coding: utf-8 -*-
"""
Title: Manage Pipe Systems
Description: Visualize, Diagnose, and Merge disconnected pipe networks using a modeless WPF interface.
Author: OXYZEN Digital
"""
import clr
import sys
import System
from System.Collections.Generic import List
from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Windows import Window, Application

# Revit API Imports
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
    OverrideGraphicSettings,
    Color,
    XYZ,
    ConnectorProfileType,
    Domain,
    ElementId,
    MepSystem,
    Plumbing
)
from Autodesk.Revit.UI import Selection

# pyRevit Imports
from pyrevit import revit, forms, script

# --- Context ---
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# --- XAML Interface (Embedded) ---
XAML_CONTENT = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Pipe Network Manager" Height="500" Width="800"
        WindowStartupLocation="CenterScreen"
        Background="#2D2D30">
    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Background" Value="#3E3E42"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="RowBackground" Value="#252526"/>
            <Setter Property="AlternatingRowBackground" Value="#1E1E1E"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Title/Header -->
            <RowDefinition Height="*"/>    <!-- DataGrid -->
            <RowDefinition Height="Auto"/> <!-- Controls -->
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Margin="0,0,0,10">
            <Label Content="Network Diagnosis &amp; Merge" FontSize="16"/>
            <TextBlock Text="Select multiple pipes to identify disconnected 'islands' and merge them." 
                       Foreground="#AAAAAA" Margin="5,0"/>
        </StackPanel>

        <!-- Network List -->
        <DataGrid x:Name="dgNetworks" Grid.Row="1" SelectionMode="Extended">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Network ID" Binding="{Binding Id}" Width="80"/>
                <DataGridTextColumn Header="System Name" Binding="{Binding SystemName}" Width="150"/>
                <DataGridTextColumn Header="Elements" Binding="{Binding ElementCount}" Width="80"/>
                <DataGridTextColumn Header="Total Length (ft)" Binding="{Binding TotalLength}" Width="100"/>
                <DataGridTextColumn Header="Volume (ft³)" Binding="{Binding TotalVolume}" Width="100"/>
                <DataGridTextColumn Header="Base Equip" Binding="{Binding BaseEquipment}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Action Panel -->
        <Grid Grid.Row="2" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <StackPanel Orientation="Horizontal" Grid.Column="0">
                <Button x:Name="btnSelect" Content="1. Select Pipes"/>
                <Button x:Name="btnAnalyze" Content="2. Analyze Selection"/>
            </StackPanel>
            
            <StackPanel Orientation="Horizontal" Grid.Column="2">
                <Button x:Name="btnVisualize" Content="3. Visualize (Color Splash)" Background="#007ACC"/>
                <Button x:Name="btnMerge" Content="4. Fix / Merge Selected" Background="#CA5100"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"""

# --- Network Logic ---

class NetworkIsland:
    """Represents a connected 'island' of piping elements."""
    def __init__(self, id, elements):
        self.Id = id
        self.Elements = elements
        self.ElementCount = len(elements)
        self.TotalLength = 0.0
        self.TotalVolume = 0.0
        self.SystemName = "Undefined"
        self.BaseEquipment = "None"
        self.SystemType = None
        self._calculate_metrics()

    def _calculate_metrics(self):
        """Calculates physical properties of the network island."""
        unique_systems = set()
        
        for el in self.Elements:
            # Length Calculation (Pipes only)
            if isinstance(el, Plumbing.Pipe):
                try:
                    length = el.get_Parameter(BuiltInCategory.CURVE_ELEM_LENGTH).AsDouble()
                    self.TotalLength += length
                    
                    # Volume Calculation (Approximate cylinder)
                    # Get Outer Diameter or Inner Diameter if available
                    # Using Inner Diameter for fluid volume is better, but Outer is standard for 'size'
                    d_param = el.get_Parameter(BuiltInCategory.RBS_PIPE_INNER_DIAM_PARAM)
                    if d_param:
                        d = d_param.AsDouble()
                        area = 3.14159 * (d / 2.0)**2
                        self.TotalVolume += area * length
                except:
                    pass

            # System Identification
            # Check the "System Type" parameter or the System Name
            sys_param = el.get_Parameter(BuiltInCategory.RBS_PIPING_SYSTEM_TYPE_PARAM)
            if sys_param and sys_param.AsElementId().IntegerValue != -1:
                unique_systems.add(doc.GetElement(sys_param.AsElementId()).Name)
            
            # Base Equipment Check (Logic for Pressure Systems)
            # This is tricky without fully traversing the system object, 
            # but we can check if any element IS equipment or connected to it.
            # Simplified: If the element is part of a system with base equipment.
            if hasattr(el, "MEPSystem"):
                mep_sys = el.MEPSystem
                if mep_sys and mep_sys.BaseEquipment:
                    self.BaseEquipment = mep_sys.BaseEquipment.Name

        if unique_systems:
            self.SystemName = ", ".join(unique_systems)
        
        # Rounding for UI
        self.TotalLength = round(self.TotalLength, 2)
        self.TotalVolume = round(self.TotalVolume, 2)


def get_connected_elements(start_element):
    """
    Recursively finds all physically connected MEP elements (BFS).
    Returns a set of Element IDs.
    """
    visited_ids = set()
    queue = [start_element.Id]
    visited_ids.add(start_element.Id)
    
    elements = []

    while len(queue) > 0:
        current_id = queue.pop(0)
        current_el = doc.GetElement(current_id)
        if not current_el: continue
        
        elements.append(current_el)

        # Get Connectors
        # For Pipes/Fittings/Accessories
        connectors = None
        try:
            if hasattr(current_el, "ConnectorManager"):
                connectors = current_el.ConnectorManager.Connectors # MEP Curves
            elif hasattr(current_el, "MEPModel") and current_el.MEPModel:
                connectors = current_el.MEPModel.ConnectorManager.Connectors # Fittings
        except Exception:
            continue

        if not connectors:
            continue

        for conn in connectors:
            # Check what this connector is connected to
            # 'AllRefs' returns connectors on other elements connected to this one
            for ref_conn in conn.AllRefs:
                owner = ref_conn.Owner
                if owner and owner.Id not in visited_ids:
                    # Filter by category to avoid jumping to unintended categories if needed
                    # For now, we accept all MEP stuff
                    if owner.Category.IsId(BuiltInCategory.OST_PipeCurves) or \
                       owner.Category.IsId(BuiltInCategory.OST_PipeFitting) or \
                       owner.Category.IsId(BuiltInCategory.OST_PipeAccessory) or \
                       owner.Category.IsId(BuiltInCategory.OST_MechanicalEquipment) or \
                       owner.Category.IsId(BuiltInCategory.OST_PlumbingFixtures):
                        
                        visited_ids.add(owner.Id)
                        queue.append(owner.Id)
    
    return elements


# --- UI Class ---

class PipeNetworkWindow(Window):
    def __init__(self):
        self.selected_elements = []
        self.network_islands = []
        
        # Parse XAML
        reader = StringReader(XAML_CONTENT)
        self.Content = XamlReader.Load(reader)
        
        # Bind Controls
        self.dgNetworks = self.FindName("dgNetworks")
        
        self.btnSelect = self.FindName("btnSelect")
        self.btnSelect.Click += self.select_pipes
        
        self.btnAnalyze = self.FindName("btnAnalyze")
        self.btnAnalyze.Click += self.analyze_networks
        
        self.btnVisualize = self.FindName("btnVisualize")
        self.btnVisualize.Click += self.visualize_networks
        
        self.btnMerge = self.FindName("btnMerge")
        self.btnMerge.Click += self.merge_networks

    def select_pipes(self, sender, args):
        try:
            # Pick Objects
            refs = uidoc.Selection.PickObjects(
                Selection.ObjectType.Element, 
                "Select pipes from the networks you want to analyze."
            )
            self.selected_elements = [doc.GetElement(r) for r in refs]
            forms.alert("Selected {} elements.".format(len(self.selected_elements)))
        except Exception as e:
            # Operation cancelled or error
            pass

    def analyze_networks(self, sender, args):
        if not self.selected_elements:
            forms.alert("Please select pipes first.")
            return

        # Algorithm:
        # 1. Iterate through selected elements.
        # 2. If element is not already visited, start traversal to find its 'Island'.
        # 3. Store Island. 
        
        global_visited = set()
        self.network_islands = []
        
        island_id = 1
        
        for el in self.selected_elements:
            if el.Id not in global_visited:
                # Start Traversal
                connected_els = get_connected_elements(el)
                
                # Mark all as visited
                for c_el in connected_els:
                    global_visited.add(c_el.Id)
                
                # Create Island Data
                island = NetworkIsland(island_id, connected_els)
                self.network_islands.append(island)
                island_id += 1
        
        # Populate Grid
        self.dgNetworks.ItemsSource = self.network_islands
        forms.alert("Found {} disconnected networks.".format(len(self.network_islands)))

    def visualize_networks(self, sender, args):
        """Applies color overrides to distinct islands."""
        if not self.network_islands:
            return

        # Pre-defined high contrast colors
        colors = [
            Color(0, 255, 0),    # Neon Green
            Color(0, 255, 255),  # Cyan
            Color(255, 0, 255),  # Magenta
            Color(255, 255, 0),  # Yellow
            Color(255, 0, 0),    # Red
            Color(0, 0, 255)     # Blue
        ]
        
        t = Transaction(doc, "Visualize Networks")
        t.Start()
        
        try:
            # Reset active view overrides first (optional, maybe too aggressive?)
            # Let's just override over existing.
            
            for idx, island in enumerate(self.network_islands):
                color = colors[idx % len(colors)]
                
                ogs = OverrideGraphicSettings()
                ogs.SetProjectionLineColor(color)
                # Make lines thicker for visibility
                ogs.SetProjectionLineWeight(6) 
                
                # Apply solid fill pattern if possible (for 3D/Fine)
                # Getting solid fill pattern id
                fill_patterns = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.FillPatternElement).ToElements()
                solid_fill = next((fp for fp in fill_patterns if fp.GetFillPattern().IsSolidFill), None)
                if solid_fill:
                    ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                    ogs.SetSurfaceForegroundPatternColor(color)

                for el in island.Elements:
                    doc.ActiveView.SetElementOverrides(el.Id, ogs)
            
            t.Commit()
            uidoc.RefreshActiveView()
            
        except Exception as e:
            t.RollBack()
            forms.alert("Error visualizing: " + str(e))

    def merge_networks(self, sender, args):
        """
        The 'Jiggle' Fix.
        Attempts to force-connect disconnected systems by slightly moving connector points.
        """
        # We need at least 2 islands to merge anything, 
        # OR we check within a single selected set for disconnects.
        # But 'islands' definition implies they are disconnected.
        
        t = Transaction(doc, "Merge Networks (The Jiggle)")
        t.Start() 
        
        merge_count = 0
        
        try:
            # Strategy:
            # Look for open connectors in all identified islands.
            # Compare distance between open connectors of DIFFERENT islands.
            # If distance < Tolerance (e.g., 0.1 ft), attempt the fix.
            
            all_open_connectors = [] # Tuple (Connector, IslandID)
            
            for island in self.network_islands:
                for el in island.Elements:
                    # Get Connectors
                    conns = None
                    if hasattr(el, "ConnectorManager"):
                        conns = el.ConnectorManager.Connectors
                    elif hasattr(el, "MEPModel") and el.MEPModel:
                        conns = el.MEPModel.ConnectorManager.Connectors
                    
                    if conns:
                        for c in conns:
                            # Check if connector is NOT connected
                            if not c.IsConnected:
                                all_open_connectors.append((c, island.Id))

            # Find pairs
            tolerance = 0.5 # ft. Generous tolerance to find nearby pipes.
            
            processed_connectors = set()

            for i in range(len(all_open_connectors)):
                c1, id1 = all_open_connectors[i]
                if c1 in processed_connectors: continue
                
                for j in range(i+1, len(all_open_connectors)):
                    c2, id2 = all_open_connectors[j]
                    if c2 in processed_connectors: continue
                    
                    # If same island, ignore (unless we are fixing internal breaks)
                    # if id1 == id2: continue 

                    dist = c1.Origin.DistanceTo(c2.Origin)
                    
                    if dist < tolerance:
                        # FOUND A MATCH!
                        # The Fix:
                        # 1. Align: Move one element so connectors match exactly?
                        # 2. Connect: Explicitly call ConnectTo
                        
                        try:
                            # Try explicit connection first (Cleanest)
                            if c1.Origin.IsAlmostEqualTo(c2.Origin):
                                c1.ConnectTo(c2)
                                merge_count += 1
                                processed_connectors.add(c1)
                                processed_connectors.add(c2)
                                break
                            
                            # If not exactly coincident but close, we must Move.
                            # Move the owner of c2 to c1
                            owner2 = c2.Owner
                            translation_vector = c1.Origin - c2.Origin
                            
                            # Only move if it's a pipe or fitting (not heavy equipment usually)
                            # Avoiding moving pinned elements
                            if not owner2.Pinned:
                                ElementTransformUtils.MoveElement(doc, owner2.Id, translation_vector)
                                doc.Regenerate() # Important!
                                c1.ConnectTo(c2)
                                merge_count += 1
                                processed_connectors.add(c1)
                                processed_connectors.add(c2)
                                break
                                
                        except Exception as inner_e:
                            # print("Failed to merge pair: " + str(inner_e))
                            pass
            
            # The 'Jiggle' Fallback
            # If explicit connection didn't work, sometimes just moving an element back and forth triggers a heal.
            # Applying to all selected elements just in case.
            # Only do this if we haven't successfully merged everything.
            if merge_count == 0:
                # Shake it off
                ids_to_shake = [el.Id for island in self.network_islands for el in island.Elements if not el.Pinned]
                if ids_to_shake:
                    # Move up 0.1mm
                    shake_vec = XYZ(0, 0, 0.003) # ~1mm
                    ElementTransformUtils.MoveElements(doc, List[ElementId](ids_to_shake), shake_vec)
                    doc.Regenerate()
                    # Move back
                    ElementTransformUtils.MoveElements(doc, List[ElementId](ids_to_shake), -shake_vec)
            
            t.Commit()
            if merge_count > 0:
                forms.alert("Successfully merged {} connection points.".format(merge_count))
            else:
                forms.alert("Jiggle complete. Check if systems propagated.")

        except Exception as e:
            t.RollBack()
            forms.alert("Error during merge: " + str(e))

# --- Main Execution ---

def run():
    # pyRevit Modeless Window Execution
    window = PipeNetworkWindow()
    window.Show() # Modeless
    
if __name__ == '__main__':
    run()