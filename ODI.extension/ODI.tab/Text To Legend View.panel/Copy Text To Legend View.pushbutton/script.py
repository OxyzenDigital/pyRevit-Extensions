#pylint: disable=import-error,invalid-name,broad-except
"""Adds formatted text in columns to the active view."""

__title__ = "Add Text\nColumns"
__doc__ = "Creates formatted text columns in the active view from a text file"
__author__ = "ODI"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System')
clr.AddReference('System.IO')

from Autodesk.Revit.DB import *
import System
from System.Windows.Forms import OpenFileDialog
from System.IO import StreamReader, File
from pyrevit import revit, DB, script, forms
import os
from Autodesk.Revit.Exceptions import ArgumentException  # Add this import

# Get output window
output = script.get_output()

# Access the current Revit document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

if not doc:
    forms.alert("No active Revit document found. Please open a project in Revit and try again.")
    script.exit()

def get_or_create_text_type(doc, type_name, font="Arial", text_size=10.0):
    """
    Finds or creates a TextNoteType with the specified format.
    """
    text_type_collector = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
    
    # Find existing text type
    for tt in text_type_collector:
        if tt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == type_name:
            return tt.Id  # Just return the Id directly
    
    text_types = list(text_type_collector)
    if not text_types:
        forms.alert("No existing TextNoteType found to duplicate.")
        script.exit()
    
    default_text_type = text_types[0]

    try:
        with revit.Transaction("Create TextNoteType"):
            new_text_type = default_text_type.Duplicate(type_name)  # This returns the new TextNoteType
            
            params_to_set = {
                "Text Font": font,
                "Text Size": text_size / 72.0,  # Convert points to feet
                "Text Width Scale": 1.0
            }

            for param_name, value in params_to_set.items():
                param = new_text_type.LookupParameter(param_name)
                if param and not param.IsReadOnly:
                    param.Set(value)
                else:
                    output.print_md("**Warning**: Unable to set {} parameter".format(param_name))

            return new_text_type.Id  # Return the Id directly

    except Exception as e:
        forms.alert("Failed to create text type: {}".format(str(e)))
        script.exit()

def calculate_column_capacity(width_inches=5.0, height_inches=24.0, point_size=10.0):

    return 3500  # Known capacity from testing

def split_text_into_columns(content, chars_per_column=3500):
    """
    Split text into columns based on known capacity of 3500 chars per column.
    Let Revit handle the actual line wrapping.
    """
    text_length = len(content)
    output.print_md("## Text Splitting Info:")
    output.print_md("- Total text length: {} characters".format(text_length))
    output.print_md("- Characters per column: 3500 (based on testing)")
    output.print_md("- Estimated columns needed: {:.1f}".format(text_length / 3500.0))
    
    columns = []
    start = 0
    
    while start < text_length:
        # Find break point around 3500 characters
        end = min(start + 3500, text_length)
        
        # If not at the end, find a good break point
        if end < text_length:
            # Try to find paragraph break first, then sentence, then word
            breaks = [
                content.rfind('\n\n', start, end + 1),  # Paragraph
                content.rfind('. ', start, end + 1),    # Sentence
                content.rfind('? ', start, end + 1),    # Question
                content.rfind('! ', start, end + 1),    # Exclamation
                content.rfind(' ', start, end + 1),     # Word
            ]
            break_point = max(breaks)
            if break_point > start:
                end = break_point + 1
        
        column_text = content[start:end].strip()
        columns.append(column_text)
        output.print_md("- Column {}: {} characters".format(len(columns), len(column_text)))
        
        start = end
    
    return columns

def clean_view_content(doc, view):
    """
    Removes all elements from the given view.
    """
    try:
        with revit.Transaction("Clean View Content"):
            collector = FilteredElementCollector(doc, view.Id)
            elements = collector.WhereElementIsNotElementType().ToElements()
            for element in elements:
                if isinstance(element, TextNote) and doc.GetElement(element.Id):  # Check if the element is a TextNote and still exists
                    try:
                        doc.Delete(element.Id)
                    except Exception as e:
                        output.print_md("Warning: Could not delete element - {}".format(str(e)))
    except Exception as e:
        output.print_md("Warning: Could not clean all elements from view - {}".format(str(e)))

def create_legend_view(doc, view_name):
    """
    Creates a new legend view by duplicating an existing one or gets existing view and cleans it.
    """
    # First try to find existing view
    for view in FilteredElementCollector(doc).OfClass(View):
        if view.ViewType == ViewType.Legend and view.Name == view_name:
            clean_view_content(doc, view)
            return view

    # Find a legend view to duplicate from
    base_legend = None
    for view in FilteredElementCollector(doc).OfClass(View):
        if view.ViewType == ViewType.Legend:
            base_legend = view
            break
    
    if not base_legend:
        forms.alert("No Legend view found in project to duplicate from. Please add a legend view to your template.")
        script.exit()

    try:
        with revit.Transaction("Create Legend View"):
            # Duplicate the legend view
            new_legend = base_legend.Duplicate(ViewDuplicateOption.Duplicate)
            
            if not new_legend:
                forms.alert("Failed to duplicate legend view.")
                script.exit()
                
            # Get the view element
            new_view = doc.GetElement(new_legend)
            
            # Set view properties
            new_view.Name = view_name
            new_view.Scale = 48
            
            # Force regeneration
            doc.Regenerate()
            
            output.print_md("Created legend view: **{}**".format(view_name))
            return new_view

    except Exception as e:
        forms.alert("Failed to create legend view: {}".format(str(e)))
        script.exit()

def create_text_with_columns(doc, view_id, content, text_type_id, paper_width, paper_height, 
                           paper_gap, scale_value, start_position):
    """
    Creates multi-column text notes in the legend view.
    """
    try:
        # Split content into columns using known capacity
        text_columns = split_text_into_columns(content)
        
        # Convert dimensions to model space
        model_width = paper_width / 12.0  # Convert 5" to feet (1:1 scale)
        model_gap = paper_gap / 12.0      # Convert 0.5" to feet (1:1 scale)
        
        # Get the valid width limits from the TextNote class
        min_width = TextNote.GetMinimumAllowedWidth(doc, text_type_id)
        max_width = TextNote.GetMaximumAllowedWidth(doc, text_type_id)
        
        # Ensure the width is within valid bounds
        width_to_use = max(min(model_width, max_width), min_width)
        
        # output.print_md("## Width Limits:")
        # output.print_md("- Minimum width: {:.2f}'".format(min_width))
        # output.print_md("- Maximum width: {:.2f}'".format(max_width))
        # output.print_md("- Requested width: {:.2f}'".format(model_width))
        # output.print_md("- Using width: {:.2f}'".format(width_to_use))
        
        created_text_notes = []
        current_position = start_position

        with revit.Transaction("Create Legend Text"):
            for column_text in text_columns:
                options = TextNoteOptions(text_type_id)
                options.HorizontalAlignment = HorizontalTextAlignment.Left
                
                # Create text note with validated width parameter
                text_note = TextNote.Create(
                    doc,
                    view_id,
                    current_position,
                    width_to_use,  # Use validated width
                    column_text,
                    options
                )
                
                # Ensure the width is set correctly after creation
                text_note.Width = width_to_use
                
                doc.Regenerate()
                created_text_notes.append(text_note.Id)
                
                # Move to next column position using original width for spacing
                current_position = XYZ(
                    current_position.X + (width_to_use + model_gap) * scale_value,  # Convert to inches
                    current_position.Y,
                    current_position.Z
                )

        return created_text_notes

    except Exception as e:
        output.print_md("Error details: {}".format(str(e)))
        forms.alert("Error creating text columns: {}".format(str(e)))
        script.exit()

# Main script
try:
    # Paper space dimensions (actual printed size)
    paper_width = 5.0   # inches
    paper_height = 24.0  # inches
    paper_gap = 0.5     # inches
    
    # View scale and text parameters
    scale_value = 48    # 1/4" = 1'-0"
    point_size = 10.0   # 10pt text
    font_name = "Arial" # Font family name
    font_size = ((point_size/72.0)/12.0) * scale_value  # Convert to model space feet
    
    # Start position in model space (convert from paper space)
    start_x = (0.5/12.0) * scale_value  # 0.5" from left
    start_y = (0.5/12.0) * scale_value  # 0.5" from bottom
    start_position = XYZ(start_x, start_y, 0)

    # Select a text file
    dialog = OpenFileDialog()
    dialog.Filter = "Text Files (*.txt)|*.txt"
    dialog.Multiselect = False
    dialog.Title = "Select Text File for Import"
    
    if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK):
        forms.alert("Operation cancelled by user.")
        script.exit()

    # Get file name without extension for legend name
    legend_name = System.IO.Path.GetFileNameWithoutExtension(dialog.FileName)
    
    # Create or get legend view with additional verification
    legend_view = create_legend_view(doc, legend_name)
    if not legend_view:
        forms.alert("Failed to create or get legend view.")
        script.exit()
    
    output.print_md("## Working with Legend View:")
    output.print_md("- Name: {}".format(legend_view.Name))
    output.print_md("- ID: {}".format(legend_view.Id.IntegerValue))
    output.print_md("- Type: {}".format(legend_view.ViewType))
    
    # Force UI update
    uidoc.RefreshActiveView()

    # Read and validate file content
    try:
        reader = StreamReader(dialog.FileName)
        content = reader.ReadToEnd()
        reader.Close()
    except Exception as e:
        forms.alert("Error reading file: {}".format(str(e)))
        script.exit()

    if not content.strip():
        forms.alert("Selected file is empty")
        script.exit()

    # Get or create the TextNoteType
    text_type_id = get_or_create_text_type(
        doc, 
        "Arial10pt", 
        font=font_name, 
        text_size=font_size
    )

    # Create text columns in legend
    created_notes = create_text_with_columns(
        doc, 
        legend_view.Id, 
        content, 
        text_type_id, 
        paper_width, 
        paper_height, 
        paper_gap, 
        scale_value, 
        start_position
    )

    output.print_md("# Success!")
    output.print_md("Created {} text columns in legend view: {}".format(len(created_notes), legend_name))

except Exception as e:
    forms.alert("Error: {}".format(str(e)))
    script.exit()