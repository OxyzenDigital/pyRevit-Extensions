#pylint: disable=import-error,invalid-name,broad-except,missing-docstring
"""
This script creates a new legend view in Revit and populates it with text from a 
user-selected text file.
"""

__title__ = "Bring Text \nto Legend View"
__author__ = "ODI"

import os
import System

# --- CRITICAL FIX: Load the Windows Forms Assembly ---
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog
# -----------------------------------------------------

from System.IO import StreamReader

# Import Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    TextNoteType,
    BuiltInParameter,
    Transaction,
    ViewDuplicateOption,
    ViewType,
    XYZ,
    TextNote,
    TextNoteOptions,
    HorizontalTextAlignment,
    View
)
from pyrevit import revit, script, forms

doc = revit.doc
uidoc = revit.uidoc

def get_or_create_text_type(doc, type_name, font="Arial", point_size=10.0):
    """Finds or creates a TextNoteType safely."""
    collector = FilteredElementCollector(doc).OfClass(TextNoteType)
    
    # 1. Search for existing type
    for text_type in collector:
        name_param = text_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if name_param and name_param.AsString() == type_name:
            return text_type.Id

    # 2. Duplicate to create new if not found
    default_text_type = collector.FirstElement()
    if not default_text_type:
        forms.alert("No TextNoteTypes found in project to duplicate.")
        return None

    try:
        with Transaction(doc, "Create Text Type") as t:
            t.Start()
            new_text_type = default_text_type.Duplicate(type_name)
            
            # Convert points to feet (1 pt = 1/72 inch)
            text_size_feet = point_size / (72.0 * 12.0)

            # Set Text Size (Language Independent)
            p_size = new_text_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
            if p_size and not p_size.IsReadOnly:
                p_size.Set(text_size_feet)

            # Set Font (Language Independent)
            p_font = new_text_type.get_Parameter(BuiltInParameter.TEXT_FONT)
            if p_font and not p_font.IsReadOnly:
                p_font.Set(font)
            
            t.Commit()
            return new_text_type.Id
    except Exception as e:
        forms.alert("Failed to create text type: {}".format(e))
        return None

def split_text_into_columns(content, chars_per_column=3500):
    """Splits text into chunks without breaking words."""
    text_length = len(content)
    columns = []
    start_index = 0
    
    while start_index < text_length:
        end_index = min(start_index + chars_per_column, text_length)
        
        if end_index < text_length:
            # Find natural break points (paragraph, sentence, space)
            break_point = content.rfind('\n\n', start_index, end_index)
            if break_point == -1:
                break_point = content.rfind('. ', start_index, end_index)
            if break_point == -1:
                break_point = content.rfind(' ', start_index, end_index)
            
            if break_point > start_index:
                end_index = break_point + 1
        
        columns.append(content[start_index:end_index].strip())
        start_index = end_index
    
    return columns

def create_or_get_legend_view(doc, view_name):
    """Finds an existing legend or creates a new one by duplicating a template."""
    
    # 1. Check if the legend already exists
    all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
    
    for view in all_views:
        if view.ViewType == ViewType.Legend and view.Name == view_name:
            # It exists, so clear old text out of it
            try:
                with Transaction(doc, "Clean View") as t:
                    t.Start()
                    notes = FilteredElementCollector(doc, view.Id).OfClass(TextNote).ToElementIds()
                    if notes:
                        doc.Delete(notes)
                    t.Commit()
            except Exception:
                pass 
            return view

    # 2. Find a template to duplicate (must be a Legend)
    base_legend = None
    for view in all_views:
        if view.ViewType == ViewType.Legend and not view.IsTemplate:
            base_legend = view
            break
            
    if not base_legend:
        forms.alert("No Legend Views found in project. Please create at least one Legend manually to serve as a template.")
        return None

    # 3. Create the new legend
    try:
        with Transaction(doc, "Create Legend") as t:
            t.Start()
            new_view_id = base_legend.Duplicate(ViewDuplicateOption.Duplicate)
            new_view = doc.GetElement(new_view_id)
            new_view.Name = view_name
            t.Commit()
            return new_view
    except Exception as e:
        forms.alert("Error creating legend view: {}".format(e))
        return None

def main():
    # 1. Select Text File
    dialog = OpenFileDialog()
    dialog.Filter = "Text Files (*.txt)|*.txt"
    dialog.Title = "Select Text File to Import"
    
    if dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK:
        return
    
    file_path = dialog.FileName
    legend_name = os.path.splitext(os.path.basename(file_path))[0]

    # 2. Read File Content
    try:
        with StreamReader(file_path) as reader:
            content = reader.ReadToEnd()
    except Exception as e:
        forms.alert("Could not read file: {}".format(e))
        return

    if not content:
        forms.alert("File is empty.")
        return

    # 3. Get/Create View
    legend_view = create_or_get_legend_view(doc, legend_name)
    if not legend_view: 
        return
    
    # 4. Get/Create Text Type
    text_type_id = get_or_create_text_type(doc, "10pt_Imported", "Arial", 10.0)
    if not text_type_id: 
        return

    # 5. Split Text
    cols = split_text_into_columns(content, 3000)
    
    # 6. Place Text Columns
    col_w_feet = 5.0 / 12.0  # 5 inches wide
    gap_feet = 0.5 / 12.0    # 0.5 inch gap
    curr_x = 0.5 / 12.0
    curr_y = 2.0             # Start 2 feet up
    
    with Transaction(doc, "Place Text Columns") as t:
        t.Start()
        for txt in cols:
            opts = TextNoteOptions(text_type_id)
            opts.HorizontalAlignment = HorizontalTextAlignment.Left
            
            TextNote.Create(doc, legend_view.Id, XYZ(curr_x, curr_y, 0), col_w_feet, txt, opts)
            
            curr_x += (col_w_feet + gap_feet)
        t.Commit()

    # 7. Open the view
    try:
        uidoc.ActiveView = legend_view
    except:
        pass

if __name__ == "__main__":
    main()