"""Export Text Notes to a Keynote file."""
# -*- coding: utf-8 -*-
import os
import re
import io
import datetime
from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()

def sanitize_text_for_keynote(text):
    """
    Cleans text to be safe for Keynote file format (Tab-separated).
    Removes tabs and newlines.
    """
    if not text:
        return ""
    # Replace tabs and newlines with spaces to maintain one line per entry
    clean = text.replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')
    # Remove multiple spaces
    clean = re.sub(r'\s+', ' ', clean)
    return clean.strip()

def sanitize_key(text):
    """
    Creates a safe key from text (alphanumeric only + hyphens/underscores).
    """
    if not text:
        return "UNKNOWN"
    # Replace invalid chars with underscore
    clean = re.sub(r'[^a-zA-Z0-9\-_]', '_', text)
    return clean

def get_text_notes_in_view(view_id):
    """Collects all text notes in a specific view."""
    return (DB.FilteredElementCollector(doc, view_id)
            .OfClass(DB.TextNote)
            .ToElements())

def main():
    # 1. Select Sheets
    sheets = forms.select_sheets(include_placeholder=False)
    if not sheets:
        return

    # Prepare Data Structure
    # List of tuples: (Key, Text, ParentKey)
    keynote_data = []
    
    # Track keys to ensure uniqueness (though logic should enforce it)
    # Using a set is good practice if we were merging, but here we build hierarchically.
    
    with forms.ProgressBar(title="Gathering Text Notes...", cancellable=True) as pb:
        total = len(sheets)
        for i, sheet in enumerate(sheets):
            if pb.cancelled:
                break
            pb.update_progress(i, total)
            
            sheet_num = sheet.SheetNumber
            sheet_name = sheet.Name
            
            # Level 1: Sheet
            # Key: SheetNumber (e.g. A101)
            sheet_key = sanitize_key(sheet_num)
            keynote_data.append((sheet_key, sheet_name, ""))
            
            # --- A. Notes directly on Sheet ---
            sheet_notes = get_text_notes_in_view(sheet.Id)
            if sheet_notes:
                # Create a "Sheet Notes" section to keep structure consistent
                # Key: A101.00
                gen_notes_key = "{}.00".format(sheet_key)
                keynote_data.append((gen_notes_key, "Sheet Notes", sheet_key))
                
                for i, note in enumerate(sheet_notes, 1):
                    note_text = sanitize_text_for_keynote(note.Text)
                    if not note_text: continue
                    
                    # Key: A101.00.01
                    note_key = "{}.{:02d}".format(gen_notes_key, i)
                    keynote_data.append((note_key, note_text, gen_notes_key))

            # --- B. Views placed on Sheet ---
            # Get all viewports
            viewports_ids = sheet.GetAllViewports()
            
            view_counter = 1
            for vp_id in viewports_ids:
                vp = doc.GetElement(vp_id)
                view_id = vp.ViewId
                view = doc.GetElement(view_id)
                
                # Filter: Only Drafting Views and Legend Views
                if view.ViewType in [DB.ViewType.DraftingView, DB.ViewType.Legend]:
                    view_name = view.Name
                    
                    # Level 2: View
                    # Key: A101.01, A101.02, etc.
                    view_key = "{}.{:02d}".format(sheet_key, view_counter)
                    view_counter += 1
                    
                    keynote_data.append((view_key, view_name, sheet_key))
                    
                    # Get Notes in View
                    view_notes = get_text_notes_in_view(view.Id)
                    
                    for k, note in enumerate(view_notes, 1):
                        note_text = sanitize_text_for_keynote(note.Text)
                        if not note_text: continue
                        
                        # Level 3: Note
                        # Key: A101.01.01
                        note_key = "{}.{:02d}".format(view_key, k)
                        keynote_data.append((note_key, note_text, view_key))

    if not keynote_data:
        forms.alert("No text notes found in selected sheets/views.", title="Export Result")
        return

    # 2. Export
    # Default filename: ProjectTitle_YYYY-MM-DD_HH-MM.txt
    project_title = doc.Title
    # Remove extension if present (common in detached files or local copies)
    if project_title.lower().endswith(".rvt"):
        project_title = project_title[:-4]
        
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    default_name = "{}_{}.txt".format(project_title, timestamp)
    
    default_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    
    dest_file = forms.save_file(file_ext='txt', 
                                default_name=default_name, 
                                init_dir=default_dir,
                                title="Save Keynote File")
                                
    if dest_file:
        try:
            with io.open(dest_file, 'w', encoding='utf-8') as f:
                # Format: Key \t Text \t ParentKey
                for item in keynote_data:
                    line = "{}\t{}\t{}\n".format(item[0], item[1], item[2])
                    f.write(line)
            
            # Show Result
            forms.alert("Export Successful!\nFile saved to: {}".format(dest_file), 
                              title="Export Complete", 
                              warn_icon=False)
            
        except Exception as e:
            forms.alert("Error saving file: \n{}".format(str(e)), title="Export Error")

if __name__ == '__main__':
    main()
