# -*- coding: utf-8 -*-
__title__ = "Manage Sheets"
__version__ = "4.2"
__doc__ = """A WPF Window to align active Revit sheets against dynamically generated AIA UDS schemas using a Card-Style Tree Grid UI."""

import traceback
from pyrevit import forms
from System.Windows import MessageBox

if __name__ == '__main__':
    try:
        from managesheets.panel import ManageSheetsPanel
        from managesheets import project_settings
        from managesheets import classification
        
        doc = __revit__.ActiveUIDocument.Document
        
        # Check if custom parameters exist
        from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction
        sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
        param_exists = False
        if sheets:
            for test_sheet in sheets:
                if test_sheet.LookupParameter("Sheet Series"):
                    param_exists = True
                break
        else:
            # No sheets in project, assume true or safe to skip
            param_exists = True
            
        if not param_exists:
            res = forms.alert("The 'Sheet Series' parameter is missing from your project.\n\nThis parameter is required to persist custom Series assignments for your sheets. Would you like to automatically create it now?", options=["Yes", "No"])
            import sys
            if res == "Yes":
                from managesheets.panel import ensure_sheet_parameter
                with Transaction(doc, "Add Manage Sheets Parameters") as t:
                    t.Start()
                    ensure_sheet_parameter(doc, "Sheet Series")
                    ensure_sheet_parameter(doc, "Discipline")
                    ensure_sheet_parameter(doc, "Content Group")
                    t.Commit()
                forms.alert("Parameters created successfully!\n\nPlease restart the Manage Sheets tool to continue.")
                sys.exit(0)
            else:
                sys.exit(0)
        
        window = ManageSheetsPanel()
        window.show_dialog() # Opens as a true modal window, blocking Revit until closed
        
        # After window closes, execute any pending saves synchronously
        if getattr(window, 'classification_needs_saving', False):
            project_settings.save_classification_dict_sync(doc, classification.CLASSIFICATION_DICT)
        
        if getattr(window, 'setup_needs_saving', False):
            project_settings.save_project_setup_sync(doc, window.pending_setup_dict)

    except Exception as e:
        error_msg = traceback.format_exc()
        MessageBox.Show("Error initializing Manage Sheets:\n\n" + error_msg, "Crash Dump")
