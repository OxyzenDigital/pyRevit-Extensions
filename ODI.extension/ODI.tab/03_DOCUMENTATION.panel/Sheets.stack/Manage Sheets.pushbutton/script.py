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
