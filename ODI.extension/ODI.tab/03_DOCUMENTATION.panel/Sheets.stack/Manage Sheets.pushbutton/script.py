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
        window = ManageSheetsPanel()
        window.show_dialog() # Opens as a true modal window, blocking Revit until closed
    except Exception as e:
        error_msg = traceback.format_exc()
        MessageBox.Show("Error initializing Manage Sheets:\n\n" + error_msg, "Crash Dump")
