# -*- coding: utf-8 -*-
__title__ = "Manage Sheets"
__version__ = "4.2"
__doc__ = """A WPF Dockable Pane to align active Revit sheets against dynamically generated AIA UDS schemas using a Card-Style Tree Grid UI."""

from pyrevit import forms

from pyrevit import forms, EXEC_PARAMS

def register_pane():
    from managesheets.panel import ManageSheetsPanel
    if not forms.is_registered_dockable_panel(ManageSheetsPanel):
        forms.register_dockable_panel(ManageSheetsPanel)

def instantiate_pane():
    # Force reload of the panel module for hot-reloading
    import sys
    if 'managesheets.panel' in sys.modules:
        del sys.modules['managesheets.panel']

def show_pane():
    from managesheets.panel import ManageSheetsPanel
    panel = forms.get_dockable_panel(ManageSheetsPanel.panel_id)
    # PyRevit WPFPanel custom setup helps recreate the UI elements
    panel.custom_setup() 
    if panel.IsShown():
        panel.Hide()
    else:
        panel.Show()

if __name__ == '__main__':
    try:
        if getattr(EXEC_PARAMS, 'event', '') == 'startup':
            register_pane()
        else:
            instantiate_pane()
            show_pane()
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        from System.Windows import MessageBox
        MessageBox.Show("Error initializing Manage Sheets:\n\n" + error_msg, "Crash Dump")
