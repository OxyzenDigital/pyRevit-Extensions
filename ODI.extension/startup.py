# -*- coding: utf-8 -*-
from pyrevit import forms, script

# debug print removed

try:
    from managesheets.panel import ManageSheetsPanel

    logger = script.get_logger()

    # Panel Registration
    panel_id = ManageSheetsPanel.panel_id

    if not forms.is_registered_dockable_panel(ManageSheetsPanel):
        try:
            forms.register_dockable_panel(ManageSheetsPanel)
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            from System.Windows import MessageBox
            MessageBox.Show("Failed to register Manage Sheets dockable panel in startup.py:\n\n" + error_msg, "Crash Dump")
except Exception as main_e:
    import traceback
    error_msg = traceback.format_exc()
    from System.Windows import MessageBox
    MessageBox.Show("Fatal Error in Manage Sheets startup.py:\n\n" + error_msg, "Crash Dump")
