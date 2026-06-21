# -*- coding=utf-8 -*-
"""Manage Sheets.
An architectural tool utilizing a decoupled 'Data Revisor' schema builder
and layout simulator to manage sheets, views, and scopes in Revit models.
"""

from pyrevit import forms, script
import System

# Configure logger
logger = script.get_logger()

def main():
    # Present options to user
    actions = [
        "Export Model State & Open Browser UI",
        "Import Schema Intent & Sync Model"
    ]
    
    choice = forms.CommandSwitchWindow.show(
        actions,
        message="Select Manage Sheets Action:"
    )
    
    if not choice:
        script.exit()
        
    if choice == "Export Model State & Open Browser UI":
        try:
            import export_state
            export_state.run()
        except System.Exception as e:
            import traceback
            tb = traceback.format_exc()
            forms.alert("Error running exporter (CLR):\n{}".format(e.Message), title="Exporter Error")
            logger.error("Exporter failed (CLR):\n{}".format(tb))
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            forms.alert("Error running exporter:\n{}".format(e), title="Exporter Error")
            logger.error("Exporter failed:\n{}".format(tb))
            
    elif choice == "Import Schema Intent & Sync Model":
        try:
            import import_schema
            import_schema.run()
        except System.Exception as e:
            import traceback
            tb = traceback.format_exc()
            forms.alert("Error running importer (CLR):\n{}".format(e.Message), title="Importer Error")
            logger.error("Importer failed (CLR):\n{}".format(tb))
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            forms.alert("Error running importer:\n{}".format(e), title="Importer Error")
            logger.error("Importer failed:\n{}".format(tb))

if __name__ == "__main__":
    main()
