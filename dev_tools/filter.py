import os
import subprocess
import glob

# Remove pycache files
pyc_files = glob.glob("**/*.pyc", recursive=True)
if pyc_files:
    subprocess.run(["git", "rm", "--cached", "--ignore-unmatch", "-f"] + pyc_files, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

files_to_remove = [
    "ODI.extension/ODI.tab/03_DOCUMENTATION.panel/Sheets.stack/Manage Sheets.pushbutton/.backup",
    "ODI.extension/lib/managesheets/temp_target.txt",
    "ODI.extension/ODI.tab/03_DOCUMENTATION.panel/Manage.pulldown/Manage Graphic Scales.pushbutton/config.json",
    "ODI.extension/ODI.tab/03_DOCUMENTATION.panel/Manage.pulldown/Manage ViewTitles.pushbutton/manage_views_config.json",
    "ODI.extension/ODI.tab/03_DOCUMENTATION.panel/Import.pulldown/Excel Named Range to Annotation.pushbutton/excel_table_map.json"
]

subprocess.run(["git", "rm", "--cached", "--ignore-unmatch", "-r", "-f"] + files_to_remove, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
