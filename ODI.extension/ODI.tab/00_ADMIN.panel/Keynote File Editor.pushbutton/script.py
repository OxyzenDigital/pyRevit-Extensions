# -*- coding: utf-8 -*-

import os
from pyrevit import forms

# Construct path to HTML file relative to this script
html_file = os.path.join(os.path.dirname(__file__), 'ODI Keynotes Editor.html')

if os.path.exists(html_file):
    os.startfile(html_file)
else:
    forms.alert('HTML file not found: {}'.format(html_file))
