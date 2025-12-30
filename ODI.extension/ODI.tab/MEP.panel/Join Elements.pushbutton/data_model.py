# -*- coding: utf-8 -*-
from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId, XYZ

class JoinSolution(object):
    """
    Represents a single calculated solution for joining two pipes.
    """
    def __init__(self, name, description):
        self.name = name
        self.description = description
        
        # Metrics requested by user
        self.curve_count = 0
        self.fitting_count = 0
        self.slopes = [] # List of slope values (float)
        self.angles = [] # List of angle values (degrees)
        
        # The abstract geometry/instructions
        # List of tuples/objects defining the path: ('pipe', p1, p2), ('fitting', pt, type)
        self.steps = [] 
        self.is_valid = False
        self.error_message = ""
        self.context = {} # Dictionary to store points, IDs for RevitService

class AppState(object):
    """
    The Single Source of Truth for the application.
    Serialized/Passed between Window and Script.
    """
    def __init__(self):
        # Window Position
        self.win_top = 200.0
        self.win_left = 200.0
        
        # 1. Selections
        self.source_id = None # ElementId
        self.target_id = None # ElementId
        self.source_desc = "[None]"
        self.target_desc = "[None]"
        
        # 2. Logic / Solutions
        self.solutions = [] # List[JoinSolution]
        self.selected_solution_index = -1
        
        # 3. Settings (User Inputs)
        self.allow_rolling = True
        self.allow_vertical = True
        self.use_45_angles = True
        self.use_90_angles = True
        
        # 4. Control Flow
        self.next_action = None # 'select', 'preview_next', 'preview_prev', 'commit', 'settings'
        self.status_message = "Ready to select pipes."
        self.is_ready = False
        self.preview_ids = [] # IDs of temporary DirectShapes

    @property
    def current_solution(self):
        if 0 <= self.selected_solution_index < len(self.solutions):
            return self.solutions[self.selected_solution_index]
        return None
