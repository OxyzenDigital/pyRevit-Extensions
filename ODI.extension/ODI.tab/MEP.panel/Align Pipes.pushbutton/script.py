# -*- coding: utf-8 -*-
"""
Align Pipes Tool

Allows the user to align one pipe to another in the XY plane.
Supports:
1. Parallel Horizontal Pipes (Collinear Alignment)
2. Vertical Riser to Vertical Riser (Concentric Alignment)
3. Vertical Riser to Horizontal Pipe (Intersecting Alignment)
   - Moves the horizontal pipe so its axis points directly at the vertical riser.
"""

from Autodesk.Revit.DB import (
    Transaction,
    BuiltInCategory,
    XYZ,
    Line,
    ElementTransformUtils,
    LocationCurve
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, script, forms
import math

# --- Configuration ---
TOLERANCE = 0.01 # Tolerance for parallel check (radians)

# --- Logging Setup ---
log_buffer = []

def log_section(title):
    log_buffer.append("\n### {}".format(title))

def log_item(key, value):
    log_buffer.append("- **{}:** {}".format(key, value))

def log_point(name, point):
    log_buffer.append("- **{}:** ({:.4f}, {:.4f}, {:.4f})".format(name, point.X, point.Y, point.Z))

def log_vector(name, vector):
    if vector:
        log_buffer.append("- **{}:** <{:.4f}, {:.4f}, {:.4f}>".format(name, vector.X, vector.Y, vector.Z))
    else:
        log_buffer.append("- **{}:** None".format(name))

def show_log():
    output = script.get_output()
    output.close_others()
    for msg in log_buffer:
        output.print_md(msg)

# --- Helpers ---

def get_id_value(element_id):
    """Safely gets the integer value of an ElementId."""
    if hasattr(element_id, "Value"): # Revit 2024+
        return element_id.Value
    elif hasattr(element_id, "IntegerValue"): # Pre-2024
        return element_id.IntegerValue
    else:
        return int(element_id)

def get_xy_vector(v):
    """Projects a vector to the XY plane (Z=0) and normalizes it."""
    v_xy = XYZ(v.X, v.Y, 0)
    if v_xy.IsZeroLength():
        return None
    return v_xy.Normalize()

def is_vertical(v):
    """Checks if a vector is vertical (Z-aligned)."""
    return abs(v.Z) > 0.99 

def are_parallel(v1, v2, tolerance=TOLERANCE):
    """Checks if two vectors are parallel within a tolerance."""
    if not v1 or not v2:
        return False
    return v1.CrossProduct(v2).IsZeroLength() or v1.AngleTo(v2) < tolerance or v1.AngleTo(v2.Negate()) < tolerance

def project_point_to_line_infinite_xy(point, line_origin, line_dir_xy):
    """
    Projects a point onto an infinite line defined by origin and direction in XY plane.
    line_dir_xy must be normalized and in XY plane.
    """
    v_to_point = point - line_origin
    v_to_point_xy = XYZ(v_to_point.X, v_to_point.Y, 0)
    
    dot_prod = v_to_point_xy.DotProduct(line_dir_xy)
    projected_vector = line_dir_xy.Multiply(dot_prod)
    
    closest_point_xy = XYZ(line_origin.X, line_origin.Y, 0) + projected_vector
    
    # Return projected point keeping original Z
    return XYZ(closest_point_xy.X, closest_point_xy.Y, point.Z)

def is_pipe(element):
    if not element or not element.Category:
        return False
    try:
        cat_id_val = get_id_value(element.Category.Id)
    except Exception:
        return False
    return cat_id_val == int(BuiltInCategory.OST_PipeCurves) or cat_id_val == int(BuiltInCategory.OST_PlaceHolderPipes)

def pick_pipe_safely(uidoc, doc, prompt):
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
        element = doc.GetElement(ref)
        if is_pipe(element):
            return element
        else:
            forms.alert("Selected element is not a Pipe.", title="Invalid Selection")
            return None
    except OperationCanceledException:
        return None

def align_pipe_geometry(ref_pipe, target_pipe):
    """
    Calculates the move vector for target_pipe to align with ref_pipe in XY.
    """
    log_section("Geometry Analysis")
    
    # 1. Get Location Curves
    loc_ref = ref_pipe.Location
    loc_target = target_pipe.Location
    
    if not isinstance(loc_ref, LocationCurve) or not isinstance(loc_target, LocationCurve):
        return None, "One of the elements does not have a location curve."

    curve_ref = loc_ref.Curve
    curve_target = loc_target.Curve
    
    if not isinstance(curve_ref, Line) or not isinstance(curve_target, Line):
        return None, "Only straight pipes are supported."

    # 2. Get Vectors and Points
    p1 = curve_ref.Origin
    v1 = curve_ref.Direction.Normalize()
    
    p2 = curve_target.Origin
    v2 = curve_target.Direction.Normalize()
    
    log_point("Ref Pipe Origin", p1)
    log_vector("Ref Pipe Direction", v1)
    log_point("Target Pipe Origin", p2)
    log_vector("Target Pipe Direction", v2)

    ref_is_vert = is_vertical(v1)
    target_is_vert = is_vertical(v2)
    
    log_item("Ref Pipe Vertical?", ref_is_vert)
    log_item("Target Pipe Vertical?", target_is_vert)

    # CASE A: Both Vertical (Concentric)
    if ref_is_vert and target_is_vert:
        log_item("Alignment Mode", "Vertical -> Vertical (Concentric)")
        # Move Target (P2) to match Ref (P1) in X,Y
        target_pos_xy = XYZ(p1.X, p1.Y, p2.Z)
        move_vector = target_pos_xy - p2
        return move_vector, None

    # CASE B: Ref Vertical, Target Horizontal/Sloped
    if ref_is_vert and not target_is_vert:
        log_item("Alignment Mode", "Vertical Ref -> Horizontal Target (Intersection)")
        # Ref is a point (P1.X, P1.Y).
        # Target is a line (P2, V2).
        # We need to move Target (perpendicularly to V2) so that its line passes through Ref.
        
        # 1. Project Ref Point onto Target Line (Current Position)
        v2_xy = get_xy_vector(v2)
        if not v2_xy: return None, "Target pipe has no XY length."
        
        # Where is Ref Point currently relative to Target Line?
        ref_projected_on_target = project_point_to_line_infinite_xy(p1, p2, v2_xy)
        
        # Vector from Projected Point (on Target Line) to Ref Point (Desired Line)
        # We want the line to move TO the Ref point.
        move_vector = XYZ(p1.X, p1.Y, 0) - XYZ(ref_projected_on_target.X, ref_projected_on_target.Y, 0)
        
        log_point("Ref Point (XY)", p1)
        log_point("Projected Ref on Target Line", ref_projected_on_target)
        log_vector("Calculated Move Vector", move_vector)
        return move_vector, None

    # CASE C: Ref Horizontal/Sloped, Target Vertical
    if not ref_is_vert and target_is_vert:
        log_item("Alignment Mode", "Horizontal Ref -> Vertical Target (Intersection)")
        # Ref is a Line (P1, V1).
        # Target is a Point (P2.X, P2.Y).
        # We need to move Target (the point) to lie on Ref Line.
        
        v1_xy = get_xy_vector(v1)
        if not v1_xy: return None, "Ref pipe has no XY length."
        
        # Project Target Point onto Ref Line
        target_projected_on_ref = project_point_to_line_infinite_xy(p2, p1, v1_xy)
        
        # Move Target TO the projected point
        move_vector = XYZ(target_projected_on_ref.X, target_projected_on_ref.Y, 0) - XYZ(p2.X, p2.Y, 0)
        
        log_vector("Calculated Move Vector", move_vector)
        return move_vector, None

    # CASE D: Both Horizontal/Sloped
    log_item("Alignment Mode", "Horizontal -> Horizontal")
    v1_xy = get_xy_vector(v1)
    v2_xy = get_xy_vector(v2)

    if are_parallel(v1_xy, v2_xy):
        log_item("Sub-Mode", "Parallel (Collinear)")
        # Existing logic: Move P2 to line of P1
        target_pos = project_point_to_line_infinite_xy(p2, p1, v1_xy)
        move_vector = target_pos - p2
        move_vector = XYZ(move_vector.X, move_vector.Y, 0)
        return move_vector, None
    else:
        log_item("Sub-Mode", "Non-Parallel (Skew)")
        # In this case, non-parallel lines in 2D already intersect. 
        # "Aligning" them implies we want them to intersect? They do.
        # Unless user wants to move Target so the Intersection Point aligns with something else?
        # Given "vector... intersects", we assume infinite lines. They intersect.
        
        return XYZ.Zero, "Non-parallel horizontal pipes already intersect in Plan View. No alignment needed."

# --- Main Execution ---

def main():
    uidoc = revit.uidoc
    doc = revit.doc
    
    log_section("Initialization")
    log_item("Tool", "Align Pipes")
    log_item("Active View", doc.ActiveView.Name)

    # 1. Select Reference Pipe
    ref_pipe = pick_pipe_safely(uidoc, doc, "Select REFERENCE Pipe (Stationary)")
    if not ref_pipe:
        log_item("Status", "Cancelled or Invalid Ref selection")
        return
    log_item("Reference Pipe Selected", get_id_value(ref_pipe.Id))

    # 2. Select Target Pipe
    target_pipe = pick_pipe_safely(uidoc, doc, "Select TARGET Pipe (To Move)")
    if not target_pipe:
        log_item("Status", "Cancelled or Invalid Target selection")
        return
    log_item("Target Pipe Selected", get_id_value(target_pipe.Id))

    if ref_pipe.Id == target_pipe.Id:
        forms.alert("You selected the same pipe twice.", title="Error")
        show_log()
        return

    # 3. Calculate Alignment
    move_vector, error = align_pipe_geometry(ref_pipe, target_pipe)
    
    if error:
        forms.alert(error, title="Alignment Failed")
        log_item("Error", error)
        show_log()
        return
        
    if move_vector.IsZeroLength():
        forms.alert("Pipes are already aligned.", title="Info")
        log_item("Result", "Already Aligned")
        show_log()
        return

    # 4. Execute Move
    try:
        log_section("Execution")
        with Transaction(doc, "Align Pipes") as t:
            t.Start()
            ElementTransformUtils.MoveElement(doc, target_pipe.Id, move_vector)
            t.Commit()
        log_item("Transaction", "Success")
        log_item("Result", "Pipe Aligned")
        
    except Exception as e:
        forms.alert("An error occurred: {}".format(str(e)), title="Error")
        log_item("Transaction Error", str(e))
    
    show_log()

if __name__ == '__main__':
    main()
