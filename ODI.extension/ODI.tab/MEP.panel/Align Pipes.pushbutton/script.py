# -*- coding: utf-8 -*-
"""
Align Pipes Tool
Version: 1.0
Author: ODI (Generated via Gemini CLI)

Description:
    Allows the user to align multiple pipes to a single reference pipe in the XY plane.
    Designed to facilitate the joining of pipes by ensuring their centerlines intersect 
    or are collinear in the Plan View (XY projection).

Usage:
    1. Run the command.
    2. Select the 'Reference Pipe' (the stationary pipe).
    3. Continuously select 'Target Pipes' (the pipes to be moved).
       - Pipes move immediately upon selection.
    4. Press ESC to finish the command and view the log report.

Supported Alignment Modes:
    1. Horizontal/Sloped -> Horizontal/Sloped (Parallel):
       - Moves the Target pipe laterally so it becomes collinear with the Reference pipe in Plan View.
    
    2. Vertical Riser -> Vertical Riser:
       - Moves the Target pipe so it is concentric (same X,Y) with the Reference pipe.
    
    3. Vertical Riser -> Horizontal/Sloped:
       - Moves the Horizontal pipe laterally so its centerline axis passes through the Vertical Riser in Plan View.
       - Useful for aligning a branch pipe to hit a main riser.

    4. Horizontal/Sloped -> Vertical Riser:
       - Moves the Vertical pipe so it sits on the centerline axis of the Horizontal pipe in Plan View.

Limitations:
    - Alignment is calculated in the XY plane (Plan View) only. Z-elevations are not modified.
    - Non-parallel horizontal pipes (Skew lines) are not aligned (as they technically already intersect in 2D).
    - Requires pipes to be straight lines (no arcs/splines).
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

__title__ = "Align Pipes"
__doc__ = "Aligns multiple pipes to a reference pipe in the XY plane (Plan View)."
__version__ = "1.0"

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
    if not log_buffer: return
    output = script.get_output()
    output.close_others()
    for msg in log_buffer:
        output.print_md(msg)

# --- Helpers ---

def get_id_value(element_id):
    """Safely gets the integer value of an ElementId (Revit 2024+ compatible)."""
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
    """
    v_to_point = point - line_origin
    v_to_point_xy = XYZ(v_to_point.X, v_to_point.Y, 0)
    
    dot_prod = v_to_point_xy.DotProduct(line_dir_xy)
    projected_vector = line_dir_xy.Multiply(dot_prod)
    
    closest_point_xy = XYZ(line_origin.X, line_origin.Y, 0) + projected_vector
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
            return None 
    except OperationCanceledException:
        return None # Signal to stop loop

def align_pipe_geometry(ref_pipe, target_pipe):
    """
    Calculates the move vector for target_pipe to align with ref_pipe in XY.
    """
    # 1. Get Location Curves
    loc_ref = ref_pipe.Location
    loc_target = target_pipe.Location
    
    if not isinstance(loc_ref, LocationCurve) or not isinstance(loc_target, LocationCurve):
        return None, "No location curve."

    curve_ref = loc_ref.Curve
    curve_target = loc_target.Curve
    
    if not isinstance(curve_ref, Line) or not isinstance(curve_target, Line):
        return None, "Not straight line."

    # 2. Get Vectors
    p1 = curve_ref.Origin
    v1 = curve_ref.Direction.Normalize()
    p2 = curve_target.Origin
    v2 = curve_target.Direction.Normalize()
    
    ref_is_vert = is_vertical(v1)
    target_is_vert = is_vertical(v2)

    # CASE A: Both Vertical (Concentric)
    if ref_is_vert and target_is_vert:
        target_pos_xy = XYZ(p1.X, p1.Y, p2.Z)
        move_vector = target_pos_xy - p2
        return move_vector, None

    # CASE B: Ref Vertical, Target Horizontal/Sloped
    if ref_is_vert and not target_is_vert:
        v2_xy = get_xy_vector(v2)
        if not v2_xy: return None, "Target has no XY length."
        ref_projected_on_target = project_point_to_line_infinite_xy(p1, p2, v2_xy)
        move_vector = XYZ(p1.X, p1.Y, 0) - XYZ(ref_projected_on_target.X, ref_projected_on_target.Y, 0)
        return move_vector, None

    # CASE C: Ref Horizontal/Sloped, Target Vertical
    if not ref_is_vert and target_is_vert:
        v1_xy = get_xy_vector(v1)
        if not v1_xy: return None, "Ref has no XY length."
        target_projected_on_ref = project_point_to_line_infinite_xy(p2, p1, v1_xy)
        move_vector = XYZ(target_projected_on_ref.X, target_projected_on_ref.Y, 0) - XYZ(p2.X, p2.Y, 0)
        return move_vector, None

    # CASE D: Both Horizontal/Sloped
    v1_xy = get_xy_vector(v1)
    v2_xy = get_xy_vector(v2)

    if are_parallel(v1_xy, v2_xy):
        target_pos = project_point_to_line_infinite_xy(p2, p1, v1_xy)
        move_vector = target_pos - p2
        move_vector = XYZ(move_vector.X, move_vector.Y, 0)
        return move_vector, None
    else:
        return XYZ.Zero, "Non-parallel horizontal pipes."

# --- Main Execution ---

def main():
    uidoc = revit.uidoc
    doc = revit.doc
    
    log_section("Initialization")
    log_item("Tool", "Align Multiple Pipes (v1.0)")
    log_item("Active View", doc.ActiveView.Name)

    # 1. Select Reference Pipe (Once)
    ref_pipe = pick_pipe_safely(uidoc, doc, "Select REFERENCE Pipe (Stationary)")
    if not ref_pipe:
        return # Exit if no reference picked
    
    ref_id = get_id_value(ref_pipe.Id)
    log_item("Reference Pipe", ref_id)

    count_success = 0
    count_fail = 0

    # 2. Loop for Target Pipes
    while True:
        try:
            target_ref = uidoc.Selection.PickObject(
                ObjectType.Element, 
                "Select Pipe to ALIGN (ESC to finish)"
            )
            target_pipe = doc.GetElement(target_ref)
            
            if not is_pipe(target_pipe):
                continue 
            
            t_id = get_id_value(target_pipe.Id)
            
            if t_id == ref_id:
                continue 

            # Calculate
            move_vector, error = align_pipe_geometry(ref_pipe, target_pipe)
            
            if error:
                log_item("Pipe {}".format(t_id), "Failed: {}".format(error))
                count_fail += 1
                continue
            
            if move_vector.IsZeroLength():
                log_item("Pipe {}".format(t_id), "Already Aligned")
                count_success += 1
                continue

            # Execute Immediate Move
            try:
                with Transaction(doc, "Align Pipe") as t:
                    t.Start()
                    ElementTransformUtils.MoveElement(doc, target_pipe.Id, move_vector)
                    t.Commit()
                log_item("Pipe {}".format(t_id), "Aligned Success")
                count_success += 1
                
                uidoc.RefreshActiveView() 
                
            except Exception as e:
                log_item("Pipe {}".format(t_id), "Transaction Error: {}".format(str(e)))
                count_fail += 1

        except OperationCanceledException:
            # User pressed ESC
            break

    log_section("Summary")
    log_item("Total Aligned", count_success)
    log_item("Total Failed", count_fail)
    
    show_log()

if __name__ == '__main__':
    main()
