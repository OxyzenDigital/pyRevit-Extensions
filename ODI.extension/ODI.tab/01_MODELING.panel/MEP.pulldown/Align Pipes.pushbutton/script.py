# -*- coding: utf-8 -*-
"""
Align Pipes Tool
Version: 1.3
Author: ODI (Generated via Gemini CLI)

Description:
    Allows the user to align multiple pipes to a single reference pipe in the XY plane.
    Designed to facilitate the joining of pipes by ensuring their centerlines intersect 
    or are collinear in the Plan View (XY projection).

Usage:
    1. Run the command.
    2. Select the 'Reference Pipe' (the stationary pipe).
    3. Continuously select 'Target Pipes' to align them to the current Reference.
    4. Press ESC to finish Target selection and pick a NEW Reference Pipe.
    5. Press ESC during Reference selection to Exit the tool.

Supported Alignment Modes:
    1. Horizontal/Sloped -> Horizontal/Sloped:
       - Moves the Target pipe vertically (Z-axis) to match the Reference pipe's elevation.
    
    2. Vertical Riser -> Vertical Riser:
       - Moves the Target pipe in X or Y (whichever is shorter) to align with the Reference pipe's grid.
    
    3. Vertical Riser -> Horizontal/Sloped:
       - Moves the Horizontal pipe laterally so its centerline axis passes through the Vertical Riser in Plan View.
       - Useful for aligning a branch pipe to hit a main riser.

    4. Horizontal/Sloped -> Vertical Riser:
       - Moves the Vertical pipe so it sits on the centerline axis of the Horizontal pipe in Plan View.

Limitations:
    - Requires pipes to be straight lines (no arcs/splines).
"""

from Autodesk.Revit.DB import (
    Transaction,
    BuiltInCategory,
    XYZ,
    Line,
    ElementTransformUtils,
    LocationCurve,
    GraphicsStyle,
    PartType,
    ConnectorType,
    ElementId,
    OverrideGraphicSettings,
    Color
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List
from pyrevit import revit, script, forms


__title__ = "Align Pipes"
__doc__ = "Aligns multiple pipes to a reference pipe in the XY plane (Plan View)."
__version__ = "1.3"
__context__ = "doc-project"

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
    while True:
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
            element = doc.GetElement(ref)
            if is_pipe(element):
                return element
            # If not a pipe, loop continues (effectively ignoring the click)
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

    # CASE A: Both Vertical (Orthogonal Alignment)
    if ref_is_vert and target_is_vert:
        dx = p1.X - p2.X
        dy = p1.Y - p2.Y
        if abs(dx) < abs(dy):
            return XYZ(dx, 0, 0), None # Align X
        else:
            return XYZ(0, dy, 0), None # Align Y

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

    # CASE D: Both Horizontal/Sloped (Align Elevation)
    dz = p1.Z - p2.Z
    return XYZ(0, 0, dz), None

def is_movable_category(category):
    if not category: return False
    cid = get_id_value(category.Id)
    return cid in [
        int(BuiltInCategory.OST_PipeCurves),
        int(BuiltInCategory.OST_PlaceHolderPipes),
        int(BuiltInCategory.OST_PipeFitting),
        int(BuiltInCategory.OST_PipeAccessory)
    ]

def smart_move_pipe(doc, pipe, move_vector, ref_id=None):
    """
    Moves the pipe and intelligently handles connections:
    - Moves attached fittings/accessories to maintain connectivity.
    - Disconnects if connected to Equipment/Fixtures to preserve their location.
    """
    ids_to_move = List[ElementId]()
    ids_to_move.Add(pipe.Id)
    
    # We need to check both ends
    connectors = pipe.ConnectorManager.Connectors
    
    for conn in connectors:
        if not conn.IsConnected:
            continue
            
        neighbor_conn = None
        for c in conn.AllRefs:
            if c.Owner.Id != pipe.Id and c.ConnectorType != ConnectorType.Logical:
                 neighbor_conn = c
                 break
        
        if not neighbor_conn:
            continue
            
        neighbor_elem = neighbor_conn.Owner
        
        # Safety: Never move the Reference Pipe
        if ref_id and get_id_value(neighbor_elem.Id) == ref_id:
            conn.DisconnectFrom(neighbor_conn)
            continue

        # Logic: Should we move the neighbor fitting too?
        should_move_neighbor = False
        
        # Check if neighbor is a fitting or accessory
        if neighbor_elem.Category:
            cat_id = get_id_value(neighbor_elem.Category.Id)
            if cat_id in [int(BuiltInCategory.OST_PipeFitting), int(BuiltInCategory.OST_PipeAccessory)]:
                # Check if this fitting is anchored to something immovable
                is_anchored = False
                if hasattr(neighbor_elem, "MEPModel") and neighbor_elem.MEPModel:
                    try:
                        fit_conns = neighbor_elem.MEPModel.ConnectorManager.Connectors
                        for fc in fit_conns:
                            if not fc.IsConnected: continue
                            for fcr in fc.AllRefs:
                                if fcr.ConnectorType == ConnectorType.Logical: continue
                                other_elem = fcr.Owner
                                if other_elem.Id == neighbor_elem.Id: continue
                                if get_id_value(other_elem.Id) == get_id_value(pipe.Id): continue # Back to target
                                
                                # Check if connected to Reference Pipe
                                if ref_id and get_id_value(other_elem.Id) == ref_id:
                                    is_anchored = True
                                    break
                                
                                # Check if connected to non-movable category
                                if not is_movable_category(other_elem.Category):
                                    is_anchored = True
                                    break
                            if is_anchored: break
                    except Exception:
                        pass
                
                if not is_anchored:
                    should_move_neighbor = True

        if should_move_neighbor:
            if neighbor_elem.Id not in ids_to_move:
                ids_to_move.Add(neighbor_elem.Id)
        else:
            # Disconnect to preserve the neighbor's position (Equipment, Fixtures, etc.)
            conn.DisconnectFrom(neighbor_conn)
            
    # Execute Move
    ElementTransformUtils.MoveElements(doc, ids_to_move, move_vector)

def toggle_highlight(doc, element_id, enable=True):
    """Applies or removes a color override to the element in the active view."""
    try:
        with Transaction(doc, "Toggle Highlight") as t:
            t.Start()
            if enable:
                ogs = OverrideGraphicSettings()
                ogs.SetProjectionLineColor(Color(255, 128, 0)) # Orange
                ogs.SetProjectionLineWeight(6) # Thick line
                doc.ActiveView.SetElementOverrides(element_id, ogs)
            else:
                doc.ActiveView.SetElementOverrides(element_id, OverrideGraphicSettings())
            t.Commit()
        revit.uidoc.RefreshActiveView()
    except Exception:
        pass

# --- Main Execution ---

def main():
    uidoc = revit.uidoc
    doc = revit.doc
    
    log_section("Initialization")
    log_item("Tool", "Align Pipes (Multiple) (v1.3)")
    log_item("Active View", doc.ActiveView.Name)

    count_success = 0
    count_fail = 0

    # Outer Loop: Reference Selection
    while True:
        # 1. Select Reference Pipe
        ref_pipe = pick_pipe_safely(uidoc, doc, "Select REFERENCE Pipe (Stationary) - ESC to Exit")
        if not ref_pipe:
            break # Exit tool
        
        ref_id = get_id_value(ref_pipe.Id)
        
        # Highlight Reference Pipe
        toggle_highlight(doc, ref_pipe.Id, enable=True)

        # Inner Loop: Target Selection
        while True:
            # 2. Select Target Pipe
            target_pipe = pick_pipe_safely(uidoc, doc, "Select Target Pipe to ALIGN to Ref {} - ESC to New Ref".format(ref_id))
            
            if not target_pipe:
                break # Break inner loop -> Go to select new Reference
                
            t_id = get_id_value(target_pipe.Id)
            
            if t_id == ref_id:
                continue 

            if target_pipe.Pinned:
                log_item("Pipe {}".format(t_id), "Skipped: Pinned")
                continue

            # Calculate
            move_vector, error = align_pipe_geometry(ref_pipe, target_pipe)
            
            if error:
                log_item("Pair {} -> {}".format(ref_id, t_id), "Failed: {}".format(error))
                count_fail += 1
                continue
            
            if move_vector.IsZeroLength():
                log_item("Pair {} -> {}".format(ref_id, t_id), "Already Aligned")
                count_success += 1
                continue

            # Execute Immediate Move
            try:
                with Transaction(doc, "Align Pipe") as t:
                    t.Start()
                    smart_move_pipe(doc, target_pipe, move_vector, ref_id)
                    t.Commit()
                log_item("Pair {} -> {}".format(ref_id, t_id), "Aligned Success")
                count_success += 1
                
                uidoc.RefreshActiveView()
                
            except Exception as e:
                log_item("Pair {} -> {}".format(ref_id, t_id), "Transaction Error: {}".format(str(e)))

                count_fail += 1
        
        # Clear highlight on the current reference before picking a new one
        toggle_highlight(doc, ref_pipe.Id, enable=False)


    log_section("Summary")
    log_item("Total Aligned", count_success)
    log_item("Total Failed", count_fail)
    
    show_log()

if __name__ == '__main__':
    main()
