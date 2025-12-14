from Autodesk.Revit.DB import XYZ

def calculate_spade_profile(start_pt, end_pt, curve, width=6.0, bank=0.0, step=2.0):
    """
    Returns a list of XYZ points creating a 'spade' profile along the curve.
    """
    grading_points = []
    
    line_length = curve.Length
    offset_dist = width / 2.0
    
    # Direction Check
    dist_start = start_pt.DistanceTo(curve.GetEndPoint(0))
    dist_end   = start_pt.DistanceTo(curve.GetEndPoint(1))
    is_reversed = dist_end < dist_start

    # Slope
    total_rise = end_pt.Z - start_pt.Z
    slope = total_rise / line_length
    
    current_dist = 0.0

    while current_dist <= line_length:
        # Parameter (0.0 to 1.0)
        raw_param = current_dist / line_length
        param = (1.0 - raw_param) if is_reversed else raw_param
        
        # Transform (Position + Tangent)
        transform = curve.ComputeDerivatives(param, True)
        center_pt = transform.Origin
        tangent   = transform.BasisX.Normalize() 
        
        # Normal (Perpendicular in XY plane)
        normal = XYZ(-tangent.Y, tangent.X, 0).Normalize()
        
        # Elevations
        center_z = start_pt.Z + (current_dist * slope)
        bank_z   = center_z + bank

        # Create 3 Points
        pt_center = XYZ(center_pt.X, center_pt.Y, center_z)
        
        pt_left_loc = center_pt + (normal * offset_dist)
        pt_left = XYZ(pt_left_loc.X, pt_left_loc.Y, bank_z)
        
        pt_right_loc = center_pt - (normal * offset_dist)
        pt_right = XYZ(pt_right_loc.X, pt_right_loc.Y, bank_z)

        grading_points.extend([pt_center, pt_left, pt_right])
        current_dist += step

    # Ensure End Point is locked
    grading_points.append(XYZ(end_pt.X, end_pt.Y, end_pt.Z))
    
    return grading_points