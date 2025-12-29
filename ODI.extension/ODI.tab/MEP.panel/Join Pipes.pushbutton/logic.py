# -*- coding: utf-8 -*-
import math
from data_model import JoinSolution

# Constants
TO_RAD = math.pi / 180.0
TO_DEG = 180.0 / math.pi

class Solver(object):
    """
    Pure Logic Engine. 
    Does not import Revit API. Operates on Abstract Geometry (Points, Vectors).
    """
    def __init__(self, settings):
        self.settings = settings

    def calculate_solutions(self, source_data, target_data):
        solutions = []
        if not source_data or not target_data: return []
        
        # 1. Extract Geometry
        p1 = source_data["p1"]
        p2 = source_data["p2"]
        p3 = target_data["p1"]
        p4 = target_data["p2"]
        
        # Convert to vectors relative to origin for math
        # Helper: Closest Points between two infinite lines
        def closest_points_lines(p1, p2, p3, p4):
            # Algorithm: http://paulbourke.net/geometry/pointlineplane/
            p13 = p1 - p3
            p43 = p4 - p3
            if abs(p43.X) < 1e-9 and abs(p43.Y) < 1e-9 and abs(p43.Z) < 1e-9: return None
            p21 = p2 - p1
            if abs(p21.X) < 1e-9 and abs(p21.Y) < 1e-9 and abs(p21.Z) < 1e-9: return None

            d1343 = p13.DotProduct(p43)
            d4321 = p43.DotProduct(p21)
            d1321 = p13.DotProduct(p21)
            d4343 = p43.DotProduct(p43)
            d2121 = p21.DotProduct(p21)

            denom = d2121 * d4343 - d4321 * d4321
            if abs(denom) < 1e-9:
                return None # Parallel

            numer = d1343 * d4321 - d1321 * d4343
            mua = numer / denom
            mub = (d1343 + d4321 * mua) / d4343

            pa = p1 + p21 * mua
            pb = p3 + p43 * mub
            return (pa, pb)

        # 2. Calculate Intersection
        pts = closest_points_lines(p1, p2, p3, p4)
        
        if pts:
            pt_a, pt_b = pts
            dist = pt_a.DistanceTo(pt_b)
            
            # Case A: Intersection (Coplanar-ish)
            if dist < 0.1: # 0.1 ft tolerance
                s1 = JoinSolution("Standard Trim/Elbow", "Extend pipes to intersection and add elbow.")
                s1.is_valid = True
                s1.context = {
                    "meet_pt_a": pt_a,
                    "meet_pt_b": pt_a, # Same point
                    "id_a": source_data["id"],
                    "id_b": target_data["id"]
                }
                solutions.append(s1)
            else:
                # Case B: Skew (Rolling Offset)
                s2 = JoinSolution("Rolling Offset", "Connect skew pipes with intermediate segment.")
                s2.is_valid = True
                s2.context = {
                    "meet_pt_a": pt_a,
                    "meet_pt_b": pt_b,
                    "id_a": source_data["id"],
                    "id_b": target_data["id"]
                }
                solutions.append(s2)
                
        else:
            # Parallel Case
            s3 = JoinSolution("Parallel Offset", "Pipes are parallel.")
            s3.is_valid = False # Not implemented yet
            solutions.append(s3)

        return solutions