# -*- coding: utf-8 -*-
"""
Point Cloud (.rcp / .rcs) to Toposolid
Extracts bare-earth ground points from a Point Cloud and either generates a
brand-new Toposolid or updates an existing Toposolid using shape editing.
"""

__title__ = "PC to Toposolid"
__doc__ = "Extracts bare-earth ground points from a Point Cloud (.rcp/.rcs) and creates or updates a native Revit Toposolid."
__author__ = "Oxyzen Digital"
__context__ = "doc-project"

import math
import os
import heapq
import clr

# --- ASSEMBLIES ---
clr.AddReference("System")
clr.AddReference("System.Collections")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Collections.Generic import List
from System.Windows.Media import SolidColorBrush, Color as WpfColor
from System.Windows import Window, Visibility

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    XYZ,
    Plane,
    CurveLoop,
    Line,
    Toposolid,
    ToposolidType,
    Level,
    Transaction,
    TransactionStatus,
    IFailuresPreprocessor,
    FailureProcessingResult,
    FailureSeverity,
    PointCloudInstance,
    BoundingBoxXYZ
)
from Autodesk.Revit.DB.PointClouds import PointCloudFilterFactory, PointCloudFilter
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType, PickBoxStyle
from Autodesk.Revit.Exceptions import OperationCanceledException

# Try to import UIThemeManager (Revit 2024+)
try:
    from Autodesk.Revit.UI import UIThemeManager, UITheme
    HAS_THEME = True
except ImportError:
    HAS_THEME = False

from pyrevit import revit, forms
from data_model import (
    PCToToposolidViewModel,
    LevelItem,
    ToposolidTypeItem,
    ToposolidInstanceItem
)

doc = revit.doc
uidoc = revit.uidoc


def get_id_val(elem_id):
    """Safely retrieves integer or long value of an ElementId across Revit 2024+ and earlier versions."""
    if elem_id is None:
        return -1
    if hasattr(elem_id, "Value"):
        return elem_id.Value
    return elem_id.IntegerValue


def get_element_name(elem):
    """Safely retrieves the name of an Element or ElementType across Revit versions."""
    if elem is None:
        return ""
    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p and p.AsString():
            return p.AsString()
    except Exception:
        pass
    try:
        p = elem.get_Parameter(BuiltInParameter.ELEM_NAME_PARAM)
        if p and p.AsString():
            return p.AsString()
    except Exception:
        pass
    try:
        if hasattr(elem, "Name") and elem.Name:
            return str(elem.Name)
    except Exception:
        pass
    try:
        elem_id_val = get_id_val(elem.Id)
        return "Element {}".format(elem_id_val)
    except Exception:
        return "Unnamed Element"


class SilentErrorPreprocessor(IFailuresPreprocessor):
    """Suppresses minor triangulation/geometry warnings so they don't roll back the transaction."""
    def PreprocessFailures(self, failuresAccessor):
        failures = failuresAccessor.GetFailureMessages()
        for failure in failures:
            severity = failure.GetSeverity()
            if severity == FailureSeverity.Warning:
                try:
                    failuresAccessor.DeleteWarning(failure)
                except Exception:
                    pass
        return FailureProcessingResult.Continue


def ensure_toposolid_type_thickness(toposolid_type, required_thickness):
    """
    Analyzes the compound structure of the ToposolidType.
    If the current total thickness is less than required_thickness,
    expands the variable layer so the terrain never exceeds slab thickness.
    """
    if not toposolid_type:
        return 0.0
    try:
        cs = toposolid_type.GetCompoundStructure()
        if not cs:
            return 0.0

        current_total = 0.0
        for i in range(cs.LayerCount):
            current_total += cs.GetLayerWidth(i)

        if current_total < required_thickness:
            var_idx = cs.VariableLayerIndex
            if var_idx < 0 or var_idx >= cs.LayerCount:
                max_w = 0.0
                var_idx = 0
                for i in range(cs.LayerCount):
                    w = cs.GetLayerWidth(i)
                    if w > max_w:
                        max_w = w
                        var_idx = i
                try:
                    cs.VariableLayerIndex = var_idx
                except Exception:
                    pass

            current_var_w = cs.GetLayerWidth(var_idx)
            additional_needed = (required_thickness - current_total) + 5.0
            new_var_w = current_var_w + additional_needed
            cs.SetLayerWidth(var_idx, new_var_w)
            toposolid_type.SetCompoundStructure(cs)
            return current_total + additional_needed
        return current_total
    except Exception:
        return 0.0


class PointCloudSelectionFilter(ISelectionFilter):
    """Filters element selection strictly to PointCloudInstance elements."""
    def AllowElement(self, element):
        if element is None:
            return False
        if isinstance(element, PointCloudInstance) or element.GetType().Name == "PointCloudInstance":
            return True
        try:
            if element.Category and get_id_val(element.Category.Id) == int(BuiltInCategory.OST_PointClouds):
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, reference, position):
        return True


class ToposolidSelectionFilter(ISelectionFilter):
    """Filters element selection strictly to Toposolid elements."""
    def AllowElement(self, element):
        if element is None:
            return False
        if isinstance(element, Toposolid) or element.GetType().Name == "Toposolid":
            return True
        try:
            if element.Category and get_id_val(element.Category.Id) == int(BuiltInCategory.OST_Toposolid):
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, reference, position):
        return True


def get_point_cloud():
    """Returns pre-selected or user-picked PointCloudInstance (.rcp or .rcs)."""
    sel_ids = uidoc.Selection.GetElementIds()
    for eid in sel_ids:
        el = doc.GetElement(eid)
        if el and (isinstance(el, PointCloudInstance) or el.GetType().Name == "PointCloudInstance"):
            return el

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            PointCloudSelectionFilter(),
            "Click on an attached Point Cloud (.rcp / .rcs) in the view"
        )
        if ref:
            el = doc.GetElement(ref.ElementId)
            if el:
                return el
    except OperationCanceledException:
        return None
    except Exception as ex:
        forms.alert("Point cloud selection error:\n{}".format(str(ex)), title="Selection Error")
        return None
    return None


def smooth_surface_bilateral(cell_elevations, passes=2, sigma_r=2.5):
    """
    Applies edge-preserving bilateral Gaussian smoothing across the 2D elevation grid.
    """
    if not cell_elevations or passes <= 0:
        return dict(cell_elevations)

    current = dict(cell_elevations)
    kernel = [
        (0, 0, 1.0),
        (1, 0, 0.6065), (-1, 0, 0.6065), (0, 1, 0.6065), (0, -1, 0.6065),
        (1, 1, 0.3679), (-1, 1, 0.3679), (1, -1, 0.3679), (-1, -1, 0.3679)
    ]
    two_sigma_r_sq = 2.0 * (sigma_r * sigma_r)

    for _ in range(passes):
        next_elevations = {}
        for (cx, cy), z0 in current.items():
            total_weight = 0.0
            weighted_z = 0.0
            for dx, dy, w_spatial in kernel:
                k = (cx + dx, cy + dy)
                if k in current:
                    z_n = current[k]
                    dz = z_n - z0
                    w_range = math.exp(-(dz * dz) / two_sigma_r_sq)
                    w = w_spatial * w_range
                    weighted_z += w * z_n
                    total_weight += w
            if total_weight > 0.0:
                next_elevations[(cx, cy)] = weighted_z / total_weight
            else:
                next_elevations[(cx, cy)] = z0
        current = next_elevations

    return current


def _separable_morph_open(z_dict, col_count, row_count, radius):
    """
    Exact 2D Morphological Opening (Erosion followed by Dilation) using fast 1D separable
    horizontal + vertical passes. Evaluates 100% of cells in a (2*radius+1)x(2*radius+1) box
    without sparse-ring artifacts.
    """
    # 1. Erosion Horizontal
    erode_h = {}
    for (c, r), val in z_dict.items():
        m = val
        for dc in range(-radius, radius + 1):
            nb = (c + dc, r)
            if nb in z_dict and z_dict[nb] < m:
                m = z_dict[nb]
        erode_h[(c, r)] = m

    # 2. Erosion Vertical
    erode_hv = {}
    for (c, r), val in erode_h.items():
        m = val
        for dr in range(-radius, radius + 1):
            nb = (c, r + dr)
            if nb in erode_h and erode_h[nb] < m:
                m = erode_h[nb]
        erode_hv[(c, r)] = m

    # 3. Dilation Horizontal
    dilate_h = {}
    for (c, r), val in erode_hv.items():
        m = val
        for dc in range(-radius, radius + 1):
            nb = (c + dc, r)
            if nb in erode_hv and erode_hv[nb] > m:
                m = erode_hv[nb]
        dilate_h[(c, r)] = m

    # 4. Dilation Vertical
    opened = {}
    for (c, r), val in dilate_h.items():
        m = val
        for dr in range(-radius, radius + 1):
            nb = (c, r + dr)
            if nb in dilate_h and dilate_h[nb] > m:
                m = dilate_h[nb]
        opened[(c, r)] = m

    return opened


def _compute_2d_convex_hull(points, min_edge_len=2.0):
    """
    Computes a simplified 2D Convex Hull (Monotone Chain) of Revit XYZ points.
    Returns a list of XYZ points in counter-clockwise order with every edge >= min_edge_len.
    """
    pts = sorted(points, key=lambda p: (round(p.X, 4), round(p.Y, 4)))
    unique_pts = []
    for p in pts:
        if not unique_pts or math.hypot(p.X - unique_pts[-1].X, p.Y - unique_pts[-1].Y) > 0.05:
            unique_pts.append(p)

    if len(unique_pts) < 3:
        return []

    def cross(o, a, b):
        return (a.X - o.X) * (b.Y - o.Y) - (a.Y - o.Y) * (b.X - o.X)

    lower = []
    for p in unique_pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 1e-4:
            lower.pop()
        lower.append(p)

    upper = []
    for p in reversed(unique_pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 1e-4:
            upper.pop()
        upper.append(p)

    hull = lower[:-1] + upper[:-1]
    if len(hull) < 3:
        return []

    simplified = [hull[0]]
    for p in hull[1:]:
        if math.hypot(p.X - simplified[-1].X, p.Y - simplified[-1].Y) >= min_edge_len:
            simplified.append(p)

    if len(simplified) > 2 and math.hypot(simplified[0].X - simplified[-1].X, simplified[0].Y - simplified[-1].Y) < min_edge_len:
        simplified.pop()

    return simplified if len(simplified) >= 3 else []


def build_progressive_ground_terrain(
    grid_clusters,
    col_count,
    row_count,
    min_x,
    min_y,
    resolution,
    strategy_idx=0,
    filter_noise=True,
    max_slope=0.35,
    elevation_tolerance=2.0,
    obstacle_mode=0,
    smoothing_mode=0,
    footprint_out=None
):
    """
    6-Stage Bare-Earth Terrain Extractor:
    - Stage 1: Vertical Layer Clustering per cell (splits overhangs, ceilings, tree canopies, and wires from ground).
    - Stage 2: Sub-Floor Trench/Pit Recovery & Disconnected Floating Island Pruning.
    - Stage 3: 8-Ray Cliff-Bounded Plateau & Protrusion Detector (removes commercial roofs up to 120 ft wide, cars, boxes, poles).
    - Stage 4: Slope-Continuous Connected Component Segmentation & 3x3 Median Despiker.
    - Stage 5: Interior Obstacle Infilling (8-Directional IDW + Laplace Relaxation strictly inside the scanned footprint).
    - Stage 6: Edge-Preserving Bilateral Smoothing.
    """
    neighbors_8 = [
        (-1, 0, 1.0), (1, 0, 1.0), (0, -1, 1.0), (0, 1, 1.0),
        (-1, -1, 1.414), (-1, 1, 1.414), (1, -1, 1.414), (1, 1, 1.414)
    ]
    dir_8 = [
        (-1, 0), (1, 0), (0, -1), (0, 1),
        (-1, -1), (-1, 1), (1, -1), (1, 1)
    ]

    # =========================================================================
    # STAGE 1: Vertical Layer Clustering per Cell
    # Splits each cell's Z points by vertical air gaps (> 1.2 ft) to isolate the
    # lowest physical surface cluster (ground/floor) from overhanging roofs,
    # ceilings, tree canopies, and overhead wires.
    # =========================================================================
    raw_layers = {}
    cell_min = {}
    cell_span = {}
    cell_count = {}

    for (c, r), z_list in grid_clusters.items():
        if not z_list or c < 0 or c >= col_count or r < 0 or r >= row_count:
            continue
        z_s = sorted(z_list)
        n = len(z_s)
        cell_count[(c, r)] = n

        clusters = []
        curr_cluster = [z_s[0]]
        for i in range(1, n):
            if (z_s[i] - z_s[i - 1]) > 1.2:
                clusters.append(curr_cluster)
                curr_cluster = [z_s[i]]
            else:
                curr_cluster.append(z_s[i])
        clusters.append(curr_cluster)

        min_cluster_pts = max(2, int(n * 0.06))
        valid_clusters = [cl for cl in clusters if len(cl) >= min_cluster_pts]
        if not valid_clusters:
            valid_clusters = clusters

        layer_infos = []
        for cl in valid_clusters:
            cn = len(cl)
            if strategy_idx == 0:
                l_rep = cl[max(0, min(int(cn * 0.20), cn - 1))]
            elif strategy_idx == 1:
                l_rep = cl[0]
            else:
                l_rep = cl[cn // 2]
            l_span = cl[max(0, min(int(cn * 0.90), cn - 1))] - cl[max(0, min(int(cn * 0.10), cn - 1))]
            layer_infos.append((l_rep, l_span, cn))

        raw_layers[(c, r)] = layer_infos
        cell_min[(c, r)] = layer_infos[0][0]
        cell_span[(c, r)] = layer_infos[0][1]

    if not cell_min:
        return List[XYZ]()

    # =========================================================================
    # STAGE 2: Sub-Floor Pit Recovery & Disconnected Floating Island Pruning
    # =========================================================================
    for (c, r), layers in raw_layers.items():
        if len(layers) >= 2:
            ring_z = []
            for dc in range(-3, 4):
                for dr in range(-3, 4):
                    if abs(dc) <= 1 and abs(dr) <= 1:
                        continue
                    nb = (c + dc, r + dr)
                    if nb in cell_min:
                        ring_z.append(cell_min[nb])
            if ring_z:
                ring_z.sort()
                local_ref_z = ring_z[len(ring_z) // 2]
                if (local_ref_z - layers[0][0]) > 2.0 and abs(layers[1][0] - local_ref_z) < abs(layers[0][0] - local_ref_z):
                    cell_min[(c, r)] = layers[1][0]
                    cell_span[(c, r)] = layers[1][1]

    # Remove sparse single-cell airborne wire hits
    for (c, r) in list(cell_min.keys()):
        nb_vals = [cell_min[(c + dc, r + dr)] for dc, dr, _ in neighbors_8 if (c + dc, r + dr) in cell_min]
        if nb_vals:
            nb_vals.sort()
            med_nb = nb_vals[len(nb_vals) // 2]
            if cell_count[(c, r)] <= 3 and abs(cell_min[(c, r)] - med_nb) > 2.5:
                del cell_min[(c, r)]
                del cell_span[(c, r)]

    # Prune disconnected spatial islands in cell_min that sit above/below the main site across empty space
    visited_spatial = set()
    spatial_comps = []
    for cell in cell_min.keys():
        if cell in visited_spatial:
            continue
        comp = []
        q = [cell]
        visited_spatial.add(cell)
        while q:
            curr = q.pop()
            comp.append(curr)
            cc, cr = curr
            for dc, dr, _ in neighbors_8:
                nb = (cc + dc, cr + dr)
                if nb in cell_min and nb not in visited_spatial:
                    visited_spatial.add(nb)
                    q.append(nb)
        spatial_comps.append(comp)

    if len(spatial_comps) > 1:
        spatial_comps.sort(key=lambda item: len(item), reverse=True)
        main_spatial = set(spatial_comps[0])
        for comp in spatial_comps[1:]:
            comp_z_list = [cell_min[c] for c in comp]
            comp_z_list.sort()
            comp_med_z = comp_z_list[len(comp_z_list) // 2]
            best_d = float("inf")
            near_z = comp_med_z
            for sc, sr in comp[::max(1, len(comp) // 15)]:
                for mc, mr in list(main_spatial)[::max(1, len(main_spatial) // 200)]:
                    d = math.hypot(sc - mc, sr - mr)
                    if d < best_d:
                        best_d = d
                        near_z = cell_min[(mc, mr)]
            # Only prune tiny specks or islands whose elevation is inconsistent with nearby terrain
            # (size alone must never drop a legit corner of the site separated by a scan gap)
            if len(comp) < 8 or abs(comp_med_z - near_z) > (max_slope * best_d * resolution + elevation_tolerance):
                for c in comp:
                    del cell_min[c]
                    if c in cell_span:
                        del cell_span[c]

    if not cell_min:
        return List[XYZ]()

    # Record the full scanned footprint (incl. cells later classified as objects) so the
    # Toposolid boundary covers the whole point cloud extent, not only confirmed-ground cells
    if footprint_out is not None:
        for (c, r) in cell_min.keys():
            footprint_out.append((min_x + c * resolution, min_y + r * resolution))

    # =========================================================================
    # STAGE 3: 8-Ray Cliff-Bounded Plateau & Protrusion Detector
    # Marches up to 120 ft in 8 compass directions from every cell.
    # Any building roof (even 100+ ft wide, or touching a ramp/tree bridge, or
    # sitting on the scan boundary), car, box, canopy, or wire drops off a steep
    # wall cliff in a majority of valid directions, whereas true terrain continues
    # smoothly without cliff drops.
    # =========================================================================
    is_object = set()
    local_step_tol = min(0.65, elevation_tolerance * 0.35)

    if filter_noise:
        for (c, r), span in cell_span.items():
            if span > 1.5:
                is_object.add((c, r))

        max_ray_steps = max(6, int(math.ceil(120.0 / resolution)))

        for (c, r), z_p in cell_min.items():
            if (c, r) in is_object:
                continue

            cliff_drop_dirs = 0
            cliff_pit_dirs = 0
            smooth_reach_dirs = 0
            valid_dirs = 0

            for dc, dr in dir_8:
                step_unit = math.hypot(dc, dr) * resolution
                prev_z = z_p
                prev_prev_z = z_p
                prev_s = 0
                prev_prev_s = 0
                found_any = False
                hit_cliff_drop = False
                hit_cliff_pit = False
                max_smooth_dist = 0.0

                for s in range(1, max_ray_steps + 1):
                    nc = c + dc * s
                    nr = r + dr * s
                    if not (0 <= nc < col_count and 0 <= nr < row_count):
                        break
                    if (nc, nr) not in cell_min:
                        continue

                    found_any = True
                    curr_z = cell_min[(nc, nr)]
                    ds_1 = (s - prev_s) * step_unit
                    ds_2 = (s - prev_prev_s) * step_unit
                    total_d = s * step_unit

                    # Check if ray stepped down a steep wall cliff (over 1 or 2 steps) to lower ground
                    step_drop_1 = prev_z - curr_z
                    step_drop_2 = prev_prev_z - curr_z
                    allow_1 = (max_slope * ds_1) + local_step_tol
                    allow_2 = (max_slope * ds_2) + (local_step_tol * 1.3)

                    if (step_drop_1 > allow_1 or step_drop_2 > allow_2) and (z_p - curr_z) > 1.8:
                        hit_cliff_drop = True
                        break

                    # Also check if ray stepped UP a steep wall cliff out of a sub-floor trench/pit
                    if (-step_drop_1 > allow_1 or -step_drop_2 > allow_2) and (curr_z - z_p) > 1.8:
                        hit_cliff_pit = True
                        break

                    # Check if within close range (<= 20 ft) the cell protrudes sharply above ground
                    # (skipped across unscanned gaps: slope is unknown there, full-slope cliff test above still applies)
                    if (s - prev_s) == 1 and total_d <= 20.0 and (z_p - curr_z) > ((max_slope * total_d * 0.6) + 1.2):
                        hit_cliff_drop = True
                        break

                    if not hit_cliff_drop and not hit_cliff_pit:
                        max_smooth_dist = total_d

                    prev_prev_z = prev_z
                    prev_prev_s = prev_s
                    prev_z = curr_z
                    prev_s = s

                if found_any:
                    valid_dirs += 1
                    if hit_cliff_drop:
                        cliff_drop_dirs += 1
                    elif hit_cliff_pit:
                        cliff_pit_dirs += 1
                    elif max_smooth_dist >= min(45.0, max_ray_steps * resolution * 0.5):
                        smooth_reach_dirs += 1

            if valid_dirs > 0:
                if cliff_drop_dirs >= max(3, int(math.ceil(valid_dirs * 0.60))) or (cliff_drop_dirs >= 4 and smooth_reach_dirs <= 2):
                    is_object.add((c, r))
                elif cliff_pit_dirs >= max(3, int(math.ceil(valid_dirs * 0.60))):
                    is_object.add((c, r))

        # =========================================================================
        # STAGE 4: Slope-Continuous Connected Component Segmentation
        # =========================================================================
        smooth_cells = set()
        for k in cell_min.keys():
            if k not in is_object:
                smooth_cells.add(k)
        adj = {k: [] for k in smooth_cells}

        for (c, r) in smooth_cells:
            z_curr = cell_min[(c, r)]
            for dc, dr, dist_f in neighbors_8:
                nb = (c + dc, r + dr)
                if nb in smooth_cells:
                    allow_dz = (max_slope * resolution * dist_f) + local_step_tol
                    if abs(cell_min[nb] - z_curr) <= allow_dz:
                        adj[(c, r)].append(nb)
                elif nb not in cell_min:
                    # Bridge across 1-to-3 empty scan-shadow cells ONLY (never jump over populated obstacles!)
                    for step in (2, 3, 4):
                        far_nb = (c + dc * step, r + dr * step)
                        if far_nb in smooth_cells:
                            far_dist = dist_f * step
                            allow_dz = (max_slope * resolution * far_dist) + local_step_tol
                            if abs(cell_min[far_nb] - z_curr) <= allow_dz:
                                adj[(c, r)].append(far_nb)
                                adj[far_nb].append((c, r))
                            break
                        elif far_nb in cell_min:
                            break

        visited = set()
        components = []
        for cell in smooth_cells:
            if cell in visited:
                continue
            comp = []
            q = [cell]
            visited.add(cell)
            while q:
                curr = q.pop()
                comp.append(curr)
                for nb in adj.get(curr, ()):
                    if nb not in visited:
                        visited.add(nb)
                        q.append(nb)
            components.append(comp)

        if components:
            components.sort(key=lambda item: len(item), reverse=True)
            confirmed_ground_set = set(components[0])

            main_cells_list = list(confirmed_ground_set)
            bucket_size = 8
            main_buckets = {}
            for mc, mr in main_cells_list:
                bk = (mc // bucket_size, mr // bucket_size)
                if bk not in main_buckets:
                    main_buckets[bk] = []
                main_buckets[bk].append((mc, mr))

            min_valid_terrace_size = 8

            for comp in components[1:]:
                comp_size = len(comp)
                if comp_size < min_valid_terrace_size:
                    for c in comp:
                        is_object.add(c)
                    continue

                sample_step = max(1, comp_size // 30)
                diffs = []
                slopes = []
                for sc, sr in comp[::sample_step]:
                    sz = cell_min[(sc, sr)]
                    best_d = float("inf")
                    best_mz = sz
                    sbk_c, sbk_r = sc // bucket_size, sr // bucket_size
                    for rad in range(0, max(col_count, row_count) // bucket_size + 1):
                        found = False
                        for dbc in range(-rad, rad + 1):
                            for dbr in range(-rad, rad + 1):
                                if max(abs(dbc), abs(dbr)) != rad:
                                    continue
                                for mc, mr in main_buckets.get((sbk_c + dbc, sbk_r + dbr), ()):
                                    d = math.hypot(sc - mc, sr - mr)
                                    if d < best_d:
                                        best_d = d
                                        best_mz = cell_min[(mc, mr)]
                                        found = True
                        if found and best_d <= (rad + 1) * bucket_size * 1.414:
                            break
                    if best_d < float("inf") and best_d > 0:
                        dz = sz - best_mz
                        diffs.append(dz)
                        slopes.append(abs(dz) / (best_d * resolution))

                if diffs:
                    diffs.sort()
                    slopes.sort()
                    boundary_slope = slopes[max(0, len(slopes) // 4)]
                    med_dz = diffs[len(diffs) // 2]
                    # Reject only cliff-bounded plateaus (steep AND tall); gently graded terraces/embankments stay
                    if boundary_slope > (max_slope * 1.15) and abs(med_dz) > 1.8:
                        for c in comp:
                            is_object.add(c)
                    else:
                        for c in comp:
                            confirmed_ground_set.add(c)
                else:
                    for c in comp:
                        is_object.add(c)

            confirmed_ground = {k: cell_min[k] for k in confirmed_ground_set}
        else:
            confirmed_ground = dict(cell_min)
            is_object.clear()

        if len(confirmed_ground) < max(10, int(len(cell_min) * 0.15)):
            confirmed_ground = dict(cell_min)
            is_object.clear()
    else:
        confirmed_ground = dict(cell_min)

    # 3x3 Median Despiker for any single-cell residual spike or dip
    spike_limit = max(0.8, max_slope * resolution * 1.2)
    for (c, r), z_val in list(confirmed_ground.items()):
        nb_z = [
            confirmed_ground[(c + dc, r + dr)]
            for dc in (-1, 0, 1) for dr in (-1, 0, 1)
            if not (dc == 0 and dr == 0) and (c + dc, r + dr) in confirmed_ground
        ]
        if len(nb_z) >= 3:
            nb_z.sort()
            med = nb_z[len(nb_z) // 2]
            if abs(z_val - med) > spike_limit:
                confirmed_ground[(c, r)] = med

    # =========================================================================
    # STAGE 5: Interior Obstacle Infilling (8-Directional IDW + Laplace Relaxation)
    # Strictly infills obstacles that lie inside the scanned site boundary!
    # =========================================================================
    ground_dem = dict(confirmed_ground)
    infill_cells = set()

    if obstacle_mode == 0 and is_object:
        max_search = max(col_count, row_count)
        for (c, r) in is_object:
            hit_map = {}
            for idx_d, (dc, dr) in enumerate(dir_8):
                step_len = math.hypot(dc, dr)
                for s in range(1, max_search):
                    nc = c + dc * s
                    nr = r + dr * s
                    if not (0 <= nc < col_count and 0 <= nr < row_count):
                        break
                    if (nc, nr) in confirmed_ground:
                        hit_map[idx_d] = (s * step_len, confirmed_ground[(nc, nr)])
                        break

            bracketed = (
                (0 in hit_map and 1 in hit_map) or
                (2 in hit_map and 3 in hit_map) or
                (4 in hit_map and 7 in hit_map) or
                (5 in hit_map and 6 in hit_map)
            )
            if (bracketed and len(hit_map) >= 4) or len(hit_map) >= 5:
                infill_cells.add((c, r))
                hits = list(hit_map.values())
                w_sum = sum(1.0 / (d * d) for d, _ in hits)
                z_sum = sum((1.0 / (d * d)) * z for d, z in hits)
                ground_dem[(c, r)] = z_sum / w_sum

        if infill_cells:
            ortho = [(-1, 0), (1, 0), (0, -1), (0, 1)]
            for _ in range(50):
                next_v = {}
                for (c, r) in infill_cells:
                    vals = [ground_dem[(c + dc, r + dr)] for dc, dr in ortho if (c + dc, r + dr) in ground_dem]
                    if vals:
                        next_v[(c, r)] = sum(vals) / float(len(vals))
                for k, v in next_v.items():
                    ground_dem[k] = v

    # =========================================================================
    # STAGE 6: Bilateral Smoothing & Final Point List Assembly
    # =========================================================================
    if smoothing_mode == 0:
        ground_dem = smooth_surface_bilateral(ground_dem, passes=2, sigma_r=2.5)
    elif smoothing_mode == 1:
        ground_dem = smooth_surface_bilateral(ground_dem, passes=1, sigma_r=1.5)
    elif smoothing_mode == 2:
        ground_dem = smooth_surface_bilateral(ground_dem, passes=3, sigma_r=3.5)

    sampled_points = List[XYZ]()
    for (c, r), node_z in ground_dem.items():
        node_x = min_x + (c * resolution)
        node_y = min_y + (r * resolution)
        sampled_points.Add(XYZ(node_x, node_y, node_z))

    return sampled_points


# Backward compatibility aliases
build_smrf_grid_terrain = build_progressive_ground_terrain
build_regular_grid_terrain = build_progressive_ground_terrain


class PCToToposolidWindow(forms.WPFWindow):
    """WPF Host Window implementing MVVM and semantic theming."""
    def __init__(self, xaml_file_name, pc_instance):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.pc = pc_instance

        raw_levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
        level_items = []
        for lvl in raw_levels:
            lvl_name = get_element_name(lvl) or "Level"
            level_items.append(LevelItem(lvl.Id, lvl_name, lvl.Elevation))
        level_items.sort(key=lambda lvl: lvl.Elevation)

        raw_types = list(FilteredElementCollector(doc).OfClass(ToposolidType).ToElements())
        if not raw_types:
            raw_types = list(
                FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Toposolid)
                .WhereElementIsElementType()
                .ToElements()
            )
        type_items = []
        for t in raw_types:
            t_name = get_element_name(t) or "Toposolid Type"
            type_items.append(ToposolidTypeItem(t.Id, t_name))
        type_items.sort(key=lambda item: item.Name)

        raw_toposolids = list(
            FilteredElementCollector(doc)
            .OfClass(Toposolid)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        existing_items = []
        for ts in raw_toposolids:
            t_type = doc.GetElement(ts.GetTypeId())
            type_name = get_element_name(t_type) if t_type else "Toposolid"
            ts_name = get_element_name(ts) or "Toposolid"
            existing_items.append(ToposolidInstanceItem(ts.Id, ts_name, type_name))

        pc_min_z, pc_max_z = None, None
        try:
            bbox = self.pc.get_BoundingBox(None)
            if bbox:
                pc_min_z = round(bbox.Min.Z, 1)
                pc_max_z = round(bbox.Max.Z, 1)
        except Exception:
            pass

        self.view_model = PCToToposolidViewModel(level_items, type_items, existing_items, pc_min_z, pc_max_z)
        self.DataContext = self.view_model

        if not self.view_model.HasExistingToposolids:
            self.rbCreateNew.IsChecked = True
            self.rbUpdateExisting.IsEnabled = False
            self.txtNoToposolidNote.Visibility = Visibility.Visible
            self.pnlUpdateExisting.Visibility = Visibility.Collapsed
            self.pnlCreateNew.Visibility = Visibility.Visible
        else:
            self.rbUpdateExisting.IsEnabled = True
            self.txtNoToposolidNote.Visibility = Visibility.Collapsed

        self.apply_theme()
        if hasattr(self, "Btn_WinClose") and self.Btn_WinClose is not None:
            self.Btn_WinClose.Click += self.Close_Click
        if hasattr(self, "HeaderDrag") and self.HeaderDrag is not None:
            self.HeaderDrag.MouseLeftButtonDown += self.Header_MouseDown
        self.set_status(self.view_model.StatusMessage, "info")

    _STATUS_COLORS = {
        "idle":      ((100, 100, 100), (120, 120, 120)),
        "info":      ((100, 100, 100), (120, 120, 120)),
        "success":   (( 16, 124,  16), ( 16, 124,  16)),
        "error":     ((197,  15,  31), (197,  15,  31)),
        "busy":      ((  0, 120, 215), (  0, 100, 200)),
    }

    def set_status(self, msg, level="info"):
        try:
            dot_rgb, txt_rgb = self._STATUS_COLORS.get(level, self._STATUS_COLORS["info"])
            dot_brush = SolidColorBrush(WpfColor.FromRgb(*dot_rgb))
            txt_brush = SolidColorBrush(WpfColor.FromRgb(*txt_rgb))
            if hasattr(self, "Lb_Status") and self.Lb_Status is not None:
                self.Lb_Status.Text = msg
                self.Lb_Status.Foreground = txt_brush
            if hasattr(self, "Dot_Status") and self.Dot_Status is not None:
                self.Dot_Status.Fill = dot_brush
            if hasattr(self, "Tb_StatusHeader") and self.Tb_StatusHeader is not None:
                self.Tb_StatusHeader.Text = msg if len(msg) <= 46 else msg[:43] + "..."
            self.view_model.StatusMessage = msg
        except Exception:
            pass

    def apply_theme(self):
        self._is_dark = False
        if HAS_THEME:
            try:
                self._is_dark = (UIThemeManager.CurrentTheme == UITheme.Dark)
            except Exception:
                pass

        from System.Windows import SystemColors
        res = self.Resources
        if self._is_dark:
            res["WindowBrush"]        = SolidColorBrush(WpfColor.FromRgb( 45,  52,  64))
            res["ControlBrush"]       = SolidColorBrush(WpfColor.FromRgb( 35,  41,  51))
            res["TextBrush"]          = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
            res["TextLightBrush"]     = SolidColorBrush(WpfColor.FromRgb(160, 165, 175))
            res["AccentBrush"]        = SolidColorBrush(WpfColor.FromRgb(  0, 120, 215))
            res["HeaderTextBrush"]    = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            res["HeaderSubTextBrush"] = SolidColorBrush(WpfColor.FromRgb(180, 210, 240))
            res["BorderBrush"]        = SolidColorBrush(WpfColor.FromRgb( 75,  85, 100))
            res["StatusReadyBrush"]   = SolidColorBrush(WpfColor.FromRgb( 90, 210,  90))
            res["StatusErrorBrush"]   = SolidColorBrush(WpfColor.FromRgb(240, 100, 100))
            res["ButtonBrush"]        = SolidColorBrush(WpfColor.FromRgb( 60,  68,  82))
            res["HoverBrush"]         = SolidColorBrush(WpfColor.FromRgb( 70,  85, 105))
            res["PressedBrush"]       = SolidColorBrush(WpfColor.FromRgb(  0,  80, 160))
            res["SelectionBrush"]     = SolidColorBrush(WpfColor.FromRgb( 30,  58, 138))
            res["SelectionTextBrush"] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
            try:
                res[SystemColors.HighlightBrushKey] = SolidColorBrush(WpfColor.FromRgb(0, 120, 215))
                res[SystemColors.HighlightTextBrushKey] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
                res[SystemColors.WindowBrushKey] = SolidColorBrush(WpfColor.FromRgb(35, 41, 51))
                res[SystemColors.WindowTextBrushKey] = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
                res[SystemColors.ControlBrushKey] = SolidColorBrush(WpfColor.FromRgb(35, 41, 51))
                res[SystemColors.ControlTextBrushKey] = SolidColorBrush(WpfColor.FromRgb(240, 240, 240))
            except Exception:
                pass
        else:
            try:
                res[SystemColors.HighlightBrushKey] = SolidColorBrush(WpfColor.FromRgb(0, 120, 212))
                res[SystemColors.HighlightTextBrushKey] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
                res[SystemColors.WindowBrushKey] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
                res[SystemColors.WindowTextBrushKey] = SolidColorBrush(WpfColor.FromRgb(30, 30, 30))
                res[SystemColors.ControlBrushKey] = SolidColorBrush(WpfColor.FromRgb(255, 255, 255))
                res[SystemColors.ControlTextBrushKey] = SolidColorBrush(WpfColor.FromRgb(30, 30, 30))
            except Exception:
                pass

    def Header_MouseDown(self, sender, args):
        try:
            self.DragMove()
        except Exception:
            pass

    def Action_Changed(self, sender, args):
        if not hasattr(self, "pnlCreateNew") or not hasattr(self, "pnlUpdateExisting"):
            return

        is_create = self.rbCreateNew.IsChecked == True
        self.view_model.TargetModeIndex = 0 if is_create else 1

        if is_create:
            self.pnlCreateNew.Visibility = Visibility.Visible
            self.pnlUpdateExisting.Visibility = Visibility.Collapsed
            self.btnAction.Content = "Generate New Toposolid"
        else:
            self.pnlCreateNew.Visibility = Visibility.Collapsed
            self.pnlUpdateExisting.Visibility = Visibility.Visible
            self.btnAction.Content = "Update Existing Toposolid"

    def PickToposolid_Click(self, sender, args):
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                ToposolidSelectionFilter(),
                "Click on a Toposolid to update with point cloud data"
            )
            if ref:
                picked_id = ref.ElementId
                for item in self.view_model.ExistingToposolids:
                    if item.Id == picked_id:
                        self.view_model.SelectedExistingToposolid = item
                        break
        except OperationCanceledException:
            pass
        finally:
            self.Show()

    def Close_Click(self, sender, args):
        self.Close()

    def PickDetailAreas_Click(self, sender, args):
        """Lets the user drag one box per detail area; Esc finishes picking."""
        vm = self.view_model
        picked = 0
        self.Hide()
        try:
            while True:
                try:
                    pbox = uidoc.Selection.PickBox(
                        PickBoxStyle.Enclosing,
                        "Drag detail area #{} (press Esc when finished)".format(len(vm.DetailAreas) + 1)
                    )
                except OperationCanceledException:
                    break
                box = (
                    min(pbox.Min.X, pbox.Max.X), min(pbox.Min.Y, pbox.Max.Y),
                    max(pbox.Min.X, pbox.Max.X), max(pbox.Min.Y, pbox.Max.Y)
                )
                if (box[2] - box[0]) < 1.0 or (box[3] - box[1]) < 1.0:
                    continue
                vm.add_detail_area(box)
                picked += 1
        finally:
            self.Show()

        if picked:
            vm.EnableDetailAreas = True
            self.set_status("Added {} detail area(s). {}".format(picked, vm.DetailAreasSummary), "success")
        else:
            self.set_status("No detail areas added.", "info")

    def ClearDetailAreas_Click(self, sender, args):
        self.view_model.clear_detail_areas()
        self.set_status("Detail areas cleared.", "info")

    def Generate_Click(self, sender, args):
        vm = self.view_model
        is_create = vm.IsCreateNew

        try:
            min_z = float(vm.MinZOffset)
            max_z = float(vm.MaxZOffset)
            resolution = float(vm.GridResolution)
        except ValueError:
            forms.alert("Please enter valid numeric values for Z Offsets and Grid Resolution.", title="Input Error")
            return

        if resolution < 0.5:
            forms.alert("Grid resolution must be at least 0.5 ft.", title="Resolution Error")
            return

        if min_z >= max_z:
            forms.alert("Min Z Offset must be strictly less than Max Z Offset.", title="Elevation Band Error")
            return

        if not vm.SelectedLevel:
            forms.alert("Please select a valid Reference Level.", title="Selection Required")
            return

        if is_create and not vm.SelectedToposolidType:
            forms.alert("No Toposolid Type selected.", title="Type Required")
            return

        if not is_create and not vm.SelectedExistingToposolid:
            forms.alert("Please select or pick an existing Toposolid to update.", title="Toposolid Required")
            return

        filter_noise = vm.FilterNoise
        max_slope = 0.35
        elevation_tolerance = 2.0
        if filter_noise:
            try:
                max_slope = max(0.05, float(vm.MaxGroundSlope))
            except ValueError:
                forms.alert("Please enter a valid numeric value for Max Natural Terrain Slope.", title="Input Error")
                return
            try:
                elevation_tolerance = max(0.5, float(vm.GroundElevationTolerance))
            except ValueError:
                forms.alert("Please enter a valid numeric value for Elevation Step Tolerance.", title="Input Error")
                return

        obstacle_mode = vm.SelectedObstacleModeIndex
        smoothing_mode = vm.SelectedSmoothingModeIndex

        detail_areas = None
        detail_resolution = None
        if vm.EnableDetailAreas:
            if not vm.HasDetailAreas:
                forms.alert(
                    "Detail Areas are enabled but none are selected.\n\n"
                    "Click 'Pick Detail Areas in View...' or uncheck the option.",
                    title="Detail Areas Required"
                )
                return
            try:
                detail_resolution = float(vm.DetailResolution)
            except ValueError:
                forms.alert("Please enter a valid numeric Detail Sampling Spacing.", title="Input Error")
                return
            if detail_resolution < 0.5 or detail_resolution >= resolution:
                forms.alert(
                    "Detail Sampling Spacing must be at least 0.5 ft and finer than the "
                    "base Grid Sampling Spacing ({:.2f} ft).".format(resolution),
                    title="Detail Resolution Error"
                )
                return
            detail_areas = list(vm.DetailAreas)
            est_pts = sum((b[2] - b[0]) * (b[3] - b[1]) for b in detail_areas) / (detail_resolution ** 2)
            if est_pts > 40000:
                if not forms.alert(
                    "Detail areas will add roughly {:,.0f} points at {:.2f} ft spacing.\n\n"
                    "This may be slow. Continue?".format(est_pts, detail_resolution),
                    title="High Detail Density", yes=True, no=True
                ):
                    return

        self.Hide()
        try:
            success = run_extraction_and_creation(
                self.pc,
                is_create,
                vm.SelectedLevel,
                vm.SelectedToposolidType,
                vm.SelectedExistingToposolid,
                min_z,
                max_z,
                resolution,
                vm.SelectedStrategy.Index,
                vm.SelectedBoundaryMode.Index,
                filter_noise,
                max_slope,
                elevation_tolerance,
                obstacle_mode,
                smoothing_mode,
                vm,
                detail_areas=detail_areas,
                detail_resolution=detail_resolution
            )
            if success:
                self.Close()
            else:
                self.Show()
        except Exception as ex:
            forms.alert("Error generating Toposolid:\n{}".format(str(ex)), title="Operation Failed")
            self.Show()


class SamplingCancelled(Exception):
    """Raised when the user cancels the point cloud sampling progress bar."""
    pass


def sample_point_cloud_grid(
    pc,
    transform,
    min_x,
    max_x,
    min_y,
    max_y,
    resolution,
    world_min_z,
    world_max_z,
    progress_title="Sampling Point Cloud"
):
    """
    Tiled point cloud sampling: bins every point within the Z band into (c, r) grid cells.
    Returns (grid_clusters, col_count, row_count, total_raw_points). Raises SamplingCancelled.
    """
    tile_size = max(80.0, min(200.0, resolution * 30.0))

    x_steps = []
    curr_x = min_x
    while curr_x < max_x:
        next_x = min(curr_x + tile_size, max_x)
        x_steps.append((curr_x, next_x))
        curr_x = next_x

    y_steps = []
    curr_y = min_y
    while curr_y < max_y:
        next_y = min(curr_y + tile_size, max_y)
        y_steps.append((curr_y, next_y))
        curr_y = next_y

    tiles = [(tx0, tx1, ty0, ty1) for tx0, tx1 in x_steps for ty0, ty1 in y_steps]

    col_count = int(math.ceil((max_x - min_x) / resolution)) + 1
    row_count = int(math.ceil((max_y - min_y) / resolution)) + 1
    grid_clusters = {}
    total_raw_points = 0
    avg_sample_dist = max(0.15, resolution * 0.15)
    max_pts_per_tile = 800000
    tile_overlap = max(1.0, resolution * 0.5)

    pb = None
    if len(tiles) > 1:
        try:
            pb = forms.ProgressBar(title=progress_title + ": {value} of {max} tiles...", cancellable=True)
        except Exception:
            pb = None

    try:
        for idx, (tx0, tx1, ty0, ty1) in enumerate(tiles):
            if pb:
                if pb.cancelled:
                    raise SamplingCancelled()
                pb.update_progress(idx + 1, len(tiles))

            tile_model_planes = List[Plane]()
            tile_model_planes.Add(Plane.CreateByNormalAndOrigin(XYZ.BasisX, XYZ(tx0 - tile_overlap, 0, 0)))
            tile_model_planes.Add(Plane.CreateByNormalAndOrigin(-XYZ.BasisX, XYZ(tx1 + tile_overlap, 0, 0)))
            tile_model_planes.Add(Plane.CreateByNormalAndOrigin(XYZ.BasisY, XYZ(0, ty0 - tile_overlap, 0)))
            tile_model_planes.Add(Plane.CreateByNormalAndOrigin(-XYZ.BasisY, XYZ(0, ty1 + tile_overlap, 0)))
            tile_model_planes.Add(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, world_min_z)))
            tile_model_planes.Add(Plane.CreateByNormalAndOrigin(-XYZ.BasisZ, XYZ(0, 0, world_max_z)))

            tile_filter = PointCloudFilterFactory.CreateMultiPlaneFilter(tile_model_planes)

            try:
                tile_points = pc.GetPoints(tile_filter, avg_sample_dist, max_pts_per_tile)
            except Exception:
                continue

            if not tile_points or tile_points.Count == 0:
                continue

            total_raw_points += tile_points.Count

            for pt in tile_points:
                world_pt = transform.OfPoint(XYZ(pt.X, pt.Y, pt.Z))
                if world_min_z <= world_pt.Z <= world_max_z:
                    c = int(round((world_pt.X - min_x) / resolution))
                    r = int(round((world_pt.Y - min_y) / resolution))
                    if 0 <= c < col_count and 0 <= r < row_count:
                        k = (c, r)
                        if k not in grid_clusters:
                            grid_clusters[k] = []
                        grid_clusters[k].append(world_pt.Z)
    finally:
        if pb:
            try:
                pb.close()
            except Exception:
                pass

    return grid_clusters, col_count, row_count, total_raw_points


def _point_in_box(x, y, box):
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def extract_detail_area_points(pc, transform, box, detail_res, base_res, world_min_z, world_max_z, terrain_kwargs):
    """
    Re-extracts bare-earth ground inside `box` at `detail_res`.
    The box is sampled with a context margin so roof/cliff detection (Stage 3) and slope
    segmentation (Stage 4) still see the surrounding ground; only points inside `box` are returned.
    """
    margin = max(60.0, base_res * 10.0)
    x0, y0 = box[0] - margin, box[1] - margin
    x1, y1 = box[2] + margin, box[3] + margin

    grid_clusters, col_count, row_count, total = sample_point_cloud_grid(
        pc, transform, x0, x1, y0, y1, detail_res, world_min_z, world_max_z,
        progress_title="Sampling Detail Area"
    )
    if not grid_clusters or total == 0:
        return []

    pts = build_progressive_ground_terrain(
        grid_clusters, col_count, row_count, x0, y0, detail_res, **terrain_kwargs
    )
    return [p for p in pts if _point_in_box(p.X, p.Y, box)]


def merge_detail_points(coarse_points, detail_sets, boxes, detail_res):
    """
    Drops coarse points inside any detail box and appends the fine points.
    Fine points from overlapping boxes are de-duplicated with a spatial hash
    (no two points closer than 0.4 x detail_res).
    """
    merged = List[XYZ]()
    removed = 0
    for p in coarse_points:
        px = p.X
        py = p.Y
        in_any_box = False
        for b in boxes:
            if b[0] <= px <= b[2] and b[1] <= py <= b[3]:
                in_any_box = True
                break
        if in_any_box:
            removed += 1
            continue
        merged.Add(p)

    min_d = detail_res * 0.4
    min_d2 = min_d * min_d
    buckets = {}
    added = 0
    for pts in detail_sets:
        for p in pts:
            bk = (int(math.floor(p.X / min_d)), int(math.floor(p.Y / min_d)))
            too_close = False
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for q in buckets.get((bk[0] + dx, bk[1] + dy), ()):
                        if (q.X - p.X) ** 2 + (q.Y - p.Y) ** 2 < min_d2:
                            too_close = True
                            break
                    if too_close:
                        break
                if too_close:
                    break
            if too_close:
                continue
            buckets.setdefault(bk, []).append(p)
            merged.Add(p)
            added += 1
    return merged, removed, added


def refine_existing_toposolid(pc, transform, existing_toposolid_item, boxes, detail_res, base_res,
                              world_min_z, world_max_z, terrain_kwargs):
    """
    Update-mode refinement: replaces the existing Toposolid vertices inside each detail box
    with fine-resolution ground points. Everything outside the boxes is left untouched.
    """
    target_ts = doc.GetElement(existing_toposolid_item.Id)
    if not target_ts:
        forms.alert("The selected Toposolid is no longer valid in the document.", title="Update Failed")
        return False

    detail_sets = []
    for box in boxes:
        detail_sets.append(
            extract_detail_area_points(pc, transform, box, detail_res, base_res,
                                       world_min_z, world_max_z, terrain_kwargs)
        )
    fine_points, _, added = merge_detail_points([], detail_sets, boxes, detail_res)
    if fine_points.Count == 0:
        forms.alert("No ground points were found inside the selected detail areas.", title="No Points Found")
        return False

    all_z = [p.Z for p in fine_points]
    required_thickness = max(15.0, (max(all_z) - min(all_z)) + 10.0)
    ts_type = doc.GetElement(target_ts.GetTypeId())

    deleted = 0
    kept_boundary = 0
    with Transaction(doc, "Refine Toposolid Detail Areas from Point Cloud") as t:
        t.Start()
        fail_opts = t.GetFailureHandlingOptions()
        fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
        t.SetFailureHandlingOptions(fail_opts)

        if ts_type:
            ensure_toposolid_type_thickness(ts_type, required_thickness)

        shape_editor = target_ts.GetSlabShapeEditor()
        if not shape_editor:
            t.RollBack()
            return False
        if not shape_editor.IsEnabled:
            shape_editor.Enable()

        # Collect existing vertices inside the detail boxes
        doomed = []
        for v in shape_editor.SlabShapeVertices:
            pos = v.Position
            vx = pos.X
            vy = pos.Y
            in_any_box = False
            for b in boxes:
                if b[0] <= vx <= b[2] and b[1] <= vy <= b[3]:
                    in_any_box = True
                    break
            if in_any_box:
                doomed.append(v)

        if doomed:
            if hasattr(shape_editor, "DeletePoint"):
                for v in doomed:
                    try:
                        shape_editor.DeletePoint(v)
                        deleted += 1
                    except Exception:
                        # Boundary/profile vertices cannot be deleted - they simply stay
                        kept_boundary += 1
            else:
                # Older API without DeletePoint: rebuild shape keeping only vertices outside the boxes
                keep = List[XYZ]()
                for v in shape_editor.SlabShapeVertices:
                    pos = v.Position
                    vx = pos.X
                    vy = pos.Y
                    in_any_box = False
                    for b in boxes:
                        if b[0] <= vx <= b[2] and b[1] <= vy <= b[3]:
                            in_any_box = True
                            break
                    if not in_any_box:
                        keep.Add(XYZ(vx, vy, pos.Z))
                deleted = len(doomed)
                shape_editor.ResetSlabShape()
                if keep.Count > 0:
                    try:
                        shape_editor.AddPoints(keep)
                    except Exception:
                        for p in keep:
                            try:
                                shape_editor.AddPoint(p)
                            except Exception:
                                pass

        try:
            shape_editor.AddPoints(fine_points)
        except Exception:
            for p in fine_points:
                try:
                    shape_editor.AddPoint(p)
                except Exception:
                    pass

        status = t.Commit()

    if status != TransactionStatus.Committed:
        forms.alert("Refining the Toposolid could not be committed by Revit.", title="Update Failed")
        return False

    forms.toast(
        "Refined {} detail area(s) @ {:.2f} ft\n"
        "• Fine points added: {}\n"
        "• Old points removed: {}{}".format(
            len(boxes), detail_res, added, deleted,
            " ({} boundary kept)".format(kept_boundary) if kept_boundary else ""
        ),
        title="Toposolid Refined"
    )
    return True


def run_extraction_and_creation(
    pc,
    is_create,
    level_item,
    toposolid_type_item,
    existing_toposolid_item,
    min_z_off,
    max_z_off,
    resolution,
    strategy_idx,
    boundary_mode,
    filter_noise,
    max_slope,
    elevation_tolerance,
    obstacle_mode,
    smoothing_mode,
    vm,
    detail_areas=None,
    detail_resolution=None
):
    """Performs spatial filtering, extraction, clustering, and Toposolid creation or updating."""
    lvl_elev = level_item.Elevation
    world_min_z = lvl_elev + min_z_off
    world_max_z = lvl_elev + max_z_off

    transform = pc.GetTotalTransform()

    use_detail = bool(detail_areas) and detail_resolution is not None
    terrain_kwargs = dict(
        strategy_idx=strategy_idx,
        filter_noise=filter_noise,
        max_slope=max_slope,
        elevation_tolerance=elevation_tolerance,
        obstacle_mode=obstacle_mode,
        smoothing_mode=smoothing_mode
    )

    # Update mode + detail areas = refine only inside the boxes (no full-site pass)
    if use_detail and not is_create:
        try:
            return refine_existing_toposolid(
                pc, transform, existing_toposolid_item, detail_areas, detail_resolution, resolution,
                world_min_z, world_max_z, terrain_kwargs
            )
        except SamplingCancelled:
            forms.alert("Operation cancelled by user.", title="Cancelled")
            return False

    if boundary_mode == 1:
        try:
            pbox = uidoc.Selection.PickBox(
                PickBoxStyle.Enclosing,
                "Drag a box around the area of the Point Cloud to extract"
            )
            min_x = min(pbox.Min.X, pbox.Max.X)
            max_x = max(pbox.Min.X, pbox.Max.X)
            min_y = min(pbox.Min.Y, pbox.Max.Y)
            max_y = max(pbox.Min.Y, pbox.Max.Y)
        except OperationCanceledException:
            return False
    else:
        bbox = pc.get_BoundingBox(None)
        if not bbox:
            forms.alert("Unable to read Point Cloud bounding box.", title="Geometry Error")
            return False
        min_x = bbox.Min.X
        max_x = bbox.Max.X
        min_y = bbox.Min.Y
        max_y = bbox.Max.Y

    if abs(max_x - min_x) < resolution or abs(max_y - min_y) < resolution:
        forms.alert(
            "Target extraction area is too small for the selected grid resolution ({:.1f} ft).".format(resolution),
            title="Area Too Small"
        )
        return False

    try:
        grid_clusters, col_count, row_count, total_raw_points = sample_point_cloud_grid(
            pc, transform, min_x, max_x, min_y, max_y, resolution, world_min_z, world_max_z
        )
    except SamplingCancelled:
        forms.alert("Operation cancelled by user.", title="Cancelled")
        return False

    if not grid_clusters or total_raw_points == 0:
        forms.alert(
            "No points were found within the specified boundary and Z-elevation band.\n\n"
            "• Level: {} (Elev: {:.2f} ft)\n"
            "• Elevation Band: {:.2f} ft to {:.2f} ft\n\n"
            "Try expanding the Min/Max Z offsets or verifying the Reference Level.".format(
                level_item.Name, lvl_elev, world_min_z, world_max_z
            ),
            title="No Points Found"
        )
        return False

    footprint_xy = []
    sampled_points = build_progressive_ground_terrain(
        grid_clusters,
        col_count,
        row_count,
        min_x,
        min_y,
        resolution,
        strategy_idx=strategy_idx,
        filter_noise=filter_noise,
        max_slope=max_slope,
        elevation_tolerance=elevation_tolerance,
        obstacle_mode=obstacle_mode,
        smoothing_mode=smoothing_mode,
        footprint_out=footprint_xy
    )

    if not sampled_points or sampled_points.Count < 3:
        forms.alert("At least 3 valid ground points are required.", title="Insufficient Points")
        return False

    # Create mode + detail areas: swap coarse points inside each box for fine-resolution points
    detail_added = 0
    detail_removed = 0
    if use_detail and is_create:
        try:
            detail_sets = []
            for box in detail_areas:
                detail_sets.append(
                    extract_detail_area_points(pc, transform, box, detail_resolution, resolution,
                                               world_min_z, world_max_z, terrain_kwargs)
                )
        except SamplingCancelled:
            forms.alert("Operation cancelled by user.", title="Cancelled")
            return False
        sampled_points, detail_removed, detail_added = merge_detail_points(
            sampled_points, detail_sets, detail_areas, detail_resolution
        )
        # Fine points near the site edge must stay inside the Toposolid boundary hull
        for pts in detail_sets:
            for p in pts:
                footprint_xy.append((p.X, p.Y))

    if sampled_points.Count > 25000:
        proceed = forms.alert(
            "Extracted {} ground points.\n\nProceed with generating Toposolid?".format(sampled_points.Count),
            title="High Point Count Warning",
            yes=True,
            no=True
        )
        if not proceed:
            return False

    all_z = [pt.Z for pt in sampled_points]
    z_min = min(all_z)
    z_max = max(all_z)
    relief = z_max - z_min

    required_thickness = max(
        15.0,
        (z_max - z_min) + 10.0,
        (lvl_elev - z_min) + 10.0
    )

    if is_create:
        ts_type = doc.GetElement(toposolid_type_item.Id)

        with Transaction(doc, "Generate Toposolid from Point Cloud") as t:
            t.Start()
            fail_opts = t.GetFailureHandlingOptions()
            fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
            t.SetFailureHandlingOptions(fail_opts)

            if ts_type:
                ensure_toposolid_type_thickness(ts_type, required_thickness)

            pts_x = [pt.X for pt in sampled_points]
            pts_y = [pt.Y for pt in sampled_points]
            bound_min_x = min(pts_x)
            bound_max_x = max(pts_x)
            bound_min_y = min(pts_y)
            bound_max_y = max(pts_y)

            # Build the 2D Convex Hull from the FULL scanned footprint (not just confirmed ground),
            # so corners occupied by trees/roads/objects are still inside the Toposolid boundary
            hull_src = [XYZ(fx, fy, 0.0) for fx, fy in footprint_xy] if footprint_xy else list(sampled_points)
            hull_pts = _compute_2d_convex_hull(hull_src, min_edge_len=max(2.0, resolution * 0.5))

            # Every hull vertex needs a ground elevation: reuse an existing point or extrapolate
            # (inverse-distance weighting from the 6 nearest confirmed ground points)
            if len(hull_pts) >= 3:
                ground_xyz = [(p.X, p.Y, p.Z) for p in sampled_points]
                resolved_hull = []
                for hp in hull_pts:
                    hp_x = hp.X
                    hp_y = hp.Y
                    # Avoid closure variables in lambda to prevent IronPython DLR UnboundNameException
                    dist_tuples = []
                    for gx, gy, gz in ground_xyz:
                        d2 = (gx - hp_x) * (gx - hp_x) + (gy - hp_y) * (gy - hp_y)
                        dist_tuples.append((d2, gx, gy, gz))
                    dist_tuples.sort(key=lambda t_entry: t_entry[0])
                    nearest = dist_tuples[:6]
                    d0 = nearest[0][0]
                    if d0 < 0.01:
                        hz = nearest[0][3]
                    else:
                        weight_sum = 0.0
                        z_weighted = 0.0
                        for d2, gx, gy, gz in nearest:
                            w = 1.0 / max(1e-6, d2)
                            weight_sum += w
                            z_weighted += w * gz
                        hz = z_weighted / weight_sum if weight_sum > 0.0 else nearest[0][3]
                        sampled_points.Add(XYZ(hp_x, hp_y, hz))
                    resolved_hull.append(XYZ(hp_x, hp_y, hz))
                hull_pts = resolved_hull

            profile_loops = List[CurveLoop]()
            if len(hull_pts) >= 3:
                avg_hull_z = sum(p.Z for p in hull_pts) / float(len(hull_pts))
                boundary_loop = CurveLoop()
                n_h = len(hull_pts)
                for i in range(n_h):
                    pa = hull_pts[i]
                    pb = hull_pts[(i + 1) % n_h]
                    boundary_loop.Append(Line.CreateBound(XYZ(pa.X, pa.Y, avg_hull_z), XYZ(pb.X, pb.Y, avg_hull_z)))
                profile_loops.Add(boundary_loop)
            else:
                pts_list = list(sampled_points)
                def get_corner_z(tx, ty, pts_ref, default_z):
                    best_d = float("inf")
                    best_z = default_z
                    for pt in pts_ref:
                        d = (pt.X - tx)**2 + (pt.Y - ty)**2
                        if d < best_d:
                            best_d = d
                            best_z = pt.Z
                    return best_z

                z_c1 = get_corner_z(bound_min_x, bound_min_y, pts_list, z_min)
                z_c2 = get_corner_z(bound_max_x, bound_min_y, pts_list, z_min)
                z_c3 = get_corner_z(bound_max_x, bound_max_y, pts_list, z_min)
                z_c4 = get_corner_z(bound_min_x, bound_max_y, pts_list, z_min)
                avg_corner_z = (z_c1 + z_c2 + z_c3 + z_c4) / 4.0

                p1 = XYZ(bound_min_x, bound_min_y, avg_corner_z)
                p2 = XYZ(bound_max_x, bound_min_y, avg_corner_z)
                p3 = XYZ(bound_max_x, bound_max_y, avg_corner_z)
                p4 = XYZ(bound_min_x, bound_max_y, avg_corner_z)

                boundary_loop = CurveLoop()
                boundary_loop.Append(Line.CreateBound(p1, p2))
                boundary_loop.Append(Line.CreateBound(p2, p3))
                boundary_loop.Append(Line.CreateBound(p3, p4))
                boundary_loop.Append(Line.CreateBound(p4, p1))
                profile_loops.Add(boundary_loop)

            toposolid = Toposolid.Create(
                doc,
                profile_loops,
                sampled_points,
                toposolid_type_item.Id,
                level_item.Id
            )

            status = t.Commit()

        if status != TransactionStatus.Committed:
            forms.alert("Toposolid creation could not be committed by Revit.", title="Creation Failed")
            return False

        forms.toast(
            "New Toposolid created successfully!\n"
            "• Ground Points: {}\n"
            "• Footprint: {:.1f} x {:.1f} ft\n"
            "• Relief Span: {:.1f} ft{}".format(
                sampled_points.Count,
                (bound_max_x - bound_min_x),
                (bound_max_y - bound_min_y),
                relief,
                "\n• Detail Areas: {} (+{} fine / -{} coarse pts)".format(
                    len(detail_areas), detail_added, detail_removed
                ) if use_detail else ""
            ),
            title="Toposolid Created"
        )
    else:
        target_ts = doc.GetElement(existing_toposolid_item.Id)
        if not target_ts:
            forms.alert("The selected Toposolid is no longer valid in the document.", title="Update Failed")
            return False

        ts_type = doc.GetElement(target_ts.GetTypeId())

        with Transaction(doc, "Add Points to Toposolid from Point Cloud") as t:
            t.Start()
            fail_opts = t.GetFailureHandlingOptions()
            fail_opts.SetFailuresPreprocessor(SilentErrorPreprocessor())
            t.SetFailureHandlingOptions(fail_opts)

            if ts_type:
                ensure_toposolid_type_thickness(ts_type, required_thickness)

            shape_editor = target_ts.GetSlabShapeEditor()
            if not shape_editor:
                t.RollBack()
                return False

            if not shape_editor.IsEnabled:
                shape_editor.Enable()

            try:
                shape_editor.AddPoints(sampled_points)
            except Exception:
                for pt in sampled_points:
                    try:
                        shape_editor.AddPoint(pt)
                    except Exception:
                        pass

            status = t.Commit()

        if status != TransactionStatus.Committed:
            forms.alert("Updating Toposolid could not be committed by Revit.", title="Update Failed")
            return False

        forms.toast(
            "Existing Toposolid updated successfully!\n• Points Added: {}".format(sampled_points.Count),
            title="Toposolid Updated"
        )

    return True


if __name__ == "__main__":
    pc_element = get_point_cloud()
    if pc_element:
        w = PCToToposolidWindow("ui.xaml", pc_element)
        w.ShowDialog()