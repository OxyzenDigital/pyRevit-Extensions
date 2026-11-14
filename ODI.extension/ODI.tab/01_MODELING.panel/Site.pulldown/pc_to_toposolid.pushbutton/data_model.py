# -*- coding: utf-8 -*-
"""
ViewModel and Data Structures for Point Cloud to Toposolid Tool.
Implements MVVM reactive data flow and INotifyPropertyChanged.
"""

import clr
clr.AddReference("System")
clr.AddReference("System.Core")

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.ObjectModel import ObservableCollection


class ViewModelBase(INotifyPropertyChanged):
    """Base class for ViewModels implementing INotifyPropertyChanged."""
    def __init__(self):
        self._property_changed_handlers = []

    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)

    def OnPropertyChanged(self, property_name):
        args = PropertyChangedEventArgs(property_name)
        for handler in self._property_changed_handlers:
            handler(self, args)


class LevelItem(ViewModelBase):
    """Wraps a Revit Level element for display and binding."""
    def __init__(self, level_id, name, elevation):
        ViewModelBase.__init__(self)
        self._id = level_id
        self._name = name
        self._elevation = elevation

    @property
    def Id(self):
        return self._id

    @property
    def Name(self):
        return self._name

    @property
    def Elevation(self):
        return self._elevation

    @property
    def DisplayText(self):
        return "{} ({:+.2f} ft)".format(self._name, self._elevation)


class ToposolidTypeItem(ViewModelBase):
    """Wraps a Revit ToposolidType element for display and binding."""
    def __init__(self, type_id, name):
        ViewModelBase.__init__(self)
        self._id = type_id
        self._name = name

    @property
    def Id(self):
        return self._id

    @property
    def Name(self):
        return self._name


class ToposolidInstanceItem(ViewModelBase):
    """Wraps an existing Revit Toposolid element for updating."""
    def __init__(self, element_id, name, type_name):
        ViewModelBase.__init__(self)
        self._id = element_id
        self._name = name
        self._type_name = type_name

    @property
    def Id(self):
        return self._id

    @property
    def Name(self):
        return self._name

    @property
    def TypeName(self):
        return self._type_name

    @property
    def DisplayText(self):
        id_val = self._id.IntegerValue if hasattr(self._id, "IntegerValue") else self._id.Value
        return "{} [ID: {}] ({})".format(self._name, id_val, self._type_name)


class StrategyItem(ViewModelBase):
    """Represents a ground point extraction strategy."""
    def __init__(self, index, name, description):
        ViewModelBase.__init__(self)
        self._index = index
        self._name = name
        self._description = description

    @property
    def Index(self):
        return self._index

    @property
    def Name(self):
        return self._name

    @property
    def Description(self):
        return self._description


class BoundaryModeItem(ViewModelBase):
    """Represents an extraction boundary mode."""
    def __init__(self, index, name):
        ViewModelBase.__init__(self)
        self._index = index
        self._name = name

    @property
    def Index(self):
        return self._index

    @property
    def Name(self):
        return self._name


class PCToToposolidViewModel(ViewModelBase):
    """Reactive ViewModel for the PC to Toposolid dialog."""
    def __init__(self, level_items, type_items, existing_toposolids, pc_min_z=None, pc_max_z=None):
        ViewModelBase.__init__(self)
        self._levels = ObservableCollection[LevelItem](level_items)
        self._toposolid_types = ObservableCollection[ToposolidTypeItem](type_items)
        self._existing_toposolids = ObservableCollection[ToposolidInstanceItem](existing_toposolids)

        self._pc_min_z = pc_min_z
        self._pc_max_z = pc_max_z

        self._has_existing_toposolids = len(self._existing_toposolids) > 0
        self._selected_existing_toposolid = self._existing_toposolids[0] if self._has_existing_toposolids else None

        # Mode: 0 = Create New, 1 = Update Existing
        self._target_mode_index = 0

        # Strategy options
        self._strategies = ObservableCollection[StrategyItem]([
            StrategyItem(0, "Lowest Percentile (5th) - Bare-Earth", "Filters out vegetation, vehicles & debris"),
            StrategyItem(1, "Minimum Z (Absolute Lowest)", "Uses lowest recorded point in cell"),
            StrategyItem(2, "Average Z (Mean)", "Averages all point elevations in cell")
        ])

        # Boundary options
        self._boundary_modes = ObservableCollection[BoundaryModeItem]([
            BoundaryModeItem(0, "Full Point Cloud Bounding Box"),
            BoundaryModeItem(1, "Interactive PickBox (Drag crop area in view)")
        ])

        # Obstacle handling options
        self._obstacle_modes = ObservableCollection[str]([
            "Predict Continuous Ground Under Obstacles (Recommended)",
            "Trim / Exclude Outlier Points"
        ])
        self._selected_obstacle_mode_index = 0

        # Surface smoothing options (Civil 3D DEM denoising)
        self._smoothing_modes = ObservableCollection[str]([
            "Moderate (Civil Grade - Recommended)",
            "Light (Subtle Denoising)",
            "Strong (Smooth Parkland Contours)",
            "None (Raw Surface Points)"
        ])
        self._selected_smoothing_mode_index = 0

        # Progressive Ground Extraction & Filtering defaults (Wang et al., Keller et al., Morel et al.)
        self._filter_noise = True
        self._max_ground_slope = "0.35"
        self._ground_elevation_tolerance = "2.0"
        self._smrf_elevation_threshold = "2.0"
        self._smrf_max_window_radius = "60.0"
        self._smrf_slope_threshold = "0.15"
        self._max_ground_jump = "2.0"

        # Default level to "Level 01" or lowest elevation
        default_lvl = None
        for lvl in self._levels:
            if "01" in lvl.Name or "Level 1" in lvl.Name:
                default_lvl = lvl
                break
        if not default_lvl and len(self._levels) > 0:
            default_lvl = self._levels[0]

        self._selected_level = default_lvl
        self._selected_toposolid_type = self._toposolid_types[0] if len(self._toposolid_types) > 0 else None
        self._selected_strategy = self._strategies[0]
        self._selected_boundary_mode = self._boundary_modes[0]

        # Numeric input defaults automatically tailored to detected point cloud extents
        self._grid_resolution = "5.0"
        self._calculate_default_z_offsets()

        # Detail Areas (selective higher-resolution refinement)
        self._enable_detail_areas = False
        self._detail_resolution = "1.5"
        self._detail_areas = []  # list of (min_x, min_y, max_x, max_y) in world XY

        # Status and state
        if self._has_existing_toposolids:
            self._status_message = "Ready. Found {} existing Toposolid(s).".format(len(self._existing_toposolids))
        else:
            self._status_message = "Ready. No existing Toposolid found - will generate a new one."
        self._is_busy = False

    def _calculate_default_z_offsets(self):
        """Calculates Min/Max Z offsets that safely encompass the entire point cloud."""
        if self._pc_min_z is not None and self._pc_max_z is not None and self._selected_level:
            lvl_elev = self._selected_level.Elevation
            # Add 2.0 ft safety buffer so points at extents are never clipped
            min_off = round(self._pc_min_z - lvl_elev - 2.0, 1)
            max_off = round(self._pc_max_z - lvl_elev + 2.0, 1)
            self._min_z_offset = str(min_off)
            self._max_z_offset = str(max_off)
        else:
            self._min_z_offset = "-20.0"
            self._max_z_offset = "50.0"

    @property
    def PointCloudExtentsText(self):
        if self._pc_min_z is not None and self._pc_max_z is not None:
            relief = self._pc_max_z - self._pc_min_z
            return "Point Cloud Z Range: {:.1f} ft to {:.1f} ft (Total Relief: {:.1f} ft)".format(
                self._pc_min_z, self._pc_max_z, relief
            )
        return "Point Cloud Z Range: Auto-detecting..."

    @property
    def FilterNoise(self):
        return self._filter_noise

    @FilterNoise.setter
    def FilterNoise(self, value):
        self._filter_noise = value
        self.OnPropertyChanged("FilterNoise")

    @property
    def MaxGroundSlope(self):
        return self._max_ground_slope

    @MaxGroundSlope.setter
    def MaxGroundSlope(self, value):
        self._max_ground_slope = value
        self.OnPropertyChanged("MaxGroundSlope")

    @property
    def GroundElevationTolerance(self):
        return self._ground_elevation_tolerance

    @GroundElevationTolerance.setter
    def GroundElevationTolerance(self, value):
        self._ground_elevation_tolerance = value
        self._smrf_elevation_threshold = value
        self._max_ground_jump = value
        self.OnPropertyChanged("GroundElevationTolerance")
        self.OnPropertyChanged("SmrfElevationThreshold")
        self.OnPropertyChanged("MaxGroundJump")

    @property
    def SmrfElevationThreshold(self):
        return self._ground_elevation_tolerance

    @SmrfElevationThreshold.setter
    def SmrfElevationThreshold(self, value):
        self.GroundElevationTolerance = value

    @property
    def SmrfMaxWindowRadius(self):
        return self._smrf_max_window_radius

    @SmrfMaxWindowRadius.setter
    def SmrfMaxWindowRadius(self, value):
        self._smrf_max_window_radius = value
        self.OnPropertyChanged("SmrfMaxWindowRadius")

    @property
    def SmrfSlopeThreshold(self):
        return self._max_ground_slope

    @SmrfSlopeThreshold.setter
    def SmrfSlopeThreshold(self, value):
        self.MaxGroundSlope = value

    @property
    def MaxGroundJump(self):
        return self._ground_elevation_tolerance

    @MaxGroundJump.setter
    def MaxGroundJump(self, value):
        self.GroundElevationTolerance = value

    @property
    def ObstacleModes(self):
        return self._obstacle_modes

    @property
    def SelectedObstacleModeIndex(self):
        return self._selected_obstacle_mode_index

    @SelectedObstacleModeIndex.setter
    def SelectedObstacleModeIndex(self, value):
        self._selected_obstacle_mode_index = value
        self.OnPropertyChanged("SelectedObstacleModeIndex")

    @property
    def SmoothingModes(self):
        return self._smoothing_modes

    @property
    def SelectedSmoothingModeIndex(self):
        return self._selected_smoothing_mode_index

    @SelectedSmoothingModeIndex.setter
    def SelectedSmoothingModeIndex(self, value):
        self._selected_smoothing_mode_index = value
        self.OnPropertyChanged("SelectedSmoothingModeIndex")

    @property
    def HasExistingToposolids(self):
        return self._has_existing_toposolids

    @property
    def TargetModeIndex(self):
        return self._target_mode_index

    @TargetModeIndex.setter
    def TargetModeIndex(self, value):
        self._target_mode_index = value
        self.OnPropertyChanged("TargetModeIndex")
        self.OnPropertyChanged("IsCreateNew")
        self.OnPropertyChanged("IsUpdateExisting")
        self.OnPropertyChanged("ActionButtonText")

    @property
    def IsCreateNew(self):
        return self._target_mode_index == 0

    @property
    def IsUpdateExisting(self):
        return self._target_mode_index == 1

    @property
    def ActionButtonText(self):
        return "Generate New Toposolid" if self._target_mode_index == 0 else "Update Existing Toposolid"

    @property
    def ExistingToposolids(self):
        return self._existing_toposolids

    @property
    def SelectedExistingToposolid(self):
        return self._selected_existing_toposolid

    @SelectedExistingToposolid.setter
    def SelectedExistingToposolid(self, value):
        self._selected_existing_toposolid = value
        self.OnPropertyChanged("SelectedExistingToposolid")

    @property
    def Levels(self):
        return self._levels

    @property
    def SelectedLevel(self):
        return self._selected_level

    @SelectedLevel.setter
    def SelectedLevel(self, value):
        self._selected_level = value
        self.OnPropertyChanged("SelectedLevel")
        # Automatically update Z offsets to match new level
        self._calculate_default_z_offsets()
        self.OnPropertyChanged("MinZOffset")
        self.OnPropertyChanged("MaxZOffset")

    @property
    def ToposolidTypes(self):
        return self._toposolid_types

    @property
    def SelectedToposolidType(self):
        return self._selected_toposolid_type

    @SelectedToposolidType.setter
    def SelectedToposolidType(self, value):
        self._selected_toposolid_type = value
        self.OnPropertyChanged("SelectedToposolidType")

    @property
    def Strategies(self):
        return self._strategies

    @property
    def SelectedStrategy(self):
        return self._selected_strategy

    @SelectedStrategy.setter
    def SelectedStrategy(self, value):
        self._selected_strategy = value
        self.OnPropertyChanged("SelectedStrategy")

    @property
    def BoundaryModes(self):
        return self._boundary_modes

    @property
    def SelectedBoundaryMode(self):
        return self._selected_boundary_mode

    @SelectedBoundaryMode.setter
    def SelectedBoundaryMode(self, value):
        self._selected_boundary_mode = value
        self.OnPropertyChanged("SelectedBoundaryMode")

    @property
    def MinZOffset(self):
        return self._min_z_offset

    @MinZOffset.setter
    def MinZOffset(self, value):
        self._min_z_offset = value
        self.OnPropertyChanged("MinZOffset")

    @property
    def MaxZOffset(self):
        return self._max_z_offset

    @MaxZOffset.setter
    def MaxZOffset(self, value):
        self._max_z_offset = value
        self.OnPropertyChanged("MaxZOffset")

    @property
    def GridResolution(self):
        return self._grid_resolution

    @GridResolution.setter
    def GridResolution(self, value):
        self._grid_resolution = value
        self.OnPropertyChanged("GridResolution")

    # ------------------------------------------------------------------
    # Detail Areas (selective higher-resolution refinement)
    # ------------------------------------------------------------------
    @property
    def EnableDetailAreas(self):
        return self._enable_detail_areas

    @EnableDetailAreas.setter
    def EnableDetailAreas(self, value):
        self._enable_detail_areas = bool(value)
        self.OnPropertyChanged("EnableDetailAreas")

    @property
    def DetailResolution(self):
        return self._detail_resolution

    @DetailResolution.setter
    def DetailResolution(self, value):
        self._detail_resolution = value
        self.OnPropertyChanged("DetailResolution")

    @property
    def DetailAreas(self):
        return self._detail_areas

    @property
    def HasDetailAreas(self):
        return len(self._detail_areas) > 0

    @property
    def DetailAreasSummary(self):
        if not self._detail_areas:
            return "No detail areas selected"
        total_area = sum((b[2] - b[0]) * (b[3] - b[1]) for b in self._detail_areas)
        return "{} area(s) selected - {:,.0f} sq ft".format(len(self._detail_areas), total_area)

    def _notify_detail_areas(self):
        self.OnPropertyChanged("DetailAreas")
        self.OnPropertyChanged("HasDetailAreas")
        self.OnPropertyChanged("DetailAreasSummary")

    def add_detail_area(self, box):
        """Adds a (min_x, min_y, max_x, max_y) world-XY box."""
        self._detail_areas.append(tuple(box))
        self._notify_detail_areas()

    def clear_detail_areas(self):
        self._detail_areas = []
        self._notify_detail_areas()

    @property
    def StatusMessage(self):
        return self._status_message

    @StatusMessage.setter
    def StatusMessage(self, value):
        self._status_message = value
        self.OnPropertyChanged("StatusMessage")

    @property
    def IsBusy(self):
        return self._is_busy

    @IsBusy.setter
    def IsBusy(self, value):
        self._is_busy = value
        self.OnPropertyChanged("IsBusy")
