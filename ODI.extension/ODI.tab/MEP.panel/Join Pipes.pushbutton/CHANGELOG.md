# Changelog - Join Pipes Tool

## [v0.5.0] - 2025-12-29
### Added
- **MVVM Architecture:** Rebuilt the tool using a robust Model-View-ViewModel pattern similar to the Grading tool.
  - `data_model.py`: Manages application state (`AppState`) and solution data (`JoinSolution`).
  - `logic.py`: Pure Python geometric solver for calculating intersection points and rolling offsets.
  - `revit_service.py`: Handles all Revit API interactions (Selection, Transactions, DirectShapes).
  - `script.py`: Implements a "Modal Loop" to manage window state and Revit transactions safely.
- **Robust Selection:**
  - Implemented `PipeSelectionFilter` that supports `OST_PipeCurves`, `OST_DuctCurves`, `OST_Conduit`, `OST_CableTray`, and importantly **Fabrication Parts**.
  - Added safe `ElementId` handling (`get_id_val`) to support both Revit 2023 (and older) and Revit 2024+.
- **Geometric Solving:**
  - Implemented a vector-math based solver in `logic.py` to find the closest points between two infinite lines.
  - Detects **Intersections** (coplanar) vs **Skew Lines** (rolling offset).
- **Preview Visualization:**
  - Added transient `DirectShape` visualization (using `OST_PipeCurves` category) to show the proposed connection path before committing.
  - Implemented auto-cleanup of preview elements.
- **Transaction Handling:**
  - Implemented `commit_solution` to physically move pipe endpoints and create fittings.
  - Added `doc.Regenerate()` calls to ensure Connectors are updated before fitting creation.

### Fixed
- **Revit Crashes:** Resolved "Modifiable Document" and "Transaction" crashes by strictly separating UI logic (window open) from Transaction logic (window closed).
- **Selection Issues:** Fixed `ISelectionFilter` rejecting elements due to incorrect `BuiltInCategory` mapping and `ElementId.IntegerValue` usage in 2024+.
- **Fitting Creation:** Fixed silent failures in `NewElbowFitting` by ensuring geometry regeneration and correct connector matching.

---

## Future Development Pointers (Next Steps)

### 1. Multi-Solution Support
- **Current State:** The solver currently returns only one "best geometry" solution (Direct Connect).
- **Todo:** Implement alternative routing strategies in `logic.py`:
  - **45-Degree Solutions:** Calculate paths using two 45-degree elbows instead of 90s.
  - **90-Degree Routing:** For skew lines, offer a "Square" route (Out -> Up/Down -> In) instead of a direct diagonal rolling offset.
  - **Smart Routing:** Check for collisions along the proposed path.

### 2. Parallel Pipe Handling
- **Current State:** Parallel pipes return a "Parallel Offset" solution which is currently invalid/unimplemented.
- **Todo:** Implement logic to connect parallel pipes (e.g., S-curve or 90-90 offset).

### 3. Fittings for Rolling Offsets
- **Current State:** The rolling offset logic creates the intermediate pipe but the elbow creation code is wrapped in a generic `try/catch`.
- **Todo:** Robustify the connector matching for the *newly created* intermediate pipe. Ensure the correct ends are found even if the pipe creation flips the start/end points.

### 4. UI Refinement
- **Todo:** Add a "Diameter" display to the UI so users verify they are connecting same-size pipes.
- **Todo:** Add visual feedback (color coding) for "Valid" vs "Invalid" solutions in the preview.

### 5. Settings Persistence
- **Todo:** Save user preferences (Allow Rolling, Allow Vertical) to `settings.json` so they persist between sessions.
