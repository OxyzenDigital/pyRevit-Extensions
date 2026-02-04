# Changelog - Quantity & Measures

## [v0.2] - 2025-02-04
### Added
- **Export to CSV:** Added functionality to export quantified data to CSV.
  - Supports exporting checked items or currently selected node.
  - Generates hierarchical filenames based on selection (e.g., `Width-Doors-Single_Flush`).
- **Unit Handling:** Improved unit conversion and label display.
  - Added support for Revit 2022+ `ForgeTypeId`.
  - Implemented fallback parsing for older versions.
  - Ensures correct unit labels (SF, CF, ft, m, etc.) are displayed.

### Changed
- **Code Structure:** Refactored code into modular files (`revit_utils.py`, `data_model.py`) for better maintainability.
- **UI Logic:**
  - Fixed checkbox recursion to correctly select/deselect child nodes.
  - Updated button states (Visualize, Export) to react to checkbox changes.
  - Improved Dark Theme detection and resource application.