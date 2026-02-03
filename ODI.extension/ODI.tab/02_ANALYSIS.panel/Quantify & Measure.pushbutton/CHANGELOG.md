# Changelog - Quantity & Measures

## [v0.1]
### Initial Release
- **Core Functionality:**
  - Scans visible elements in the active view.
  - Aggregates measurable quantities (Area, Volume, Length, etc.) by Category and Type.
  - Displays data in a hierarchical Tree View (Measurement -> Category -> Type).
- **UI/UX:**
  - **Dashboard:** Shows aggregated totals for selected nodes.
  - **DataGrid:** Lists individual instances with ID, Name, and Value when a Type is selected.
  - **Visualization:**
    - **Highlight:** Selected elements are highlighted in Orange (Instances) or Blue (Types) in the active view.
    - **Auto-Zoom:** Automatically zooms to selected elements.
    - **Colorize:** Applies distinct colors to checked categories/types for visual analysis.
  - **Theme Support:** Fully supports Revit Dark and Light themes.