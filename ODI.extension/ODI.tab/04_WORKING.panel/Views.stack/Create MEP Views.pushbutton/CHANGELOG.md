# Changelog

## [v1.0.0] - 2026-01-04
### Added
- **Initial Release:** Comprehensive "MEP Views" creation tool.
- **View Types:** Supports creation of HVAC, Pipes (Sanitary, Water, Gas), and 3D views.
- **Auto-Naming:** Automatically names views based on the selected level (e.g., `_Work Pipes - Level 1`).
- **3D Logic:** Smartly handles 3D view naming (Single global view for small projects, per-level for larger ones).
- **Filters:** Automatically creates and applies essential MEP filters (Cold/Hot Water, Gas, Sanitary, Vent, Storm, Hydronic).
- **Discipline Enforcement:** Enforces correct View Discipline (Mechanical/Plumbing) regardless of template selection.
- **Visibility Overrides:**
  - Isolates relevant categories (e.g., Pipes only in Pipe views).
  - Automatically turns on Insulation/Linings.
  - "Grays out" (Halftones) architectural context (Walls, Floors, etc.) when no template is applied.
- **Safety:** Implemented robust error handling, null checks, and Revit version compatibility (2024+ ElementId fixes).
