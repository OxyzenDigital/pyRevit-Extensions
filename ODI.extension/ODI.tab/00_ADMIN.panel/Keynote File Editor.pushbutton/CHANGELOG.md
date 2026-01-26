# Changelog - Keynote File Editor

## [v1.1.1] - 2025-01-02
### Fixed
- **Drag and Drop:** Fixed a critical bug where the Undo operation caused data type mismatches (String vs Number) for Pane IDs, breaking drag-and-drop functionality and causing incorrect merge behavior.

## [v1.1] - 2025-01-02
### Changed
- **Renumbering Logic:**
  - Added protection for CSI MasterFormat keys (e.g., `03 10 00`) to prevent accidental renumbering.
  - Implemented smart sequence detection (Numeric `01`, `1`, Alpha `A`, Alphanumeric `A1`) based on existing children.
  - Added confirmation dialog before processing.
- **Drag and Drop:**
  - **Visual Feedback:** Added distinct styling for drop zones:
    - **Sibling (Above/Below):** Orange gradient with solid line.
    - **Child (Inside):** Blue dashed outline with background tint.
  - **Logic:**
    - Internal moves now preserve the original key instead of prompting for a new one.
    - Cross-pane copies check for duplicates and prompt only if necessary.
    - Scroll position is preserved after drag operations.
- **UI/UX:**
  - Added version number `v1.1` to the toolbar.
  - Improved tooltip behavior and visual styling.