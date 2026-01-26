# Changelog - Keynote File Editor

## [v1.2] - 2025-01-02
### Added
- **Undo Functionality:** Added a global Undo button to revert destructive actions (Drag & Drop, Renumber, Delete, Text Edits).
- **Recent Files Management:** Added a "Manage" button (gear icon) to the recent files dropdown, allowing users to view and remove specific entries from history.
- **Tree State Persistence:** Expanded/Collapsed state of tree nodes is now preserved across refreshes and drag-and-drop operations.

### Changed
- **File Export:** Saved files now default to the original filename with a timestamp suffix (e.g., `Keynotes_20250102_123000.txt`) instead of generic `Revised_Database.txt`.
- **Change Case:** Default scope for "Change Case" is now "Selected Section & Children".
- **Drag & Drop:**
  - Files dragged onto specific tree rows are now explicitly rejected to prevent accidental loading.
  - **Unformatted Text:** Loading unformatted text (single column) is now restricted to Reference panes only. Primary Editor ignores non-tab-delimited lines to maintain data integrity.

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