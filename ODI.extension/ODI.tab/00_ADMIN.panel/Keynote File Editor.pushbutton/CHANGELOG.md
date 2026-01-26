# Changelog - Keynote File Editor

## [v1.1.0] - 2025-01-02
### Enhancements
- **Smart Drag & Drop:**
  - Added logic to detect CSI/MasterFormat numbering patterns.
  - Implemented auto-accept for moves that fit the target sequence or pattern.
  - Added step detection for renumbering suggestions.
- **UI Improvements:**
  - Updated selection highlight to red for better visibility.
  - Improved contrast for Root Nodes in Light/Dark themes.
  - Added "File Info Box" with a "Locate" button to copy filenames.
  - Added Toast notifications for user feedback.
  - Added "Close Pane" button for secondary panes.

## [v1.0.0] - 2025-01-02
### Initial Release
- **Core Functionality:**
  - Standalone HTML-based Keynote Editor.
  - Supports hierarchical editing of Revit Keynote files (tab-delimited).
  - Dual-pane interface for comparing and auditing against reference files.

### Features
- **File Management:**
  - Load/Save .txt files.
  - "Recent Files" list persisted in local storage for quick access.
  - Status label indicating the currently loaded filename.
- **Editing Tools:**
  - **Drag-and-Drop:** Re-parenting and ordering within the primary pane.
  - **Cross-Pane Copy:** Drag items from Reference to Primary pane with optional deep copy (children).
  - **CRUD:** "Add Sibling" and "Delete Item" (recursive) context actions.
  - **Renumbering:** Tool to auto-increment child keys for a selected parent.
  - **Find & Replace:** Bulk text replacement for descriptions.
- **Visualization & Navigation:**
  - Tree view with "Expand All" / "Collapse All" controls.
  - **Deep Search:** Filters tree nodes and auto-expands parents of matches.
  - **Coordination Audit:** Visual highlighting (faded green) for rows with matching descriptions across panes.
  - **Sync Scroll:** Clicking an item finds and centers matching descriptions in other panes.
  - **Highlighting:** Selected items are highlighted in gold across all panes.
- **UI/UX:**
  - **Dark Mode:** Fully supported theme with persistence.
  - **Layout:** Resizable split-pane layout.
  - **Tooltips:** Custom theme-aware tooltips for all interactive elements.
  - **Subtle Design:** Refined tree indentation and drag-and-drop visual cues.