# Manage Sheets — UI Workflow & Interaction Spec

## Overview
This specification defines the interactive behaviors, validation feedback loops, and user flows for the **Manage Sheets** WPF tool. It establishes standard visual cues for data validation (specifically duplicate name collisions) and bulk operations.

---

## 1. Interactive States & Visual Cues

### 1.1 List Item Feedback
Each list item (both parent Sheets and child Views) dynamically responds to user interaction and validation status.

- **Normal State:** Neutral control backgrounds with standard text contrast.
- **Validation Error State:**
  - **Visual Treatment:** The offending list item is underlined with a bold alert color (e.g., Red `#EF4444`).
  - **Graying Out:** Input fields (TextBoxes) associated with the invalid item are visually "grayed out" by significantly lowering their opacity (e.g., 50%). This signals that the current input state is unacceptable and must be modified.
  - **Warning Text:** An italicized, color-matched warning sentence appears beneath the item explaining the issue (e.g., *"Name already exists in project"*).
- **Live Validation:** Validation occurs instantly on keystrokes (`UpdateSourceTrigger=PropertyChanged`). As soon as the user types a character that resolves the conflict, the alert underline, gray-out effect, and warning text disappear instantly.

### 1.2 Interactive Controls
- **Detail Numbers (Legends/Schedules):** These inputs are explicitly disabled (`IsEnabled=False`) to prevent editing, as their numbering is managed differently by Revit.
- **Bulk Action Buttons:** Hoverable text buttons (`UPPER`, `lower`, `Title`, `Sync Names`) appear in the nested view header to trigger mass operations.

---

## 2. Validation Gates

### 2.1 Live DataGrid Validation
- **Duplicate Sheet Numbers:** Checks across the entire grid for identical sheet numbers *within the same Sheet Collection*. If found, both rows are flagged.
- **Duplicate View Names (Grid):** Ensures two new views within the current session do not share the exact same name.
- **Duplicate View Names (Project):** Ensures new views do not conflict with existing views already present in the active Revit project.
- **Duplicate View Numbers (Sheet):** Prevents multiple views on the exact same sheet from holding the same Detail Number.
- **Semantic Mismatch:** A soft warning (non-blocking) if a view's Level name logically conflicts with the parent Sheet's intended name.

### 2.2 Execution Block
The main "Execute" button (`Btn_ApplyChanges`) is dynamically disabled (`IsEnabled=False`) as long as *any* checked row contains a validation error. It re-enables only when all conflicts are resolved.

---

## 3. Workflow Operations

### 3.1 Two-Phase Transaction (Collision Avoidance)
To prevent temporary naming collisions when swapping sheet numbers (e.g., renaming Sheet 1 -> 2, and 2 -> 1):
- **Phase 1 (Park):** All selected sheets are temporarily renamed with a unique hash suffix (e.g., `[SheetNumber]_temp_[Hash]`).
- **Phase 2 (Finalize):** All parked sheets are renamed to their final, user-desired target numbers.

### 3.2 View Naming Strategy
Revit requires all views to have a globally unique name. To allow users to see clean names on sheets without backend collisions:
- **Title on Sheet:** Set to the user's desired short name.
- **Backend View Name:** Automatically generated as `[Desired Name] - [Sheet Collection Shorthand] - [Detail Number] - [Sheet Number]`.
