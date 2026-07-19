# Manage Sheets: Nested Views Event Flow Specification

**Purpose:** This specification elaborates on the operations that occur inside the temporary workbench specifically regarding the **Nested Views** feature. It outlines the process of adding, defining, managing, and reconciling views within a targeted sheet slot before pushing changes to the Revit model.

## 1. The Nested View Context
Each Sheet slot within the temporary workbench must be capable of hosting a nested collection of Views.
* **Data Structure:** A sheet object (ViewModel) in the workbench must maintain an observable list (`Views`) of `ViewItemViewModel` instances.
* **UI Representation:** Users can access a dedicated nested list for any selected sheet, allowing them to inspect or define the views that reside on it.

## 2. Defining a Theoretical View (The "Add View" Action)
Users can manually define the views that *should* be placed on a sheet before any physical changes occur in Revit.
* **Action:** The user clicks the "+ Add View" button.
* **Intent:** Creates a theoretical blueprint for a view. No actual Revit Views are created or placed during this action.

## 3. View Validation, Attributes, and Exceptions
Each nested view must capture the core attributes required for Revit.
* **Core Attributes:**
  * **View Number (Detail Number):** The local identifier for the view on the sheet.
  * **View Name (Display Intent):** The name the user actually wants to see printed on the sheet (e.g., "Floor Plan").
  * **Plan Type (ViewFamilyType):** The Revit view type (e.g., Floor Plan, Section).
* **Exception Handling (Legends & Schedules):** View types that do not rely on standard Detail Numbers (like Legends) or can exist on multiple sheets must be identified by the system. The UI must **gray out (disable)** their number input fields and exclude them from auto-sequencing logic so they don't get in the user's way.
* **Local Validation Rule:** For standard views, the UI must ensure that no two views on the same sheet share the exact same View Number.

## 4. Automation and Utility Tools (UI Enhancements)
To accelerate workflow within the workbench, the UI must provide utility tools that act upon the nested views.
* **Auto-Sequencing & Renumbering:** The UI must include controls (e.g., drag-and-drop, up/down arrows, or an "Auto-Number" button) allowing the user to rapidly change the sequence of views. The detail numbers should automatically recalculate (1, 2, 3...) based on their order in the UI.
* **Case Formatting:** Provide bulk action tools to instantly change the text casing of selected view names (e.g., UPPERCASE, Title Case, lowercase) for quick standardization.
* **Semantic Name Matching:** Provide a smart tool that synchronizes view names with the parent Sheet's intent. For example, if the parent sheet is named "Enlarged Plans", applying this tool should automatically suffix or prefix the nested view names with "Enlarged Plan".

## 5. The Reality Check & Semantic Auditing
The workbench compares ideal views against actual placed views and audits them for logical consistency.
* **Action:** When a sheet is matched to an existing Revit Sheet, the system must harvest the actual views currently placed on it.
* **Goal:** Populate the `Views` collection with real, existing placed views alongside theoretical ones, allowing for direct comparison and mending.
* **Semantic Auditing (Mismatch Detection):** This is a critical validation step. The system must analyze the harvested Revit views against the theoretical intent of the parent sheet.
  * *Example:* If the Sheet slot is designated for "Level 1", but the harvested view's internal Revit properties (Associated Level or Name) indicate it belongs to "Level 2" or "Roof".
  * *UI Response:* The UI must NOT automatically move the view (to avoid overcomplicating the workbench). Instead, it must visually flag the mismatched view (e.g., a warning icon ⚠️, red/orange highlighting) and provide a tooltip explaining the discrepancy (e.g., "Warning: View Level does not match Sheet Intent").
* **Relocation Strategy:** To resolve a flagged mismatch, the user can use the "Remove" (X) button to queue the view for removal from the incorrect sheet during execution, allowing it to be properly placed later.

## 6. Execution and Placement (Push to Revit)
When the user finishes and initiates the push to Revit, the nested views must be processed.
* **Global Unique Naming Logic (The Revit Duplicate Name Fix):**
  * Revit enforces a strict rule: *No two views in the entire project can have the same View Name.* This is problematic because users frequently want multiple views named "Floor Plan" on different sheets.
  * **Resolution:** The engine must decouple the *Display Name* from the *System Name*.
  * When pushing to Revit, the engine must write the user's desired name to the View's **"Title on Sheet"** parameter.
  * The engine must automatically generate a globally unique backend **"View Name"** by concatenating the parent container data.
  * **Format:** `[Desired View Name] - [Sheet Collection Shorthand] - [Detail Number (Double Digits)] - [Sheet Number]` (e.g., `Floor Plan - PS - 01 - A101`). If no Sheet Collection exists, omit that shorthand.
* **Execution Logic:**
  * **Updates:** If a view exists but differs in the workbench, update its "Title on Sheet" and backend "View Name" according to the logic above.
  * **Creation/Placement:** Generate any new theoretical views using the specified `Plan Type`, apply the unique naming logic, and physically place them onto the sheet.

---
**Agent Directive:** When implementing the nested view data models, adhere strictly to this separation of concerns: The workbench is a theoretical space. Do not call Revit API creation or modification methods until the final execution phase described in Step 6. Ensure Legend views are specifically targeted for UI disabling and excluded from detail numbering validation. Always use the "Title on Sheet" workaround in Step 6 to prevent duplicate name API crashes.
