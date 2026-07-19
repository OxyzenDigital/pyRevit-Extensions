# Manage Sheets: Transition Event Flow Specification

**Purpose:** This specification details the exact logical sequence of operations and validation events that occur when a user transitions from the **Project Setup** state to the **Sheet Reviewer** state in the Manage Sheets application. Agents modifying the UI, ViewModel, or backend logic must adhere to this sequence to maintain application integrity and data safety.

## 1. The Validation Gate (Access Control)
Before allowing the user to access the Sheet Reviewer, the application must perform a validation check to ensure a meaningful review context can be generated.
* **Trigger:** User attempts to select or transition to the "Sheet Reviewer" tab/view.
* **Action:** The system evaluates current user selections.
* **Gate Logic:** The intent is to **not proceed** if the data generation is invalid. It protects the backend engine, prevents an empty screen, and keeps the UI disabled until valid targets exist.

## 2. Securing the Blueprint (State Persistence)
The system captures the blueprint (the user's setup choices) to create a definitive list to compare against the actual Revit project.
* **Storage Intent:** Saving this to the Revit file (Extensible Storage) is a matter of convenience. The primary goal is that this blueprint can be exchanged, exported to JSON, or shared for future reference. 
* **UI Requirement:** The UI must support retrieving this blueprint (e.g., importing a previously saved JSON) for processing.

## 3. Calculating the "Ideal" Schema (Target Generation)
The application calculates the theoretical blueprint for the project's AIA standard sheet structure.
* **Action:** The system runs the generation algorithm in a **pure vacuum**. It calculates exactly what *should* exist completely independent of the current Revit model.
* **Requirement:** No Revit sheets are created or modified during this step.

## 4. Organizing the Workspace (Hierarchy Construction)
The flat list of "ideal" target sheets is converted into a structured tree map for presentation in the UI.
* **Grouping Rules:** While currently presented as a Discipline -> Content Group hierarchy, the objective is simply to present AIA defined sheets. Ultimately, how the user groups them in the UI should be their choice.

## 5. The Reality Check & Temporary Workbench (STRICT RULE)
The system prepares the main review dashboard (the data grid) as a **temporary workbench** to compare the "Ideal" list against the "Actual" Revit environment.
* **Primary Goal:** Find the deviation between the ideal list and the actual sheets, and mend the names and numbers (via fuzzy matching) so existing sheets can be updated to match the ideal set.
* **Sheet Collection Targeting (STRICT RULE):** 
  * The user is expected to select a **Sheet Collection** if it is present in the Revit model and use that specific data for the workbench.
  * If no Sheet Collection exists in the model, then all sheets in the root may be used.
  * If the Revit model is completely empty, then the pure ideal list is presented to the datagrid.

## 6. Validation and Execution (Action Station)
Once the user has worked on the temporary workbench, the system must validate the proposed changes before they can be pushed back into Revit.
* **Action:** Contextual action panels containing execution commands are revealed.
* **Validation Rules:** Before any push to Revit, the system must validate against Revit's strict acceptance rules:
  * No duplicate sheet numbers within the same Sheet Collection.
  * No unacceptable or illegal characters in sheet names/numbers.
* **Execution:** Once validated, the user is allowed to take action and push the workbench changes back into the Revit model.

---
**Agent Directive:** When refactoring, adding features, or changing state management logic, you **MUST** ensure that this 6-step logical flow is preserved. Pay special attention to the Strict Rule in Step 5 regarding Sheet Collections and the Validation requirements in Step 6.
