# Manage Sheets - Product Requirements & Workflow

## Overview
The "Manage Sheets" tool is a PyRevit extension designed to enforce architectural naming conventions, automate sheet generation, and manage view placements across complex Revit projects. It acts as a strict "workbench" where the user can compare the native Revit model against an external CAD Standard Schema, reconcile differences, and batch-push updates.

## Working Method & Philosophy
1. **Revit as the Read-Only Ground Truth:** Upon launch, the tool reads all Sheets and Views from the active Revit document. This state is cached in memory as a pristine baseline. The tool does not write to Revit during the user's interactive session.
2. **The Editor as a Sandbox:** The WPF UI acts as a sandbox. Users can map sheets, resolve naming conflicts, add missing views, and configure parameters without affecting the live Revit model.
3. **Strict Schema Compliance:** The tool expects sheets to conform to a specific organizational hierarchy (e.g., Collection -> Discipline -> Content Group -> Series). Non-conforming sheets are flagged for user reconciliation.
4. **Intelligent Auto-Tracking:** When a user modifies a sheet in the sandbox (e.g., renames it, or adds a view), the tool automatically transitions the sheet's status from `MATCHED` to `UPDATE` and queues it for synchronization.
5. **Batch Transaction:** Once the user is satisfied with the sandbox state, they click "Push To Revit". The tool compares the sandbox against the baseline, opens a single Revit `TransactionGroup`, and applies all changes (creations, parameter updates, view placements).

## End-to-End User Flow
1. **Parameter Pre-flight Check:** 
   - Upon clicking the tool, `script.py` checks if required Project Parameters (e.g., "Sheet Series", "Discipline") exist.
   - If missing, it prompts the user to auto-generate them using a temporary Shared Parameters file.
2. **Data Ingestion & Mapping:**
   - The tool loads all native sheets and groups them by `Sheet Collection`.
   - It cross-references them against the target schema defined in the project settings.
   - Sheets are categorized into matching, missing, or mismatched states.
3. **User Reconciliation (WPF UI):**
   - The user interacts with a split-pane WPF window.
   - **Left Pane:** Hierarchical TreeView showing sheet collections and disciplines.
   - **Right Pane:** DataGrid showing sheet details, matching statuses (🟢, 🟡, 🔴), and editable fields.
4. **Push to Revit:**
   - The user selects which sheets to sync.
   - The tool executes the changes. If a native sheet blocks a schema requirement, it is temporarily "auto-parked" (renamed with a suffix) to avoid fatal collisions.
   - A final summary report (Markdown) is printed to the PyRevit output window.
