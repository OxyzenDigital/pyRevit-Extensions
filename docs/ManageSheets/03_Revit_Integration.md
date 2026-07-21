# Manage Sheets - Revit Integration & Execution API

## Parameter Injection Lifecycle
The tool heavily relies on custom parameters (e.g., `Sheet Series`, `Discipline`) being associated with the `OST_Sheets` category in the active Revit document. 

Because the Revit API does not expose a direct `doc.ProjectParameters.Add()` method for Project Documents, the tool employs a specialized injection routine (`ensure_sheet_parameter` in `panel.py`):
1. **Validation:** It checks if the parameter already exists on Sheets. If not, it checks if an `ExternalDefinition` already exists in the project (perhaps bound to Views).
2. **Generation:** If missing entirely, it creates a temporary, correctly formatted `.txt` Shared Parameters file in the OS temp directory, opens it via `app.OpenSharedParameterFile()`, and generates an `ExternalDefinition`.
3. **Binding:** It uses `doc.ParameterBindings.Insert()` or `ReInsert()` to map the definition to the `OST_Sheets` category.
4. **Version Safety:** The API signature for `ParameterBindings.Insert` changed fundamentally between Revit 2023 and Revit 2024 (moving from `BuiltInParameterGroup` to `ForgeTypeId`). The `ensure_sheet_parameter` function uses explicit `import Autodesk.Revit.DB as DB` and stacked `try/except` blocks to guarantee cross-version compatibility without crashing IronPython.

## The Push Transaction (`_sync_action`)
When the user clicks "Push To Revit", a highly defensive batch execution cycle occurs:
1. **Baseline Isolation:** A `TransactionGroup` is opened.
2. **Auto-Parking:** Native elements that have a unique name or number collision with the incoming schema (but are not part of the current execution batch) are identified and renamed with a temporary suffix (e.g., `_TEMP`). This "auto-parking" ensures the API does not throw fatal collision exceptions during processing.
3. **Execution:** The script iterates through the `Sheets` collection in the ViewModel. It executes renames, creates new sheets, assigns views to viewports, and updates parameter values.
4. **PyRevit Transaction Gotcha:** The script must exit gracefully (no aggressive `sys.exit(0)` calls immediately after a commit). If PyRevit detects a forced script abort, it flags the external command as failed, which causes Revit's internal database monitor to immediately roll back the transaction group to protect the model. The tool avoids this by cleanly unwinding the call stack.
