

## Manage Sheets Tool Constraints
- **Revit 2025 Sheet Collections:** When querying or assigning native Sheet Collections in Revit 2025, strictly use ParameterTypeId.SheetCollection to get the parameter and avoid fuzzy fallback loops iterating through ViewSheetSet or Worksets. Extract the actual parameter via sheet.get_Parameter(ParameterTypeId.SheetCollection).
- **Fallback for older Revit versions:** Use a strict Project Parameter lookup via sheet.LookupParameter(" Sheet Collection\).
