# ANTIGRAVITY AGENT INSTRUCTIONS: CORE DIRECTIVES

**CRITICAL:** You are interacting with an established codebase. You must prioritize the preservation of existing logic, strictly follow documented specifications, and never assume a blank slate.

## 1. THE "READ-FIRST" PROTOCOL
Before writing, modifying, or suggesting any code, you MUST:
1. Locate and read the core architectural specification for the current task (e.g., `aia_sheet_reconciliation_algorithm_v1.md`, `ARCHITECTURE.md`, or the active `.py` file).
2. Explicitly state in your response: "I have reviewed [Filename] and understand the current state."
3. Wait for confirmation or proceed ONLY if the objective is completely clear based on the existing code context.

## 2. THE "PATCH, DON'T REPLACE" RULE (ANTI-REWRITE BIAS)
1. **Never rewrite entire files** from scratch unless explicitly commanded by the user.
2. When modifying existing logic, provide **only the specific classes, functions, or lines** that are changing.
3. Show exactly where the new code should be inserted using clear `// BEFORE` and `// AFTER` comments, or standard diff formatting. 
4. If a function works, do not refactor it just to change its style.

## 3. REVIT API & PYREVIT CONSTRAINTS
1. **No Hallucinations:** The Revit API is rigid. Do not invent methods (e.g., assuming `db.guidegrid.create` exists). If you are unsure of a class property or method, search the official Revit API docs or ask the user to verify before writing the code.
2. **Transaction Safety:** Any modification to the Revit database MUST occur within a `Transaction` or `TransactionGroup`. Always account for RollBacks on failure.
3. **Data Types:** Pay strict attention to Revit internal data types (e.g., `ElementId` vs. `IntegerValue`, `BuiltInParameter` vs. `String`). 

## 4. LOCAL ENVIRONMENT & SECRETS
1. Rely on standard Python libraries (e.g., `requests`, `os`, `json`) where possible to avoid dependency bloat.
2. Never hardcode credentials, API keys, or directory paths. Always read from local `.env` files or project-relative directories.

## 5. EXECUTION & COMMUNICATION
1. Before executing multi-file changes or terminal commands, output a brief "Implementation Plan" for user approval.
2. Keep conversational filler to zero. Output the code, the location to place it, and any necessary API warnings.