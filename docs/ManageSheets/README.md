# Manage Sheets - Project Documentation Hub

Welcome to the documentation for the **Manage Sheets** PyRevit extension. 
This project is built using IronPython, WPF (XAML), and the Revit API to provide a comprehensive architectural sheet and view management sandbox.

## Purpose of this Repository Section
These specifications are meant to allow developers, agents, and maintainers to rapidly understand the underlying mechanics, API gotchas, and architectural decisions of the Manage Sheets tool.

## Available Specifications

| File | Description |
| :--- | :--- |
| `01_Product_Requirements.md` | Defines the end-to-end user flow, application expectations, and the "Sandbox vs. Revit" working philosophy. |
| `02_Architecture_and_DataModel.md` | Details the Strict MVVM separation, WPF Semantic Brush conventions, and XAML/Python binding validation rules. |
| `03_Revit_Integration.md` | Explains the specialized logic for API parameter injection, cross-version compatibility (Revit 2022 vs 2024), and PyRevit transaction management. |

## Developer Guidelines
- **Always read the existing specs before planning new features.**
- **Enforce UI consistency:** Use `dev_tools/validate_bindings.py` to ensure XAML bindings match the `data_model.py` `@property` schema.
- **Respect `AGENTS.md`:** Global development rules for the repository (including WPF quirks and the Semantic Brush Dictionary) are strictly documented in the workspace `.agents/AGENTS.md` file.
