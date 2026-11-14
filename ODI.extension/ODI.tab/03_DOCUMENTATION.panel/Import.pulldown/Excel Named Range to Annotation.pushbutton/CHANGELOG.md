# Changelog

## [1.1.1] - 2026-11-14
### Added
- **Source File Path Persistence & History**: Persists `LastFilePath` and an ordered, deduplicated list of `RecentFiles` under `__meta__` in both document `ExtensibleStorage` and local JSON.
- **Saved Sources Selector**: Added a quick ComboBox in the table configuration card to switch among project saved sources without having to browse files repeatedly.
- **Link Hints & Inferred Metadata**: Implemented `infer_hint_metadata` to infer likely source files and range metadata for unlinked tables based on naming heuristics and workbook usage history.
- **File Column & Rich Link Tooltips**: Added a dedicated `File` column in the Project Excel Tables DataGrid and multi-line descriptive tooltips showing recorded/hinted paths, source range, and status.
- **Contextual Link Hint Banner**: Added a status banner above the worksheet picker indicating link state, missing file alerts, and 1-click relinking feedback.
- **Smart Initial Directory Resolution**: Added `get_best_init_dir()` to navigate directly to the nearest existing folder on disk when browsing.

## [1.1.0] - 2026-11-13
### Added
- **MVVM Refactor & Split-Pane UI**: Decoupled UI state into reactive `data_model.py` and modern borderless split-pane `ui.xaml` layout adhering to `.agents/AGENTS.md`.
- **Project Excel Tables Management**: Left pane DataGrid displaying all project linked tables with status detection (`Ready`, `Missing File`, `No Link`).
- **Batch Table Updates**: Support for multi-select batch updates with TransactionGroup encapsulation and summary alerts.
- **Interactive Instance Placement**: Added `place_table_instance()` using Revit's `PromptForFamilyInstancePlacement` for seamless placement directly into active views or sheets after table creation.
- **ExtensibleStorage Link Sharing**: Persists table link metadata directly into the Revit project document via `DataStorage` with local JSON fallback, allowing seamless multi-user collaboration without lost paths.
- **Worksheet-Wide Extraction**: Added `<Entire Worksheet>` option alongside named ranges and print areas to import whole sheets.
- **Revit 2026+ Template Resolution**: Added fallback search path across Revit application settings and program data versions `2020` through `2027`.

### Changed
- **Adaptive Dark Theme**: Full compliance with `.agents/AGENTS.md` semantic brush keys (`WindowBrush`, `ControlBrush`, `ToolbarBrush`, `CardBrush`, `AltRowBrush`, etc.) and WPF DataGrid selection highlight override keys.
- **DataFormatter Value Extraction**: Utilizes NPOI `DataFormatter` and formula evaluation for high-fidelity cell representation (currency, percentages, formatted dates) with fallback to raw types.

## [1.0.0] - 2026-01-05
### Added
- **Excel Import**: Import Excel ranges as Revit Annotation Families.
- **Formatting Support**: Preserves Fonts, Borders, Background Fills, and Column/Row dimensions.
- **Update Mechanism**: One-click update to refresh data from the source Excel file.
- **Settings Persistence**: Settings are stored in `%APPDATA%\ODI_ExcelTable\excel_table_map.json` to ensure persistence across sessions and Revit restarts.

### Fixed
- **Color Accuracy**: Improved RGB and Theme Color extraction logic for better fidelity (fixed Blue/Gray tint issues).
- **Selection Handling**: Fixed "Silent Rejection" bugs by ensuring elements are deselected before family updates.
- **Stability**: Enhanced transaction handling during family loading and updating.
- **Invisible Lines**: Improved detection of invisible line styles for cleaner table borders.
