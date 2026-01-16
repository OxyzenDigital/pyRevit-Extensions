# Changelog

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