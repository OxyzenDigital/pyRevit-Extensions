# Changelog - List Identicals

## [v1.1] - 2025-01-02
### Changed
- **Detection Method:** Switched from geometric comparison to scanning Revit's "Identical Instances" warnings. This ensures 100% alignment with Revit's internal duplicate detection and significantly improves performance on large models.

## [v1.0] - 2025-01-02
### Initial Release
- **Features:**
  - Geometric duplicate detection based on Location and Bounding Box.
  - Grouping by Category, Family, Type.
  - HTML Report with "Select" and "Purge" buttons.