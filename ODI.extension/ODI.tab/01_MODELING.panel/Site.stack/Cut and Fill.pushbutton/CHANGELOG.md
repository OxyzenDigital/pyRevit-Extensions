# Changelog - Cut and Fill Tool

## [v1.0] - 2025-01-02
### Initial Release
- **Core Calculation:**
  - Extracts Cut, Fill, and Total Volumes from Revit Toposolids.
  - Supports Design Options and Phasing.
  - Calculates Net Volume and Truck Logistics (Import/Export estimates).
- **Reporting:**
  - Generates a styled HTML report with visual balance bars.
  - Includes a "Raw Data Log" for auditing individual elements (ID, Type, Subdivision).
  - Added timestamp to report header.
- **UI/UX:**
  - High-contrast, large font layout for readability.
  - Optimized CSS for printing (embedded styles, forced background colors).
  - Removed horizontal scrollbars.

### Fixed
- Fixed printing issues where CSS styles and background colors were stripped by the browser.