# Changelog - Cut and Fill Tool

## [v2.1] - 2025-01-02
### Added
- **Interactive Report:**
  - **Cost Estimation:** Added real-time cost calculator. Users can input Unit Cost and Basis (Truck/CY) directly in the browser.
  - **Logistics Settings:** Added controls to adjust Truck Capacity, Swell Factor, and Compaction Factor. The report dynamically recalculates Net Trucks and updates visual balance bars based on these inputs.

## [v2.0] - 2025-01-02
### Changed
- **Workflow:** Removed the pyRevit output window entirely. The tool now generates the report directly to the user's **Downloads** folder and opens it in the default system browser.
- **Cleanup:** Removed redundant imports (`tempfile`, `clr`, `System`) and legacy pyRevit chart logic.

## [v1.9] - 2025-01-02
### Changed
- **Auto-Open:** Report now automatically opens in the default system browser upon generation. This provides access to full print driver options (Margins, Scale, Shrink to Fit) which are missing in the embedded viewer.
- **CSS:** Removed `@page` size and margin constraints to allow user overrides in the print dialog.
- **Link Fix:** Updated file path generation to use `System.Uri` for robust handling of spaces and special characters in the "Open in Browser" link.

## [v1.8] - 2025-01-02
### Fixed
- Fixed "Open in Browser" link functionality by converting the local file path to a valid `file:///` URI, ensuring it opens correctly in default browsers.

## [v1.7] - 2025-01-02
### Fixed
- Fixed `NameError: global name 'os' is not defined` by importing the missing `os` module.

## [v1.6] - 2025-01-02
### Added
- Added "Open in Browser" button which generates a standalone HTML file (with embedded Chart.js) and opens it in the default system browser. This resolves "Print Preview" limitations in the embedded viewer.

## [v1.5] - 2025-01-02
### Fixed
- Fixed chart scaling issue in print mode where the chart would appear disproportionately large or overflow the page width.

## [v1.4] - 2025-01-02
### Fixed
- Fixed `AttributeError` when creating charts by calling `make_line_chart` on the output instance (`out`) instead of the module (`output`).

## [v1.3] - 2025-01-02
### Changed
- Refactored output window initialization to use `pyrevit.output.get_output()` directly, ensuring better compatibility with print engines.

## [v1.2] - 2025-01-02
### Fixed
- Fixed critical issue where CSS styles were ignored during printing due to UTF-8 BOM encoding characters invalidating the style tag.

## [v1.1] - 2025-01-02
### Added
- Added a "Print Report" button directly to the report footer to trigger the browser print dialog.

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
- Added robust BOM handling for CSS files to prevent style injection failures.