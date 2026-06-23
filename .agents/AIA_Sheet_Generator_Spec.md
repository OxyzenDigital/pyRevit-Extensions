AIA Sheet Naming Generator Specification for Antigravity 2.0

This Markdown file is the authoritative specification for a pyRevit app built with Antigravity 2.0 to generate AIA UDS Level 1 compliant sheet numbers and names dynamically. It defines standards, defaults, user inputs, name derivation rules, grid and enlargement logic, validation rules, Revit integration notes, and example outputs. Use this as the single source of truth for implementation.

Naming Convention Summary

Format: D-NNN[Suffix]

Example: A-101, A-101A

Discipline codes: A, C, S, M, E, P, FP, T, L, V

Series digit meanings (first digit of NNN):

0 = General

1 = Plans

2 = Elevations

3 = Sections

5 = Details

6 = Schedules

7 = Diagrams and Risers

9 = 3D and Perspectives

Enlargement suffix: recommended single uppercase letter appended to parent sheet (e.g., A-101A). Alternative suffix styles supported by configuration.

Matchlines and quadrant index: parent plan must show matchlines referencing enlarged sheets; cover sheet or A-001 must include a quadrant index with extents and scales.

Hard coded Standards and Defaults

These constants are embedded in the implementation.

DISCIPLINE_CODES = {
  "CS": "COVER SHEET", "G": "GENERAL", "H": "HAZARDOUS MATERIALS",
  "V": "SURVEY", "B": "GEOTECHNICAL", "C": "CIVIL", "L": "LANDSCAPE",
  "S": "STRUCTURAL", "A": "ARCHITECTURAL", "I": "INTERIORS",
  "Q": "EQUIPMENT", "F": "FIRE PROTECTION", "P": "PLUMBING",
  "D": "PROCESS", "M": "MECHANICAL", "E": "ELECTRICAL",
  "W": "DISTRIBUTED ENERGY", "T": "TELECOMMUNICATIONS", "R": "RESOURCE",
  "X": "OTHER DISCIPLINES", "Z": "CONTRACTOR SHOP DRAWINGS",
  "O": "OPERATIONS", "AD": "ARCHITECTURAL DEMOLITION", "AF": "ARCHITECTURAL FINISHES",
  "AG": "ARCHITECTURAL GRAPHICS", "AI": "ARCHITECTURAL INTERIORS",
  "FA": "FIRE ALARM", "MH": "HVAC", "MP": "HVAC PIPING",
  "EL": "ELECTRICAL LIGHTING", "EP": "ELECTRICAL POWER",
  "RA": "EXISTING ARCHITECTURAL", "RS": "EXISTING STRUCTURAL",
  "RP": "EXISTING PLUMBING", "RM": "EXISTING MECHANICAL"
}

SERIES_MAP = {
  0:"General",1:"Plans",2:"Elevations",3:"Sections",
  5:"Details",6:"Schedules",7:"Diagrams",9:"ThreeD"
}

DEFAULT_SUFFIX_ORDER = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
DEFAULT_GRID = (2,2)  # rows, cols
MATCHLINE_LABEL_FORMAT = "See {sheet_number}"
SHEET_SCHEMA_COLUMNS = ["Template ID","Template Name","Sheet Number","Sheet Name","Sheet Collection","Sheet Discipline","Sheet Use"]
FILENAME_SAFE_CHARS = "A-Z a-z 0-9 - _ ."

User Inputs and Configuration

The UI must collect these inputs before generation. All are configurable per project.

levels ordered list, e.g., ["Level 1","Level 2","Roof"]

disciplines_to_include subset of DISCIPLINE_CODES

enlargement_grid tuple (rows, cols) default (2,2); user may choose any reasonable grid

sheet_collection_map optional mapping to override collection names

Name Derivation Rules

Base sheet number composition

discipline_code = DISCIPLINE_CODES[discipline_full]

series_digit = series (0,1,2,3,5,6,7,9)

sequence = NN (01..99)

sheet_number = f"{discipline_code}-{series_digit}{NN:02d}"Example: A-101 for Architectural plans sheet 01.

Sheet name template

Default template: "{Discipline Full} {Sheet Type} {Qualifier}"

Examples:

A-101 → Floor Plan Level 1

A-111 → Dimension Plan Level 1

M-131 → Mechanical Piping Plan Level 1

Apply naming_overrides before finalizing.

Enlargement grid and suffix logic

For a parent plan P and grid R x C:

Generate suffixes in top-left to bottom-right order (row-major). For 2×2 the recommended mapping is:

A = NW, B = NE, C = SW, D = SE

Compose enlarged sheet number: f"{P}{suffix}" e.g., A-101A

Sheet name: "Enlarged Plan {Level} {QuadrantLabel}" or use user override

Support arbitrary grid sizes; if R*C > 26 use double-letter suffixes (AA, AB) or decimal style per suffix_style.

Parent sheet must show matchline with label MATCHLINE — See A-101A and scale.

Multi building and phasing

Prefer to keep building and phase information in Sheet Name or Sheet Collection rather than altering Sheet Number.

If office policy requires building code in number, document the rule and keep it consistent; avoid breaking AIA format where possible.

Validation Rules

No duplicate Sheet Number across the project.

Suffixes must be unique per parent sheet.

Sequence numbers should be contiguous within a series when practical.

Sheet numbers must use only filename safe characters.

Provide preview and require user confirmation before writing to Revit.

Data Model and Core Functions

Use these function signatures and data model as the implementation scaffold.

# Data model
class SheetRow:
    template_id: str
    template_name: str
    sheet_number: str
    sheet_name: str
    sheet_collection: str
    sheet_discipline: str
    sheet_use: str

# Core utilities
def format_sheet_number(discipline_code: str, series_digit: int, seq: int) -> str:
    return f"{discipline_code}-{series_digit}{seq:02d}"

def generate_suffixes(rows: int, cols: int) -> List[str]:
    # returns list like ["A","B","C","D"] or ["A","B","C","D","E",...]
    pass

def generate_discipline_sheets(discipline_full: str, levels: List[str],
                               grid: Tuple[int,int]) -> List[SheetRow]:
    # returns list of SheetRow objects for that discipline
    pass

def validate_sheet_rows(sheet_rows: List[SheetRow]) -> List[str]:
    # returns list of validation error messages
    pass

Implementations should follow the rules in this spec and return a preview before committing changes.

Revit Integration Notes

Parameters to write:

Sheet Number → Revit Sheet Number

Sheet Name → Revit Sheet Name

Sheet Collection → custom shared parameter SheetCollection (create if missing)

Workflow:

Collect user inputs in Antigravity UI.

Generate sheet rows and run validation.

Present preview table with Sheet Number, Sheet Name, Sheet Collection, and conflicts highlighted.

On user confirmation, update existing sheets or create new sheets and write parameters.

Undo and safety: do not delete or overwrite sheets without explicit user confirmation. Provide a log of changes.

Example 2x2 and Arbitrary Grid Outputs

Example 2x2 for Level 1

A-101, Floor Plan Level 1
A-101A, Enlarged Plan Level 1 NW
A-101B, Enlarged Plan Level 1 NE
A-101C, Enlarged Plan Level 1 SW
A-101D, Enlarged Plan Level 1 SE

Example 3x3 for Level 1

A-101, Floor Plan Level 1
A-101A, Enlarged Plan Level 1 Cell 1
A-101B, Enlarged Plan Level 1 Cell 2
A-101C, Enlarged Plan Level 1 Cell 3
A-101D, Enlarged Plan Level 1 Cell 4
A-101E, Enlarged Plan Level 1 Cell 5
A-101F, Enlarged Plan Level 1 Cell 6
A-101G, Enlarged Plan Level 1 Cell 7
A-101H, Enlarged Plan Level 1 Cell 8
A-101I, Enlarged Plan Level 1 Cell 9

Suffix mapping order must be documented in the UI and on the cover sheet quadrant index.



User Interaction Summary

This specification instructs the Antigravity app to present the following user interactions and inputs:

Project Setup Dialog

Enter project_code optional

Define levels in order (e.g., Level 1, Level 2, Roof)

Select disciplines_to_include from the supported list

Choose enlargement_grid (rows, cols) dynamically; default is 2x2

Optionally provide sheet_collection_map

Preview and Validation

The app generates a preview list of sheet rows

Validation runs and highlights duplicates, suffix collisions, and overflow warnings

User reviews and can edit names or override sequence numbers

Confirmation and Execution

On confirmation, the app writes Sheet Number, Sheet Name, and Sheet Collection to Revit sheets

The app logs changes and provides a summary of created or updated sheets

Post Generation Guidance

The app instructs the user to place matchlines on parent plans and to include a quadrant index on the cover sheet

The app recommends keeping building and phase qualifiers in Sheet Name rather than Sheet Number

Implementation Recommendations and Best Practices

Keep Sheet Number canonical and stable; use Sheet Name for qualifiers.

Use single letter suffixes for enlargements to keep sheets adjacent to parent.

Avoid encoding phase or revision into sheet numbers; use Revit parameters.

Document the suffix mapping and include a quadrant index on the cover sheet.

Next Steps for Antigravity 2.0 Developer

Implement the scaffold functions and UI described above.

Add unit tests for duplicate detection, suffix overflow, and multi building scenarios.

Provide a sample project template and sample CSV export for QA.

Include a short help panel in the UI that explains the AIA series digits and suffix rules.

Short Project Use Case Summary

This MD file instructs Antigravity to generate a complete AIA UDS Level 1 sheet index for projects of any size. The user selects levels, disciplines, and an enlargement grid. The tool produces canonical sheet numbers and names, supports arbitrary grid sizes, validates results, previews changes, and writes to Revit. The workflow is designed for clarity, office standardization, and safe Revit integration.

End of specification