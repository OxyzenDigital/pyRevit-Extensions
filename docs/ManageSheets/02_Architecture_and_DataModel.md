# Manage Sheets - Architecture & Data Model

## Pattern: Strict MVVM (Model-View-ViewModel)
The UI architecture strictly separates the layout (`ui.xaml`), the reactive state (`data_model.py`), and the core Revit execution logic (`script.py` / `panel.py`). 

### 1. The View (`ui.xaml`)
- **Responsibility:** Purely aesthetic layout and data binding definitions. No business logic resides here.
- **Components:** Uses WPF DataGrids, TreeViews, and CheckBoxes. 
- **Styling:** Adheres to a centralized `DynamicResource` semantic brush dictionary (defined in `AGENTS.md`) to dynamically support Revit Light/Dark modes.
- **Rule:** Never hardcode layout colors; always use the Semantic Brush Dictionary.

### 2. The ViewModel (`data_model.py`)
- **Responsibility:** Manages the sandbox state, cascading updates, and triggers validation logic.
- **Structure:**
  - `ViewModelBase`: Implements `INotifyPropertyChanged` for reactive WPF updates.
  - `MainViewModel`: The root context. Holds the `Sheets` collection and `IsPushEnabled` logic.
  - `SheetViewModel`: Represents an individual sheet. Handles `MatchStatus` transitions (e.g., `MATCHED` to `UPDATE`).
- **Auto-Check Logic:** If a `SheetViewModel` detects a modification (via `update_action`), it automatically flips `IsChecked = True` and calls its `validation_callback` to enable the Push button on the `MainViewModel`.
- **Validation:** Provides properties that XAML reads to change colors (e.g., green for matched, red for missing).

### 3. The Controller (`script.py` & `panel.py`)
- **Responsibility:** Interacts with the Revit API, orchestrates startup/shutdown, and executes the final Push transaction.
- **`script.py`:** The entry point. Handles parameter injection and invokes the WPF window. Crucially, it manages the `Transaction` lifecycle and prevents PyRevit from rolling back changes by avoiding aggressive `sys.exit()` calls when a transaction has just been committed.
- **`panel.py`:** Contains the heavy-lifting logic for reading the Revit database (`FilteredElementCollector`), applying color overrides, and physically pushing sheet/view modifications to Revit.

## Safety & Regression Prevention
To prevent "silent UI regressions" (where a XAML binding asks for a property that does not exist in Python), this project enforces:
1. **Explicit Data-Flow Planning:** Any modification to XAML bindings or ViewModel properties requires cross-verification.
2. **Automated Validation:** The script `dev_tools/validate_bindings.py` must be run to parse `ui.xaml` against `data_model.py` and warn developers of unmatched bindings.
