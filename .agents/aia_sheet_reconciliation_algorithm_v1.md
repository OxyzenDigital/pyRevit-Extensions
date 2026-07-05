# AIA Sheet Reconciliation & Sequencing Algorithm — Implementation Spec v1
Target: Antigravity 2.0 / pyRevit codebase
Mode: instruction spec for coding agent. No UI design here — pure algorithm + data contracts + acceptance tests.

---

## 0. HARD INVARIANTS (Revit rules — never violate)

| ID | Invariant | Enforcement point |
|----|-----------|-------------------|
| INV-1 | `SheetNumber` is unique project-wide (across ALL disciplines). | Every write. Discipline prefix alone does not guarantee this — validate the full string. |
| INV-2 | View names are unique project-wide. | View creation/rename. |
| INV-3 | A model view (plan/elevation/section/3D) can be placed on exactly ONE sheet. Legends and schedules may repeat. | View-to-sheet assignment. |
| INV-4 | Renumbering a sheet to a number that currently exists throws. All renumber operations MUST use the two-phase protocol (§6). | Apply phase. |
| INV-5 | All writes happen in a single `Transaction` (or `TransactionGroup`) that rolls back atomically on any failure. Never leave the project half-renamed. | Apply phase. |
| INV-6 | No writes occur before user approval. Phases 1–5 are read-only. | Pipeline gate. |

---

## 1. PIPELINE OVERVIEW

```
Phase 1  HARVEST      read-only scan of project (sheets, views, scope boxes, levels)
Phase 2  SCHEMA       generate canonical AIA sheet slots for this project
Phase 3  MATCH        assign existing sheets to slots (deterministic scoring)
Phase 4  RESOLVE      views: reuse / create / ask-user
Phase 5  PLAN         emit curated list + rename/create operation plan (JSON)
-------- USER APPROVAL GATE --------
Phase 6  APPLY        two-phase renumber, renames, sheet creation, view placement
Phase 7  VERIFY       re-harvest, assert final state == approved plan
```

---

## 2. PHASE 1 — HARVEST

Collect into an in-memory model (no mutation):

```python
Sheet:      element_id, number, name, placed_view_ids, is_placeholder
View:       element_id, name, view_type, level_name, scope_box_name,
            crop_active, crop_area_sqft, scale, on_sheet_id (or None)
ScopeBox:   element_id, name, bbox
Level:      element_id, name, elevation   # sorted ascending by elevation
```

Derived facts:
- `level_order`: levels sorted by elevation, index 1..N. This — not level *name* string sorting — drives plan sheet sequence. "Level 1", "Ground", "Mezzanine" must sequence by elevation.
- `grid_size`: count of segment scope boxes per level (0, 4, 9, or 16 → none, 2×2, 3×3, 4×4). If the count is not one of these, flag `GRID_IRREGULAR` and ask user before proceeding.

### 2.1 Scope box canonical names
Expected: `Overall` (or `Building Extent`) plus segments.
- 2×2 → `Segment NW / NE / SW / SE`
- 3×3, 4×4 → `Segment A` … `Segment I/P` (row-major)

If scope boxes exist but names don't match, propose renames (Before → After) in the plan. Match misnamed boxes to canonical positions **geometrically**: sort segment bboxes by (−Y centroid, +X centroid) → row-major order → assign letters. Never guess from names alone.

---

## 3. PHASE 2 — SCHEMA GENERATION (deterministic slot list)

A **slot** = one required sheet: `(discipline, series, ordinal, suffix, canonical_name, required_view_spec)`.

### 3.1 Sequence number assignment rule
`SheetNumber = {DISC}-{SERIES}{NN}{SUFFIX?}` where `NN` is a zero-padded 2-digit ordinal assigned by the **fixed ordering below**, restarting at 01 for each `(discipline, series)` pair. Never derive `NN` from level names or existing numbers — always from canonical order. This is what makes output deterministic.

### 3.2 Canonical ordering within each series

**Series 1 (Plans)** — order:
1. Demolition Plan(s) (per level, ascending)
2. Existing Plan(s)
3. Site Plan (A/C/L only)
4. Floor Plan per level, ascending elevation — **each level's overall sheet immediately followed by its enlarged segment sheets** (see §3.3)
5. Roof Plan
6. Discipline-specific overlay plans in this order, each expanding per-level ascending:
   - A: Dimension → RCP → Finish → Furniture → Equipment
   - M: HVAC → Mechanical Piping
   - E: Power → Lighting → Emergency Lighting
   - P: Plumbing
   - FP: Sprinkler → Fire Alarm (if FP owns FA; else FA under E or T per project setting `fa_owner`, default `FP`)
   - T: Data/IT → Security Device → AV Equipment
   - L: Landscape → Planting → Irrigation

**Series 0 (General)**: General Notes → Code Analysis → Life Safety Plan(s) per level → Legend & Symbols → Survey Control.

**Series 2 (Elevations)**: Exterior N, S, E, W → Interior Sets 1..n.

**Series 3 (Sections)**: Building Sections → Wall Sections → discipline sections.

**Series 5 (Details)**: numbered sequentially as demand requires (1..n).

**Series 6 (Schedules)**: Door → Window → Finish → Equipment → Fixture.

**Series 7 (Diagrams/Risers)**: per the modifier list in the source prompt, in listed order.

**Series 9 (3D)**: Axonometric → Perspective → Cutaway.

### 3.3 Enlargement suffix rule
Enlarged plans do **not** consume their own `NN`. They share the parent's number + suffix:

```
A-101    Floor Plan – Level 1 (Overall)
A-101A   Enlarged Plan – Level 1 – Area A   (or "– Northwest" for 2×2)
A-101B   ...
A-102    Floor Plan – Level 2 (Overall)
A-102A   ...
```

- Suffix set is `A..D`, `A..I`, or `A..P` per `grid_size`. **No gaps**: if the grid is 3×3, all nine A–I slots exist, even if some segments are empty shell — mark unneeded ones as user-skippable in the curated list, but never emit A, B, D with C missing.
- 2×2 names use compass words (`Northwest` etc.); 3×3 and 4×4 use `Area A..P`.
- Revit browser sorts sheet numbers as strings; `A-101, A-101A … A-102` sorts correctly with this scheme. Do not use `A-101.1` style — dot-suffixes break string ordering against three-digit ordinals elsewhere.

### 3.4 Canonical sheet name format
`{Modifier} – {Qualifier}` with en-dash separator, e.g.:
- `Floor Plan – Level 1`
- `Enlarged Plan – Level 1 – Area C`
- `Reflected Ceiling Plan – Level 2`
- `Exterior Elevation – North`

One format everywhere. Matching (§4) normalizes hyphen/en-dash/em-dash before comparing, but writes always use en-dash.

---

## 4. PHASE 3 — MATCHING (existing sheets → slots)

Greedy, multi-pass, deterministic. Each existing sheet matches at most one slot; each slot accepts at most one sheet. Tie-breaks always by lowest `element_id`.

**Normalization** for all string comparison: lowercase; collapse whitespace; unify `-`/`–`/`—`; strip leading zeros in numbers; expand abbreviations from a fixed table (`flr→floor`, `plt→plant`, `elev→elevation`, `dtl→detail`, `sched→schedule`, `enl→enlarged`, `rcp→reflected ceiling plan`, `demo→demolition`, `lvl/l→level`).

**Pass A — exact number + exact name** → `MATCH`.
**Pass B — exact number, name differs** → `RENAME_NAME` (number stays; propose name change).
**Pass C — exact name, number differs** → `RENAME_NUMBER`.
**Pass D — fuzzy**: score = `0.5·token_set_ratio(name) + 0.3·number_affinity + 0.2·content_affinity`
- `number_affinity`: same discipline +0.5, same series digit +0.3, |ordinal diff| ≤ 2 +0.2.
- `content_affinity`: 1.0 if the placed view satisfies the slot's `required_view_spec` (§5.1), else 0.
- Accept ≥ 0.80 → `RENAME_BOTH`. 0.60–0.79 → `CLOSE_MATCH` (user must confirm). < 0.60 → no match.

**Leftovers**: unmatched slots → `MISSING`; unmatched sheets → `EXTRA` (user decides: keep-as-is / renumber-into-schema / delete-nothing-automatically — the tool never deletes sheets).

---

## 5. PHASE 4 — VIEW RESOLUTION

### 5.1 Overall vs Enlarged classification (priority order, stop at first decisive rule)
1. **Scope box** assigned to view: `Overall`/`Building Extent` → overall; `Segment X` → enlarged segment X. Decisive.
2. **Level + scope box** combination pins the exact slot.
3. **Crop area**: crop_area ≥ 70% of overall scope box footprint → overall; ≤ 40% → enlarged. 40–70% → indecisive, fall through.
4. **Scale**: view scale coarser (numerically larger, e.g. 1:100) than the project's median plan scale → overall; finer (e.g. 1:25) → enlarged. Equal → fall through.
5. **Name tokens** (last resort): `enlarged|area [a-p]|nw|ne|sw|se|segment` → enlarged.
6. Otherwise → `UNCLASSIFIED`, queue for manual user classification. Never guess.

### 5.2 Per-slot view decision
For each slot in the approved-pending plan:
- Matched sheet already carries a satisfying view → `VIEW_REUSE`.
- Unplaced view exists that satisfies the spec → `VIEW_PLACE` (respect INV-3: if it's already on another sheet, offer duplicate-as-dependent instead).
- No view, but deterministically creatable (floor/ceiling plans: level + scope box known) → `VIEW_CREATE` with recipe `{view_type, level_id, scope_box_id, scale, name}`. View name mirrors sheet name (`Floor Plan – Level 1`, `Enlarged Plan – Level 1 – Area C`) — this satisfies the "view names match sheet intent" requirement.
- Not deterministically creatable (elevations, sections, details, 3D) → `VIEW_ASK_USER` with a filtered candidate list (same view_type, unplaced first).

---

## 6. PHASE 6 — APPLY: TWO-PHASE RENUMBER PROTOCOL

Renumber collisions are the classic failure (A-101→A-102 while A-102 exists → exception, or worse, a cascade that half-applies). Protocol:

```python
with TransactionGroup("AIA Reconciliation"):
    with Transaction("Phase 1 – park"):
        for op in renumber_ops:
            op.sheet.SheetNumber = "ZZ~" + str(op.sheet.Id.IntegerValue)  # guaranteed unique, never collides with schema space
    with Transaction("Phase 2 – finalize"):
        for op in renumber_ops:
            op.sheet.SheetNumber = op.target_number
            if op.target_name: op.sheet.Name = op.target_name
        # then: create missing sheets, create views, place views, rename scope boxes
    # any exception → RollBack the whole group
```

Pre-apply validation (fail fast, before any transaction):
- Assert target numbers are unique among themselves AND against all untouched sheets.
- Assert target view names unique against all untouched views.
- Assert every `VIEW_PLACE` view is currently unplaced or flagged dependent-duplicate.

Phase 7 re-harvests and diffs against the approved plan; any drift → report, do not silently retry.

---

## 7. CURATED OUTPUT CONTRACT (Phase 5 → UI)

One flat JSON list, already in final AIA order (overall then its suffixes, no gaps), so the UI just renders rows:

```json
{
  "schema_version": "1.0",
  "grid_size": "3x3",
  "rows": [
    {
      "seq": 12,
      "target_number": "A-101",
      "target_name": "Floor Plan – Level 1",
      "status": "MATCH | RENAME_NAME | RENAME_NUMBER | RENAME_BOTH | CLOSE_MATCH | MISSING | EXTRA | UNCLASSIFIED",
      "existing_number": "A-1.01",
      "existing_name": "1st Flr Plan",
      "sheet_element_id": 445123,
      "view_action": "VIEW_REUSE | VIEW_PLACE | VIEW_CREATE | VIEW_ASK_USER | NONE",
      "view_element_id": 445200,
      "view_recipe": null,
      "requires_user_decision": false,
      "match_score": 0.91
    }
  ],
  "scope_box_renames": [{"element_id": 1, "before": "Box 3", "after": "Segment C"}],
  "warnings": ["GRID_IRREGULAR: 7 segment boxes found on Level 2"]
}
```

Approval semantics: user may toggle any `RENAME_*` back to keep-existing, skip any `MISSING`, and must resolve every `requires_user_decision: true` row before Apply enables.

---

## 8. ACCEPTANCE TESTS (must pass before this module ships)

1. **Determinism**: run Phases 1–5 twice on the same file → byte-identical JSON.
2. **Collision swap**: project has A-101 and A-102 that must swap numbers → applies cleanly via park phase, no exception.
3. **No-gap suffixes**: 3×3 grid with only 5 enlarged views existing → plan still emits all A–I slots; missing ones flagged, none skipped.
4. **Uniqueness cross-discipline**: schema proposing `FP-101` while an existing misfiled sheet holds `FP-101` as EXTRA → pre-apply validation catches it before transaction.
5. **Level ordering**: levels named `Ground, Mezz, Level 2` (elevations 0, 12, 24) → sequence by elevation, not alphabetical.
6. **View reuse guard**: view already on a sheet is never proposed for `VIEW_PLACE`; dependent-duplicate offered instead.
7. **Rollback**: inject a failure mid-Phase-2 apply → project state identical to pre-apply.
8. **Unclassifiable view**: plan view with no scope box, mid-range crop, median scale, neutral name → lands in `UNCLASSIFIED`, never auto-assigned.
9. **En-dash tolerance**: existing sheet `Floor Plan - Level 1` (hyphen) matches slot `Floor Plan – Level 1` as Pass A, not RENAME.
10. **Never delete**: EXTRA sheets survive Apply untouched unless user chose renumber-into-schema.
