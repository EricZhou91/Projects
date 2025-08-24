# Government Workforce Automation (VBA)

This repo contains **refactored** Excel/VBA modules for:
1) **Split Raw Data → Branch/Director/Manager** sheets + Visio helpers.
2) **FTE Headcount & Branch Consolidation** across 10+ branch tabs.

- Centralized constants (`mod_Constants.bas`)
- Safer macro wrappers (`mod_Helpers.bas`)
- No real data included — see `/examples` for synthetic input/output workbooks.

## Modules

- `src/split-data/Module_SplitRawData.bas` — entry point `SplitRawData` opens a workbook with `Raw Data` + `Branch Identifier (NEW)` and emits clean sheets per branch and role.
- `src/fte-summary/Module_FTECount.bas` — `CombineBranchesToList` builds **Combined list**; `FTECountOnAllSheets` creates **FTE Summary**.
- `src/fte-summary/Module_CombineBranches.bas` — post-processing helpers for highlighting missing reporting relationships and merging duplicate Position Numbers.
- `src/fte-summary/Module_OpenSaveClose.bas` — simple per-file utility.
- `src/fte-summary/mod_Constants.bas` — shared constants and lists.
- `src/mod_Helpers.bas` — wrapper helpers for safe execution.

## Examples (synthetic)

- `examples/sample_raw_data.xlsx` — input for Split pipeline (Raw Data + Branch Identifier).
- `examples/sample_structure.xlsx` — branch tabs for FTE pipeline.
- `examples/sample_outputs.xlsx` — empty output shells (for reference).

## Usage (high-level)

- **Split pipeline**
  1. `ALT+F11` → `File > Import` all modules from `src/mod_Helpers.bas`, `src/fte-summary/mod_Constants.bas`, and `src/split-data/Module_SplitRawData.bas`.
  2. Run `SplitRawData` → pick `examples/sample_raw_data.xlsx`.
  3. Output workbook will contain **Director**, **Manager**, and branch sheets.

- **FTE pipeline**
  1. `ALT+F11` → `File > Import` modules from `src/mod_Helers.bas`, `src/fte-summary/mod_Constants.bas`, `src/fte-summary/Module_FTECount.bas`, `src/fte-summary/Module_CombineBranches.bas`.
  2. In a workbook containing branch tabs, run `CombineBranchesToList` then `FTECountOnAllSheets`.
  3. Optionally run `HighlightMissingReportsTo` and `MergeDuplicatePositionNumbers` on the active sheet.