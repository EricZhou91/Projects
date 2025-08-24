# Architecture Notes

## Inputs

- **Split pipeline** expects a workbook with:
  - `Raw Data` sheet (columns: First Name, Last Name, Position Title, PS_DEPT, Position Number, Job Code, Reports To, Incumbency Status, Expected Return Date, Class Description).
  - `Branch Identifier (NEW)` sheet mapping PS_DEPT → Branch code (e.g., OCIA).

- **FTE pipeline** expects branch-named sheets where columns A:F are:
  - Name, Position Title, Department ID, Position Number, Job Code, Reports to

## Rules

- FTE excludes names containing: `(A/O)`, `(LoA)`, `(M/L)`, `(S/O)`, `(LTIP)`, `(FxT)`.
- Split pipeline appends tokens to names based on Raw Data fields:
  - `Incumbency Status` = OUT → `(A/O)`
  - `Incumbency Status` = IN  → `(A/I)`
  - `Expected Return Date` non-empty → `(LoA)`
  - `Class Description` = Fixed Term → `(FxT)`

## Outputs

- Split pipeline:
  - New workbook with sheets: **Director**, **Manager**, and one per Branch.
  - Manager sheet has Directors appended; duplicates by Position Number are merged; missing “Reports To” highlighted.

- FTE pipeline:
  - **Combined list** sheet consolidated from all branch tabs.
  - **FTE Summary** with per-branch counts and total.

## Extensibility

- Edit `mod_Constants.bas` to add/rename branches and tokens.
- Column letters for Raw Data are centralized in constants so they’re easy to adjust.
- Helpers in `mod_Helpers.bas` ensure Excel state resets on error.