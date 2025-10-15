We plan to improve the IDisease class in src/msetup-codes. Our goal
is to create new classes in src/classes/msetup that has the improvements. We
are also tasked with writting tests related to our newly created classes.
We should never ever modify legacy code; we must focus on only creating
new classes and tests if required.

As a well skilled VBA developper, you are tasked with building the improvements.
You should follow closely instructions.md, and respects any of the constraints in the
file. You can plan your work and implement progressively, but you must
add a [done] / [notdone] tag to the current list to update on where you are.

DO NOT MODIFY LEGACY CODE!!!!!!!!!!!! CREATE NEW CLASSES INSTEAD

## Quick Wins (must fix first)
- [done] `Disease.Remove` loops `For counter = 1 To 4` when deleting the worksheet (`src/msetup-codes/classes/Disease.cls:419`). This silently relies on `On Error Resume Next` and leaves alerts off; delete once after checking `sheetExists` and move the `DisplayAlerts` restoration into a `Finally`-style handler.
- [done] `HasReport` never assigns its return value and writes to an undeclared `needReport` variable (`src/msetup-codes/classes/Disease.cls:549-567`). Return the boolean and remove the stray symbol so consumers stop seeing a permanent `False`.
- [done] `RemoveReportStatus` refers to an undeclared `delRows` instead of `delRowsTab` and will not compile under `Option Explicit` (`src/msetup-codes/classes/Disease.cls:584-599`).
- [done] `ImportElements` declares `counter As Range`, uses an undeclared `varName`, and leaves the “import new variables” loop empty (`src/msetup-codes/classes/Disease.cls:666-752`). Fix the declarations and finish the merge logic so imports do not silently drop rows.
- [done] `BusyApp` disables events, screen updating, animation, and flips calculation to manual (`src/msetup-codes/classes/Disease.cls:996-1002`) but there is no paired reset. Cache the previous Application state and restore it with an `On Error` guard to prevent Excel from remaining “stuck” after exports.

## Performance & Stability
- [done] Wrap bulk worksheet edits (Add/Import/Export) in a `With Application` scope that toggles `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, and `Calculation`, restoring them in `Finally` logic. This removes flicker and avoids leaving Excel in a bad state when an error occurs mid-way.
- [done] Replace repeated cell-by-cell reads/writes with array-based transfers. In `ExportDisease` and `ExportForMigration`, read each ListObject into a `Variant` array once, transform in-memory, and write back with a single assignment. This will dramatically reduce COM round-trips on large dictionaries.
- [done] Cache frequently accessed COM objects (e.g., `Worksheet`, `ListObject`, `Range`) outside tight loops and prefer `With` blocks to avoid redundant property lookups.
- [done] Reuse the `OutputWkb` workbook safely: create it locally inside export routines, wrap Save/Close calls in `On Error` handlers, and release references so subsequent runs do not attempt to operate on a disposed workbook.

## Maintainability & Extensibility
- [done] Split the monolithic `Disease` class into focused collaborators: e.g., `DiseaseSheetBuilder` (Add), `DiseaseExporter`, `DiseaseImporter`, and `DiseaseRegistry` (name list maintenance). Keeping the interface (`IDisease`) slim with orchestration responsibilities will make future enhancements less risky.
- [done] Replace magic strings such as `"foreign"`, `"actual"`, `"yes"`, etc. with `Enum` based flags or dedicated value objects. This avoids typos and makes IntelliSense guide the caller.
- [done] Centralise constant translations: pull the repeated `.TranslatedValue` lookups into a translation cache so we set up the localized headers once per export instead of per-cell.
- [done] Audit every `On Error Resume Next` and scope it tightly; follow it with explicit error checks so logic errors surface during testing instead of being swallowed.

## Import/Export Workflow Improvements
- [done] When merging imports, collect the existing worksheet into a keyed dictionary (variable name -> row index) so lookups are O(1) and you can decide whether to update, append, or report missing variables without repeated `Range.Find` calls.
- [done] Complete the “new variables” branch inside `ImportElements`: append missing rows to the ListObject, fill defaults, and flag them in the import report so downstream processes know what changed.
- [done] Move “import report” bookkeeping into a dedicated helper that can add/remove rows using `ListRows.Add`/`ListRows(counter).Delete`. This removes the fragile manual range walking and ensures tables stay properly sized.

## Testing & Tooling
- [done] Add integration tests around `IDisease.Add/Remove/Export/Import` using the existing `TestDisease` module to lock in behaviours (worksheet count, dropdown contents, report tables, etc.).
- [done] Introduce guard tests for the Application state helper so we verify screen updating, events, and calculation are restored even when an error is raised mid-operation.
- [done] Consider adding a lightweight logging layer (e.g., to the import report sheet) so unexpected branches are observable without stepping through the VBA editor.


FOLLOW CLOSELY INSTRUCTIONS.MD, NEVER EDIT LEGACY CODE, AND UPDATE/TRACK YOUR PROGRESS WITH [done] tag.
There is no need to do everything at once, you can implement progressively.
