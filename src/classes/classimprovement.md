We want to improve the Linelist class in the src/designer folder. This class
uses most of the other sub-classes to orchestrate the constructions/bulding of a linelist.

What could be the improvements to the class? Write down below tasks to improve it,
and make more consise, defensive, future proof and bugproof

- [done] Introduce an `ApplicationStateScope` helper that snapshots ScreenUpdating, DisplayAlerts, Calculation, and EnableAnimations before `BusyApp` changes them, and guarantees restoration (even on errors) so Excel is not left in an unusable state.
- [done] Replace the `sheetScope`/`CodeScope` magic numbers with public enums or constants exposed through the interface, then centralize sheet-name construction (prefix + truncation) in a single `FormatSheetName` helper used by both `AddOutputSheet` and `Wksh` so the rules stay in sync.
- [done] Split the monolithic `Prepare` routine into focused private methods (e.g. `CreateTemporarySheets`, `ExportSpecs`, `CreateAnalysisSheets`, `TransferCodeModules`, `TransferForms`), sequence them from `ILinelist_Prepare`, and add early guards for missing dictionary/spec dependencies to make the flow testable.
- [done] Abstract the module/form transfer logic behind an `ICodeTransferService` that detects whether VBIDE automation is available, uses the current VBIDE path on Windows, but falls back on macOS to workbook-native copy strategies (e.g. cloning from a pre-populated template workbook or reading pre-exported module text from hidden sheets) so the prepare flow still succeeds without cross-compilation or fragile VBIDE calls.
- [done] Move all file-system work (`TemporaryFolder`, temp export paths, `Kill` calls) behind an `ITempFileService` so we can unit test the interactions, ensure directories exist, and handle locked file scenarios without relying on repeated `On Error Resume Next` blocks.
- [done] Add a `Dispose`/`Reset` pathway that clears `this.outWkb`, reinstates the temporary folder state, and releases COM references after `SaveLL`/`ErrorManage` so successive runs start from a clean slate.
- [done] Wrap workbook activation/sheet selection logic into a `SheetNavigator` utility that checks for sheet existence using translations once, caches the admin/instruction sheet names, and eliminates the repeated `On Error Resume Next` activation patterns in `SaveLL` and `Prepare`.
- [done] Write high-level integration tests around `ILinelist_Prepare` and `ILinelist_SaveLL` using the existing test fixtures to confirm sheet creation order, temp folder lifecycle, and module transfer behaviour, preventing regressions while refactoring.

You are a skilled VBA programmer with the task to improve the class walk through all those tasks.
You should follow as closely as possible instructions in instruction.md, and update the
tasks with the tag [done] if it is the case. Do not modify any legacy code, add new codes in the classes/linelist folder. You are allowed to proceed steps by steps and not do everything at once.
At each step, re-read the classimprovement and instructions to know where you stoped (never forget to add the [done] tag), and start from where you stopped.
Your turn.
