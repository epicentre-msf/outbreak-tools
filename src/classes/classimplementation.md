EventsMasterSetup
==========================

Goal
----
Implement a Event orchestration class for masterSetup, like for the setup
with a simple wrapper in the corresponding manager delegating to EventMasterSetup
service, pretty much exactly like in the setup.

When we open the workbook, we do the same as in the setup.
Our caches are the updater, the dropdown, the ribbon. There is no analysis / dictionary in the master setup, but we have a variable clas (MasterSetupVariables)

If something changes on the variable worksheet, if the name of the variable changes
for example, for each of the disease worksheet 



We want to implement a event orchestration class for mastersetup like in
setup. The class should mimic probably sheet changes like in 
The class must wrap a `CustomTable` and favour minimal, fast workbook interactions.

Functional requirements
-----------------------
1. [DONE] **Factory** – `Create(listObject As ListObject)` must verify that all required columns exist (case-insensitive). When missing, initialise them before instantiating the wrapper.
2. [DONE] **Choice synchronisation** – expose a method that updates the “Choices Values” column for a given variable using an `ILLChoices` instance. Concatenate the values with `" | "` as the delimiter.
3. [DONE] **Initialisation** – accept an `IDropdownLists` dependency to apply validation to the “Default Status” column. During initialisation:
   - Register worksheet-level hidden names pointing to each required column index. We will use those names to point to columns in the future (or tags) as columns could be translated during the lifecycle of the class.
   - Store a hidden-name flag indicating that the table has been initialised.
   - Apply conditional formatting to the “Variable Name” column that highlights duplicates (reuse the colour scheme employed by `SetupTranslations`, keeping the tone subtle).
4. [DONE] **Translations** – provide helpers that consume an `ITranslationObject` to:
   - Translate column headers.
   - Translate the “Variable Section”, “Variable Name”, “Label”, “Choices Values”, and “Comments” ranges.
   - Translate the entire table payload when needed (only all specific translatable columns ranges should be translated).
   Use name tags for lookups so that translated headers do not break column detection.
5. [DONE] **Lookup API** – expose quick getters for a single variable that return its label, default choice, choice values, default status, and section.
6. [DONE] **Row management** – support adding and removing rows (similar ergonomics to `CustomTable.ManageRows`).
7. [DONE] **Performance** – minimise recalculation cost. Cache references where practical, employ `HiddenNames` for repeated lookups, and avoid expensive workbook interactions.
8. [DONE] **Import / export** – allow cloning the table into another workbook and restoring from a peer instance. Persist the hidden names during export so the imported copy retains metadata.

Implementation notes
--------------------
- Maintain a small internal state record (e.g., column indices, cached `CustomTable`, hidden names).
- Keep error handling defensive but lightweight; surface meaningful errors when prerequisites are absent.
- Ensure public members are easy to unit test alongside existing Setup helpers.

As a well skilled VBA developper, you are tasked with building the class, with
different tasks listed below

Implementation plan
-------------------
1. [DONE] Define module-level state structure tracking the source `ListObject`, cached `CustomTable`, column indexes, and hidden-name manager.
2. [DONE] Implement the `Create` factory: verify/initialise required columns, seed the state structure, and return the predeclared instance.
3. [DONE] Add column discovery utilities that load indices from hidden names or the header row, persisting any new mappings back to hidden names.
4. [DONE] Implement the choice synchronisation method that leverages an `ILLChoices` dependency to populate the “Choices Values” column.
5. [DONE] Provide an initialisation routine that applies dropdown validation, registers hidden-name metadata, and wires conditional formatting for duplicate names.
6. [DONE] Build translation helpers: header translation, targeted column translations, and whole-table translation while preserving column tagging.
7. [DONE] Implement lightweight lookup methods that return the requested field values for a supplied variable key.
8. [DONE] Add row-management helpers for inserting and deleting rows, deferring to the underlying `CustomTable` where possible.
9. [DONE] Optimise performance paths with caching, hidden names, and minimal worksheet interaction (including reset hooks for cache invalidation).
10. [DONE] Implement export/import routines that clone the table state, copy hidden names, and reconstruct the class from an exported workbook.

Your role is to implement progressively, but you must add a [done] / [notdone] tag to the current list to update on where you are. You should follow closely instructions.md, and respects any of the constraints in the
instruction file. Once again aim for performance. There is no need to implement everything at once,
you can implement features progressively.
