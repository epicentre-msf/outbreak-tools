**Instructions**

Implement your improvements steps by steps, incrementally. Do not touch legacy code
in deisgner-code or in setup-code. NEVER EVER EVER EDIT legacy code; instead
create new classes when required.

0. NEVER USE DLLS OR SCRIPTING DICTIONARY, I AM WORKING FOR MACOS ALSO, USE BETTERARRAY IF NEEDED.
1. ADD DETAILED COMMENTS AND ANNOTATIONS TO THE CODE !!!
2. Use the same Error management logic based on ProjectError
3. Add checks and checkings notifications if required
4. Always keep the interface at the end of class.
5. Aim for efficiency, always write a code that should execute fast.
6. Do not forget to add sections and other parameter annotation tags to the class implementation.
7. Clean, document, and comment also the interface.
8. Write some tests associated to the class, leverage the TestHelpers if needed.
10. Use TypeName instead of TypeOf for object type checking.
11. Write new classes in src/classes and new tests in src/tests. Do not overwrite code in src/designer
12. Add Comments to the current instructions.md file on what is done in the Progress section. There is no need to do everything at once, you can go progressively; but do not overwrite all of what is written in the current file. Add [DONE] or [NOTDONE] tags to instructions
items once you are done with them, on the progress session.
13. Classes should NOT DEPEND on modules. Classes can use other classes, but they should be self contained and not rely on modules code outside the class.
14. ALWAYS Make sure the new classes fit well with the other classes created in src/classes
15. Pays extremely attention when working with quotes and mutiple quotes. Avoid syntax errors either by using Chr(34) or by escaping correctly double quotes using
required VBA syntax.
16. Always run unix2dos src once you are done with the files
17. DO NOT EDIT DICTIONARYTESTFIXTURE.BAS!

**Progress**

- [DONE] Implemented DiseaseApplicationState guard to capture BusyApp toggles and guarantee restoration.
- [DONE] Added DiseaseWorksheetManager to remove sheets safely without legacy display alert loops.
- [DONE] Delivered DiseaseReportManager to fix HasReport/RemoveReportStatus bugs with accurate row pruning.
- [DONE] Built DiseaseImporter/DiseaseImportSummary to complete ImportElements merge logic with reporting data.
- [DONE] Introduced DiseaseExportWorkbook so export sessions create, save, and release workbooks safely.
- [DONE] Added DiseaseExporter with array-based dictionary/migration exports, plus translation cache for header lookups and dedicated tests.
- [DONE] Introduced DiseaseSheetBuilder to construct new disease worksheets with translated headers, validations, and smoke tests.
- [DONE] Added DiseaseLogger to capture import/export actions and wired logging into DiseaseImporter with accompanying tests.
- [DONE] Added integration tests covering disease add/export/import/remove collaboration to validate end-to-end workflows.
- [NOTDONE] Remaining tasks: wire collaborators into the legacy IDisease orchestrator when refactor begins.



