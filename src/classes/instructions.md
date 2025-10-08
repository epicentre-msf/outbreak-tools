**Instructions**

Implement your improvements steps by steps, incrementally. Do not touch legacy code

0. NEVER USE DLLS OR SCRIPTING DICTIONARY, I AM WORKING FOR MACOS ALSO, USE BETTERARRAY IF NEEDED.
1. ADD DETAILED COMMENTS AND ANNOTATIONS TO THE CODE !!!
2. Use the same Error management logic based on ProjectError
3. Add checks and checkings notifications if required
4. Always keep the interface at the end of class.
5. Aim for efficiency, always write a code that should execute fast.
6. Do not forget to add sections and other parameter annotation tags to the class implementation.
7. Clean, document, and comment also the interface.
8. Write some tests associated to the class, leverage the TestHelpers if needed.
10.Use TypeName instead of TypeOf for object type checking.
11. Write new classes in src/classes and new tests in src/tests. Do not overwrite code in src/designer
12. Add Comments to the current instructions.md file on what is done in the Progress section. There is no need to do everything at once, you can go progressively; but do not overwrite all of what is written in the current file. Add [DONE] or [NOTDONE] tags to instructions
items once you are done with them, on the progress session.
13. Classes should NOT DEPEND on modules. Classes can use other classes, but they should be
self contained and not rely on modules code outside the class.
14. ALWAYS Make sure the new classes fit well with the other classes created in src/classes
15. Pays extremely attention when working with quotes and mutiple quotes. Avoid syntax errors either by using Chr(34) or by escaping correctly double quotes using
required VBA syntax.
16. Always run unix2dos src once you are done with the files

**Progress**

- [DONE] Introduced `LinelistApplicationStateScope` and companion tests to guard Excel busy state transitions while leaving legacy code untouched.
- [DONE] Added `LinelistSheetNameFormatter` with scoped enums and tests to replace Linelist sheet-name magic numbers and centralise formatting logic.
- [DONE] Added `LinelistCodeTransferService` with VBIDE/fallback strategies, wired the preparation step and workbook accessor abstractions, and covered the new pipeline orchestration with targeted tests.
- [DONE] Added `LinelistSheetNavigator` with translation caching and activation helpers to replace ad-hoc sheet selection logic, backed by focused tests.
- [DONE] Built `LinelistSaveWorkflow` to consolidate admin/instruction activation, protection, and lifecycle disposal for saves, including a suite of workflow-focused tests.
- [DONE] Added high-level integration tests exercising Linelist preparation and save workflows to guard sheet ordering, module transfers, and temp lifecycle expectations.
- [DONE] Introduced `LinelistLifecycleManager` plus supporting interfaces to reset/Dispose output workbooks, ensuring temp folders reset and behaviour validated with dedicated tests.
- [DONE] Added `LinelistTempFileService` with sanitised path helpers and tests to replace scattered temporary folder handling.
- [DONE] Built `LinelistPreparationPipeline` with stage-specific steps, context, guards, and tests to break down the Prepare flow and validate dependencies.
