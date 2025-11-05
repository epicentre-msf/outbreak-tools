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
18. Avoid naming variables that are ambiguous such as "sheet", "workbook", etc. use sh for sheet, wb for workbook. And avoid using a name variable that can conflict with a function or sub. On example of non acceptable usage is something like "devSheet = DevSheet()". VBA is not that strict with case sensitive and will clash in case you have those type of variable namings. NAMING CONVENTION IS SUPER IMPORTANT!!!!!!!!
19. Use existing classes, do not reinvent the wheel.
20. ALWAYS READ IMPLEMENTATION OR IMPROVEMENT FILE CORRECTLY BEFORE PROCEEDING.
21. Use BetterArray instead of collections if necessary.
22. ALWAYS add failure management for tests.

**Progress**
- [DONE] Expanded SetupErrors tests to cover dictionary, choices, and exports error constants.
- [DONE] Added SetupErrors test coverage for factory argument validation and busy state restoration on failure.
- [DONE] Reworked SetupErrors tests to use worksheet fixtures and assert dictionary/choices inconsistencies end-to-end.
- [DONE] Modernized UpdatedValues with optional identifiers, sheet-wide registration helpers, targeted removal, and refreshed tests.
- [DONE] Restored UpdatedValues helper stack to initialise identifiers and registry names consistently.
- [DONE] Reworked UpdatedValues to maintain per-table registries, prevent cross-table pruning, and expanded tests for multi-table coverage.
- [DONE] Added structured failure logging to TestUpdatedValues to meet test failure management requirements.
- [DONE] Simplified UpdatedValues registry naming to table+sheet keys, centralised name index tracking, introduced targeted CheckUpdate tags, and aligned tests with the new workflow.
- [DONE] Extracted EventSetup workbook orchestration into a dedicated class/interface pair with supporting tests and workbook delegates.
- [DONE] Consolidated EventSetup analysis caching and sheet unlock helpers to cut duplication and tighten Worksheet_Change execution paths.
- [DONE] Added SetupPreparation class and interface to handle dropdown registration and updated values initialisation with accompanying unit tests.
- [DONE] Ported setup table validation logic into SetupPreparation with comprehensive tests covering dictionary, exports, analysis, and checking sheets.
- [DONE] Implemented MasterSetupVariables manager with metadata caching, dropdown validation, translation helpers, clone/import support, and unit tests exercising column creation, choice synchronisation, and export behaviour.
- [DONE] Added MasterSetupPreparation workflow to register master dropdowns, initialise variables, and cover the behaviour with dedicated unit tests.
