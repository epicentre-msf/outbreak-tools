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
18. Avoid naming variables that are ambiguous such as "sheet", "workbook", etc. use sh for sheet, wb for workbook. And avoid using a name variable that can conflict with a function or sub. On example of non acceptable usage is something like "devSheet = DevSheet()". VBA is not that strict with case sensitive and will clash in case you have those type of variable namings.
19. Use existing classes, do not reinvent the wheel.

**Progress**
- [DONE] Synced `LLdictionary` total export tracking with hidden sheet name storage and interface setter support.
- [DONE] Added regression tests in `TestLLdictionary.bas` covering setter persistence, import, and export of the total export counter.
- [DONE] Hardened `LLExport` dictionary coordination and caching, plus added regression coverage for dictionary-optional row management.
- [DONE] Exposed `LLExport.SyncDictionaryExports` via `ILLExport` with regression tests for optional dictionary usage.
