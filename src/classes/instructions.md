Linelist class improvement

Create a new Linelist class in src/classes with improvements.


Outlined improvements are:



**Instructions**


Implement your improvements steps by steps, incrementally. Do not touch the source
classes TableSpecs and GraphSpecs and focus only on reading them to create new one and improvements in src/classes, 
based on discussions and abstraction strategies. Follow these rules:

0. NEVER USE DLLS OR SCRIPTING DICTIONARY, I AM WORKING FOR MACOS ALSO, USE BETTERARRAY IF NEEDED.
1. ADD DETAILED COMMENTS TO THE CODE !!!
2. Use the same Error management logic based on ProjectError
3. Add checks and checkings notifications if required
4. Always keep the interface at the end of class.
5. Aim for efficiency, always write a code that should execute fast.
6. Do not forget to add sections and other parameter tags to the class implementation.
7. Clean, document, and comment also the interface.
8. Write some tests associated to the class, leverage the TestHelpers if needed.
9. Never use Enums as variable type; use bytes instead because enum can generate weird bugs on MacOS
10. Use TypeName instead of TypeOf for object type checking.
11. Write new classes in src/classes and new tests in src/tests. Do not overwrite code in src/designer
12. Add Comments to the current instructions.md file on what is done in the Progress section. There is no need to do everything at once, you can go progressively; but do not overwrite all of what is written in the current file. Add [DONE] or [NOTDONE] tags to instructions
items once you are done with them, on the progress session.
13. Classes should NOT DEPEND on modules. Classes can use other classes, but they should be
self contained and not rely on modules code outside the class.
14. ALWAYS Make sure the new classes fit well with the other classes created LLVar* Sections* classes in src/classes.


[PROGRESS] _write down what is done/not done down here_

[DONE] ListBuilderCoordinator orchestrates layout, section builder, and worksheet preparer with detailed tests (src/classes/implements/ListBuilderCoordinator.cls, src/tests/TestListBuilderCoordinator.bas).

[DONE] ListWorksheetPreparer centralises busy state handling with supporting unit tests (src/classes/implements/ListWorksheetPreparer.cls, src/tests/TestListWorksheetPreparer.bas).

[DONE] ListContextCache collects reusable linelist collaborators to minimise repeated lookups (src/classes/implements/ListContextCache.cls, src/tests/TestListContextCache.bas).

[DONE] Shared interfaces and value objects introduced for layouts, sections, and contexts powering future HList/VList refactors (src/classes/interfaces/IList*.cls, src/classes/implements/ListSectionDescriptor.cls).

[DONE] ListBuilderFactory composes caches, section builders, and orientation-specific strategies (src/classes/implements/ListBuilderFactory.cls, src/tests/TestListBuilderFactory.bas).

[DONE] ListBuildService exposes a simplified build entry-point backed by the factory (src/classes/implements/ListBuildService.cls, src/tests/TestListBuildService.bas).


