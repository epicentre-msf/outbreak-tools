We plan to create a new class based on the importmoduleandclasses.bas file called 
Development.cls hat will add simple things on top of what we already have in the bas
file. The bas file manage the import/export of codes into one specific file, basically
the "development" part of our work.

0- The class should instanciate with one worksheet and make sure the worksheet has the
ranges for classes/modules, same name as in the bas file. Test range is optional.

1- The class should expose a AddCodeSheets method to add a worksheet with names of files to import.

2- It should also expose a AddClassTable where we add a listObject in the fourth column of the Code sheet for classes.
    Something simple: A listobject with two columns: classes and hasinterface. The top of the listobject should have
    the tags "general classes", the folder is added by the user.

3- The class should expose a AddModuleTable which is a listObject with only one column

4- The class should expose a AddFormsTable which is a listObject with two columns "modules", "corresponding form"
    the tag for this is "general form modules", but the forms columns are filled by the user.

4- Always ask for validation before adding something to the worksheet

5- The class should implement the import/export processes in classimprovement using the tags on top of the listObjects.

6- The class should implement a AddHiddenSheet and a AddProtectedSheet method so some worksheets are hidden/protected for deployment

7- The class should implement a Deploy method (that takes a passwords object as argument) that protects protected sheets,
hide those that need to be hidden, lock the workbook if needed

8- The class should implement a AddFormsCodes, that will look for all the forms table in the codes worksheet, and move the
codes to the corresponding forms in the second columns

10- The class should implement a AddTestTable using same rules previously defined

9- You should track the classes/modules/tests added using a counter (it could be a named attached to the worksheet). Classes ListObjects should be ClassesLo[number], 
modules listobjects ModulesLo[number], tests listobjects TestsLo[number].

Use IOSFile class for adding paths.


My plan is to wrap the existing `ImportModuleAndClasses` procedural workflow into a stateful object that owns the Dev worksheet and the code tables so we can progressively build more tooling around it. Concretely:

- [done] Constructor: accept (and validate) a worksheet reference. On creation ensure the mandatory named ranges exist and cache them (modules/classes/tests) while allowing the tests range to be missing. Store counters for the next ListObject suffix in worksheet-level names so that the naming survives workbook reloads.

- [done] Worksheet validation helpers: centralize checks for existing named ranges/list objects, re-use the prompting/validation logic before any insert. I’ll expose a `RequestConfirmation(ByVal action As String) As Boolean` that wraps the “always ask before adding” rule.

- [done] Table builders: `AddCodeSheets`, `AddClassTable`, `AddModuleTable`, `AddFormsTable`, `AddTestTable` will share a private `CreateListObject(tag, columnNames(), counterName)` that handles header rows, tagging, naming (`ClassesLo1`, `ModulesLo2`, …), and tagging metadata in the top-left cell. For class tables the second column (`hasinterface`) will be seeded with validation list yes/no. For forms tables add the general tag in the header row and leave user-fillable columns blank as required.

- [done] Hidden/protected sheet support: methods `AddHiddenSheet`/`AddProtectedSheet` will rely on workbook-level helpers look , while registering them in an internal BetterArray so `Deploy` can iterate them. Protection passwords will be provided through a `Passwords` interface (likely the project’s `IPasswords`). Do not create new worksheets, the worksheets should exists, you should betterArray to stock them and the apply the protection matrix if required at the deploy step.

- [done] Import/export: encapsulate the logic currently in `ImportModuleAndClasses`. I’ll pull the directory resolution (`ResolveOutputDirs`) and `TransferCode` behaviours into private methods, but use the ListObject metadata on each table (tag + folder + interface flag) to determine scope. The new class will expose `ImportAll`/`ExportAll` methods that orchestrate modules/classes/tests using a common `ProcessTable` routine with the appropriate `ImportedFileScope` value. IOSFile/IOSFiles will be used for path handling to respect the existing abstraction.

- [done] Forms code deployment: `AddFormsCodes` will enumerate each forms table, derive the module/form pair, and move code modules to their associated form modules using the same VBProject component manipulation as the `.bas` module. I’ll encapsulate the VBIDE access behind a helper that can be reused for the other transfer operations. There is a CopyCodes sub that you can leverage.

- [done] Deployment: `Deploy(passwords)` will iterate registered protected sheets, apply protection, hide the hidden sheets, and optionally lock the workbook. It will respect the user confirmations and log actions so tests can inspect side effects.

- [done] Testing/supportability: expose small, testable helpers (e.g., `NextCounterValue`, `EnsureNamedRange`) so the existing test fixtures (`ImportModuleAndClasses.bas`) can be adapted. I’ll scaffold new tests verifying table creation, counter increment, and deploy behaviour, reusing the formulas fixtures structure.




As a well skilled VBA developper, you are tasked with building the class.
You should follow closely instructions.md, and respects any of the constraints in the
file. You can plan your work and implement progressively, but you must
add a [done] / [notdone] tag to the current list to update on where you are.
You should right very efficient, compact and tightened code like in LinelistTranslation. No need to implement mutiple classes or add a lot of layers. Efficiency and compactness should be your leitmotiv. We aim to
reach an output with as minimum as possible codes and create a tightened coherent class we can improve progressively. Always add annotations and comment.

