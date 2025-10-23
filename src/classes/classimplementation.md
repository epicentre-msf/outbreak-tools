We want to convert improve the LLShowHide class that hides columns in a linelist.
The class is  src/designer-codes/LLShowHide.cls. We want to be able to hide elements
on a CRF worksheet, a HList worksheet and a printed HList worksheet.

We want to register the show/hide state in listObjects like done in the class,
but also to be able to import/export the show hide state into workbook.

We also want to implement on a crf or print worksheet, an "applyHlistShowHideLogic" that
will apply the show hide logic for the corresponding Hlist, and applyDictionaryShowHideLogic
that will apply show hide logic from the dictionary on all the worksheets (crf, vlist, hlist, printed). On crf and printed though, formulas, choice_formulas, list auto, are always hidden.

The class has a lot of heavy codes and turn arround wich keep it difficult to maintain,
understand and read. Here are your improvements that you are proposing. We want
to write tigthened and integrated code that uses existing structures (CustomTables, etc.) to
ease the process and reduce code. We also want to rely on classes as much as possible instead
of modules, although using modules is not prohibited. Classes should be really compact,
with strictly what is needed for the implementation of features, like for example the LinelisTranslations or the Development classes.

Create new modules in src/modules/showhide and classes in src/classes/showhide.
Tests should be in src/tests/showhide.


Start by splitting the responsibilities into helpers We don't want a lot of them, if
you can factor in this part in the class, that should be the way to go. This keeps the core class focused on orchestration.

Represent show/hide rules with a clear data model (e.g., dedicated ShowHideState/ShowHideRule class). Store flags such as isCRF, isPrinted, isDictionaryHidden, forceHidden (for formulas, choice_formulas, list auto on CRF/printed) so logic becomes declarative and easier to extend. [done]

Introduce strategy-style helpers for each worksheet type (CRF, HList, printed HList, dictionary). Each implements ApplyRules showHideState, allowing the main class to call the appropriate strategy via applyHlistShowHideLogic and applyDictionaryShowHideLogic without large Select Case blocks. [done]

Extract workbook import/export into a separate service (ShowHideExport). It should take the shared data model and handle serialization so the UI logic is independent of storage. [done]
Normalize repeated column-visibility code into reusable procedures that take a listObject, rule set, and default behaviors. This eliminates duplicated loops and conditional checks sprinkled throughout the class. [notdone]

Add guard/validation methods up front (e.g., ensure target listObject exists, required named ranges available). By failing fast with descriptive errors, the remaining code can assume valid state, reducing defensive checks. [done]

Write unit tests around the newly extracted helpers (especially the rule evaluation and persistence services). With smaller, pure procedures, the current heavy integration logic becomes easier to exercise in isolation and prevent regressions. [done]

Once split, document briefly at the top of each helper why it exists and which worksheet layer it supports—future contributors will no longer need to reverse engineer the entire original class. [done]


As a well skilled VBA developper, you are tasked with building the improvements.
You should follow closely instructions.md, and respects any of the constraints in the
file. You can plan your work and implement progressively, but you must
add a [done] / [notdone] tag to the current list to update on where you are.
