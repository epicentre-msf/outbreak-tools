You are a high skilled VBA developper with task to improve the current project,
and move from legacy code to actual future proof, bug proof project.

We want to extend our formula class with the possiblity to parse formulas
on groups. We have specific formulas on group allowed. Those are the followings:

- SUMIFS
- COUNTIFS, NIFS (should be the same functions)
- MINIFS
- MAXIFS
- MEDIANIFS
- MEANIFS
And a general GROUP_FUN() where "GROUP" is a tag to notice that it is a group
function and FUN the function where to apply the group.

In General, the formula class should implement (or augment) with these new approaches.

Here are the examples:

Example 1
In the setup, the formula is written SUMIFS(var1, var2, var3). The formula
should check for validity: var1 and var3 should be on the same table; var2 is 
the condition and can be elsewhere. The parsed formula should be
SUMIFS(table1[var3], table1[var1], address(var2)) where table1 is the table of
both variables var1 and var3 and address(var2) is the variable address of var2, with row absolute = False. You 
can easily draw the adress from the LLSheets class like in the ParsedLinelistFormula.

Example 2
In the setup, the formula is written MEANIFS(var1, var2, var3). The formula
should check for validity: var1 and var3 should be on the same table; var2 is 
the condition and can be elsewhere. The parsed formula should be
AVERAGE(IF(table1[var1] = address(var2), table1[var3])) where table1 is the table of
both variables var1 and var3 and address(var2) is the variable address of var2, with row absolute = False. You 
can easily draw the adress from the LLSheets class like in the ParsedLinelistFormula.


Example 3
In the setup, the formula is written COUNTIFS(var1, var2, var3) or NIFS(var1, var2, var3). The formula
should check for validity: var1 and var3 should be on the same table; var2 is 
the condition and can be elsewhere. The parsed formula should be
COUNTIFS(table1[var1], address(var2), table1[var3], "<>") where table1 is the table of
both variables var1 and var3 and address(var2) is the variable address of var2, with row absolute = False. You 
can easily draw the adress from the LLSheets class like in the ParsedLinelistFormula. NIFS should behave the same way as COUNTIFS.


Example 4
In the setup, the formula is written GROUPS_FUN(var1, var2, var3). The formula
should check for validity: var1 and var3 should be on the same table; var2 is 
the condition and can be elsewhere. The parsed formula should be
FUN(IF(table1[var1] = address(var2), table1[var3])) where table1 is the table of
both variables var1 and var3 and address(var2) is the variable address of var2, with row absolute = False. You 
can easily draw the adress from the LLSheets class like in the ParsedLinelistFormula, and FUN is 
an excel formula allowed by formulaData.
You should also populate a isGrouped property that should be "Yes" for grouped function so that .FormulaArray is used instead of .Formula when 
assigning the formula.


Here is a detailed implementation plan for the support of group formula

## Implementation Plan for Group Formula Support

1. **Baseline Audit** [DONE]
   - Review `Formulas.cls` parsing pipeline (`EnsureEvaluation`, tokenisation helpers, linelist/analysis builders) to pinpoint injection points for group-aware logic.
   - Trace how variables, tables, and sheet addresses are retrieved today (`LLVariables`, `LLSheets`, `ParsedLinelistFormula`) to confirm we have the metadata needed for grouped validation/output.
   - Catalogue where consuming code reads formula metadata (e.g., `isValid`, `ParsedLinelistFormula`, existing flags) to understand how and where the new `isGrouped` state must surface.

2. **Extend Formula Metadata Contracts** [DONE]
   - Update `IFormulaData` (and its concrete provider) to register the new group-capable tokens (`SUMIFS`, `COUNTIFS`, `NIFS`, `MINIFS`, `MAXIFS`, `MEDIANIFS`, `MEANIFS`, `GROUP`) and expose any lookup the parser will require (e.g., mapping to canonical Excel aggregation functions).
   - Ensure other modules that use `IFormulaData` remain compatible—add targeted unit/integration checks if the contract changes (new methods, collections, or flags).

3. **Introduce Group Function Detection** [DONE]
   - Inside `Formulas.cls`, create a dedicated helper that inspects the parsed token sequence and recognises when the root function is one of the supported grouped names (including `GROUP_*(...)` wrapping another function).
   - Normalise aliases (`COUNTIFS`/`NIFS`) into a shared internal representation so downstream logic can treat them uniformly.
   - Flag unsupported function names wrapped in `GROUP_*()` early with a descriptive validation error that leverages `formulaData` for the allow-list.

4. **Capture Group Parsing Context** [DONE]
   - Extend the internal `TFormulas` type (and any associated class state) with fields to store: resolved aggregator name, controlling variable tokens (filter vs aggregation targets), resolved table name, and an `isGrouped` Boolean.
   - Populate these fields during evaluation so consumers can access consistent information without recomputing the parse.

5. **Validate Variable/Table Alignment** [DONE]
   - Implement a validation routine that confirms argument arity (exactly three for fixed group functions) and leverages `LLVariables` metadata to verify `var1` and `var3` share the same table.
   - Extend the current Valid method with new validation routine to include checking on grouped formula.
   - Enhance the error messaging path (`invalidReason`) to surface precise guidance when tables differ or a variable is unknown.
   - Allow `var2` to reside on any table but ensure we can still resolve its sheet address using `LLSheets`.

6. **Resolve Addresses with Correct Anchoring** [DONE]
   - Introduce (or reuse) a helper that converts variable names into structured references, ensuring `var2` is converted via `LLSheets.Address` as required.
   - Confirm the helper gracefully handles missing metadata and bubbles coherent errors back into the validation pipeline.

7. **Build Grouped Excel Formula Output** [DONE]
   - Extend `ParsedLinelistFormula` (and any other output-building methods) to construct grouped expressions, using native Excel functions where available (`SUMIFS` and `COUNTIFS`/`NIFS`) and the `FUNCTION(IF(table[var1] = address(var2), table[var3]))` pattern for other aggregators.
   - Map each grouped setup function to the correct Excel aggregator: `SUMIFS → SUMIFS`, `COUNTIFS/NIFS → COUNTIFS`, `MEANIFS → AVERAGE`, `MINIFS → MIN`, `MAXIFS → MAX`, `MEDIANIFS → MEDIAN`, and ensure the generic `GROUP_FUN()` path reuses the validated `FUN`.
   - Guard the transformation so it only executes when validation succeeded; otherwise fall back to current error-handling behaviour.

8. **Expose `isGrouped` to Consumers** [DONE]
   - Add a public read-only property (and, if required, update `IFormulas`) that returns `"Yes"`/`"No"` or a Boolean flag per the requirement for downstream `.FormulaArray` usage.
   - Review writer modules that output formulas to Excel to switch to `.FormulaArray` when `isGrouped` is true; ensure this change is backwards-compatible for non-grouped formulas.

9. **Augment Existing Tests and Add New Coverage** [DONE]
   - Identify current test harnesses (unit tests, integration scripts, workbook-driven checks) and extend them with cases covering each grouped function, invalid table pairings, and the generic `GROUP_FUN()` success/failure paths.
   - Add regression tests that verify `isGrouped` flagging, generated Excel strings, and error messages to protect against future regressions.

10. **Documentation and Developer Guidance**
    - Update developer-facing documentation, in-code XML/annotation comments, and any onboarding materials to describe the new grouped formula capabilities and constraints.
    - Provide sample setup expressions and the resulting Excel translations (mirroring the examples) so future maintainers can quickly understand expected behaviour.

11. **Roll-out Checklist**
    - Perform a targeted manual QA pass in Excel to confirm `FormulaArray` usage works with representative datasets.
    - Communicate any necessary configuration updates (e.g., refreshed `IFormulaData` catalogue) to deployment scripts or runtime configuration owners before releasing the feature.


Your task is to implement this feature by following closely instructions.md. You can skip steps 10 and 11; You don't have to implement
all at once, you can do that progressively (2 to 3 taks), until you are done.  REMEMBER, YOU SHOULD TRACK YOUR PROGRESS IN INSTRUCTIONS.MD AND ADD A [DONE] TAG
HERE IN THE CURRENT FILE FOR TASKS THAT ARE DONE. YOU ARE NOT ALLOWED TO OVERWRITE THE CONTENT OF THIS CURRENT EXCEPT FOR THE [DONE] TAG.

Implement the next steps.
