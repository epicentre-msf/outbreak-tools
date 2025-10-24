We use a lot storage of values at worksheet level that are used for various things
in LLExport to store the totalexportnumber, in some cases to know if the a class
is in a specific mode, etc.

We will probably use it a lot in the future, and we are thinking about
reducing duplication by creating a class in src/classes/general that will
handle this for us:

- Add one or more hidden name to a worksheet, optionally specify the type (a simple AsString boolean) for retrieval
- Retrieve one or more hidden name of the worksheet,
- Export the hidden names to another worksheet
- Update/import the hidden names from a worksheet

On retrieval, use .Value instead of Evaluation as this fails.

What are your plans for such a class? reply bellow

Proposed plan:

1. Define new class `HiddenNames` in `src/classes/general` implementing an `IHiddenNames` interface for testability.
2. Constructor accepts a `Worksheet` reference plus optional validator so unit tests can inject fakes.
3. Provide methods:
   - `EnsureName(nameId As String, initialValue As Variant, Optional valueType As HiddenNameValueType)` to create the name when missing using `.Value`.
   - `Value(nameId As String)` returning the stored value, plus helpers `ValueAsBoolean`, `ValueAsLong`, `ValueAsString`.
   - `SetValue(nameId As String, value As Variant)` to update existing stored values.
   - `HasName(nameId As String)` to validate name existence.
   - `ListNames(Optional prefixFilter As String)` for discovery/export.
   - `ExportNames(targetWorksheet As Worksheet)` and `ImportNames(sourceWorksheet As Worksheet)` to support migration between sheets while preserving type metadata.
4. Store lightweight metadata (name, type, lastUpdated) in a private BetterArray to avoid repeated sheet lookups and simplify export routines.
5. Add error handling: raise descriptive errors when requested names are missing, wrap worksheet failures, and log via existing `ApplicationState` facilities.

Next steps would be to confirm naming conventions, agree on enum values for `HiddenNameValueType`, and prioritize which modules migrate first.


As a well skilled VBA developper, you are tasked with building the improvements.
You should follow closely instructions.md, and respects any of the constraints in the
file. You can plan your work and implement progressively, but you must
add a [done] / [notdone] tag to the current list to update on where you are.
