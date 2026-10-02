Attribute VB_Name = "TestHiddenNames"
Attribute VB_Description = "Regression tests for HiddenNames worksheet name manager"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TEST_SHEET_NAME As String = "hn_main"
Private Const OTHER_SHEET_NAME As String = "hn_other"
Private Const WORKBOOK_SCOPE_NAME As String = "__hn_workbook_scope__"
Private Const WORKBOOK_HEADER_NAME As String = "__hn_table_header__"

'@Folder("CustomTests")
'@ModuleDescription("Regression tests for HiddenNames worksheet name manager")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Validates the HiddenNames class, which persists typed key-value pairs as
'hidden Excel Name definitions scoped to a worksheet or workbook. Tests
'cover factory guard clauses (Nothing raises ObjectNotInitialized),
'workbook-scoped name creation and update, export and import of names
'between worksheets and workbooks (with overwrite flag semantics),
'SetListObjectHeader for binding workbook names to table column references,
'CRUD operations via EnsureName/SetValue/HasName/RemoveName, Value with
'default fallback that avoids side-effects, typed round-trips for String
'(including embedded-quote encoding), Boolean, and Long values, ListNames
'metadata retrieval, and prefix-based filtering of listed names.
'
'Two more groups sit at the end. The first pins the keyed index that holds each
'record's position: a removal shifts every record behind it, and a lookup may
'ask in a different case from the one stored. The second covers QuickValue,
'which reads one value off a host with no instance built.
'
'The fixture allocates two temporary worksheets (hn_main, hn_other) and a
'lazy-loaded manager instance that are reset before every test to guarantee
'full isolation.
'@depends HiddenNames, BetterArray, CustomTest, TestHelpersLite

Private Assert As CustomTest
Private testSh As Worksheet
Private otherSh As Worksheet
Private manager As HiddenNames


'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
'@sub-title Initialise the test harness, suppress UI updates, and prepare fixture sheets
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestHiddenNames"
    ResetSheets
End Sub

'@ModuleCleanup
'@sub-title Print results, tear down sheets, and restore the application state
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    ReleaseManager
    RestoreSheets
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
'@sub-title Reset fixture worksheets and release the manager before each test
Private Sub TestInitialize()
    ResetSheets
    ReleaseManager
End Sub

'@TestCleanup
'@sub-title Flush assertion output, release the manager, and reset sheets after each test
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    ReleaseManager
    ResetSheets
End Sub


'@section Helper routines
'===============================================================================

'@sub-title Ensure both fixture worksheets exist and are cleared of content and names
Private Sub ResetSheets()
    Set testSh = EnsureWorksheet(TEST_SHEET_NAME)
    Set otherSh = EnsureWorksheet(OTHER_SHEET_NAME)
    ClearWorksheet testSh
    ClearWorksheet otherSh
End Sub

'@sub-title Delete the two fixture worksheets from the host workbook
Private Sub RestoreSheets()
    DeleteWorksheet TEST_SHEET_NAME
    DeleteWorksheet OTHER_SHEET_NAME
End Sub

'@sub-title Release the cached HiddenNames manager instance
Private Sub ReleaseManager()
    Set manager = Nothing
End Sub

'@sub-title Create a new empty workbook for cross-workbook export/import tests
Private Function NewTemporaryWorkbook() As Workbook
    Set NewTemporaryWorkbook = NewWorkbook
End Function

'@sub-title Close and delete a temporary workbook, swallowing errors on cleanup
Private Sub CloseTemporaryWorkbook(ByRef wb As Workbook)
    On Error Resume Next
        DeleteWorkbook wb
    On Error GoTo 0
    Set wb = Nothing
End Sub

'@sub-title Remove a workbook-scoped Name definition by identifier, swallowing errors if absent
Private Sub DeleteWorkbookName(ByVal nameId As String)
    Dim wb As Workbook

    Set wb = testSh.Parent
    On Error Resume Next
        wb.Names(nameId).Delete
    On Error GoTo 0
End Sub

'@sub-title Lazy-create and return the shared HiddenNames manager scoped to testSh
Private Function EnsureManager() As HiddenNames
    If manager Is Nothing Then
        Set manager = HiddenNames.Create(testSh)
    End If
    Set EnsureManager = manager
End Function

'@sub-title Safely retrieve a worksheet-scoped Name definition, returning Nothing if absent
Private Function NameDefinition(ByVal sh As Worksheet, ByVal nameId As String) As Name
    On Error Resume Next
        Set NameDefinition = sh.Names(nameId)
    On Error GoTo 0
End Function

'@sub-title Build a two-column ListObject on testSh for SetListObjectHeader tests
'@details
'Clears testSh, writes headers "alpha" and "beta" with one data row, removes
'any prior TST_HN_TABLE ListObject, then creates a new ListObject from the
'range A1:B2 and returns it.
Private Function BuildTestListObject() As ListObject
    Dim tableRange As Range
    Dim lo As ListObject

    testSh.Cells.Clear
    testSh.Range("A1").Value = "alpha"
    testSh.Range("B1").Value = "beta"
    testSh.Range("A2").Value = "one"
    testSh.Range("B2").Value = "two"
    Set tableRange = testSh.Range("A1:B2")

    On Error Resume Next
        testSh.ListObjects("TST_HN_TABLE").Delete
    On Error GoTo 0

    Set lo = testSh.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    lo.Name = "TST_HN_TABLE"
    Set BuildTestListObject = lo
End Function


'@section Test cases
'===============================================================================

'@TestMethod("HiddenNames")
'@sub-title Factory guard: Create raises ObjectNotInitialized when passed Nothing
'@details
'Verifies that the HiddenNames.Create factory method rejects a Nothing
'argument by raising ProjectError.ObjectNotInitialized. The test arranges
'no worksheet, acts by calling Create with Nothing inside an error trap,
'and asserts that the trapped error number matches ObjectNotInitialized.
'If no error is raised the test logs a failure explicitly.
Public Sub TestCreateRequiresWorksheet()
    CustomTestSetTitles Assert, "HiddenNames", "CreateRequiresWorksheet"

    On Error GoTo ExpectError
        HiddenNames.Create Nothing
        Assert.LogFailure "Create should raise when worksheet is missing"
        GoTo TestExit
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Create should raise ObjectNotInitialized when worksheet argument is missing"
    Err.Clear
TestExit:
    On Error GoTo 0
End Sub

'@TestMethod("HiddenNames")
'@sub-title Workbook scope: names created via a Workbook argument persist globally
'@details
'Creates a HiddenNames instance scoped to the host Workbook (not a Worksheet),
'ensures a string name, updates its value with SetValue, and asserts that
'ValueAsString returns the updated value. It then looks up the raw Name
'definition in the Workbook.Names collection to confirm the name exists and
'is hidden. Cleanup removes the name via both RemoveName and a direct
'workbook-level deletion to avoid leaking state.
Public Sub TestWorkbookScopeStoresGlobalName()
    CustomTestSetTitles Assert, "HiddenNames", "WorkbookScopeStoresGlobalName"

    Dim names As HiddenNames
    Dim wb As Workbook
    Dim definition As Name

    On Error GoTo UnexpectedError

    Set wb = testSh.Parent
    Set names = HiddenNames.Create(wb)

    names.EnsureName WORKBOOK_SCOPE_NAME, "wb-value", HiddenNameTypeString
    names.SetValue WORKBOOK_SCOPE_NAME, "wb-updated"

    Assert.AreEqual "wb-updated", names.ValueAsString(WORKBOOK_SCOPE_NAME), _
                     "Workbook-scoped HiddenNames should persist values."

    On Error Resume Next
        Set definition = wb.Names(WORKBOOK_SCOPE_NAME)
    On Error GoTo 0
    Assert.IsTrue Not definition Is Nothing, "Workbook scope should create a global hidden name."
    Assert.AreEqual False, definition.Visible, "Workbook-scoped names should remain hidden."

    names.RemoveName WORKBOOK_SCOPE_NAME
    DeleteWorkbookName WORKBOOK_SCOPE_NAME
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    DeleteWorkbookName WORKBOOK_SCOPE_NAME
    CustomTestLogFailure Assert, "TestWorkbookScopeStoresGlobalName", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title ExportNamesToWorkbook copies sheet-scoped names into a target workbook
'@details
'Creates a string name on the fixture manager, updates it, then exports all
'names to a freshly created temporary workbook via ExportNamesToWorkbook.
'A second HiddenNames instance is created against the target workbook and
'asserts that the exported name exists and retains its value. The temporary
'workbook is closed and deleted in both the normal and error paths.
Public Sub TestExportNamesToWorkbookCopiesValues()
    CustomTestSetTitles Assert, "HiddenNames", "TestExportNamesToWorkbookCopiesValues"

    Dim names As HiddenNames
    Dim targetWb As Workbook
    Dim destination As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_export__", "alpha", HiddenNameTypeString
    names.SetValue "__hn_export__", "bravo"

    Set targetWb = NewTemporaryWorkbook()
    names.ExportNamesToWorkbook targetWb

    Set destination = HiddenNames.Create(targetWb)
    Assert.IsTrue destination.HasName("__hn_export__"), "ExportNamesToWorkbook should create the name on the destination workbook."
    Assert.AreEqual "bravo", destination.ValueAsString("__hn_export__"), _
                     "Exported workbook name should keep the stored value."

    CloseTemporaryWorkbook targetWb
    Exit Sub

UnexpectedError:
    CloseTemporaryWorkbook targetWb
    CustomTestLogFailure Assert, "TestExportNamesToWorkbookCopiesValues", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title A name Excel reserves for itself is stored under a safe name
'@details
'Excel writes hidden names of its own -- _xlfn.* for a function it cannot
'resolve, _xleta.* and _xlpm.* for LAMBDA -- whose RefersTo reads "=#NAME?".
'A workbook-scoped Names.Add on one raises 1004, "the syntax of this name
'isn't correct", and an export that carried one took a whole linelist
'generation down with it.
'
'The store keeps the caller's name anyway: the Excel tag is stripped and the
'rest goes under the "__obt_xl_" prefix. The test asserts that the raw
'reserved name is never written, that the caller reads it back under the
'name it gave, that the safe name is what the sheet carries, and that the
'export carries the safe name beside the ordinary one.
Public Sub TestExcelReservedNamesAreRewritten()
    CustomTestSetTitles Assert, "HiddenNames", "TestExcelReservedNamesAreRewritten"

    Dim names As HiddenNames
    Dim targetWb As Workbook
    Dim destination As HiddenNames
    Dim rawDefinition As Name

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_real__", "kept", HiddenNameTypeString
    names.EnsureName "_xlfn.SINGLE", "stored safe", HiddenNameTypeString

    Set rawDefinition = NameDefinition(testSh, "_xlfn.SINGLE")
    Assert.IsTrue (rawDefinition Is Nothing), _
                  "The raw reserved name should never be written to the sheet."
    Assert.IsTrue names.HasName("__obt_xl_SINGLE"), _
                  "The reserved name should be stored under the safe prefix."
    Assert.IsTrue names.HasName("_xlfn.SINGLE"), _
                  "The caller should find the name under the identifier it gave."
    Assert.AreEqual "stored safe", names.ValueAsString("_xlfn.SINGLE"), _
                     "The caller should read the value back under the identifier it gave."

    Set targetWb = NewTemporaryWorkbook()
    names.ExportNamesToWorkbook targetWb

    Set destination = HiddenNames.Create(targetWb)
    Assert.IsTrue destination.HasName("__hn_real__"), _
                  "The export should still carry the ordinary name."
    Assert.IsTrue destination.HasName("__obt_xl_SINGLE"), _
                  "The export should carry the safe name."
    Assert.IsTrue (NameDefinition(targetWb.Worksheets(1), "_xlfn.SINGLE") Is Nothing), _
                  "The export should never write the raw reserved name."

    CloseTemporaryWorkbook targetWb
    Exit Sub

UnexpectedError:
    CloseTemporaryWorkbook targetWb
    CustomTestLogFailure Assert, "TestExcelReservedNamesAreRewritten", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title ImportNamesFromWorkbook honours the overwriteExisting flag
'@details
'Creates a Long name on the fixture manager, then creates a separate
'workbook containing the same name with a different value. Calls
'ImportNamesFromWorkbook with overwriteExisting:=False and asserts the
'original value is preserved. Then calls again with overwriteExisting:=True
'and asserts the value is updated to the source workbook value. This
'validates both branches of the overwrite flag for cross-workbook imports.
Public Sub TestImportNamesFromWorkbookRespectsOverwrite()
    CustomTestSetTitles Assert, "HiddenNames", "TestImportNamesFromWorkbookRespectsOverwrite"

    Dim target As HiddenNames
    Dim sourceWb As Workbook
    Dim sourceStore As HiddenNames

    On Error GoTo UnexpectedError

    Set target = EnsureManager()
    target.EnsureName "__hn_import__", 5, HiddenNameTypeLong

    Set sourceWb = NewTemporaryWorkbook()
    Set sourceStore = HiddenNames.Create(sourceWb)
    sourceStore.EnsureName "__hn_import__", 42, HiddenNameTypeLong
    sourceStore.SetValue "__hn_import__", 42

    target.ImportNamesFromWorkbook sourceWb, overwriteExisting:=False
    Assert.AreEqual 5, target.ValueAsLong("__hn_import__"), _
                     "ImportNamesFromWorkbook should preserve values when overwriteExisting is False."

    target.ImportNamesFromWorkbook sourceWb, overwriteExisting:=True
    Assert.AreEqual 42, target.ValueAsLong("__hn_import__"), _
                     "ImportNamesFromWorkbook should update values when overwriteExisting is True."

    CloseTemporaryWorkbook sourceWb
    Exit Sub

UnexpectedError:
    CloseTemporaryWorkbook sourceWb
    CustomTestLogFailure Assert, "TestImportNamesFromWorkbookRespectsOverwrite", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title SetListObjectHeader creates a workbook-scoped Name referencing a table column
'@details
'Builds a two-column ListObject on testSh, creates a workbook-scoped
'HiddenNames instance, and calls SetListObjectHeader to bind a workbook
'Name to the "alpha" column. Asserts the Name exists and its RefersTo
'formula matches the expected structured reference (=TableName[alpha]).
'Then re-calls SetListObjectHeader with "beta" and asserts the RefersTo
'formula is overwritten, verifying that the method supports updating an
'existing workbook Name to a different column.
Public Sub TestSetListObjectHeaderCreatesWorkbookName()
    CustomTestSetTitles Assert, "HiddenNames", "TestSetListObjectHeaderCreatesWorkbookName"

    Dim names As HiddenNames
    Dim lo As ListObject
    Dim targetBook As Workbook
    Dim createdName As Name
    Dim expectedRefersTo As String

    On Error GoTo UnexpectedError

    Set lo = BuildTestListObject()
    Set names = HiddenNames.Create(testSh.Parent)
    expectedRefersTo = "=" & lo.Name & "[alpha]"

    names.SetListObjectHeader WORKBOOK_HEADER_NAME, lo, "alpha"

    Set targetBook = testSh.Parent
    Set createdName = targetBook.Names(WORKBOOK_HEADER_NAME)
    Assert.IsTrue Not createdName Is Nothing, "Workbook name should exist after SetListObjectHeader."
    Assert.AreEqual expectedRefersTo, createdName.RefersTo, "Workbook name should reference the table header."

    names.SetListObjectHeader WORKBOOK_HEADER_NAME, lo, "beta"
    expectedRefersTo = "=" & lo.Name & "[beta]"
    Assert.AreEqual expectedRefersTo, targetBook.Names(WORKBOOK_HEADER_NAME).RefersTo, _
                     "SetListObjectHeader should overwrite existing workbook names."

    DeleteWorkbookName WORKBOOK_HEADER_NAME
    Exit Sub

UnexpectedError:
    DeleteWorkbookName WORKBOOK_HEADER_NAME
    CustomTestLogFailure Assert, "TestSetListObjectHeaderCreatesWorkbookName", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title EnsureName creates a hidden, sheet-scoped Name definition
'@details
'Calls EnsureName with a Long default value, then inspects the raw
'worksheet Name definition to confirm it exists and is hidden. Also
'verifies that ValueAsLong returns the initial value and that HasName
'reports the name as present. This test validates the full create path
'of the CRUD lifecycle.
Public Sub TestEnsureNameCreatesDefinition()
    CustomTestSetTitles Assert, "HiddenNames", "EnsureNameCreatesDefinition"

    Dim names As HiddenNames
    Dim definition As Name

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_counter__", 7, HiddenNameTypeLong

    Set definition = NameDefinition(testSh, "__hn_counter__")
    Assert.IsTrue Not definition Is Nothing, "EnsureName should create a sheet-scoped name"
    Assert.AreEqual False, definition.Visible, "Created name should be hidden"
    Assert.AreEqual 7, names.ValueAsLong("__hn_counter__"), "ValueAsLong should return the stored long value"
    Assert.IsTrue names.HasName("__hn_counter__"), "HasName should report the ensured name"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestEnsureNameCreatesDefinition", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title SetValue updates the stored value of an existing Name definition
'@details
'Ensures a Long name with initial value 1, then calls SetValue to change
'it to 42. Asserts that ValueAsLong returns 42, confirming that SetValue
'overwrites the previously stored value without creating a duplicate
'definition.
Public Sub TestSetValueUpdatesExistingDefinition()
    CustomTestSetTitles Assert, "HiddenNames", "SetValueUpdatesExistingDefinition"

    Dim names As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_total__", 1, HiddenNameTypeLong
    names.SetValue "__hn_total__", 42&

    Assert.AreEqual 42&, names.ValueAsLong("__hn_total__"), _
                     "SetValue should update the stored long value"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestSetValueUpdatesExistingDefinition", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title Value with default returns the fallback without creating a Name
'@details
'Calls Value on a name that has never been ensured, passing 99 as the
'default. Asserts that the returned value equals 99, then asserts that
'HasName returns False, confirming that merely reading with a default
'does not have the side-effect of creating a Name definition. This is
'important for read-only queries that should not mutate state.
Public Sub TestValueWithDefaultDoesNotCreateName()
    CustomTestSetTitles Assert, "HiddenNames", "ValueWithDefaultDoesNotCreateName"

    Dim names As HiddenNames
    Dim defaultValue As Variant

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    defaultValue = names.Value("__hn_missing__", 99&)

    Assert.AreEqual 99&, defaultValue, "Value should return provided default when name is absent"
    Assert.IsFalse names.HasName("__hn_missing__"), _
                   "Value default retrieval should not create a name definition"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestValueWithDefaultDoesNotCreateName", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title ExportNames copies name definitions from one worksheet to another
'@details
'Ensures a Boolean name on testSh via the fixture manager, sets it to True,
'then exports all names to otherSh using ExportNames. A new HiddenNames
'instance scoped to otherSh asserts that HasName is True, ValueAsBoolean
'returns True, and the raw Name definition on otherSh exists and remains
'hidden. This confirms both the value fidelity and the visibility flag
'during sheet-to-sheet export.
Public Sub TestExportNamesCopiesDefinitions()
    CustomTestSetTitles Assert, "HiddenNames", "ExportNamesCopiesDefinitions"

    Dim source As HiddenNames
    Dim destination As HiddenNames
    Dim definition As Name

    On Error GoTo UnexpectedError

    Set source = EnsureManager()
    source.EnsureName "__hn_export__", True, HiddenNameTypeBoolean
    source.SetValue "__hn_export__", True

    source.ExportNames otherSh

    Set destination = HiddenNames.Create(otherSh)
    Assert.IsTrue destination.HasName("__hn_export__"), "ExportNames should copy name definition to destination sheet"
    Assert.IsTrue destination.ValueAsBoolean("__hn_export__"), "Exported name should retain boolean value"

    Set definition = NameDefinition(otherSh, "__hn_export__")
    Assert.IsTrue Not definition Is Nothing, "Destination worksheet should expose the exported name"
    Assert.AreEqual False, definition.Visible, "Exported name should remain hidden"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestExportNamesCopiesDefinitions", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title ImportNames honours the overwriteExisting flag for sheet-to-sheet import
'@details
'Creates the same Long name on both testSh (value 10) and otherSh (value 25).
'Calls ImportNames with overwriteExisting:=False and asserts the target
'retains 10. Then calls with overwriteExisting:=True and asserts the target
'is updated to 25. This validates both branches of the overwrite flag when
'importing between worksheets, as opposed to the cross-workbook variant.
Public Sub TestImportNamesRespectsOverwriteFlag()
    CustomTestSetTitles Assert, "HiddenNames", "ImportNamesRespectsOverwriteFlag"

    Dim target As HiddenNames
    Dim source As HiddenNames

    On Error GoTo UnexpectedError

    Set target = EnsureManager()
    target.EnsureName "__hn_import__", 10, HiddenNameTypeLong

    Set source = HiddenNames.Create(otherSh)
    source.EnsureName "__hn_import__", 25, HiddenNameTypeLong
    source.SetValue "__hn_import__", 25

    target.ImportNames otherSh, overwriteExisting:=False

    Assert.AreEqual 10, target.ValueAsLong("__hn_import__"), _
                     "ImportNames overwriteExisting:=False should preserve existing values"

    target.ImportNames otherSh, overwriteExisting:=True

    Assert.AreEqual 25, target.ValueAsLong("__hn_import__"), _
                     "ImportNames overwriteExisting:=True should update values from source sheet"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestImportNamesRespectsOverwriteFlag", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title ListNames returns a BetterArray of metadata records for stored names
'@details
'Ensures a String name and updates its value, then calls ListNames without
'a prefix filter. Asserts the returned BetterArray is not Nothing and has
'exactly one entry. Inspects the record array to verify the name identifier
'is at index 0, the HiddenNameType at index 1 matches HiddenNameTypeString,
'and the timestamp at index 2 is non-zero. This validates the metadata
'structure returned by ListNames.
Public Sub TestListNamesReturnsMetadata()
    CustomTestSetTitles Assert, "HiddenNames", "ListNamesReturnsMetadata"

    Dim names As HiddenNames
    Dim records As BetterArray
    Dim record As Variant

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_meta__", "sample", HiddenNameTypeString
    names.SetValue "__hn_meta__", "updated"

    Set records = names.ListNames()
    Assert.IsTrue Not records Is Nothing, "ListNames should return a BetterArray instance"
    Assert.AreEqual 1, records.Length, "ListNames should include ensured name metadata"

    record = records.Item(records.LowerBound)
    Assert.AreEqual "__hn_meta__", record(0), "Metadata should expose the name identifier"
    Assert.AreEqual HiddenNameTypeString, record(1), "Metadata should track the value type"
    Assert.IsTrue record(2) <> 0, "Metadata should include a last-updated timestamp"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestListNamesReturnsMetadata", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title String round-trip preserves embedded double-quote characters
'@details
'Stores a string containing an embedded double-quote (beta"quote) via
'EnsureName and SetValue, then retrieves it with ValueAsString. Asserts
'the retrieved value matches the original, confirming that the internal
'quote-encoding and decoding logic does not corrupt or strip embedded
'quote characters during serialisation into the Name RefersTo formula.
Public Sub TestStringValuesDecodeQuotes()
    CustomTestSetTitles Assert, "HiddenNames", "TestStringValuesDecodeQuotes"

    Dim names As HiddenNames
    Dim expected As String

    On Error GoTo UnexpectedError

    expected = "beta""quote"

    Set names = EnsureManager()
    names.EnsureName "__hn_text__", "alpha", HiddenNameTypeString
    names.SetValue "__hn_text__", expected

    Assert.AreEqual expected, names.ValueAsString("__hn_text__"), _
                     "ValueAsString should return the stored text without serialized quotes"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestStringValuesDecodeQuotes", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title Boolean round-trip: False -> True survives EnsureName/SetValue/ValueAsBoolean
'@details
'Ensures a Boolean name with initial value False, updates it to True via
'SetValue, then retrieves it with ValueAsBoolean. Asserts the returned
'value is True, confirming that Boolean values survive the serialisation
'round-trip through the hidden Name definition.
Public Sub TestBooleanRoundTrip()
    CustomTestSetTitles Assert, "HiddenNames", "TestBooleanRoundTrip"

    Dim names As HiddenNames
    Dim stored As Boolean

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_bool__", False, HiddenNameTypeBoolean
    names.SetValue "__hn_bool__", True

    stored = names.ValueAsBoolean("__hn_bool__")
    Assert.IsTrue stored, "ValueAsBoolean should return the stored boolean"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestBooleanRoundTrip", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title Long round-trip: 0 -> 123456 survives EnsureName/SetValue/ValueAsLong
'@details
'Ensures a Long name with initial value 0, updates it to 123456 via
'SetValue, then retrieves it with ValueAsLong. Asserts the returned value
'matches 123456, confirming that Long integer values survive the
'serialisation round-trip through the hidden Name definition.
Public Sub TestLongRoundTrip()
    CustomTestSetTitles Assert, "HiddenNames", "TestLongRoundTrip"

    Dim names As HiddenNames
    Dim stored As Long

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_long__", 0&, HiddenNameTypeLong
    names.SetValue "__hn_long__", 123456&

    stored = names.ValueAsLong("__hn_long__")
    Assert.AreEqual 123456&, stored, "ValueAsLong should return the stored long value"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestLongRoundTrip", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title RemoveName deletes both the metadata entry and the worksheet Name definition
'@details
'Ensures a Long name, then immediately removes it via RemoveName. Asserts
'that HasName returns False (metadata deleted) and that looking up the raw
'Name definition on the worksheet returns Nothing (Excel definition
'deleted). This validates the delete path of the CRUD lifecycle.
Public Sub TestRemoveDeletesDefinition()
    CustomTestSetTitles Assert, "HiddenNames", "TestRemoveDeletesDefinition"

    Dim names As HiddenNames
    Dim definition As Name

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_remove__", 5, HiddenNameTypeLong
    names.RemoveName "__hn_remove__"

    Assert.IsFalse names.HasName("__hn_remove__"), "RemoveName should clear existence from metadata"
    Set definition = NameDefinition(testSh, "__hn_remove__")
    Assert.IsTrue definition Is Nothing, "RemoveName should delete the worksheet definition"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestRemoveDeletesDefinition", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title ListNames with a prefix filter returns only matching entries
'@details
'Ensures three names: two with the prefix "__hn_" and one with "zz_".
'Calls ListNames("__hn_") and asserts the returned BetterArray has
'exactly two entries, confirming that the prefix filter excludes names
'that do not start with the specified string. This validates the optional
'prefix filtering parameter of ListNames.
Public Sub TestListNamesFiltersByPrefix()
    CustomTestSetTitles Assert, "HiddenNames", "TestListNamesFiltersByPrefix"

    Dim names As HiddenNames
    Dim records As BetterArray

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_alpha__", 1, HiddenNameTypeLong
    names.EnsureName "__hn_beta__", 2, HiddenNameTypeLong
    names.EnsureName "zz_skip__", 3, HiddenNameTypeLong

    Set records = names.ListNames("__hn_")
    Assert.AreEqual 2, records.Length, "ListNames should filter entries using the provided prefix"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestListNamesFiltersByPrefix", Err.Number, Err.Description
    Err.Clear
End Sub


'@section The keyed name index
'===============================================================================
'The position of every record is held in a keyed Collection. These tests pin
'the two places that index can go wrong: a removal that shifts every record
'after it, and a lookup whose case differs from the stored name.

'@TestMethod("HiddenNames")
'@sub-title Removing a name leaves every later name answering its own value
'@details
'Four names are stored and the first is removed, so the three behind it each
'move down one place. A reader that kept the positions of the old store would
'answer each of the three with its neighbour's value, and it would do so in
'silence. The values are all different so that a shift shows as a wrong answer
'rather than as a raise.
Public Sub TestRemoveKeepsLaterNamesReachable()
    CustomTestSetTitles Assert, "HiddenNames", "TestRemoveKeepsLaterNamesReachable"

    Dim names As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_pos1__", "one", HiddenNameTypeString
    names.EnsureName "__hn_pos2__", "two", HiddenNameTypeString
    names.EnsureName "__hn_pos3__", "three", HiddenNameTypeString
    names.EnsureName "__hn_pos4__", "four", HiddenNameTypeString

    names.RemoveName "__hn_pos1__"

    Assert.AreEqual "two", names.ValueAsString("__hn_pos2__"), _
                     "The name after the removed one should still answer its own value"
    Assert.AreEqual "three", names.ValueAsString("__hn_pos3__"), _
                     "Every name after the removed one should still answer its own value"
    Assert.AreEqual "four", names.ValueAsString("__hn_pos4__"), _
                     "The last name should still answer its own value after a removal"
    Assert.IsFalse names.HasName("__hn_pos1__"), _
                     "The removed name should be gone from the index"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestRemoveKeepsLaterNamesReachable", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title A write after a removal reaches the name it was given
'@details
'The middle of three names is removed and the last one is then written. This is
'where a stale index does real damage: SetValue resolves a position and writes
'through the definition it finds there, so a position pointing at a neighbour
'overwrites the wrong hidden name. Both survivors are read back.
Public Sub TestSetValueAfterRemoveWritesRightName()
    CustomTestSetTitles Assert, "HiddenNames", "TestSetValueAfterRemoveWritesRightName"

    Dim names As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_wr1__", "first", HiddenNameTypeString
    names.EnsureName "__hn_wr2__", "second", HiddenNameTypeString
    names.EnsureName "__hn_wr3__", "third", HiddenNameTypeString

    names.RemoveName "__hn_wr2__"
    names.SetValue "__hn_wr3__", "changed"

    Assert.AreEqual "changed", names.ValueAsString("__hn_wr3__"), _
                     "SetValue should write the name it was given"
    Assert.AreEqual "first", names.ValueAsString("__hn_wr1__"), _
                     "SetValue should leave every other name alone"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestSetValueAfterRemoveWritesRightName", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title A lookup finds a name whose case differs from the stored one
'@details
'The index is keyed on the lower-cased identifier, which is what keeps the
'case-insensitive matching the scan before it did with StrComp.
Public Sub TestLookupIgnoresNameCase()
    CustomTestSetTitles Assert, "HiddenNames", "TestLookupIgnoresNameCase"

    Dim names As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_MixedCase__", "kept", HiddenNameTypeString

    Assert.IsTrue names.HasName("__hn_mixedcase__"), _
                     "HasName should find a stored name whatever the case asked for"
    Assert.AreEqual "kept", names.ValueAsString("__HN_MIXEDCASE__"), _
                     "ValueAsString should find a stored name whatever the case asked for"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestLookupIgnoresNameCase", Err.Number, Err.Description
    Err.Clear
End Sub


'@section QuickValue
'===============================================================================
'QuickValue reads one stored value off a host with no instance behind it. It
'skips the Names walk Create pays for, so it holds nothing and sees every write
'as it happens.

'@TestMethod("HiddenNames")
'@sub-title QuickValue reads a worksheet-scoped value with no instance built
'@details
'The name is written through an instance and read back through the class
'itself. Nothing is created for the read, which is the whole point of the
'member: the linelist event handlers read sheet_type and table_name this way on
'sheets that hold hundreds of names.
Public Sub TestQuickValueReadsWorksheetName()
    CustomTestSetTitles Assert, "HiddenNames", "TestQuickValueReadsWorksheetName"

    Dim names As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_quick__", "HList", HiddenNameTypeString

    Assert.AreEqual "HList", HiddenNames.QuickValue(testSh, "__hn_quick__"), _
                     "QuickValue should read a worksheet-scoped stored value"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestQuickValueReadsWorksheetName", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title QuickValue undoubles the quotes inside a stored string
'@details
'SerializeValue doubles every quote inside a string before it writes the
'RefersTo formula, so the reader has to undo that. ValueAsString is asserted
'beside it to show both readers answer the same thing.
Public Sub TestQuickValueDecodesQuotes()
    CustomTestSetTitles Assert, "HiddenNames", "TestQuickValueDecodesQuotes"

    Dim names As HiddenNames
    Dim expected As String

    On Error GoTo UnexpectedError

    expected = "beta""quote"

    Set names = EnsureManager()
    names.EnsureName "__hn_qtext__", expected, HiddenNameTypeString

    Assert.AreEqual expected, HiddenNames.QuickValue(testSh, "__hn_qtext__"), _
                     "QuickValue should return the text without its serialized quotes"
    Assert.AreEqual expected, names.ValueAsString("__hn_qtext__"), _
                     "Both readers should answer the same stored text"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestQuickValueDecodesQuotes", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title QuickValue answers a stored number as its digits
'@details
'A Long is written as =123 with no quote wrapper, so the caller gets the digits
'and coerces them. The typed readers stay on the instance.
Public Sub TestQuickValueReadsNumberAsText()
    CustomTestSetTitles Assert, "HiddenNames", "TestQuickValueReadsNumberAsText"

    Dim names As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_qlong__", 0&, HiddenNameTypeLong
    names.SetValue "__hn_qlong__", 4321&

    Assert.AreEqual "4321", HiddenNames.QuickValue(testSh, "__hn_qlong__"), _
                     "QuickValue should answer a stored number as its digits"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestQuickValueReadsNumberAsText", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title QuickValue answers the default when the host holds no such name
'@details
'An absent name is the ordinary answer for a caller reading a tag off a sheet
'that carries none, so it comes back as the default rather than as a raise.
'A Nothing host and an empty identifier answer the same way.
Public Sub TestQuickValueAnswersDefaultWhenAbsent()
    CustomTestSetTitles Assert, "HiddenNames", "TestQuickValueAnswersDefaultWhenAbsent"

    On Error GoTo UnexpectedError

    Assert.AreEqual "none", HiddenNames.QuickValue(testSh, "__hn_qmissing__", "none"), _
                     "QuickValue should answer the default for a name the sheet does not hold"
    Assert.AreEqual vbNullString, HiddenNames.QuickValue(testSh, "__hn_qmissing__"), _
                     "QuickValue should answer an empty string when no default is given"
    Assert.AreEqual "none", HiddenNames.QuickValue(Nothing, "__hn_qmissing__", "none"), _
                     "QuickValue should answer the default for a host that is Nothing"
    Assert.AreEqual "none", HiddenNames.QuickValue(testSh, vbNullString, "none"), _
                     "QuickValue should answer the default for an empty identifier"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestQuickValueAnswersDefaultWhenAbsent", Err.Number, Err.Description
    Err.Clear
End Sub

'@TestMethod("HiddenNames")
'@sub-title QuickValue reads a workbook-scoped value
'@details
'The host is taken as an Object, so a Workbook is read the same way a Worksheet
'is. CustomLinelistFunctions reads its workbook settings through this.
Public Sub TestQuickValueReadsWorkbookName()
    CustomTestSetTitles Assert, "HiddenNames", "TestQuickValueReadsWorkbookName"

    Dim names As HiddenNames
    Dim wb As Workbook

    On Error GoTo UnexpectedError

    Set wb = testSh.Parent
    DeleteWorkbookName WORKBOOK_SCOPE_NAME

    Set names = HiddenNames.Create(wb)
    names.EnsureName WORKBOOK_SCOPE_NAME, "global", HiddenNameTypeString

    Assert.AreEqual "global", HiddenNames.QuickValue(wb, WORKBOOK_SCOPE_NAME), _
                     "QuickValue should read a workbook-scoped stored value"

    names.RemoveName WORKBOOK_SCOPE_NAME
    DeleteWorkbookName WORKBOOK_SCOPE_NAME
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestQuickValueReadsWorkbookName", Err.Number, Err.Description
    Err.Clear
    DeleteWorkbookName WORKBOOK_SCOPE_NAME
End Sub

'@TestMethod("HiddenNames")
'@sub-title QuickValue sees a write made after an instance had already read it
'@details
'QuickValue holds nothing, so it answers what the host says right now. This is
'the contract that lets an event handler read a tag without worrying about how
'stale some other object's cache has become.
Public Sub TestQuickValueSeesLaterWrites()
    CustomTestSetTitles Assert, "HiddenNames", "TestQuickValueSeesLaterWrites"

    Dim names As HiddenNames

    On Error GoTo UnexpectedError

    Set names = EnsureManager()
    names.EnsureName "__hn_qlive__", "before", HiddenNameTypeString

    Assert.AreEqual "before", HiddenNames.QuickValue(testSh, "__hn_qlive__"), _
                     "QuickValue should read the value the name was created with"

    names.SetValue "__hn_qlive__", "after"

    Assert.AreEqual "after", HiddenNames.QuickValue(testSh, "__hn_qlive__"), _
                     "QuickValue should read the value the name holds now"
    On Error GoTo 0
    Exit Sub

UnexpectedError:
    CustomTestLogFailure Assert, "TestQuickValueSeesLaterWrites", Err.Number, Err.Description
    Err.Clear
End Sub
