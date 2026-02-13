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
'The fixture allocates two temporary worksheets (hn_main, hn_other) and a
'lazy-loaded manager instance that are reset before every test to guarantee
'full isolation.
'@depends HiddenNames, IHiddenNames, BetterArray, CustomTest, TestHelpers

Private Assert As ICustomTest
Private testSh As Worksheet
Private otherSh As Worksheet
Private manager As IHiddenNames


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
    Set testSh = TestHelpers.EnsureWorksheet(TEST_SHEET_NAME)
    Set otherSh = TestHelpers.EnsureWorksheet(OTHER_SHEET_NAME)
    TestHelpers.ClearWorksheet testSh
    TestHelpers.ClearWorksheet otherSh
End Sub

'@sub-title Delete the two fixture worksheets from the host workbook
Private Sub RestoreSheets()
    TestHelpers.DeleteWorksheet TEST_SHEET_NAME
    TestHelpers.DeleteWorksheet OTHER_SHEET_NAME
End Sub

'@sub-title Release the cached HiddenNames manager instance
Private Sub ReleaseManager()
    Set manager = Nothing
End Sub

'@sub-title Create a new empty workbook for cross-workbook export/import tests
Private Function NewTemporaryWorkbook() As Workbook
    Set NewTemporaryWorkbook = TestHelpers.NewWorkbook
End Function

'@sub-title Close and delete a temporary workbook, swallowing errors on cleanup
Private Sub CloseTemporaryWorkbook(ByRef wb As Workbook)
    On Error Resume Next
        TestHelpers.DeleteWorkbook wb
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

'@sub-title Lazy-create and return the shared IHiddenNames manager scoped to testSh
Private Function EnsureManager() As IHiddenNames
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

    Dim names As IHiddenNames
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

    Dim names As IHiddenNames
    Dim targetWb As Workbook
    Dim destination As IHiddenNames

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

    Dim target As IHiddenNames
    Dim sourceWb As Workbook
    Dim sourceStore As IHiddenNames

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

    Dim names As IHiddenNames
    Dim lo As ListObject
    Dim workbook As Workbook
    Dim createdName As Name
    Dim expectedRefersTo As String

    On Error GoTo UnexpectedError

    Set lo = BuildTestListObject()
    Set names = HiddenNames.Create(testSh.Parent)
    expectedRefersTo = "=" & lo.Name & "[alpha]"

    names.SetListObjectHeader WORKBOOK_HEADER_NAME, lo, "alpha"

    Set workbook = testSh.Parent
    Set createdName = workbook.Names(WORKBOOK_HEADER_NAME)
    Assert.IsTrue Not createdName Is Nothing, "Workbook name should exist after SetListObjectHeader."
    Assert.AreEqual expectedRefersTo, createdName.RefersTo, "Workbook name should reference the table header."

    names.SetListObjectHeader WORKBOOK_HEADER_NAME, lo, "beta"
    expectedRefersTo = "=" & lo.Name & "[beta]"
    Assert.AreEqual expectedRefersTo, workbook.Names(WORKBOOK_HEADER_NAME).RefersTo, _
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

    Dim names As IHiddenNames
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

    Dim names As IHiddenNames

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

    Dim names As IHiddenNames
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

    Dim source As IHiddenNames
    Dim destination As IHiddenNames
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

    Dim target As IHiddenNames
    Dim source As IHiddenNames

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

    Dim names As IHiddenNames
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

    Dim names As IHiddenNames
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

    Dim names As IHiddenNames
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

    Dim names As IHiddenNames
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

    Dim names As IHiddenNames
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

    Dim names As IHiddenNames
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
