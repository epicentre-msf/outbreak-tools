Attribute VB_Name = "TestHiddenNames"
Attribute VB_Description = "Regression tests for HiddenNames worksheet name manager"

Option Explicit

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TEST_SHEET_NAME As String = "hn_main"
Private Const OTHER_SHEET_NAME As String = "hn_other"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private testSh As Worksheet
Private otherSh As Worksheet
Private manager As IHiddenNames


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestHiddenNames"
    ResetSheets
End Sub

'@ModuleCleanup
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
Private Sub TestInitialize()
    ResetSheets
    ReleaseManager
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    ReleaseManager
    ResetSheets
End Sub


'@section Helper routines
'===============================================================================
Private Sub ResetSheets()
    Set testSh = TestHelpers.EnsureWorksheet(TEST_SHEET_NAME)
    Set otherSh = TestHelpers.EnsureWorksheet(OTHER_SHEET_NAME)
    TestHelpers.ClearWorksheet testSh
    TestHelpers.ClearWorksheet otherSh
End Sub

Private Sub RestoreSheets()
    TestHelpers.DeleteWorksheet TEST_SHEET_NAME
    TestHelpers.DeleteWorksheet OTHER_SHEET_NAME
End Sub

Private Sub ReleaseManager()
    Set manager = Nothing
End Sub

Private Function EnsureManager() As IHiddenNames
    If manager Is Nothing Then
        Set manager = HiddenNames.Create(testSh)
    End If
    Set EnsureManager = manager
End Function

Private Function NameDefinition(ByVal sh As Worksheet, ByVal nameId As String) As Name
    On Error Resume Next
        Set NameDefinition = sh.Names(nameId)
    On Error GoTo 0
End Function


'@section Test cases
'===============================================================================
'@TestMethod("HiddenNames")
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
