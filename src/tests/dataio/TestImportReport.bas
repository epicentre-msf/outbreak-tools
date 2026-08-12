Attribute VB_Name = "TestImportReport"
Attribute VB_Description = "Unit tests for ImportReport"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for ImportReport")

Option Explicit

'@description
'Drives ImportReport, the four ListObjects a linelist keeps the last import
'report in.
'
'WHAT THESE TESTS ARE GUARDING
'-------------------------------------------------------------------------------
'Two things. The store makes its own tables, which is the lesson the show/hide
'store was written from: a class that assumes something else provisioned its
'table saves nothing and reads nothing, in silence. And what is written survives
'the object, which is what lets F_ImportRep open the last report days after the
'import ran.
'
'ONE WORKBOOK PER MODULE
'-------------------------------------------------------------------------------
'A bare workbook with no __import_rep worksheet at all, so the provisioning is
'driven from nothing. The tests that write put the store back with Clear.
'@depends ImportReport, LLImporter, CustomTest

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ImportReport"

Private Const STORE_SHEET As String = "__import_rep"
Private Const TAB_SHEETS_NOT_IMP As String = "reptab_sheetsNotImp"
Private Const TAB_SHEETS_NOT_TOUCH As String = "reptab_sheetsNotTouch"
Private Const TAB_VARS_NOT_IMP As String = "reptab_varsNotImp"
Private Const TAB_VARS_NOT_TOUCH As String = "reptab_varsNotTouch"


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness and the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()

    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestImportReport"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FixtureWorkbook = NewWorkbook()
        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print the results and drop the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set FixtureWorkbook = Nothing

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Put the application into its test state.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of the test that just ran.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Factory and provisioning
'===============================================================================

'@sub-title Create raises when the workbook argument is Nothing.
'@TestMethod("ImportReport")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim store As ImportReport
    On Error Resume Next
    Set store = ImportReport.Create(Nothing)
    Assert.IsTrue Err.Number <> 0, "Factory should raise for a Nothing workbook"
    Err.Clear
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title The store makes its own worksheet and its own four tables.
'@details
'The whole point of the class. A linelist built by an older generation carries
'none of them, and a class that assumes something else provisioned its table
'saves nothing and reads nothing, in silence.
'@TestMethod("ImportReport")
Public Sub TestTheStoreMakesItsOwnTables()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheStoreMakesItsOwnTables"
    If Not FixtureReady("TestTheStoreMakesItsOwnTables") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Dim sh As Worksheet

    Set store = ImportReport.Create(FixtureWorkbook)
    Set sh = store.Wksh

    Assert.IsNotNothing sh, "The store binds to a worksheet"
    Assert.AreEqual STORE_SHEET, sh.Name, "And it is the one a linelist keeps"

    Assert.IsTrue TableExists(sh, TAB_SHEETS_NOT_IMP), "The sheets not imported table"
    Assert.IsTrue TableExists(sh, TAB_SHEETS_NOT_TOUCH), "The sheets not touched table"
    Assert.IsTrue TableExists(sh, TAB_VARS_NOT_IMP), "The variables not imported table"
    Assert.IsTrue TableExists(sh, TAB_VARS_NOT_TOUCH), "The variables not touched table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheStoreMakesItsOwnTables", Err.Number, Err.Description
End Sub

'@sub-title Binding a second time leaves the four tables as they are.
'@TestMethod("ImportReport")
Public Sub TestASecondBindingAddsNoTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestASecondBindingAddsNoTable"
    If Not FixtureReady("TestASecondBindingAddsNoTable") Then Exit Sub
    On Error GoTo TestFail

    Dim first As ImportReport
    Dim second As ImportReport
    Dim countAfterFirst As Long

    Set first = ImportReport.Create(FixtureWorkbook)
    countAfterFirst = first.Wksh.ListObjects.Count

    Set second = ImportReport.Create(FixtureWorkbook)

    Assert.AreEqual countAfterFirst, CLng(second.Wksh.ListObjects.Count), _
                    "Provisioning an already provisioned worksheet adds nothing"
    Assert.AreEqual CLng(4), CLng(second.Wksh.ListObjects.Count), _
                    "And the worksheet carries exactly the four tables"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASecondBindingAddsNoTable", Err.Number, Err.Description
End Sub

'@sub-title The worksheet is left very hidden.
'@details
'A user never opens it; the form reads the tables. Activating a very hidden
'worksheet raises, so nothing in the class selects or activates.
'@TestMethod("ImportReport")
Public Sub TestTheStoreWorksheetStaysOutOfSight()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheStoreWorksheetStaysOutOfSight"
    If Not FixtureReady("TestTheStoreWorksheetStaysOutOfSight") Then Exit Sub
    On Error GoTo TestFail

    Assert.AreEqual CLng(xlSheetVeryHidden), _
                    CLng(ImportReport.Create(FixtureWorkbook).Wksh.Visible), _
                    "The store worksheet is very hidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheStoreWorksheetStaysOutOfSight", Err.Number, Err.Description
End Sub


'@section Writing and reading back
'===============================================================================

'@sub-title A list of worksheet names is written and read back.
'@TestMethod("ImportReport")
Public Sub TestSheetNamesGoInAndComeBack()
    CustomTestSetTitles Assert, TESTMODULE, "TestSheetNamesGoInAndComeBack"
    If Not FixtureReady("TestSheetNamesGoInAndComeBack") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Dim written As BetterArray
    Dim readBack As BetterArray

    Set store = ImportReport.Create(FixtureWorkbook)
    store.Clear

    Set written = New BetterArray
    written.LowerBound = 1
    written.Push "one-sheet", "two-sheet", "three-sheet"

    store.SaveSheets ImportReportNotImported, written
    Set readBack = store.SheetNames(ImportReportNotImported)

    Assert.AreEqual CLng(3), readBack.Length, "Every name written is read back"
    Assert.IsTrue readBack.Includes("one-sheet"), "The first of them"
    Assert.IsTrue readBack.Includes("three-sheet"), "And the last of them"

    store.Clear

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSheetNamesGoInAndComeBack", Err.Number, Err.Description
End Sub

'@sub-title A list of variables and their sheets is written and read back.
'@TestMethod("ImportReport")
Public Sub TestVariableEntriesGoInAndComeBack()
    CustomTestSetTitles Assert, TESTMODULE, "TestVariableEntriesGoInAndComeBack"
    If Not FixtureReady("TestVariableEntriesGoInAndComeBack") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Dim written As BetterArray
    Dim readBack As BetterArray
    Dim entry As Variant

    Set store = ImportReport.Create(FixtureWorkbook)
    store.Clear

    Set written = New BetterArray
    written.LowerBound = 1
    written.Push Array("date_h2", "hlist2D-sheet1")
    written.Push Array("int_h2", "hlist2D-sheet1")

    store.SaveVariables ImportReportNotTouched, written
    Set readBack = store.VariableEntries(ImportReportNotTouched)

    Assert.AreEqual CLng(2), readBack.Length, "Both entries are read back"

    'BetterArray.Push rebases an inner array to the outer LowerBound, so the
    'pair is read through LBound rather than from 0.
    entry = readBack.Item(readBack.LowerBound)
    Assert.AreEqual "date_h2", CStr(entry(LBound(entry))), _
                    "The variable is in the first column"
    Assert.AreEqual "hlist2D-sheet1", CStr(entry(LBound(entry) + 1)), _
                    "And its sheet in the second"

    store.Clear

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestVariableEntriesGoInAndComeBack", Err.Number, Err.Description
End Sub

'@sub-title The two scopes write into two different tables.
'@TestMethod("ImportReport")
Public Sub TestTheTwoScopesDoNotShareATable()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheTwoScopesDoNotShareATable"
    If Not FixtureReady("TestTheTwoScopesDoNotShareATable") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Dim notImported As BetterArray
    Dim notTouched As BetterArray

    Set store = ImportReport.Create(FixtureWorkbook)
    store.Clear

    Set notImported = New BetterArray
    notImported.LowerBound = 1
    notImported.Push "from-the-file"

    Set notTouched = New BetterArray
    notTouched.LowerBound = 1
    notTouched.Push "from-the-linelist"

    store.SaveSheets ImportReportNotImported, notImported
    store.SaveSheets ImportReportNotTouched, notTouched

    Assert.IsTrue store.SheetNames(ImportReportNotImported).Includes("from-the-file"), _
                  "The first scope holds its own name"
    Assert.IsFalse store.SheetNames(ImportReportNotImported).Includes("from-the-linelist"), _
                   "And none of the other's"
    Assert.IsTrue store.SheetNames(ImportReportNotTouched).Includes("from-the-linelist"), _
                  "And the second scope holds its own"

    store.Clear

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheTwoScopesDoNotShareATable", Err.Number, Err.Description
End Sub

'@sub-title What was written survives the object that wrote it.
'@details
'The reason the report is a worksheet and not four arrays. The four arrays it
'replaced died with the import object at the end of the import, and the message
'box asking whether the user wanted a report was shown with an OK button, so
'there was nothing to answer and nothing behind it.
'@TestMethod("ImportReport")
Public Sub TestTheReportSurvivesTheObjectThatWroteIt()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheReportSurvivesTheObjectThatWroteIt"
    If Not FixtureReady("TestTheReportSurvivesTheObjectThatWroteIt") Then Exit Sub
    On Error GoTo TestFail

    Dim writer As ImportReport
    Dim reader As ImportReport
    Dim written As BetterArray

    Set writer = ImportReport.Create(FixtureWorkbook)
    writer.Clear

    Set written = New BetterArray
    written.LowerBound = 1
    written.Push "kept-on-the-worksheet"

    writer.SaveSheets ImportReportNotTouched, written
    Set writer = Nothing

    Set reader = ImportReport.Create(FixtureWorkbook)

    Assert.IsTrue reader.SheetNames(ImportReportNotTouched).Includes("kept-on-the-worksheet"), _
                  "A second store over the same workbook reads what the first wrote"

    reader.Clear

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheReportSurvivesTheObjectThatWroteIt", Err.Number, Err.Description
End Sub

'@sub-title Writing a second time replaces what was there.
'@details
'The store describes the last import rather than every import that ever ran.
'@TestMethod("ImportReport")
Public Sub TestASecondWriteReplacesTheFirst()
    CustomTestSetTitles Assert, TESTMODULE, "TestASecondWriteReplacesTheFirst"
    If Not FixtureReady("TestASecondWriteReplacesTheFirst") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Dim firstRun As BetterArray
    Dim secondRun As BetterArray

    Set store = ImportReport.Create(FixtureWorkbook)
    store.Clear

    Set firstRun = New BetterArray
    firstRun.LowerBound = 1
    firstRun.Push "old-one", "old-two", "old-three"

    Set secondRun = New BetterArray
    secondRun.LowerBound = 1
    secondRun.Push "new-one"

    store.SaveSheets ImportReportNotImported, firstRun
    store.SaveSheets ImportReportNotImported, secondRun

    Assert.AreEqual CLng(1), store.SheetNames(ImportReportNotImported).Length, _
                    "The second write leaves one row"
    Assert.IsFalse store.SheetNames(ImportReportNotImported).Includes("old-two"), _
                   "And nothing of the run before it"

    store.Clear

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASecondWriteReplacesTheFirst", Err.Number, Err.Description
End Sub


'@section Emptiness
'===============================================================================

'@sub-title Clear empties every table and HasEntries says so.
'@TestMethod("ImportReport")
Public Sub TestClearEmptiesEveryTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestClearEmptiesEveryTable"
    If Not FixtureReady("TestClearEmptiesEveryTable") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Dim written As BetterArray

    Set store = ImportReport.Create(FixtureWorkbook)

    Set written = New BetterArray
    written.LowerBound = 1
    written.Push "something"

    store.SaveSheets ImportReportNotImported, written
    Assert.IsTrue store.HasEntries, "The store holds something"

    store.Clear

    Assert.IsFalse store.HasEntries, "And nothing after Clear"
    Assert.AreEqual CLng(0), store.SheetNames(ImportReportNotImported).Length, _
                    "Every table is empty"
    Assert.AreEqual CLng(0), store.VariableEntries(ImportReportNotTouched).Length, _
                    "The variable tables included"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestClearEmptiesEveryTable", Err.Number, Err.Description
End Sub

'@sub-title An unknown scope answers an empty list on both sides.
'@TestMethod("ImportReport")
Public Sub TestAnUnknownScopeAnswersEmpty()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnUnknownScopeAnswersEmpty"
    If Not FixtureReady("TestAnUnknownScopeAnswersEmpty") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Set store = ImportReport.Create(FixtureWorkbook)

    Assert.AreEqual CLng(0), store.SheetNames(99).Length, _
                    "An unknown scope names no table and answers an empty list"
    Assert.AreEqual CLng(0), store.VariableEntries(99).Length, _
                    "And so does the variable side"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnUnknownScopeAnswersEmpty", Err.Number, Err.Description
End Sub

'@sub-title Writing an empty list leaves the table empty and raises nothing.
'@TestMethod("ImportReport")
Public Sub TestWritingAnEmptyListIsSafe()
    CustomTestSetTitles Assert, TESTMODULE, "TestWritingAnEmptyListIsSafe"
    If Not FixtureReady("TestWritingAnEmptyListIsSafe") Then Exit Sub
    On Error GoTo TestFail

    Dim store As ImportReport
    Dim emptyList As BetterArray

    Set store = ImportReport.Create(FixtureWorkbook)
    store.Clear

    'Named emptyList rather than empty: Empty is the VBA literal, and a variable
    'of that name is a compile error at its declaration.
    Set emptyList = New BetterArray
    emptyList.LowerBound = 1

    store.SaveSheets ImportReportNotImported, emptyList
    store.SaveVariables ImportReportNotImported, emptyList

    Assert.AreEqual CLng(0), store.SheetNames(ImportReportNotImported).Length, _
                    "An import that found nothing writes nothing"
    Assert.IsFalse store.HasEntries, "And the store stays empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWritingAnEmptyListIsSafe", Err.Number, Err.Description
End Sub


'@section Fixture helpers
'===============================================================================

'@fun-title Report a fixture that could not be built as this test's failure.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is usable.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 And Not FixtureWorkbook Is Nothing Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@fun-title Whether one worksheet carries a ListObject of that name.
'@param sh Worksheet. The worksheet to look on.
'@param tableName String. The table to look for.
'@return Boolean. True when it is there.
Private Function TableExists(ByVal sh As Worksheet, ByVal tableName As String) As Boolean
    Dim lo As ListObject

    On Error Resume Next
    Set lo = sh.ListObjects(tableName)
    On Error GoTo 0

    TableExists = (Not lo Is Nothing)
End Function
