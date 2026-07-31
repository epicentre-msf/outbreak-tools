Attribute VB_Name = "TestCustomPivotTable"
Attribute VB_Description = "Tests for CustomPivotTable class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for CustomPivotTable class")

Option Explicit

Private Assert As CustomTest
Private FixtureWkb As Workbook
Private PivotSheet As Worksheet

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "CustomPivotTable"
Private Const PIVOT_SHEET_NAME As String = "PivotSheet"
Private Const DATA_TABLE_NAME As String = "TestDataTable"
Private Const SECOND_TABLE_NAME As String = "SecondDataTable"
Private Const FORMAT_SHEET As String = "PivotFormatFixture"

'Geometry the class publishes: first block base 2, title at +2, pivot at +6,
'next base 66 rows on.
Private Const FIRST_TITLE_ROW As Long = 4
Private Const SECOND_TITLE_ROW As Long = 70
Private Const TITLE_COLUMN As Long = 2


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestCustomPivotTable"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWkb = NewWorkbook
    SeedFixture FixtureWkb
    Set PivotSheet = FixtureWkb.Worksheets(PIVOT_SHEET_NAME)
    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWkb Is Nothing Then DeleteWorkbook FixtureWkb
    On Error GoTo 0

    Set PivotSheet = Nothing
    Set FixtureWkb = Nothing
End Sub


'@section Test Fixture Helpers
'===============================================================================

'@sub-title Build a workbook with two ListObject sources and an empty pivot sheet
Private Sub SeedFixture(ByVal targetWkb As Workbook)
    Dim sh As Worksheet

    Set sh = targetWkb.Worksheets.Add
    sh.Name = "DataSheet"
    SeedTable sh, 1, DATA_TABLE_NAME
    SeedTable sh, 5, SECOND_TABLE_NAME

    'Create empty pivot sheet
    Set sh = targetWkb.Worksheets.Add
    sh.Name = PIVOT_SHEET_NAME
End Sub

'@sub-title Write a three-row table at a given column and name it
Private Sub SeedTable(ByVal sh As Worksheet, _
                      ByVal firstColumn As Long, _
                      ByVal tableName As String)
    Dim rng As Range

    sh.Cells(1, firstColumn).Value = "Name"
    sh.Cells(1, firstColumn + 1).Value = "Age"
    sh.Cells(2, firstColumn).Value = "Alice"
    sh.Cells(2, firstColumn + 1).Value = 30
    sh.Cells(3, firstColumn).Value = "Bob"
    sh.Cells(3, firstColumn + 1).Value = 25

    Set rng = sh.Range(sh.Cells(1, firstColumn), sh.Cells(3, firstColumn + 1))
    sh.ListObjects.Add(SourceType:=xlSrcRange, _
                       Source:=rng, _
                       XlListObjectHasHeaders:=xlYes).Name = tableName
End Sub

'@sub-title A HiddenNames manager reading the pivot sheet fresh from the workbook
Private Function PivotNames() As HiddenNames
    Set PivotNames = HiddenNames.Create(PivotSheet)
End Function

'@sub-title A real LLFormat built on a fixture design sheet
Private Function BuildDesign() As LLFormat
    Set BuildDesign = LLFormat.Create(PrepareLLFormatFixture(FORMAT_SHEET, FixtureWkb))
End Function


'@section Factory Tests
'===============================================================================

'@TestMethod("CustomPivotTable")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable
    Set sut = CustomPivotTable.Create(PivotSheet)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Assert.IsTrue sut.Wksh Is PivotSheet, _
                  "Create should bind the instance to the sheet it was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestCreateRejectsNothingSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSheet"
    On Error GoTo ExpectError

    Dim sut As CustomPivotTable
    Set sut = CustomPivotTable.Create(Nothing)

    CustomTestLogFailure Assert, "TestCreateRejectsNothingSheet", , _
                         "Expected error when sheet is Nothing"
    Exit Sub
ExpectError:
    'The number is asserted, not just its presence: this class used to offset
    'every ProjectError by vbObjectError, so it raised a different number than
    'every collaborator for the same condition.
    Assert.AreEqual CStr(CLng(ProjectError.InvalidArgument)), CStr(Err.Number), _
                    "Create(Nothing) should raise a plain ProjectError.InvalidArgument"
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestCreateInitialisesHiddenNames()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateInitialisesHiddenNames"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable
    Dim shNames As HiddenNames

    Set sut = CustomPivotTable.Create(PivotSheet)
    Set shNames = PivotNames()

    Assert.AreEqual "2", shNames.ValueAsString("pivot_output_row"), _
                    "pivot_output_row should start at 2"

    Assert.AreEqual "1", shNames.ValueAsString("pivot_counter"), _
                    "pivot_counter should start at 1"

    Assert.IsTrue shNames.HasName("pivot_has_content"), _
                  "Create should seed pivot_has_content"

    Assert.IsTrue shNames.HasName("pivot_formatted"), _
                  "Create should seed pivot_formatted"

    Assert.IsTrue Not shNames.ValueAsBoolean("pivot_has_content"), _
                  "pivot_has_content should start False"

    Assert.IsTrue Not shNames.ValueAsBoolean("pivot_formatted"), _
                  "pivot_formatted should start False"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateInitialisesHiddenNames", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestCreateResumesFromCurrentPosition()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateResumesFromCurrentPosition"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"

    'A second Create over the same sheet must not reset the cursor
    Set sut = CustomPivotTable.Create(PivotSheet)

    Assert.AreEqual "68", PivotNames().ValueAsString("pivot_output_row"), _
                    "Re-wrapping a populated sheet should keep the current position"

    Assert.AreEqual "2", PivotNames().ValueAsString("pivot_counter"), _
                    "Re-wrapping a populated sheet should keep the current counter"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateResumesFromCurrentPosition", Err.Number, Err.Description
End Sub


'@section Add Tests
'===============================================================================

'@TestMethod("CustomPivotTable")
Public Sub TestAddCreatesPivotTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddCreatesPivotTable"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable
    Dim pt As PivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"

    Assert.IsTrue PivotSheet.PivotTables.Count > 0, _
                  "Add should create at least one PivotTable on the sheet"

    On Error Resume Next
    Set pt = PivotSheet.PivotTables("PivotTable_" & DATA_TABLE_NAME)
    On Error GoTo 0

    Assert.IsTrue Not pt Is Nothing, _
                  "PivotTable should be named PivotTable_" & DATA_TABLE_NAME

    Assert.IsTrue Not sut.HasCheckings, _
                  "A pivot that lands should record no checking"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddCreatesPivotTable", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestAddStoresTitleText()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddStoresTitleText"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"

    'The name holds the TITLE TEXT. It used to hold the cell's address, which
    'the linelist then failed to resolve as a range — asserting only that the
    'name exists is what let that ship.
    Assert.AreEqual "Pivot Table 1 - patients", _
                    PivotNames().ValueAsString("pivot_title_" & DATA_TABLE_NAME), _
                    "Add should store the title text under pivot_title_<tableName>"

    Assert.AreEqual "Pivot Table 1 - patients", _
                    CStr(PivotSheet.Cells(FIRST_TITLE_ROW, TITLE_COLUMN).Value), _
                    "The stored title should match what was written to the cell"

    Assert.AreEqual "Pivot Table 1 - patients", sut.Title(DATA_TABLE_NAME), _
                    "Title should answer the stored title for that table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddStoresTitleText", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestAddAdvancesPositionAndCounter()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddAdvancesPositionAndCounter"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"

    Assert.AreEqual "68", PivotNames().ValueAsString("pivot_output_row"), _
                    "pivot_output_row should advance by one block of 66"

    Assert.AreEqual "2", PivotNames().ValueAsString("pivot_counter"), _
                    "pivot_counter should advance to 2 after the first Add"

    Assert.IsTrue PivotNames().ValueAsBoolean("pivot_has_content"), _
                  "pivot_has_content should be True once a pivot has landed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddAdvancesPositionAndCounter", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestSecondAddStacksBelowTheFirst()
    CustomTestSetTitles Assert, TESTMODULE, "TestSecondAddStacksBelowTheFirst"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"
    sut.Add "contacts", SECOND_TABLE_NAME, "Pivot Table"

    Assert.AreEqual "Pivot Table 2 - contacts", _
                    CStr(PivotSheet.Cells(SECOND_TITLE_ROW, TITLE_COLUMN).Value), _
                    "The second title should land 66 rows below the first"

    Assert.AreEqual "Pivot Table 2 - contacts", _
                    PivotNames().ValueAsString("pivot_title_" & SECOND_TABLE_NAME), _
                    "The second block should store its own title"

    Assert.AreEqual "Pivot Table 1 - patients", _
                    PivotNames().ValueAsString("pivot_title_" & DATA_TABLE_NAME), _
                    "The first block's title should survive the second Add"

    Assert.AreEqual "134", PivotNames().ValueAsString("pivot_output_row"), _
                    "pivot_output_row should advance one block per Add"

    Assert.AreEqual "3", PivotNames().ValueAsString("pivot_counter"), _
                    "pivot_counter should advance one per Add"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSecondAddStacksBelowTheFirst", Err.Number, Err.Description
End Sub

'THE FAILURE PATH OF Add HAS NO TEST, AND IT CANNOT HAVE ONE HERE.
'===============================================================================
'The obvious way to reach the PivotFailed handler is a SourceData naming no
'ListObject. Excel answers that with a MODAL DIALOG rather than a trappable
'error, so `On Error` never sees it, the headless run hangs on the dialog, and
'the harness writes a 0-byte test-results.csv — which reads exactly like an
'infinite loop. Cost one run on 2026-07-31. Do not add that test back.
'
'The duplicate-source case does not reach the handler either: Excel accepts a
'second pivot over the same table (see the test below). What the handler is
'left guarding is whatever Excel raises rather than displays, which nothing in
'the fixture can provoke. The happy path asserts HasCheckings is False, so a
'handler firing when it should not would still be caught.

'@TestMethod("CustomPivotTable")
Public Sub TestAddingTheSameTableTwiceLetsTheLaterBlockWin()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddingTheSameTableTwiceLetsTheLaterBlockWin"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"
    sut.Add "patients again", DATA_TABLE_NAME, "Pivot Table"

    'Excel does NOT refuse a second pivot over the same source — it renames the
    'PivotTable itself and builds the block. So a rebuild over a populated sheet
    'stacks a second block for that table and the stored title follows the newer
    'one. Pinned here because it reads like a failure case and is not one.
    Assert.IsTrue Not sut.HasCheckings, _
                  "A second pivot over the same table should not be treated as a failure"

    Assert.AreEqual "Pivot Table 2 - patients again", _
                    PivotNames().ValueAsString("pivot_title_" & DATA_TABLE_NAME), _
                    "The later block's title should win for that table"

    Assert.AreEqual "Pivot Table 2 - patients again", _
                    CStr(PivotSheet.Cells(SECOND_TITLE_ROW, TITLE_COLUMN).Value), _
                    "The second block should still be written a full block below"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddingTheSameTableTwiceLetsTheLaterBlockWin", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestAddRejectsUseWithoutCreate()
    CustomTestSetTitles Assert, TESTMODULE, "TestAddRejectsUseWithoutCreate"
    On Error GoTo ExpectError

    'The predeclared instance holds no worksheet
    CustomPivotTable.Add "patients", DATA_TABLE_NAME, "Pivot Table"

    CustomTestLogFailure Assert, "TestAddRejectsUseWithoutCreate", , _
                         "Expected error when Add is called without Create"
    Exit Sub
ExpectError:
    Assert.AreEqual CStr(CLng(ProjectError.ObjectNotInitialized)), CStr(Err.Number), _
                    "Add without Create should raise ObjectNotInitialized"
End Sub


'@section Format Tests
'===============================================================================

'@TestMethod("CustomPivotTable")
Public Sub TestFormatHidesSheetWithNoPivot()
    CustomTestSetTitles Assert, TESTMODULE, "TestFormatHidesSheetWithNoPivot"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Format BuildDesign()

    Assert.AreEqual CStr(CLng(xlSheetVeryHidden)), CStr(CLng(PivotSheet.Visible)), _
                    "A pivot sheet that took no pivot should stay very hidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormatHidesSheetWithNoPivot", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestFormatShowsSheetHoldingAPivot()
    CustomTestSetTitles Assert, TESTMODULE, "TestFormatShowsSheetHoldingAPivot"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable

    Set sut = CustomPivotTable.Create(PivotSheet)
    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"
    sut.Format BuildDesign()

    'This is the whole point of the class reaching the user: the sheet was
    'always left very hidden while LLFormat looked for a cell nobody wrote.
    Assert.AreEqual CStr(CLng(xlSheetVisible)), CStr(CLng(PivotSheet.Visible)), _
                    "A pivot sheet holding a pivot should be visible"

    Assert.IsTrue PivotNames().ValueAsBoolean("pivot_formatted"), _
                  "Format should record that the cosmetic pass has run"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormatShowsSheetHoldingAPivot", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestFormatRunsCosmeticPassOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestFormatRunsCosmeticPassOnce"
    On Error GoTo TestFail

    Dim sut As CustomPivotTable
    Dim design As LLFormat

    Set sut = CustomPivotTable.Create(PivotSheet)
    Set design = BuildDesign()

    sut.Add "patients", DATA_TABLE_NAME, "Pivot Table"
    sut.Format design

    'The second call short-circuits on pivot_formatted, so the whole-sheet
    'row-height write happens once however many data-entry sheets are built.
    sut.Add "contacts", SECOND_TABLE_NAME, "Pivot Table"
    sut.Format design

    Assert.AreEqual CStr(CLng(xlSheetVisible)), CStr(CLng(PivotSheet.Visible)), _
                    "The sheet should stay visible across repeated Format calls"

    Assert.IsTrue Not sut.HasCheckings, _
                  "Repeated Format calls should record no checking"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestFormatRunsCosmeticPassOnce", Err.Number, Err.Description
End Sub

'@TestMethod("CustomPivotTable")
Public Sub TestFormatRejectsNothingDesign()
    CustomTestSetTitles Assert, TESTMODULE, "TestFormatRejectsNothingDesign"
    On Error GoTo ExpectError

    Dim sut As CustomPivotTable
    Set sut = CustomPivotTable.Create(PivotSheet)

    sut.Format Nothing

    CustomTestLogFailure Assert, "TestFormatRejectsNothingDesign", , _
                         "Expected error when design is Nothing"
    Exit Sub
ExpectError:
    Assert.AreEqual CStr(CLng(ProjectError.InvalidArgument)), CStr(Err.Number), _
                    "Format(Nothing) should raise ProjectError.InvalidArgument"
End Sub
