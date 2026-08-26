Attribute VB_Name = "TestMasterSetupLog"
Attribute VB_Description = "Tests for MasterSetupLog class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for MasterSetupLog class")

'@description
'Drives MasterSetupLog, the store of the master setup user log. The class
'binds to the very hidden __log worksheet, builds it when the workbook has
'none, and flushes one Checking bundle per event through CheckingOutput,
'which appends across renders. The suite covers the provisioning, the three
'sections, the append across two flushes, the outcome colours, the rotation
'past the row cap, the separator guard on the detail and the source, and
'the text report with its disease worksheet count. The last two tests
'reach the log the way the master setup modules do, through the one
'EventMasterSetup of the workbook.
'
'THE ROTATION IS REACHED BY SEEDING
'-------------------------------------------------------------------------------
'The cap is ten thousand rows. The test writes one cell past the cap in
'the output column and logs one event, so the rotation runs without the
'suite writing ten thousand lines.
'@depends MasterSetupLog, EventMasterSetup, CustomTest, Checking, HiddenNames

Private Assert As CustomTest
Private FixtureWkb As Workbook

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "MasterSetupLog"
Private Const LOG_SHEET As String = "__log"
Private Const TITLE_COLUMN As Long = 2
Private Const OUTPUT_COLUMN As Long = 3
Private Const DATE_COLUMN As Long = 4
Private Const DETAIL_COLUMN As Long = 5

'The three section titles of the log sheet, held here as the suite reads
'them off the sheet and the class holds them Private.
Private Const SECTION_OPENCLOSE As String = "open/close"
Private Const SECTION_DATAIO As String = "data input/output"
Private Const SECTION_LIFECYCLE As String = "master setup lifecycle"

'@section Lifecycle
'===============================================================================

'@sub-title Set up the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestMasterSetupLog"
End Sub

'@sub-title Print results and tear down.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Build a fresh workbook before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWkb = NewWorkbook()
End Sub

'@sub-title Flush assert state and drop the fixture workbook.
'@TestCleanup
Private Sub TestCleanup()
    'The assertions of a test reach the results sheet only once flushed.
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWkb Is Nothing Then DeleteWorkbook FixtureWkb
    On Error GoTo 0

    Set FixtureWkb = Nothing
End Sub

'@section Test Fixture Helpers
'===============================================================================

'@sub-title Answer the last written row of the output column of a worksheet.
'@param sh Worksheet. The log worksheet.
'@return Long. The last written row of the output column.
Private Function LastOutputRow(ByVal sh As Worksheet) As Long
    LastOutputRow = sh.Cells(sh.Rows.Count, OUTPUT_COLUMN).End(xlUp).Row
End Function

'@sub-title Answer the first row whose cell in a column contains a text.
'@details
'A loop, because Range.Find inherits LookIn and SearchOrder from the last
'search of the Excel session. Answers 0 on a miss.
'@param sh Worksheet. The log worksheet.
'@param columnIndex Long. The column to read.
'@param searched String. The text to look for.
'@return Long. The first matching row, or 0.
Private Function RowOfText(ByVal sh As Worksheet, ByVal columnIndex As Long, _
                           ByVal searched As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = LastOutputRow(sh)
    For rowIndex = 1 To lastRow
        If InStr(1, CStr(sh.Cells(rowIndex, columnIndex).Value), searched, vbTextCompare) > 0 Then
            RowOfText = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

'@sub-title Count the rows of a column that carry a text.
'@param sh Worksheet. The log worksheet.
'@param columnIndex Long. The column to read.
'@param searched String. The text to look for.
'@return Long. How many rows carry it.
Private Function CountOfText(ByVal sh As Worksheet, ByVal columnIndex As Long, _
                             ByVal searched As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim total As Long

    lastRow = LastOutputRow(sh)
    For rowIndex = 1 To lastRow
        If InStr(1, CStr(sh.Cells(rowIndex, columnIndex).Value), searched, vbTextCompare) > 0 Then
            total = total + 1
        End If
    Next rowIndex

    CountOfText = total
End Function

'@sub-title Answer the first report line that contains a text.
'@details
'The report is read by content, because the heading block and the entries
'grow at their own pace. Answers an empty string on a miss.
'@param lines BetterArray. The report lines.
'@param searched String. The text to look for.
'@return String. The first matching line, or an empty string.
Private Function LineOfText(ByVal lines As BetterArray, _
                            ByVal searched As String) As String
    Dim index As Long

    For index = lines.LowerBound To lines.UpperBound
        If InStr(1, CStr(lines.Item(index)), searched, vbTextCompare) > 0 Then
            LineOfText = CStr(lines.Item(index))
            Exit Function
        End If
    Next index
End Function

'@sub-title Add a worksheet to the fixture and tag it as a disease sheet.
'@details
'The tag is the hidden sheetTag name MasterSetupHelpers.IsMasterDiseaseSheet
'reads, written the way TestMasterSetupHelpers writes it.
'@param sheetName String. The name of the new sheet.
'@return Worksheet. The tagged sheet.
Private Function AddDiseaseSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet
    Dim tagStore As HiddenNames

    Set sh = FixtureWkb.Worksheets.Add( _
        After:=FixtureWkb.Worksheets(FixtureWkb.Worksheets.Count))
    sh.Name = sheetName

    Set tagStore = HiddenNames.Create(sh)
    tagStore.EnsureName "sheetTag", "disease", HiddenNameTypeString

    Set AddDiseaseSheet = sh
End Function

'@section Factory Tests
'===============================================================================

'@sub-title A Nothing workbook is refused, and the number says why.
'@TestMethod("MasterSetupLog")
Public Sub TestCreateRejectsNothingWorkbook()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim errNumber As Long

    On Error Resume Next
        Set sut = MasterSetupLog.Create(Nothing)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing workbook is refused and the number names the reason"
    Assert.IsTrue (sut Is Nothing), "Nothing comes back from the refused create"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title A workbook without a log sheet grows one, very hidden.
'@TestMethod("MasterSetupLog")
Public Sub TestCreateBuildsTheLogSheetVeryHidden()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateBuildsTheLogSheetVeryHidden"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet

    Set sut = MasterSetupLog.Create(FixtureWkb)

    Assert.IsTrue (Not sut Is Nothing), "Create gives back an instance"

    Set sh = FixtureWkb.Worksheets(LOG_SHEET)
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(sh.Visible), _
                    "The built sheet is very hidden"
    Assert.AreEqual LOG_SHEET, sut.Wksh().Name, _
                    "The instance is bound to the log sheet"
    Assert.AreEqual LOG_SHEET, sut.SheetName(), _
                    "The class answers the internal sheet name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateBuildsTheLogSheetVeryHidden", Err.Number, Err.Description
End Sub

'@sub-title An existing log sheet is reused and keeps its visibility.
'@TestMethod("MasterSetupLog")
Public Sub TestCreateBindsToAnExistingLogSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateBindsToAnExistingLogSheet"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim countBefore As Long

    Set sh = FixtureWkb.Worksheets.Add
    sh.Name = LOG_SHEET
    countBefore = FixtureWkb.Worksheets.Count

    Set sut = MasterSetupLog.Create(FixtureWkb)

    Assert.AreEqual countBefore, FixtureWkb.Worksheets.Count, _
                    "The existing sheet is reused"
    Assert.IsTrue (sut.Wksh() Is sh), "The instance is bound to the existing sheet"
    Assert.AreEqual CLng(xlSheetVisible), CLng(sh.Visible), _
                    "An existing sheet keeps its visibility"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateBindsToAnExistingLogSheet", Err.Number, Err.Description
End Sub

'@sub-title A fresh workbook grows its log sheet on the first logged event.
'@details
'The acceptance line of the session: nothing is built ahead of time, and
'one logged event leaves the sheet in the workbook with one line on it.
'@TestMethod("MasterSetupLog")
Public Sub TestTheFirstEventGrowsTheLogSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFirstEventGrowsTheLogSheet"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim countBefore As Long

    countBefore = FixtureWkb.Worksheets.Count
    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogInfo "workbook-open", FixtureWkb.Name
    Set sh = sut.Wksh()

    Assert.AreEqual countBefore + 1, FixtureWkb.Worksheets.Count, _
                    "The workbook gained the log sheet"
    Assert.IsTrue (RowOfText(sh, DETAIL_COLUMN, "workbook-open") > 0), _
                  "The first event is on the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFirstEventGrowsTheLogSheet", Err.Number, Err.Description
End Sub

'@section Logging Tests
'===============================================================================

'@sub-title One logged event carries the section, the date, the action and the detail.
'@TestMethod("MasterSetupLog")
Public Sub TestLogSuccessWritesActionDateAndDetail()
    CustomTestSetTitles Assert, TESTMODULE, "TestLogSuccessWritesActionDateAndDetail"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim titleRow As Long
    Dim entryRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "export-setup", "measles.xlsb"
    Set sh = sut.Wksh()

    titleRow = RowOfText(sh, OUTPUT_COLUMN, SECTION_DATAIO)
    Assert.IsTrue (titleRow > 0), "The section heads the block"

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Success")
    Assert.IsTrue (entryRow > titleRow), "The outcome line sits under the section"
    'The platform tag follows the timestamp in the same cell, so the date is
    'read off the front of it.
    Assert.IsTrue IsDate(Left$(CStr(sh.Cells(entryRow, DATE_COLUMN).Value), 19)), _
                  "The entry label opens with the date"
    Assert.AreEqual "export-setup: measles.xlsb", _
                    CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
                    "The action and the detail sit in the last written column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLogSuccessWritesActionDateAndDetail", Err.Number, Err.Description
End Sub

'@sub-title Every entry names the platform it was written on.
'@TestMethod("MasterSetupLog")
Public Sub TestEveryEntryNamesThePlatform()
    CustomTestSetTitles Assert, TESTMODULE, "TestEveryEntryNamesThePlatform"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim entryRow As Long
    Dim dateCell As String

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "export-setup", "measles.xlsb"
    Set sh = sut.Wksh()

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Success")
    dateCell = CStr(sh.Cells(entryRow, DATE_COLUMN).Value)

    Assert.IsTrue (InStr(1, dateCell, "mac-", vbTextCompare) > 0 Or _
                   InStr(1, dateCell, "win-", vbTextCompare) > 0), _
                  "The entry names its platform, read: " & dateCell
    Assert.IsTrue (InStr(1, dateCell, "excel-", vbTextCompare) > 0), _
                  "The entry names its Excel version, read: " & dateCell

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEveryEntryNamesThePlatform", Err.Number, Err.Description
End Sub

'@sub-title A section title is written once, however many actions follow it.
'@TestMethod("MasterSetupLog")
Public Sub TestTheSectionTitleIsWrittenOnce()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSectionTitleIsWrittenOnce"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim firstRow As Long
    Dim thirdRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "add-disease", "Measles"
    sut.LogWarning "sort-tables", "no header"
    sut.LogSuccess "add-rows", "Measles"
    Set sh = sut.Wksh()

    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE), _
                    "Three actions of one section stand under one title"
    Assert.AreEqual 0&, RowOfText(sh, OUTPUT_COLUMN, "add-disease"), _
                    "No title row carries the action code"

    firstRow = RowOfText(sh, DETAIL_COLUMN, "add-disease")
    thirdRow = RowOfText(sh, DETAIL_COLUMN, "add-rows")
    Assert.AreEqual 2&, thirdRow - firstRow, _
                    "The three entries follow one another with no blank row"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSectionTitleIsWrittenOnce", Err.Number, Err.Description
End Sub

'@sub-title A second section brings its own title, and an action rejoins its block.
'@details
'A title heads a block of its own kind. The close comes last here and lands
'at the foot of the open/close block, above the blocks opened after it,
'with a blank row between two blocks.
'@TestMethod("MasterSetupLog")
Public Sub TestASecondSectionBringsItsOwnTitle()
    CustomTestSetTitles Assert, TESTMODULE, "TestASecondSectionBringsItsOwnTitle"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim openRow As Long
    Dim closeRow As Long
    Dim firstTitleRow As Long
    Dim secondTitleRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogInfo "workbook-open", "msetup.xlsb"
    sut.LogSuccess "add-disease", "Measles"
    sut.LogSuccess "export-setup", "measles.xlsb"
    sut.LogInfo "workbook-close", "msetup.xlsb"
    Set sh = sut.Wksh()

    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_OPENCLOSE), _
                    "The session section is written once"
    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE), _
                    "The lifecycle section is written once"
    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_DATAIO), _
                    "The data section is written once"

    openRow = RowOfText(sh, DETAIL_COLUMN, "workbook-open")
    closeRow = RowOfText(sh, DETAIL_COLUMN, "workbook-close")
    Assert.AreEqual 1&, closeRow - openRow, _
                    "The close sits at the foot of the block the open opened"
    Assert.AreEqual SECTION_OPENCLOSE, _
                    Trim$(CStr(sh.Cells(closeRow, TITLE_COLUMN).Value)), _
                    "The close line carries its section in the hidden column"

    firstTitleRow = RowOfText(sh, OUTPUT_COLUMN, SECTION_OPENCLOSE)
    secondTitleRow = RowOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE)
    Assert.AreEqual 0&, CLng(LenB(CStr(sh.Cells(secondTitleRow - 1, OUTPUT_COLUMN).Value))), _
                    "A blank row stands between two blocks"
    Assert.IsTrue (secondTitleRow > firstTitleRow), _
                  "The lifecycle block stands below the session block"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASecondSectionBringsItsOwnTitle", Err.Number, Err.Description
End Sub

'@sub-title Every action falls in one of the three sections.
'@details
'The boundary cases are the ones worth stating: the two workbook codes and
'the restart line are session boundaries; every export, every import and
'the comparison move data; a code named by no case falls in the lifecycle
'section; and a code opening with export- or import- moves data on its own.
'@TestMethod("MasterSetupLog")
Public Sub TestActionsFallInTheThreeSections()
    CustomTestSetTitles Assert, TESTMODULE, "TestActionsFallInTheThreeSections"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog

    Set sut = MasterSetupLog.Create(FixtureWkb)

    Assert.AreEqual SECTION_OPENCLOSE, sut.SectionOf("workbook-open"), _
                    "The workbook open is a session boundary"
    Assert.AreEqual SECTION_OPENCLOSE, sut.SectionOf("workbook-close"), _
                    "The workbook close is a session boundary"
    Assert.AreEqual SECTION_OPENCLOSE, sut.SectionOf("log"), _
                    "The restart line is a session boundary"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("export-setup"), _
                    "A setup export moves data out"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("import-migration"), _
                    "A migration import moves data in"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("import-passwords"), _
                    "A password import moves data in"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("compare-diseases"), _
                    "The disease comparison belongs with the data"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("export-log"), _
                    "Writing the log out moves data out"
    Assert.AreEqual SECTION_DATAIO, sut.SectionOf("export-choices"), _
                    "A new export code reaches the data block on its prefix"
    Assert.AreEqual SECTION_LIFECYCLE, sut.SectionOf("add-disease"), _
                    "Adding a disease is work inside the file"
    Assert.AreEqual SECTION_LIFECYCLE, sut.SectionOf("sort-tables"), _
                    "A sort is work inside the file"
    Assert.AreEqual SECTION_LIFECYCLE, sut.SectionOf("MSG_ErrUpdate"), _
                    "A code named by no case falls in the lifecycle section"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestActionsFallInTheThreeSections", Err.Number, Err.Description
End Sub

'@sub-title A second flush appends under the first one.
'@TestMethod("MasterSetupLog")
Public Sub TestTwoFlushesAppend()
    CustomTestSetTitles Assert, TESTMODULE, "TestTwoFlushesAppend"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim firstRow As Long
    Dim secondRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows", "10 rows"
    sut.LogFailure "resize-tables", "err 1004"
    Set sh = sut.Wksh()

    firstRow = RowOfText(sh, DETAIL_COLUMN, "add-rows")
    secondRow = RowOfText(sh, DETAIL_COLUMN, "resize-tables")

    Assert.IsTrue (firstRow > 0), "The first event survives the second flush"
    Assert.IsTrue (secondRow > firstRow), "The second event lands below the first"
    Assert.IsTrue (RowOfText(sh, OUTPUT_COLUMN, "Error") > 0), _
                  "The failure line carries the error outcome"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTwoFlushesAppend", Err.Number, Err.Description
End Sub

'@sub-title Each outcome paints its entry with its own colour.
'@details
'The colours come from CheckingOutput.ResolveFormatting: green for a
'success, red for an error, orange for a warning, grey for an info line.
'The action codes of this test stay clear of the outcome words, so each
'search lands on an entry line.
'@TestMethod("MasterSetupLog")
Public Sub TestColoursFollowTheOutcome()
    CustomTestSetTitles Assert, TESTMODULE, "TestColoursFollowTheOutcome"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows"
    sut.LogFailure "resize-tables"
    sut.LogWarning "sort-tables"
    sut.LogInfo "workbook-open"
    Set sh = sut.Wksh()

    Assert.AreEqual RGB(0, 120, 50), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Success"), OUTPUT_COLUMN).Font.Color, _
                    "A success line is green"
    Assert.AreEqual RGB(192, 0, 0), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Error"), OUTPUT_COLUMN).Font.Color, _
                    "A failure line is red"
    Assert.AreEqual RGB(167, 106, 0), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Warning"), OUTPUT_COLUMN).Font.Color, _
                    "A warning line is orange"
    Assert.AreEqual RGB(96, 97, 100), _
                    sh.Cells(RowOfText(sh, OUTPUT_COLUMN, "Info"), OUTPUT_COLUMN).Font.Color, _
                    "An info line is grey"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestColoursFollowTheOutcome", Err.Number, Err.Description
End Sub

'@sub-title A separator inside the detail or the source stays in the last column.
'@TestMethod("MasterSetupLog")
Public Sub TestSeparatorStaysInTheLastColumn()
    CustomTestSetTitles Assert, TESTMODULE, "TestSeparatorStaysInTheLastColumn"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim entryRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogFailure "resize-tables", "range -- missing", "click--Resize"
    Set sh = sut.Wksh()

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Error")
    Assert.AreEqual "click-Resize > resize-tables: range - missing", _
                    CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
                    "The softened source and detail sit whole in the last column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestSeparatorStaysInTheLastColumn", Err.Number, Err.Description
End Sub

'@sub-title The procedure that raised an event opens the entry column.
'@details
'The section is still read off the action: a data move raised from a
'procedure whose name says nothing lands under data input/output.
'@TestMethod("MasterSetupLog")
Public Sub TestTheSourceOpensTheEntryColumn()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSourceOpensTheEntryColumn"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim entryRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "export-setup", "Measles to measles.xlsb", "clickExpSheet"
    Set sh = sut.Wksh()

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Success")
    Assert.AreEqual "clickExpSheet > export-setup: Measles to measles.xlsb", _
                    CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
                    "The procedure, the action and the detail read in that order"
    Assert.IsTrue (RowOfText(sh, OUTPUT_COLUMN, SECTION_DATAIO) > 0), _
                  "The export still opens the data input/output block"
    Assert.AreEqual 0&, RowOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE), _
                    "The procedure name does not move it to lifecycle"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSourceOpensTheEntryColumn", Err.Number, Err.Description
End Sub

'@sub-title A caller that names no procedure writes a line opening at the action.
'@TestMethod("MasterSetupLog")
Public Sub TestAnUnnamedSourceLeavesTheLineAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnUnnamedSourceLeavesTheLineAlone"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim entryRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows", "5 rows on Measles"
    Set sh = sut.Wksh()

    entryRow = RowOfText(sh, OUTPUT_COLUMN, "Success")
    Assert.AreEqual "add-rows: 5 rows on Measles", _
                    CStr(sh.Cells(entryRow, DETAIL_COLUMN).Value), _
                    "With no procedure named the line opens at the action"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnUnnamedSourceLeavesTheLineAlone", Err.Number, Err.Description
End Sub

'@sub-title An empty action code is refused, and the number says why.
'@TestMethod("MasterSetupLog")
Public Sub TestEmptyActionIsRefused()
    CustomTestSetTitles Assert, TESTMODULE, "TestEmptyActionIsRefused"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim errNumber As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)

    On Error Resume Next
        sut.LogInfo vbNullString
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "An empty action code is refused and the number names the reason"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEmptyActionIsRefused", Err.Number, Err.Description
End Sub

'@section Rotation Tests
'===============================================================================

'@sub-title A flush past the cap clears the sheet and restarts the log.
'@details
'The overflow is seeded as one cell past the cap in the output column,
'so the rotation runs without the suite writing ten thousand lines.
'@TestMethod("MasterSetupLog")
Public Sub TestRotationClearsThePastLog()
    CustomTestSetTitles Assert, TESTMODULE, "TestRotationClearsThePastLog"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim sh As Worksheet
    Dim restartRow As Long
    Dim eventRow As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogSuccess "add-rows", "before the cap"
    Set sh = sut.Wksh()
    sh.Cells(sut.MaxEntries + 1, OUTPUT_COLUMN).Value = "overflow marker"

    sut.LogInfo "workbook-open", "after the cap"

    Assert.AreEqual 0&, CLng(LenB(CStr(sh.Cells(sut.MaxEntries + 1, OUTPUT_COLUMN).Value))), _
                    "The overflow cell was cleared"
    Assert.AreEqual 0&, RowOfText(sh, DETAIL_COLUMN, "add-rows"), _
                    "The old log lines are gone"

    restartRow = RowOfText(sh, DETAIL_COLUMN, "restarted")
    eventRow = RowOfText(sh, DETAIL_COLUMN, "workbook-open: after the cap")
    Assert.IsTrue (restartRow > 0), "The fresh log opens with the restart line"
    Assert.IsTrue (eventRow > restartRow), _
                  "The event that met the cap lands after the restart line"
    Assert.IsTrue (LastOutputRow(sh) < sut.MaxEntries), _
                  "The fresh log sits far below the cap"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestRotationClearsThePastLog", Err.Number, Err.Description
End Sub

'@section Text Export Tests
'===============================================================================
'The report is built in memory and read back in memory. Writing a file is
'left to ExportText, which is one Open, one loop and one Close over the same
'lines, and a file-writing test would meet the macOS file-access panel.

'@sub-title The report opens with the workbook, the platform and the disease count.
'@details
'A master setup has no Metadata sheet, so the header carries the file name,
'the date, the platform and the count of disease worksheets. Two tagged
'sheets and one plain sheet are built, and the count reads two.
'@TestMethod("MasterSetupLog")
Public Sub TestTheReportOpensWithTheWorkbookAndTheDiseaseCount()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheReportOpensWithTheWorkbookAndTheDiseaseCount"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim lines As BetterArray
    Dim writtenLine As String

    AddDiseaseSheet "Measles"
    AddDiseaseSheet "Cholera"
    FixtureWkb.Worksheets.Add(After:=FixtureWkb.Worksheets(FixtureWkb.Worksheets.Count)).Name = "Choices"

    Set sut = MasterSetupLog.Create(FixtureWkb)
    Assert.AreEqual 2&, sut.DiseaseSheetCount(), _
                    "Two tagged sheets are counted and the plain one is left out"

    Set lines = sut.ReportLines()

    Assert.IsTrue (lines.Length > 0), "The report carries lines"
    Assert.IsTrue InStr(1, CStr(lines.Item(lines.LowerBound)), _
                        "master setup log", vbTextCompare) > 0, _
                  "The first line names what the file is"
    Assert.IsTrue (LenB(LineOfText(lines, FixtureWkb.Name)) > 0), _
                  "The report names the workbook it came out of"
    Assert.AreEqual "disease worksheets 2", LineOfText(lines, "disease worksheets"), _
                    "The report counts the disease worksheets"

    writtenLine = LineOfText(lines, "written")
    Assert.IsTrue (InStr(1, writtenLine, "mac-", vbTextCompare) > 0 Or _
                   InStr(1, writtenLine, "win-", vbTextCompare) > 0), _
                  "The written line names the platform, read: " & writtenLine

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheReportOpensWithTheWorkbookAndTheDiseaseCount", _
                         Err.Number, Err.Description
End Sub

'@sub-title The logged entries reach the report with their section and outcome.
'@details
'The outcome reads as the bare word: the sheet paints a symbol before it,
'and Print # would write that symbol as a stray mark.
'@TestMethod("MasterSetupLog")
Public Sub TestTheReportCarriesTheLoggedEntries()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheReportCarriesTheLoggedEntries"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim lines As BetterArray
    Dim entryLine As String

    Set sut = MasterSetupLog.Create(FixtureWkb)
    sut.LogFailure "import-migration", "old.xlsb would not open", "clickImpMig"
    Set lines = sut.ReportLines()

    Assert.IsTrue (LenB(LineOfText(lines, "-- log --")) > 0), _
                  "The report carries a log block"
    Assert.IsTrue (LenB(LineOfText(lines, "[" & SECTION_DATAIO & "]")) > 0), _
                  "The section opens its block in the report with its name whole"

    entryLine = LineOfText(lines, "old.xlsb would not open")
    Assert.IsTrue (LenB(entryLine) > 0), "The entry reaches the report"
    Assert.IsTrue InStr(1, entryLine, "[Error]", vbBinaryCompare) > 0, _
                  "The entry carries its outcome as the bare word, read: " & entryLine
    Assert.AreEqual 0&, CLng(InStr(1, entryLine, "[_ ", vbBinaryCompare)), _
                    "No stray mark stands in front of it, read: " & entryLine
    Assert.IsTrue InStr(1, entryLine, "clickImpMig", vbTextCompare) > 0, _
                  "The procedure reaches the report, read: " & entryLine

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheReportCarriesTheLoggedEntries", _
                         Err.Number, Err.Description
End Sub

'@sub-title Exporting to no folder is refused before any file is opened.
'@TestMethod("MasterSetupLog")
Public Sub TestExportTextRefusesAnEmptyFolder()
    CustomTestSetTitles Assert, TESTMODULE, "TestExportTextRefusesAnEmptyFolder"
    On Error GoTo TestFail

    Dim sut As MasterSetupLog
    Dim writtenPath As String
    Dim errNumber As Long

    Set sut = MasterSetupLog.Create(FixtureWkb)

    On Error Resume Next
        writtenPath = sut.ExportText(vbNullString)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "An empty folder is refused and the number names the reason"
    Assert.AreEqual 0&, CLng(LenB(writtenPath)), _
                    "Nothing is written when the folder is empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExportTextRefusesAnEmptyFolder", _
                         Err.Number, Err.Description
End Sub

'@section Event Service Tests
'===============================================================================

'@sub-title The event service builds the log once and answers the same one.
'@details
'Every module reaches the log through EventMasterSetup.UserLog, so one
'workbook has one log and one __log sheet whatever the number of callers.
'@TestMethod("MasterSetupLog")
Public Sub TestTheEventServiceHoldsOneLog()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheEventServiceHoldsOneLog"
    On Error GoTo TestFail

    Dim service As EventMasterSetup
    Dim firstLog As MasterSetupLog
    Dim secondLog As MasterSetupLog
    Dim countBefore As Long

    countBefore = FixtureWkb.Worksheets.Count
    Set service = EventMasterSetup.Create(FixtureWkb)

    Set firstLog = service.UserLog()
    Set secondLog = service.UserLog()

    Assert.IsTrue (Not firstLog Is Nothing), "The service answers a log"
    Assert.IsTrue (firstLog Is secondLog), "The second call answers the held log"
    Assert.AreEqual countBefore + 1, FixtureWkb.Worksheets.Count, _
                    "The workbook gained one log sheet"
    Assert.AreEqual LOG_SHEET, firstLog.Wksh().Name, _
                    "The held log is bound to the log sheet"
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(FixtureWkb.Worksheets(LOG_SHEET).Visible), _
                    "The sheet the service built is very hidden"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheEventServiceHoldsOneLog", Err.Number, Err.Description
End Sub

'@sub-title An open, a disease added and an export leave three lines in three sections.
'@details
'The acceptance line of the wiring: the three actions a session takes
'first, written through the held log the way MasterSetupEventsManager,
'EventsMasterSetupRibbon and MasterSetupExports write them.
'@TestMethod("MasterSetupLog")
Public Sub TestAnOpenAnAddAndAnExportLeaveThreeLines()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnOpenAnAddAndAnExportLeaveThreeLines"
    On Error GoTo TestFail

    Dim service As EventMasterSetup
    Dim heldLog As MasterSetupLog
    Dim sh As Worksheet

    Set service = EventMasterSetup.Create(FixtureWkb)
    Set heldLog = service.UserLog()

    heldLog.LogInfo "workbook-open", FixtureWkb.Name, "MsWorkbookOpened"
    heldLog.LogSuccess "add-disease", "Measles", "clickAddSheet"
    heldLog.LogSuccess "export-setup", "Measles to C:\exports\measles.xlsx", "ExportToSetup"
    Set sh = heldLog.Wksh()

    Assert.IsTrue (RowOfText(sh, DETAIL_COLUMN, "workbook-open") > 0), "The open line is on the sheet"
    Assert.IsTrue (RowOfText(sh, DETAIL_COLUMN, "add-disease: Measles") > 0), "The add line is on the sheet"
    Assert.IsTrue (RowOfText(sh, DETAIL_COLUMN, "export-setup: Measles") > 0), "The export line is on the sheet"
    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_OPENCLOSE), _
                    "The open line sits under open/close"
    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_LIFECYCLE), _
                    "The add line sits under master setup lifecycle"
    Assert.AreEqual 1&, CountOfText(sh, OUTPUT_COLUMN, SECTION_DATAIO), _
                    "The export line sits under data input/output"
    Assert.IsTrue (RowOfText(sh, DETAIL_COLUMN, "clickAddSheet") > 0), _
                  "The add line names the callback that wrote it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnOpenAnAddAndAnExportLeaveThreeLines", _
                         Err.Number, Err.Description
End Sub
