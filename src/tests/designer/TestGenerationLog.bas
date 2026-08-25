Attribute VB_Name = "TestGenerationLog"
Attribute VB_Description = "Unit tests for the GenerationLog class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates GenerationLog: the run header, the stamped in-memory record, the append across two flushes, the run window, the report sections and their subsections, the closing bundle, the text export and the re-run reset.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private FixtureWorkbook As Workbook

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'The worksheet the log writes on
Private Const SHEET_CHECKING As String = "__check"

'Base name of the text export written into this workbook's folder
Private Const EXPORT_BASE As String = "obt_generationlog_test"

'The first visible column of the report. The writer puts the first
'fragment of every row it writes there, so a heading is one cell of it.
Private Const HEADING_COLUMN As Long = 3


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestGenerationLog"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    BusyApp

    Set FixtureWorkbook = NewWorkbook
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        Kill ExportFilePath()
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set FixtureWorkbook = Nothing

    'A mid-test workbook close can hand the screen to another workbook
    ThisWorkbook.Activate

    RestoreApp
End Sub


'@section Start Tests
'===============================================================================
'@TestMethod("GenerationLog.Start")
Public Sub TestStartCreatesCheckSheetAndHeader()
    CustomTestSetTitles Assert, "GenerationLog", "TestStartCreatesCheckSheetAndHeader"
    On Error GoTo Fail

    'Act: a run opens with the setup path and the linelist name known
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start "setup.xlsb", "linelist_v1"

    'Assert: the sheet exists and the header carries its three entries
    Assert.IsTrue WorksheetExists(SHEET_CHECKING, FixtureWorkbook), _
                  "Start should create the " & SHEET_CHECKING & " worksheet."
    Assert.AreEqual CLng(3), runLog.RecordLength, _
                    "The run header should file the start time, the setup path and the linelist name."
    Assert.IsTrue InStr(1, runLog.RecordLine(1), "Generation run") > 0, _
                  "The first record line should carry the run header title."
    Assert.IsTrue InStr(1, runLog.RecordLine(2), "setup.xlsb") > 0, _
                  "The second record line should carry the setup path."
    'The platform rides on the start line, so the record keeps its length. The
    'check reads the name and the Excel version separately, since the version
    'moves with the host and cannot be written down here.
    Assert.IsTrue (InStr(1, runLog.RecordLine(1), "mac-", vbTextCompare) > 0 Or _
                   InStr(1, runLog.RecordLine(1), "win-", vbTextCompare) > 0), _
                  "The start line should name the platform of the run."
    Assert.IsTrue InStr(1, runLog.RecordLine(1), "excel-", vbTextCompare) > 0, _
                  "The start line should name the Excel version."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestStartCreatesCheckSheetAndHeader", Err.Number, Err.Description
End Sub


'@section Record Tests
'===============================================================================
'@TestMethod("GenerationLog.Record")
Public Sub TestCollectAppendsAcrossTwoFlushes()
    CustomTestSetTitles Assert, "GenerationLog", "TestCollectAppendsAcrossTwoFlushes"
    On Error GoTo Fail

    'Arrange: a bare run, whose header holds one entry
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start

    'Act: two flushes
    runLog.Collect MakeBundle("first bundle", 2)
    runLog.Collect MakeBundle("second bundle", 1)

    'Assert: the record holds the header plus both bundles, in order
    Assert.AreEqual CLng(4), runLog.RecordLength, _
                    "The record should hold the header entry plus the three collected entries."
    Assert.IsTrue InStr(1, runLog.RecordLine(2), "first bundle") > 0, _
                  "The first collected entry should follow the header."
    Assert.IsTrue InStr(1, runLog.RecordLine(4), "second bundle") > 0, _
                  "The second flush should append after the first."

    'Assert: both bundles reached the worksheet
    Dim sh As Worksheet
    Set sh = FixtureWorkbook.Worksheets(SHEET_CHECKING)
    Assert.IsNotNothing sh.Cells.Find(What:="first bundle", LookIn:=xlValues, LookAt:=xlPart), _
                        "The first bundle should land on the worksheet."
    Assert.IsNotNothing sh.Cells.Find(What:="second bundle", LookIn:=xlValues, LookAt:=xlPart), _
                        "The second flush should append on the worksheet."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCollectAppendsAcrossTwoFlushes", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.Record")
Public Sub TestRecordLineCarriesStampScopeTitleAndLabel()
    CustomTestSetTitles Assert, "GenerationLog", "TestRecordLineCarriesStampScopeTitleAndLabel"
    On Error GoTo Fail

    'Arrange
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start

    Dim faultBundle As Checking
    Set faultBundle = Checking.Create("geo checks")
    faultBundle.Add "geo file", "The geobase would not open", checkingError

    'Act
    runLog.Collect faultBundle

    'Assert: the line reads hh:mm:ss  [scope]  title  label
    Dim line As String
    line = runLog.RecordLine(2)
    Assert.AreEqual ":", Mid$(line, 3, 1), _
                    "The line should open with the hh:mm:ss stamp."
    Assert.AreEqual ":", Mid$(line, 6, 1), _
                    "The stamp should carry its second colon."
    Assert.IsTrue InStr(1, line, "[Error]") > 0, _
                  "The scope should read as the plain word in brackets."
    Assert.IsTrue InStr(1, line, "geo checks") > 0, _
                  "The line should carry the bundle title."
    Assert.IsTrue InStr(1, line, "The geobase would not open") > 0, _
                  "The line should carry the entry label."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestRecordLineCarriesStampScopeTitleAndLabel", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.Record")
Public Sub TestCollectOutsideRunWindowIsIgnored()
    CustomTestSetTitles Assert, "GenerationLog", "TestCollectOutsideRunWindowIsIgnored"
    On Error GoTo Fail

    'Arrange: a log that never started
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)

    'Act and assert: a bundle before Start is dropped with no raise
    runLog.Collect MakeBundle("early bundle", 1)
    Assert.AreEqual CLng(0), runLog.RecordLength, _
                    "A bundle before Start should leave the record empty."

    'Act and assert: a bundle after Finish is dropped too
    runLog.Start
    runLog.Finish "done"
    Dim closedLength As Long
    closedLength = runLog.RecordLength
    runLog.Collect MakeBundle("late bundle", 1)
    Assert.AreEqual closedLength, runLog.RecordLength, _
                    "A bundle after Finish should leave the record as Finish closed it."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCollectOutsideRunWindowIsIgnored", Err.Number, Err.Description
End Sub


'@TestMethod("GenerationLog.Record")
Public Sub TestARecordOnlyBundleStaysOffTheWorksheet()
    CustomTestSetTitles Assert, "GenerationLog", "TestARecordOnlyBundleStaysOffTheWorksheet"
    On Error GoTo Fail

    'Arrange: a bare run, whose header holds one entry
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start

    'Act: one bundle each way
    runLog.Collect MakeBundle("sheet bundle", 1)
    runLog.Collect MakeBundle("detail bundle", 3), recordOnly:=True

    'Assert: both bundles are in the record, which is what the text file reads
    Assert.AreEqual CLng(5), runLog.RecordLength, _
                    "A record-only bundle should still reach the in-memory record."
    Assert.IsTrue InStr(1, runLog.RecordLine(3), "detail bundle") > 0, _
                  "The record-only entries should follow the flushed bundle in order."

    'Assert: only the flushed bundle reached the worksheet. This is what keeps
    'an EntireColumn.AutoFit from running behind every variable of a build.
    Dim sh As Worksheet
    Set sh = FixtureWorkbook.Worksheets(SHEET_CHECKING)
    Assert.IsNotNothing sh.Cells.Find(What:="sheet bundle", LookIn:=xlValues, LookAt:=xlPart), _
                        "A flushed bundle should land on the worksheet."
    Assert.IsNothing sh.Cells.Find(What:="detail bundle", LookIn:=xlValues, LookAt:=xlPart), _
                     "A record-only bundle should stay off the worksheet."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestARecordOnlyBundleStaysOffTheWorksheet", Err.Number, Err.Description
End Sub


'@section Section Tests
'===============================================================================
'@TestMethod("GenerationLog.Section")
Public Sub TestASectionHeadsItsBundlesOnce()
    CustomTestSetTitles Assert, "GenerationLog", "TestASectionHeadsItsBundlesOnce"
    On Error GoTo Fail

    'Arrange
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start

    'Act: two parts of one build, under one section
    runLog.OpenSection "linelist one"
    runLog.Collect MakeBundle("dictionary", 1)
    runLog.Collect MakeBundle("choices", 1)
    runLog.CloseSection

    'Assert: the section title stands once, with both parts under it
    Dim sh As Worksheet
    Set sh = FixtureWorkbook.Worksheets(SHEET_CHECKING)
    Assert.AreEqual CLng(1), TitleCount(sh, "linelist one"), _
                    "The section title should reach the worksheet once."
    Assert.IsNotNothing sh.Cells.Find(What:="dictionary", LookIn:=xlValues, LookAt:=xlPart), _
                        "The first part should land as a subsection of the section."
    Assert.IsNotNothing sh.Cells.Find(What:="choices", LookIn:=xlValues, LookAt:=xlPart), _
                        "The second part should land as a subsection of the section."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestASectionHeadsItsBundlesOnce", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.Section")
Public Sub TestAnOpenSectionHoldsItsBundlesBack()
    CustomTestSetTitles Assert, "GenerationLog", "TestAnOpenSectionHoldsItsBundlesBack"
    On Error GoTo Fail

    'Arrange
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start
    Dim sh As Worksheet
    Set sh = FixtureWorkbook.Worksheets(SHEET_CHECKING)

    'Act: one bundle inside a section that stays open
    runLog.OpenSection "linelist two"
    runLog.Collect MakeBundle("held bundle", 1)

    'Assert: the record has it and the worksheet stays clear
    Assert.AreEqual "linelist two", runLog.SectionTitle, _
                    "The open section should answer its own title."
    Assert.AreEqual CLng(2), runLog.RecordLength, _
                    "A held bundle should reach the in-memory record straight away."
    Assert.IsNothing sh.Cells.Find(What:="held bundle", LookIn:=xlValues, LookAt:=xlPart), _
                     "A bundle of an open section should stay off the worksheet."

    'Act: closing writes the whole section
    runLog.CloseSection

    'Assert
    Assert.AreEqual vbNullString, runLog.SectionTitle, _
                    "A closed section should leave no title open."
    Assert.IsNotNothing sh.Cells.Find(What:="held bundle", LookIn:=xlValues, LookAt:=xlPart), _
                        "Closing the section should write the bundles it held."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestAnOpenSectionHoldsItsBundlesBack", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.Section")
Public Sub TestOpeningASectionClosesTheOneBeforeIt()
    CustomTestSetTitles Assert, "GenerationLog", "TestOpeningASectionClosesTheOneBeforeIt"
    On Error GoTo Fail

    'Arrange
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start

    'Act: a driver walking two rows and closing neither by hand
    runLog.OpenSection "row one"
    runLog.Collect MakeBundle("first part", 1)
    runLog.OpenSection "row two"
    runLog.Collect MakeBundle("second part", 1)
    runLog.Finish "done"

    'Assert: both sections stand, each once, and Finish wrote the last one
    Dim sh As Worksheet
    Set sh = FixtureWorkbook.Worksheets(SHEET_CHECKING)
    Assert.AreEqual CLng(1), TitleCount(sh, "row one"), _
                    "The first section should stand on the worksheet once."
    Assert.AreEqual CLng(1), TitleCount(sh, "row two"), _
                    "Finish should write the section that was still open."
    Assert.AreEqual vbNullString, runLog.SectionTitle, _
                    "Finish should leave no section open."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestOpeningASectionClosesTheOneBeforeIt", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.Section")
Public Sub TestARecordLineOfASectionNamesBothTitles()
    CustomTestSetTitles Assert, "GenerationLog", "TestARecordLineOfASectionNamesBothTitles"
    On Error GoTo Fail

    'Arrange
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start

    'Act
    runLog.OpenSection "linelist three"
    runLog.Collect MakeBundle("analyses", 1)
    runLog.CloseSection

    'Assert: the text file reads the way the worksheet does
    Assert.IsTrue InStr(1, runLog.RecordLine(2), "linelist three - analyses") > 0, _
                  "A record line of a section should name the section and the part."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestARecordLineOfASectionNamesBothTitles", Err.Number, Err.Description
End Sub


'@section Finish Tests
'===============================================================================
'@TestMethod("GenerationLog.Finish")
Public Sub TestFinishWritesOutcomeAndShowsSheet()
    CustomTestSetTitles Assert, "GenerationLog", "TestFinishWritesOutcomeAndShowsSheet"
    On Error GoTo Fail

    'Arrange
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start

    'Act
    runLog.Finish "All sheets built", sheetCount:=3, variableCount:=42

    'Assert: the closing bundle carries the outcome, the counts and the clock
    Dim lastLine As String
    lastLine = runLog.RecordLine(runLog.RecordLength)
    Assert.IsTrue InStr(1, lastLine, "seconds") > 0, _
                  "The closing bundle should end on the elapsed seconds."
    Assert.IsTrue InStr(1, runLog.RecordLine(2), "All sheets built") > 0, _
                  "The outcome should be the first closing entry."
    Assert.IsTrue InStr(1, runLog.RecordLine(3), "3 data entry sheet(s) built") > 0, _
                  "The sheet count should be written when the driver knows it."
    Assert.IsTrue InStr(1, runLog.RecordLine(4), "42 variable(s) written") > 0, _
                  "The variable count should be written when the driver knows it."

    'Assert: the sheet is shown, and a second Finish changes nothing
    Assert.AreEqual CLng(xlSheetVisible), _
                    CLng(FixtureWorkbook.Worksheets(SHEET_CHECKING).Visible), _
                    "Finish should make the report sheet visible."
    Dim closedLength As Long
    closedLength = runLog.RecordLength
    runLog.Finish "again"
    Assert.AreEqual closedLength, runLog.RecordLength, _
                    "A second Finish should leave the record as the first one closed it."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestFinishWritesOutcomeAndShowsSheet", Err.Number, Err.Description
End Sub


'@section ShowReport Tests
'===============================================================================
'Finish used to activate the report sheet itself, under On Error Resume Next.
'Worksheet.Activate raises 1004 when the sheet belongs to a workbook that is not
'the active one, and a generation ends with the built linelist open and in
'front, so that call raised on every real run and the raise was swallowed. The
'report never appeared at the end of a generation.
'
'The first test below is that run, in miniature: another workbook holds the
'screen when the report is asked for.

'@TestMethod("GenerationLog.ShowReport")
Public Sub TestShowReportComesUpFromUnderAnotherWorkbook()
    CustomTestSetTitles Assert, "GenerationLog", "TestShowReportComesUpFromUnderAnotherWorkbook"
    On Error GoTo Fail

    'Arrange: a finished run on the fixture workbook
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start
    runLog.Finish "built"

    'Act: another workbook takes the screen, then the report is asked for
    ThisWorkbook.Activate
    Dim shown As Boolean
    shown = runLog.ShowReport()

    'Assert
    Assert.IsTrue shown, _
                  "ShowReport should answer True when it has a report to show."
    Assert.AreEqual SHEET_CHECKING, ActiveSheet.Name, _
                    "The report sheet should be the active sheet, whatever held the screen."

    ThisWorkbook.Activate

    Exit Sub
Fail:
    ThisWorkbook.Activate
    CustomTestLogFailure Assert, "TestShowReportComesUpFromUnderAnotherWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.ShowReport")
Public Sub TestShowReportWorksAfterFinishReleasedTheWriter()
    CustomTestSetTitles Assert, "GenerationLog", "TestShowReportWorksAfterFinishReleasedTheWriter"
    On Error GoTo Fail

    'Arrange: a run that is over, and its report put away again
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start
    runLog.Finish "built"
    FixtureWorkbook.Worksheets(SHEET_CHECKING).Visible = xlSheetVeryHidden

    'Act: the ribbon button, pressed some time after the run
    Dim shown As Boolean
    shown = runLog.ShowReport()

    'Assert: the sheet outlives the writer Finish released
    Assert.IsTrue shown, _
                  "A report should still open after the run that wrote it has closed."
    Assert.AreEqual CLng(xlSheetVisible), _
                    CLng(FixtureWorkbook.Worksheets(SHEET_CHECKING).Visible), _
                    "ShowReport should bring the sheet back out of hiding."

    ThisWorkbook.Activate

    Exit Sub
Fail:
    ThisWorkbook.Activate
    CustomTestLogFailure Assert, "TestShowReportWorksAfterFinishReleasedTheWriter", Err.Number, Err.Description
End Sub

'@sub-title A workbook that has never generated anything has no report.
'@details
'The ribbon button builds a log over ThisWorkbook when no run has been opened
'this session, so it asks this question on a designer that was just opened. An
'answer of True there would show the user an empty grid.
'@TestMethod("GenerationLog.ShowReport")
Public Sub TestShowReportAnswersFalseWithNoReportSheet()
    CustomTestSetTitles Assert, "GenerationLog", "TestShowReportAnswersFalseWithNoReportSheet"
    On Error GoTo Fail

    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)

    Assert.IsFalse runLog.ShowReport(), _
                   "A workbook with no " & SHEET_CHECKING & " sheet has no report to show."
    Assert.IsFalse WorksheetExists(SHEET_CHECKING, FixtureWorkbook), _
                   "And asking should not have created one."

    Exit Sub
Fail:
    ThisWorkbook.Activate
    CustomTestLogFailure Assert, "TestShowReportAnswersFalseWithNoReportSheet", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.ShowReport")
Public Sub TestShowReportAnswersFalseOnAnEmptyReportSheet()
    CustomTestSetTitles Assert, "GenerationLog", "TestShowReportAnswersFalseOnAnEmptyReportSheet"
    On Error GoTo Fail

    'Arrange: the sheet is there but nothing was ever written on it, which is
    'what a run that raised before its first flush leaves behind
    Dim sh As Worksheet
    Set sh = FixtureWorkbook.Worksheets.Add
    sh.Name = SHEET_CHECKING
    sh.Cells.Clear
    sh.Visible = xlSheetVeryHidden

    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)

    Assert.IsFalse runLog.ShowReport(), _
                   "An empty sheet is not a report."
    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(sh.Visible), _
                    "And it should be left where it was."

    Exit Sub
Fail:
    ThisWorkbook.Activate
    CustomTestLogFailure Assert, "TestShowReportAnswersFalseOnAnEmptyReportSheet", Err.Number, Err.Description
End Sub


'@section Export Tests
'===============================================================================
'@TestMethod("GenerationLog.Export")
Public Sub TestExportTextWritesOneLinePerEntry()
    CustomTestSetTitles Assert, "GenerationLog", "TestExportTextWritesOneLinePerEntry"
    On Error GoTo Fail

    'Arrange: a short finished run
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start "setup.xlsb", "linelist_v1"
    runLog.Collect MakeBundle("export bundle", 2)
    runLog.Finish "done"

    'Act
    Dim writtenPath As String
    writtenPath = runLog.ExportText(ThisWorkbook.Path, EXPORT_BASE)

    'Assert: the file sits where the caller asked, one line per entry
    Assert.AreEqual ExportFilePath(), writtenPath, _
                    "The export should land as <baseName>-generation.txt in the given folder."

    Dim fileLines As Collection
    Set fileLines = ReadAllLines(writtenPath)
    Assert.AreEqual runLog.RecordLength, CLng(fileLines.Count), _
                    "The file should hold one line per record entry."
    Assert.AreEqual runLog.RecordLine(1), CStr(fileLines.Item(1)), _
                    "The first file line should be the first record line."
    Assert.IsTrue InStr(1, CStr(fileLines.Item(fileLines.Count)), "seconds") > 0, _
                  "The last file line should be the elapsed entry of the closing bundle."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestExportTextWritesOneLinePerEntry", Err.Number, Err.Description
End Sub

'@TestMethod("GenerationLog.Export")
Public Sub TestExportTextBeforeStartRaises()
    CustomTestSetTitles Assert, "GenerationLog", "TestExportTextBeforeStartRaises"
    On Error GoTo Fail

    'Arrange: a log that never started
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)

    'Act
    Dim raisedNumber As Long
    On Error Resume Next
    runLog.ExportText ThisWorkbook.Path, EXPORT_BASE
    raisedNumber = Err.Number
    On Error GoTo Fail

    'Assert
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), raisedNumber, _
                    "An export with no record should raise ObjectNotInitialized."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestExportTextBeforeStartRaises", Err.Number, Err.Description
End Sub


'@section Re-run Tests
'===============================================================================
'@TestMethod("GenerationLog.ReRun")
Public Sub TestSecondStartResetsRecordMarkerAndSheet()
    CustomTestSetTitles Assert, "GenerationLog", "TestSecondStartResetsRecordMarkerAndSheet"
    On Error GoTo Fail

    'The stage carries into the failure line. Five calls sit between the
    'start of this test and its first assertion, and a bare error number
    'names none of them.
    Dim stage As String

    'Arrange: a finished first run with one bundle on the sheet
    Dim runLog As GenerationLog
    stage = "create"
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    stage = "first start"
    runLog.Start "setup.xlsb", "linelist_v1"
    stage = "collect"
    runLog.Collect MakeBundle("stale bundle", 1)
    stage = "finish"
    runLog.Finish "done"

    'Act: the re-run
    stage = "second start"
    runLog.Start
    stage = "assertions"

    'Assert: the record starts over with the bare header
    Assert.AreEqual CLng(1), runLog.RecordLength, _
                    "A re-run should reset the record to the new header."

    'Assert: the previous run's rows are cleared. Start drops the render
    'marker before the header flush, so the writer formats the cleared
    'sheet as a first render and the new marker belongs to the new run.
    Dim sh As Worksheet
    Set sh = FixtureWorkbook.Worksheets(SHEET_CHECKING)
    Assert.IsNothing sh.Cells.Find(What:="stale bundle", LookIn:=xlValues, LookAt:=xlPart), _
                     "A re-run should clear the previous run from the sheet."

    'Assert: the new run writes on the cleared sheet
    Assert.IsNotNothing sh.Cells.Find(What:="Generation run", LookIn:=xlValues, LookAt:=xlPart), _
                        "The new header should land on the cleared sheet."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, _
        "TestSecondStartResetsRecordMarkerAndSheet at " & stage & _
        " [source " & Err.Source & "]", Err.Number, Err.Description
End Sub


'@section Test helpers
'===============================================================================

'@sub-title Build one info bundle with the asked number of entries
Private Function MakeBundle(ByVal title As String, ByVal entryCount As Long) As Checking
    Dim bundle As Checking
    Dim index As Long

    Set bundle = Checking.Create(title)
    For index = 1 To entryCount
        bundle.Add "entry " & CStr(index), title & " label " & CStr(index), checkingInfo
    Next index

    Set MakeBundle = bundle
End Function

'@sub-title How many rows of the report carry one text as their heading
'@details
'The writer puts the first fragment of every row it writes in the first
'visible column, so a level one heading is one cell of that column. The
'count answers how many times a section title reached the sheet.
Private Function TitleCount(ByVal sh As Worksheet, ByVal titleText As String) As Long
    Dim cell As Range
    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = sh.Cells(sh.Rows.Count, HEADING_COLUMN).End(xlUp).Row

    For rowIndex = 1 To lastRow
        Set cell = sh.Cells(rowIndex, HEADING_COLUMN)
        If StrComp(CStr(cell.Value), titleText, vbTextCompare) = 0 Then
            TitleCount = TitleCount + 1
        End If
    Next rowIndex
End Function

'@sub-title The full path of the test's text export
Private Function ExportFilePath() As String
    ExportFilePath = ThisWorkbook.Path & Application.PathSeparator & _
                     EXPORT_BASE & "-generation.txt"
End Function

'@sub-title Read a text file into a collection of lines
Private Function ReadAllLines(ByVal filePath As String) As Collection
    Dim fileNumber As Long
    Dim lineText As String
    Dim fileLines As Collection

    Set fileLines = New Collection
    fileNumber = FreeFile

    Open filePath For Input As #fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineText
        fileLines.Add lineText
    Loop
    Close #fileNumber

    Set ReadAllLines = fileLines
End Function
