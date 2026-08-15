Attribute VB_Name = "TestGenerationLog"
Attribute VB_Description = "Unit tests for the GenerationLog class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates GenerationLog: the run header, the stamped in-memory record, the append across two flushes, the run window, the closing bundle, the text export and the re-run reset.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private FixtureWorkbook As Workbook

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'The worksheet the log writes on
Private Const SHEET_CHECKING As String = "__check"

'Base name of the text export written into this workbook's folder
Private Const EXPORT_BASE As String = "obt_generationlog_test"


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

    'Arrange: a finished first run with one bundle on the sheet
    Dim runLog As GenerationLog
    Set runLog = GenerationLog.Create(FixtureWorkbook)
    runLog.Start "setup.xlsb", "linelist_v1"
    runLog.Collect MakeBundle("stale bundle", 1)
    runLog.Finish "done"

    'Act: the re-run
    runLog.Start

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
    CustomTestLogFailure Assert, "TestSecondStartResetsRecordMarkerAndSheet", Err.Number, Err.Description
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
