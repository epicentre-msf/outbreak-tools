Attribute VB_Name = "TestLLImportTiming"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Times the import walk against a real linelist")
'
'MEASURED -- 2026-09-01, macOS 27.0, Excel 16.111 headless, run-tests.R
'--build over a registry narrowed to helpers plus this module. Recorded
'here so the next reader has a yardstick without paying for a run. The
'numbers move with whatever else the machine is doing, so read each one
'beside its fixture line, which says what size it was taken at. This box
'is en_FR, so a number printed by Format$ carries a DECIMAL COMMA.
'
'   TIMING linelist shape (table rows per sheet): dropdown_lists__=3
'       ana_tabnames__=144 import_rep__=0 spatial_tables__=1 Translations=110
'       LinelistTranslation=10 __pass=9000 print_Linelist patients=11 Linelist
'       patients=357 fLinelist patients=201 Geo=1 Dictionary=40 | dictionary
'       rows=41| platform macOS
'   TIMING export migration, whole walk, REAL linelist: 1.148 s
'   TIMING steps: the linelist carries no __log sheet
'   TIMING 1 reading what the file says about itself: 0.008 s
'   TIMING same language: False (False means the last three steps did not
'       run)
'
'@description
'   A MEASURING PROBE, not a behaviour suite, and it stays commented out of the
'   registry the way TestCustomTableTiming and TestLLExporterTiming do.
'   Uncomment the row, run, read the TIMING lines out of the run CSV, comment it
'   back out.
'
'   WHAT IT IS FOR
'   ---------------------------------------------------------------------------
'   The export and import speed block owes a baseline for BOTH walks. Session
'   117 landed the stopwatch, TestLLExporterTiming drives the export walk over a
'   synthetic linelist, and this module is the import half. No session in the
'   block had timed the import walk when this was written.
'
'   WHY THIS IS A SEPARATE MODULE FROM TestLLExporterTiming
'   ---------------------------------------------------------------------------
'   It runs against a REAL linelist -- src/tests/.input/old_linelist.xlsb, a
'   measles linelist with data in it, the same file TestExportOtherLinelist
'   uses. That file is gitignored (.gitignore:32), so it does not travel to the
'   Windows peer. Keeping the two probes apart means the synthetic export probe
'   still runs where the .xlsb is absent, and only this one reports MISSING.
'
'   ONE RUN, BOTH WALKS, THE SAME LINELIST
'   ---------------------------------------------------------------------------
'   The import needs a file to read, so the test exports a migration first and
'   imports that back. Both walks therefore run over the same real data on the
'   same machine in the same minute, which is what makes the two numbers
'   comparable -- the one thing the synthetic probe could not give.
'   The export leaves its own step list on the linelist's __log sheet under
'   `export-migration`; the import steps are timed here call by call instead,
'   for the reason in the next block.
'
'   IT TIMES THE CLASS METHODS, NOT THE WRAPPER, AND HERE IS WHY
'   ---------------------------------------------------------------------------
'   LinelistRun.HandleImportData could NOT be driven from inside the headless
'   harness. Five runs: the full registry answered -50 with a zero-byte CSV and
'   no result for any module, and narrowed to this module alone it came back
'   with modulesRun=10 testsFound=1 total=0 -- the test never returned from the
'   call, so TestCleanup never flushed what had already been logged. That is the
'   same shape as the GenerationHost.Run guard this suite already carries.
'
'   So the probe calls the LLImporter methods in the order HandleImportData
'   calls them, timing one at a time.
'     LOST:  the wrapper's own overhead, and the LLLog step list, because the
'            stopwatch lives in the wrapper rather than in the class.
'     KEPT:  every piece of work the walk actually does.
'   LLImporter carries no dialog of its own -- grep answers zero for Messenger,
'   MsgBox and InputBox in that class -- so nothing here can put a box on a
'   screen nobody is watching, and no messenger arming is needed.
'
'   WHAT THIS DOES NOT DO THAT PRODUCTION DOES
'   ---------------------------------------------------------------------------
'   The wrapper opens with ApplicationState.ApplyBusyState, minimizes the import
'   window, and RunImportData adds LLEnterQuietState, TrimDataTables and
'   UpdateAllListAuto around it. None of those is timed here. The trim in
'   particular changes how many rows the walk then writes, so read these numbers
'   as the cost of the import work itself, not as a stopwatch on a user's click.
'
'   SECONDS GO THROUGH Str$, NEVER Format$: this box is en_FR and Format$ writes
'   a comma here and a period on an English Windows machine, and two runs are
'   read side by side.
'@depends LLImporter, LLExporter, ImportMetadata, LLLog, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLImportTiming"

'The real linelist and the password it carries, the same pair
'TestExportOtherLinelist resolves.
Private Const INPUT_FOLDER As String = ".input"
Private Const INPUT_FILE As String = "old_linelist.xlsb"
Private Const INPUT_PASSWORD As String = "5678"

Private Const FILES_FOLDER As String = "ImportTimingFiles"

'The text LLLog writes a step list under, and the two actions that write one.
Private Const STEP_MARKER As String = "step times"
Private Const IMPORT_ACTION As String = "import-data"
Private Const EXPORT_ACTION As String = "export-migration"

Private Const LOG_SHEET As String = "__log"

Private Assert As CustomTest

'The workbooks that were open before this module ran. Only what appears after
'this is swept, so a workbook belonging to the harness or to a module that ran
'earlier is never closed by this one.
Private baselineWorkbooks As BetterArray


'@section Module lifecycle
'===============================================================================

'@sub-title Build the assertion harness.
'@details Public because the harness calls it by name through Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLImportTiming"

    CaptureWorkbookBaseline
End Sub

'@sub-title Records what was already open, so the sweep only takes what this
'   module added.
Private Sub CaptureWorkbookBaseline()
    Dim wkb As Workbook

    Set baselineWorkbooks = New BetterArray

    On Error Resume Next
    For Each wkb In Application.Workbooks
        baselineWorkbooks.Push wkb.Name
    Next wkb
    On Error GoTo 0
End Sub

'@sub-title Put back what the WALK changed, then report.
'@details
'THIS ROUTINE IS WHY A FAILING TEST HERE REPORTS INSTEAD OF KILLING THE RUN.
'The import writes thousands of rows into a real linelist, and a raise part way
'through can leave application flags where it left them. Flags left wrong break
'the harness itself rather than one test: a run answered -50 with a zero-byte CSV
'and no result for ANY module, which reads exactly like the flaky wedge and is
'not one. So the flags go back by hand, the workbooks this module opened are
'swept, and the driver is brought to the front before anything is printed.
'@ModuleCleanup
Public Sub ModuleCleanup()
    SweepStrayWorkbooks

    'By hand, not RestoreApp: RestoreApp undoes what BusyApp did, and what needs
    'undoing here is what the walk did. The calculation mode is deliberately NOT
    'touched -- forcing it to automatic here made every module that ran after
    'this one recalculate on every write, and the run died in the NEXT module.
    On Error Resume Next
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    BringDriverToFront

    'PrintResults is the last thing that can lose a whole module's results, so
    'it is guarded rather than left to raise.
    On Error Resume Next
    If Not Assert Is Nothing Then Assert.PrintResults TEST_OUTPUT_SHEET
    On Error GoTo 0

    Application.EnableEvents = True
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Hands the screen back to the driver, unminimized.
'@details
'THE ACTIVATE ALONE IS NOT ENOUGH, and this is what cost three runs.
'Both walks minimize a window on purpose: LLExporter.CreateOutputWorkbook does
'`ActiveWindow.WindowState = xlMinimized` on the output workbook, and
'HandleImportData does the same to the import file it opens. When those
'workbooks close, the driver can be left as the active workbook with a
'MINIMIZED window, and PrintResults cannot write to that -- "Method
'PrintResults of Object CustomTest failed". Activate puts the workbook in
'front; it does not un-minimize the window, so the window is asked separately.
Private Sub BringDriverToFront()
    On Error Resume Next
    ThisWorkbook.Activate

    If ThisWorkbook.Windows.Count > 0 Then
        ThisWorkbook.Windows(1).Visible = True
        If ThisWorkbook.Windows(1).WindowState = xlMinimized Then _
            ThisWorkbook.Windows(1).WindowState = xlNormal
        ThisWorkbook.Windows(1).Activate
    End If
    On Error GoTo 0
End Sub

'@sub-title Closes the workbooks THIS MODULE opened, and only those.
'@details
'HandleImportData opens the import file itself and closes it on its way out. A
'walk that raised leaves it open, holding the screen.
'The list is a BASELINE taken in ModuleInitialize, the shape
'TestExportOtherLinelist uses, rather than "everything but ThisWorkbook".
'Closing everything else took down workbooks this module never opened, and the
'run then died in the module that ran AFTER this one.
Private Sub SweepStrayWorkbooks()
    Dim counter As Long
    Dim wkbName As String

    If baselineWorkbooks Is Nothing Then Exit Sub

    For counter = Application.Workbooks.Count To 1 Step -1
        wkbName = vbNullString
        On Error Resume Next
        wkbName = Application.Workbooks(counter).Name
        On Error GoTo 0

        If LenB(wkbName) > 0 Then
            If StrComp(wkbName, ThisWorkbook.Name, vbTextCompare) <> 0 Then
                If Not baselineWorkbooks.Includes(wkbName) Then
                    On Error Resume Next
                    Application.Workbooks(counter).Close savechanges:=False
                    On Error GoTo 0
                End If
            End If
        End If
    Next counter
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanUp
Public Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.FlushCurrentTest
End Sub


'@section The walk
'===============================================================================

'@sub-title Times an export and the import of what it wrote, on a real linelist.
'@details
'The export runs first because the import needs a file to read. Both step lists
'land on the same __log sheet and both are reported, so the two walks of this
'block can be read against each other on the same data for the first time.
'@TestMethod("LLImportTiming")
Public Sub TestTimeImportWalkOnARealLinelist()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeImportWalkOnARealLinelist"
    On Error GoTo TestFail

    Dim llBook As Workbook
    Dim impwb As Workbook
    Dim exporter As LLExporter
    Dim impObj As LLImporter
    Dim meta As ImportMetadata
    Dim inputPath As String
    Dim exportPath As String
    Dim refused As Boolean
    Dim sameLanguage As Boolean
    Dim startedAt As Double
    Dim exportSeconds As Double
    Dim failedNumber As Long
    Dim failedText As String

    'What "replace" means to the walk: the rows this linelist holds go before
    'the file's rows are written. It is the heavier of the two rules, so it is
    'the one worth timing, and it exercises ClearData, which Session 121 rewrites.
    Const pasteAtBottom As Boolean = False

    inputPath = ResolvedInputPath()
    If LenB(inputPath) = 0 Then
        'Not a failure of the code under test. The file is gitignored, so a
        'checkout without it reports this rather than an opaque raise.
        Assert.LogSuccesses "TIMING SKIPPED: " & INPUT_FILE & " is not on this " & _
                            "machine, so there is no real linelist to time"
        Exit Sub
    End If

    Set llBook = Workbooks.Open(fileName:=inputPath, password:=INPUT_PASSWORD, _
                                  UpdateLinks:=0)

    'The export half. It is timed here as well so the pair travels together:
    'a later run comparing import numbers wants to know the machine was in the
    'same mood when the export ran.
    Set exporter = LLExporter.Create(llBook)
    startedAt = Timer
    exportPath = exporter.ExportMigration(BuildTempFolder(ThisWorkbook, FILES_FOLDER))
    exportSeconds = ElapsedSince(startedAt)

    Assert.IsTrue (LenB(exportPath) > 0), _
                  "The export wrote a file, so the import has something to read"

    LogLinelistShape llBook
    LogTiming "export migration, whole walk, REAL linelist", exportSeconds
    LogStepLists llBook

    'THE CLASS METHODS, IN THE ORDER THE WALK CALLS THEM, NOT THE WALK ITSELF.
    'LinelistRun.HandleImportData could not be driven from inside the headless
    'harness: the run answered -50 with a zero-byte CSV and no result for any
    'module, and narrowed to this module alone it returned with zero result
    'rows, meaning the test never came back from the call and TestCleanup never
    'flushed what had already been logged. Same shape as the GenerationHost.Run
    'guard already in this suite.
    'What is lost by going one level down: the wrapper's own overhead, and the
    'LLLog step list, since the stopwatch lives in the wrapper. What is kept is
    'every piece of work the walk actually does, timed one call at a time, which
    'is the number this block wants. LLImporter carries no dialog of its own --
    'grep answers zero for Messenger, MsgBox and InputBox in that class -- so
    'nothing here can put a box on a screen nobody is watching.
    Set impwb = Workbooks.Open(fileName:=exportPath, ReadOnly:=True, UpdateLinks:=0)
    Set impObj = LLImporter.Create(llBook)

    startedAt = Timer
    Set meta = ImportMetadata.Create(impwb)
    refused = Not impObj.CheckImportFile(impwb, meta)
    LogTiming "1 reading what the file says about itself", ElapsedSince(startedAt)

    Assert.IsFalse refused, _
                   "The file this linelist just exported is one it accepts, so " & _
                   "the steps below timed real work rather than a refusal"

    sameLanguage = impObj.HasSameLanguage(meta)
    Assert.LogSuccesses "TIMING same language: " & CStr(sameLanguage) & _
                        " (False means the last three steps did not run)"

    startedAt = Timer
    impObj.ImportData impwb, pasteAtBottom, meta
    LogTiming "2 importing the data", ElapsedSince(startedAt)

    startedAt = Timer
    impObj.ImportCustomDropdown impwb, pasteAtBottom
    LogTiming "3 importing the custom dropdowns", ElapsedSince(startedAt)

    startedAt = Timer
    impObj.CompareWithImportFile impwb
    LogTiming "4 comparing with the import file", ElapsedSince(startedAt)

    startedAt = Timer
    impObj.FinalizeReport
    LogTiming "5 finishing the report", ElapsedSince(startedAt)

    If sameLanguage Then
        startedAt = Timer
        impObj.ImportShowHide impwb, meta
        LogTiming "6 importing the show/hide choices", ElapsedSince(startedAt)

        startedAt = Timer
        impObj.ImportEditableLabels impwb, meta
        LogTiming "7 importing the editable labels", ElapsedSince(startedAt)

        startedAt = Timer
        impObj.ImportSingleValues meta
        LogTiming "8 importing the single values", ElapsedSince(startedAt)
    End If

    impwb.Close savechanges:=False
    Set impwb = Nothing

    DropArtefacts llBook, exporter, exportPath
    Exit Sub
TestFail:
    failedNumber = Err.Number
    failedText = Err.Description
    If Not exporter Is Nothing Then _
        failedText = failedText & " | exporter said: " & exporter.LastFailure
    On Error Resume Next
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    DropArtefacts llBook, exporter, exportPath
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTimeImportWalkOnARealLinelist", failedNumber, failedText
End Sub


'@section Reporting
'===============================================================================

'@sub-title Writes one timing line into the results.
'@param label String. What was measured.
'@param elapsed Double. Seconds it took.
Private Sub LogTiming(ByVal label As String, ByVal elapsed As Double)
    Assert.LogSuccesses "TIMING " & label & ": " & SecondsText(elapsed) & " s"
End Sub

'@sub-title Writes the size of the linelist beside the numbers.
'@details
'A timing line means nothing without the shape it was measured on. The owner
'asked for the row count of each data sheet and the variable count, so those are
'what this writes.
'@param llBook Workbook. The linelist being timed.
Private Sub LogLinelistShape(ByVal llBook As Workbook)
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim shapeText As String
    Dim rowCount As Long

    On Error Resume Next

    For Each sh In llBook.Worksheets
        If sh.ListObjects.Count > 0 Then
            Set lo = sh.ListObjects(1)
            rowCount = 0
            If Not lo.DataBodyRange Is Nothing Then rowCount = lo.DataBodyRange.Rows.Count
            shapeText = shapeText & sh.Name & "=" & rowCount & " "
        End If
    Next sh

    'The two counts the baseline asks for beside the rows per sheet: the
    'dictionary is not a table, so it is read off its used range, header
    'row excluded. Added for the block's baseline run of 2026-09-01.
    Dim dictRows As Long
    dictRows = llBook.Worksheets("Dictionary").UsedRange.Rows.Count - 1
    shapeText = shapeText & "| dictionary rows=" & dictRows & " "

    On Error GoTo 0

    Assert.LogSuccesses "TIMING linelist shape (table rows per sheet): " & _
                        Trim$(shapeText) & "| platform " & PlatformTag()
End Sub

'@sub-title Reads every step list off the linelist log and reports it.
'@details
'The column is not hard-coded. An earlier probe of this block looked in column 3
'when the entry detail lands in column 5, reported every other number, raised
'nothing, and simply had no step list -- a missing measurement that reads as a
'clean run. The whole used range is scanned, and a run that finds nothing says
'so out loud.
'@param llBook Workbook. The workbook whose log to read.
Private Sub LogStepLists(ByVal llBook As Workbook)
    Dim sh As Worksheet
    Dim usedRng As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim detail As String
    Dim foundImport As Boolean
    Dim foundExport As Boolean

    On Error Resume Next
    Set sh = llBook.Worksheets(LOG_SHEET)
    On Error GoTo 0

    If sh Is Nothing Then
        Assert.LogSuccesses "TIMING steps: the linelist carries no " & LOG_SHEET & " sheet"
        Exit Sub
    End If

    Set usedRng = sh.UsedRange
    If usedRng Is Nothing Then Exit Sub

    For rowIndex = usedRng.Row To usedRng.Row + usedRng.Rows.Count - 1
        For colIndex = usedRng.Column To usedRng.Column + usedRng.Columns.Count - 1
            detail = CStr(sh.Cells(rowIndex, colIndex).Value)
            If InStr(1, detail, STEP_MARKER, vbTextCompare) > 0 Then
                Assert.LogSuccesses "TIMING steps: " & detail
                If InStr(1, detail, IMPORT_ACTION, vbTextCompare) > 0 Then foundImport = True
                If InStr(1, detail, EXPORT_ACTION, vbTextCompare) > 0 Then foundExport = True
            End If
        Next colIndex
    Next rowIndex

    If Not foundImport Then _
        Assert.LogSuccesses "TIMING steps: NO " & IMPORT_ACTION & " list on the log"
    If Not foundExport Then _
        Assert.LogSuccesses "TIMING steps: NO " & EXPORT_ACTION & " list on the log"
End Sub

'@fun-title Seconds as text, locale-independent.
'@details Format$ writes a comma on an en_FR box and a period on an English one,
'   and two runs are read side by side. Str$ always writes a period, and drops
'   the leading zero of a value under one, which is put back.
'@param seconds Double. The reading.
'@return String. The seconds, three decimals, always with a period.
Private Function SecondsText(ByVal seconds As Double) As String
    Dim text As String

    text = Trim$(Str$(Int(seconds * 1000 + 0.5) / 1000))
    If Left$(text, 1) = "." Then text = "0" & text
    If Left$(text, 2) = "-." Then text = "-0" & Mid$(text, 2)

    SecondsText = text
End Function

'@fun-title Seconds since a Timer reading, safe across midnight.
'@details Timer restarts at midnight, so a walk running across it reads as a
'   negative difference and a day is added back on.
'@param startedAt Double. The earlier Timer reading.
'@return Double. Seconds elapsed.
Private Function ElapsedSince(ByVal startedAt As Double) As Double
    Dim elapsed As Double

    elapsed = CDbl(Timer) - startedAt
    If elapsed < 0 Then elapsed = elapsed + 86400#

    ElapsedSince = elapsed
End Function

'@fun-title Which box the number was measured on.
'@return String. "macOS" or "Windows".
Private Function PlatformTag() As String
    #If Mac Then
        PlatformTag = "macOS"
    #Else
        PlatformTag = "Windows"
    #End If
End Function


'@section The file
'===============================================================================

'@fun-title Where the real linelist sits, or an empty string.
'@details
'The run dir and the repo are both looked in, the same candidates
'TestExportOtherLinelist walks. The file is gitignored, so an empty answer is an
'ordinary outcome on a machine that does not carry it, not a fault.
'@return String. The path, or an empty string when the file is not there.
Private Function ResolvedInputPath() As String
    Dim candidates As BetterArray
    Dim candidate As String
    Dim counter As Long

    Set candidates = New BetterArray
    candidates.Push JoinPath(ThisWorkbook.Path, "tests", INPUT_FOLDER, INPUT_FILE)
    candidates.Push JoinPath(RepoRoot(), "src", "tests", INPUT_FOLDER, INPUT_FILE)
    candidates.Push JoinPath(ThisWorkbook.Path, INPUT_FOLDER, INPUT_FILE)

    For counter = candidates.LowerBound To candidates.UpperBound
        candidate = CStr(candidates.Item(counter))
        If LenB(Dir$(candidate)) > 0 Then
            ResolvedInputPath = candidate
            Exit Function
        End If
    Next counter
End Function

'@sub-title Closes the linelist and removes the file the export wrote.
'@details
'The linelist is closed WITHOUT saving, which is what puts back everything the
'import just wrote into it. Nothing this probe does reaches disk.
'@param llBook Workbook. The linelist, or Nothing.
'@param exporter LLExporter. The exporter to close, or Nothing.
'@param exportPath String. The file the export wrote, or an empty string.
Private Sub DropArtefacts(ByVal llBook As Workbook, _
                          ByVal exporter As LLExporter, _
                          ByVal exportPath As String)
    On Error Resume Next
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not llBook Is Nothing Then llBook.Close savechanges:=False
    If LenB(exportPath) > 0 Then
        If Dir$(exportPath) <> vbNullString Then Kill exportPath
    End If
    On Error GoTo 0
End Sub
