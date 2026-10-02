Attribute VB_Name = "TestExportOtherLinelist"
Attribute VB_Description = "Step by step walk of the other-linelist export against a real linelist"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Step by step walk of the other-linelist export against a real linelist")

Option Explicit

'@description
'Walks the other-linelist export of F_ExportMig one step at a time against a
'real linelist that the path is known to fail on. FormLogicExportMig
'.OtherLinelistWalk does five things in order -- open the file, read it, write
'the migration file, write the geobase, write the historic geobase -- and a
'failure anywhere in that run reaches the user as one message box that names
'no step. Each step is one test here, so the report says which step refused.
'
'THE FILE UNDER TEST
'-------------------------------------------------------------------------------
'src/tests/.input/old_linelist.xlsb, a measles linelist with data in it, opened
'with the password below. The file is encrypted (the container reads as CDFV2
'Encrypted, so nothing can be read out of it without Excel and the password)
'and it is gitignored, travelling through the push and pull asset scripts
'instead. src/tests/.input reaches the run dir on its own: run-tests.R lists
'subfolders with list.dirs, which includes dotted names, and copies them with
'all.files = TRUE.
'
'A MISSING FILE IS NOT A FAILING EXPORT
'-------------------------------------------------------------------------------
'The first test is the one that says whether the file is there at all. When it
'is absent every later test says the same thing and says it as a skip, so a
'machine that has never pulled the asset does not read as a broken export.
'
'EVENTS ARE OFF FOR EVERY OPEN
'-------------------------------------------------------------------------------
'A linelist carries a Workbook_Open handler, and it builds managers and can put
'a dialog on the screen. A dialog in a headless run takes the whole run down
'with it, so every open here runs with Application.EnableEvents False. BusyApp
'does not cover that -- it sets ScreenUpdating, DisplayAlerts, Calculation and
'EnableAnimations and leaves events alone -- so this module holds events itself.
'
'WHAT THE PATH ANSWERS WHEN IT FAILS
'-------------------------------------------------------------------------------
'LLExporter.LastFailure is filled by ExportGeo alone. ExportMigration clears it
'on entry and never sets it, so a migration that raises answers an empty
'LastFailure and the form logs a description that has already crossed out of the
'class. Each test here reads the raise at the call instead, so the number and
'the text reach the report.
'@depends LLExporter, LLGeo, HiddenNames, CustomTest, BetterArray

Private Assert As CustomTest

'The file, resolved once by ResolvedInputPath and held for every test.
Private inputPathValue As String
Private inputPathTried As Boolean
Private inputPathCandidates As String

'The exporter the factory built, opened once and shared by the read and write
'steps. The two fields beside it carry what the open raised, so a test that
'depends on the open reports the open's own failure rather than error 91.
Private sharedExporter As LLExporter
Private openTried As Boolean
Private openError As Long
Private openMessage As String

'Every file an export step wrote, deleted in ModuleCleanup.
Private writtenFiles As BetterArray

'The workbooks that were open before this module ran. Anything open later and
'not on this list belongs to a test, and SweepStrayWorkbooks closes it.
Private baselineWorkbooks As BetterArray

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ExportOtherLinelist"

'The linelist under test and the password it opens with.
Private Const INPUT_FILE As String = "old_linelist.xlsb"
Private Const INPUT_PASSWORD As String = "5678"
Private Const INPUT_FOLDER As String = ".input"

'Where the export steps put the files they write.
'
'THERE ARE TWO, AND THAT IS THE POINT
'-------------------------------------------------------------------------------
'LLExport.ExportFileName stamps the name to the MINUTE. The two migration tests
'ask for the same export of the same file, so inside one minute they compose the
'same name, and the second SaveAs lands on the file the first one wrote and
'raises 1004. It passed or failed depending on which side of a minute boundary
'the run fell -- green by luck. A folder each removes the collision instead of
'waiting on the clock.
Private Const OUT_FOLDER As String = "ExportOtherFiles"
Private Const OUT_FOLDER_SECOND As String = "ExportOtherFiles2"

'The five sheets LLExporter.CreateFromFile requires before it hands back an
'exporter. The temp sheet answers to either name, since a file may be older
'than the internal-sheet rename.
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const EXPORTS_SHEET As String = "Exports"
Private Const PASS_SHEET As String = "__pass"
Private Const GEO_SHEET As String = "Geo"
Private Const TEMP_SHEET As String = "__temp"
Private Const TEMP_SHEET_OLD As String = "temp__"

'The sheet LLTranslation.Create is built over
Private Const LINELIST_TRANSLATION_SHEET As String = "LinelistTranslation"


'@section Lifecycle
'===============================================================================

'@sub-title Build the harness and put the application into its test state.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run. Nothing is opened here: the open is a step under test and it
'belongs in a test that can report its own failure.
'@ModuleInitialize
Public Sub ModuleInitialize()

    BusyApp
    Application.EnableEvents = False

    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestExportOtherLinelist"

    Set writtenFiles = New BetterArray
    RecordOpenWorkbooks

    inputPathTried = False
    inputPathValue = vbNullString
    openTried = False
    openError = 0
    openMessage = vbNullString
End Sub

'@sub-title Close what was opened, delete what was written, print the results.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run. CloseAll is what closes the linelist the factory opened, and
'it is called before the results are printed so PrintResults writes onto a
'ThisWorkbook that holds the screen.
'@ModuleCleanup
Public Sub ModuleCleanup()

    On Error Resume Next
        If Not sharedExporter Is Nothing Then sharedExporter.CloseAll
    On Error GoTo 0

    Set sharedExporter = Nothing

    'Anything the exports left standing goes now. CloseAll drops the ONE output
    'workbook the exporter still points at, and an export that raised part way
    'through had its output workbook replaced by the next call, so the earlier
    'ones are open with nothing referring to them.
    SweepStrayWorkbooks

    DeleteWrittenFiles

    'BuildTempFolder makes the folder it answers, so the folder is this
    'module's to remove. RmDir fails quietly when it is absent or still holds
    'a file, and a file left behind is reported by DeleteWrittenFiles already.
    On Error Resume Next
        RmDir ThisWorkbook.Path & Application.PathSeparator & OUT_FOLDER
        RmDir ThisWorkbook.Path & Application.PathSeparator & OUT_FOLDER_SECOND
    On Error GoTo 0

    'ThisWorkbook has to hold the screen for PrintResults, which raises 1004
    'whenever another workbook has it.
    On Error Resume Next
        ThisWorkbook.Activate
    On Error GoTo 0

    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    Application.EnableEvents = True
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Put the application into its test state.
'@details
'Events go off again on every test. An export walk opens and closes workbooks
'of its own, and a step that raised part way through may have left the flag
'wherever it was.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    Application.EnableEvents = False
End Sub

'@sub-title Flush the results of the test that just ran.
'@details
'The stray sweep runs first. An export step that raised leaves its output
'workbook open, and one of those holding the screen is what makes
'CustomTest.PrintResults raise 1004 at the end of the module. Sweeping after
'every test keeps at most one of them alive at a time.
'@TestCleanup
Public Sub TestCleanup()
    SweepStrayWorkbooks

    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Step 1 -- the file
'===============================================================================

'@sub-title The linelist under test is on disk where a run can reach it.
'@details
'Every later step needs this file. When it is missing the report says so once
'here and says it as a skip everywhere else, and the message names every path
'that was tried so an operator can see where the asset should land.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheMeaslesLinelistIsOnDisk()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheMeaslesLinelistIsOnDisk"
    On Error GoTo TestFail

    Dim resolvedPath As String

    resolvedPath = ResolvedInputPath()

    Assert.IsTrue LenB(resolvedPath) > 0, _
                  "The linelist under test is on disk. Tried: " & inputPathCandidates

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheMeaslesLinelistIsOnDisk", Err.Number, Err.Description
End Sub


'@section Step 2 -- the open
'===============================================================================

'@sub-title Excel opens the file with the password the user gave.
'@details
'The raw open, with no factory in the way. This separates a password or a
'container Excel refuses from a workbook Excel opens and the factory then
'rejects, which is the next step. A wrong password is not tested: an explicitly
'wrong one raises here, but a prompt instead of a raise would be a modal, and a
'modal takes a headless run down.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheFileOpensWithItsPassword()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFileOpensWithItsPassword"
    On Error GoTo TestFail

    Dim probeWkb As Workbook
    Dim openedHere As Boolean
    Dim raisedNumber As Long
    Dim raisedMessage As String

    If SkipWhenFileMissing("TestTheFileOpensWithItsPassword") Then Exit Sub

    On Error Resume Next
        Set probeWkb = OpenInputWorkbook(openedHere)
        raisedNumber = Err.Number
        raisedMessage = Err.Description
    On Error GoTo TestFail

    Assert.IsNotNothing probeWkb, _
                       "The file opens with its password - raise was " & _
                       CStr(raisedNumber) & " [" & raisedMessage & "]"

    CloseInputWorkbook probeWkb, openedHere

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFileOpensWithItsPassword", Err.Number, Err.Description
End Sub

'@sub-title The opened workbook carries every sheet the factory demands.
'@details
'CreateFromFile refuses a workbook missing Dictionary, Exports, __pass, Geo or
'a temp sheet under either name, and the refusal names one sheet and stops. An
'old linelist can be missing more than one, so this reads them all and names
'every one that is absent. The factory is not involved: this is the raw
'workbook, so a refusal in the next step can be read against this list.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheOpenedFileCarriesEverySheetTheFactoryDemands()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheOpenedFileCarriesEverySheetTheFactoryDemands"
    On Error GoTo TestFail

    Dim probeWkb As Workbook
    Dim openedHere As Boolean
    Dim missingNames As String

    If SkipWhenFileMissing("TestTheOpenedFileCarriesEverySheetTheFactoryDemands") Then Exit Sub

    On Error Resume Next
        Set probeWkb = OpenInputWorkbook(openedHere)
    On Error GoTo TestFail

    If probeWkb Is Nothing Then
        Assert.Fail "The file did not open, so its sheets were not read. " & _
                    "TestTheFileOpensWithItsPassword carries the reason."
        Exit Sub
    End If

    missingNames = MissingRequiredSheets(probeWkb)

    Assert.AreEqual vbNullString, missingNames, _
                    "Every sheet the factory demands is in the file"

    'The sheet list of the file, whatever the outcome, so the report says what
    'the linelist actually holds rather than only what it lacks.
    Assert.LogSuccesses "The file holds these worksheets: " & WorksheetNamesOf(probeWkb)

    CloseInputWorkbook probeWkb, openedHere

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheOpenedFileCarriesEverySheetTheFactoryDemands", _
                         Err.Number, Err.Description
End Sub


'@section Step 3 -- the factory
'===============================================================================

'@sub-title CreateFromFile hands back an exporter bound to the file.
'@details
'The seam OtherLinelistWalk arms its own ErrOpen handler around. A raise here
'is the path telling the user to check the file and the password; a raise after
'here is an export failing, and the two reach the user as different messages.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheFactoryBuildsAnExporterOnTheFile()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFactoryBuildsAnExporterOnTheFile"
    On Error GoTo TestFail

    Dim exporter As LLExporter

    If SkipWhenFileMissing("TestTheFactoryBuildsAnExporterOnTheFile") Then Exit Sub

    Set exporter = EnsureExporter()

    Assert.IsNotNothing exporter, _
                        "CreateFromFile builds an exporter - raise was " & _
                        CStr(openError) & " [" & openMessage & "]"

    If exporter Is Nothing Then Exit Sub

    Assert.IsTrue exporter.OpenedFromFile, _
                  "The factory opened the file itself, so CloseAll owns it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFactoryBuildsAnExporterOnTheFile", _
                         Err.Number, Err.Description
End Sub


'@section Step 4 -- the reads an export depends on
'===============================================================================

'@sub-title The epiweek value of the file is readable.
'@details
'One of the three read seams of LLExporter, and the one AddMetadataTags writes
'into every export file. It writes nothing, so a raise here is the file
'refusing a read rather than an export refusing to run. An empty answer is a
'pass: a linelist that names no epiweek start is a linelist, and the tag it
'writes is then blank by design.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheEpiweekValueOfTheFileIsReadable()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheEpiweekValueOfTheFileIsReadable"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim epiweek As String
    Dim raisedNumber As Long
    Dim raisedMessage As String

    If SkipWhenNoExporter("TestTheEpiweekValueOfTheFileIsReadable", exporter) Then Exit Sub

    On Error Resume Next
        epiweek = exporter.EpiWeekStart()
        raisedNumber = Err.Number
        raisedMessage = Err.Description
    On Error GoTo TestFail

    Assert.AreEqual 0&, raisedNumber, _
                    "Reading the epiweek start raises nothing - text was [" & _
                    raisedMessage & "]"
    Assert.LogSuccesses "The file answers epiweek start [" & epiweek & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheEpiweekValueOfTheFileIsReadable", _
                         Err.Number, Err.Description
End Sub

'@sub-title The options line a migration would write is readable.
'@details
'MigrationOptions reads the Exports sheet and the dictionary of the file and
'writes nothing, so it is the cheapest read that touches both. A migration that
'cannot compose its own options line will not write a file either, and this
'says so without waiting for the write.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheMigrationOptionsLineIsReadable()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheMigrationOptionsLineIsReadable"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim optionsLine As String
    Dim raisedNumber As Long
    Dim raisedMessage As String

    If SkipWhenNoExporter("TestTheMigrationOptionsLineIsReadable", exporter) Then Exit Sub

    On Error Resume Next
        optionsLine = exporter.MigrationOptions()
        raisedNumber = Err.Number
        raisedMessage = Err.Description
    On Error GoTo TestFail

    Assert.AreEqual 0&, raisedNumber, _
                    "Composing the migration options raises nothing - text was [" & _
                    raisedMessage & "]"
    Assert.LogSuccesses "The migration options line reads [" & optionsLine & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheMigrationOptionsLineIsReadable", _
                         Err.Number, Err.Description
End Sub


'@section Step 5 -- the migration file
'===============================================================================

'@sub-title ExportMigration writes a file that lands on disk.
'@details
'The first write of the walk, called with the two options the form passes when
'both boxes are ticked. The path it answers is asserted against the disk: the
'function answers vbNullString on failure and a path on success, and a path
'with no file behind it would be a save that did not happen.
'
'LastFailure is read into the message and is expected to be empty even on a
'failure: ExportMigration clears it on entry and never sets it. The raise
'number captured here is the only thing that names the fault.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheMigrationExportWritesItsFile()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheMigrationExportWritesItsFile"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim savedPath As String
    Dim raisedNumber As Long
    Dim raisedMessage As String

    If SkipWhenNoExporter("TestTheMigrationExportWritesItsFile", exporter) Then Exit Sub

    On Error Resume Next
        savedPath = exporter.ExportMigration(OutputFolder(), _
                                             includeShowHide:=True, _
                                             keepLabels:=True)
        raisedNumber = Err.Number
        raisedMessage = Err.Description
    On Error GoTo TestFail

    Assert.AreEqual 0&, raisedNumber, _
                    "ExportMigration raises nothing - text was [" & raisedMessage & _
                    "], LastFailure was [" & exporter.LastFailure & "]"

    AssertFileWasWritten savedPath, "the migration export"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheMigrationExportWritesItsFile", _
                         Err.Number, Err.Description
End Sub


'@section Step 6 -- the geobase files
'===============================================================================

'@sub-title ExportGeo writes the whole geobase to a file.
'@details
'The one export that fills LastFailure, through its own RecordFailure handler,
'so the message here carries the step name the class was on when it refused.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheGeobaseExportWritesItsFile()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheGeobaseExportWritesItsFile"
    On Error GoTo TestFail

    RunGeoStep onlyHistoric:=False, _
               testName:="TestTheGeobaseExportWritesItsFile", _
               label:="the geobase export"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheGeobaseExportWritesItsFile", _
                         Err.Number, Err.Description
End Sub

'@sub-title ExportGeo writes the historic geobase to a file.
'@details
'The third box of the form, and a separate call on the same exporter. A file
'carrying no historic geobase is a real case, so an empty answer with no raise
'is reported rather than failed.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheHistoricGeobaseExportWritesItsFile()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheHistoricGeobaseExportWritesItsFile"
    On Error GoTo TestFail

    RunGeoStep onlyHistoric:=True, _
               testName:="TestTheHistoricGeobaseExportWritesItsFile", _
               label:="the historic geobase export"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheHistoricGeobaseExportWritesItsFile", _
                         Err.Number, Err.Description
End Sub


'@section Private -- the geo step both geo tests share
'===============================================================================

'@sub-title Run one ExportGeo call and report what it answered.
'@param onlyHistoric Boolean. Which of the two geobase exports to run.
'@param testName String. The test the report files this under.
'@param label String. What to call this export in the messages.
Private Sub RunGeoStep(ByVal onlyHistoric As Boolean, _
                       ByVal testName As String, _
                       ByVal label As String)

    Dim exporter As LLExporter
    Dim savedPath As String
    Dim raisedNumber As Long
    Dim raisedMessage As String
    Dim recordedFailure As String

    If SkipWhenNoExporter(testName, exporter) Then Exit Sub

    On Error Resume Next
        savedPath = exporter.ExportGeo(OutputFolder(), onlyHistoric:=onlyHistoric)
        raisedNumber = Err.Number
        raisedMessage = Err.Description
        recordedFailure = exporter.LastFailure
    On Error GoTo 0

    Assert.AreEqual 0&, raisedNumber, _
                    "ExportGeo raises nothing for " & label & " - text was [" & _
                    raisedMessage & "], LastFailure was [" & recordedFailure & "]"

    If LenB(savedPath) = 0 And raisedNumber = 0 Then
        Assert.LogSuccesses label & " answered no path and raised nothing, so " & _
                            "this file carries no such geobase"
        Exit Sub
    End If

    AssertFileWasWritten savedPath, label
End Sub


'@section Private -- the file under test
'===============================================================================

'@sub-title The path of the linelist under test, resolved once.
'@details
'Three places are tried, nearest first. The run dir copy comes first because it
'sits inside the folder macOS has already granted Excel access to; the repo
'copy is the fallback for a run driven from somewhere else, and a copy dropped
'beside the driver workbook is the last.
'@return String. The first path that has a file behind it, or vbNullString.
Private Function ResolvedInputPath() As String

    Dim candidates As BetterArray
    Dim onePath As String
    Dim counter As Long

    If inputPathTried Then
        ResolvedInputPath = inputPathValue
        Exit Function
    End If

    inputPathTried = True

    Set candidates = New BetterArray
    candidates.Push JoinPath(ThisWorkbook.Path, "tests", INPUT_FOLDER, INPUT_FILE)
    candidates.Push JoinPath(RepoRoot(), "src", "tests", INPUT_FOLDER, INPUT_FILE)
    candidates.Push JoinPath(ThisWorkbook.Path, INPUT_FOLDER, INPUT_FILE)

    inputPathCandidates = vbNullString

    For counter = candidates.LowerBound To candidates.UpperBound
        onePath = CStr(candidates.Item(counter))

        If LenB(inputPathCandidates) > 0 Then _
            inputPathCandidates = inputPathCandidates & " | "
        inputPathCandidates = inputPathCandidates & onePath

        If LenB(onePath) > 0 Then
            'Dir raises on a path whose folder is not reachable, and an
            'unreachable candidate is simply not the answer.
            On Error Resume Next
            If LenB(Dir$(onePath)) > 0 Then
                inputPathValue = onePath
            End If
            On Error GoTo 0
        End If

        If LenB(inputPathValue) > 0 Then Exit For
    Next counter

    ResolvedInputPath = inputPathValue
End Function

'@sub-title Open the linelist under test, or answer the copy already open.
'@details
'A workbook already open in the session is answered as it stands, which is what
'LLExporter.FindOpenWorkbook does, so a test running after the factory has
'opened the file does not open a second copy and does not close the copy the
'shared exporter is holding.
'@param openedHere Boolean. Receives True when this call is what opened it.
'@return Workbook. The open workbook, or Nothing when the open failed.
Private Function OpenInputWorkbook(ByRef openedHere As Boolean) As Workbook

    Dim wkb As Workbook
    Dim filePath As String

    openedHere = False
    filePath = ResolvedInputPath()
    If LenB(filePath) = 0 Then Exit Function

    Set wkb = OpenWorkbookByPath(filePath)
    If Not wkb Is Nothing Then
        Set OpenInputWorkbook = wkb
        Exit Function
    End If

    'Events are already off for the module, and they are set again here because
    'this is the line a stray dialog would come out of.
    Application.EnableEvents = False

    On Error Resume Next
    Set wkb = Workbooks.Open(fileName:=filePath, _
                             ReadOnly:=True, _
                             password:=INPUT_PASSWORD)
    On Error GoTo 0

    If Not wkb Is Nothing Then openedHere = True
    Set OpenInputWorkbook = wkb
End Function

'@sub-title Close the linelist under test when this test is what opened it.
'@param wkb Workbook. The workbook to close. Nothing is skipped.
'@param openedHere Boolean. True when the caller opened it.
Private Sub CloseInputWorkbook(ByVal wkb As Workbook, ByVal openedHere As Boolean)
    If wkb Is Nothing Then Exit Sub
    If Not openedHere Then Exit Sub

    On Error Resume Next
    wkb.Close savechanges:=False
    On Error GoTo 0
End Sub

'@sub-title The open workbook sitting at a path, when there is one.
'@param filePath String. Full path to look for.
'@return Workbook. The open workbook, or Nothing.
Private Function OpenWorkbookByPath(ByVal filePath As String) As Workbook

    Dim wkb As Workbook

    For Each wkb In Application.Workbooks
        On Error Resume Next
        If StrComp(wkb.FullName, filePath, vbTextCompare) = 0 Then
            Set OpenWorkbookByPath = wkb
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next wkb
End Function


'@section Private -- the shared exporter
'===============================================================================

'@sub-title The exporter on the file, built once and held.
'@details
'The factory is called once for the whole module. Its raise is kept in the two
'module fields so a test that needs the exporter reports the open's own failure
'instead of failing on a Nothing it cannot explain.
'@return LLExporter. The held exporter, or Nothing when the factory refused.
Private Function EnsureExporter() As LLExporter

    Dim filePath As String

    If openTried Then
        Set EnsureExporter = sharedExporter
        Exit Function
    End If

    openTried = True

    filePath = ResolvedInputPath()
    If LenB(filePath) = 0 Then
        openError = -1
        openMessage = "the file under test is not on disk"
        Exit Function
    End If

    Application.EnableEvents = False

    On Error Resume Next
        Set sharedExporter = LLExporter.CreateFromFile(filePath, INPUT_PASSWORD)
        openError = Err.Number
        openMessage = Err.Description
    On Error GoTo 0

    Set EnsureExporter = sharedExporter
End Function


'@section Private -- skips
'===============================================================================

'@sub-title Report a skip when the file under test is absent.
'@param testName String. The test the skip is filed under.
'@return Boolean. True when the caller should stop.
Private Function SkipWhenFileMissing(ByVal testName As String) As Boolean

    If LenB(ResolvedInputPath()) > 0 Then Exit Function

    Assert.LogSuccesses testName & " did not run: the linelist under test is " & _
                        "not on disk. Tried: " & inputPathCandidates
    SkipWhenFileMissing = True
End Function

'@sub-title Report a skip when the exporter could not be built.
'@details
'A step after the factory has nothing to say about a factory that refused, and
'the refusal is already one test's failure. So this reports rather than fails,
'and it names the raise so the report is readable on its own.
'@param testName String. The test the skip is filed under.
'@param exporter LLExporter. Receives the held exporter when there is one.
'@return Boolean. True when the caller should stop.
Private Function SkipWhenNoExporter(ByVal testName As String, _
                                    ByRef exporter As LLExporter) As Boolean

    If SkipWhenFileMissing(testName) Then
        SkipWhenNoExporter = True
        Exit Function
    End If

    Set exporter = EnsureExporter()
    If Not exporter Is Nothing Then Exit Function

    Assert.LogSuccesses testName & " did not run: CreateFromFile refused the " & _
                        "file with error " & CStr(openError) & " [" & _
                        openMessage & "]. " & _
                        "TestTheFactoryBuildsAnExporterOnTheFile carries it."
    SkipWhenNoExporter = True
End Function


'@section Private -- reading the file
'===============================================================================

'@sub-title Every sheet CreateFromFile demands that the workbook does not hold.
'@param wkb Workbook. The workbook to read.
'@return String. The missing names on one line, empty when none are missing.
Private Function MissingRequiredSheets(ByVal wkb As Workbook) As String

    Dim required As BetterArray
    Dim missing As String
    Dim sheetName As String
    Dim counter As Long

    Set required = New BetterArray
    required.Push DICTIONARY_SHEET, EXPORTS_SHEET, PASS_SHEET, GEO_SHEET

    For counter = required.LowerBound To required.UpperBound
        sheetName = CStr(required.Item(counter))
        If Not SheetIsInWorkbook(wkb, sheetName) Then _
            missing = AppendName(missing, sheetName)
    Next counter

    'The temp sheet answers to either name, since the file may be older than
    'the internal-sheet rename.
    If Not SheetIsInWorkbook(wkb, TEMP_SHEET) Then
        If Not SheetIsInWorkbook(wkb, TEMP_SHEET_OLD) Then _
            missing = AppendName(missing, TEMP_SHEET & " or " & TEMP_SHEET_OLD)
    End If

    MissingRequiredSheets = missing
End Function

'@sub-title Whether a workbook holds a worksheet of that name.
'@param wkb Workbook. The workbook to read.
'@param sheetName String. The name to look for.
'@return Boolean. True when the sheet is there.
Private Function SheetIsInWorkbook(ByVal wkb As Workbook, _
                                   ByVal sheetName As String) As Boolean

    Dim foundSheet As Worksheet

    'A missing sheet raises on the read, and Nothing is the answer read back.
    On Error Resume Next
    Set foundSheet = wkb.Worksheets(sheetName)
    On Error GoTo 0

    SheetIsInWorkbook = Not (foundSheet Is Nothing)
End Function

'@sub-title The worksheet names of a workbook on one line.
'@param wkb Workbook. The workbook to read.
'@return String. The names, comma separated.
Private Function WorksheetNamesOf(ByVal wkb As Workbook) As String

    Dim sh As Worksheet
    Dim names As String

    For Each sh In wkb.Worksheets
        names = AppendName(names, sh.Name)
    Next sh

    WorksheetNamesOf = names
End Function

'@sub-title Add one name to a comma separated line.
'@param collected String. The line so far.
'@param oneName String. The name to add.
'@return String. The line with the name on the end.
Private Function AppendName(ByVal collected As String, _
                            ByVal oneName As String) As String
    If LenB(collected) = 0 Then
        AppendName = oneName
    Else
        AppendName = collected & ", " & oneName
    End If
End Function


'@section Private -- the files the exports write
'===============================================================================

'@sub-title The folder the export steps write into.
'@return String. The folder path, created when it is not there.
Private Function OutputFolder() As String
    OutputFolder = BuildTempFolder(ThisWorkbook, OUT_FOLDER)
End Function

'@sub-title The folder the second migration run writes into.
'@details
'Its own folder, so the two migration runs cannot compose the same path inside
'one minute. See the note on the two constants.
'@return String. The folder path, created when it is not there.
Private Function SecondOutputFolder() As String
    SecondOutputFolder = BuildTempFolder(ThisWorkbook, OUT_FOLDER_SECOND)
End Function

'@sub-title Assert that an export answered a path and put a file behind it.
'@param savedPath String. What the export answered.
'@param label String. What to call this export in the messages.
Private Sub AssertFileWasWritten(ByVal savedPath As String, ByVal label As String)

    Assert.IsTrue LenB(savedPath) > 0, _
                  label & " answers the path it saved"

    If LenB(savedPath) = 0 Then Exit Sub

    writtenFiles.Push savedPath

    Assert.IsTrue FileIsOnDisk(savedPath), _
                  label & " left a file at the path it answered [" & savedPath & "]"
End Sub

'@sub-title Whether a file sits at a path.
'@param filePath String. The path to look at.
'@return Boolean. True when a file is there.
Private Function FileIsOnDisk(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileIsOnDisk = (LenB(Dir$(filePath)) > 0)
    On Error GoTo 0
End Function

'@sub-title Note which workbooks were already open before this module ran.
'@details
'Names rather than references: a workbook closed by something else leaves a
'reference that raises on every read, and the sweep only needs to recognise a
'name it must not touch.
Private Sub RecordOpenWorkbooks()

    Dim wkb As Workbook

    Set baselineWorkbooks = New BetterArray

    On Error Resume Next
    For Each wkb In Application.Workbooks
        baselineWorkbooks.Push wkb.Name
    Next wkb
    On Error GoTo 0
End Sub

'@sub-title Close every workbook a test opened and did not close.
'@details
'Three are spared: the driver, whatever was already open before the module ran,
'and the linelist under test while the held exporter is still using it. The
'walk is backwards over the index because closing a workbook renumbers the
'collection under a For Each.
Private Sub SweepStrayWorkbooks()

    Dim counter As Long
    Dim wkbName As String
    Dim sourceName As String

    If baselineWorkbooks Is Nothing Then Exit Sub

    'The source is spared while an exporter still holds it. Reading the name
    'off a closed workbook raises, and an empty name spares nothing.
    On Error Resume Next
    If Not sharedExporter Is Nothing Then _
        sourceName = sharedExporter.SourceWorkbook.Name
    On Error GoTo 0

    For counter = Application.Workbooks.count To 1 Step -1

        wkbName = vbNullString
        On Error Resume Next
        wkbName = Application.Workbooks(counter).Name
        On Error GoTo 0

        If LenB(wkbName) > 0 Then
            If StrComp(wkbName, ThisWorkbook.Name, vbTextCompare) <> 0 Then
                If StrComp(wkbName, sourceName, vbTextCompare) <> 0 Then
                    If Not baselineWorkbooks.Includes(wkbName) Then
                        On Error Resume Next
                        Application.Workbooks(counter).Close savechanges:=False
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next counter
End Sub

'@sub-title Delete every file the export steps wrote.
'@details
'A suite cleans up what it writes. Each Kill is guarded on its own so one file
'Excel still holds does not leave the rest behind.
Private Sub DeleteWrittenFiles()

    Dim counter As Long
    Dim onePath As String

    If writtenFiles Is Nothing Then Exit Sub

    For counter = writtenFiles.LowerBound To writtenFiles.UpperBound
        onePath = CStr(writtenFiles.Item(counter))

        On Error Resume Next
        If LenB(Dir$(onePath)) > 0 Then Kill onePath
        On Error GoTo 0
    Next counter

    Set writtenFiles = Nothing
End Sub


'@section Step 7 -- what the geo manager refuses
'===============================================================================
'@description
'Both geobase exports answered error 1007, ElementNotFound, and ExportGeo's own
'RecordFailure named the step: reading the Geo sheet, out of LLGeo.Create. That
'is one call, and it checks nine tables, four hidden names and one named range
'before it hands anything back. The three tests here read those requirements off
'the file so the report names the missing element rather than the error number.

'@sub-title The Geo sheet carries every table the geo manager demands.
'@details
'LLGeo.CheckRequirements calls LoExists over nine names, and LoExists stops on
'the first one it cannot find. This reads all nine, so one run says everything
'that is absent, and it logs the tables the sheet does hold beside them.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheGeoSheetCarriesEveryTableTheGeoManagerNeeds()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheGeoSheetCarriesEveryTableTheGeoManagerNeeds"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim geoSh As Worksheet
    Dim required As BetterArray
    Dim missing As String
    Dim tableName As String
    Dim counter As Long

    If SkipWhenNoExporter("TestTheGeoSheetCarriesEveryTableTheGeoManagerNeeds", exporter) Then Exit Sub

    Set geoSh = GeoSheetOf(exporter)
    If geoSh Is Nothing Then
        Assert.Fail "The file carries no Geo worksheet"
        Exit Sub
    End If

    Set required = New BetterArray
    required.Push "T_ADM1", "T_ADM2", "T_ADM3", "T_ADM4", "T_HF", _
                  "T_NAMES", "T_HISTOGEO", "T_HISTOHF", "T_METADATA"

    For counter = required.LowerBound To required.UpperBound
        tableName = CStr(required.Item(counter))
        If Not ListObjectIsOnSheet(geoSh, tableName) Then _
            missing = AppendName(missing, tableName)
    Next counter

    Assert.AreEqual vbNullString, missing, _
                    "Every table the geo manager demands is on the Geo sheet"
    Assert.LogSuccesses "The Geo sheet holds these tables: " & ListObjectNamesOf(geoSh)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheGeoSheetCarriesEveryTableTheGeoManagerNeeds", _
                         Err.Number, Err.Description
End Sub

'@sub-title The Geo sheet carries every hidden name the geo manager demands.
'@details
'Four hidden names on the sheet and one named range. The names the sheet does
'carry are logged beside the verdict, since a name under an older spelling is
'the likeliest reason one is missing.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheGeoSheetCarriesEveryHiddenNameTheGeoManagerNeeds()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheGeoSheetCarriesEveryHiddenNameTheGeoManagerNeeds"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim geoSh As Worksheet
    Dim geoStore As HiddenNames
    Dim required As BetterArray
    Dim missing As String
    Dim nameId As String
    Dim counter As Long

    If SkipWhenNoExporter("TestTheGeoSheetCarriesEveryHiddenNameTheGeoManagerNeeds", exporter) Then Exit Sub

    Set geoSh = GeoSheetOf(exporter)
    If geoSh Is Nothing Then
        Assert.Fail "The file carries no Geo worksheet"
        Exit Sub
    End If

    Set geoStore = HiddenNames.Create(geoSh)

    Set required = New BetterArray
    required.Push "RNG_GeoName", "RNG_GeoUpdated", "RNG_GeoLangCode", "RNG_MetaLang"

    For counter = required.LowerBound To required.UpperBound
        nameId = CStr(required.Item(counter))
        If Not geoStore.HasName(nameId) Then missing = AppendName(missing, nameId)
    Next counter

    'RNG_PastingGeoCol is checked as a range rather than a stored value, which
    'is what RangeExists does inside CheckRequirements.
    If Not NameIsOnSheet(geoSh, "RNG_PastingGeoCol") Then _
        missing = AppendName(missing, "RNG_PastingGeoCol (as a range)")

    'This is reported and never failed. A linelist older than the hidden name
    'store kept its geo metadata in plain named ranges, so an old file having
    'none of these is what an old file looks like rather than a fault in it.
    If LenB(missing) = 0 Then
        Assert.LogSuccesses "Every hidden name the geo manager demands is on " & _
                            "the Geo sheet"
    Else
        Assert.LogSuccesses "The Geo sheet is missing these hidden names, " & _
                            "which is expected of a linelist older than the " & _
                            "store: " & missing
    End If

    Assert.LogSuccesses "The Geo sheet holds these hidden names: " & _
                        StoredNamesOf(geoStore)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheGeoSheetCarriesEveryHiddenNameTheGeoManagerNeeds", _
                         Err.Number, Err.Description
End Sub

'@sub-title The geo manager builds on the Geo sheet of the file.
'@details
'The call ExportGeo names when it refuses, made here with no exporter in the
'way. A raise proves the geobase failure is the geo manager refusing the sheet
'rather than anything the export does with what it hands back.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheGeoManagerBuildsOnTheFilesGeoSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheGeoManagerBuildsOnTheFilesGeoSheet"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim geoSh As Worksheet
    Dim geoManager As LLGeo
    Dim raisedNumber As Long
    Dim raisedMessage As String
    Dim strictNumber As Long
    Dim strictMessage As String

    If SkipWhenNoExporter("TestTheGeoManagerBuildsOnTheFilesGeoSheet", exporter) Then Exit Sub

    Set geoSh = GeoSheetOf(exporter)
    If geoSh Is Nothing Then
        Assert.Fail "The file carries no Geo worksheet"
        Exit Sub
    End If

    'What the export path does: the checks are off, so the tables are what the
    'geobase is judged on.
    On Error Resume Next
        Set geoManager = LLGeo.Create(geoSh, runChecks:=False)
        raisedNumber = Err.Number
        raisedMessage = Err.Description
    On Error GoTo TestFail

    Assert.AreEqual 0&, raisedNumber, _
                    "The geo manager builds on this Geo sheet with the checks " & _
                    "off - text was [" & raisedMessage & "]"

    'And what the strict form does, reported rather than failed. An old file
    'refusing here is the reason the export path turns the checks off.
    Set geoManager = Nothing
    On Error Resume Next
        Set geoManager = LLGeo.Create(geoSh)
        strictNumber = Err.Number
        strictMessage = Err.Description
    On Error GoTo TestFail

    If strictNumber = 0 Then
        Assert.LogSuccesses "The strict build passes on this file too"
    Else
        Assert.LogSuccesses "The strict build refuses this file with error " & _
                            CStr(strictNumber) & " [" & strictMessage & "], " & _
                            "which is expected of a linelist older than the store"
    End If

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheGeoManagerBuildsOnTheFilesGeoSheet", _
                         Err.Number, Err.Description
End Sub


'@section Step 8 -- bisecting the migration export by its options
'===============================================================================
'@description
'ExportMigration answered 1007 with both boxes ticked. It is called again here
'with both untick, which is the one bisection its public surface allows: a pass
'puts the fault in AddShowHide, the only thing the first call did that this one
'does not, and a failure puts it in the part both runs share.

'@sub-title ExportMigration writes its file with show/hide and labels left out.
'@details
'The exporter is the one the module holds, so the output workbook the failed
'run left open is still open. CloseAll would take the source down with it, so
'the leftover is dropped in ModuleCleanup instead of here.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheMigrationExportWritesItsFileWithoutShowHide()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheMigrationExportWritesItsFileWithoutShowHide"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim savedPath As String
    Dim raisedNumber As Long
    Dim raisedMessage As String

    If SkipWhenNoExporter("TestTheMigrationExportWritesItsFileWithoutShowHide", exporter) Then Exit Sub

    On Error Resume Next
        savedPath = exporter.ExportMigration(SecondOutputFolder(), _
                                             includeShowHide:=False, _
                                             keepLabels:=False)
        raisedNumber = Err.Number
        raisedMessage = Err.Description
    On Error GoTo TestFail

    Assert.AreEqual 0&, raisedNumber, _
                    "ExportMigration with show/hide off raises nothing - text was [" & _
                    raisedMessage & "], LastFailure was [" & exporter.LastFailure & "]"

    AssertFileWasWritten savedPath, "the migration export with show/hide off"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheMigrationExportWritesItsFileWithoutShowHide", _
                         Err.Number, Err.Description
End Sub


'@section Private -- reading the Geo sheet
'===============================================================================

'@sub-title The Geo worksheet of the file under test.
'@param exporter LLExporter. The exporter bound to the file.
'@return Worksheet. The Geo sheet, or Nothing when the file has none.
Private Function GeoSheetOf(ByVal exporter As LLExporter) As Worksheet

    Dim geoSh As Worksheet

    'A missing sheet raises on the read, and Nothing is the answer read back.
    On Error Resume Next
    Set geoSh = exporter.SourceWorkbook.Worksheets(GEO_SHEET)
    On Error GoTo 0

    Set GeoSheetOf = geoSh
End Function

'@sub-title Whether a ListObject of that name is on a sheet.
'@param sh Worksheet. The sheet to read.
'@param tableName String. The name to look for.
'@return Boolean. True when the table is there.
Private Function ListObjectIsOnSheet(ByVal sh As Worksheet, _
                                     ByVal tableName As String) As Boolean

    Dim lo As ListObject

    On Error Resume Next
    Set lo = sh.ListObjects(tableName)
    On Error GoTo 0

    ListObjectIsOnSheet = Not (lo Is Nothing)
End Function

'@sub-title The ListObject names of a sheet on one line.
'@param sh Worksheet. The sheet to read.
'@return String. The names, comma separated.
Private Function ListObjectNamesOf(ByVal sh As Worksheet) As String

    Dim lo As ListObject
    Dim names As String

    For Each lo In sh.ListObjects
        names = AppendName(names, lo.Name)
    Next lo

    If LenB(names) = 0 Then names = "(none)"
    ListObjectNamesOf = names
End Function

'@sub-title Whether a sheet-scoped name resolves to a range.
'@details
'This is what RangeExists does inside LLGeo.CheckRequirements. A HiddenNames
'entry holding a quoted literal is a name with no range behind it, so asking
'for the range is the check that matters here.
'@param sh Worksheet. The sheet to read.
'@param nameId String. The name to resolve.
'@return Boolean. True when the name answers a range.
Private Function NameIsOnSheet(ByVal sh As Worksheet, ByVal nameId As String) As Boolean

    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(nameId)
    On Error GoTo 0

    NameIsOnSheet = Not (rng Is Nothing)
End Function

'@sub-title The hidden names a store holds, on one line.
'@param store HiddenNames. The store to read.
'@return String. The names, comma separated.
Private Function StoredNamesOf(ByVal store As HiddenNames) As String

    Dim held As BetterArray
    Dim names As String
    Dim counter As Long
    Dim record As Variant

    On Error Resume Next
    Set held = store.ListNames()
    On Error GoTo 0

    If held Is Nothing Then
        StoredNamesOf = "(unreadable)"
        Exit Function
    End If

    'ListNames answers one RECORD per name -- Array(nameId, type, updated) --
    'so the item is an array and CStr over it raises 13. The name is its first
    'element, and a record that is not an array is taken as the name itself.
    For counter = held.LowerBound To held.UpperBound
        record = held.Item(counter)
        If IsArray(record) Then
            names = AppendName(names, CStr(record(LBound(record))))
        Else
            names = AppendName(names, CStr(record))
        End If
    Next counter

    If LenB(names) = 0 Then names = "(none)"
    StoredNamesOf = names
End Function


'@section Step 9 -- why the store sees no name
'===============================================================================
'@description
'The Geo sheet answered "(none)" for its hidden names while RNG_PastingGeoCol
'resolved as a range on the same sheet, so names are on that sheet and the
'store is not seeing them. HiddenNames.ShouldTrack drops any name whose Visible
'property is True, before it looks at anything else. This reads the raw Excel
'names and their Visible flag, so the reason is measured rather than inferred.

'@sub-title The four names the geo manager wants, read raw with their visibility.
'@details
'Each name is looked for at sheet scope and then at workbook scope, and what is
'found is reported with its Visible flag. A name that is present and visible is
'a name the store refuses by design; a name absent at both scopes is a file
'that never carried it.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheGeoNamesAreOnTheFileButVisible()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheGeoNamesAreOnTheFileButVisible"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim geoSh As Worksheet
    Dim wanted As BetterArray
    Dim report As String
    Dim nameId As String
    Dim counter As Long
    Dim presentAndVisible As Long

    If SkipWhenNoExporter("TestTheGeoNamesAreOnTheFileButVisible", exporter) Then Exit Sub

    Set geoSh = GeoSheetOf(exporter)
    If geoSh Is Nothing Then
        Assert.Fail "The file carries no Geo worksheet"
        Exit Sub
    End If

    Set wanted = New BetterArray
    wanted.Push "RNG_GeoName", "RNG_GeoUpdated", "RNG_GeoLangCode", _
                "RNG_MetaLang", "RNG_PastingGeoCol"

    For counter = wanted.LowerBound To wanted.UpperBound
        nameId = CStr(wanted.Item(counter))
        report = AppendName(report, nameId & " -> " & RawNameState(geoSh, nameId))
        If RawNameIsVisible(geoSh, nameId) Then _
            presentAndVisible = presentAndVisible + 1
    Next counter

    Assert.LogSuccesses "Raw name state on the Geo sheet: " & report
    Assert.LogSuccesses "Sheet-scoped names on Geo: " & _
                        CStr(SheetScopedNameCount(geoSh)) & " total, " & _
                        CStr(SheetScopedHiddenCount(geoSh)) & " of them hidden"

    'Reported and never failed. The point is to record what the file carries,
    'and an old file may hold these as visible names or not at all.
    Assert.LogSuccesses "Names present and visible, which is what ShouldTrack " & _
                        "refuses: " & CStr(presentAndVisible)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheGeoNamesAreOnTheFileButVisible", _
                         Err.Number, Err.Description
End Sub


'@section Private -- raw Excel names
'===============================================================================

'@sub-title The Name object of that id, at sheet scope then workbook scope.
'@param sh Worksheet. The sheet to look on.
'@param nameId String. The name to look for.
'@return Name. The definition, or Nothing when neither scope holds it.
Private Function RawNameOf(ByVal sh As Worksheet, ByVal nameId As String) As Name

    Dim definition As Name

    On Error Resume Next
    Set definition = sh.Names(nameId)
    On Error GoTo 0

    If definition Is Nothing Then
        On Error Resume Next
        Set definition = sh.Parent.Names(nameId)
        On Error GoTo 0
    End If

    Set RawNameOf = definition
End Function

'@sub-title Whether that name is on the file and visible.
'@param sh Worksheet. The sheet to look on.
'@param nameId String. The name to look for.
'@return Boolean. True when it is there and Visible is True.
Private Function RawNameIsVisible(ByVal sh As Worksheet, ByVal nameId As String) As Boolean

    Dim definition As Name

    Set definition = RawNameOf(sh, nameId)
    If definition Is Nothing Then Exit Function

    On Error Resume Next
    RawNameIsVisible = definition.Visible
    On Error GoTo 0
End Function

'@sub-title What a name is on the file, on one short phrase.
'@param sh Worksheet. The sheet to look on.
'@param nameId String. The name to look for.
'@return String. absent, visible or hidden.
Private Function RawNameState(ByVal sh As Worksheet, ByVal nameId As String) As String

    Dim definition As Name
    Dim isVisible As Boolean

    Set definition = RawNameOf(sh, nameId)
    If definition Is Nothing Then
        RawNameState = "absent"
        Exit Function
    End If

    On Error Resume Next
    isVisible = definition.Visible
    On Error GoTo 0

    If isVisible Then
        RawNameState = "present, visible"
    Else
        RawNameState = "present, hidden"
    End If
End Function

'@sub-title How many names are scoped to that sheet.
'@param sh Worksheet. The sheet to read.
'@return Long. The count.
Private Function SheetScopedNameCount(ByVal sh As Worksheet) As Long
    On Error Resume Next
    SheetScopedNameCount = sh.Names.count
    On Error GoTo 0
End Function

'@sub-title How many of that sheet's names are hidden.
'@param sh Worksheet. The sheet to read.
'@return Long. The count.
Private Function SheetScopedHiddenCount(ByVal sh As Worksheet) As Long

    Dim definition As Name
    Dim hiddenCount As Long

    On Error Resume Next
    For Each definition In sh.Names
        If Not definition.Visible Then hiddenCount = hiddenCount + 1
    Next definition
    On Error GoTo 0

    SheetScopedHiddenCount = hiddenCount
End Function


'@section Step 10 -- what the password sheet is missing
'===============================================================================
'@description
'The migration export now names its step: it fails reading the passwords, out of
'Passwords.Create. That factory calls ValidateSheet, which demands two tables and
'five named ranges on __pass. Neither check goes through HiddenNames -- both are
'plain sheet.ListObjects and sheet.Range lookups -- so this one is NOT the
'visible-name problem the Geo sheet has. Something is genuinely absent.

'@sub-title The password sheet carries everything Passwords.Create demands.
'@details
'ValidateSheet stops on the first thing it cannot find, so this reads all seven
'and names every one that is absent. Reported and never failed: an old linelist
'lacking a table this version added is a fact about the file.
'@TestMethod("ExportOtherLinelist")
Public Sub TestThePasswordSheetCarriesWhatItsFactoryDemands()
    CustomTestSetTitles Assert, TESTMODULE, "TestThePasswordSheetCarriesWhatItsFactoryDemands"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim passSh As Worksheet
    Dim wantedTables As BetterArray
    Dim wantedNames As BetterArray
    Dim missing As String
    Dim oneName As String
    Dim counter As Long

    If SkipWhenNoExporter("TestThePasswordSheetCarriesWhatItsFactoryDemands", exporter) Then Exit Sub

    Set passSh = Nothing
    On Error Resume Next
    Set passSh = exporter.SourceWorkbook.Worksheets(PASS_SHEET)
    On Error GoTo TestFail

    If passSh Is Nothing Then
        Assert.Fail "The file carries no " & PASS_SHEET & " worksheet"
        Exit Sub
    End If

    Set wantedTables = New BetterArray
    wantedTables.Push "T_keys", "T_ProtectedSheets"

    For counter = wantedTables.LowerBound To wantedTables.UpperBound
        oneName = CStr(wantedTables.Item(counter))
        If Not ListObjectIsOnSheet(passSh, oneName) Then _
            missing = AppendName(missing, oneName & " (table)")
    Next counter

    Set wantedNames = New BetterArray
    wantedNames.Push "RNG_DebuggingPassword", "RNG_PublicKey", "RNG_PrivateKey", _
                     "RNG_DebugMode", "RNG_Version"

    For counter = wantedNames.LowerBound To wantedNames.UpperBound
        oneName = CStr(wantedNames.Item(counter))
        If Not NameIsOnSheet(passSh, oneName) Then _
            missing = AppendName(missing, oneName & " (range)")
    Next counter

    If LenB(missing) = 0 Then
        Assert.LogSuccesses "The password sheet carries everything its factory demands"
    Else
        Assert.LogSuccesses "The password sheet is missing: " & missing
    End If

    Assert.LogSuccesses "The password sheet holds these tables: " & _
                        ListObjectNamesOf(passSh)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestThePasswordSheetCarriesWhatItsFactoryDemands", _
                         Err.Number, Err.Description
End Sub


'@section Step 11 -- what the translation sheet is missing
'===============================================================================
'@description
'With the passwords behind it, the migration export now names its next step:
'adding the metadata sheets, out of LLTranslation.Create. That factory demands
'five tables on the LinelistTranslation sheet, by plain ListObjects lookup, and
'stops on the first one it cannot find.

'@sub-title The translation sheet carries every table its factory demands.
'@details
'Reported and never failed: an old linelist missing a table a later version
'added is a fact about the file, the same as the password sheet above.
'@TestMethod("ExportOtherLinelist")
Public Sub TestTheTranslationSheetCarriesWhatItsFactoryDemands()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheTranslationSheetCarriesWhatItsFactoryDemands"
    On Error GoTo TestFail

    Dim exporter As LLExporter
    Dim tradSh As Worksheet
    Dim wanted As BetterArray
    Dim missing As String
    Dim oneName As String
    Dim counter As Long

    If SkipWhenNoExporter("TestTheTranslationSheetCarriesWhatItsFactoryDemands", exporter) Then Exit Sub

    Set tradSh = Nothing
    On Error Resume Next
    Set tradSh = exporter.SourceWorkbook.Worksheets(LINELIST_TRANSLATION_SHEET)
    On Error GoTo TestFail

    If tradSh Is Nothing Then
        Assert.Fail "The file carries no " & LINELIST_TRANSLATION_SHEET & " worksheet"
        Exit Sub
    End If

    Set wanted = New BetterArray
    wanted.Push "T_TradLLMsg", "T_TradLLShapes", "T_TradLLForms", _
                "T_TradLLRibbon", "Tab_Translations"

    For counter = wanted.LowerBound To wanted.UpperBound
        oneName = CStr(wanted.Item(counter))
        If Not ListObjectIsOnSheet(tradSh, oneName) Then _
            missing = AppendName(missing, oneName)
    Next counter

    If LenB(missing) = 0 Then
        Assert.LogSuccesses "The translation sheet carries every table its factory demands"
    Else
        Assert.LogSuccesses "The translation sheet is missing: " & missing
    End If

    Assert.LogSuccesses "The translation sheet holds these tables: " & _
                        ListObjectNamesOf(tradSh)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheTranslationSheetCarriesWhatItsFactoryDemands", _
                         Err.Number, Err.Description
End Sub
