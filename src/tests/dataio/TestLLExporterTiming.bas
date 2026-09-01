Attribute VB_Name = "TestLLExporterTiming"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Times the LLExporter export walks against a sized fixture")
'
'@description
'   A MEASURING PROBE, not a behaviour suite. It stays commented out of the
'   registry the way TestCustomTableTiming does: every test here reports a
'   number rather than a pass, and the numbers move with whatever else the
'   machine is doing. Uncomment the row, run, read the TIMING lines out of the
'   run CSV, comment it back out.
'
'   WHAT IT IS FOR
'   ---------------------------------------------------------------------------
'   The export and import speed block (`.obt/implementations.md` Sessions 117 to
'   125) owes a baseline every later session measures against. Session 117 put
'   the stopwatch on the walk; this module is what drives the walk so the
'   stopwatch has something to time.
'
'   To compare two versions of LLExporter, run this module, check the other
'   version of `src/classes/dataio/LLExporter.cls` out, run it again, and put
'   the file back. Nothing here touches a private member, so the same module
'   compiles against both.
'
'   WHAT THE FIXTURE SIZE MEANS, AND WHICH COSTS IT CAN SEE
'   ---------------------------------------------------------------------------
'   Not every cost scales with the same thing, so a fixture that measures one
'   well measures another badly. Read a number against the list before believing
'   it says something about a real linelist:
'
'     Whole-grid formatting   INDEPENDENT of the data size. Setting a row height
'                             over 1,048,576 rows costs the same on a three-row
'                             sheet as on a three-hundred-thousand-row one, and
'                             it is paid once per data sheet. SHEET_COUNT is
'                             what moves it, not DATA_ROWS.
'     The save                grows with what the grid carries. A sheet whose
'                             whole grid holds a format is written out with a
'                             record per row and per column, so this one shows
'                             up in the save step whatever the data size.
'     HiddenNames.Create      grows with the NUMBER OF NAMES on the source
'                             sheet, not with the rows. HIDDEN_NAMES_PER_SHEET
'                             is what moves it. A real data entry sheet carries
'                             hundreds.
'     The data move itself    grows with rows times columns. DATA_ROWS and
'                             VARS_PER_SHEET move it.
'
'   The fixture is a synthetic linelist, not a generated one. It has no
'   analyses, no geobase and no formulas, so it says nothing about those stages.
'   A real linelist baseline is still owner hand work.
'
'   SECONDS ARE WRITTEN WITH Str$, NEVER Format$
'   ---------------------------------------------------------------------------
'   This box is en_FR and `Format$(2.15, "0.000")` writes `2,150` here and
'   `2.150` on an English Windows machine. Two timing runs are read side by
'   side, which is the whole point, so the seconds go through `Str$`, which is
'   locale-independent.
'@depends LLExporter, LLLog, HiddenNames, CustomTest, TestHelpersLite, DictionaryTestFixture, ChoicesTestFixture

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "LLExporterTiming"

'THE FIXTURE SIZE. Each of these moves a different cost -- see the list above.
'Six data sheets rather than the two TestLLExporter builds, because the
'whole-grid formatting is paid once per sheet and per-sheet is the whole shape
'of that cost.
Private Const SHEET_COUNT As Long = 6
Private Const VARS_PER_SHEET As Long = 30
Private Const DATA_ROWS As Long = 300

'What makes HiddenNames.Create expensive. Create walks every tracked name of the
'host and reads three properties off each; QuickValue reads the one name asked
'for. A generated data entry sheet carries one name per variable plus the sheet
'metadata, so a few hundred is the real shape.
Private Const HIDDEN_NAMES_PER_SHEET As Long = 250

'The linelist sheets LLExporter reads before it writes anything.
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const EXPORTS_SHEET As String = "Exports"
Private Const CHOICES_SHEET As String = "Choices"
Private Const METADATA_SHEET As String = "Metadata"
Private Const PASS_SHEET As String = "__pass"
Private Const GEO_SHEET As String = "Geo"
Private Const TEMP_SHEET As String = "__temp"
Private Const LOG_SHEET As String = "__log"

'AddShowHide reads the source store with ShowHideStore.CreateForRead and merges
'whatever it holds, with no guard for a source that has no store sheet. An empty
'one carries nothing and merges nothing, which is what a linelist that has never
'had the form opened looks like.
Private Const SHOWHIDE_SHEET As String = "__show_hide"

Private Const SHEET_PREFIX As String = "timing-sheet"
Private Const FILES_FOLDER As String = "ExporterTimingFiles"

'The text LLLog writes its step list under. The column is NOT hard-coded: the
'first version of this module looked in column 3, the entry detail lands in
'column 5, and the step list came back empty with nothing raised and nothing
'reported. A missing measurement that looks like a clean run is the worst shape
'a probe can have, so the whole used range is scanned instead.
Private Const STEP_MARKER As String = "step times"

Private Assert As CustomTest
Private SetupError As Long
Private SetupMessage As String


'@section Module lifecycle
'===============================================================================

'@sub-title Build the assertion harness.
'@details Public because the harness calls it by name through Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLLExporterTiming"
    SetupError = 0
    SetupMessage = vbNullString
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'PrintResults raises 1004 whenever another workbook holds the screen, so the
    'driver is handed the screen back first.
    On Error Resume Next
    ThisWorkbook.Activate
    On Error GoTo 0

    If Not Assert Is Nothing Then Assert.PrintResults TEST_OUTPUT_SHEET
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanUp
Public Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.FlushCurrentTest
End Sub


'@section Tests
'===============================================================================

'@sub-title Times one migration export end to end and reports its step list.
'@details
'The migration is the widest walk: it carries the whole dictionary, every data
'sheet, the metadata set and the show/hide choices. The wall time of the call is
'the headline number; the step list LLLog wrote is what attributes it, and both
'go into the results so a later run can be read against this one.
'@TestMethod("LLExporterTiming")
Public Sub TestTimeMigrationExport()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeMigrationExport"
    On Error GoTo TestFail

    Dim sourceBook As Workbook
    Dim exporter As LLExporter
    Dim savedPath As String
    Dim startedAt As Double
    Dim elapsed As Double
    Dim failedNumber As Long
    Dim failedText As String

    Set sourceBook = TimingSourceWorkbook()
    Set exporter = LLExporter.Create(sourceBook)

    startedAt = Timer
    savedPath = exporter.ExportMigration(BuildTempFolder(ThisWorkbook, FILES_FOLDER), _
                                         includeShowHide:=True, _
                                         keepLabels:=False)
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (LenB(savedPath) > 0), _
                  "The migration wrote a file, so the number below timed a real walk"

    LogTiming "migration export, whole walk", elapsed
    LogFixtureShape
    LogSavedFileSize savedPath
    LogStepListFrom sourceBook

    DropArtefacts sourceBook, exporter, savedPath
    Exit Sub
TestFail:
    failedNumber = Err.Number
    failedText = Err.Description
    If Not exporter Is Nothing Then _
        failedText = failedText & " | " & exporter.LastFailure
    On Error Resume Next
    DropArtefacts sourceBook, exporter, savedPath
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTimeMigrationExport", failedNumber, failedText
End Sub

'@sub-title Times a second migration on the same fixture.
'@details
'The first export of a session pays whatever Excel warms up on the way -- the
'save path, the styles of a new workbook. A second run on an identical fixture
'says how much of the first number was that. Read the two together; the smaller
'one is the steadier estimate.
'@TestMethod("LLExporterTiming")
Public Sub TestTimeMigrationExportAgain()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeMigrationExportAgain"
    On Error GoTo TestFail

    Dim sourceBook As Workbook
    Dim exporter As LLExporter
    Dim savedPath As String
    Dim startedAt As Double
    Dim elapsed As Double
    Dim failedNumber As Long
    Dim failedText As String

    Set sourceBook = TimingSourceWorkbook()
    Set exporter = LLExporter.Create(sourceBook)

    startedAt = Timer
    savedPath = exporter.ExportMigration(BuildTempFolder(ThisWorkbook, FILES_FOLDER), _
                                         includeShowHide:=True, _
                                         keepLabels:=False)
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (LenB(savedPath) > 0), _
                  "The second migration wrote a file too"

    LogTiming "migration export, second run", elapsed
    LogStepListFrom sourceBook

    DropArtefacts sourceBook, exporter, savedPath
    Exit Sub
TestFail:
    failedNumber = Err.Number
    failedText = Err.Description
    If Not exporter Is Nothing Then _
        failedText = failedText & " | " & exporter.LastFailure
    On Error Resume Next
    DropArtefacts sourceBook, exporter, savedPath
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTimeMigrationExportAgain", failedNumber, failedText
End Sub

'@sub-title Times HiddenNames.Create against QuickValue on a fixture sheet.
'@details
'The fourth cut of Session 118, measured on its own rather than inside the walk,
'because inside the walk it is one cost among several. Both routes answer the
'same string off the same sheet. The sheet carries HIDDEN_NAMES_PER_SHEET names,
'which is the shape a generated data entry sheet has.
'@TestMethod("LLExporterTiming")
Public Sub TestTimeHiddenNameRoutes()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeHiddenNameRoutes"
    On Error GoTo TestFail

    Dim sourceBook As Workbook
    Dim sh As Worksheet
    Dim store As HiddenNames
    Dim viaCreate As String
    Dim viaQuick As String
    Dim startedAt As Double
    Dim createSeconds As Double
    Dim quickSeconds As Double
    Dim counter As Long
    Dim failedNumber As Long
    Dim failedText As String

    Const REPEATS As Long = 20

    Set sourceBook = TimingSourceWorkbook()
    Set sh = sourceBook.Worksheets(SHEET_PREFIX & "1")

    startedAt = Timer
    For counter = 1 To REPEATS
        Set store = HiddenNames.Create(sh)
        viaCreate = store.ValueAsString("sheet_type")
    Next counter
    createSeconds = ElapsedSince(startedAt)

    startedAt = Timer
    For counter = 1 To REPEATS
        viaQuick = HiddenNames.QuickValue(sh, "sheet_type")
    Next counter
    quickSeconds = ElapsedSince(startedAt)

    Assert.AreEqual viaCreate, viaQuick, _
                    "Both routes answer the same tag, so the two times compare"

    LogTiming "HiddenNames.Create x" & REPEATS & " over " & _
              HIDDEN_NAMES_PER_SHEET & " names", createSeconds
    LogTiming "HiddenNames.QuickValue x" & REPEATS & " over " & _
              HIDDEN_NAMES_PER_SHEET & " names", quickSeconds

    On Error Resume Next
    DeleteWorkbook sourceBook
    On Error GoTo 0
    Exit Sub
TestFail:
    failedNumber = Err.Number
    failedText = Err.Description
    On Error Resume Next
    DeleteWorkbook sourceBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTimeHiddenNameRoutes", failedNumber, failedText
End Sub


'@section Where the dictionary step goes
'===============================================================================
'The migration walk spends about 60% of itself in one step, `adding the
'dictionary`. That step is LLExporter.AddDictionary, four statements, of which
'the first is LLdictionary.Export. Export calls EnsureFormatColumns, which
'registers SEVEN format columns on the DataSheet under it, and DataSheet.Export
'then runs ApplyFormat once per registered column over the whole data column.
'ApplyFormat is already on record as the measured 65% cost centre of the
'analyses build, so it is the suspect. These three tests decide it rather than
'assuming it.

'@sub-title Times the whole call AddDictionary makes.
'@details The upper bound of the step, and the number the other two are read
'   against. Same arguments LLExporter.AddDictionary passes.
'@TestMethod("LLExporterTiming")
Public Sub TestTimeDictionaryExportWholeCall()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeDictionaryExportWholeCall"
    On Error GoTo TestFail

    Dim sourceBook As Workbook
    Dim targetBook As Workbook
    Dim dict As LLdictionary
    Dim startedAt As Double
    Dim elapsed As Double
    Dim failedNumber As Long
    Dim failedText As String

    Set sourceBook = TimingDictionaryWorkbook()
    Set targetBook = NewWorkbook()
    Set dict = LLdictionary.Create(sourceBook.Worksheets(DICTIONARY_SHEET), 1, 1, 5)

    startedAt = Timer
    dict.Export toWkb:=targetBook, exportType:="__all__", _
                addListObject:=False, Hide:=xlSheetVisible
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (elapsed >= 0), "The dictionary export ran"

    LogTiming "LLdictionary.Export, the whole call AddDictionary makes", elapsed
    LogDictionaryShape sourceBook

    DropBooks sourceBook, targetBook
    Exit Sub
TestFail:
    failedNumber = Err.Number
    failedText = Err.Description
    On Error Resume Next
    DropBooks sourceBook, targetBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTimeDictionaryExportWholeCall", failedNumber, failedText
End Sub

'@sub-title Times the same export with NO format columns registered.
'@details
'DataSheet.Export skips its whole format loop when the list is empty, so this is
'the data move and the cosmetic pass and nothing else. It goes through
'DataSheet.Export directly rather than LLdictionary.Export, because
'EnsureFormatColumns would put the seven columns straight back.
'@TestMethod("LLExporterTiming")
Public Sub TestTimeDictionaryExportWithoutFormatColumns()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeDictionaryExportWithoutFormatColumns"
    On Error GoTo TestFail

    Dim sourceBook As Workbook
    Dim targetBook As Workbook
    Dim dict As LLdictionary
    Dim startedAt As Double
    Dim elapsed As Double
    Dim failedNumber As Long
    Dim failedText As String

    Set sourceBook = TimingDictionaryWorkbook()
    Set targetBook = NewWorkbook()
    Set dict = LLdictionary.Create(sourceBook.Worksheets(DICTIONARY_SHEET), 1, 1, 5)

    'Clear the list. resetColumns True with no column names leaves it empty.
    dict.Data.AddFormatsColumns True, True

    startedAt = Timer
    dict.Data.Export toWkb:=targetBook, Hide:=xlSheetVisible
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (elapsed >= 0), "The export with no format columns ran"

    LogTiming "DataSheet.Export, NO format columns", elapsed

    DropBooks sourceBook, targetBook
    Exit Sub
TestFail:
    failedNumber = Err.Number
    failedText = Err.Description
    On Error Resume Next
    DropBooks sourceBook, targetBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTimeDictionaryExportWithoutFormatColumns", failedNumber, failedText
End Sub

'@sub-title Times the same export with the seven format columns registered.
'@details
'The same DataSheet.Export as the test above, same fixture, same target, with
'the seven columns EnsureFormatColumns registers. The difference between the two
'is the ApplyFormat loop and nothing else.
'@TestMethod("LLExporterTiming")
Public Sub TestTimeDictionaryExportWithFormatColumns()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeDictionaryExportWithFormatColumns"
    On Error GoTo TestFail

    Dim sourceBook As Workbook
    Dim targetBook As Workbook
    Dim dict As LLdictionary
    Dim startedAt As Double
    Dim elapsed As Double
    Dim failedNumber As Long
    Dim failedText As String

    Set sourceBook = TimingDictionaryWorkbook()
    Set targetBook = NewWorkbook()
    Set dict = LLdictionary.Create(sourceBook.Worksheets(DICTIONARY_SHEET), 1, 1, 5)

    'The seven names EnsureFormatColumns registers, written out here so the
    'test says what it is measuring rather than reaching a private routine.
    dict.Data.AddFormatsColumns False, True, _
                                "formatting condition", "formatting values", _
                                "variable name", "control", "main label", _
                                "lock cells", "dev comments"

    startedAt = Timer
    dict.Data.Export toWkb:=targetBook, Hide:=xlSheetVisible
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (elapsed >= 0), "The export with seven format columns ran"

    LogTiming "DataSheet.Export, SEVEN format columns", elapsed

    DropBooks sourceBook, targetBook
    Exit Sub
TestFail:
    failedNumber = Err.Number
    failedText = Err.Description
    On Error Resume Next
    DropBooks sourceBook, targetBook
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestTimeDictionaryExportWithFormatColumns", failedNumber, failedText
End Sub


'@section Reporting
'===============================================================================

'@sub-title Writes one timing line into the results.
'@param label String. What was measured.
'@param elapsed Double. Seconds it took.
Private Sub LogTiming(ByVal label As String, ByVal elapsed As Double)
    Assert.LogSuccesses "TIMING " & label & ": " & SecondsText(elapsed) & " s"
End Sub

'@sub-title Writes the fixture size into the results beside the numbers.
'@details A timing line means nothing without the shape it was measured on, and
'   the two have to travel together or a later run is compared against a
'   different fixture without anybody noticing.
Private Sub LogFixtureShape()
    Assert.LogSuccesses "TIMING fixture: " & SHEET_COUNT & " data sheets, " & _
                        VARS_PER_SHEET & " variables each, " & DATA_ROWS & _
                        " rows each, " & HIDDEN_NAMES_PER_SHEET & _
                        " hidden names per sheet, platform " & PlatformTag()
End Sub

'@sub-title Writes the size of the file the export wrote.
'@details The whole-grid formatting costs twice, and the second time is here: a
'   sheet whose whole grid carries a row height, a width and a font is saved
'   with a format record for every row and every column. The byte count is the
'   only place that cost is visible as a number.
'@param savedPath String. The file the export wrote.
Private Sub LogSavedFileSize(ByVal savedPath As String)
    Dim sizeBytes As Double

    If LenB(savedPath) = 0 Then Exit Sub

    On Error Resume Next
    sizeBytes = FileLen(savedPath)
    On Error GoTo 0

    Assert.LogSuccesses "TIMING saved file: " & Trim$(Str$(sizeBytes)) & " bytes"
End Sub

'@sub-title Reads the step list LLLog wrote and puts it in the results.
'@details
'The walk writes one info line per export carrying every step with its seconds.
'It lands on the source workbook's __log sheet, so it is read back off that
'sheet here rather than through LLLog, which has no reader for one entry.
'@param sourceBook Workbook. The workbook the export ran against.
Private Sub LogStepListFrom(ByVal sourceBook As Workbook)
    Dim sh As Worksheet
    Dim usedRng As Range
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim detail As String
    Dim found As Boolean

    If sourceBook Is Nothing Then Exit Sub

    On Error Resume Next
    Set sh = sourceBook.Worksheets(LOG_SHEET)
    On Error GoTo 0
    If sh Is Nothing Then
        Assert.LogSuccesses "TIMING step list: the source carries no " & LOG_SHEET & " sheet"
        Exit Sub
    End If

    Set usedRng = sh.UsedRange
    If usedRng Is Nothing Then Exit Sub

    found = False
    For rowIndex = usedRng.Row To usedRng.Row + usedRng.Rows.Count - 1
        For colIndex = usedRng.Column To usedRng.Column + usedRng.Columns.Count - 1
            detail = CStr(sh.Cells(rowIndex, colIndex).Value)
            If InStr(1, detail, STEP_MARKER, vbTextCompare) > 0 Then
                Assert.LogSuccesses "TIMING steps: " & detail
                found = True
            End If
        Next colIndex
    Next rowIndex

    'Say so out loud. A probe that quietly reports one number fewer than it
    'promised reads as a clean run.
    If Not found Then _
        Assert.LogSuccesses "TIMING steps: NOT FOUND on the " & LOG_SHEET & _
                            " sheet - the walk logged no step list"
End Sub

'@fun-title Seconds as text, locale-independent.
'@details Format$ writes a comma on an en_FR box and a period on an English one,
'   and two runs are read side by side. Str$ always writes a period. It also
'   drops the leading zero of a value under one, which is put back.
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


'@section The fixture
'===============================================================================

'@fun-title A synthetic linelist wide enough for the walk to cost something.
'@details
'The sheets LLExporter reads before it writes anything, plus SHEET_COUNT HList
'data sheets. Every sheet carries the hidden names a generated one does, and
'HIDDEN_NAMES_PER_SHEET filler names on top, because HiddenNames.Create walks
'all of them and that walk is what Session 118 cut.
'@return Workbook. The new workbook, open and unsaved.
Private Function TimingSourceWorkbook() As Workbook
    Dim sourceBook As Workbook
    Dim counter As Long

    Set sourceBook = NewWorkbook()

    BuildDictionarySheet sourceBook
    ChoicesTestFixture.PrepareChoicesFixture CHOICES_SHEET, sourceBook
    BuildMetadataSheet sourceBook
    BuildExportsSheet sourceBook

    EnsureWorksheet PASS_SHEET, sourceBook
    EnsureWorksheet GEO_SHEET, sourceBook
    EnsureWorksheet TEMP_SHEET, sourceBook
    EnsureWorksheet SHOWHIDE_SHEET, sourceBook

    For counter = 1 To SHEET_COUNT
        BuildDataSheet sourceBook, SHEET_PREFIX & counter
    Next counter

    Set TimingSourceWorkbook = sourceBook
End Function

'@sub-title Writes a dictionary naming every fixture sheet and variable.
'@details
'The header row is the one the shared dictionary fixture uses, because the
'dictionary is read by column name. One row per variable per sheet.
'@param wkb Workbook. The workbook to write into.
Private Sub BuildDictionarySheet(ByVal wkb As Workbook)
    Dim sh As Worksheet
    Dim headers As Variant
    Dim block() As Variant
    Dim headerCount As Long
    Dim rowCount As Long
    Dim sheetIndex As Long
    Dim varIndex As Long
    Dim rowIndex As Long
    Dim colIndex As Long

    headers = DictionaryHeaders()
    headerCount = UBound(headers) - LBound(headers) + 1
    rowCount = SHEET_COUNT * VARS_PER_SHEET

    Set sh = EnsureWorksheet(DICTIONARY_SHEET, wkb, clearSheet:=True)
    sh.Range("A1").Resize(1, headerCount).Value = headers

    ReDim block(1 To rowCount, 1 To headerCount)

    rowIndex = 0
    For sheetIndex = 1 To SHEET_COUNT
        For varIndex = 1 To VARS_PER_SHEET
            rowIndex = rowIndex + 1
            For colIndex = 1 To headerCount
                block(rowIndex, colIndex) = vbNullString
            Next colIndex

            block(rowIndex, HeaderPosition(headers, "Variable Name")) = _
                VariableNameFor(sheetIndex, varIndex)
            block(rowIndex, HeaderPosition(headers, "Main Label")) = _
                "Label " & sheetIndex & "-" & varIndex
            block(rowIndex, HeaderPosition(headers, "Sheet Name")) = _
                SHEET_PREFIX & sheetIndex
            block(rowIndex, HeaderPosition(headers, "Sheet Type")) = "HList"
            block(rowIndex, HeaderPosition(headers, "Main Section")) = "section 1"
            block(rowIndex, HeaderPosition(headers, "Status")) = "mandatory"
            block(rowIndex, HeaderPosition(headers, "Variable Type")) = "text"
            block(rowIndex, HeaderPosition(headers, "Control")) = "free"
            block(rowIndex, HeaderPosition(headers, "Personal Identifier")) = "no"
            block(rowIndex, HeaderPosition(headers, "Export 1")) = "yes"
        Next varIndex
    Next sheetIndex

    sh.Range("A2").Resize(rowCount, headerCount).Value = block
End Sub

'@fun-title The dictionary column names, in order.
'@details Taken from the shared fixture so the two stay in step: LLdictionary
'   reads by column name and a header this module invented would be read as a
'   missing column rather than as an error.
'@return Variant. A zero-based array of header names.
Private Function DictionaryHeaders() As Variant
    DictionaryHeaders = DictionaryTestFixture.DictionaryFixtureHeaders()
End Function

'@fun-title Where one header sits in the header array, one-based.
'@param headers Variant. The header array.
'@param columnName String. The header wanted.
'@return Long. Its one-based position, or 1 when it is absent.
Private Function HeaderPosition(headers As Variant, ByVal columnName As String) As Long
    Dim idx As Long

    HeaderPosition = 1
    For idx = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(idx)), columnName, vbTextCompare) = 0 Then
            HeaderPosition = idx - LBound(headers) + 1
            Exit Function
        End If
    Next idx
End Function

'@fun-title The variable name of one column of one fixture sheet.
'@param sheetIndex Long. Which data sheet.
'@param varIndex Long. Which variable on it.
'@return String. The name the dictionary and the table header both carry.
Private Function VariableNameFor(ByVal sheetIndex As Long, ByVal varIndex As Long) As String
    VariableNameFor = "v" & sheetIndex & "_" & varIndex
End Function

'@sub-title Writes the Metadata sheet a migration appends its tags to.
'@details The header has to be variable/value with a row under it: the tags are
'   appended below DataRange("variable"), and that answers nothing at all for a
'   sheet holding a header and no body.
'@param wkb Workbook. The workbook to write into.
Private Sub BuildMetadataSheet(ByVal wkb As Workbook)
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(METADATA_SHEET, wkb, clearSheet:=True)
    WriteRow sh.Cells(1, 1), "variable", "value"
    WriteRow sh.Cells(2, 1), "language", "English"
End Sub

'@sub-title Writes the one Exports row a migration export reads.
'@param wkb Workbook. The workbook to write into.
Private Sub BuildExportsSheet(ByVal wkb As Workbook)
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(EXPORTS_SHEET, wkb, clearSheet:=True)

    WriteRow sh.Cells(1, 1), "export number", "status", "label button", _
                             "file format", "file name", "password", _
                             "include personal identifiers", "include p-codes", _
                             "header format", "export metadata sheets", _
                             "export analyses sheets", "admin levels"

    WriteRow sh.Cells(2, 1), 1, "active", "migration", "xlsx", "timing", "no", _
                             "yes", "yes", "default", "no", "no", vbNullString
End Sub

'@sub-title Builds one HList data sheet with its table and its hidden names.
'@details
'The table header row carries the variable names the dictionary names, because
'the export resolves every dictionary name against the table and writes the ones
'it finds. The body is written as one block rather than cell by cell, so
'building the fixture does not cost more than the walk being measured.
'@param wkb Workbook. The workbook to add the sheet to.
'@param sheetName String. The name the dictionary gives this sheet.
Private Sub BuildDataSheet(ByVal wkb As Workbook, ByVal sheetName As String)
    Dim sh As Worksheet
    Dim headerRow() As Variant
    Dim body() As Variant
    Dim listRange As Range
    Dim tableName As String
    Dim sheetIndex As Long
    Dim colIndex As Long
    Dim rowIndex As Long

    Set sh = EnsureWorksheet(sheetName, wkb, clearSheet:=True)
    sheetIndex = CLng(Mid$(sheetName, Len(SHEET_PREFIX) + 1))
    tableName = "Tab_" & Replace(sheetName, "-", "_")

    ReDim headerRow(1 To 1, 1 To VARS_PER_SHEET)
    For colIndex = 1 To VARS_PER_SHEET
        headerRow(1, colIndex) = VariableNameFor(sheetIndex, colIndex)
    Next colIndex
    sh.Range("A1").Resize(1, VARS_PER_SHEET).Value = headerRow

    ReDim body(1 To DATA_ROWS, 1 To VARS_PER_SHEET)
    For rowIndex = 1 To DATA_ROWS
        For colIndex = 1 To VARS_PER_SHEET
            body(rowIndex, colIndex) = "r" & rowIndex & "c" & colIndex
        Next colIndex
    Next rowIndex
    sh.Range("A2").Resize(DATA_ROWS, VARS_PER_SHEET).Value = body

    Set listRange = sh.Range("A1").Resize(DATA_ROWS + 1, VARS_PER_SHEET)
    sh.ListObjects.Add(SourceType:=xlSrcRange, _
                       Source:=listRange, _
                       XlListObjectHasHeaders:=xlYes).Name = tableName

    AddSheetNames sh, tableName
End Sub

'@sub-title Puts the metadata names and the filler names onto one data sheet.
'@details
'sheet_type and table_name are what the export reads. The filler names are what
'makes reading them expensive: HiddenNames.Create walks every tracked name and
'reads three properties off each, so a sheet with none of them would say the two
'routes cost the same. A generated data entry sheet carries one name per
'variable plus its own metadata.
'@param sh Worksheet. The data sheet.
'@param tableName String. The ListObject on it.
Private Sub AddSheetNames(ByVal sh As Worksheet, ByVal tableName As String)
    Dim store As HiddenNames
    Dim counter As Long

    Set store = HiddenNames.Create(sh)

    store.EnsureName "sheet_type", "HList", HiddenNameTypeString
    store.SetValue "sheet_type", "HList"
    store.EnsureName "table_name", tableName, HiddenNameTypeString
    store.SetValue "table_name", tableName

    For counter = 1 To HIDDEN_NAMES_PER_SHEET
        store.EnsureName "filler_" & counter, "value " & counter, HiddenNameTypeString
        store.SetValue "filler_" & counter, "value " & counter
    Next counter
End Sub

'@fun-title A workbook holding only the Dictionary sheet.
'@details
'The dictionary tests need a dictionary and a target workbook, nothing else.
'Building the six data sheets for them would cost more than the call being
'measured and would put the fixture build inside the number.
'@return Workbook. The new workbook, open and unsaved.
Private Function TimingDictionaryWorkbook() As Workbook
    Dim sourceBook As Workbook

    Set sourceBook = NewWorkbook()
    BuildDictionarySheet sourceBook

    Set TimingDictionaryWorkbook = sourceBook
End Function

'@sub-title Writes the dictionary size beside the dictionary numbers.
'@param sourceBook Workbook. The workbook holding the Dictionary sheet.
Private Sub LogDictionaryShape(ByVal sourceBook As Workbook)
    Dim sh As Worksheet
    Dim usedRng As Range

    On Error Resume Next
    Set sh = sourceBook.Worksheets(DICTIONARY_SHEET)
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    Set usedRng = sh.UsedRange
    If usedRng Is Nothing Then Exit Sub

    Assert.LogSuccesses "TIMING dictionary: " & (usedRng.Rows.Count - 1) & _
                        " rows over " & usedRng.Columns.Count & " columns"
End Sub

'@sub-title Closes two workbooks a dictionary timing test opened.
'@param sourceBook Workbook. The fixture, or Nothing.
'@param targetBook Workbook. The export target, or Nothing.
Private Sub DropBooks(ByVal sourceBook As Workbook, ByVal targetBook As Workbook)
    On Error Resume Next
    If Not targetBook Is Nothing Then DeleteWorkbook targetBook
    If Not sourceBook Is Nothing Then DeleteWorkbook sourceBook
    On Error GoTo 0
End Sub

'@sub-title Closes the source, the exporter and the file one timing test made.
'@param sourceBook Workbook. The source workbook, or Nothing.
'@param exporter LLExporter. The exporter to close, or Nothing.
'@param savedPath String. The saved export to remove, or an empty string.
Private Sub DropArtefacts(ByVal sourceBook As Workbook, _
                          ByVal exporter As LLExporter, _
                          ByVal savedPath As String)
    On Error Resume Next
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not sourceBook Is Nothing Then DeleteWorkbook sourceBook
    If LenB(savedPath) > 0 Then
        If Dir$(savedPath) <> vbNullString Then Kill savedPath
    End If
    On Error GoTo 0
End Sub
