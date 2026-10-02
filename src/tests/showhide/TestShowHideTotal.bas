Attribute VB_Name = "TestShowHideTotal"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Times a whole show/hide session on a big sheet")
'
'MEASURED -- 2026-09-01, macOS 27.0, Excel 16.111 headless, run-tests.R
'--build over a registry narrowed to helpers plus this module. Recorded
'here so the next reader has a yardstick without paying for a run. The
'numbers move with whatever else the machine is doing, so read each one
'beside its fixture line, which says what size it was taken at. This box
'is en_FR, so a number printed by Format$ carries a DECIMAL COMMA.
'
'   TOTAL open the form: 1.344 s
'   TOTAL one section press: 0.391 s
'   TOTAL close the form: 0.352 s
'   TOTAL FULL SHOW/HIDE SESSION: 2.086 s
'   TOTAL writes the layout refused during the session: 0
'   TOTAL N size writes, which the form open no longer pays: 2.438 s
'   TOTAL writes the layout refused while every column was shown: 0
'   TOTAL the first position whose width write was refused: 0
'   TOTAL fixture: V=2000 N=800 H=200 section=100, sheet columns 16384,
'       platform macOS
'   TOTAL open, build the entry list: 0.055 s
'   TOTAL open, read the store with stored sizes: 1.328 s
'   TOTAL open, read the store with no stored sizes: 0.211 s
'   TOTAL writes the layout refused over the breakdown: 0
'   TOTAL fixture: V=2000 N=800 H=200 section=100, sheet columns 16384,
'       platform macOS
'
'@description
'   A MEASURING PROBE, not a behaviour suite, and it stays commented out of the
'   registry the way TestShowHideTiming and TestLLExporterTiming do. Every test
'   here reports a number rather than a pass.
'
'   WHAT IT IS FOR
'   ---------------------------------------------------------------------------
'   One number for a whole show/hide session on a sheet with a lot of variables:
'   the open, one section press and the close, added up. TestShowHideTiming
'   times the five class calls one at a time on a small sheet; this one times
'   what a user waits for, at a size a big linelist reaches.
'
'   IT COMPILES AGAINST THE WHOLE SPEED BLOCK, BEFORE AND AFTER
'   ---------------------------------------------------------------------------
'   THIS IS THE POINT OF THE MODULE and every call below is chosen for it. It
'   uses only what ShowHide, ShowHideLayout and ShowHideStore carried before
'   Session 127 and still carry now:
'
'     ShowHide.Create / EntryCount / IsFree / PositionIndex / SetHidden / Apply
'     ShowHideLayout.Create / Size / SetSize / FailureCount
'     ShowHideStore.CreateOnSheet / Load(entries, layout) / Save(entries, layout)
'
'   So the same file runs in a worktree at the pre-block commit and in the tree
'   today, and the two answers are read side by side. Nothing here passes
'   `writeSizes`, `force` or a handed-in store: those arrived with Sessions 130
'   and 131 and naming one would stop the file compiling against the old code.
'   ADDING SUCH A CALL BREAKS THE COMPARISON THIS MODULE EXISTS FOR.
'
'   WHAT THE `Load` LINE DOES AND DOES NOT SAY
'   ---------------------------------------------------------------------------
'   `Load(entries, layout)` writes the stored size of every entry in both
'   versions, so the open below pays N size writes on both sides and the
'   comparison is fair. The form does NOT pay them any more: since Session 130
'   the open passes `writeSizes:=False`. The last line of the test times those
'   N writes on their own, so what the real open saves can be taken off the
'   open figure.
'
'   The layout carries no Passwords, so no protection bracket and no
'   ScreenUpdating toggle is inside any number here. Seconds go through Str$,
'   because this box is en_FR and Format$ would write a comma.
'@depends ShowHide, ShowHideLayout, ShowHideStore, LLdictionary, CustomTest, TestHelpersLite, DictionaryTestFixture

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHideTotal"

'THE FIXTURE SIZE. A big linelist rather than the audit's guess: the small
'fixture of TestShowHideTiming hides what the walk costs when a sheet is wide.
Private Const VARIABLE_COUNT As Long = 2000
Private Const SHEET_VARS As Long = 800
Private Const HIDDEN_ENTRIES As Long = 200
Private Const SECTION_VARS As Long = 100

'How many variables each of the other sheets carries, so V is bigger than N
Private Const OTHER_SHEET_VARS As Long = 100

Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const TIMED_SHEET As String = "timing-sheet"
Private Const OTHER_SHEET_PREFIX As String = "timing-other"
Private Const STORE_SHEET As String = "__show_hide"

'What the dictionary calls a horizontal data entry sheet
Private Const HLIST_TYPE As String = "hlist2D"

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private Dict As LLdictionary
Private SetupError As Long
Private SetupMessage As String


'@section Module lifecycle
'===============================================================================

'@sub-title Build the assertion harness, the fixture workbook and the dictionary.
'@details
'Public because the harness calls it by name through Application.Run. An error
'escaping here is a modal dialog, which stops the whole run, so the setup
'captures its error and FixtureReady reports it as the test's failure.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestShowHideTotal"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FixtureWorkbook = WideWorkbook()
        BuildDictionarySheet FixtureWorkbook
        EnsureWorksheet TIMED_SHEET, FixtureWorkbook
        EnsureWorksheet STORE_SHEET, FixtureWorkbook

        Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
        Dict.Prepare

        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print the results and drop the fixture workbook.
'@details
'Public because the harness calls it by name through Application.Run. The
'driver is brought to the front before PrintResults: it raises 1004 whenever
'another workbook holds the screen.
'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
        ThisWorkbook.Activate
    On Error GoTo 0

    If Not Assert Is Nothing Then Assert.PrintResults TEST_OUTPUT_SHEET

    Set Dict = Nothing
    Set FixtureWorkbook = Nothing
    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanUp
Public Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush
End Sub


'@section The test
'===============================================================================

'@sub-title Times a whole show/hide session: the open, one section press, the close.
'@details
'The three steps a user waits for, in the order the form pays them, on a sheet
'of N variables with a store that already holds its rows:
'
'  open            ShowHide.Create walks the dictionary, ShowHideStore.Load
'                  reads the table and writes the stored sizes
'  section press   SECTION_VARS entries are marked hidden and Apply reconciles
'                  the sheet
'  close           ShowHideStore.Save reads the table, measures every entry and
'                  writes the table back
'
'The store and the sheet are put in the state a real open finds them in before
'the clock starts: the rows are saved once and H entries are already hidden.
'@TestMethod("ShowHideTotal")
Public Sub TestTimeFullShowHideSession()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeFullShowHideSession"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim seed As ShowHide
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim store As ShowHideStore
    Dim hiddenCount As Long
    Dim pressed As Long
    Dim startedAt As Double
    Dim elapsed As Double
    Dim sessionSeconds As Double
    Dim refusedInSession As Long

    If Not FixtureReady("TestTimeFullShowHideSession") Then Exit Sub

    Set sh = FixtureWorkbook.Worksheets(TIMED_SHEET)
    ResetSheetGeometry sh

    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)
    Set store = ShowHideStore.CreateOnSheet(FixtureWorkbook.Worksheets(STORE_SHEET))

    'The state a real open finds: the store holds this sheet's rows and H
    'entries are hidden on the sheet. None of this is timed.
    Set seed = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)
    hiddenCount = HideFirstFreeEntries(seed, HIDDEN_ENTRIES, 1)
    Assert.IsTrue (hiddenCount = HIDDEN_ENTRIES), _
                  "H entries were hidden before the clock started: " & _
                  CStr(hiddenCount) & " of " & CStr(HIDDEN_ENTRIES)
    seed.Apply layout
    store.Save seed, layout

    'The open
    startedAt = Timer
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)
    store.Load entries, layout
    elapsed = ElapsedSince(startedAt)
    sessionSeconds = elapsed
    LogTiming "open the form", elapsed

    Assert.IsTrue (entries.EntryCount = SHEET_VARS), _
                  "The entry list holds the N variables of the sheet: " & _
                  CStr(entries.EntryCount) & " of " & CStr(SHEET_VARS)

    'One section press: a run of entries changes state and the sheet follows
    startedAt = Timer
    pressed = HideFirstFreeEntries(entries, SECTION_VARS, HIDDEN_ENTRIES + 1)
    entries.Apply layout
    elapsed = ElapsedSince(startedAt)
    sessionSeconds = sessionSeconds + elapsed
    LogTiming "one section press", elapsed

    Assert.IsTrue (pressed = SECTION_VARS), _
                  "The press moved a section's worth of entries: " & _
                  CStr(pressed) & " of " & CStr(SECTION_VARS)

    'The close
    startedAt = Timer
    store.Save entries, layout
    elapsed = ElapsedSince(startedAt)
    sessionSeconds = sessionSeconds + elapsed
    LogTiming "close the form", elapsed

    LogTiming "FULL SHOW/HIDE SESSION", sessionSeconds

    'How many writes Excel refused over the whole session, reported rather
    'than asserted on: this is a measuring module, and a refusal count is a
    'reading like the others. A hidden column refuses a width write on this
    'box, so the count is expected to be non-zero while columns are hidden.
    refusedInSession = layout.FailureCount
    LogCount "writes the layout refused during the session", refusedInSession

    'The N size writes, on their own. Load pays them on both sides of the
    'comparison; the form stopped paying them at Session 130, so this is what
    'comes off the open figure to read what a form open costs today. The sheet
    'is shown again first, so every write lands and the number is write time
    'rather than error time.
    ResetSheetGeometry sh
    startedAt = Timer
    WriteEverySize entries, layout
    elapsed = ElapsedSince(startedAt)
    LogTiming "N size writes, which the form open no longer pays", elapsed
    LogCount "writes the layout refused while every column was shown", _
             layout.FailureCount - refusedInSession
    LogCount "the first position whose width write was refused", _
             FirstRefusedPosition(entries, layout)

    LogFixtureShape
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeFullShowHideSession", Err.Number, Err.Description
End Sub



'@sub-title Splits the open into its two calls, with and without stored sizes.
'@details
'The open is ShowHide.Create then ShowHideStore.Load, and the session test
'reports the pair as one number. This one takes them apart, and takes Load
'twice over:
'
'  build the entry list        ShowHide.Create walks the dictionary
'  read the store, sizes kept  Load matches N rows and writes every stored
'                              width to the sheet
'  read the store, no sizes    the same Load against a store whose widths were
'                              never recorded, so the size branch is skipped
'
'The third is what the form pays today. Since Session 130 the open passes
'`writeSizes:=False`, and a store saved with no layout holds no widths, so
'both routes do the same work and this one is portable to the old code, which
'has no such argument. The difference between the second and the third is the
'N size writes.
'@TestMethod("ShowHideTotal")
Public Sub TestTimeOpenBreakdown()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeOpenBreakdown"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim seed As ShowHide
    Dim fresh As ShowHide
    Dim layout As ShowHideLayout
    Dim store As ShowHideStore
    Dim matched As Long
    Dim startedAt As Double
    Dim elapsed As Double

    If Not FixtureReady("TestTimeOpenBreakdown") Then Exit Sub

    Set sh = FixtureWorkbook.Worksheets(TIMED_SHEET)
    ResetSheetGeometry sh

    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)
    Set store = ShowHideStore.CreateOnSheet(FixtureWorkbook.Worksheets(STORE_SHEET))

    'The store as a close leaves it: the rows, and a width for every entry
    Set seed = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)
    HideFirstFreeEntries seed, HIDDEN_ENTRIES, 1
    seed.Apply layout
    store.Save seed, layout

    startedAt = Timer
    Set fresh = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)
    elapsed = ElapsedSince(startedAt)
    LogTiming "open, build the entry list", elapsed

    startedAt = Timer
    matched = store.Load(fresh, layout)
    elapsed = ElapsedSince(startedAt)
    LogTiming "open, read the store with stored sizes", elapsed

    Assert.IsTrue (matched = SHEET_VARS), _
                  "The load matched the N rows of the sheet: " & _
                  CStr(matched) & " of " & CStr(SHEET_VARS)

    'The same rows with no widths recorded: a save with no layout writes
    'visibility alone, which leaves entry_size empty on every row
    store.Save seed

    Set fresh = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)

    startedAt = Timer
    matched = store.Load(fresh, layout)
    elapsed = ElapsedSince(startedAt)
    LogTiming "open, read the store with no stored sizes", elapsed

    Assert.IsTrue (matched = SHEET_VARS), _
                  "The load without sizes matched the same N rows: " & _
                  CStr(matched) & " of " & CStr(SHEET_VARS)

    LogCount "writes the layout refused over the breakdown", layout.FailureCount
    LogFixtureShape
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeOpenBreakdown", Err.Number, Err.Description
End Sub

'@section Reporting
'===============================================================================

'@sub-title Writes one count into the results.
'@param label String. What was counted.
'@param value Long. The count.
Private Sub LogCount(ByVal label As String, ByVal value As Long)
    Assert.LogSuccesses "TOTAL " & label & ": " & CStr(value)
End Sub

'@sub-title Writes one timing line into the results.
'@param label String. What was measured.
'@param elapsed Double. Seconds it took.
Private Sub LogTiming(ByVal label As String, ByVal elapsed As Double)
    Assert.LogSuccesses "TOTAL " & label & ": " & SecondsText(elapsed) & " s"
End Sub

'@sub-title Writes the fixture size into the results beside the numbers.
'@details A timing line means nothing without the shape it was measured on.
Private Sub LogFixtureShape()
    Assert.LogSuccesses "TOTAL fixture: V=" & VARIABLE_COUNT & " N=" & SHEET_VARS & _
                        " H=" & HIDDEN_ENTRIES & " section=" & SECTION_VARS & _
                        ", sheet columns " & _
                        CStr(FixtureWorkbook.Worksheets(TIMED_SHEET).Columns.Count) & _
                        ", platform " & PlatformTag()
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

'@fun-title Whether the module setup got through, filing its error otherwise.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is there to time.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "ModuleInitialize failed: " & SetupMessage
End Function


'@section The fixture
'===============================================================================

'@fun-title A new workbook whose sheets are wider than 256 columns.
'@details
'THE FIXTURE MUST BE WIDER THAN THE SHEET IT MEASURES. `Workbooks.Add` follows
'Excel's default save format, and on a box set to the 97-2003 format that is a
'workbook of 256 columns. Every ShowHideLayout write past column 256 is then
'refused, the walk measures an error path rather than a write, and the reading
'is worthless without saying so. The default format is moved to the current one
'for the Add and put back at once, so nothing else on the machine sees it
'changed.
'@return Workbook. A new workbook of the current file format.
Private Function WideWorkbook() As Workbook
    Dim heldFormat As Long

    BusyApp
    heldFormat = Application.DefaultSaveFormat

    'xlOpenXMLWorkbook. Named by its value because the constant is not
    'available in every host this module compiles in.
    Application.DefaultSaveFormat = 51
    Set WideWorkbook = Workbooks.Add
    Application.DefaultSaveFormat = heldFormat
End Function

'@sub-title Writes a dictionary of V variables, N of them on the timed sheet.
'@details
'The header row is the one the shared dictionary fixture uses, because the
'dictionary is read by column name. Every row is optional and visible, so every
'entry follows the user. Prepare derives `column index`, one per variable in
'dictionary order, which is the position the layout writes.
'@param wkb Workbook. The workbook to write into.
Private Sub BuildDictionarySheet(ByVal wkb As Workbook)
    Dim sh As Worksheet
    Dim headers As Variant
    Dim block() As Variant
    Dim headerCount As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim otherIndex As Long

    headers = DictionaryTestFixture.DictionaryFixtureHeaders()
    headerCount = UBound(headers) - LBound(headers) + 1

    Set sh = EnsureWorksheet(DICTIONARY_SHEET, wkb, clearSheet:=True)
    sh.Range("A1").Resize(1, headerCount).Value = headers

    ReDim block(1 To VARIABLE_COUNT, 1 To headerCount)

    For rowIndex = 1 To VARIABLE_COUNT
        For colIndex = 1 To headerCount
            block(rowIndex, colIndex) = vbNullString
        Next colIndex

        'Four characters at least: Prepare refuses a shorter variable name
        block(rowIndex, HeaderPosition(headers, "Variable Name")) = "tvar" & rowIndex
        block(rowIndex, HeaderPosition(headers, "Main Label")) = "Label " & rowIndex
        block(rowIndex, HeaderPosition(headers, "Sheet Type")) = HLIST_TYPE
        block(rowIndex, HeaderPosition(headers, "Main Section")) = _
            "section " & (((rowIndex - 1) \ SECTION_VARS) + 1)
        block(rowIndex, HeaderPosition(headers, "Status")) = "optional, visible"
        block(rowIndex, HeaderPosition(headers, "Variable Type")) = "text"
        block(rowIndex, HeaderPosition(headers, "Control")) = "free"
        block(rowIndex, HeaderPosition(headers, "Personal Identifier")) = "no"

        If rowIndex <= SHEET_VARS Then
            block(rowIndex, HeaderPosition(headers, "Sheet Name")) = TIMED_SHEET
        Else
            otherIndex = ((rowIndex - SHEET_VARS - 1) \ OTHER_SHEET_VARS) + 1
            block(rowIndex, HeaderPosition(headers, "Sheet Name")) = _
                OTHER_SHEET_PREFIX & otherIndex
        End If
    Next rowIndex

    sh.Range("A2").Resize(VARIABLE_COUNT, headerCount).Value = block
End Sub

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

'@fun-title Marks a run of free entries hidden, starting at one index.
'@details The start index is what keeps the seed hides and the section press
'   off each other: the press moves entries the seed left alone.
'@param entries ShowHide. The list to mark.
'@param wanted Long. How many to hide.
'@param fromIndex Long. The entry index to start looking at.
'@return Long. How many were marked.
Private Function HideFirstFreeEntries(ByVal entries As ShowHide, _
                                      ByVal wanted As Long, _
                                      ByVal fromIndex As Long) As Long
    Dim idx As Long
    Dim done As Long

    For idx = fromIndex To entries.EntryCount
        If done >= wanted Then Exit For
        If entries.IsFree(idx) And entries.PositionIndex(idx) > 0 Then
            entries.SetHidden idx, True
            done = done + 1
        End If
    Next idx

    HideFirstFreeEntries = done
End Function

'@sub-title Writes a size to every positioned entry of the list.
'@details What ShowHideStore.Load does on top of its table read when it is
'   asked for sizes. Timed on its own so the open figure can be read with and
'   without it.
'@param entries ShowHide. The list whose positions are written.
'@param layout ShowHideLayout. The sheet to write on.
Private Sub WriteEverySize(ByVal entries As ShowHide, ByVal layout As ShowHideLayout)
    Dim idx As Long
    Dim position As Long

    For idx = 1 To entries.EntryCount
        position = entries.PositionIndex(idx)
        If position > 0 Then layout.SetSize position, 10
    Next idx
End Sub

'@fun-title The first position whose size write Excel refuses.
'@details A width write that fails costs the sheet the width it was told to
'   keep, and nothing upstream reads FailureCount, so the loss is silent. This
'   walks the positions one at a time and answers the first that refuses, so
'   the boundary is a measured number rather than one inferred from a count.
'@param entries ShowHide. The list whose positions are written.
'@param layout ShowHideLayout. The sheet to write on.
'@return Long. The first refusing position, or 0 when every write landed.
Private Function FirstRefusedPosition(ByVal entries As ShowHide, _
                                      ByVal layout As ShowHideLayout) As Long
    Dim idx As Long
    Dim position As Long
    Dim before As Long

    For idx = 1 To entries.EntryCount
        position = entries.PositionIndex(idx)
        If position > 0 Then
            before = layout.FailureCount
            layout.SetSize position, 11
            If layout.FailureCount > before Then
                FirstRefusedPosition = position
                Exit Function
            End If
        End If
    Next idx
End Function

'@sub-title Shows every column of the timed sheet again and resets its widths.
'@details Nothing here deletes a worksheet: Worksheet.Delete is unreliable on
'   macOS Excel. The columns the layout wrote are put back by hand.
'@param sh Worksheet. The sheet to reset.
Private Sub ResetSheetGeometry(ByVal sh As Worksheet)
    Dim span As Range

    On Error Resume Next
        Set span = sh.Range(sh.Cells(1, 1), sh.Cells(1, SHEET_VARS + 20))
        span.EntireColumn.Hidden = False
        span.EntireColumn.ColumnWidth = 10
    On Error GoTo 0
End Sub
