Attribute VB_Name = "TestShowHideTiming"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Times the show/hide classes against a sized fixture")
'
'@description
'   A MEASURING PROBE, not a behaviour suite. It stays commented out of the
'   registry the way TestLLExporterTiming does: every test here reports a
'   number rather than a pass, and the numbers move with whatever else the
'   machine is doing. Uncomment the row, run, read the TIMING lines out of the
'   run CSV, comment it back out.
'
'   WHAT IT IS FOR
'   ---------------------------------------------------------------------------
'   The show/hide speed block (`.obt/implementations.md` Sessions 127 to 132)
'   changes five calls one session at a time: ShowHide.Create, ShowHideStore.Load,
'   ShowHide.Apply, ShowHideStore.Save and SectionMap.Create. This module times
'   each of them on the same fixture, so a session can be read against the one
'   before it. To compare two versions of a class, run this module, check the
'   other version out, run it again, and put the file back. Nothing here touches
'   a private member, so the same module compiles against both.
'
'   The step lines LLLog writes on a real linelist are the other yardstick; this
'   module is the one that runs without a person clicking.
'
'   WHAT THE FIXTURE SIZE MEANS, AND WHICH COSTS IT CAN SEE
'   ---------------------------------------------------------------------------
'   The audit (`.obt/plans/showhide-speed.md`) counts every cost in five
'   numbers, and each constant below is one of them:
'
'     V   VARIABLE_COUNT   variables in the whole dictionary. ShowHide.Create
'                          walks all of them (2V) to find the N of its sheet.
'     N   SHEET_VARS       variables on the timed sheet. Create reads seven more
'                          columns for each (14N); Apply writes N positions.
'     H   HIDDEN_ENTRIES   entries hidden when Save runs. Save shows and hides
'                          each one again to read its width (2H).
'     M   HIDDEN_NAMES     hidden names on the sheet. SectionMap.Create builds a
'                          HiddenNames over all of them (3M).
'     B   SECTION_BLOCKS   section blocks on the sheet.
'
'   What the fixture cannot see: the protection bracket. The layout is built
'   with no Passwords, so no test here pays a Protect or an Unprotect, and the
'   ScreenUpdating toggle of the bracket is not in these numbers either. Session
'   129 is about the bracket and is measured on a real linelist through the
'   log, not here.
'
'   The dictionary is one block write and one Prepare, so building the fixture
'   is not inside any number. Seconds go through Str$, because this box is
'   en_FR and Format$ would write a comma.
'@depends ShowHide, ShowHideLayout, ShowHideStore, SectionMap, LLdictionary, HiddenNames, CustomTest, TestHelpersLite, DictionaryTestFixture

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ShowHideTiming"

'THE FIXTURE SIZE. The audit's guess at a real linelist, so the numbers here
'read against the counts in the audit at the same scale.
Private Const VARIABLE_COUNT As Long = 400
Private Const SHEET_VARS As Long = 150
Private Const HIDDEN_ENTRIES As Long = 30
Private Const HIDDEN_NAMES As Long = 300
Private Const SECTION_BLOCKS As Long = 12

'How many variables each of the other sheets of the dictionary carries. They
'exist so that V is bigger than N: Create has to walk past them.
Private Const OTHER_SHEET_VARS As Long = 50

Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const TIMED_SHEET As String = "timing-sheet"
Private Const OTHER_SHEET_PREFIX As String = "timing-other"
Private Const STORE_SHEET As String = "__show_hide"
Private Const SECTIONS_SHEET As String = "timing-sections"

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
'Public because the harness calls it by name through Application.Run. The
'fixture is built once: every test reads the same dictionary, and a Prepare
'over V rows per test would be paid five times for nothing.
'
'An error escaping here is a modal dialog, which stops the whole run, so the
'setup captures its error and FixtureReady reports it as each test's failure.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestShowHideTiming"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set FixtureWorkbook = NewWorkbook()
        BuildDictionarySheet FixtureWorkbook
        EnsureWorksheet TIMED_SHEET, FixtureWorkbook
        EnsureWorksheet STORE_SHEET, FixtureWorkbook
        EnsureWorksheet SECTIONS_SHEET, FixtureWorkbook

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
    If Not Assert Is Nothing Then Assert.FlushCurrentTest
End Sub


'@section Tests
'===============================================================================

'@sub-title Times ShowHide.Create over V variables with N on the sheet.
'@details
'This is the walk Session 128 rewrites to read the dictionary in blocks. The
'audit counts it at 2V + 14N crossings; the time here is what that costs on
'this box.
'@TestMethod("ShowHideTiming")
Public Sub TestTimeCreate()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeCreate"
    On Error GoTo TestFail

    Dim entries As ShowHide
    Dim startedAt As Double
    Dim elapsed As Double

    If Not FixtureReady("TestTimeCreate") Then Exit Sub

    startedAt = Timer
    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (entries.EntryCount = SHEET_VARS), _
                  "The entry list holds the N variables of the sheet: " & _
                  CStr(entries.EntryCount) & " of " & CStr(SHEET_VARS)

    LogTiming "ShowHide.Create", elapsed
    LogFixtureShape
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeCreate", Err.Number, Err.Description
End Sub

'@sub-title Times Apply, Save, Load and a second Apply on one sheet.
'@details
'The four calls are run in the order a session pays them, on one entry list
'and one layout, so each number is taken with the sheet in the state the real
'call finds it in:
'
'  apply, H hidden   the entries hold H hidden entries, the sheet shows all of
'                    them; Apply writes N positions and hides H columns
'  save, H hidden    Save measures every entry, showing and re-hiding the H
'                    hidden columns to read their widths (2N + 2H)
'  load              a fresh entry list reads the N rows back and the layout
'                    takes N size writes
'  apply, in step    the sheet already matches the entries; today Apply still
'                    writes all N positions, and Session 130 is what changes that
'
'The layout carries no Passwords, so no bracket is paid: see the module note.
'@TestMethod("ShowHideTiming")
Public Sub TestTimeApplySaveLoad()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeApplySaveLoad"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim store As ShowHideStore
    Dim fresh As ShowHide
    Dim hiddenCount As Long
    Dim matched As Long
    Dim startedAt As Double
    Dim elapsed As Double

    If Not FixtureReady("TestTimeApplySaveLoad") Then Exit Sub

    Set sh = FixtureWorkbook.Worksheets(TIMED_SHEET)
    ResetSheetGeometry sh

    Set entries = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)
    Set layout = ShowHideLayout.Create(sh, ShowHideLayerHList)
    Set store = ShowHideStore.CreateOnSheet(FixtureWorkbook.Worksheets(STORE_SHEET))

    hiddenCount = HideFirstFreeEntries(entries, HIDDEN_ENTRIES)
    Assert.IsTrue (hiddenCount = HIDDEN_ENTRIES), _
                  "H free entries were marked hidden before the timed calls: " & _
                  CStr(hiddenCount) & " of " & CStr(HIDDEN_ENTRIES)

    startedAt = Timer
    entries.Apply layout
    elapsed = ElapsedSince(startedAt)
    LogTiming "ShowHide.Apply, H hidden, sheet not in step", elapsed

    startedAt = Timer
    store.Save entries, layout
    elapsed = ElapsedSince(startedAt)
    LogTiming "ShowHideStore.Save, H hidden", elapsed

    Set fresh = ShowHide.Create(Dict, ShowHideLayerHList, TIMED_SHEET)

    startedAt = Timer
    matched = store.Load(fresh, layout)
    elapsed = ElapsedSince(startedAt)
    LogTiming "ShowHideStore.Load, with sizes", elapsed

    Assert.IsTrue (matched = SHEET_VARS), _
                  "Load matched the N rows Save wrote: " & CStr(matched) & _
                  " of " & CStr(SHEET_VARS)

    startedAt = Timer
    fresh.Apply layout
    elapsed = ElapsedSince(startedAt)
    LogTiming "ShowHide.Apply, sheet already in step", elapsed

    Assert.IsTrue (layout.FailureCount = 0), _
                  "The layout refused no write, so the numbers timed real writes"

    LogFixtureShape
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeApplySaveLoad", Err.Number, Err.Description
End Sub

'@sub-title Times SectionMap.Create on a sheet carrying M hidden names.
'@details
'SectionMap.Create builds a HiddenNames over the sheet, which walks every name
'and reads three properties off each. The map's own blocks are a few names among
'the M; the M is what the walk pays for. Session 131 hands the map the store the
'event service already holds.
'@TestMethod("ShowHideTiming")
Public Sub TestTimeSectionMapCreate()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeSectionMapCreate"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim secMap As SectionMap
    Dim startedAt As Double
    Dim elapsed As Double

    If Not FixtureReady("TestTimeSectionMapCreate") Then Exit Sub

    Set sh = FixtureWorkbook.Worksheets(SECTIONS_SHEET)
    AddFillerNames sh, HIDDEN_NAMES
    WriteSectionBlocks sh, SECTION_BLOCKS

    startedAt = Timer
    Set secMap = SectionMap.Create(sh)
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (secMap.Count = SECTION_BLOCKS), _
                  "The map read the B blocks back: " & CStr(secMap.Count) & _
                  " of " & CStr(SECTION_BLOCKS)

    LogTiming "SectionMap.Create, M hidden names", elapsed
    LogFixtureShape
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSectionMapCreate", Err.Number, Err.Description
End Sub

'@sub-title Times SectionMap.Create handed a store already built over the sheet.
'@details
'What the section button pays since Session 131: the event service holds one
'HiddenNames per sheet, and the map reads its B blocks through it. The store
'is built outside the clock, the way EventLinelist builds it once per sheet,
'so the number beside TestTimeSectionMapCreate's is the walk of M names that
'the press no longer pays.
'@TestMethod("ShowHideTiming")
Public Sub TestTimeSectionMapCreateWithStore()
    CustomTestSetTitles Assert, TESTMODULE, "TestTimeSectionMapCreateWithStore"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim store As HiddenNames
    Dim secMap As SectionMap
    Dim startedAt As Double
    Dim elapsed As Double

    If Not FixtureReady("TestTimeSectionMapCreateWithStore") Then Exit Sub

    Set sh = FixtureWorkbook.Worksheets(SECTIONS_SHEET)
    AddFillerNames sh, HIDDEN_NAMES
    WriteSectionBlocks sh, SECTION_BLOCKS
    Set store = HiddenNames.Create(sh)

    startedAt = Timer
    Set secMap = SectionMap.Create(sh, store)
    elapsed = ElapsedSince(startedAt)

    Assert.IsTrue (secMap.Count = SECTION_BLOCKS), _
                  "The map read the B blocks back through the store: " & _
                  CStr(secMap.Count) & " of " & CStr(SECTION_BLOCKS)

    LogTiming "SectionMap.Create, store handed in", elapsed
    LogFixtureShape
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTimeSectionMapCreateWithStore", Err.Number, Err.Description
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
    Assert.LogSuccesses "TIMING fixture: V=" & VARIABLE_COUNT & " N=" & SHEET_VARS & _
                        " H=" & HIDDEN_ENTRIES & " M=" & HIDDEN_NAMES & _
                        " B=" & SECTION_BLOCKS & ", platform " & PlatformTag()
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

'@sub-title Writes a dictionary of V variables, N of them on the timed sheet.
'@details
'The header row is the one the shared dictionary fixture uses, because the
'dictionary is read by column name. The first N rows name the timed sheet and
'the rest are spread over other sheets in runs of OTHER_SHEET_VARS, so Create
'walks past V - N rows that are not its own. Every row is optional and visible,
'so every entry follows the user and the H hidden ones are chosen by the test.
'Prepare derives `column index`, one per variable in dictionary order, which is
'the position the layout writes.
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
            "section " & (((rowIndex - 1) \ (SHEET_VARS \ SECTION_BLOCKS)) + 1)
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

'@fun-title Marks the first free entries of a list hidden.
'@param entries ShowHide. The list to mark.
'@param wanted Long. How many to hide.
'@return Long. How many were marked, which is fewer only when the list is short.
Private Function HideFirstFreeEntries(ByVal entries As ShowHide, ByVal wanted As Long) As Long
    Dim idx As Long
    Dim done As Long

    For idx = 1 To entries.EntryCount
        If done >= wanted Then Exit For
        If entries.IsFree(idx) And entries.PositionIndex(idx) > 0 Then
            entries.SetHidden idx, True
            done = done + 1
        End If
    Next idx

    HideFirstFreeEntries = done
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

'@sub-title Puts M filler names onto one sheet.
'@details A generated data entry sheet carries one hidden name per variable plus
'   its own metadata, and SectionMap.Create walks all of them. The names are
'   written through HiddenNames so they are the shape the walk reads.
'@param sh Worksheet. The sheet to write on.
'@param count Long. How many names.
Private Sub AddFillerNames(ByVal sh As Worksheet, ByVal count As Long)
    Dim store As HiddenNames
    Dim counter As Long

    Set store = HiddenNames.Create(sh)

    For counter = 1 To count
        store.EnsureName "filler_" & counter, "value " & counter, HiddenNameTypeString
        store.SetValue "filler_" & counter, "value " & counter
    Next counter
End Sub

'@sub-title Writes B section blocks on one sheet, side by side.
'@param sh Worksheet. The sheet to write on.
'@param count Long. How many blocks.
Private Sub WriteSectionBlocks(ByVal sh As Worksheet, ByVal count As Long)
    Dim secMap As SectionMap
    Dim counter As Long
    Dim width As Long

    width = SHEET_VARS \ count
    If width < 1 Then width = 1

    Set secMap = SectionMap.Create(sh)
    secMap.Clear
    For counter = 1 To count
        secMap.Add "section " & counter, (counter - 1) * width + 1, counter * width
    Next counter
End Sub
