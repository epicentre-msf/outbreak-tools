Attribute VB_Name = "TestCustomTableTiming"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Times the two row removal branches of CustomTable")
'
'MEASURED -- 2026-09-01, macOS 27.0, Excel 16.111 headless, run-tests.R
'--build over a registry narrowed to helpers plus this module. Recorded
'here so the next reader has a yardstick without paying for a run. The
'numbers move with whatever else the machine is doing, so read each one
'beside its fixture line, which says what size it was taken at. This box
'is en_FR, so a number printed by Format$ carries a DECIMAL COMMA.
'
'   TIMING narrow ListRows branch: 0,031 s for 597 blank rows over 60
'       columns
'   TIMING narrow whole row branch: 0,016 s for 597 blank rows over 60
'       columns
'   TIMING narrow scan only: 0,008 s for 0 blank rows over 60 columns
'   TIMING narrow resize and clear: 0,000 s for 597 blank rows over 60
'       columns
'   TIMING wide plain ListRows branch: 0,008 s for 500 blank rows over 150
'       columns
'   TIMING wide plain whole row branch: 0,016 s for 500 blank rows over 150
'       columns
'   TIMING dressed scan only: 0,016 s for 0 blank rows over 150 columns
'   TIMING dressed ListRows branch: 0,016 s for 500 blank rows over 150
'       columns
'   TIMING dressed whole row branch: 0,023 s for 500 blank rows over 150
'       columns
'   TIMING dressed resize and clear: 0,008 s for 500 blank rows over 150
'       columns
'   TIMING watched ListRows branch: 0,016 s for 300 blank rows over 150
'       columns
'   TIMING watched whole row branch: 0,008 s for 300 blank rows over 150
'       columns
'   TIMING watched resize and clear: 0,000 s for 300 blank rows over 150
'       columns
'   TIMING recalculation of 1000 formulas: 0,008 s for 300 blank rows over
'       150 columns
'   TIMING add rows plain: 0,000 s for 199 blank rows over 150 columns
'   TIMING add rows dressed: 0,008 s for 199 blank rows over 150 columns
'   TIMING add rows dressed, no ids: 0,008 s for 199 blank rows over 150
'       columns
'   TIMING add rows watched: 0,008 s for 199 blank rows over 150 columns
'   TIMING handed-back recalculation of 1000 formulas: 0,008 s for 199 blank
'       rows over 150 columns
'   TIMING add rows with 5 VBA columns: 0,023 s for 199 blank rows over 150
'       columns
'   TIMING handed-back recalculation over 5 VBA columns: 0,023 s for 199
'       blank rows over 150 columns
'   TIMING add rows on a 2000 row table: 0,031 s for 199 blank rows over 150
'       columns
'   TIMING handed-back recalculation on a 2000 row table: 0,016 s for 199
'       blank rows over 150 columns
'   TIMING add rows on a 2000 row table, no ids: 0,023 s for 199 blank rows
'       over 150 columns
'
'@description
'   A measuring probe. It answers where the wait goes when a user presses
'   Resize on an HList sheet of a generated linelist. Every test asserts the
'   table ended the right shape and then logs how long the call took, so the
'   numbers land in the results CSV beside the pass.
'
'   The fixture matches the shape a data entry sheet reaches after three
'   clicks of Add rows: a wide table, a few filled rows at the top and about
'   six hundred blank ones under them.
'
'   The four numbers:
'     ListRows branch   what ClickResize runs today, one delete per blank row.
'     Whole row branch  what forceShift runs, one delete for the whole block.
'     Scan only         the same call over a table with nothing to remove,
'                       which is the block read and the cell walk alone.
'     Resize and clear  shrinking the ListObject and clearing the tail, the
'                       floor a trailing block fast path could reach.
'
'   Subtract Scan only from either branch to get the cost of the deletes.
'@depends CustomTable, CustomTest, TestHelpersLite

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TIMING_SHEETNAME As String = "CustomTableTiming"
Private Const TIMING_TABLENAME As String = "tblTiming"
Private Const TIMING_COLUMNS As Long = 60
Private Const TIMING_FILLED_ROWS As Long = 5
Private Const TIMING_BLANK_ROWS As Long = 597

'The realistic fixture. A generated HList sheet is wide, and its rows carry
'computed columns and dropdowns. Five hundred blank rows is under three clicks
'of Add rows, which keeps the probe inside one run.
Private Const WIDE_COLUMNS As Long = 150
Private Const WIDE_FILLED_ROWS As Long = 5
Private Const WIDE_BLANK_ROWS As Long = 500
Private Const DRESSED_FORMULA_COLUMNS As Long = 10
Private Const DRESSED_VALIDATION_COLUMNS As Long = 10

'How many columns of the fixture answer through VBA. A data entry sheet carries
'GEOCONCAT and the other custom functions in several columns, and a formula Excel
'has to cross into VBA for costs orders of magnitude more than one it answers
'itself. Nothing else in this module has such a column, which is why every number
'it reported before was at the resolution floor of Timer.
Private Const DRESSED_UDF_COLUMNS As Long = 5

'A table at the length a data entry sheet reaches after a few weeks of use. The
'short fixture above says what a press of Add rows costs on a table nobody has
'filled yet, which is not the table the complaint came from. Every cost that
'grows with the table rather than with the 199 rows added shows up as the
'difference between the two.
Private Const LONG_TABLE_ROWS As Long = 2000

'What one press of Add rows asks for on a data entry sheet.
'EventsLinelistButtons.ClickAddRows passes 199 on an HList sheet and 10 on a
'print sheet, so 199 is the number a user waits on.
Private Const ADDROWS_CLICK_ROWS As Long = 199

'The sheet of formulas that reads the fixture, standing in for the analysis
'sheets of a linelist. Fewer blank rows here, because a delete that has to
'remap references is the slow case this block is looking for.
Private Const DEPENDENTS_SHEETNAME As String = "CustomTableTimingDeps"
Private Const DEPENDENTS_FORMULAS As Long = 1000
Private Const DEPENDENTS_BLANK_ROWS As Long = 300

Private Assert As CustomTest

'@section A worksheet function answered by VBA
'===============================================================================

'@sub-title Stands in for GEOCONCAT and the other custom linelist functions
'@details Called from a worksheet formula, so every cell holding it costs Excel
'   a crossing into VBA. The body is deliberately trivial: what is being measured
'   is the crossing, not the work, because that is what a data entry sheet pays
'   per row and what nothing else in this module has ever included.
'
'   It lives here rather than in CustomLinelistFunctions because that module
'   belongs to the linelist and is not in this suite's compile closure.
'@param anchor Range. Any cell of the row, so the call is not volatile.
'@return Double. The row number of the anchor.
Public Function TIMINGVBACELL(ByVal anchor As Range) As Double
    If anchor Is Nothing Then Exit Function
    TIMINGVBACELL = anchor.Row
End Function

'@section Helpers
'===============================================================================

'@sub-title Builds a block of distinct text values
'@details Every cell holds its own row and column number, so no cell of the
'   block reads as empty and the header row comes out with unique names.
Private Function TextBlock(ByVal rowCount As Long, ByVal columnCount As Long) As Variant

    Dim block() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long

    ReDim block(1 To rowCount, 1 To columnCount)
    For rowIndex = 1 To rowCount
        For colIndex = 1 To columnCount
            block(rowIndex, colIndex) = "r" & rowIndex & "c" & colIndex
        Next colIndex
    Next rowIndex

    TextBlock = block
End Function

'@sub-title Builds the timing fixture table
'@details A ListObject of TIMING_COLUMNS columns holding filledRows rows of
'   text followed by blankRows empty ones. The sheet is rebuilt for every
'   test, because each test consumes the rows it measures.
'@param filledRows Long. Rows carrying a value in every column.
'@param blankRows Long. Empty rows under them.
'@return ListObject. The fixture table.
Private Function PrepareTimingTable(ByVal filledRows As Long, _
                                    ByVal blankRows As Long) As ListObject

    Dim sh As Worksheet
    Dim listRange As Range
    Dim timingTable As ListObject

    Set sh = EnsureWorksheet(TIMING_SHEETNAME)
    ClearWorksheet sh

    sh.Range("A1").Resize(1, TIMING_COLUMNS).Value = TextBlock(1, TIMING_COLUMNS)
    If filledRows > 0 Then
        sh.Range("A2").Resize(filledRows, TIMING_COLUMNS).Value = _
            TextBlock(filledRows, TIMING_COLUMNS)
    End If

    Set listRange = sh.Range("A1").Resize(1 + filledRows + blankRows, TIMING_COLUMNS)
    Set timingTable = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                         Source:=listRange, _
                                         XlListObjectHasHeaders:=xlYes)
    timingTable.Name = TIMING_TABLENAME

    Set PrepareTimingTable = timingTable
End Function

'@sub-title Writes one timing line into the results
'@param label String. Which branch was measured.
'@param elapsed Single. Seconds the measured call took.
'@param blankRows Long. How many blank rows it was handed.
Private Sub LogTiming(ByVal label As String, ByVal elapsed As Single, _
                      ByVal blankRows As Long, ByVal columnCount As Long)

    Assert.LogSuccesses "TIMING " & label & ": " & Format$(elapsed, "0.000") & _
                        " s for " & blankRows & " blank rows over " & _
                        columnCount & " columns"
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCustomTableTiming"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    DeleteWorksheet TIMING_SHEETNAME
    DeleteWorksheet DEPENDENTS_SHEETNAME
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanUp
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.FlushCurrentTest
    End If
End Sub

'@section Tests
'===============================================================================

'@sub-title Times the branch ClickResize runs today
'@details No shift tracker and no forceShift, so RemoveEmptyDataRows walks
'   DeleteRowsThroughTable and asks Excel for one ListRows delete per blank
'   row. Renumbering is left out so the number is the deletes alone.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeListRowsBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeListRowsBranch"
    On Error GoTo Fail

    Dim timingTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set timingTable = PrepareTimingTable(TIMING_FILLED_ROWS, TIMING_BLANK_ROWS)
    Set tbl = CustomTable.Create(timingTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual TIMING_FILLED_ROWS, timingTable.ListRows.Count, _
                    "The ListRows branch should keep the filled rows alone"
    LogTiming "narrow ListRows branch", elapsed, TIMING_BLANK_ROWS, TIMING_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeListRowsBranch", Err.Number, Err.Description
End Sub

'@sub-title Times the batched whole row branch
'@details forceShift sends the same call down the branch that gathers the
'   empty rows and deletes them in one go.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeWholeRowBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeWholeRowBranch"
    On Error GoTo Fail

    Dim timingTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set timingTable = PrepareTimingTable(TIMING_FILLED_ROWS, TIMING_BLANK_ROWS)
    Set tbl = CustomTable.Create(timingTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False, forceShift:=True
    elapsed = Timer - startedAt

    Assert.AreEqual TIMING_FILLED_ROWS, timingTable.ListRows.Count, _
                    "The whole row branch should keep the filled rows alone"
    LogTiming "narrow whole row branch", elapsed, TIMING_BLANK_ROWS, TIMING_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeWholeRowBranch", Err.Number, Err.Description
End Sub

'@sub-title Times the read and the walk with nothing to delete
'@details The table is the same size, and every row carries a value, so the
'   call reads the body, walks every cell and deletes nothing. This is the
'   part of both branches that no batching can remove.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeScanOnly()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeScanOnly"
    On Error GoTo Fail

    Dim timingTable As ListObject
    Dim tbl As CustomTable
    Dim totalRows As Long
    Dim startedAt As Single
    Dim elapsed As Single

    totalRows = TIMING_FILLED_ROWS + TIMING_BLANK_ROWS
    Set timingTable = PrepareTimingTable(totalRows, 0)
    Set tbl = CustomTable.Create(timingTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual totalRows, timingTable.ListRows.Count, _
                    "A table with no empty row should keep every row"
    LogTiming "narrow scan only", elapsed, 0, TIMING_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeScanOnly", Err.Number, Err.Description
End Sub

'@sub-title Times shrinking the table and clearing the tail
'@details The floor a trailing block fast path could reach: one Resize of the
'   ListObject and one Clear over the rows it let go, with no delete at all.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeResizeAndClear()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeResizeAndClear"
    On Error GoTo Fail

    Dim timingTable As ListObject
    Dim sh As Worksheet
    Dim headerRow As Long
    Dim headerCol As Long
    Dim keptRange As Range
    Dim tailRange As Range
    Dim startedAt As Single
    Dim elapsed As Single

    Set timingTable = PrepareTimingTable(TIMING_FILLED_ROWS, TIMING_BLANK_ROWS)
    Set sh = timingTable.Range.Worksheet
    headerRow = timingTable.Range.Row
    headerCol = timingTable.Range.Column

    Set keptRange = sh.Range(sh.Cells(headerRow, headerCol), _
                             sh.Cells(headerRow + TIMING_FILLED_ROWS, _
                                      headerCol + TIMING_COLUMNS - 1))
    Set tailRange = sh.Range(sh.Cells(headerRow + TIMING_FILLED_ROWS + 1, headerCol), _
                             sh.Cells(headerRow + TIMING_FILLED_ROWS + TIMING_BLANK_ROWS, _
                                      headerCol + TIMING_COLUMNS - 1))

    startedAt = Timer
    timingTable.Resize keptRange
    tailRange.Clear
    elapsed = Timer - startedAt

    Assert.AreEqual TIMING_FILLED_ROWS, timingTable.ListRows.Count, _
                    "The resize should leave the filled rows alone in the table"
    LogTiming "narrow resize and clear", elapsed, TIMING_BLANK_ROWS, TIMING_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeResizeAndClear", Err.Number, Err.Description
End Sub

'@section A realistic data entry table
'===============================================================================
'The narrow fixture above is a bare table. A real HList row carries computed
'columns, dropdowns and a conditional format, and the sheet is wide. These
'tests measure the same four things over that shape, so the difference between
'the two blocks says how much of the wait is the dressing.

'@sub-title Builds the wide fixture, plain or dressed
'@details Dressed adds an =ROW() formula to the first DRESSED_FORMULA_COLUMNS
'   columns, a list validation to the DRESSED_VALIDATION_COLUMNS after them,
'   one conditional format over the whole body and the autofilter. The formulas
'   answer a number, so a blank row reads as DRESSED_FORMULA_COLUMNS filled
'   cells and RemoveRows counts that threshold itself, the way the linelist
'   stores it under blank_row_count.
'@param filledRows Long. Rows carrying a value in every column.
'@param blankRows Long. Empty rows under them.
'@param dressed Boolean. When True, adds the formulas, dropdowns and format.
'@return ListObject. The fixture table.
Private Function PrepareWideTable(ByVal filledRows As Long, _
                                  ByVal blankRows As Long, _
                                  ByVal dressed As Boolean) As ListObject

    Dim sh As Worksheet
    Dim listRange As Range
    Dim wideTable As ListObject
    Dim bodyRange As Range
    Dim colRange As Range
    Dim colIndex As Long

    Set sh = EnsureWorksheet(TIMING_SHEETNAME)
    ClearWorksheet sh

    sh.Range("A1").Resize(1, WIDE_COLUMNS).Value = TextBlock(1, WIDE_COLUMNS)
    If filledRows > 0 Then
        sh.Range("A2").Resize(filledRows, WIDE_COLUMNS).Value = _
            TextBlock(filledRows, WIDE_COLUMNS)
    End If

    Set listRange = sh.Range("A1").Resize(1 + filledRows + blankRows, WIDE_COLUMNS)
    Set wideTable = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                       Source:=listRange, _
                                       XlListObjectHasHeaders:=xlYes)
    wideTable.Name = TIMING_TABLENAME

    If Not dressed Then
        Set PrepareWideTable = wideTable
        Exit Function
    End If

    For colIndex = 1 To DRESSED_FORMULA_COLUMNS
        Set colRange = wideTable.ListColumns(colIndex).DataBodyRange
        colRange.FormulaR1C1 = "=ROW()"
    Next colIndex

    For colIndex = DRESSED_FORMULA_COLUMNS + 1 To _
                   DRESSED_FORMULA_COLUMNS + DRESSED_VALIDATION_COLUMNS
        Set colRange = wideTable.ListColumns(colIndex).DataBodyRange
        colRange.Validation.Delete
        colRange.Validation.Add Type:=xlValidateList, _
                                AlertStyle:=xlValidAlertInformation, _
                                Formula1:="one,two,three"
    Next colIndex

    Set bodyRange = wideTable.DataBodyRange
    bodyRange.FormatConditions.Delete
    bodyRange.FormatConditions.Add Type:=xlExpression, Formula1:="=$A2=""x"""
    bodyRange.FormatConditions(1).Interior.Color = RGB(255, 235, 156)

    wideTable.ShowAutoFilter = True

    Set PrepareWideTable = wideTable
End Function

'@sub-title Times the ListRows branch over a wide plain table
'@TestMethod("CustomTableTiming")
Public Sub TestTimeWideListRowsBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeWideListRowsBranch"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, False)
    Set tbl = CustomTable.Create(wideTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The wide plain table should keep the filled rows alone"
    LogTiming "wide plain ListRows branch", elapsed, WIDE_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeWideListRowsBranch", Err.Number, Err.Description
End Sub

'@sub-title Times the whole row branch over a wide plain table
'@TestMethod("CustomTableTiming")
Public Sub TestTimeWideWholeRowBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeWideWholeRowBranch"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, False)
    Set tbl = CustomTable.Create(wideTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False, forceShift:=True
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The wide plain table should keep the filled rows alone"
    LogTiming "wide plain whole row branch", elapsed, WIDE_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeWideWholeRowBranch", Err.Number, Err.Description
End Sub

'@sub-title Times the read and the walk over a dressed table with nothing to delete
'@TestMethod("CustomTableTiming")
Public Sub TestTimeDressedScanOnly()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeDressedScanOnly"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim totalRows As Long
    Dim startedAt As Single
    Dim elapsed As Single

    totalRows = WIDE_FILLED_ROWS + WIDE_BLANK_ROWS
    Set wideTable = PrepareWideTable(totalRows, 0, True)
    Set tbl = CustomTable.Create(wideTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual totalRows, wideTable.ListRows.Count, _
                    "A dressed table with no empty row should keep every row"
    LogTiming "dressed scan only", elapsed, 0, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeDressedScanOnly", Err.Number, Err.Description
End Sub

'@sub-title Times the ListRows branch over a dressed table
'@details The closest this probe comes to the real thing: what a user waits
'   for when they press Resize on an HList sheet today.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeDressedListRowsBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeDressedListRowsBranch"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, True)
    Set tbl = CustomTable.Create(wideTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The dressed table should keep the filled rows alone"
    LogTiming "dressed ListRows branch", elapsed, WIDE_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeDressedListRowsBranch", Err.Number, Err.Description
End Sub

'@sub-title Times the whole row branch over a dressed table
'@TestMethod("CustomTableTiming")
Public Sub TestTimeDressedWholeRowBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeDressedWholeRowBranch"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, True)
    Set tbl = CustomTable.Create(wideTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False, forceShift:=True
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The dressed table should keep the filled rows alone"
    LogTiming "dressed whole row branch", elapsed, WIDE_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeDressedWholeRowBranch", Err.Number, Err.Description
End Sub

'@sub-title Times shrinking a dressed table and clearing its tail
'@TestMethod("CustomTableTiming")
Public Sub TestTimeDressedResizeAndClear()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeDressedResizeAndClear"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim sh As Worksheet
    Dim headerRow As Long
    Dim headerCol As Long
    Dim keptRange As Range
    Dim tailRange As Range
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, True)
    Set sh = wideTable.Range.Worksheet
    headerRow = wideTable.Range.Row
    headerCol = wideTable.Range.Column

    Set keptRange = sh.Range(sh.Cells(headerRow, headerCol), _
                             sh.Cells(headerRow + WIDE_FILLED_ROWS, _
                                      headerCol + WIDE_COLUMNS - 1))
    Set tailRange = sh.Range(sh.Cells(headerRow + WIDE_FILLED_ROWS + 1, headerCol), _
                             sh.Cells(headerRow + WIDE_FILLED_ROWS + WIDE_BLANK_ROWS, _
                                      headerCol + WIDE_COLUMNS - 1))

    startedAt = Timer
    wideTable.Resize keptRange
    tailRange.Clear
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The resize should leave the filled rows alone in the table"
    LogTiming "dressed resize and clear", elapsed, WIDE_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeDressedResizeAndClear", Err.Number, Err.Description
End Sub

'@section A table the rest of the workbook points at
'===============================================================================
'The two blocks above measure a table nothing else reads. In a generated
'linelist the HList table is read by the filtered companion, by the analysis
'sheets and by the pivot, so every delete makes Excel remap those references,
'and the busy state hands calculation back at the end. These tests put a sheet
'of formulas over the fixture and measure both costs.

'@sub-title Fills a second sheet with formulas reading the fixture table
'@details One COUNTIF per row over a column of the table, addressed the way an
'   analysis sheet addresses it, through the table name. The sheet is rebuilt
'   for every test, after the table it reads.
'@param formulaCount Long. How many formulas to write.
'@return Worksheet. The sheet holding them.
Private Function PrepareDependents(ByVal formulaCount As Long) As Worksheet

    Dim sh As Worksheet
    Dim formulaBlock() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long

    Set sh = EnsureWorksheet(DEPENDENTS_SHEETNAME)
    ClearWorksheet sh

    ReDim formulaBlock(1 To formulaCount, 1 To 1)
    For rowIndex = 1 To formulaCount
        colIndex = ((rowIndex - 1) Mod WIDE_COLUMNS) + 1
        formulaBlock(rowIndex, 1) = "=COUNTIF(" & TIMING_TABLENAME & "[r1c" & colIndex & _
                                "],""r2c" & colIndex & """)"
    Next rowIndex

    sh.Range("A1").Resize(formulaCount, 1).Formula = formulaBlock

    Set PrepareDependents = sh
End Function

'@sub-title Times the ListRows branch with the workbook reading the table
'@TestMethod("CustomTableTiming")
Public Sub TestTimeDependentsListRowsBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeDependentsListRowsBranch"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, DEPENDENTS_BLANK_ROWS, True)
    PrepareDependents DEPENDENTS_FORMULAS
    Set tbl = CustomTable.Create(wideTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The watched table should keep the filled rows alone"
    LogTiming "watched ListRows branch", elapsed, DEPENDENTS_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeDependentsListRowsBranch", Err.Number, Err.Description
End Sub

'@sub-title Times the whole row branch with the workbook reading the table
'@TestMethod("CustomTableTiming")
Public Sub TestTimeDependentsWholeRowBranch()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeDependentsWholeRowBranch"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, DEPENDENTS_BLANK_ROWS, True)
    PrepareDependents DEPENDENTS_FORMULAS
    Set tbl = CustomTable.Create(wideTable)

    startedAt = Timer
    tbl.RemoveRows totalCount:=0, includeIds:=False, forceShift:=True
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The watched table should keep the filled rows alone"
    LogTiming "watched whole row branch", elapsed, DEPENDENTS_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeDependentsWholeRowBranch", Err.Number, Err.Description
End Sub

'@sub-title Times shrinking a watched table and clearing its tail
'@TestMethod("CustomTableTiming")
Public Sub TestTimeDependentsResizeAndClear()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeDependentsResizeAndClear"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim sh As Worksheet
    Dim headerRow As Long
    Dim headerCol As Long
    Dim keptRange As Range
    Dim tailRange As Range
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, DEPENDENTS_BLANK_ROWS, True)
    PrepareDependents DEPENDENTS_FORMULAS

    Set sh = wideTable.Range.Worksheet
    headerRow = wideTable.Range.Row
    headerCol = wideTable.Range.Column
    Set keptRange = sh.Range(sh.Cells(headerRow, headerCol), _
                             sh.Cells(headerRow + WIDE_FILLED_ROWS, _
                                      headerCol + WIDE_COLUMNS - 1))
    Set tailRange = sh.Range(sh.Cells(headerRow + WIDE_FILLED_ROWS + 1, headerCol), _
                             sh.Cells(headerRow + WIDE_FILLED_ROWS + DEPENDENTS_BLANK_ROWS, _
                                      headerCol + WIDE_COLUMNS - 1))

    startedAt = Timer
    wideTable.Resize keptRange
    tailRange.Clear
    elapsed = Timer - startedAt

    Assert.AreEqual WIDE_FILLED_ROWS, wideTable.ListRows.Count, _
                    "The resize should leave the filled rows alone in the table"
    LogTiming "watched resize and clear", elapsed, DEPENDENTS_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeDependentsResizeAndClear", Err.Number, Err.Description
End Sub

'@sub-title Times the recalculation the busy state hands back at the end
'@details ClickResize runs under manual calculation and LLExitBusyState puts
'   automatic back, which recalculates. Nothing about batching the deletes
'   touches this number, so it is measured on its own.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeRecalculation()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeRecalculation"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim dependentsSheet As Worksheet
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, DEPENDENTS_BLANK_ROWS, True)
    Set dependentsSheet = PrepareDependents(DEPENDENTS_FORMULAS)

    startedAt = Timer
    dependentsSheet.Calculate
    elapsed = Timer - startedAt

    Assert.AreEqual DEPENDENTS_FORMULAS, dependentsSheet.Range("A1").Resize(DEPENDENTS_FORMULAS, 1).Count, _
                    "The dependents sheet should hold every formula"
    LogTiming "recalculation of " & DEPENDENTS_FORMULAS & " formulas", elapsed, _
              DEPENDENTS_BLANK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeRecalculation", Err.Number, Err.Description
End Sub

'@section Add rows
'===============================================================================
'The block above measures Resize, which is the button that takes rows away. Add
'rows is the other one and it was never measured here, so what follows is where
'a press of it goes.
'
'Four numbers, over the same dressed wide table the Resize block uses:
'  plain            the resize alone, on a table carrying no formulas,
'                   validation or conditional format.
'  dressed          the same call once the table is dressed. The difference
'                   between the two is what Excel spends carrying the dressing
'                   into 199 new rows.
'  without ids      dressed again with includeIds off, so subtracting it from
'                   dressed gives the cost of renumbering the whole ID column.
'  handed-back      the recalculation the old busy-state exit forced. AddRows
'  recalculation    runs under manual calculation and LLExitBusyState used to
'                   restore automatic, which recalculates every formula that
'                   reads the table. It is measured on its own because it
'                   belongs to the manager, not to CustomTable, and because it
'                   is the number the resting-state fix removes.

'@sub-title Times one press of Add rows over a wide plain table
'@TestMethod("CustomTableTiming")
Public Sub TestTimeAddRowsPlain()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeAddRowsPlain"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim rowsBefore As Long
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, False)
    Set tbl = CustomTable.Create(wideTable)
    rowsBefore = wideTable.ListRows.Count

    startedAt = Timer
    tbl.AddRows nbRows:=ADDROWS_CLICK_ROWS, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual rowsBefore + ADDROWS_CLICK_ROWS, wideTable.ListRows.Count, _
                    "The add should leave the table longer by the rows asked for"
    LogTiming "add rows plain", elapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeAddRowsPlain", Err.Number, Err.Description
End Sub

'@sub-title Times one press of Add rows over a dressed table, ids and all
'@details The whole of what CustomTable does for the click: the resize that
'   carries the formulas, the validation and the conditional format into the new
'   rows, the calculation of those rows, and the renumbering of the ID column.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeAddRowsDressed()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeAddRowsDressed"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim rowsBefore As Long
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, True)
    Set tbl = CustomTable.Create(wideTable)
    rowsBefore = wideTable.ListRows.Count

    startedAt = Timer
    tbl.AddRows nbRows:=ADDROWS_CLICK_ROWS
    elapsed = Timer - startedAt

    Assert.AreEqual rowsBefore + ADDROWS_CLICK_ROWS, wideTable.ListRows.Count, _
                    "The add should leave the table longer by the rows asked for"
    LogTiming "add rows dressed", elapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeAddRowsDressed", Err.Number, Err.Description
End Sub

'@sub-title Times the same press with the ID column left alone
'@details Subtract this from the dressed number to get what AddIds costs. It
'   rewrites every ID in the table, not only the new ones, so it grows with the
'   table rather than with the rows added.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeAddRowsWithoutIds()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeAddRowsWithoutIds"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim rowsBefore As Long
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, True)
    Set tbl = CustomTable.Create(wideTable)
    rowsBefore = wideTable.ListRows.Count

    startedAt = Timer
    tbl.AddRows nbRows:=ADDROWS_CLICK_ROWS, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual rowsBefore + ADDROWS_CLICK_ROWS, wideTable.ListRows.Count, _
                    "The add should leave the table longer by the rows asked for"
    LogTiming "add rows dressed, no ids", elapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeAddRowsWithoutIds", Err.Number, Err.Description
End Sub

'@sub-title Times the recalculation the busy-state exit used to hand back
'@details The two halves of one press, told apart. First the add itself under
'   manual calculation, which is what the user waits for now. Then the line
'   ApplicationState.Restore ran on the way out: putting automatic calculation
'   back, which recalculates every formula reading the table. The linelist rests
'   on manual, so that restore only ever happened because the open had failed to
'   make the resting mode stick, and the second number is what the fix removes.
'
'   The dependents sheet holds COUNTIF formulas, which are cheap. A data entry
'   sheet answers a GEOCONCAT call per row out of VBA, so the real workbook pays
'   more than this number, not less.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeAddRowsHandedBackRecalculation()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeAddRowsHandedBackRecalculation"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim rowsBefore As Long
    Dim startedAt As Single
    Dim addElapsed As Single
    Dim recalcElapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, DEPENDENTS_BLANK_ROWS, True)
    PrepareDependents DEPENDENTS_FORMULAS
    Set tbl = CustomTable.Create(wideTable)
    rowsBefore = wideTable.ListRows.Count

    'Everything open is settled before the clock starts. Application.Calculation
    'is an application setting, so restoring automatic recalculates every dirty
    'cell of every open workbook -- and the driver of this run carries fixture
    'sheets built under manual by the modules before this one. Settling first is
    'what makes the second number below the cost of THIS add and nothing else,
    'which is also the real situation: a linelist sitting settled under the
    'user's hands when they press the button.
    Application.Calculate

    'BusyApp left the application on manual, which is the mode the click runs
    'under either way.
    startedAt = Timer
    tbl.AddRows nbRows:=ADDROWS_CLICK_ROWS
    addElapsed = Timer - startedAt

    'The line the old exit ran. Automatic is put back and Excel recalculates
    'before the assignment returns.
    startedAt = Timer
    Application.Calculation = xlCalculationAutomatic
    recalcElapsed = Timer - startedAt
    Application.Calculation = xlCalculationManual

    Assert.AreEqual rowsBefore + ADDROWS_CLICK_ROWS, wideTable.ListRows.Count, _
                    "The add should leave the table longer by the rows asked for"
    LogTiming "add rows watched", addElapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    LogTiming "handed-back recalculation of " & DEPENDENTS_FORMULAS & " formulas", _
              recalcElapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    'The mode is put back on the failure path too: a probe that walks away from
    'automatic leaves every module after it recalculating on every write.
    Application.Calculation = xlCalculationManual
    CustomTestLogFailure Assert, "TestTimeAddRowsHandedBackRecalculation", _
                         Err.Number, Err.Description
End Sub

'@sub-title Times a press of Add rows on a table whose columns answer through VBA
'@details The measurement the four above could not make. Their fixture holds
'   =ROW() and the dependents sheet holds COUNTIF, and Excel answers both itself
'   in well under a microsecond, so every number they reported was one or two
'   ticks of Timer and said nothing. This one puts DRESSED_UDF_COLUMNS columns of
'   TIMINGVBACELL on the table, which is the shape a data entry sheet has.
'
'   Two halves, told apart:
'     add           the click as it runs now. Manual calculation throughout, and
'                   CalculateAddedRows settles the rows the add brought in.
'     handed back   the line ApplicationState.Restore used to run on the way out
'                   of the busy state, putting automatic calculation back. The
'                   session is settled first, so what it recalculates is what
'                   this add dirtied and nothing left over from another module.
'
'   A real linelist is longer than this fixture and carries more such columns, so
'   read the two numbers as a ratio rather than as seconds a user would feel.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeAddRowsWithVbaFormulas()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeAddRowsWithVbaFormulas"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim colRange As Range
    Dim colIndex As Long
    Dim firstUdfColumn As Long
    Dim rowsBefore As Long
    Dim startedAt As Single
    Dim addElapsed As Single
    Dim recalcElapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, WIDE_BLANK_ROWS, True)
    PrepareDependents DEPENDENTS_FORMULAS

    'After the =ROW() columns and the validated ones, so no cell of this block
    'reads itself.
    firstUdfColumn = DRESSED_FORMULA_COLUMNS + DRESSED_VALIDATION_COLUMNS + 1
    For colIndex = firstUdfColumn To firstUdfColumn + DRESSED_UDF_COLUMNS - 1
        Set colRange = wideTable.ListColumns(colIndex).DataBodyRange
        colRange.FormulaR1C1 = "=TIMINGVBACELL(RC1)"
    Next colIndex

    Set tbl = CustomTable.Create(wideTable)
    rowsBefore = wideTable.ListRows.Count

    'Everything open is settled before the clock starts. Calculation is an
    'application setting, so the restore below would otherwise recalculate every
    'dirty cell of every workbook this run has built.
    Application.Calculate

    startedAt = Timer
    tbl.AddRows nbRows:=ADDROWS_CLICK_ROWS
    addElapsed = Timer - startedAt

    startedAt = Timer
    Application.Calculation = xlCalculationAutomatic
    recalcElapsed = Timer - startedAt
    Application.Calculation = xlCalculationManual

    Assert.AreEqual rowsBefore + ADDROWS_CLICK_ROWS, wideTable.ListRows.Count, _
                    "The add should leave the table longer by the rows asked for"
    LogTiming "add rows with " & DRESSED_UDF_COLUMNS & " VBA columns", _
              addElapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    LogTiming "handed-back recalculation over " & DRESSED_UDF_COLUMNS & " VBA columns", _
              recalcElapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    Application.Calculation = xlCalculationManual
    CustomTestLogFailure Assert, "TestTimeAddRowsWithVbaFormulas", Err.Number, Err.Description
End Sub

'@sub-title Times a press of Add rows on a table already thousands of rows long
'@details Same shape as the test above and the same two halves, on a table of
'   LONG_TABLE_ROWS rows instead of WIDE_BLANK_ROWS. Read the pair against that
'   one: what grows with the table is paid on every press for the life of the
'   sheet, and what does not is paid once per press whatever the sheet holds.
'
'   AddIds is the known member of the first group. It rewrites every identifier
'   in the table, not only the ones the add brought in, so the no-ids number
'   beside it is what says how much of the press that is.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeAddRowsOnALongTable()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeAddRowsOnALongTable"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim colRange As Range
    Dim colIndex As Long
    Dim firstUdfColumn As Long
    Dim rowsBefore As Long
    Dim startedAt As Single
    Dim addElapsed As Single
    Dim recalcElapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, LONG_TABLE_ROWS, True)
    PrepareDependents DEPENDENTS_FORMULAS

    firstUdfColumn = DRESSED_FORMULA_COLUMNS + DRESSED_VALIDATION_COLUMNS + 1
    For colIndex = firstUdfColumn To firstUdfColumn + DRESSED_UDF_COLUMNS - 1
        Set colRange = wideTable.ListColumns(colIndex).DataBodyRange
        colRange.FormulaR1C1 = "=TIMINGVBACELL(RC1)"
    Next colIndex

    Set tbl = CustomTable.Create(wideTable)
    rowsBefore = wideTable.ListRows.Count

    Application.Calculate

    startedAt = Timer
    tbl.AddRows nbRows:=ADDROWS_CLICK_ROWS
    addElapsed = Timer - startedAt

    startedAt = Timer
    Application.Calculation = xlCalculationAutomatic
    recalcElapsed = Timer - startedAt
    Application.Calculation = xlCalculationManual

    Assert.AreEqual rowsBefore + ADDROWS_CLICK_ROWS, wideTable.ListRows.Count, _
                    "The add should leave the table longer by the rows asked for"
    LogTiming "add rows on a " & LONG_TABLE_ROWS & " row table", _
              addElapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    LogTiming "handed-back recalculation on a " & LONG_TABLE_ROWS & " row table", _
              recalcElapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    Application.Calculation = xlCalculationManual
    CustomTestLogFailure Assert, "TestTimeAddRowsOnALongTable", Err.Number, Err.Description
End Sub

'@sub-title Times the same long press with the ID column left alone
'@details Subtract this from the number above to get what AddIds costs on a long
'   table. It is the one part of a press that grows with the sheet rather than
'   with the rows added.
'@TestMethod("CustomTableTiming")
Public Sub TestTimeAddRowsOnALongTableWithoutIds()
    CustomTestSetTitles Assert, "CustomTableTiming", "TestTimeAddRowsOnALongTableWithoutIds"
    On Error GoTo Fail

    Dim wideTable As ListObject
    Dim tbl As CustomTable
    Dim colRange As Range
    Dim colIndex As Long
    Dim firstUdfColumn As Long
    Dim rowsBefore As Long
    Dim startedAt As Single
    Dim elapsed As Single

    Set wideTable = PrepareWideTable(WIDE_FILLED_ROWS, LONG_TABLE_ROWS, True)

    firstUdfColumn = DRESSED_FORMULA_COLUMNS + DRESSED_VALIDATION_COLUMNS + 1
    For colIndex = firstUdfColumn To firstUdfColumn + DRESSED_UDF_COLUMNS - 1
        Set colRange = wideTable.ListColumns(colIndex).DataBodyRange
        colRange.FormulaR1C1 = "=TIMINGVBACELL(RC1)"
    Next colIndex

    Set tbl = CustomTable.Create(wideTable)
    rowsBefore = wideTable.ListRows.Count

    Application.Calculate

    startedAt = Timer
    tbl.AddRows nbRows:=ADDROWS_CLICK_ROWS, includeIds:=False
    elapsed = Timer - startedAt

    Assert.AreEqual rowsBefore + ADDROWS_CLICK_ROWS, wideTable.ListRows.Count, _
                    "The add should leave the table longer by the rows asked for"
    LogTiming "add rows on a " & LONG_TABLE_ROWS & " row table, no ids", _
              elapsed, ADDROWS_CLICK_ROWS, WIDE_COLUMNS
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTimeAddRowsOnALongTableWithoutIds", _
                         Err.Number, Err.Description
End Sub
