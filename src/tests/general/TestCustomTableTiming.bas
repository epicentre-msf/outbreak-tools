Attribute VB_Name = "TestCustomTableTiming"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Times the two row removal branches of CustomTable")
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

'The sheet of formulas that reads the fixture, standing in for the analysis
'sheets of a linelist. Fewer blank rows here, because a delete that has to
'remap references is the slow case this block is looking for.
Private Const DEPENDENTS_SHEETNAME As String = "CustomTableTimingDeps"
Private Const DEPENDENTS_FORMULAS As Long = 1000
Private Const DEPENDENTS_BLANK_ROWS As Long = 300

Private Assert As CustomTest

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
