VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Data manipulation in a worksheet"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Data manipulation in a worksheet")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, IIfSideEffect, UseMeaningfulName, HungarianNotation

Option Explicit

'@class DataSheet
'@description
'Works with a block of data on a worksheet, marked out by a start row and a
'start column. The block is plain cells and needs no ListObject. ColumnExists
'and ColumnIndex look a column up by its header, HeaderRange and DataRange
'answer the ranges, RenameColumn rewrites a header, FilterData and FiltersData
'filter on one column or on several, and Import, Export and ImportFormat move
'values and formats between this block and another DataSheet or a CustomTable.
'
'WHAT THE INSTANCE HOLDS
'-------------------------------------------------------------------------------
'The start row and the start column are sealed at creation. The rest of the
'geometry is worked out on each call: DataEndRow, DataEndColumn and the header
'row are read again every time they are asked for, so a caller can insert and
'delete rows and columns while this object stays alive. Within one call the
'header row is read as a single block.
'@depends Checking, BetterArray, HiddenNames
'@version 1.0 (2026-02-09)

Private Type TDataSheet
    sh As Worksheet                        'Worksheet of the data
    startLn As Long                              'The line where the data starts
    startCl As Long                             'The column where the data starts (the first column is the column of the variable names.
    forceEnd As Boolean
    hasChecks As Boolean
    objName As String
    sealed As Boolean
    formatColumns As BetterArray           'Columns whose formats travel with the data
    checkCounter As Long                   'Unique key for each logged message
    internalChecks As Checking
End Type

Private this As TDataSheet

'@section Factory
'===============================================================================
'@description Entry point for creating DataSheet instances.

'@label:create
'@sub-title Create a DataSheet wrapper around a worksheet
'@details
'Wraps the supplied worksheet with convenience methods for column
'lookup, filtering, import/export, and formatting transfer. The data
'structure is identified by a starting row and column; end row and
'column are computed dynamically. Unlike CustomTable, no ListObject
'is required.
'@param sh Worksheet. The worksheet hosting the data.
'@param startLn Long. The 1-based header row index.
'@param startCl Long. The 1-based starting column index.
'@param forceEndRow Optional Boolean. When True, scans all columns to find the last row. Defaults to False.
'@param objName Optional String. Object name. Defaults to the worksheet name.
'@return DataSheet. A fully initialised DataSheet instance.
'@throws ProjectError.ObjectNotInitialized When sh is Nothing.
'@throws ProjectError.InvalidArgument When startLn or startCl is less than 1.
'@depends Checking, BetterArray
'@export
Public Function Create(ByVal sh As Worksheet, ByVal startLn As Long, _
                       ByVal startCl As Long, _
                       Optional ByVal forceEndRow As Boolean = False, _
                       Optional ByVal objName As String = vbNullString) As DataSheet

    'Throw illegal errors before the creation of the dictionary

    If sh Is Nothing Then ThrowErrorEmptySheetGiven
    If startLn < 1 Then ThrowErrorStartNumber startLn, "start row"
    If startCl < 1 Then ThrowErrorStartNumber startCl, "start column"

    Dim newDataSheet As DataSheet
    Set newDataSheet = New DataSheet

    Set newDataSheet.Wksh = sh
    newDataSheet.DataStartRow = startLn
    newDataSheet.DataStartColumn = startCl
    newDataSheet.StrictEnd = forceEndRow
    newDataSheet.Name = objName
    newDataSheet.Seal
    Set Create = newDataSheet
End Function

'@label:wksh-set
'@prop-title Assign the worksheet reference
'@param sh Worksheet. The worksheet to assign.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sh = sh
End Property

'@label:data-start-row-set
'@prop-title Assign the header row index
'@param startLn Long. The 1-based header row index.
Public Property Let DataStartRow(ByVal startLn As Long)
    GuardNotSealed "DataStartRow"
    this.startLn = startLn
End Property

'@label:data-start-column-set
'@prop-title Assign the starting column index
'@param startCl Long. The 1-based starting column index.
Public Property Let DataStartColumn(ByVal startCl As Long)
    GuardNotSealed "DataStartColumn"
    this.startCl = startCl
End Property

'@label:strict-end-set
'@prop-title Assign the strict end row flag
'@param forceEnd Boolean. When True, scans all columns for the last row.
Public Property Let StrictEnd(ByVal forceEnd As Boolean)
    GuardNotSealed "StrictEnd"
    this.forceEnd = forceEnd
End Property

'@label:name-set
'@prop-title Assign the DataSheet name
'@param objName String. The name to assign.
Public Property Let Name(ByVal objName As String)
    GuardNotSealed "Name"
    this.objName = objName
End Property


'@section Elements
'===============================================================================
'@description Properties that expose the worksheet layout and column access.

'@label:wksh
'@prop-title Worksheet backing the DataSheet
'@details
'Returns the worksheet that hosts the table-like data.
'@return Worksheet. The host worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:data-start-row
'@prop-title First row of data (header row)
'@return Long. The 1-based header row index.
'@export
Public Property Get DataStartRow() As Long
    DataStartRow = this.startLn
End Property

'@label:data-start-column
'@prop-title First column of data
'@return Long. The 1-based starting column index.
'@export
Public Property Get DataStartColumn() As Long
    DataStartColumn = this.startCl
End Property

'@label:strict-end
'@prop-title Whether to scan all columns for the last row
'@return Boolean. True when all columns are scanned.
Public Property Get StrictEnd() As Boolean
    StrictEnd = this.forceEnd
End Property

'@label:name
'@prop-title Name of the DataSheet
'@details
'Returns the assigned object name. When no name was provided during
'creation, defaults to the worksheet name.
'@return String. The DataSheet name.
'@export
Public Property Get Name() As String

    If this.objName = vbNullString Then
        this.objName = Wksh.Name
    End If

    Name = this.objName
End Property

'@label:data-end-row
'@prop-title Last row containing data
'@details
'Dynamically computes the last occupied row. When StrictEnd is True,
'scans all non-formula columns for the furthest row. Guarantees at
'least one data row below the header.
'This is read fresh every single time. It is the value that moves most:
'the class itself collapses it in Clean and grows it back in Import, and
'callers insert and delete body rows without telling this object.
'@return Long. The 1-based last data row index.
'@export
Public Property Get DataEndRow() As Long

    Dim endrw As Long
    Dim headerRng As Range
    Dim headers As Variant
    Dim sh As Worksheet
    Dim counter As Long
    Dim columnNumber As Long
    Dim colEndRw As Long 'column end row, used in loop

    Set sh = Wksh()

    If StrictEnd() Then
        'The header row is read once as a block. Walking it cell by cell
        'with Offset used to cost one worksheet read per column just to
        'ask whether the cell was empty.
        Set headerRng = HeaderRange()
        headers = RowValues(headerRng)
        endrw = headerRng.Row

        For counter = 1 To UBound(headers)
            'Stop at the first empty header cell. That rule is NOT the same
            'as DataEndColumn, which jumps over a hole with End(xlToLeft).
            'The two disagree when the header row has a gap, and both are
            'pinned by the tests. Keep each as it is.
            If IsEmpty(headers(counter)) Then Exit For

            columnNumber = headerRng.Column + counter - 1
            'For each column, search and update the last row. Columns with
            'formulas are not included in the search
            If Not sh.Cells(DataStartRow(), columnNumber).HasFormula Then
                colEndRw = sh.Cells(sh.Rows.Count, columnNumber).End(xlUp).Row
                If endrw < colEndRw Then endrw = colEndRw
            End If
        Next
    Else
        endrw = sh.Cells(sh.Rows.Count, DataStartColumn()).End(xlUp).Row
    End If

    'In case the endrow is the headerRow, add 1
    If (endrw = DataStartRow()) Then
        DataEndRow = DataStartRow() + 1
    Else
        DataEndRow = endrw
    End If
End Property

'@label:data-end-column
'@prop-title Last column containing data
'@return Long. The 1-based last column index.
'@export
Public Property Get DataEndColumn() As Long
    With Wksh
        DataEndColumn = .Cells(DataStartRow(), .Columns.Count).End(xlToLeft).Column
    End With
End Property

'@label:header-range
'@prop-title Header row Range
'@details
'Returns the Range spanning from the start column to the end column
'on the header row.
'@return Range. The header row range.
'@export
Public Property Get HeaderRange() As Range
    'updating the header range
    Set HeaderRange = SheetBlock(DataStartRow(), DataStartColumn(), _
                                 DataStartRow(), DataEndColumn())
End Property

'@label:sheet-block
'@sub-title Build a Range on the host worksheet from four numbers
'@details
'The "With sh ... .Range(.Cells(a, b), .Cells(c, d))" shape was written
'out in five places. One small builder replaces all of them.
'@param firstRow Long. Top row.
'@param firstColumn Long. Left column.
'@param lastRow Long. Bottom row.
'@param lastColumn Long. Right column.
'@return Range. The block.
Private Function SheetBlock(ByVal firstRow As Long, ByVal firstColumn As Long, _
                            ByVal lastRow As Long, ByVal lastColumn As Long) As Range
    With Wksh
        Set SheetBlock = .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn))
    End With
End Function

'@label:row-values
'@sub-title Read a one-row Range as a 1-based list of values
'@details
'Range.Value hands back a plain value when the Range is one cell, and a
'two-dimensional array otherwise. This flattens both cases so callers see
'one shape and do not repeat the test.
'@param blockRng Range. A single-row block.
'@return Variant. A 1-based one-dimensional array of cell values.
Private Function RowValues(ByVal blockRng As Range) As Variant
    Dim rawValues As Variant
    Dim flat() As Variant
    Dim counter As Long
    Dim nbColumns As Long

    nbColumns = blockRng.Columns.Count
    ReDim flat(1 To nbColumns)

    If nbColumns = 1 Then
        flat(1) = blockRng.Cells(1, 1).Value
    Else
        rawValues = blockRng.Value
        For counter = 1 To nbColumns
            flat(counter) = rawValues(1, counter)
        Next
    End If

    RowValues = flat
End Function

'@label:cell-text
'@sub-title Read one cell value as text without raising
'@details
'A cell holding an error value (#N/A and the like) makes CStr raise a
'type mismatch. Such a header can never match a column name, so it reads
'as empty text instead of stopping the caller.
'@param cellValue Variant. The raw cell value.
'@return String. The value as text, or empty text.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:find-column
'@sub-title Find one column in the header row, reading that row once
'@details
'This is the single lookup behind ColumnExists, ColumnIndex, RenameColumn
'and AddFormatsColumns. Each of them used to build the header Range and
'run Range.Find on its own, so one ColumnIndex call built the range twice
'and searched twice. Here the row is read once and matched in memory. The
'out argument carries the third answer: the header text as the sheet
'spells it.
'The four match rules are kept exactly as Range.Find applied them:
'strictSearch True is whole-string equality (xlWhole) and False is a
'substring test (xlPart); matchCase True compares binary and False
'compares text.
'The header block is deliberately NOT kept on the instance. Callers write
'to these sheets without going through this object, so it is read again
'on the next call.
'@param colName String. Column header to look for.
'@param strictSearch Boolean. True for whole-string matching.
'@param matchCase Boolean. True for case-sensitive matching.
'@param matchedHeader String. Out. The header text exactly as the sheet spells it.
'@return Long. The worksheet column of the match, or -1 when there is none.
Private Function FindColumn(ByVal colName As String, _
                            ByVal strictSearch As Boolean, _
                            ByVal matchCase As Boolean, _
                            Optional ByRef matchedHeader As String) As Long

    Dim headerRng As Range
    Dim headers As Variant
    Dim currentHeader As String
    Dim compareMethod As Long
    Dim isMatch As Boolean
    Dim counter As Long

    matchedHeader = vbNullString
    FindColumn = -1

    'An empty name never matches, so answer it without touching the sheet.
    If colName = vbNullString Then Exit Function

    compareMethod = IIF(matchCase, vbBinaryCompare, vbTextCompare)

    Set headerRng = HeaderRange()
    headers = RowValues(headerRng)

    For counter = 1 To UBound(headers)
        currentHeader = CellText(headers(counter))

        If strictSearch Then
            isMatch = (StrComp(currentHeader, colName, compareMethod) = 0)
        Else
            isMatch = (InStr(1, currentHeader, colName, compareMethod) > 0)
        End If

        If isMatch Then
            matchedHeader = currentHeader
            FindColumn = headerRng.Column + counter - 1
            Exit Function
        End If
    Next
End Function

'@label:column-exists
'@sub-title Check whether a column exists in the DataSheet
'@details
'Searches the header row for the specified column name.
'Returns False for empty column names.
'@param colName String. Column header to search for.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return Boolean. True when the column is found.
'@export
Public Function ColumnExists(ByVal colName As String, _
                              Optional ByVal strictSearch As Boolean = True, _
                              Optional ByVal matchCase As Boolean = True) As Boolean
    'The empty name is answered before any worksheet read.
    If colName = vbNullString Then Exit Function
    ColumnExists = (FindColumn(colName, strictSearch, matchCase) > 0)
End Function


'@label:column-index
'@sub-title Resolve the column index for a header name
'@details
'Returns the worksheet column index, or the data-range-relative index
'when inDataRange is True. Returns -1 when the column is not found and
'shouldExist is False; raises an error when shouldExist is True.
'@param colName String. Column header to look up.
'@param inDataRange Optional Boolean. When True, returns the index relative to the data range. Defaults to False.
'@param shouldExist Optional Boolean. When True, raises an error if not found. Defaults to False.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return Long. The column index, or -1 when not found.
'@throws ProjectError.ElementNotFound When the column is not found and shouldExist is True.
'@export
Public Function ColumnIndex(ByVal colName As String, _
                             Optional ByVal inDataRange As Boolean = False, _
                             Optional ByVal shouldExist As Boolean = False, _
                             Optional ByVal strictSearch As Boolean = True, _
                             Optional ByVal matchCase As Boolean = True) As Long
    Dim colIndex As Long

    colIndex = FindColumn(colName, strictSearch, matchCase)

    If colIndex < 0 Then
        If shouldExist Then ThrowErrorUnFoundColumn colName
        'By default, unfound columns will have -1 as Index
        ColumnIndex = -1
        Exit Function
    End If

    If inDataRange Then
        ColumnIndex = colIndex - DataStartColumn() + 1
    Else
        ColumnIndex = colIndex
    End If
End Function


'@label:data-range
'@prop-title Range of a column or the entire data body
'@details
'Returns a data Range for the specified column, or the entire data body
'when colName is "__all__". When strictSearch is False, partial matching
'is used. Use strictSearch when column names share common text.
'@param colName Optional String. Column header to retrieve. Defaults to "__all__".
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return Range. The requested data range.
'@export
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = True, _
                              Optional ByVal matchCase As Boolean = True) As Range
    Dim customStartRow As Long
    Dim endCol As Long
    Dim customStartCol As Long

    customStartRow = DataStartRow() + 1
    If includeHeaders Then customStartRow = DataStartRow()

    If colName = "__all__" Then
        customStartCol = DataStartColumn()
        endCol = DataEndColumn()
    Else
        customStartCol = ColumnIndex(colName, shouldExist:=True, strictSearch:=strictSearch, matchCase:=matchCase)
        endCol = customStartCol
    End If

    Set DataRange = SheetBlock(customStartRow, customStartCol, DataEndRow(), endCol)
End Property


'@label:rename-column
'@sub-title Rename a header in the DataSheet
'@details
'Replaces the header text of the matched column with newName. Raises
'an error when newName is empty or the column is not found.
'@param currentName String. Current column header to find.
'@param newName String. New header text to assign.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to False.
'@throws ProjectError.InvalidArgument When newName is empty.
'@throws ProjectError.ElementNotFound When the column is not found.
'@export
Public Sub RenameColumn(ByVal currentName As String, _
                        ByVal newName As String, _
                        Optional ByVal strictSearch As Boolean = True, _
                        Optional ByVal matchCase As Boolean = False)

    Dim headerColumn As Long

    If LenB(newName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "New column name cannot be empty"
    End If

    headerColumn = FindColumn(currentName, strictSearch, matchCase)
    If headerColumn < 0 Then ThrowErrorUnFoundColumn currentName

    'The header row moves under this object, so the cell is taken straight
    'from the column that was just found rather than from a second build
    'of the header Range.
    Wksh.Cells(DataStartRow(), headerColumn).Value = newName
End Sub

'@section Operations
'===============================================================================
'@description Filtering, cleaning, import, and export operations.

'@label:clean
'@sub-title Clear all non-formula columns
'@details
'Clears the body of every column that does not hold a formula in its
'first data cell. Formula columns are generated and must survive an
'import.
'The column range is built from the loop position. Looking each header
'text back up paid for a full column lookup on a column whose place was
'already known, and it had two faults: a blank header inside the range
'made the lookup raise and killed the whole import, and two columns
'sharing a header cleared the first twice and the second never.
'The last row is read once, before the first clear. Reading it per column
'shrinks it as soon as the first column is emptied, so the columns after
'it were left with old values below the first row.
Private Sub Clean()
    Dim headerRng As Range
    Dim headers As Variant
    Dim columnRng As Range
    Dim counter As Long
    Dim columnNumber As Long
    Dim firstRow As Long
    Dim lastRow As Long

    Set headerRng = HeaderRange()
    headers = RowValues(headerRng)
    firstRow = DataStartRow() + 1
    lastRow = DataEndRow()

    For counter = 1 To UBound(headers)
        columnNumber = headerRng.Column + counter - 1
        Set columnRng = SheetBlock(firstRow, columnNumber, lastRow, columnNumber)
        'clear only "non formula" columns
        If (Not columnRng.Cells(1, 1).HasFormula) Then columnRng.ClearContents
    Next
End Sub

'@label:remove-autofilter
'@sub-title Turn off AutoFilter on the host worksheet
'@details
'Reliably removes any active AutoFilter. Called before and after
'filtering operations to guarantee a clean state.
'Known, and left as it is on purpose: this clears EVERY filter on the host
'sheet, including one the user set by hand on a ListObject of their own.
'The two filter routines below need the sheet unfiltered to read the right
'rows, and a stray filter left behind hides data from the user, which is
'worse than losing their view. Do not narrow this sweep without a way to
'prove the sheet ends up unfiltered on every path.
Private Sub RemoveAutoFilter()
    Dim sh As Worksheet
    Dim Lo As ListObject

    Set sh = Wksh()

    'Nothing to undo when the sheet carries no filter and holds no table.
    'This routine runs twice per filter call, so the cheap exit matters.
    If (Not sh.AutoFilterMode) And (sh.ListObjects.Count = 0) Then Exit Sub

    For Each lo In sh.ListObjects
        If Not lo.AutoFilter Is Nothing Then
            On Error Resume Next
                lo.AutoFilter.ShowAllData
            On Error GoTo 0
        End If
    Next lo

    'ShowAllData unhides rows hidden by filter criteria.
    'It errors when no criteria are active, so guard with On Error.
    On Error Resume Next
    sh.AutoFilter.ShowAllData
    On Error GoTo 0

    'Remove the AutoFilter entirely (dropdown arrows)
    If sh.AutoFilterMode Then sh.AutoFilterMode = False
End Sub


'@label:area-values
'@sub-title Read one solid Range as a 1-based two-dimensional array
'@details
'Range.Value hands back a plain value when the Range is a single cell and
'a two-dimensional array otherwise. This gives callers one shape so they do
'not repeat the test.
'@param blockRng Range. One solid block, never a multi-area Range.
'@return Variant. A two-dimensional array, 1 To rows, 1 To columns.
Private Function AreaValues(ByVal blockRng As Range) As Variant
    Dim oneCell(1 To 1, 1 To 1) As Variant

    If blockRng.Cells.Count = 1 Then
        oneCell(1, 1) = blockRng.Value
        AreaValues = oneCell
    Else
        AreaValues = blockRng.Value
    End If
End Function

'@label:visible-block
'@sub-title Read a filtered Range as one solid 1-based two-dimensional array
'@details
'SpecialCells(xlCellTypeVisible) hands back one Area per unbroken run of
'visible rows, and Range.Rows.Count on such a Range answers for Areas(1)
'ONLY. Anything sized from that number is short as soon as the filter
'leaves two blocks, which is what used to break the two routines below:
'they sized a scratch block on the sheet from it, pasted the WHOLE
'selection into it (Excel anchors the paste at the top-left and writes
'everything, a smaller destination neither truncates nor raises), read the
'short block back and cleared the short block. So the caller got a
'truncated list and the rows past the block stayed on the worksheet.
'Every area is read with one .Value instead and the areas are stacked in
'the order Excel hands them out, which is top to bottom. Nothing is written
'to the worksheet, so there is no block to clear, no clipboard to reset and
'no fixed spot past the last column to collide with.
'BetterArray.FromExcelRange cannot do this job: it reads .Row, .Column,
'.Rows.Count and .Columns.Count off the Range - all Areas(1) answers - and
'then rebuilds one contiguous block from those bounds, which would pull in
'the hidden rows sitting inside the first area.
'Only the ROW runs are taken from the areas. A column hidden inside the
'block splits the areas sideways as well, and each row run then shows up
'more than once, so a run already taken is never taken twice and every run
'is read across the WHOLE requested column span. That also keeps a hidden
'column from stopping the routine: Copy raises 1004 on areas that do not
'line up.
'@param visibleRng Range. The visible part of a filtered block.
'@param firstColumn Long. Left worksheet column of the block that was filtered.
'@param lastColumn Long. Right worksheet column of the block that was filtered.
'@return Variant. A two-dimensional array, 1 To total visible rows, 1 To columns.
Private Function VisibleBlock(ByVal visibleRng As Range, _
                              ByVal firstColumn As Long, _
                              ByVal lastColumn As Long) As Variant
    Dim areaRng As Range
    Dim blockValues As Variant
    Dim block() As Variant
    Dim runFirstRow() As Long
    Dim runRowCount() As Long
    Dim nbRuns As Long
    Dim totalRows As Long
    Dim nbColumns As Long
    Dim rowCursor As Long
    Dim rowCounter As Long
    Dim columnCounter As Long
    Dim counter As Long
    Dim areaFirstRow As Long
    Dim alreadyTaken As Boolean

    nbColumns = lastColumn - firstColumn + 1
    ReDim runFirstRow(1 To visibleRng.Areas.Count)
    ReDim runRowCount(1 To visibleRng.Areas.Count)

    For Each areaRng In visibleRng.Areas
        areaFirstRow = areaRng.Row
        alreadyTaken = False

        For counter = nbRuns To 1 Step -1
            If runFirstRow(counter) = areaFirstRow Then
                alreadyTaken = True
                Exit For
            End If
        Next

        If Not alreadyTaken Then
            nbRuns = nbRuns + 1
            runFirstRow(nbRuns) = areaFirstRow
            runRowCount(nbRuns) = areaRng.Rows.Count
            totalRows = totalRows + areaRng.Rows.Count
        End If
    Next

    ReDim block(1 To totalRows, 1 To nbColumns)

    For counter = 1 To nbRuns
        blockValues = AreaValues(SheetBlock(runFirstRow(counter), firstColumn, _
                                            runFirstRow(counter) + runRowCount(counter) - 1, _
                                            lastColumn))
        For rowCounter = 1 To runRowCount(counter)
            For columnCounter = 1 To nbColumns
                block(rowCursor + rowCounter, columnCounter) = blockValues(rowCounter, columnCounter)
            Next
        Next
        rowCursor = rowCursor + runRowCount(counter)
    Next

    VisibleBlock = block
End Function

'@label:store-block
'@sub-title Put a block into a BetterArray with the shape callers expect
'@details
'This reproduces exactly what BetterArray.FromExcelRange did for the block
'the two filter routines used to build, and callers depend on all four
'cases: one cell gives a one-element array, a single column gives a FLAT
'one-dimensional array, a single row gives a flat one-dimensional array,
'and anything wider gives a two-dimensional array (BA_MULTIDIMENSION,
'pinned by TestDataSheet). The array is handed over in ONE Items
'assignment. Pushing row by row would reallocate and copy the whole array
'on every element.
'@param target BetterArray. The array to fill. Its LowerBound is already 1.
'@param block Variant. A 1-based two-dimensional array.
Private Sub StoreBlock(ByVal target As BetterArray, ByRef block As Variant)
    Dim nbRows As Long
    Dim nbColumns As Long
    Dim flat() As Variant
    Dim counter As Long

    nbRows = UBound(block, 1)
    nbColumns = UBound(block, 2)

    If (nbRows = 1) And (nbColumns = 1) Then
        target.Push block(1, 1)
    ElseIf nbColumns = 1 Then
        ReDim flat(1 To nbRows)
        For counter = 1 To nbRows
            flat(counter) = block(counter, 1)
        Next
        target.Items = flat
    ElseIf nbRows = 1 Then
        ReDim flat(1 To nbColumns)
        For counter = 1 To nbColumns
            flat(counter) = block(1, counter)
        Next
        target.Items = flat
    Else
        target.Items = block
    End If
End Sub

'@label:filter-data
'@sub-title Filter data on a single column and return the result
'@details
'Applies an AutoFilter on varName with criteriaName and reads the visible
'rows of returnedColumnName straight into a BetterArray. The filter is
'removed on every path, including the failure paths, so the DataSheet is
'left as it was found.
'Nothing is written to the worksheet. The routine used to copy the visible
'rows into a scratch block two columns past the last data column, read it
'back and clear it; see VisibleBlock for why that block was sized wrong and
'what it left behind.
'@param varName String. Column to filter on.
'@param criteriaName String. Filter criteria to apply.
'@param returnedColumnName String. Column whose values to return ("__all__" for all columns).
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param matchCase Optional Boolean. When True, column lookup is case-sensitive. Defaults to True.
'@return BetterArray. The filtered values.
'@throws ProjectError.ElementNotFound When a column name is not found.
'@export
Public Function FilterData(ByVal varName As String, _
                            ByVal criteriaName As String, _
                            ByVal returnedColumnName As String, _
                            Optional ByVal includeHeaders As Boolean = False, _
                            Optional ByVal matchCase As Boolean = True) _
        As BetterArray

    Dim dataRng As Range
    Dim visibleRng As Range
    Dim returnedData As BetterArray
    Dim startRetColIndex As Long
    Dim endRetColIndex As Long
    Dim startRetRow As Long
    Dim endRow As Long
    Dim endCol As Long
    Dim filterField As Long
    Dim block As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    'The caller always gets an array back, never Nothing. LowerBound 1 and
    'the Clone() at the end are part of what the callers read.
    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    'One build answers three questions: the range to filter, the last row
    'and the last column. Reading the two ends off it saves asking the
    'worksheet for them a second time, and it is the same answer, because
    'DataRange("__all__") spans exactly DataStartRow to DataEndRow and
    'DataStartColumn to DataEndColumn.
    Set dataRng = DataRange(includeHeaders:=True)
    endCol = dataRng.Column + dataRng.Columns.Count - 1
    endRow = dataRng.Row + dataRng.Rows.Count - 1

    startRetRow = DataStartRow() + 1
    If includeHeaders Then startRetRow = (startRetRow - 1)

    If returnedColumnName = "__all__" Then
        startRetColIndex = DataStartColumn()
        endRetColIndex = endCol
    Else
        'A WORKSHEET column, because SheetBlock below works in worksheet
        'coordinates. This used to ask for inDataRange:=True, which answers
        'with a position relative to the data range; the two only agree
        'while the data starts at column 1, and not every DataSheet does.
        'An unknown column raises ElementNotFound on purpose, and it raises
        'before any filter is applied.
        startRetColIndex = ColumnIndex(returnedColumnName, shouldExist:=True, matchCase:=matchCase)
        endRetColIndex = startRetColIndex
    End If

    'AutoFilter counts its Field from the left edge of the filtered range,
    'so THIS one really is relative to the data range.
    filterField = ColumnIndex(varName, inDataRange:=True, shouldExist:=True, matchCase:=matchCase)

    'From here on the sheet carries a filter, so every path has to take it
    'off again.
    On Error GoTo FilterFail

    'A filter left over from an earlier call would hide rows before the
    'new one is applied, so the sheet is cleaned first.
    RemoveAutoFilter

    dataRng.AutoFilter Field:=filterField, Criteria1:=criteriaName

    'SpecialCells raises when every row is hidden. That raise IS how
    '"the filter matched nothing" is detected: visibleRng stays Nothing
    'and the caller gets the empty array built above, with no error.
    On Error Resume Next
    Set visibleRng = SheetBlock(startRetRow, startRetColIndex, endRow, endRetColIndex).SpecialCells(xlCellTypeVisible)
    On Error GoTo FilterFail

    If Not visibleRng Is Nothing Then
        block = VisibleBlock(visibleRng, startRetColIndex, endRetColIndex)
        StoreBlock returnedData, block
    End If

    'Leave the sheet as it was found: no filter.
    RemoveAutoFilter

    Set FilterData = returnedData.Clone()
    Exit Function

FilterFail:
    'The sheet is mid-filter with its rows hidden. Put it back before the
    'error goes up, or the caller is left looking at part of the data.
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    RemoveAutoFilter
    On Error GoTo 0
    Err.Raise errNumber, errSource, errDescription
End Function


'@label:filters-data
'@sub-title Filter data on multiple columns and return the result
'@details
'Applies multiple AutoFilters (one per entry in varData/criteriaData)
'and returns the visible values from each column in returnedColumnsData
'as a two-dimensional BetterArray. Exits with an empty array when varData
'and criteriaData have different lengths. The filters are removed on every
'path, including the failure paths.
'Nothing is written to the worksheet. Each returned column used to be
'pasted into its own scratch column past the last data column and the block
'measured back off the sheet, which had three ways to answer wrong: it
'gated BOTH the read and the clear on ONE cell being non-empty, so a blank
'first value (normal for a short-label column) dropped the whole block; it
'took the bottom edge from the first scratch column only; and it took the
'right edge from the header row, so it stopped short when the last column
'started blank and ran into the user's own content when anything already
'sat further right.
'@param varData BetterArray. Column names to filter on.
'@param criteriaData BetterArray. Criteria corresponding to each column.
'@param returnedColumnsData BetterArray. Columns whose values to return.
'@return BetterArray. A two-dimensional array of filtered values.
'@throws ProjectError.ElementNotFound When a column name is not found.
'@export
Public Function FiltersData(ByVal varData As BetterArray, _
                             ByVal criteriaData As BetterArray, _
                             ByVal returnedColumnsData As BetterArray) As BetterArray

    Dim dataRng As Range
    Dim visibleRng As Range
    Dim returnedData As BetterArray
    Dim returnedColumnIndex As Long
    Dim counter As Long
    Dim columnCursor As Long
    Dim rowCounter As Long
    Dim endRow As Long
    Dim nbColumns As Long
    Dim nbRows As Long
    Dim block As Variant
    Dim columnBlock As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    'The caller always gets an array back, on every exit path below.
    Set returnedData = New BetterArray
    returnedData.LowerBound = 1
    Set FiltersData = returnedData.Clone()

    'Mismatched counts are a caller mistake that is logged and answered
    'with an empty array. An unknown column name further down RAISES
    'instead. Two different answers to two different mistakes.
    If varData.Length <> criteriaData.Length Or varData.Length = 0 Then
        LogAndTrace "Unable to filter DataSheet: Column names and criteria tables have different length"
        Exit Function
    End If

    nbColumns = returnedColumnsData.Length
    If nbColumns = 0 Then Exit Function

    Set dataRng = DataRange(includeHeaders:=True)
    endRow = dataRng.Row + dataRng.Rows.Count - 1

    'From here on the sheet carries filters, so every path has to take them
    'off again.
    On Error GoTo FiltersFail

    'A filter left over from an earlier call would hide rows before the
    'new ones are applied.
    RemoveAutoFilter

    With dataRng
        For counter = varData.LowerBound To varData.UpperBound
            .AutoFilter Field:=ColumnIndex(varData.Item(counter), inDataRange:=True, shouldExist:=True), Criteria1:=criteriaData.Item(counter)
        Next
    End With

    For counter = returnedColumnsData.LowerBound To returnedColumnsData.UpperBound
        returnedColumnIndex = ColumnIndex(returnedColumnsData.Item(counter), shouldExist:=True)
        columnCursor = columnCursor + 1

        'SpecialCells raises when every row is hidden. That raise is how
        '"the filter matched nothing" is detected, so the guard here is
        'real error handling and not laziness.
        Set visibleRng = Nothing
        On Error Resume Next
        Set visibleRng = SheetBlock(DataStartRow() + 1, returnedColumnIndex, endRow, returnedColumnIndex).SpecialCells(xlCellTypeVisible)
        On Error GoTo FiltersFail

        'The same filters are active for every returned column, so when one
        'answers "nothing is visible" they all do.
        If visibleRng Is Nothing Then Exit For

        columnBlock = VisibleBlock(visibleRng, returnedColumnIndex, returnedColumnIndex)

        'The first column that answers sets the height of the result.
        If IsEmpty(block) Then
            nbRows = UBound(columnBlock, 1)
            ReDim block(1 To nbRows, 1 To nbColumns)
        End If

        For rowCounter = 1 To nbRows
            If rowCounter > UBound(columnBlock, 1) Then Exit For
            block(rowCounter, columnCursor) = columnBlock(rowCounter, 1)
        Next
    Next

    RemoveAutoFilter

    If Not IsEmpty(block) Then StoreBlock returnedData, block

    Set FiltersData = returnedData.Clone()
    Exit Function

FiltersFail:
    'The sheet is mid-filter with its rows hidden. Put it back before the
    'error goes up.
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    RemoveAutoFilter
    On Error GoTo 0
    Err.Raise errNumber, errSource, errDescription
End Function




'@label:import
'@sub-title Import data from another DataSheet or CustomTable
'@details
'Cleans this DataSheet, then iterates the source headers and copies
'matching column data. Formula columns are preserved. Applies format
'import for registered formatting columns after the data copy.
'@param importData Object. A DataSheet or CustomTable to import from.
'@param strictColumnSearch Optional Boolean. When True, uses case-sensitive column matching. Defaults to True.
'@export
Public Sub Import(ByVal importData As Object, Optional ByVal strictColumnSearch As Boolean = True)

    Dim imphRng As Range 'Import Header Range
    Dim sourceHeaders As Variant
    Dim headerValue As String
    Dim counter As Long
    Dim nbImpRows As Long
    Dim impDataRng As Range
    Dim rng As Range
    Dim destColumn As Long
    Dim firstDataRow As Long

    'You can immport either another datasheet or a custom table
    If Not IsImportable(importData, "Imports") Then Exit Sub

    'The source header row is read once as a block instead of one
    'worksheet read per column.
    Set imphRng = importData.HeaderRange()
    sourceHeaders = RowValues(imphRng)
    firstDataRow = DataStartRow() + 1

    'Clean the actual dictionary
    Clean

    For counter = 1 To UBound(sourceHeaders)
        'all the headers are in lowercase and are trimed
        headerValue = CellText(sourceHeaders(counter))

        'One lookup answers both questions that ColumnExists and DataRange
        'used to answer with two lookups and four worksheet searches.
        'Strict search is to look for whole column names and not part only
        'In some cases we need to match the column names case wise, and in other cases
        'especially when working with customtable, this is not mandatory
        destColumn = FindColumn(headerValue, True, strictColumnSearch)

        If destColumn > 0 Then
            'The source is always read whole-string and case-sensitive while
            'the destination follows the caller's flag. The two are
            'deliberately different, do not line them up.
            Set impDataRng = importData.DataRange(headerValue, strictSearch:=True, matchCase:=True)
            nbImpRows = impDataRng.Rows.Count

            'Add values if the range does not contains formulas. If the range
            'contains formulas, do nothing. Clean above already emptied every
            'column that passes this same test on this same cell, so there is
            'nothing left to clear here.
            Set rng = SheetBlock(firstDataRow, destColumn, firstDataRow + nbImpRows - 1, destColumn)
            If (Not rng.Cells(1, 1).HasFormula) Then rng.Value = impDataRng.Value
        End If
    Next

    'Calculate current DataRange after import
    DataRange.Calculate

    'importformats
    ImportFormat importData:=importData, matchColumnsCase:=strictColumnSearch
End Sub

'@label:import-format
'@sub-title Import column formatting from a source object
'@details
'Copies visual formatting (colours, font weight, italic) from the
'source for all columns registered via AddFormatsColumns. Logs a
'warning when the source and destination row counts differ.
'@param importData Object. A DataSheet or CustomTable to copy formatting from.
'@param matchColumnsCase Optional Boolean. When True, uses case-sensitive column matching. Defaults to True.
'@export
Public Sub ImportFormat(ByVal importData As Object, Optional ByVal matchColumnsCase As Boolean = True)

    Dim actRng As Range
    Dim counter As Long
    Dim impRng As Range
    Dim colName As String

    If Not IsImportable(importData, "Imports of formats") Then Exit Sub

    If Not HasFormatColumns() Then
        LogAndTrace "Trying to import formats without specifying formats columns. Maybe use AddFormatsColumns first?", checkingWarning
        Exit Sub
    End If

    'importing colunms that need formatting
    For counter = this.formatColumns.LowerBound To this.formatColumns.UpperBound

        colName = this.formatColumns.Item(counter)


        'The column might not exists in the exported DataBase
        If importData.ColumnExists(colName, strictSearch:=True, matchCase:=matchColumnsCase) Then

            Set actRng = DataRange(colName, matchCase:=matchColumnsCase)
            Set impRng = importData.DataRange(colName, matchCase:=matchColumnsCase)

            If (impRng.Rows.Count <> actRng.Rows.Count) Then
                'Add a warning on the the checkings here
                LogInfo "Importing Range " & colName & "formatting is truncated; source and target columns have different rows", checkingWarning
            End If

            'Start importing formats
            ApplyFormat actRng, impRng
        End If

    Next
End Sub

'@label:apply-format
'@sub-title Copy formatting from one Range to another
'@details
'Transfers interior colour, font colour, bold, and italic from each
'cell in impRng to the corresponding cell in actRng.
'This is not a straight format copy and xlPasteFormats would not do the
'same thing: only a colour that is not the default travels, so the
'destination keeps its own white background and black text, and borders,
'number formats and font names are left alone.
Private Sub ApplyFormat(ByVal actRng As Range, ByVal impRng As Range)

    Dim cellRng As Range
    Dim actCellRng As Range
    Dim rowCount As Long
    Dim counter As Long
    Dim interiorColor As Long
    Dim fontColor As Long
    Dim isBold As Boolean
    Dim isItalic As Boolean

    With actRng
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Font.Bold = False
        .Font.Italic = False
    End With

    'The count is read once and the loop is counted. The condition used to be
    'Do While cellRng.Row <= lastRow, which asks a cell its own row number on
    'every turn for a bound that cannot move while this runs.
    rowCount = impRng.Rows.Count
    If rowCount < 1 Then Exit Sub

    Set cellRng = impRng.Cells(1, 1)
    Set actCellRng = actRng.Cells(1, 1)

    For counter = 1 To rowCount

        'Read each source property once. Reading Interior and Font again for
        'the assignment cost a second round trip per cell.
        interiorColor = cellRng.Interior.Color
        fontColor = cellRng.Font.Color
        isBold = cellRng.Font.Bold
        isItalic = cellRng.Font.Italic

        If (interiorColor <> vbWhite) Then
            actCellRng.Interior.color = interiorColor
        End If

        If (fontColor <> vbBlack) Then
            actCellRng.Font.color = fontColor
        End If

        'Bold and Italic used to be written on every cell, unconditionally, and
        'they were the odd ones out: both colours have always been guarded. The
        'opening block above already sets both False across the whole
        'destination range, so writing False here puts back a value that is
        'already there, and on a dictionary almost every cell is neither bold
        'nor italic.
        'One difference this makes, and it is the only one: when impRng is
        'LONGER than actRng the walk runs past the destination range, and those
        'cells never got the opening block's False. They are outside the column
        'being written and ImportFormat already logs that case as truncated.
        If isBold Then actCellRng.Font.Bold = True
        If isItalic Then actCellRng.Font.Italic = True

        'Move to the next cell
        Set cellRng = cellRng.Offset(1)
        Set actCellRng = actCellRng.Offset(1)
    Next counter
End Sub

'@label:export
'@sub-title Export the DataSheet to a workbook
'@details
'Creates (or clears) a worksheet in the target workbook with the same
'name as this DataSheet, then writes all data. Optionally filters rows
'on a single column before export. The exported sheet is hidden by
'default. Applies formatting for registered columns and optionally
'exports hidden names.
'@param toWkb Workbook. Destination workbook.
'@param filteredVarName Optional String. Column used for filtering. Defaults to "__all__" (no filter).
'@param filteredCondition Optional String. Filter criteria. Defaults to "<>".
'@param Hide Optional Long. Worksheet visibility after export. Defaults to xlSheetHidden.
'@param includeNames Optional Boolean. When True, exports hidden names alongside data. Defaults to False.
'@param matchCase Optional Boolean. When True, the filter column is matched case-sensitively. Defaults to True.
'@export
Public Sub Export(ByVal toWkb As Workbook, Optional ByVal filteredVarName As String = "__all__", _
                   Optional ByVal filteredCondition As String = "<>", _
                   Optional ByVal Hide As Long = xlSheetHidden, _
                   Optional ByVal includeNames As Boolean = False, _
                   Optional ByVal matchCase As Boolean = True)

    Dim dataExportData As BetterArray
    Dim sh As Worksheet
    Dim dataExportWksh As Worksheet
    Dim sheetDataExist As Boolean
    Dim colName As String
    Dim actRng As Range
    Dim impRng As Range
    Dim exportedRng As Range
    Dim counter As Long
    Dim colIndex As Long
    Dim sourceColumn As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim store As HiddenNames

    'Trace on errors on export
    If toWkb Is Nothing Then
        LogAndTrace "Export can not be done, workbook not found", checkingError
        Exit Sub
    End If
    'Test if the sheet  exists in the workbook
    Set sh = Wksh()

    On Error Resume Next
    Set dataExportWksh = toWkb.Worksheets(sh.Name)
    sheetDataExist = (Not dataExportWksh Is Nothing)
    On Error GoTo 0

    'Trace on errors on exports
    If sheetDataExist Then
        LogInfo "There is already a worksheet of the same name in the workbook, the worksheet will be cleared completely", checkingWarning
        dataExportWksh.Cells.Clear
    Else
        toWkb.Worksheets.Add(after:=toWkb.Worksheets(toWkb.Worksheets.Count)).Name = sh.Name
        Set dataExportWksh = toWkb.Worksheets(sh.Name)
    End If

    Set dataExportData = New BetterArray

    If filteredVarName = "__all__" Then
        dataExportData.FromExcelRange DataRange(colName:="__all__", includeHeaders:=True)
    Else
        Set dataExportData = FilterData(filteredVarName, filteredCondition, "__all__", _
                                        includeHeaders:=True, matchCase:=matchCase)
    End If

    'Always export to first column, first row
    dataExportData.ToExcelRange dataExportWksh.Cells(1, 1)
    dataExportWksh.Visible = Hide

    'Change the font and other elements on exports.
    'Only the block that was just written is formatted. Formatting the whole
    'worksheet meant AutoFit ran over a million empty rows and the column
    'width was set on every column of the sheet. The .Rows(1) lines stay on
    'the whole row so the header band still runs across the sheet.
    'The target workbook belongs to the caller and its sheet may be
    'protected. Formatting is cosmetic, so a refusal here must not stop the
    'export: that is what the On Error covers, and nothing else.
    Set exportedRng = dataExportWksh.UsedRange
    With dataExportWksh
        On Error Resume Next
        exportedRng.EntireColumn.ColumnWidth = 25
        exportedRng.Font.Size = 9
        .Rows(1).Font.Size = 10
        .Rows(1).RowHeight = 20
        .Rows(1).Interior.Color = RGB(240, 240, 244)
        .Rows(1).Font.Bold = True
        exportedRng.Font.Name = "Consolas"
        exportedRng.WrapText = True
        exportedRng.HorizontalAlignment = xlHAlignLeft
        exportedRng.VerticalAlignment = xlVAlignCenter
        exportedRng.EntireRow.AutoFit
        On Error GoTo 0
    End With

    'Export the formats
    If HasFormatColumns() Then
        firstDataRow = DataStartRow() + 1
        'Nothing below writes to the source sheet, so the last row is read
        'once instead of once per format column.
        lastDataRow = DataEndRow()

        For counter = this.formatColumns.LowerBound To this.formatColumns.UpperBound
            colName = this.formatColumns.Item(counter)
            'One lookup, where ColumnIndex followed by DataRange each paid
            'for their own header read.
            sourceColumn = FindColumn(colName, True, True)

            If sourceColumn > 0 Then
                Set impRng = SheetBlock(firstDataRow, sourceColumn, lastDataRow, sourceColumn)
                colIndex = sourceColumn - DataStartColumn() + 1
                With dataExportWksh
                    Set actRng = .Range(.Cells(2, colIndex), .Cells(impRng.Rows.Count, colIndex))
                End With
                ApplyFormat actRng, impRng
            End If
        Next
    End If

    'Export HiddenNames
    If includeNames Then
        Set store = HiddenNames.Create(sh)
        If Not store Is Nothing Then
            On Error Resume Next
                store.ExportNames dataExportWksh
            On Error GoTo 0
        End If
    End If
End Sub

'@label:is-importable
'@sub-title Test whether an object can be imported from
'@details
'Import and ImportFormat both accept a DataSheet or a CustomTable and
'nothing else. The argument is declared As Object on purpose, so the test
'is on TypeName. Passing anything else logs and returns False; it does
'not raise, and a test drives that path with a Range.
'@param importData Object. The object handed to Import or ImportFormat.
'@param whatIsImported String. Word used in the message ("Imports", "Imports of formats").
'@return Boolean. True when the object can be read from.
Private Function IsImportable(ByVal importData As Object, ByVal whatIsImported As String) As Boolean
    If (TypeName(importData) = "DataSheet") Or (TypeName(importData) = "CustomTable") Then
        IsImportable = True
        Exit Function
    End If

    LogAndTrace whatIsImported & " are done only on DataSheet or CustomTable Objects", checkingWarning
End Function

'@label:has-format-columns
'@sub-title Test whether any column was registered for formatting
'@details
'TypeName is used and not Is Nothing because it answers both "never
'created" and "not a BetterArray" in one test.
'@return Boolean. True when at least one format column is registered.
Private Function HasFormatColumns() As Boolean
    If Not (TypeName(this.formatColumns) = "BetterArray") Then Exit Function
    HasFormatColumns = (this.formatColumns.Length > 0)
End Function

'@section Checkings
'===============================================================================
'@description Diagnostic logging and error handling.

'@label:log-info
'@sub-title Log a diagnostic message
'@details
'Adds a message to the internal Checking object, creating it on first
'use. The check counter provides a unique key for each entry.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingScope. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String
    checkName = "DataSheet " & Name() & " object trace"
    'First, instanciante the internalCheks Objects
    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create(checkName)

    'Add new informations to the checking object
    this.internalChecks.Add Cstr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:log-and-trace
'@sub-title Log a message and print the debug trace
'@details
'The "log, show the trace, then leave" tail was written out six times on
'the failure paths of FiltersData, Import, ImportFormat and Export.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingScope. Defaults to checkingNote.
Private Sub LogAndTrace(ByVal label As String, _
                        Optional ByVal scope As Byte = checkingNote)
    LogInfo label, scope
    ShowDebug
End Sub

'@label:has-checkings
'@prop-title Whether the DataSheet has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. An Checking instance, or Nothing.
'@export
Public Property Get CheckingValues() As Object
    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@label:show-debug
'@sub-title Print a debug trace to the Immediate window
Private Sub ShowDebug()
    Debug.Print "Check the informations of choice object " & Name()
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumb Long. ProjectError code.
'@param errorMessage String. Descriptive message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "DataSheet", "Error : " & Chr(34) & errorMessage
End Sub

'@label:seal
'@sub-title Seal the instance so setup setters can no longer be written
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against writes after sealing
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:throw-error-empty-sheet
'@sub-title Raise an error for a Nothing worksheet
'@throws ProjectError.ObjectNotInitialized Always.
Private Sub ThrowErrorEmptySheetGiven()
    ThrowError  ProjectError.ObjectNotInitialized, _
                "Unable to create an instance of DataSheet class. " & _
                "Given Worksheet object does not exists"
End Sub

'@label:throw-error-start-number
'@sub-title Raise an error for an invalid start row or column
'@param startNumber Optional Long. The invalid value. Defaults to -1.
'@param startLabel Optional String. Label identifying the parameter. Defaults to "start row".
'@throws ProjectError.InvalidArgument Always.
Private Sub ThrowErrorStartNumber(Optional ByVal startNumber As Long = -1, _
                                  Optional ByVal startLabel As String = "start row")
    ThrowError  ProjectError.InvalidArgument, _
                "Unable to create an instance of the DataSheet class. " & _
                " The " & startLabel & " should be >= 1. " & "Actual value : " & startNumber
End Sub

'@label:throw-error-unfound-column
'@sub-title Raise an error for a missing column
'@param colName String. The column name that was not found.
'@throws ProjectError.ElementNotFound Always.
Private Sub ThrowErrorUnFoundColumn(ByVal colName As String)
    ThrowError  ProjectError.ElementNotFound, _
               "Column " & colName & " does not exists" & " in DataSheet " & Name()
End Sub

'@label:add-formats-columns
'@sub-title Register columns for formatting import/export
'@details
'Adds the specified column names to the internal formatting list so that
'Import and Export preserve their visual formatting. Validates each
'column exists before registering it.
'@param matchColumnsCase Boolean. When True, uses case-sensitive column matching.
'@param resetColumns Boolean. When True, clears the existing list before adding.
'@param lsCols ParamArray Variant. Column names to register for formatting.
'@export
Public Sub AddFormatsColumns(ByVal matchColumnsCase As Boolean, ByVal resetColumns As Boolean, ParamArray lsCols() As Variant)
    Dim nbEls As Long
    Dim colName As String
    Dim newColName As String

    If this.formatColumns Is Nothing Then
        Set this.formatColumns = New BetterArray
    ElseIf resetColumns Then
        this.formatColumns.Clear
    End If

    If UBound(lsCols) = -1 Then Exit Sub
    For nbEls = 0 To UBound(lsCols)
        'Silently ignore non strings passed to the format columns
        If TypeName(lsCols(nbEls)) = "String" Then
            colName = lsCols(nbEls)
            'One lookup answers both "does it exist" and "how does the sheet
            'spell it". The name is stored as the sheet spells it, so Import
            'and Export can match it exactly afterwards.
            If FindColumn(colName, True, matchColumnsCase, newColName) > 0 Then
                If Not this.formatColumns.Includes(newColName) Then this.formatColumns.Push newColName
            End If
        End If
    Next
End Sub
