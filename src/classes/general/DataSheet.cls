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

Implements IDataSheet

'@class DataSheet
'@description
'Provides low-level data manipulation on a worksheet with a table-like
'structure identified by a starting row and column. Unlike CustomTable,
'DataSheet does not require a ListObject — it works directly with cell
'ranges. Supports column lookup, filtering (single and multi-column),
'import/export to other DataSheets or CustomTables, and formatting
'transfer for designated columns. Consumers interact through the
'IDataSheet interface.
'@depends Checking, BetterArray, HiddenNames
'@version 1.0 (2026-02-09)
'@author OutbreakTools
Private Type TDataSheet
    sh As Worksheet                        'Worksheet of the data
    startLn As Long                              'The line where the data starts
    startCl As Long                             'The column where the data starts (the first column is the column of the variable names.
    forceEnd As Boolean
    hasImport As Boolean
    hasChecks As Boolean
    objName As String
End Type

Private this As TDataSheet

'Formated columns list
Private FORMATCOLUMNS As BetterArray

'checking objects
Private checkCounter As Long
Private internalChecks As IChecking

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
'@return IDataSheet. A fully initialised DataSheet instance.
'@throws ProjectError.ObjectNotInitialized When sh is Nothing.
'@throws ProjectError.InvalidArgument When startLn or startCl is less than 1.
'@depends Checking, BetterArray
'@export
Public Function Create(ByVal sh As Worksheet, ByVal startLn As Long, _
                       ByVal startCl As Long, _
                       Optional ByVal forceEndRow As Boolean = False, _
                       Optional ByVal objName As String = vbNullString) As IDataSheet

    'Throw illegal errors before the creation of the dictionary

    If sh Is Nothing Then ThrowErrorEmptySheetGiven
    If startLn < 1 Then ThrowErrorStartNumber startLn, "start row"
    If startCl < 1 Then ThrowErrorStartNumber startCl, "start column"

    With New DataSheet
        Set .Wksh = sh
        .DataStartRow = startLn
        .DataStartColumn = startCl
        .StrictEnd = forceEndRow
        .Name = objName
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Current object instance
'@return IDataSheet. Reference to the current object.
Public Property Get Self() As IDataSheet
    Set Self = Me
End Property

'@label:wksh-set
'@prop-title Assign the worksheet reference
'@param sh Worksheet. The worksheet to assign.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

'@label:data-start-row-set
'@prop-title Assign the header row index
'@param startLn Long. The 1-based header row index.
Public Property Let DataStartRow(ByVal startLn As Long)
    this.startLn = startLn
End Property

'@label:data-start-column-set
'@prop-title Assign the starting column index
'@param startCl Long. The 1-based starting column index.
Public Property Let DataStartColumn(ByVal startCl As Long)
    this.startCl = startCl
End Property

'@label:strict-end-set
'@prop-title Assign the strict end row flag
'@param forceEnd Boolean. When True, scans all columns for the last row.
Public Property Let StrictEnd(ByVal forceEnd As Boolean)
    this.forceEnd = forceEnd
End Property

'@label:name-set
'@prop-title Assign the DataSheet name
'@param objName String. The name to assign.
Public Property Let Name(ByVal objName As String)
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
Private Property Get Name() As String
    
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
'@return Long. The 1-based last data row index.
'@export
Private Property Get DataEndRow() As Long

    Dim endrw As Long
    Dim hRng As Range
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim colEndRw As Long 'column end row, used in loop  

    Set sh = Wksh()

    With sh
        If StrictEnd() Then
            Set hRng = HeaderRange()
            Set cellRng = hRng.Cells(1, 1)
            endrw = hRng.Row
            Do While Not IsEmpty(cellRng)
                'For each column, search and update the last row. Columns with
                'formulas are not included in the search
                If (Not cellRng.HasFormula) Then
                    colEndRw = .Cells(.Rows.Count, cellRng.Column).End(xlUp).Row
                    If endrw < colEndRw Then endrw = colEndRw
                End If
                Set cellRng = cellRng.Offset(, 1)
            Loop
        Else
            endrw = .Cells(.Rows.Count, DataStartColumn()).End(xlUp).Row
        End If
    End With
    
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
Private Property Get DataEndColumn() As Long
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
    With Wksh
        Set HeaderRange = .Range(.Cells(DataStartRow(), DataStartColumn()), .Cells(DataStartRow(), DataEndColumn()))
    End With
End Property


'@label:column-exists
'@sub-title Check whether a column exists in the DataSheet
'@details
'Searches the header row for the specified column name using Range.Find.
'Returns False for empty column names.
'@param colName String. Column header to search for.
'@param strictSearch Optional Boolean. When True, uses exact whole-string matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return Boolean. True when the column is found.
'@export
Private Function ColumnExists(ByVal colName As String, _ 
                              Optional ByVal strictSearch As Boolean = True, _
                              Optional ByVal matchCase As Boolean = True) As Boolean
    Dim hRng As Range
    Dim Exist As Boolean
    Set hRng = HeaderRange()                     'Header Range of the dictionary
    If colName = vbNullString Then
        Exist = False
    Else
        Exist = Not (hRng.Find(What:=colName, _ 
                     lookAt:=IIF(strictSearch, xlWhole, xlPart), _ 
                     MatchCase:=matchCase) Is Nothing)
    End If
    ColumnExists = Exist
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
Private Function ColumnIndex(ByVal colName As String, _
                             Optional ByVal inDataRange As Boolean = False, _
                             Optional ByVal shouldExist As Boolean = False, _ 
                             Optional ByVal strictSearch As Boolean = True, _
                             Optional ByVal matchCase As Boolean = True) As Long
    Dim hRng As Range
    Dim colIndex As Long
    Set hRng = HeaderRange()

    If ColumnExists(colName, strictSearch, matchCase) Then
        colIndex = hRng.Find(What:=colName, _ 
                            lookAt:=IIF(strictSearch, xlWhole, xlPart), _ 
                            MatchCase:=matchCase).Column
        If inDataRange Then
            ColumnIndex = colIndex - DataStartColumn() + 1
        Else
            ColumnIndex = colIndex
        End If
    Else
        If shouldExist Then
            ThrowErrorUnFoundColumn colName
        Else
            'By default, unfound columns will have -1 as Index
            ColumnIndex = -1
        End If
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
    Dim sh As Worksheet
    Dim endCol As Long
    Dim customStartCol As Long
    Set sh = Wksh()

    With sh
        customStartRow = DataStartRow() + 1
        If includeHeaders Then customStartRow = DataStartRow()
        If colName = "__all__" Then
            customStartCol = DataStartColumn()
            endCol = DataEndColumn()
        Else
            customStartCol = ColumnIndex(colName, shouldExist:=True, strictSearch:=strictSearch, matchCase:=matchCase)
            endCol = customStartCol
        End If
        Set DataRange = .Range(.Cells(customStartRow, customStartCol), .Cells(DataEndRow(), endCol))
    End With
End Property


'@label:matched-column-name
'@sub-title Return the actual header text for a column name match
'@details
'Searches the header row and returns the exact text of the matched
'cell. Returns vbNullString when no match is found.
'@param colName String. Column name to search for.
'@param strictSearch Optional Boolean. When True, uses exact matching. Defaults to True.
'@param matchCase Optional Boolean. When True, matching is case-sensitive. Defaults to True.
'@return String. The matched header text, or vbNullString.
Private Function MatchedColumnName(ByVal colName As String, _
                                Optional ByVal strictSearch As Boolean = True, _
                                Optional ByVal matchCase As Boolean = True) As String
    Dim hRng As Range
    Dim colIndex As Long

    Set hRng = HeaderRange()

    On Error Resume Next
         colIndex = (hRng.Find(What:=colName, _ 
                     lookAt:=IIF(strictSearch, xlWhole, xlPart), _ 
                     MatchCase:=matchCase).Column - hRng.Column + 1  _
        )
    On Error GoTo 0

    If  (colIndex = 0) Then 
        MatchedColumnName = vbNullString
    Else
        MatchedColumnName = hRng.Cells(1, colIndex).Value
    End If
End Function

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

    Dim hRng As Range
    Dim headerIndex As Long

    If LenB(newName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "New column name cannot be empty"
    End If

    headerIndex = ColumnIndex(currentName, inDataRange:=True, _
                              shouldExist:=True, _
                              strictSearch:=strictSearch, _
                              matchCase:=matchCase)

    Set hRng = HeaderRange()
    If hRng Is Nothing Then Exit Sub
    If headerIndex <= 0 Or headerIndex > hRng.Columns.Count Then Exit Sub

    hRng.Cells(1, headerIndex).Value = newName
End Sub

'@section Operations
'===============================================================================
'@description Filtering, cleaning, import, and export operations.

'@label:clean
'@sub-title Clear all non-formula columns
'@details
'Iterates every column in the header row and clears the contents of
'those that do not contain formulas in their first data cell.
'@export
Private Sub Clean()
  Dim rng As Range
  Dim hRng As Range
  Dim headerValue As String
  Dim counter As Long

  Set hRng = HeaderRange()

  'Clean only columns without formulas that are not keys (key columns are kept as it).
  For counter = 1 To hRng.Columns.Count
    headerValue = hRng.Cells(1, counter).Value
    Set rng = DataRange(colName:=headerValue, strictSearch:=True, includeHeaders:=False)
     'On formula colunms, Or  no clearing is done, 
    'clear only "non formula" columns      
    If (Not rng.Cells(1, 1).HasFormula) Then rng.ClearContents
  Next
End Sub


'@label:filter-data
'@sub-title Filter data on a single column and return the result
'@details
'Applies an AutoFilter on varName with criteriaName, extracts the
'visible rows from returnedColumnName into a temporary output area,
'reads them into a BetterArray, then cleans up. The filter is removed
'after the operation, leaving the DataSheet untouched.
'@param varName String. Column to filter on.
'@param criteriaName String. Filter criteria to apply.
'@param returnedColumnName String. Column whose values to return ("__all__" for all columns).
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@param matchCase Optional Boolean. When True, column lookup is case-sensitive. Defaults to True.
'@return BetterArray. The filtered values.
'@export
Private Function FilterData(ByVal varName As String, _
                            ByVal criteriaName As String, _
                            ByVal returnedColumnName As String, _
                            Optional ByVal includeHeaders As Boolean = False, _
                            Optional ByVal matchCase As Boolean = True) _
        As BetterArray

    Dim dataRng As Range
    Dim returnedRng As Range
    Dim startRetColIndex As Long
    Dim startRetRow As Long
    Dim endRetColIndex As Long
    Dim returnedData As BetterArray
    Dim sh As Worksheet
    Dim endCol As Long
    Dim outRng As Range
    Dim endRow As Long

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeaders:=True) 'Plage of the dictionary data
    Set sh = Wksh()
    endCol = DataEndColumn()                     'End column and rows of the dictionary
    endRow = DataEndRow()

    
    startRetRow = DataStartRow() + 1 
    If includeHeaders Then startRetRow = (startRetRow - 1)

    If returnedColumnName = "__all__" Then
        startRetColIndex = DataStartColumn()
        endRetColIndex = endCol
    Else
        startRetColIndex = ColumnIndex(returnedColumnName, inDataRange:=True, shouldExist:=True, matchCase:=matchCase) 'Which index to return, the startcolumn is returned if the returned colname is not found
        endRetColIndex = startRetColIndex
    End If

    With sh
        If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData

        dataRng.AutoFilter Field:=ColumnIndex(varName, inDataRange:=True, shouldExist:=True, matchCase:=matchCase), Criteria1:=criteriaName

        On Error Resume Next
        Set returnedRng = .Range(.Cells(startRetRow, startRetColIndex), .Cells(endRow, endRetColIndex)).SpecialCells(xlCellTypeVisible)
        .AutoFilter.ShowAllData
        dataRng.AutoFilter.ShowAllData
        On Error GoTo 0

        If Not returnedRng Is Nothing Then
            Set outRng = .Cells(DataStartRow(), endCol + 2)
            Set outRng = outRng.Resize(returnedRng.Rows.Count, returnedRng.Columns.Count)
            returnedRng.Copy
            outRng.PasteSpecial xlPasteValues

            returnedData.FromExcelRange outRng, DetectLastRow:=False, _
                                        DetectLastColumn:=False
            outRng.Cells.ClearContents
        End If
    End With

    Set FilterData = returnedData.Clone()
End Function


'@label:filters-data
'@sub-title Filter data on multiple columns and return the result
'@details
'Applies multiple AutoFilters (one per entry in varData/criteriaData)
'and returns the visible values from each column in returnedColumnsData
'as a two-dimensional BetterArray. Exits when varData and criteriaData
'have different lengths.
'@param varData BetterArray. Column names to filter on.
'@param criteriaData BetterArray. Criteria corresponding to each column.
'@param returnedColumnsData BetterArray. Columns whose values to return.
'@return BetterArray. A two-dimensional array of filtered values.
'@export
Private Function FiltersData(ByVal varData As BetterArray, _
                             ByVal criteriaData As BetterArray, _
                             ByVal returnedColumnsData As BetterArray) As BetterArray

    Dim dataRng As Range                         'Range of the data
    Dim sh As Worksheet
    Dim returnedRng As Range
    Dim returnedColumnIndex As Long
    Dim returnedData As BetterArray
    Dim counter As Long
    Dim Col As Long
    Dim EndRow As Long
    Dim endCol As Long
    Dim retEndRow As Long
    Dim retEndCol As Long
    Dim outRng As Range

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeaders:=True)
    Set FiltersData = returnedData.Clone()

    'On Error GoTo ErrFilter
    If varData.Length <> criteriaData.Length Or varData.Length = 0 Then 
        LogInfo "Unable to filter DataSheet: Column names and criteria tables have different length"
        ShowDebug
        Exit Function
    End If

    'returned Range
    Col = 2                                      'Iterator for columns to extract
    Set sh = Wksh()

    With sh

        If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
        EndRow = DataEndRow()
        endCol = DataEndColumn()

        With dataRng
            For counter = varData.LowerBound To varData.UpperBound
                .AutoFilter Field:=ColumnIndex(varData.Item(counter), inDataRange:=True, shouldExist:=True), Criteria1:=criteriaData.Item(counter)
            Next
        End With

        For counter = returnedColumnsData.LowerBound To returnedColumnsData.UpperBound
            returnedColumnIndex = ColumnIndex(returnedColumnsData.Item(counter), shouldExist:=True)

            On Error Resume Next
            Set returnedRng = .Range(.Cells(DataStartRow() + 1, returnedColumnIndex), .Cells(EndRow, returnedColumnIndex)).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            If Not returnedRng Is Nothing Then
                Set outRng = .Cells(DataStartRow(), endCol + Col).Resize(returnedRng.Rows.Count, 1)
                returnedRng.Copy
                outRng.PasteSpecial xlPasteValues
            End If
            Col = Col + 1
        Next

        On Error Resume Next
        .AutoFilter.ShowAllData
        dataRng.AutoFilter.ShowAllData
        On Error GoTo 0

        'end row and end columns of data to return
        retEndRow = .Cells(.Rows.Count, endCol + 2).End(xlUp).Row
        retEndCol = .Cells(DataStartRow(), .Columns.Count).End(xlToLeft).Column

        If Not IsEmpty(.Cells(DataStartRow(), endCol + 2)) Then
            Set returnedRng = .Range(.Cells(DataStartRow(), endCol + 2), .Cells(retEndRow, retEndCol))
            'returned data is the data returned by the filter
            returnedData.FromExcelRange returnedRng, DetectLastRow:=False, DetectLastColumn:=False
            returnedRng.Cells.Clear
        End If
    End With
    Set FiltersData = returnedData.Clone()
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
Private Sub Import(ByVal importData As Object, Optional ByVal strictColumnSearch As Boolean = True)

    Dim imphRng As Range 'Import Header Range
    Dim headerValue As String
    Dim counter As Long
    Dim nbImpRows As Long
    Dim impDataRng As Range
    Dim currDataRng As Range
    Dim rng As Range
    Dim sh As Worksheet

    'You can immport either another datasheet or a custom table
    If (TypeName(importData) <> "DataSheet") And _
       (TypeName(importData) <> "CustomTable") Then 
        Loginfo "Imports are done only on DataSheet or CustomTable Objects", checkingWarning
        ShowDebug
        Exit Sub
    End If

    Set imphRng = importData.HeaderRange()
    Set sh = Wksh()

    'Clean the actual dictionary
    Clean

    For counter = 1 To imphRng.Columns.Count
        'all the headers are in lowercase and are trimed
        headerValue = imphRng.Cells(1, counter).Value

        If ColumnExists(headerValue, matchCase:=strictColumnSearch) Then
            'This data Range does not includes colnames. 
            'Strict search is to look for whole column names and not part only
            'In some cases we need to match the column names case wise, and in other cases
            'especially when working with customtable, this is not mandatory
            Set impDataRng = importData.DataRange(headerValue, strictSearch:=True, matchCase:=True)
            Set currDataRng = DataRange(headerValue, strictSearch:=True, matchCase:=strictColumnSearch)
            nbImpRows = impDataRng.Rows.Count
            'Add values if the range does not contains formulas. If the range contains formulas, do nothing
            Set rng = sh.Range(currDataRng.Cells(1, 1), currDataRng.Cells(nbImpRows, 1))
            If (Not rng.Cells(1, 1).HasFormula) Then
                currDataRng.Cells.ClearContents
                rng.Value = impDataRng.Value
            End If
        End If
    Next

    'Calculate current DataRange after import
    DataRange.Calculate
    
    'importformats
    ImportFormat importData:=importData, matchColumnsCase:=strictColumnSearch
    this.hasImport = True
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
Private Sub ImportFormat(ByVal importData As Object, Optional ByVal matchColumnsCase As Boolean = True)
  
    Dim actRng As Range
    Dim counter As Long
    Dim impRng As Range
    Dim colName As String


    If (TypeName(importData) <> "DataSheet") And (TypeName(importData) <> "CustomTable") Then
        LogInfo "Imports of formats are done only on DataSheet or CustomTable Objects", checkingWarning
        ShowDebug
        Exit Sub
    End If

    If (Not TypeName(FORMATCOLUMNS) = "BetterArray") Then
        LogInfo "Trying to import formats without specifying formats columns. Maybe use AddFormatsColumns first?", checkingWarning
        ShowDebug
        Exit Sub
    End If
    
    If (FORMATCOLUMNS.Length = 0) Then 
        LogInfo "Trying to import formats without specifying formats columns. Maybe use AddFormatsColumns first?", checkingWarning
        ShowDebug
        Exit Sub
    End If

    'importing colunms that need formatting
    For counter = FORMATCOLUMNS.LowerBound To FORMATCOLUMNS.UpperBound
        
        colName = FORMATCOLUMNS.Item(counter)
        

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
'@param actRng Range. Destination range to format.
'@param impRng Range. Source range with formatting to copy.
Private Sub ApplyFormat(ByVal actRng As Range, ByVal impRng As Range)

    Dim cellRng As Range
    Dim actCellRng As Range

    actRng.ClearFormats
    Set cellRng = impRng.Cells(1, 1)
    Set actCellRng = actRng.Cells(1, 1)

    Do While (cellRng.Row <= impRng.Cells(impRng.Rows.Count, 1).Row)

        If (cellRng.Interior.Color <> vbWhite) Then 
            actCellRng.Interior.color = cellRng.Interior.color
        End If

        If (cellRng.Font.Color <> vbBlack) Then 
            actCellRng.Font.color = cellRng.Font.color
        End If

        actCellRng.Font.Bold = cellRng.Font.Bold
        actCellRng.Font.Italic = cellRng.Font.Italic

        'Move to the next cell
        Set cellRng = cellRng.Offset(1)
        Set actCellRng = actCellRng.Offset(1)
    Loop
End Sub

'@label:export
'@sub-title Export the DataSheet to a workbook
'@details
'Creates (or clears) a worksheet in the target workbook with the same
'name as this DataSheet, then writes all data. Optionally filters rows
'on a single column before export. The exported sheet is hidden by
'default. Applies formatting for registered columns and optionally
'exports hidden names.
'@param Wkb Workbook. Destination workbook.
'@param filteredVarName Optional String. Column used for filtering. Defaults to "__all__" (no filter).
'@param filteredCondition Optional String. Filter criteria. Defaults to "<>".
'@param Hide Optional Long. Worksheet visibility after export. Defaults to xlSheetHidden.
'@param includeNames Optional Boolean. When True, exports hidden names alongside data. Defaults to False.
'@export
Private Sub Export(ByVal Wkb As Workbook, Optional ByVal filteredVarName As String = "__all__", _
                   Optional ByVal filteredCondition As String = "<>", _ 
                   Optional ByVal Hide As Long = xlSheetHidden, _ 
                   Optional ByVal includeNames As Boolean = False)

    Dim dataExportData As BetterArray
    Dim sh As Worksheet
    Dim dataExportWksh As Worksheet
    Dim sheetDataExist As Boolean
    Dim colName As String
    Dim actRng As Range
    Dim impRng As Range
    Dim counter As Long
    Dim colIndex As Long
    Dim store As IHiddenNames

    'Trace on errors on export
    If Wkb Is Nothing Then
        LogInfo "Export can not be done, workbook not found", checkingError
        ShowDebug
        Exit Sub
    End If
    'Test if the sheet  exists in the workbook
    Set sh = Wksh()

    On Error Resume Next
    Set dataExportWksh = Wkb.Worksheets(sh.Name)
    sheetDataExist = (Not dataExportWksh Is Nothing)
    On Error GoTo 0

    'Trace on errors on exports
    If sheetDataExist Then 
        LogInfo "There is already a worksheet of the same name in the workbook, the worksheet will be cleared completely", checkingWarning
        dataExportWksh.Cells.Clear
    Else
        Wkb.Worksheets.Add(after:=Wkb.Worksheets(Wkb.Worksheets.Count)).Name = sh.Name
        Set dataExportWksh = Wkb.Worksheets(sh.Name)
    End If

    Set dataExportData = New BetterArray
    
    If filteredVarName = "__all__" Then
        dataExportData.FromExcelRange DataRange(colName:="__all__", includeHeaders:=True)
    Else
        Set dataExportData = FilterData(filteredVarName, filteredCondition, "__all__", includeHeaders:=True)
    End If

    'Always export to first column, first row
    dataExportData.ToExcelRange dataExportWksh.Cells(1, 1)
    dataExportWksh.Visible = Hide

    'Change the font and other elements on exports
    With dataExportWksh
        On Error Resume Next
        .Cells.EntireColumn.ColumnWidth = 25
        .Cells.Font.Size = 9
        .Rows(1).Font.Size = 10
        .Rows(1).RowHeight = 20
        .Rows(1).Interior.Color = RGB(240, 240, 244)
        .Rows(1).Font.Bold = True
        .Cells.Font.Name = "Consolas"
        .Cells.WrapText = True
        .Cells.HorizontalAlignment = xlHAlignLeft
        .Cells.VerticalAlignment = xlVAlignCenter
        .Cells.EntireRow.AutoFit
        On Error GoTo 0
    End With

    'Export the formats
    If TypeName(FORMATCOLUMNS) = "BetterArray" Then
        If (FORMATCOLUMNS.Length <> 0) Then 
            For counter = FORMATCOLUMNS.LowerBound To FORMATCOLUMNS.UpperBound
                colName = FORMATCOLUMNS.Item(counter)
                colIndex = ColumnIndex(colName, inDataRange := True)
                Set impRng = DataRange(colName)
                With dataExportWksh
                    Set actRng = .Range(.Cells(2, colIndex), .Cells(impRng.Rows.Count, colIndex))
                End With
                ApplyFormat actRng, impRng
            Next
        End If
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

'@section Checkings
'===============================================================================
'@description Diagnostic logging and error handling.

'@label:log-info
'@sub-title Log a diagnostic message
'@details
'Adds a message to the internal IChecking object, creating it on first
'use. The checkCounter provides a unique key for each entry.
'@param label String. The message to record.
'@param scope Optional Byte. Severity level from CheckingScope. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _ 
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String
    checkName = "CustomTable " & Name() & " object trace"
    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    
    'Add new informations to the checking object
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Whether the DataSheet has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. An IChecking instance, or Nothing.
'@export
Private Property Get CheckingValues() As Object
    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property
    Set CheckingValues = internalChecks
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

'@section Interface
'===============================================================================
'@description Delegated members satisfying the IDataSheet contract.

Private Property Get IDataSheet_Wksh() As Worksheet
    Set IDataSheet_Wksh = Wksh()
End Property

Private Property Get IDataSheet_DataStartRow() As Long
    IDataSheet_DataStartRow = DataStartRow()
End Property

Private Property Get IDataSheet_DataStartColumn() As Long
    IDataSheet_DataStartColumn = DataStartColumn()
End Property

Private Property Get IDataSheet_Name() As String
    IDataSheet_Name = Name()
End Property

Private Property Get IDataSheet_DataEndRow() As Long
    IDataSheet_DataEndRow = DataEndRow()
End Property

Private Property Get IDataSheet_DataEndColumn() As Long
    IDataSheet_DataEndColumn = DataEndColumn()
End Property

Private Property Get IDataSheet_HeaderRange() As Range
    Set IDataSheet_HeaderRange = HeaderRange()
End Property

Private Function IDataSheet_ColumnIndex(ByVal colName As String, _
                                       Optional ByVal inDataRange As Boolean = False, _
                                       Optional shouldExist As Boolean = False, _ 
                                       Optional ByVal strictSearch As Boolean = True, _ 
                                       Optional ByVal matchCase As Boolean = True) As Long
    IDataSheet_ColumnIndex = ColumnIndex(colName, inDataRange, shouldExist, strictSearch, matchCase)
End Function

Private Function IDataSheet_ColumnExists(ByVal colName As String, _ 
                                         Optional ByVal strictSearch As Boolean = True, _ 
                                         Optional ByVal matchCase As Boolean = True) As Boolean

    IDataSheet_ColumnExists = ColumnExists(colName, strictSearch, matchCase)
End Function


Private Property Get IDataSheet_DataRange(Optional ByVal colName As String = "__all__", _
                                         Optional ByVal includeHeaders As Boolean = False, _ 
                                         Optional ByVal strictSearch As Boolean = True, _ 
                                         Optional ByVal matchCase As Boolean = True) As Range
    Set IDataSheet_DataRange = DataRange(colName, includeHeaders, strictSearch, matchCase)
End Property

Private Sub IDataSheet_Clean()
    Clean
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
Private Sub IDataSheet_AddFormatsColumns(ByVal matchColumnsCase As Boolean, ByVal resetColumns As Boolean, ParamArray lsCols() As Variant)
    Dim nbEls As Long
    Dim colName As String
    Dim newColName As String

    If FORMATCOLUMNS Is Nothing Then
        Set FORMATCOLUMNS = New BetterArray
    ElseIf resetColumns Then
        FORMATCOLUMNS.Clear
    End If

    If UBound(lsCols) = -1 Then Exit Sub
    For nbEls = 0 To UBound(lsCols)
        'Silently ignore non strings passed to the format columns
        If TypeName(lsCols(nbEls)) = "String" Then
            colName = lsCols(nbEls)
            If ColumnExists(colName, matchCase:=matchColumnsCase) Then
                newColName = MatchedColumnName(colName:=colName, matchCase:=matchColumnsCase)
                If Not FORMATCOLUMNS.Includes(newColName) Then FORMATCOLUMNS.Push newColName
            End If
        End If
    Next
End Sub


Private Function IDataSheet_FilterData(ByVal varName As String, _
                                      ByVal criteriaName As String, _
                                      ByVal returnedColumnName As String, _
                                      Optional ByVal includeHeaders As Boolean = False, _
                                      Optional ByVal matchCase As Boolean = True) As BetterArray
    Set IDataSheet_FilterData = FilterData(varName:=varName, _
                                           criteriaName:=criteriaName, _
                                           returnedColumnName:=returnedColumnName, _
                                           includeHeaders:=includeHeaders, _ 
                                           matchCase:=matchCase)
End Function

Private Function IDataSheet_FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, _ 
                                        ByVal returnedColumnsData As BetterArray) As BetterArray
    Set IDataSheet_FiltersData = FiltersData(varData, criteriaData, returnedColumnsData)
End Function

Private Sub IDataSheet_Import(ByVal importedData As Object, Optional ByVal strictColumnSearch As Boolean = True)
    Import importedData, strictColumnSearch
End Sub

Private Sub IDataSheet_ImportFormat(ByVal importData As Object, Optional ByVal matchColumnsCase As Boolean = True)
    ImportFormat importData, matchColumnsCase
End Sub

Private Sub IDataSheet_RenameColumn(ByVal currentName As String, _
                                    ByVal newName As String, _
                                    Optional ByVal strictSearch As Boolean = True, _
                                    Optional ByVal matchCase As Boolean = False)
    RenameColumn currentName, newName, strictSearch, matchCase
End Sub

Private Sub IDataSheet_Export(ByVal toWkb As Workbook, _
                             Optional ByVal filteredVarName As String = "__all__", _
                             Optional ByVal filteredCondition As String = "<>", _
                             Optional ByVal Hide As Long = xlSheetHidden, _
                             Optional ByVal includeNames As Boolean = False)
    Export toWkb, filteredVarName, filteredCondition, Hide, includeNames
End Sub

Private Property Get IDataSheet_HasCheckings() As Boolean
    IDataSheet_HasCheckings = HasCheckings()
End Property

Private Property Get IDataSheet_CheckingValues() As Object
     Set IDataSheet_CheckingValues = CheckingValues()
End Property
