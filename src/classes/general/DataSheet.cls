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

'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Data manipulation in a worksheet")

Option Explicit
Implements IDataSheet

'Types of the class
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

'@section Instanciation
'===============================================================================

'@label:Create
'@prop-title Create a DataSheet Object
'
'@details
'A datasheet object is a worksheet with table like data structure in it. This class
'deals with basic data manipulation properties to the table-like structure in the 
'worksheet. It does not assume any prior structure in the worksheet, hence
'it is more low level than its CustomTable analog which relies on listObjects.
'DataSheets are identified by the start row and start columns. The class will 
'look for end rows and end columns dynamically.
'
'@param sh A worksheet
'@param startLn A Long, the starting line for the datasheet
'@param startCl A Long, the starting column for the datasheet
'@param forceEndRow Optional, Boolean. Whether to look for end rows on all the columns or Not. Default is False.
'@param objName Optional, As String, The name of the datasheet object. If empty, will default to the name of the worksheet. Default is empty.
'@return A DataSheet Object
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

'instanciate elements of the datasheet object
Public Property Get Self() As IDataSheet
    Set Self = Me
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Let DataStartRow(ByVal startLn As Long)
    this.startLn = startLn
End Property

Public Property Let DataStartColumn(ByVal startCl As Long)
    this.startCl = startCl
End Property

Public Property Let StrictEnd(ByVal forceEnd As Boolean)
    this.forceEnd = forceEnd
End Property

Public Property Let Name(ByVal objName As String)
    this.objName = objName
End Property


'@label:elements
'@section DataSheet Elements
'===============================================================================

'@label:wksh
'@prop-title Worksheet of the DataSheet Object
'@return A Worksheet Object
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:startrow
'@prop-title Start Row of the datasheet
'@return A Long
'@export
Public Property Get DataStartRow() As Long
    DataStartRow = this.startLn
End Property

'@label:startcol
'@prop-title Start Column of the datasheet
'@return A Long
'@export
Public Property Get DataStartColumn() As Long
    DataStartColumn = this.startCl
End Property

'@label:strictend
'@prop-title Check start row by looking on all the columns
'@return A Long
'@export
Public Property Get StrictEnd() As Boolean
    StrictEnd = this.forceEnd
End Property

'@label:name
'@prop-title Name of the Object
'@return A String, the name of the worksheet of the datasheet
'@export
Private Property Get Name() As String
    
    If this.objName = vbNullString Then 
        this.objName = Wksh.Name
    End If

    Name = this.objName
End Property

'@label:endrow
'@prop-title End Row of the DataSheet
'@return A Long
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

'@label:endcol
'@prop-title End Column of the DataSheet
'@return A Long
'@export
Private Property Get DataEndColumn() As Long
    With Wksh
        DataEndColumn = .Cells(DataStartRow(), .Columns.Count).End(xlToLeft).Column
    End With
End Property

'@label:headerrange
'@prop-title DataSheet header range
'@return A Long
'@export
Public Property Get HeaderRange() As Range
    'updating the header range
    With Wksh
        Set HeaderRange = .Range(.Cells(DataStartRow(), DataStartColumn()), .Cells(DataStartRow(), DataEndColumn()))
    End With
End Property


'@label:columnexists
'@prop-title Check if a column name exists
'@param colName A String, the column name where you want to check existence
'@param strictSearch Optional, Boolean. Whether to match the whole string in a cell or not. Default is True
'@param matchCase Optional, Boolean. Do case sensitive research. Default is True
'@return A Boolean
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


'@label:columnindex
'@prop-title Check if a column name exists
'@param colName A String, the column name where you want to check existence
'@param inDataRange Optional, Boolean. Return column index w.r. to the data range instead of the worksheet. Default is False.
'@param strictSearch Optional, Boolean. Whether to match the whole string in a cell or not. Default is True
'@param matchCase Optional, Boolean. Do case sensitive research. Default is True
'@return A Long
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


'@label:datarange
'@prop-title Range of a column
'
'@details
'Return a Data Range given a column name. When strictSearch is false, the 
'search for the column name is done using partial match. The first partial match
'found is the returned column. If the column names of the DataSheet are quite
'unique, the partial match returns unique values. In case some colunm names share
'some text, make sure to use strictSearch = True to avoid unexpected output.
'
'
'@param colName String, the name of the column to look for
'@param includeHeaders Optional, Boolean. Whether to return the colunmn plus the headers, default is False
'@param strictSearch Optional, Boolean. Look for partial values when searching the column name,
'@param matchCase Optional, Boolean. Do case sensitive search. Default is True.
'@return A Range object
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


'@prop-title For a given column name, return the corresponding column name match
'@param colName A String, the column name where you want to check existence
'@param strictSearch Optional, Boolean. Whether to match the whole string in a cell or not. Default is True
'@param matchCase Optional, Boolean. Do case sensitive research. Default is True
'@return A String
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

'@label:RenameColumn
'@sub-title Rename a header in the datasheet
'@param currentName String. Existing column header to rename.
'@param newName String. Replacement header text.
'@param strictSearch Optional Boolean. Perform strict header search. Default True.
'@param matchCase Optional Boolean. Respect header case. Default False.
Public Sub RenameColumn(ByVal currentName As String, _
                        ByVal newName As String, _
                        Optional ByVal strictSearch As Boolean = True, _
                        Optional ByVal matchCase As Boolean = False)

    Dim hRng As Range
    Dim headerIndex As Long

    If LenB(newName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "New column name cannot be empty"
    End If

    headerIndex = ColumnIndex(currentName, inDataRange:=False, _
                              shouldExist:=True, _
                              strictSearch:=strictSearch, _
                              matchCase:=matchCase)

    Set hRng = HeaderRange()
    If hRng Is Nothing Then Exit Sub
    If headerIndex <= 0 Or headerIndex > hRng.Columns.Count Then Exit Sub

    hRng.Cells(1, headerIndex).Value = newName
End Sub

'@label:operate
'@section Operate with Outside world
'===============================================================================

'@label:clean
'@sub-title Clean all the columns of the DataSheet except the formulas columns
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


'@label:filterdata
'@prop-title Filter some data using criteria on a column
'@details
'
'The function will filter the dictionary using basic excel filtering properties
'and return a BetterArray containing the results of the filtering process. It 
'will remove the filter once the operation is complete, keeping the datasheet untouched.
'
'@param varName A String, column where check condition
'@param criteriaName A String, the condition  to apply on the column
'@param returnedColumnName A String, column with values  to return
'@param includeHeaders Optional, Boolean. Whether to return values with headers. Default is False
'@return A BetterArray
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
    Dim EndRow As Long
    Dim retEndCol As Long

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeaders:=True) 'Plage of the dictionary data
    Set sh = Wksh()
    endCol = DataEndColumn()                     'End column and rows of the dictionary
    EndRow = DataEndRow()

    
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
        Set returnedRng = .Range(.Cells(startRetRow, startRetColIndex), .Cells(EndRow, endRetColIndex)).SpecialCells(xlCellTypeVisible)
        .AutoFilter.ShowAllData
        dataRng.AutoFilter.ShowAllData
        On Error GoTo 0

        If Not returnedRng Is Nothing Then
            returnedRng.Copy .Cells(DataStartRow(), endCol + 2)

            'end row of data to return
            EndRow = DataStartRow() + returnedRng.Cells.Count - 1 'cells.Count gives the number of cells, even in non contiguous ranges

            retEndCol = endCol + 2 + returnedRng.Columns.Count - 1

            'return column should be greather than endCol + 2. If it is not the case, there are empty lines causing that
            Set returnedRng = .Range(.Cells(DataStartRow(), endCol + 2), .Cells(EndRow, retEndCol))

            returnedData.FromExcelRange returnedRng, DetectLastRow:=False, _
                                        DetectLastColumn:=False
            returnedRng.Cells.Clear
        End If
    End With

    Set FilterData = returnedData.Clone()
End Function


'@label:filtersdata
'@prop-title Filter some data using criteria on a column on multiple conditions
'@details
'
'The function will filter the dictionary using basic excel filtering properties
'and return a BetterArray containing the results of the filtering process. It 
'will remove the filter once the operation is complete, keeping the datasheet untouched.
'All the column names search are done in a case sensitive way.
'
'@param varData A BetterArray of String, columns where to select values based on conditions
'@param crieteriaData A BetterArray of String, the conditions
'@param returnedColumnsData A BetterArray, columns which values you want to return
'@return A two dimension BetterArray
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

            If Not returnedRng Is Nothing Then returnedRng.Copy .Cells(DataStartRow(), endCol + Col)
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
'@sub-title Import a custom table or a data sheet into the current datasheet
'@param imporData A CustomTable or DataSheet Object to import
'@param strictColumnSearch Optional, Boolean. When importing the object, do case sensitive column search. Default is true
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
    
    'importformats
    ImportFormat importData:=importData, matchColumnsCase:=strictColumnSearch
    this.hasImport = True
End Sub

'@label:importformat
'@sub-title Import columns format into the current Datasheet
'@param imporData A CustomTable or DataSheet Object to import
'@param matchColumnsCase Optional, Boolean. When importing the object, do case sensitive column search. Default is true
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
        
        Set actRng = DataRange(colName, matchCase:=matchColumnsCase)

        'The column might not exists in the exported DataBase
        If importData.ColumnExists(colName, strictSearch:=True, matchCase:=matchColumnsCase) Then
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

'@sub-title Apply one Range format to another
'@param actRng A Range, where to apply the format
'@param impRng A Range, where to look for the format
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
'@sub-title Export the current datasheet to a workbook
'@details
'DataSheets are exported on a new worksheet with the same name as the one of the
'current DataSheet. If the worksheet already exists, nothing is done and a warning is raised.
'The exported datasheet always has startrow = 1 and startcolumn = 1, as it is exported on an empty worksheet.
'A Subset of the data sheet could be exported using filtering on one single column. By default all the rows are exported.

'@param Wkb A workbook, where to export the DataSheet
'@param filteredVarName Optional, string. Variable Name for filtering. By default no filtering is done
'@param filteredCondition Optional, string. Condition for filtering. By default select all lines
'@param Hide Optional, xlSheetVisibility. Visibility of the exported DataSheet worksheet. Default is xlSheetHidden
'@export
Private Sub Export(ByVal Wkb As Workbook, Optional ByVal filteredVarName As String = "__all__", _
                   Optional ByVal filteredCondition As String = "<>", Optional ByVal Hide As Long = xlSheetHidden)

    Dim dataExportData As BetterArray
    Dim sh As Worksheet
    Dim dataExportWksh As Worksheet
    Dim sheetDataExist As Boolean
    Dim colName As String
    Dim actRng As Range
    Dim impRng As Range
    Dim counter As Long
    Dim colIndex As Long

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
        LogInfo "Export can not be done, there is already a worksheet of the same name in the workbook", checkingError
        ShowDebug
        Exit Sub
    End If

    Wkb.Worksheets.Add(after:=Wkb.Worksheets(Wkb.Worksheets.Count)).Name = sh.Name
    Set dataExportData = New BetterArray
    Set dataExportWksh = Wkb.Worksheets(sh.Name)

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
End Sub

'@label:checkings
'@section Checkings
'===============================================================================

'@sub-title Add elements to the checking object for future notes
'@param Label String, the message to add.
'@param scope Byte. The scope of the checking object.
'@export
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

'@label:hascheckings
'@prop-title Check if the object has checkings
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checkingvalues
'@prop-title Get the cheking values
'@export
Private Property Get CheckingValues() As Object
    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property
    Set CheckingValues = internalChecks
End Property


Private Sub ShowDebug()
    Debug.Print "Check the informations of choice object " & Name()
End Sub

Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "DataSheet", "Error : " & Chr(34) & errorMessage
End Sub

Private Sub ThrowErrorEmptySheetGiven()
    ThrowError  ProjectError.ObjectNotInitialized, _
                "Unable to create an instance of DataSheet class. " & _
                "Given Worksheet object does not exists"
End Sub

Private Sub ThrowErrorStartNumber(Optional ByVal startNumber As Long = -1, _
                                  Optional ByVal startLabel As String = "start row")
    ThrowError  ProjectError.InvalidArgument, _
                "Unable to create an instance of the DataSheet class. " & _
                " The " & startLabel & " should be >= 1. " & "Actual value : " & startNumber
End Sub

Private Sub ThrowErrorUnFoundColumn(ByVal colName As String)
    ThrowError  ProjectError.ElementNotFound, _
               "Column " & colName & " does not exists" & " in DataSheet " & Name()
End Sub

'@jump:Create
'@section Interface
'===============================================================================

'@jump:elements
'DataSheet Elements ------------------------------------------------------------

'@jump:wksh
Private Property Get IDataSheet_Wksh() As Worksheet
    Set IDataSheet_Wksh = Wksh()
End Property

'@jump:startrow
Private Property Get IDataSheet_DataStartRow() As Long
    IDataSheet_DataStartRow = DataStartRow()
End Property

'@jump:startcol
Private Property Get IDataSheet_DataStartColumn() As Long
    IDataSheet_DataStartColumn = DataStartColumn()
End Property

'@jump:name
Private Property Get IDataSheet_Name() As String
    IDataSheet_Name = Name()
End Property

'@jump:endrow
Private Property Get IDataSheet_DataEndRow() As Long
    IDataSheet_DataEndRow = DataEndRow()
End Property

'@jump:endcol
Private Property Get IDataSheet_DataEndColumn() As Long
    IDataSheet_DataEndColumn = DataEndColumn()
End Property

'@jump:headerrange
Private Property Get IDataSheet_HeaderRange() As Range
    Set IDataSheet_HeaderRange = HeaderRange()
End Property

'@jump:columnindex
Private Function IDataSheet_ColumnIndex(ByVal colName As String, _
                                       Optional ByVal inDataRange As Boolean = False, _
                                       Optional shouldExist As Boolean = False, _ 
                                       Optional ByVal strictSearch As Boolean = True, _ 
                                       Optional ByVal matchCase As Boolean = True) As Long
    IDataSheet_ColumnIndex = ColumnIndex(colName, inDataRange, shouldExist, strictSearch, matchCase)
End Function

'@jump:columnexists
Private Function IDataSheet_ColumnExists(ByVal colName As String, _ 
                                         Optional ByVal strictSearch As Boolean = True, _ 
                                         Optional ByVal matchCase As Boolean = True) As Boolean

    IDataSheet_ColumnExists = ColumnExists(colName, strictSearch, matchCase)
End Function


'@jump:datarange
Private Property Get IDataSheet_DataRange(Optional ByVal colName As String = "__all__", _
                                         Optional ByVal includeHeaders As Boolean = False, _ 
                                         Optional ByVal strictSearch As Boolean = True, _ 
                                         Optional ByVal matchCase As Boolean = True) As Range
    Set IDataSheet_DataRange = DataRange(colName, includeHeaders, strictSearch, matchCase)
End Property

'@jump:operate
'Oprerate with outside world---------------------------------------------------

'@jump:clean
Private Sub IDataSheet_Clean()
    Clean
End Sub

'@label:addformatscolumns
'@sub-title Add Some columns for formatting import / export
'@param ParamArray A list of Strings, separated with "," the column name where you want to import/export formatting
'@param matchColumnsCase Optional, Boolean. Do case sensitive research. Default is True
'@export
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


'@jump:filterdata
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

'@jump:filtersdata
Private Function IDataSheet_FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, _ 
                                        ByVal returnedColumnsData As BetterArray) As BetterArray
    Set IDataSheet_FiltersData = FiltersData(varData, criteriaData, returnedColumnsData)
End Function

'@jump:import
Private Sub IDataSheet_Import(ByVal importedData As Object, Optional ByVal strictColumnSearch As Boolean = True)
    Import importedData, strictColumnSearch
End Sub

'@jump:importformat
Private Sub IDataSheet_ImportFormat(ByVal importData As Object, Optional ByVal matchColumnsCase As Boolean = True)
    ImportFormat importData, matchColumnsCase
End Sub

'@jump:RenameColumn
Private Sub IDataSheet_RenameColumn(ByVal currentName As String, _
                                    ByVal newName As String, _
                                    Optional ByVal strictSearch As Boolean = True, _
                                    Optional ByVal matchCase As Boolean = False)
    RenameColumn currentName, newName, strictSearch, matchCase
End Sub

'@jump:export
Private Sub IDataSheet_Export(ByVal toWkb As Workbook, _
                             Optional ByVal filteredVarName As String = "__all__", _
                             Optional ByVal filteredCondition As String = "<>", _
                             Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, filteredVarName, filteredCondition, Hide
End Sub

'@jump:checkings
'Errors and checkings----------------------------------------------------------

'@jump:hascheckings
Private Property Get IDataSheet_HasCheckings() As Boolean
    IDataSheet_HasCheckings = HasCheckings()
End Property

'@jump:checkingvalues
Private Property Get IDataSheet_CheckingValues() As Object
     Set IDataSheet_CheckingValues = CheckingValues()
End Property
