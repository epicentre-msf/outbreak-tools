VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Work with unstructured data in a worksheet"

'@PredeclaredId
'@IgnoreModule SuspiciousPredeclaredInstanceAccess, IIfSideEffect, ExcelMemberMayReturnNothing, ReadOnlyPropertyAssignment
'@Folder("General")
'@ModuleDescription("Work with unstructured data in a worksheet")

Option Explicit
Implements IDataSheet

'Types of the class
Private Type TDataSheet
    dataWksh As Worksheet                        'Worksheet of the data
    startLn As Long                              'The line where the data starts
    STARTCOL As Long                             'The column where the data starts (the first column is the column of the variable names.
End Type

Private this As TDataSheet

Private Const CLASSNAME As String = "DataSheet"

Public Function Create(ByVal dataWksh As Worksheet, ByVal dataStartRow As Long, _
                       ByVal dataStartColumn As Long) As IDataSheet

    'Throw illegal errors before the creation of the dataionary

    If dataWksh Is Nothing Then ThrowErrorEmptySheetGiven
    If dataStartRow < 1 Then ThrowErrorStartNumber dataStartRow, " Start Row "
    If dataStartColumn < 1 Then ThrowErrorStartNumber dataStartColumn, " Start Column "

    With New DataSheet
        Set .Wksh = dataWksh
        .StartRow = dataStartRow
        .StartColumn = dataStartColumn
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IDataSheet
    Set Self = Me
End Property

'Worksheet
Public Property Get Wksh() As Worksheet
    Set Wksh = this.dataWksh
End Property

Friend Property Set Wksh(ByVal dataWksh As Worksheet)
    Set this.dataWksh = dataWksh
End Property

'StartRow
Public Property Get StartRow() As Long
    StartRow = this.startLn
End Property

Friend Property Let StartRow(ByVal dataStartRow As Long)
    this.startLn = dataStartRow
End Property

'StartColumn
Public Property Get StartColumn() As Long
    StartColumn = this.STARTCOL
End Property

Friend Property Let StartColumn(ByVal dataStartColumn As Long)
    this.STARTCOL = dataStartColumn
End Property

'End row and columns of the dictionary
Private Property Get DataEndRow() As Long
    With Wksh
        DataEndRow = .Cells(.Rows.Count, StartColumn).End(xlUp).Row
    End With
End Property

Private Property Get DataEndColumn() As Long
    With Wksh
        DataEndColumn = .Cells(StartRow, .Columns.Count).End(xlToLeft).Column
    End With
End Property

'Data header range
Public Property Get HeaderRange() As Range
    'updating the header range
    With Wksh
        Set HeaderRange = .Range(.Cells(StartRow, StartColumn), .Cells(StartRow, DataEndColumn))
    End With
End Property

'Data data Range
Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False) As Range
    Dim customStartRow As Long
    Dim dataWksh As Worksheet
    Dim endCol As Long
    Dim customStartCol As Long
    Set dataWksh = Wksh()

    With dataWksh
        customStartRow = IIf(includeHeaders, StartRow, StartRow + 1)
        If colName = "__all__" Then
            customStartCol = StartColumn
            endCol = DataEndColumn
        Else
            customStartCol = ColumnIndex(colName, shouldExist:=True)
            endCol = customStartCol
        End If
        Set DataRange = .Range(.Cells(customStartRow, customStartCol), .Cells(DataEndRow, endCol))
    End With
End Property

'Methods -----------------------------------------------------------------------

'Check if colname exists
Private Function ColumnExists(ByVal colName As String) As Boolean
    Dim hRng As Range
    Dim exist As Boolean
    Set hRng = HeaderRange()                     'Header Range of the dictionary
    If colName = vbNullString Then
        exist = False
    Else
        exist = Not (hRng.Find(What:=colName, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
    End If
    ColumnExists = exist
End Function

'Column Index within the data range of the dictionary
Private Function ColumnIndex(ByVal colName As String, _
                             Optional ByVal inDataRange As Boolean = False, _
                             Optional ByVal shouldExist As Boolean = False) As Long
    Dim hRng As Range
    Dim colIndex As Long
    Set hRng = HeaderRange()

    If ColumnExists(colName) Then
        colIndex = hRng.Find(What:=colName, lookAt:=xlWhole, MatchCase:=True).Column
        ColumnIndex = IIf(inDataRange, colIndex - StartColumn + 1, colIndex)
    Else
        If shouldExist Then
            ThrowErrorUnFoundColumn colName
        Else
            'By default, unfound columns will have -1 as Index
            ColumnIndex = -1
        End If
    End If
End Function

'Filtered Data
Private Function FilterData(ByVal varName As String, _
                            ByVal criteriaName As String, _
                            ByVal returnedColumnName As String, _
                            Optional ByVal includeHeaders As Boolean = False) _
        As BetterArray

    Dim dataRng As Range
    Dim returnedRng As Range
    Dim startRetColIndex As Long
    Dim startRetRow As Long
    Dim endRetColIndex As Long
    Dim returnedData As BetterArray
    Dim dataWksh As Worksheet
    Dim endCol As Long
    Dim EndRow As Long
    Dim retEndCol As Long

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeaders:=True) 'Plage of the dictionary data
    Set dataWksh = Wksh()
    endCol = DataEndColumn()                     'End column and rows of the dictionary
    EndRow = DataEndRow()

    startRetRow = IIf(includeHeaders, StartRow, StartRow + 1)

    If returnedColumnName = "__all__" Then
        startRetColIndex = StartColumn
        endRetColIndex = endCol
    Else
        startRetColIndex = ColumnIndex(returnedColumnName, shouldExist:=True) 'Which index to return, the startcolumn is returned if the returned colname is not found
        endRetColIndex = startRetColIndex
    End If

    With dataWksh
        If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData

        With dataRng
            .AutoFilter Field:=ColumnIndex(varName, inDataRange:=True, shouldExist:=True), Criteria1:=criteriaName
        End With

        On Error Resume Next
        Set returnedRng = .Range(.Cells(startRetRow, startRetColIndex), .Cells(EndRow, endRetColIndex)).SpecialCells(xlCellTypeVisible)
        .AutoFilter.ShowAllData
        On Error GoTo 0

        If Not returnedRng Is Nothing Then
            returnedRng.Copy .Cells(StartRow, endCol + 2)

            'end row of data to return
            EndRow = StartRow + returnedRng.Cells.Count - 1 'cells.Count gives the number of cells, even in non contiguous ranges

            retEndCol = endCol + 2 + returnedRng.Columns.Count - 1

            'return column should be greather than endCol + 2. If it is not the case, there are empty lines causing that
            Set returnedRng = .Range(.Cells(StartRow, endCol + 2), .Cells(EndRow, retEndCol))

            returnedData.FromExcelRange returnedRng, DetectLastRow:=False, _
                                        DetectLastColumn:=False
            returnedRng.Cells.Clear
        End If
    End With

    Set FilterData = returnedData.Clone()
End Function

'Multiple filters on the data (when you have mutiple conditions)
Private Function FiltersData(ByVal varData As BetterArray, _
                             ByVal criteriaData As BetterArray, _
                             ByVal returnedColumnsData As BetterArray) As BetterArray

    Dim dataRng As Range                         'Range of the data
    Dim dataWksh As Worksheet
    Dim returnedRng As Range
    Dim returnedColumnIndex As Long
    Dim returnedData As BetterArray
    Dim counter As Long
    Dim col As Long
    Dim EndRow As Long
    Dim endCol As Long
    Dim retEndRow As Long
    Dim retEndCol As Long

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeaders:=True)
    Set FiltersData = returnedData.Clone()

    'On Error GoTo ErrFilter
    If varData.Length <> criteriaData.Length Or varData.Length = 0 Then Exit Function

    'returned Range
    col = 2                                      'Iterator for columns to extract
    Set dataWksh = Wksh()

    With dataWksh

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
            Set returnedRng = .Range(.Cells(StartRow + 1, returnedColumnIndex), .Cells(EndRow, returnedColumnIndex)).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            If Not returnedRng Is Nothing Then returnedRng.Copy .Cells(StartRow, endCol + col)
            col = col + 1
        Next

        .AutoFilter.ShowAllData

        'end row and end columns of data to return
        retEndRow = .Cells(.Rows.Count, endCol + 2).End(xlUp).Row
        retEndCol = .Cells(StartRow, .Columns.Count).End(xlToLeft).Column

        If Not IsEmpty(.Cells(StartRow, endCol + 2)) Then
            Set returnedRng = .Range(.Cells(StartRow, endCol + 2), .Cells(retEndRow, retEndCol))
            'returned data is the data returned by the filter
            returnedData.FromExcelRange returnedRng, DetectLastRow:=False, DetectLastColumn:=False
            returnedRng.Cells.Clear
        End If
    End With
    Set FiltersData = returnedData.Clone()
End Function

'Import Sub
Private Sub Import(ByVal fromWksh As Worksheet, fromStartRow As Long, fromStartcol As Long)
    Dim importingData As IDataSheet
    Dim impcolRng As Range
    Dim actcolRng As Range
    Dim cellRng As Range
    Dim impColValues As BetterArray
    Dim hRng As Range
    Dim headerValue As String

    'Trace on eventual errors on import
    If (fromWksh Is Nothing) Or (fromStartRow <= 0) Or (fromStartcol <= 0) Then Exit Sub
    Set importingData = DataSheet.Create(fromWksh, fromStartRow, fromStartcol)

    Set hRng = HeaderRange()
    Set impColValues = New BetterArray

    For Each cellRng In hRng
        headerValue = cellRng.Value
        Set actcolRng = DataRange(headerValue) 'actual column range to replace

        'clear the content of the range (even if it does not exists in importing data)
        actcolRng.ClearContents

        If importingData.ColumnExists(headerValue) Then
            Set impcolRng = importingData.DataRange(headerValue)
            impColValues.FromExcelRange impcolRng, DetectLastCol:=False, DetectLastRow:=False

            'Transpose data to this range
            impColValues.ToExcelRange actcolRng.Cells(1, 1)
            impColValues.Clear
        End If
    Next
End Sub

'Export
Private Sub Export(ByVal Wkb As Workbook, Optional ByVal filteredVarName As String = "__all__", _
                   Optional ByVal filteredCondition As String = "<>", Optional ByVal Hide As Long = xlSheetHidden)

    Dim dataExportData As BetterArray
    Dim dataWksh As Worksheet
    Dim dataExportWksh As Worksheet
    Dim sheetDataExist As Boolean

    'Trace on errors on export
    If Wkb Is Nothing Then Exit Sub
    'Test if the sheet  exists in the workbook
    Set dataWksh = Wksh()

    On Error Resume Next
    Set dataExportWksh = Wkb.Worksheets(dataWksh.Name)
    sheetDataExist = (Not dataExportWksh Is Nothing)
    On Error GoTo 0

    'Trace on errors on exports
    If sheetDataExist Then Exit Sub
    Wkb.Worksheets.Add(After:=Wkb.Worksheets(Wkb.Worksheets.Count)).Name = dataWksh.Name
    Set dataExportData = New BetterArray
    Set dataExportWksh = Wkb.Worksheets(dataWksh.Name)

    If filteredVarName = "__all__" Then
        dataExportData.FromExcelRange DataRange(colName:="__all__", includeHeaders:=True)
    Else
        Set dataExportData = FilterData(filteredVarName, filteredCondition, "__all__", includeHeaders:=True)
    End If

    dataExportData.ToExcelRange dataExportWksh.Cells(StartRow, StartColumn)
    dataExportWksh.Visible = Hide
End Sub

'Errors ----------------------------------------------------------------------
Private Sub ThrowErrorEmptySheetGiven(Optional ByVal Class As String = ":")
    Err.Raise VBA.vbObjectError + 513, CLASSNAME, "Unable to create an instance of the class" & Class & " Given Worksheet object does not exists, please verify if you set it."
End Sub

Private Sub ThrowErrorStartNumber(Optional ByVal startNumber As Long = -1, Optional ByVal startLabel As String = " Start ", Optional ByVal Class As String = ":")
    Err.Raise VBA.vbObjectError + 514, CLASSNAME, "Unable to create an instance of the class" & Class & " The " & startLabel & _
                                                 "should be >= 1. " & "Actual value : " & startNumber
End Sub

Private Sub ThrowErrorUnFoundColumn(ByVal colName As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Column " & colName & " does not exists" & " in worksheet " & Wksh.Name
End Sub

'Inteface ----------------------------------------------------------------------

'Methods
Private Function IDataSheet_FilterData(ByVal varName As String, ByVal criteriaName As String, ByVal returnedColumnName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
    Set IDataSheet_FilterData = FilterData(varName:=varName, criteriaName:=criteriaName, returnedColumnName:=returnedColumnName, includeHeaders:=includeHeaders)
End Function

Private Function IDataSheet_FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray
    Set IDataSheet_FiltersData = FiltersData(varData, criteriaData, returnedColumnsData)
End Function

Private Function IDataSheet_ColumnExists(ByVal colName As String) As Boolean
    IDataSheet_ColumnExists = ColumnExists(colName)
End Function

Private Function IDataSheet_ColumnIndex(ByVal colName As String, _
                                       Optional ByVal inDataRange As Boolean = False, _
                                       Optional shouldExist As Boolean = False) As Long
    IDataSheet_ColumnIndex = ColumnIndex(colName, inDataRange, shouldExist)
End Function

'Properties
Private Property Get IDataSheet_DataRange(Optional ByVal colName As String = "__all__", _
                                         Optional ByVal includeHeaders As Boolean = False) As Range
    Set IDataSheet_DataRange = DataRange(colName, includeHeaders)
End Property

'Getters for the encapsulated data
Private Property Get IDataSheet_StartRow() As Long
    IDataSheet_StartRow = StartRow()
End Property

Private Property Get IDataSheet_StartColumn() As Long
    IDataSheet_StartColumn = StartColumn()
End Property

Private Property Get IDataSheet_Wksh() As Worksheet
    Set IDataSheet_Wksh = Wksh()
End Property

Private Property Get IDataSheet_DataEndRow() As Long
    IDataSheet_DataEndRow = DataEndRow()
End Property

Private Property Get IDataSheet_DataEndColumn() As Long
    IDataSheet_DataEndColumn = DataEndColumn()
End Property

Private Sub IDataSheet_Export(ByVal toWkb As Workbook, _
                             Optional ByVal filteredVarName As String = "__all__", _
                             Optional ByVal filteredCondition As String = "<>", Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, filteredVarName, filteredCondition, Hide
End Sub

Private Sub IDataSheet_Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Import fromWksh, fromStartRow, fromStartcol
End Sub


