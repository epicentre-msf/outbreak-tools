VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLdictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Dictionary Class"

'@Folder("Dictionary")
'@ModuleDescription("Interface for dicitonary in the linelist")
'@IgnoreModule
'@PredeclareId
Option Explicit
Implements ILLdictionary

'Types of the class
Private Type TDictionary
    dictWksh As Worksheet 'Worksheet of the dictionary
    startLn As Long  'The line where the dictionary starts
    startCol As Long 'The column where the dictionary starts (the first column is the column of the variable names.
End Type

Private this As TDictionary

Public Function Create(ByVal dictWksh As Worksheet, ByVal dictStartLine As Long, _
                       ByVal dictStartColumn As Long) As ILLdictionary
    With New LLdictionary
        Set .Wksh = dictWksh
        .StartLine = dictStartLine
        .StartColumn = dictStartColumn
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ILLdictionary
    Set Self = Me
End Property

'Worksheet
Public Property Get Wksh() As Worksheet
    Set Wksh = this.dictWksh
End Property

Friend Property Set Wksh(ByVal dictWksh As Worksheet)
    Set this.dictWksh = dictWksh
End Property

'StartLine
Public Property Get StartLine() As Long
    StartLine = this.startLn
End Property

Friend Property Let StartLine(ByVal dictStartLine As Long)
    this.startLn = dictStartLine
End Property

'StartColumn
Public Property Get StartColumn() As Long
    StartColumn = this.startCol
End Property

Friend Property Let StartColumn(ByVal dictStartColumn As Long)
    this.startCol = dictStartColumn
End Property

'End row and columns of the dictionary
Private Property Get DictEndRow() As Long
   With Wksh
        DictEndRow = .Cells(.Rows.Count, StartColumn()).End(xlUp).Row
   End With
End Property

Private Property Get DictEndColumn() As Long
     With Wksh
        DictEndColumn = .Cells(StartLine(), .Columns.Count).End(xlToLeft).Column
    End With
End Property

'Dictionary header range
Public Property Get HeaderRange() As Range
    'updating the header range
    Dim endCol As Long

    With Wksh
        Set HeaderRange = .Range(.Cells(StartLine, StartColumn), .Cells(StartLine, DictEndColumn))
    End With
End Property

'Dictionary data Range
Public Property Get DataRange(Optional includeHeader As Boolean = False) As Range
    Dim startRow As Long
    Dim dictWksh As Worksheet
    Set dictWksh = Wksh()

    With dictWksh
        startRow = IIF(includeHeader, StartLine, StartLine + 1)
        Set DataRange = .Range(.Cells(startRow, StartColumn), .Cells(DictEndRow, DictEndColumn))
    End With
End Property

'Check if colname exists
Public Property Get ColumnExists(colName As String) As Boolean
    Dim hRng As Range
    Set hRng = HeaderRange() 'Header Range of the dictionary
    ColumnExists = Not (hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Property

'Column Index within the data range of the dictionary
Private Function ColumnIndex(colName As String, Optional InDataRange As Boolean = False) As Long
    Dim hRng As Range
    Dim colIndex As Long
    Set hRng = HeaderRange()

    If ColumnExists(colName) Then
        colIndex = hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True).Column
        ColumnIndex = IIF(InDataRange, colIndex - StartColumn + 1, colIndex)
    Else
        ColumnIndex = StartColumn
    End If
End Function

'Column
Public Property Get Column(ByVal colName As String) As BetterArray

    Dim colNum As Long
    Dim colRng As Range
    Dim dictWksh As Worksheet
    Dim colData As BetterArray 'BetterArray of the column data to return

    Set colData = New BetterArray
    colData.LowerBound = 1 'Ensure it starts at 1

    Set dictWksh = Wksh() 'Dictionary worksheet

    With dictWksh

        If ColumnExists(colName) Then
            colNum = ColumnIndex(colName)
            Set colRng = .Range(.Cells(StartLine() + 1, colNum), .Cells(DictEndRow, colNum))

        ElseIf colName = "__all__" Then
            Set colRng = .Range(.Cells(StartLine() + 1, StartColumn()), .Cells(DictEndRow, DictEndColumn))

        End If
    End With

    If Not colRng Is Nothing Then
        colData.FromExcelRange colRng, DetectLastRow:=False, DetectLastColumn:=False
    End If

    Set Column = colData.Clone 'Can be empty
End Property

'Filtered Data
Public Property Get FilterData(ByVal varName As String, ByVal criteriaValue As String, ByVal returnedColumnName As String) As BetterArray

    Dim dataRng As Range
    Dim returnedRng As Range
    Dim returnedColumnIndex As Long
    Dim returnedData As BetterArray
    Dim dictWksh As Worksheet
    Dim endCol As Long
    Dim endRow As Long

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeader:=True) 'Plage of the dictionary data
    Set dictWksh = Wksh()
    endCol = DictEndColumn() 'End column and rows of the dictionary
    endRow = DictEndRow()
    returnedColumnIndex = ColumnIndex(returnedColumnName) 'Which index to return, the startcolumn is returned if the returned colname is not found

    On Error GoTo ErrFilter

    With dataRng
        .AutoFilter Field:=ColumnIndex(varName, InDataRange:=True), Criteria1:=criteriaValue
    End With

    With dictWksh
        Set returnedRng = .Range(.Cells(StartLine + 1, returnedColumnIndex), .Cells(endRow, returnedColumnIndex)).SpecialCells(xlCellTypeVisible)
        returnedRng.Copy .Cells(1, endCol + 2)
        .AutoFilter.ShowAllData
        'end row of data to return
        endRow = .Cells(.Rows.Count, endCol + 2).End(xlToUp).Row
        Set returnedRng = .Range(.Cells(1, endCol + 2), .Cells(endRow, endCol + 2))

        returnedData.FromExcelRange returnedRng, DetectLastRow:=False, DetectLastColumn:=False
        returnedRng.Cells.Clear
    End With

    Set FilterData = returnedData.Clone()
    Exit Property

ErrFilter:
    MsgBox "Unable to Filter the Dictionary : " & Err.Description & " - " & Err.Number
    If Not dictWksh.AutoFilter Is Nothing Then dictWksh.AutoFilter.ShowAllData
    Exit Property
End Property

'Multiple filters on the data (when you have mutiple conditions)
Public Property Get FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray

    Dim dataRng As Range
    Dim dictWksh As Worksheet
    Dim returnedRng As Range
    Dim returnedColumnIndex As Long
    Dim returnedData As BetterArray
    Dim counter As Long
    Dim col As Long
    Dim endRow As Long
    Dim endCol As Long
    Dim retEndRow As Long
    Dim retEndCol As Long

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1
    Set dataRng = DataRange(includeHeader:=True)
    endRow = DictEndRow
    endCol = DictEndColumn

    On Error GoTo ErrFilter
    If varData.Length <> criteriaData.Length Or varData.Length = 0 Then Exit Property

    With dataRng
        For counter = varData.LowerBound To varData.UpperBound
            .AutoFilter Field:=ColumnIndex(varData.Item(counter), InDataRange:=True), Criteria1:=criteriaData.Item(counter)
        Next
    End With

    'returned Range
    col = 1 'Iterator for columns to extract
    Set dictWksh = Wksh()

    With dictWksh

        For counter = returnedColumnsData.LowerBound To returnedColumnsData.UpperBound
            returnedColumnIndex = ColumnIndex(returnedColumnsData.Item(counter))
            Set returnedRng = .Range(.Cells(StartLine + 1, returnedColumnIndex), .Cells(endRow, returnedColumnIndex)).SpecialCells(xlCellTypeVisible)
            returnedRng.Copy .Cells(col, endCol + 2)
            col = col + 1
        Next

        .AutoFilter.ShowAllData

        'end row and end columns of data to return
        retEndRow = .Cells(.Rows.Count, endCol + 2).End(xlUp).Row
        retEndCol = .Cells(1, .Columns.Count).End(xlToLeft).Column

        Set returnedRng = .Range(.Cells(1, endCol + 2), .Cells(retEndRow, retEndCol))
        returnedColumnsData.FromExcelRange returnedRng, DetectLastRow:=False, DetectLastColumn:=False
        returnedRng.Cells.Clear
    End With

    Set FiltersData = returnedColumnsData.Clone()
    Exit Property

ErrFilter:
    If Not dataRng.AutoFilter Is Nothing Then dataRng.AutoFilter.ShowAllData
    MsgBox "Unable to Filter the Dictionary : " & Err.Description & " - " & Err.Number
End Property

'Exposed methods
Public Property Get ILLdictionary_Column(ByVal columnName As String) As BetterArray
    Set ILLdictionary_Column = Column(columnName)
End Property

Public Property Get ILLdictionary_ColumnExists(ByVal columnName As String) As Boolean
    ILLdictionary_ColumnExists = ColumnExists(columnName)
End Property

Public Property Get ILLdictionary_FilterData(ByVal varName As String, ByVal criteriaName As String, ByVal returnedColumnName As String) As BetterArray
    Set ILLdictionary_FilterData = FilterData(varName, criteriaName, returnedColumnName)
End Property

Public Property Get ILLdictionary_FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray
    Set ILLdictionary_FiltersData = FiltersData(varData, criteriaData, returnedColumnsData)
End Property

'Getters for the encapsulated data
Public Property Get ILLdictionary_StartLine() As Long
    ILLdictionary_StartLine = StartLine()
End Property
Public Property Get ILLdictionary_StartColumn() As Long
    ILLdictionary_StartColumn = StartColumn
End Property
