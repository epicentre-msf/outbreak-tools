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
Private DICTIONARYCOLUMNS As BetterArray
Private Const SHEETTYPEVERTICAL As String = "vlist1D"
Private Const SHEETTYPEHORIZONTAL As String = "hlist2D"

Public Function Create(ByVal dictWksh As Worksheet, ByVal dictStartRow As Long, _
                       ByVal dictStartColumn As Long) As ILLdictionary
    With New LLdictionary
        Set .Wksh = dictWksh
        .StartRow = dictStartRow
        .StartColumn = dictStartColumn
        Set Create = .Self
    End With
End Function

Private Sub ColumnInitialize()
    Set DICTIONARYCOLUMNS = New BetterArray
    DICTIONARYCOLUMNS.Push "variable name", "main label", _
                            "sub label", "note", "sheet name", _
                            "sheet type", "main section", "sub section", _
                            "status", "personal identifier", _
                            "type", "control", _
                            "control details", "unique", _
                            "export 1", "export 2", _
                            "export 3", "export 4", _
                            "export 5", "min", _
                            "max", "alert", "message", _
                            "table name", "column index"
End Sub



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

'StartRow
Public Property Get StartRow() As Long
    StartRow = this.startLn
End Property

Friend Property Let StartRow(ByVal dictStartRow As Long)
    this.startLn = dictStartRow
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
        DictEndRow = .Cells(.Rows.Count, StartColumn).End(xlUp).Row
   End With
End Property

Private Property Get DictEndColumn() As Long
     With Wksh
        DictEndColumn = .Cells(StartRow, .Columns.Count).End(xlToLeft).Column
    End With
End Property

'Dictionary header range
Public Property Get HeaderRange() As Range
    'updating the header range
    Dim endCol As Long

    With Wksh
        Set HeaderRange = .Range(.Cells(StartRow, StartColumn), .Cells(StartRow, DictEndColumn))
    End With
End Property

'Dictionary data Range
Public Property Get DataRange(Optional colName As String = "__all__", Optional includeHeaders As Boolean = False) As Range
    Dim customStartRow As Long
    Dim dictWksh As Worksheet
    Dim endCol As Long
    Dim customStartCol As Long
    Set dictWksh = Wksh()


    With dictWksh
        customStartRow = IIf(includeHeaders, StartRow, StartRow + 1)
        If colName = "__all__" Then
            customStartCol = StartColumn
            endCol = DictEndColumn
        Else
            customStartCol = ColumnIndex(colName)
            endCol = customStartCol
        End If
        Set DataRange = .Range(.Cells(customStartRow, customStartCol), .Cells(DictEndRow, endCol))
    End With
End Property

'Check if colname exists
Public Property Get ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean = False) As Boolean
    Dim hRng As Range
    Dim Exist As Boolean
    Set hRng = HeaderRange() 'Header Range of the dictionary

    If colName = vbNullString Then
        Exist = False
    Else
        Exist = Not (hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
        If Exist And checkValidity Then
            If DICTIONARYCOLUMNS Is Nothing Then ColumnInitialize
            Exist = DICTIONARYCOLUMNS.Includes(colName)
        End If
    End If

    ColumnExists = Exist
End Property

'Column Index within the data range of the dictionary
Private Function ColumnIndex(ByVal colName As String, Optional ByVal inDataRange As Boolean = False) As Long
    Dim hRng As Range
    Dim colIndex As Long
    Set hRng = HeaderRange()

    If ColumnExists(colName) Then
        colIndex = hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=True).Column
        ColumnIndex = IIf(inDataRange, colIndex - StartColumn + 1, colIndex)
    Else
        ColumnIndex = StartColumn
    End If
End Function


'Column
Public Property Get Column(ByVal colName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray

    Dim colNum As Long
    Dim colRng As Range
    Dim colStartRow As Long
    Dim dictWksh As Worksheet
    Dim colData As BetterArray 'BetterArray of the column data to return

    Set colData = New BetterArray
    colData.LowerBound = 1 'Ensure it starts at 1

    colStartRow = IIf(includeHeaders, StartRow, StartRow + 1)

    Set dictWksh = Wksh() 'Dictionary worksheet

    With dictWksh
        If ColumnExists(colName) Then
            colNum = ColumnIndex(colName)
            Set colRng = .Range(.Cells(colStartRow, colNum), .Cells(DictEndRow, colNum))

        ElseIf colName = "__all__" Then
            Set colRng = .Range(.Cells(colStartRow, StartColumn), .Cells(DictEndRow, DictEndColumn))

        End If
    End With

    If Not colRng Is Nothing Then
        colData.FromExcelRange colRng, DetectLastRow:=False, DetectLastColumn:=False
    End If

    Set Column = colData.Clone 'Can be empty
End Property

'Filtered Data
Public Property Get FilterData(ByVal varName As String, ByVal criteriaValue As String, ByVal returnedColumnName As String, _
                               Optional ByVal includeHeaders As Boolean = False) As BetterArray

    Dim dataRng As Range
    Dim returnedRng As Range
    Dim startRetColIndex As Long
    Dim startRetRow As Long
    Dim endRetColIndex As Long
    Dim returnedData As BetterArray
    Dim dictWksh As Worksheet
    Dim endCol As Long
    Dim endRow As Long
    Dim retEndCol As Long

    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeaders:=True) 'Plage of the dictionary data
    Set dictWksh = Wksh()
    endCol = DictEndColumn() 'End column and rows of the dictionary
    endRow = DictEndRow()

    startRetRow = IIf(includeHeaders, StartRow, StartRow + 1)

    If returnedColumnName = "__all__" Then
        startRetColIndex = StartColumn
        endRetColIndex = endCol
    Else
        startRetColIndex = ColumnIndex(returnedColumnName) 'Which index to return, the startcolumn is returned if the returned colname is not found
        endRetColIndex = startRetColIndex
    End If

    With dictWksh

        If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData

        With dataRng
            .AutoFilter Field:=ColumnIndex(varName, inDataRange:=True), Criteria1:=criteriaValue
        End With

        On Error Resume Next
            Set returnedRng = .Range(.Cells(startRetRow, startRetColIndex), .Cells(endRow, endRetColIndex)).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        .AutoFilter.ShowAllData

        If Not returnedRng Is Nothing Then
            returnedRng.Copy .Cells(StartRow, endCol + 2)

            'end row of data to return
            endRow = .Cells(.Rows.Count, endCol + 2).End(xlUp).Row
            retEndCol = .Cells(StartRow, .Columns.Count).End(xlToLeft).Column
            Set returnedRng = .Range(.Cells(StartRow, endCol + 2), .Cells(endRow, retEndCol))

            returnedData.FromExcelRange returnedRng, DetectLastRow:=False, DetectLastColumn:=False
            returnedRng.Cells.Clear
        End If
    End With

    Set FilterData = returnedData.Clone()
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
    
    Clean
    
    Set returnedData = New BetterArray
    returnedData.LowerBound = 1

    Set dataRng = DataRange(includeHeaders:=True)
    Set FiltersData = returnedData.Clone()

    'On Error GoTo ErrFilter
    If varData.Length <> criteriaData.Length Or varData.Length = 0 Then Exit Property

    'returned Range
    col = 2 'Iterator for columns to extract
    Set dictWksh = Wksh()

    With dictWksh

        If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
        endRow = DictEndRow()
        endCol = DictEndColumn()

        With dataRng
            For counter = varData.LowerBound To varData.UpperBound
                .AutoFilter Field:=ColumnIndex(varData.Item(counter), inDataRange:=True), Criteria1:=criteriaData.Item(counter)
            Next
        End With

        For counter = returnedColumnsData.LowerBound To returnedColumnsData.UpperBound
            returnedColumnIndex = ColumnIndex(returnedColumnsData.Item(counter))

            On Error Resume Next
            Set returnedRng = .Range(.Cells(StartRow + 1, returnedColumnIndex), .Cells(endRow, returnedColumnIndex)).SpecialCells(xlCellTypeVisible)
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
End Property

'------ Prepare and clean the dictionary

Public Property Get Prepared() As Boolean
    Dim dictWksh As Worksheet
    Set dictWksh = Wksh()
    Prepared = (dictWksh.Cells(DictEndRow + 1, StartColumn).Font.Color = vbBlue) And _
                ColumnExists("table name") And ColumnExists("column index") And _
                (Not ColumnExists("main section sort")) And (Not ColumnExists("main section index"))
End Property

Private Sub Clean()
    Dim endCol As Long
    Dim counter As Long
    Dim dictWksh As Worksheet

    Set dictWksh = Wksh()
    endCol = DictEndColumn
    counter = endCol
    
    With dictWksh
        Do While counter >= 1
            If Not ColumnExists(.Cells(StartRow, counter).value, checkValidity:=True) Then
                .Columns(counter).EntireColumn.Delete
            End If
            
            counter = counter - 1
         Loop
    End With
End Sub


'Prepare the dictionary for designer
Private Sub Prepare(Optional ByVal tablePrefix As String = "table")
    Dim endCol As Long
    Dim dictWksh As Worksheet
    Dim counter As Long
    Dim controlRng As Range

    If Not Prepared() Then
        'Test if the dictionary is prepared before doing the whole loop

            If Not ColumnExists("table name") Then
                 AppendNumberColumn onColName:="sheet name", _
                                    tablePrefix:=tablePrefix, _
                                    newColumn:="table name"    'Add table name
            End If
            
            
            If Not ColumnExists("column index") Then AppendColumnIndex 'Add Column Index

            'Testing control to update the Geo
            Set controlRng = DataRange("control")

            If (Not controlRng Is Nothing) Then
                If (Not controlRng.Find(What:="geo", LookAt:=xlWhole, MatchCase:=True) Is Nothing) And _
                    (controlRng.Find(What:="geo2", LookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
                    'Update the Geo
                    AppendGeoLines
                End If
            End If

            'Sort the dictionary accordingly and make unique varNames
            UniqueVarNames
            Clean

            'Add color to the last line for prepare indication
            Set dictWksh = Wksh()
            dictWksh.Cells(DictEndRow + 1, StartColumn).Font.Color = vbBlue
    End If
End Sub

'Make unique variable names
Private Sub UniqueVarNames()

    Dim dictWksh As Worksheet
    Dim endRow As Long
    Dim varNameIndex As Long
    Dim counter As Long
    Dim varName As String

    Dim varTrackData As BetterArray

    Set varTrackData = New BetterArray
    Set dictWksh = Wksh()
    endRow = DictEndRow
    varNameIndex = ColumnIndex("variable name")

    With dictWksh
        For counter = StartRow + 1 To endRow

            varName = Application.WorksheetFunction.Trim(.Cells(counter, varNameIndex).value)
            varName = Replace(varName, " ", "_")

            If varTrackData.Includes(varName) Then varName = varName & "_"

            varTrackData.Push varName
            .Cells(counter, varNameIndex).value = varName
        Next
    End With
End Sub


'Append Geo Lines

Private Sub AppendGeoLines()

    Dim dictRng As Range
    Dim geoRng As Range
    Dim pastedGeoRng As Range
    Dim dictWksh As Worksheet
    Dim geoRowIndex As Long
    Dim controlIndex As Long
    Dim varNameIndex As Long
    Dim subSectionIndex As Long
    Dim mainLabelIndex As Long
    Dim varName As String
    Dim endRow As Long
    Dim endCol As Long

    Dim counter As Long

    Set dictWksh = Wksh()
    Set dictRng = DataRange("control")
    controlIndex = ColumnIndex("control")
    varNameIndex = ColumnIndex("variable name")
    mainLabelIndex = ColumnIndex("main label")
    subSectionIndex = ColumnIndex("sub section")

    endRow = DictEndRow
    endCol = DictEndColumn
    geoRowIndex = StartRow + 1

    With dictWksh
        Do While (Not dictRng.Find(What:="geo", LookAt:=xlWhole, MatchCase:=True) Is Nothing)

            geoRowIndex = dictRng.Find(What:="geo", LookAt:=xlWhole, MatchCase:=True).Row

            Set geoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
            varName = .Cells(geoRowIndex, varNameIndex).value

            'Replace sub-section by main label (if required)
            If IsEmpty(.Cells(geoRowIndex, subSectionIndex)) And (Not IsEmpty(.Cells(geoRowIndex, mainLabelIndex))) Then
                .Cells(geoRowIndex, subSectionIndex).value = .Cells(geoRowIndex, mainLabelIndex).value
                .Cells(geoRowIndex, mainLabelIndex).value = ""
            End If

            .Cells(geoRowIndex, varNameIndex).value = "adm1" & "_" & varName

            'Insert the others geo lines
            geoRowIndex = geoRowIndex + 1
            For counter = 4 To 2 Step -1
                .Rows(geoRowIndex).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Set pastedGeoRng = .Range(.Cells(geoRowIndex, StartColumn), .Cells(geoRowIndex, endCol))
                pastedGeoRng.value = geoRng.value 'Paste dictionary values
                .Cells(geoRowIndex, varNameIndex).value = "adm" & counter & "_" & varName
                .Cells(geoRowIndex, controlIndex).value = "geo" & counter
            Next

            geoRowIndex = geoRowIndex + 3
            Set dictRng = .Range(.Cells(geoRowIndex, controlIndex), .Cells(endRow, controlIndex))
        Loop
    End With
End Sub

Private Sub AppendNumberColumn(Optional ByVal onColName As String = "sheet name", _
                               Optional ByVal tablePrefix As String = "table", _
                               Optional ByVal newColumn As String = "table name")

    Dim endRow As Long
    Dim endCol As Integer
    Dim tableOrder As Long
    Dim sheetNameIndex As Long
    Dim sheetName As String
    Dim counter As Long

    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim sheetTrackData As BetterArray
    Dim tableTrackData As BetterArray

    Set sheetTrackData = New BetterArray
    Set tableTrackData = New BetterArray

    endRow = DictEndRow
    endCol = DictEndColumn
    tableOrder = 1
    sheetNameIndex = ColumnIndex(onColName)

    Set dictWksh = Wksh()

    With dictWksh
        .Cells(StartRow, endCol + 1).value = "number_" & newColumn
        sheetTrackData.Push .Cells(StartRow + 1, sheetNameIndex).value
        tableTrackData.Push tableOrder

        For counter = StartRow + 1 To endRow
            sheetName = .Cells(counter, sheetNameIndex).value
            If sheetTrackData.Includes(sheetName) Then
                .Cells(counter, endCol + 1).value = tableTrackData.Items(sheetTrackData.IndexOf(sheetName))
            Else
                tableOrder = tableOrder + 1
                sheetTrackData.Push sheetName
                tableTrackData.Push tableOrder
                .Cells(counter, endCol + 1).value = tableOrder
            End If
        Next

        'Sort on table name
        Set dictRng = DataRange
        dictRng.Sort key1:=DataRange("number_" & newColumn), order1:=xlAscending
        
        endCol = DictEndColumn
        .Cells(StartRow, endCol + 1).value = newColumn
        .Cells(StartRow + 1, endCol + 1).Formula = "= " & Chr(34) & tablePrefix & Chr(34) & " " & Chr(38) & " " & .Cells(StartRow + 1, endCol).Address(Rowabsolute:=False)
        Set dictRng = DataRange(newColumn)
        
        .Cells(StartRow + 1, endCol + 1).AutoFill dictRng, Type:=xlFillValues
        dictRng.value = dictRng.value 'remove the link to the formulas
    
    End With
    
    
End Sub


Private Sub AppendColumnIndex()

    Dim dictWksh As Worksheet
    Dim dictRng As Range
    Dim endCol As Long
    Dim endRow As Long
    Dim counter As Long
    Dim varOrder As Long
    Dim sheetTypeIndex As Long
    Dim sheetNameIndex As Long
    Dim controlIndex As Long
    Dim mainSectionIndex As Long

    Dim firstsheetNameCell As String
    Dim firstmainSectionCell As String
    Dim sheetName As String

    Set dictWksh = Wksh()
    endCol = DictEndColumn
    endRow = DictEndRow
    sheetTypeIndex = ColumnIndex("sheet type")
    sheetNameIndex = ColumnIndex("sheet name")
    mainSectionIndex = ColumnIndex("main section")
    controlIndex = ColumnIndex("control")

    sheetName = "&&&&&&;-------"
    varOrder = 1

    With dictWksh
        'Sorting
        .Cells(StartRow, endCol + 1).value = "main section sort"
        firstsheetNameCell = .Cells(StartRow + 1, sheetNameIndex).Address(Rowabsolute:=False) & " "
        firstmainSectionCell = .Cells(StartRow + 1, mainSectionIndex).Address(Rowabsolute:=False)
        .Cells(StartRow + 1, endCol + 1).Formula = "= IF(" & firstmainSectionCell & "=" & Chr(34) & Chr(34) & _
                                                    ", ROW(" & firstmainSectionCell & ")" & ", " & _
                                                   firstsheetNameCell & Chr(38) & Chr(34) & " - " & Chr(34) & Chr(38) & firstmainSectionCell & _
                                                   ")"
        Set dictRng = DataRange("main section sort")
        .Cells(StartRow + 1, endCol + 1).AutoFill dictRng, Type:=xlFillValues

        'Now append and sort
        AppendNumberColumn onColName:="main section sort", tablePrefix:="main section", newColumn:="main section index"

        'Filling column index
        endCol = DictEndColumn
        .Cells(StartRow, endCol + 1).value = "column index"
        For counter = StartRow + 1 To endRow
            If .Cells(counter, sheetNameIndex).value <> sheetName Then
                'New sheet name
                varOrder = IIf(.Cells(counter, sheetTypeIndex).value = SHEETTYPEVERTICAL, 4, 1)
                sheetName = .Cells(counter, sheetNameIndex).value
            Else
                'Old sheet name
                varOrder = IIf(.Cells(counter - 1, controlIndex).value = "geo", varOrder + 4, varOrder + 1)
            End If
            .Cells(counter, endCol + 1).value = varOrder 'variable order
        Next
    End With
End Sub


' INTERFACE __________________________________________________________________________________________________________________________

'Exposed methods
Public Property Get ILLdictionary_Column(ByVal columnName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
    Set ILLdictionary_Column = Column(columnName, includeHeaders)
End Property

Public Property Get ILLdictionary_ColumnExists(ByVal columnName As String) As Boolean
    ILLdictionary_ColumnExists = ColumnExists(columnName)
End Property

Public Property Get ILLdictionary_FilterData(ByVal varName As String, ByVal criteriaName As String, ByVal returnedColumnName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
    Set ILLdictionary_FilterData = FilterData(varName, criteriaName, returnedColumnName, includeHeaders)
End Property

Public Property Get ILLdictionary_FiltersData(ByVal varData As BetterArray, ByVal criteriaData As BetterArray, ByVal returnedColumnsData As BetterArray) As BetterArray
    Set ILLdictionary_FiltersData = FiltersData(varData, criteriaData, returnedColumnsData)
End Property

Public Property Get ILLdictionary_Prepared() As Boolean
    ILLdictionary_Prepared = Prepared
End Property

Public Property Get ILLdictionary_DataRange(Optional ByVal colName As String = "__all__", _
                                            Optional ByVal includeHeaders As Boolean = False) As Range
    Set ILLdictionary_DataRange = DataRange(colName, includeHeaders)
End Property

'Getters for the encapsulated data
Public Property Get ILLdictionary_StartRow() As Long
    ILLdictionary_StartRow = StartRow()
End Property
Public Property Get ILLdictionary_StartColumn() As Long
    ILLdictionary_StartColumn = StartColumn
End Property

Public Property Get ILLdictionary_Wksh() As Worksheet
    Set ILLdictionary_Wksh = Wksh()
End Property

'void methods
Public Sub ILLdictionary_Prepare(Optional ByVal tablePrefix As String = "table")
   Prepare tablePrefix
End Sub

Public Sub ILLdictionary_Clean()
    Clean
End Sub


