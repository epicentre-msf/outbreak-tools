VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSectionDictionaryDataStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("IDataSheet test double for section context tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDataSheet

Private hostSheet As Worksheet
Private startRow As Long
Private startColumn As Long
Private endRow As Long
Private endColumn As Long

Public Sub Configure(ByVal wk As Worksheet, _
                     ByVal dataStartRowValue As Long, _
                     ByVal dataStartColumnValue As Long, _
                     ByVal dataEndRowValue As Long, _
                     ByVal dataEndColumnValue As Long)

    Set hostSheet = wk
    startRow = dataStartRowValue
    startColumn = dataStartColumnValue
    endRow = dataEndRowValue
    endColumn = dataEndColumnValue
End Sub

Private Property Get IDataSheet_Wksh() As Worksheet
    Set IDataSheet_Wksh = hostSheet
End Property

Private Property Get IDataSheet_DataStartRow() As Long
    IDataSheet_DataStartRow = startRow
End Property

Private Property Get IDataSheet_DataStartColumn() As Long
    IDataSheet_DataStartColumn = startColumn
End Property

Private Property Get IDataSheet_Name() As String
    If hostSheet Is Nothing Then Exit Property
    IDataSheet_Name = hostSheet.Name
End Property

Private Property Get IDataSheet_DataEndRow() As Long
    IDataSheet_DataEndRow = endRow
End Property

Private Property Get IDataSheet_DataEndColumn() As Long
    IDataSheet_DataEndColumn = endColumn
End Property

Private Property Get IDataSheet_HeaderRange() As Range
    If hostSheet Is Nothing Then Exit Property
    Set IDataSheet_HeaderRange = hostSheet.Range(hostSheet.Cells(startRow - 1, startColumn), _
                                                 hostSheet.Cells(startRow - 1, endColumn))
End Property

Private Function IDataSheet_ColumnExists(ByVal colName As String, _
                                         Optional ByVal strictSearch As Boolean = True, _
                                         Optional ByVal matchCase As Boolean = True) As Boolean

    IDataSheet_ColumnExists = (IDataSheet_ColumnIndex(colName) > 0)
End Function

Private Function IDataSheet_ColumnIndex(ByVal colName As String, _
                                        Optional ByVal inDataRange As Boolean = False, _
                                        Optional shouldExist As Boolean = False, _
                                        Optional ByVal strictSearch As Boolean = True, _
                                        Optional ByVal matchCase As Boolean = True) As Long

    Dim headers As Range
    Dim lookupName As String
    Dim matchMode As Long
    Dim headerCell As Range

    If hostSheet Is Nothing Then Exit Function

    Set headers = IDataSheet_HeaderRange
    If headers Is Nothing Then Exit Function

    lookupName = IIf(matchCase, colName, LCase$(colName))

    For Each headerCell In headers
        If matchCase Then
            If StrComp(headerCell.Value, lookupName, vbBinaryCompare) = 0 Then
                IDataSheet_ColumnIndex = headerCell.Column
                Exit Function
            End If
        Else
            If StrComp(LCase$(CStr(headerCell.Value)), lookupName, vbBinaryCompare) = 0 Then
                IDataSheet_ColumnIndex = headerCell.Column
                Exit Function
            End If
        End If
    Next headerCell

    If shouldExist Then
        Err.Raise vbObjectError + 700, "LLSectionDictionaryDataStub.ColumnIndex", _
                  "Column '" & colName & "' not located in data sheet"
    End If
End Function

Private Property Get IDataSheet_DataRange(Optional ByVal colName As String = "__all__", _
                                          Optional ByVal includeHeaders As Boolean = False, _
                                          Optional ByVal strictSearch As Boolean = True, _
                                          Optional ByVal matchCase As Boolean = True) As Range

    Dim firstRow As Long
    Dim firstColumn As Long
    Dim headerRow As Long

    If hostSheet Is Nothing Then Exit Property

    firstRow = startRow
    If includeHeaders Then firstRow = startRow - 1

    firstColumn = startColumn

    If colName <> "__all__" Then
        firstColumn = IDataSheet_ColumnIndex(colName, shouldExist:=True, matchCase:=matchCase)
    End If

    headerRow = startRow - 1

    If includeHeaders And firstRow < startRow Then
        Set IDataSheet_DataRange = hostSheet.Range(hostSheet.Cells(firstRow, firstColumn), _
                                                   hostSheet.Cells(endRow, firstColumn))
    Else
        Set IDataSheet_DataRange = hostSheet.Range(hostSheet.Cells(startRow, firstColumn), _
                                                   hostSheet.Cells(endRow, firstColumn))
    End If
End Property

Private Sub IDataSheet_Clean()
    'Not required for stub
End Sub

Private Sub IDataSheet_AddFormatsColumns(ByVal matchColumnsCase As Boolean, ParamArray lsCols() As Variant)
    'Not required for stub
End Sub

Private Function IDataSheet_FilterData(ByVal varName As String, ByVal criteriaName As String, _
                                       ByVal returnedColumnName As String, _
                                       Optional ByVal includeHeaders As Boolean = False) As BetterArray
    Err.Raise vbObjectError + 701, "LLSectionDictionaryDataStub.FilterData", "Not implemented"
End Function

Private Function IDataSheet_FiltersData(ByVal varData As BetterArray, _
                                        ByVal criteriaData As BetterArray, _
                                        ByVal returnedColumnsData As BetterArray) As BetterArray
    Err.Raise vbObjectError + 702, "LLSectionDictionaryDataStub.FiltersData", "Not implemented"
End Function

Private Sub IDataSheet_Import(ByVal importedData As Object, Optional ByVal strictColumnSearch As Boolean = True)
    Err.Raise vbObjectError + 703, "LLSectionDictionaryDataStub.Import", "Not implemented"
End Sub

Private Sub IDataSheet_ImportFormat(ByVal importData As Object, Optional ByVal matchColumnsCase As Boolean = True)
    Err.Raise vbObjectError + 704, "LLSectionDictionaryDataStub.ImportFormat", "Not implemented"
End Sub

Private Sub IDataSheet_Export(ByVal toWkb As Workbook, _
                              Optional ByVal filteredVarName As String = "__all__", _
                              Optional ByVal filteredCondition As String = "<>", _
                              Optional ByVal Hide As Long = xlSheetHidden)
    Err.Raise vbObjectError + 705, "LLSectionDictionaryDataStub.Export", "Not implemented"
End Sub

Private Property Get IDataSheet_HasCheckings() As Boolean
    IDataSheet_HasCheckings = False
End Property

Private Property Get IDataSheet_CheckingValues() As Object
    Set IDataSheet_CheckingValues = Nothing
End Property

