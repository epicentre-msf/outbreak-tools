VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSectionDictionaryStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("ILLdictionary test double for section context tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLdictionary

Private dataSheetStub As LLSectionDictionaryDataStub
Private hostSheet As Worksheet
Private startRow As Long
Private startColumn As Long
Private endRow As Long
Private endColumn As Long
Private totalExports As Long

Public Sub Configure(ByVal sheet As Worksheet, _
                     ByVal dataStartRowValue As Long, _
                     ByVal dataStartColumnValue As Long, _
                     ByVal dataEndRowValue As Long, _
                     ByVal dataEndColumnValue As Long)

    Set hostSheet = sheet
    startRow = dataStartRowValue
    startColumn = dataStartColumnValue
    endRow = dataEndRowValue
    endColumn = dataEndColumnValue

    Set dataSheetStub = New LLSectionDictionaryDataStub
    dataSheetStub.Configure hostSheet, startRow, startColumn, endRow, endColumn
End Sub

Private Function ResolveColumnRange(ByVal colName As String, _
                                    Optional ByVal includeHeaders As Boolean = False) As Range
    Set ResolveColumnRange = dataSheetStub.DataRange(colName:=colName, _
                                                   includeHeaders:=includeHeaders, _
                                                   matchCase:=False)
End Function

Private Property Get ILLdictionary_Data() As IDataSheet
    Set ILLdictionary_Data = dataSheetStub
End Property

Private Property Get ILLdictionary_DataRange(Optional ByVal colName As String = "__all__", _
                                              Optional ByVal includeHeaders As Boolean = False) As Range

    If hostSheet Is Nothing Then Exit Property

    If colName = "__all__" Then
        If includeHeaders Then
            Set ILLdictionary_DataRange = hostSheet.Range(hostSheet.Cells(startRow - 1, startColumn), _
                                                          hostSheet.Cells(endRow, endColumn))
        Else
            Set ILLdictionary_DataRange = hostSheet.Range(hostSheet.Cells(startRow, startColumn), _
                                                          hostSheet.Cells(endRow, endColumn))
        End If
    Else
        Set ILLdictionary_DataRange = ResolveColumnRange(colName, includeHeaders)
    End If
End Property

Private Property Get ILLdictionary_Prepared() As Boolean
    ILLdictionary_Prepared = True
End Property

Private Property Get ILLdictionary_TotalNumberOfExports() As Long
    ILLdictionary_TotalNumberOfExports = totalExports
End Property

Private Property Let ILLdictionary_TotalNumberOfExports(ByVal numberOfExports As Long)
    totalExports = numberOfExports
End Property

Private Function ILLdictionary_ColumnExists(ByVal colName As String, _
                                            Optional ByVal checkValidity As Boolean) As Boolean
    ILLdictionary_ColumnExists = (Not ResolveColumnRange(colName) Is Nothing)
End Function

Private Function ILLdictionary_UniqueValues(ByVal colName As String) As BetterArray
    Dim result As BetterArray
    Set result = New BetterArray
    result.LowerBound = 1
    Set ILLdictionary_UniqueValues = result
End Function

Private Function ILLdictionary_VariableExists(ByVal varName As String) As Boolean
    Dim values As Range
    Dim cell As Range

    Set values = ResolveColumnRange("variable name")
    If values Is Nothing Then Exit Function

    For Each cell In values
        If StrComp(CStr(cell.Value), varName, vbTextCompare) = 0 Then
            ILLdictionary_VariableExists = True
            Exit Function
        End If
    Next cell
End Function

Private Sub ILLdictionary_AddColumn(ByVal colName As String)
    Err.Raise vbObjectError + 710, "LLSectionDictionaryStub.AddColumn", "Not implemented"
End Sub

Private Sub ILLdictionary_InsertColumn(ByVal colName As String, ByVal After As String)
    Err.Raise vbObjectError + 711, "LLSectionDictionaryStub.InsertColumn", "Not implemented"
End Sub

Private Sub ILLdictionary_RemoveColumn(ByVal colName As String)
    Err.Raise vbObjectError + 712, "LLSectionDictionaryStub.RemoveColumn", "Not implemented"
End Sub

Private Sub ILLdictionary_AddRows()
    Err.Raise vbObjectError + 713, "LLSectionDictionaryStub.AddRows", "Not implemented"
End Sub

Private Sub ILLdictionary_RemoveRows()
    Err.Raise vbObjectError + 714, "LLSectionDictionaryStub.RemoveRows", "Not implemented"
End Sub

Private Property Get ILLdictionary_ChoicesVars() As BetterArray
    Set ILLdictionary_ChoicesVars = Nothing
End Property

Private Property Get ILLdictionary_GeoVars() As BetterArray
    Set ILLdictionary_GeoVars = Nothing
End Property

Private Property Get ILLdictionary_TimeVars() As BetterArray
    Set ILLdictionary_TimeVars = Nothing
End Property

Private Property Get ILLdictionary_SpecialVars(ByVal first_condition As String, _
                                                Optional ByVal second_condition As String = vbNullString, _
                                                Optional ByVal conditionName As String = "Control") As BetterArray
    Set ILLdictionary_SpecialVars = Nothing
End Property

Private Sub ILLdictionary_Prepare(ByVal PreservedSheetNames As BetterArray, _
                                  ByVal geoObject As Object, _
                                  ByVal msgTrads As ITranslationObject, _
                                  Optional ByVal tablePrefix As String = "table")
    'Not required for stub
End Sub

Private Sub ILLdictionary_Clean(Optional ByVal removeAddedColumns As Boolean = False)
    'Not required for stub
End Sub

Private Sub ILLdictionary_Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Err.Raise vbObjectError + 715, "LLSectionDictionaryStub.Import", "Not implemented"
End Sub

Private Sub ILLdictionary_Export(ByVal toWkb As Workbook, _
                                 Optional ByVal exportType As String = "__all__", _
                                 Optional ByVal addListObject As Boolean = True, _
                                 Optional ByVal Hide As Long = xlSheetHidden)
    Err.Raise vbObjectError + 716, "LLSectionDictionaryStub.Export", "Not implemented"
End Sub

Private Sub ILLdictionary_Translate(ByVal TransObject As ITranslationObject)
    'Not required for stub
End Sub

