VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisDictionaryStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("ILLdictionary double exposing configurable variable existence checks")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ILLdictionary

Private Type TDictionaryStubState
    variableNames As BetterArray
    totalExports As Long
End Type

Private this As TDictionaryStubState

Private Sub Class_Initialize()
    Set this.variableNames = New BetterArray
    this.variableNames.LowerBound = 1
End Sub

'Register a variable recognised by the dictionary.
Public Sub AddVariable(ByVal variableName As String)
    Dim idx As Long

    For idx = this.variableNames.LowerBound To this.variableNames.UpperBound
        If StrComp(CStr(this.variableNames.Item(idx)), variableName, vbBinaryCompare) = 0 Then Exit Sub
    Next idx

    this.variableNames.Push variableName
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get ILLdictionary_Data() As IDataSheet
    Set ILLdictionary_Data = Nothing
End Property

Private Property Get ILLdictionary_DataRange(Optional ByVal colName As String = "__all__", _
                                              Optional ByVal includeHeaders As Boolean = False) As Range
    Set ILLdictionary_DataRange = Nothing
End Property

Private Property Get ILLdictionary_Prepared() As Boolean
    ILLdictionary_Prepared = True
End Property

Private Property Get ILLdictionary_TotalNumberOfExports() As Long
    ILLdictionary_TotalNumberOfExports = this.totalExports
End Property

Private Property Let ILLdictionary_TotalNumberOfExports(ByVal numberOfExports As Long)
    this.totalExports = numberOfExports
End Property

Private Function ILLdictionary_ColumnExists(ByVal colName As String, Optional ByVal checkValidity As Boolean) As Boolean
    ILLdictionary_ColumnExists = False
End Function

Private Function ILLdictionary_UniqueValues(ByVal colName As String) As BetterArray
    Set ILLdictionary_UniqueValues = New BetterArray
End Function

Private Function ILLdictionary_VariableExists(ByVal varName As String) As Boolean
    Dim idx As Long

    For idx = this.variableNames.LowerBound To this.variableNames.UpperBound
        If StrComp(CStr(this.variableNames.Item(idx)), varName, vbBinaryCompare) = 0 Then
            ILLdictionary_VariableExists = True
            Exit Function
        End If
    Next idx
End Function

Private Sub ILLdictionary_AddColumn(ByVal colName As String)
End Sub

Private Sub ILLdictionary_InsertColumn(ByVal colName As String, ByVal After As String)
End Sub

Private Sub ILLdictionary_RemoveColumn(ByVal colName As String)
End Sub

Private Sub ILLdictionary_AddRows()
End Sub

Private Sub ILLdictionary_RemoveRows()
End Sub

Private Property Get ILLdictionary_ChoicesVars() As BetterArray
    Set ILLdictionary_ChoicesVars = New BetterArray
End Property

Private Property Get ILLdictionary_GeoVars() As BetterArray
    Set ILLdictionary_GeoVars = New BetterArray
End Property

Private Property Get ILLdictionary_TimeVars() As BetterArray
    Set ILLdictionary_TimeVars = New BetterArray
End Property

Private Property Get ILLdictionary_SpecialVars(ByVal first_condition As String, Optional ByVal second_condition As String = vbNullString, Optional ByVal conditionName As String = "Control") As BetterArray
    Set ILLdictionary_SpecialVars = New BetterArray
End Property

Private Sub ILLdictionary_Prepare(ByVal PreservedSheetNames As BetterArray, ByVal geoObject As Object, ByVal msgTrads As ITranslationObject, Optional ByVal tablePrefix As String = "table")
End Sub

Private Sub ILLdictionary_Clean(Optional ByVal removeAddedColumns As Boolean = False)
End Sub

Private Sub ILLdictionary_Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
End Sub

Private Sub ILLdictionary_Export(ByVal toWkb As Workbook, Optional ByVal exportType As String = "__all__", Optional ByVal includeHeaders As Boolean = True, Optional ByVal filterName As String = "__all__")
End Sub

Private Sub ILLdictionary_Save()
End Sub

Private Sub ILLdictionary_Close()
End Sub
