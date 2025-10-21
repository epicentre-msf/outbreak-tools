VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisVariablesStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("ILLVariables double exposing configurable control and type metadata")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ILLVariables

Private Type TVariableInfo
    control As String
    variableType As String
End Type

Private Type TVariablesStubState
    entries As BetterArray
End Type

Private this As TVariablesStubState

Private Sub Class_Initialize()
    Set this.entries = New BetterArray
    this.entries.LowerBound = 1
End Sub

Private Sub StoreEntry(ByVal variableName As String, ByVal info As TVariableInfo)
    Dim idx As Long
    Dim entry As Variant

    If this.entries.Length = 0 Then GoTo AppendEntry

    For idx = this.entries.UpperBound To this.entries.LowerBound Step -1
        If Not IsEmpty(this.entries.Item(idx)) Then
            entry = this.entries.Item(idx)
            If IsArray(entry) Then
                If StrComp(CStr(entry(1)), variableName, vbBinaryCompare) = 0 Then
                    this.entries.Remove idx
                End If
            End If
        End If
    Next idx

AppendEntry:
    this.entries.Push Array(variableName, info)
End Sub

Private Function TryGetEntry(ByVal variableName As String, ByRef info As TVariableInfo, Optional ByVal matchCase As Boolean = True) As Boolean
    Dim idx As Long
    Dim entry As Variant
    Dim candidateName As String
    Dim comparison As VbCompareMethod

    TryGetEntry = False
    comparison = IIf(matchCase, vbBinaryCompare, vbTextCompare)

    If this.entries.Length = 0 Then Exit Function

    For idx = this.entries.LowerBound To this.entries.UpperBound
        If Not IsEmpty(this.entries.Item(idx)) Then
            entry = this.entries.Item(idx)
            If IsArray(entry) Then
                candidateName = CStr(entry(1))
                If StrComp(candidateName, variableName, comparison) = 0 Then
                    info = entry(2)
                    TryGetEntry = True
                    Exit Function
                End If
            End If
        End If
    Next idx
End Function

'Register a variable with its control and type metadata.
Public Sub AddVariable(ByVal variableName As String, _
                       ByVal controlValue As String, _
                       ByVal variableType As String)
    Dim info As TVariableInfo

    info.control = controlValue
    info.variableType = variableType

    StoreEntry variableName, info
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get ILLVariables_Dictionary() As ILLdictionary
    Set ILLVariables_Dictionary = Nothing
End Property

Private Function ILLVariables_Value(ByVal colName As String, ByVal varName As String) As String
    Dim info As TVariableInfo

    If Not TryGetEntry(varName, info) Then Exit Function

    Select Case LCase$(Trim$(colName))
        Case "control"
            ILLVariables_Value = info.control
        Case "variable type"
            ILLVariables_Value = info.variableType
        Case Else
            ILLVariables_Value = vbNullString
    End Select
End Function

Private Function ILLVariables_Index(ByVal varName As String) As Long
    ILLVariables_Index = 0
End Function

Private Function ILLVariables_Contains(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Boolean
    Dim info As TVariableInfo
    ILLVariables_Contains = TryGetEntry(varName, info, matchCase)
End Function

Private Sub ILLVariables_SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)
End Sub

Private Property Get ILLVariables_CellRange(ByVal colName As String, ByVal varName As String) As Range
    Set ILLVariables_CellRange = Nothing
End Property

Private Function ILLVariables_VariableNames() As BetterArray
    Dim names As BetterArray
    Dim idx As Long
    Dim entry As Variant

    Set names = New BetterArray
    names.LowerBound = 1

    If this.entries.Length = 0 Then
        Set ILLVariables_VariableNames = names
        Exit Function
    End If

    For idx = this.entries.LowerBound To this.entries.UpperBound
        entry = this.entries.Item(idx)
        If IsArray(entry) Then names.Push CStr(entry(1))
    Next idx

    Set ILLVariables_VariableNames = names
End Function

Private Function ILLVariables_SheetName(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    ILLVariables_SheetName = vbNullString
End Function

Private Function ILLVariables_TableName(ByVal varName As String) As String
    ILLVariables_TableName = vbNullString
End Function

Private Function ILLVariables_ControlType(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    Dim info As TVariableInfo

    If Not TryGetEntry(varName, info, matchCase) Then Exit Function
    ILLVariables_ControlType = info.control
End Function

Private Sub ILLVariables_InvalidateCaches()
End Sub

Private Function ILLVariables_HasCheckings() As Boolean
    ILLVariables_HasCheckings = False
End Function

Private Property Get ILLVariables_CheckingValues() As IChecking
    Set ILLVariables_CheckingValues = Nothing
End Property
