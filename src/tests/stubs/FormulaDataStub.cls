VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaDataStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests.Stubs")
'@ModuleDescription("Lightweight IFormulaData implementation for unit tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IFormulaData

Private allowedChars As String
Private allowedFunctions As Collection
Private groupFunctions As Collection

Private Sub Class_Initialize()
    allowedChars = "+-*/^=,() "
    Set allowedFunctions = New Collection
    Set groupFunctions = New Collection
    On Error Resume Next
        allowedFunctions.Add "SUM", "SUM"
        allowedFunctions.Add "IF", "IF"
    On Error GoTo 0
End Sub

Public Sub AllowCharacter(ByVal characterToken As String)
    If InStr(allowedChars, characterToken) = 0 Then
        allowedChars = allowedChars & characterToken
    End If
End Sub

Public Sub AllowFunction(ByVal functionName As String)
    On Error Resume Next
        allowedFunctions.Add UCase$(functionName), UCase$(functionName)
    On Error GoTo 0
End Sub

Public Sub AllowGroupFunction(ByVal functionName As String, _
                              ByVal aggregatorName As String, _
                              Optional ByVal useNative As Boolean = False)
    Dim mapping As Variant
    Dim normalizedFunction As String
    Dim normalizedAggregator As String

    normalizedFunction = UCase$(Trim$(functionName))
    normalizedAggregator = UCase$(Trim$(aggregatorName))
    If LenB(normalizedFunction) = 0 Then Exit Sub
    If LenB(normalizedAggregator) = 0 Then Exit Sub

    RemoveGroupFunction normalizedFunction
    mapping = Array(normalizedFunction, normalizedAggregator, CBool(useNative))
    groupFunctions.Add mapping

    AllowFunction normalizedFunction
    AllowFunction normalizedAggregator
End Sub

Private Function IFormulaData_SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    If Len(specialChar) = 0 Then Exit Function
    IFormulaData_SpecialCharacterIncludes = InStr(allowedChars, Left$(specialChar, 1)) > 0
End Function

Private Function IFormulaData_ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
    Dim key As String
    key = UCase$(Trim$(ExcelFormula))
    On Error Resume Next
        IFormulaData_ExcelFormulasIncludes = Not IsEmpty(allowedFunctions.Item(key))
    On Error GoTo 0
End Function

Private Function IFormulaData_IsGroupFunction(ByVal functionName As String) As Boolean
    IFormulaData_IsGroupFunction = (GroupFunctionIndex(UCase$(Trim$(functionName))) > 0)
End Function

Private Function IFormulaData_GroupAggregator(ByVal functionName As String) As String
    Dim idx As Long
    idx = GroupFunctionIndex(UCase$(Trim$(functionName)))
    If idx > 0 Then
        IFormulaData_GroupAggregator = CStr(groupFunctions.Item(idx)(1))
    End If
End Function

Private Function IFormulaData_GroupUsesNativeFunction(ByVal functionName As String) As Boolean
    Dim idx As Long
    idx = GroupFunctionIndex(UCase$(Trim$(functionName)))
    If idx > 0 Then
        IFormulaData_GroupUsesNativeFunction = CBool(groupFunctions.Item(idx)(2))
    End If
End Function

Private Function GroupFunctionIndex(ByVal normalizedFunction As String) As Long
    Dim idx As Long
    Dim entry As Variant

    If groupFunctions Is Nothing Then Exit Function
    If LenB(normalizedFunction) = 0 Then Exit Function

    For idx = 1 To groupFunctions.Count
        entry = groupFunctions.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), normalizedFunction, vbBinaryCompare) = 0 Then
                GroupFunctionIndex = idx
                Exit Function
            End If
        End If
    Next idx
End Function

Private Sub RemoveGroupFunction(ByVal normalizedFunction As String)
    Dim idx As Long
    Dim entry As Variant

    If groupFunctions Is Nothing Then Exit Sub

    For idx = groupFunctions.Count To 1 Step -1
        entry = groupFunctions.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), normalizedFunction, vbBinaryCompare) = 0 Then
                groupFunctions.Remove idx
            End If
        End If
    Next idx
End Sub
