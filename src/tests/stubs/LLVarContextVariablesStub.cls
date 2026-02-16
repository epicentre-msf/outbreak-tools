VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarContextVariablesStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("Test double implementing a lightweight ILLVariables store")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVariables

Private Type TVariablesStub
    rows As BetterArray
    rangeMap As Collection
End Type

Private this As TVariablesStub

'@section Helpers
'===============================================================================
Public Sub AddValue(ByVal varName As String, ByVal colName As String, ByVal value As String)
    EnsureStore
    this.rows.Push Array(Normalise(varName), Normalise(colName), value)
End Sub

Public Sub SetCellRange(ByVal varName As String, ByVal colName As String, ByVal target As Range)
    Dim key As String

    EnsureStore
    key = BuildKey(varName, colName)

    On Error Resume Next
        this.rangeMap.Remove key
    On Error GoTo 0

    If target Is Nothing Then Exit Sub

    On Error Resume Next
        this.rangeMap.Add target, key
    On Error GoTo 0
End Sub

Private Sub EnsureStore()
    If this.rows Is Nothing Then
        Set this.rows = New BetterArray
        this.rows.LowerBound = 1
    End If

    If this.rangeMap Is Nothing Then
        Set this.rangeMap = New Collection
    End If
End Sub

Private Function BuildKey(ByVal varName As String, ByVal colName As String) As String
    BuildKey = Normalise(varName) & "|" & Normalise(colName)
End Function

Private Function Normalise(ByVal textValue As String) As String
    Normalise = LCase$(Trim$(textValue))
End Function

Private Function FindEntry(ByVal varName As String, ByVal colName As String) As Variant
    Dim idx As Long
    Dim candidate As Variant
    Dim targetVar As String
    Dim targetCol As String

    targetVar = Normalise(varName)
    targetCol = Normalise(colName)

    If this.rows Is Nothing Then Exit Function
    If this.rows.Length = 0 Then Exit Function

    For idx = this.rows.UpperBound To this.rows.LowerBound Step -1
        candidate = this.rows.Item(idx)
        If IsArray(candidate) Then
            If StrComp(CStr(candidate(0)), targetVar, vbBinaryCompare) = 0 Then
                If StrComp(CStr(candidate(1)), targetCol, vbBinaryCompare) = 0 Then
                    FindEntry = candidate
                    Exit Function
                End If
            End If
        End If
    Next idx
End Function

Private Function LookupRange(ByVal varName As String, ByVal colName As String) As Range
    Dim key As String

    If this.rangeMap Is Nothing Then Exit Function

    key = BuildKey(varName, colName)

    On Error Resume Next
        Set LookupRange = this.rangeMap.Item(key)
    On Error GoTo 0
End Function

Private Function AllVariableNames() As BetterArray
    Dim uniqueMap As Collection
    Dim idx As Long
    Dim candidate As Variant
    Dim result As BetterArray

    Set result = New BetterArray
    result.LowerBound = 1

    If this.rows Is Nothing Then
        Set AllVariableNames = result
        Exit Function
    End If

    If this.rows.Length = 0 Then
        Set AllVariableNames = result
        Exit Function
    End If

    Set uniqueMap = New Collection

    For idx = this.rows.LowerBound To this.rows.UpperBound
        candidate = this.rows.Item(idx)
        If IsArray(candidate) Then
            On Error Resume Next
                uniqueMap.Add candidate(0), CStr(candidate(0))
            On Error GoTo 0
        End If
    Next idx

    For idx = 1 To uniqueMap.Count
        result.Push uniqueMap(idx)
    Next idx

    Set AllVariableNames = result
End Function

Private Function LookupValue(ByVal varName As String, ByVal colName As String) As String
    Dim entry As Variant

    entry = FindEntry(varName, colName)
    If IsArray(entry) Then
        LookupValue = CStr(entry(2))
    Else
        LookupValue = vbNullString
    End If
End Function

Private Function LookupIndex(ByVal varName As String) As Long
    Dim indexValue As String

    indexValue = LookupValue(varName, "column index")
    If LenB(indexValue) = 0 Then Exit Function

    On Error Resume Next
        LookupIndex = CLng(indexValue)
    On Error GoTo 0
End Function

Private Sub ThrowNotImplemented(ByVal memberName As String)
    Err.Raise vbObjectError + 513, "LLVarContextVariablesStub." & memberName, _
              "Member not implemented in stub"
End Sub

'@section ILLVariables implementation
'===============================================================================
Private Property Get ILLVariables_Dictionary() As ILLdictionary
    Set ILLVariables_Dictionary = Nothing
End Property

Private Function ILLVariables_Value(ByVal colName As String, ByVal varName As String) As String
    ILLVariables_Value = LookupValue(varName, colName)
End Function

Private Function ILLVariables_Index(ByVal varName As String) As Long
    ILLVariables_Index = LookupIndex(varName)
End Function

Private Function ILLVariables_Contains(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Boolean
    Dim names As BetterArray
    Dim idx As Long
    Dim candidate As String
    Set names = AllVariableNames()

    For idx = names.LowerBound To names.UpperBound
        candidate = CStr(names.Item(idx))
        If matchCase Then
            If StrComp(candidate, Normalise(varName), vbBinaryCompare) = 0 Then
                ILLVariables_Contains = True
                Exit Function
            End If
        Else
            If StrComp(candidate, Normalise(varName), vbTextCompare) = 0 Then
                ILLVariables_Contains = True
                Exit Function
            End If
        End If
    Next idx

    ILLVariables_Contains = False
End Function

Private Sub ILLVariables_SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)
    AddValue varName, colName, newValue
End Sub

Private Property Get ILLVariables_CellRange(ByVal colName As String, ByVal varName As String) As Range
    Set ILLVariables_CellRange = LookupRange(varName, colName)
End Property

Private Function ILLVariables_VariableNames() As BetterArray
    Set ILLVariables_VariableNames = AllVariableNames()
End Function

Private Function ILLVariables_SheetName(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    ILLVariables_SheetName = LookupValue(varName, "sheet name")
End Function

Private Function ILLVariables_TableName(ByVal varName As String) As String
    ILLVariables_TableName = LookupValue(varName, "table name")
End Function

Private Function ILLVariables_ControlType(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    ILLVariables_ControlType = LookupValue(varName, "control")
End Function

Private Sub ILLVariables_InvalidateCaches()
    'No caching required for stub
End Sub

Private Function ILLVariables_HasCheckings() As Boolean
    ILLVariables_HasCheckings = False
End Function

Private Property Get ILLVariables_CheckingValues() As IChecking
    Set ILLVariables_CheckingValues = Nothing
End Property
