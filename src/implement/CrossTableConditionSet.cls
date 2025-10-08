VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableConditionSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Tracks unique condition variables for cross table formulas"

'@Folder("Implement")
'@ModuleDescription("Collects and exposes ordered condition variables/tests for cross tables")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTableConditionSet
    variables As BetterArray
    tests As BetterArray
End Type

Private this As TCrossTableConditionSet

Public Sub Initialise()
    Set this.variables = New BetterArray
    this.variables.LowerBound = 1
    Set this.tests = New BetterArray
    this.tests.LowerBound = 1
End Sub

Public Sub Add(ByVal variableName As String, ByVal testName As String)
    Dim trimmedVariable As String
    Dim trimmedTest As String
    Dim idx As Long

    trimmedVariable = Trim$(variableName)
    trimmedTest = Trim$(testName)

    If LenB(trimmedVariable) = 0 Then
        Err.Raise ProjectError.InvalidArgument, "CrossTableConditionSet", "Variable name cannot be empty"
    End If

    EnsureInitialised

    'Merge duplicate variable references while keeping the latest test definition.
    If this.variables.Length > 0 Then
        For idx = this.variables.LowerBound To this.variables.UpperBound
            If StrComp(CStr(this.variables.Item(idx)), trimmedVariable, vbTextCompare) = 0 Then
                this.tests.Item(idx) = trimmedTest
                Exit Sub
            End If
        Next idx
    End If

    this.variables.Push trimmedVariable
    this.tests.Push trimmedTest
End Sub

Public Function Count() As Long
    If this.variables Is Nothing Then Exit Function
    Count = this.variables.Length
End Function

Public Function VariableAt(ByVal index As Long) As String
    ValidateIndex index
    VariableAt = CStr(this.variables.Item(index))
End Function

Public Function TestAt(ByVal index As Long) As String
    ValidateIndex index
    TestAt = CStr(this.tests.Item(index))
End Function

Private Sub EnsureInitialised()
    If this.variables Is Nothing Then Initialise
End Sub

Private Sub ValidateIndex(ByVal index As Long)
    If this.variables Is Nothing Then
        Err.Raise ProjectError.InvalidArgument, "CrossTableConditionSet", "Condition set is empty"
    End If

    If index < this.variables.LowerBound Or index > this.variables.UpperBound Then
        Err.Raise ProjectError.InvalidArgument, "CrossTableConditionSet", "Index is out of range"
    End If
End Sub

End Class
