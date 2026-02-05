VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableFormulaBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Assembles validated formula strings for cross tables"

'@Folder("Implement")
'@ModuleDescription("Build Excel formulas for cross tables without mutating worksheets")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTableFormulaBuilder
    evaluationSheet As Worksheet
End Type

Private this As TCrossTableFormulaBuilder
Private Const CLASSNAME As String = "CrossTableFormulaBuilder"

'@section Factory
'===============================================================================

Public Function Create(ByVal evaluationSheet As Worksheet) As CrossTableFormulaBuilder
    Dim builder As CrossTableFormulaBuilder

    If evaluationSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Evaluation worksheet is required"
    End If

    Set builder = New CrossTableFormulaBuilder
    Set builder.EvaluationSheet = evaluationSheet
    Set Create = builder
End Function

Public Property Set EvaluationSheet(ByVal sh As Worksheet)
    Set this.evaluationSheet = sh
End Property

Public Property Get EvaluationSheet() As Worksheet
    Set EvaluationSheet = this.evaluationSheet
End Property

'@section Public API
'===============================================================================

Public Function BuildPercentage(ByVal numerator As Range, ByVal denominator As Range, _
                                Optional ByVal lockRow As Boolean = True) As String
    Dim numeratorAddress As String
    Dim denominatorAddress As String
    Dim expression As String

    ValidateRange numerator, "numerator"
    ValidateRange denominator, "denominator"

    numeratorAddress = numerator.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    denominatorAddress = denominator.Address(RowAbsolute:=lockRow, ColumnAbsolute:=True)
    'Guard the denominator by locking the column and optionally the row.
    expression = numeratorAddress & "/" & denominatorAddress

    BuildPercentage = "IF(ISERR(" & expression & "),"""""""," & expression & ")"
End Function

Public Function BuildConditional(ByVal conditionRange As Range, ByVal formulaText As String, _
                                 Optional ByVal lockRow As Boolean = False) As String
    Dim conditionAddress As String
    Dim trimmedFormula As String

    ValidateRange conditionRange, "conditionRange"

    trimmedFormula = Trim$(formulaText)
    If LenB(trimmedFormula) = 0 Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Formula text cannot be empty"
    End If

    conditionAddress = conditionRange.Address(RowAbsolute:=lockRow, ColumnAbsolute:=False)
    BuildConditional = "IF(" & conditionAddress & "="""""",""""""," & trimmedFormula & ")"
End Function

Public Function Validate(ByVal formulaText As String) As Boolean
    Dim evaluationResult As Variant

    If this.evaluationSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Evaluation worksheet is required"
    End If

    If LenB(Trim$(formulaText)) = 0 Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Formula text cannot be empty"
    End If

    'Use Worksheet.Evaluate to validate syntax without mutating workbook state.
    On Error GoTo EvaluationFailed
        evaluationResult = this.evaluationSheet.Evaluate("=" & formulaText)
        If IsError(evaluationResult) Then
            Validate = False
        Else
            Validate = True
        End If
        Exit Function

EvaluationFailed:
    Validate = False
    Err.Clear
End Function

Private Sub ValidateRange(ByVal rng As Range, ByVal argumentName As String)
    If rng Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, argumentName & " range is required"
    End If
End Sub

End Class
