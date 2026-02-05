VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableLegacyAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Bridges legacy cross table workflows to the new helper services"

'@Folder("Implement")
'@ModuleDescription("Adapter exposing helper-based layout, registration, and write operations")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTableLegacyAdapter
    services As CrossTableServices
    planner As CrossTableLayoutPlanner
    registrar As CrossTableRangeRegistrar
    targetSheet As Worksheet
    layoutSummary As CrossTableLayoutSummary
    planTracker As CrossTablePerformanceTracker
    writeTracker As CrossTablePerformanceTracker
End Type

Private this As TCrossTableLegacyAdapter
Private Const CLASSNAME As String = "CrossTableLegacyAdapter"

'@section Factory
'===============================================================================

Public Function Create(ByVal specs As ITablesSpecs, ByVal targetSheet As Worksheet) As CrossTableLegacyAdapter
    Dim adapter As CrossTableLegacyAdapter

    If specs Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Table specifications are required"
    End If

    If targetSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Target worksheet is required"
    End If

    Set adapter = New CrossTableLegacyAdapter
    adapter.Initialise specs, targetSheet
    Set Create = adapter
End Function

Private Sub Initialise(ByVal specs As ITablesSpecs, ByVal targetSheet As Worksheet)
    Set this.services = New CrossTableServices
    Set this.planner = this.services.BuildLayoutPlanner(specs, targetSheet)
    Set this.registrar = this.services.BuildRangeRegistrar(targetSheet)
    Set this.targetSheet = targetSheet
    Set this.planTracker = New CrossTablePerformanceTracker
    Set this.writeTracker = New CrossTablePerformanceTracker
End Sub

'@section Public API
'===============================================================================

Public Function PlanAndRegister(ByVal namePrefix As String) As CrossTableLayoutSummary
    Dim summary As CrossTableLayoutSummary

    If LenB(Trim$(namePrefix)) = 0 Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Name prefix cannot be empty"
    End If

    this.planTracker.Reset
    this.planTracker.BeginMeasurement
    Set summary = this.planner.ComputeLayout()
    Set this.layoutSummary = summary

    this.registrar.RegisterLayoutName namePrefix & "_STARTROW", summary
    this.registrar.RegisterLayoutName namePrefix & "_TABLE", summary
    this.planTracker.EndMeasurement

    Set PlanAndRegister = summary
End Function

Public Sub WriteContent(headerMatrix As Variant, bodyMatrix As Variant)
    Dim buffer As CrossTableWriteBuffer

    EnsureSummaryAvailable

    Set buffer = this.services.BuildWriteBuffer(this.layoutSummary.StartRow, _
                                                this.layoutSummary.StartColumn)
    buffer.AssignHeader headerMatrix
    buffer.AssignBody bodyMatrix
    this.writeTracker.Reset
    this.writeTracker.BeginMeasurement
    buffer.Commit this.targetSheet
    this.writeTracker.EndMeasurement
End Sub

Public Function BuildPercentageFormula(ByVal numerator As Range, ByVal denominator As Range) As String
    Dim builder As CrossTableFormulaBuilder

    Set builder = this.services.BuildFormulaBuilder(this.targetSheet)
    BuildPercentageFormula = builder.BuildPercentage(numerator, denominator)
End Function

Public Function BuildConditionalFormula(ByVal conditionRange As Range, ByVal formulaText As String) As String
    Dim builder As CrossTableFormulaBuilder

    Set builder = this.services.BuildFormulaBuilder(this.targetSheet)
    BuildConditionalFormula = builder.BuildConditional(conditionRange, formulaText)
End Function

Public Function NewConditionSet() As CrossTableConditionSet
    Set NewConditionSet = this.services.NewConditionSet
End Function

Public Property Get LayoutSummary() As CrossTableLayoutSummary
    Set LayoutSummary = this.layoutSummary
End Property

Public Property Get PlanDurationMilliseconds() As Double
    PlanDurationMilliseconds = this.planTracker.DurationMilliseconds
End Property

Public Property Get WriteDurationMilliseconds() As Double
    WriteDurationMilliseconds = this.writeTracker.DurationMilliseconds
End Property

'@section Internal helpers
'===============================================================================

Private Sub EnsureSummaryAvailable()
    If this.layoutSummary Is Nothing Then
        Err.Raise ProjectError.ErrorUnexpectedState, CLASSNAME, _
                  "PlanAndRegister must be executed before writing content"
    End If
End Sub

End Class
