VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTableWriterContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Provides the concrete operations required by AnalysisTableWriter to materialise plan items")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Invoked before any table writing occurs so the context can prepare the target sheet.
Public Sub BeginPlan(ByVal outputSheet As Worksheet)
End Sub

'Write a single table described by the supplied plan item into the target sheet.
Public Sub WriteTable(ByVal planItem As IAnalysisTablePlanItem, ByVal outputSheet As Worksheet)
End Sub

'Apply navigation metadata (section/header labels) once tables have been written.
Public Sub ApplyNavigation(ByVal sectionLabels As BetterArray, _
                           ByVal headerLabels As BetterArray, _
                           ByVal outputSheet As Worksheet)
End Sub

'Invoked after all tables and navigation metadata have been processed.
Public Sub CompletePlan(ByVal outputSheet As Worksheet)
End Sub
