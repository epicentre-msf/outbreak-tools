VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisGraphExecutionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Represents the graph output generated for a specific plan item")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the originating graph plan item.
Public Property Get PlanItem() As IAnalysisGraphPlanItem
End Property

'Return the chart specifications produced for the plan item.
Public Property Get Charts() As BetterArray
End Property

'Return the number of charts produced for the plan item.
Public Property Get ChartCount() As Long
End Property
