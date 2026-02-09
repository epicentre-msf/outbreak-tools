VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for analysis output orchestration"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for analysis output orchestration")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

' Scope for analysis worksheet selection
Public Enum AnalysisScope
    AnalysisScopeTimeSeries = 1
    AnalysisScopeSpatial = 2
    AnalysisScopeNormal = 3
    AnalysisScopeSpatioTemporal = 4
End Enum

' @description Write all analysis tables, formulas, and graphs to the output worksheets
Public Sub WriteAnalysis(): End Sub
