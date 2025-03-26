VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Analysis Output"

'@IgnoreModule UseMeaningfulName, HungarianNotation
'@Folder("Analysis")
'@ModuleDescription("Interface of Analysis Output")
'@Interface

Option Explicit
Public Enum AnalysisScope
  AnalysisScopeTimeSeries = 1 'A time series analysis
  AnalysisScopeSpatial = 2 'A spatial analysis
  AnalysisScopeNormal = 3 'A standard univariate / bivariate analysis
  AnalysisScopeSpatioTemporal = 4 'spatio-Temporal analysis
End Enum

'Exposed methods
Public Sub WriteAnalysis(): End Sub
