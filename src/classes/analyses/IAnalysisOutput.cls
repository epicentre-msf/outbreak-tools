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

'@class IAnalysisOutput
'@description
'Defines the public contract for the AnalysisOutput class. Exposes a single
'method that orchestrates the full analysis output pipeline: building
'cross-tables, writing formulas, registering graphs, and rendering charts
'across all analysis scopes.


' Scope for analysis worksheet selection
Public Enum AnalysisScope
    AnalysisScopeTimeSeries = 1
    AnalysisScopeSpatial = 2
    AnalysisScopeNormal = 3
    AnalysisScopeSpatioTemporal = 4
    AnalysisScopeTimeSeriesTablesOnly = 5
End Enum


'@section Methods
'===============================================================================

'@jump:write-analysis
'@sub-title Write all analysis tables, formulas, and graphs to the output
'@details
'Orchestrates the complete analysis output pipeline. Iterates through
'each analysis scope (normal, time series, spatial, spatio-temporal),
'builds cross-tables from the table specifications, writes formulas via
'CrossTableFormula, registers table and graph metadata in AnaTabIds,
'and renders all tracked graphs on the output worksheets.
'@export
Public Sub WriteAnalysis(Optional ByVal part As Byte = 0): End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access for analysis output results.

'@jump:has-checkings
'@prop-title Whether the analysis output has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
Public Property Get HasCheckings() As Boolean
End Property

'@jump:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return IChecking. The checking instance, or Nothing.
Public Property Get CheckingValues() As IChecking
End Property
