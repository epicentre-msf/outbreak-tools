VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEventSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing workbook-level setup event orchestration"
Option Explicit

'@Folder("Setup")
'@ModuleDescription("Defines the contract implemented by EventSetup to service workbook level events")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'@section Lifecycle
'===============================================================================

'@description Configure the service with its host workbook context.
'@param hostBook Workbook providing setup worksheets and infrastructure.
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@description React to workbook open notifications.
Public Sub OnWorkbookOpen()
End Sub

'@section Workbook Events
'===============================================================================

'@description Handle worksheet activation events.
'@param sh Worksheet that became active.
Public Sub OnSheetActivate(ByVal sh As Worksheet)
End Sub

'@description Handle worksheet change events fired by Excel.
'@param sh Worksheet raising the change event.
'@param target Range describing the edited cells.
Public Sub OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
End Sub

'@section Auxiliary Operations
'===============================================================================

'@description Recalculate analysis tables in response to structural edits.
Public Sub RecalculateAnalysis()
End Sub

'@description Refresh cached dropdowns and related analysis helpers.
'@param forceUpdate Optional Boolean forcing a full refresh regardless of update tags.
Public Sub UpdateAnalysisDropdowns(Optional ByVal forceUpdate As Boolean = False)
End Sub

'@description Reset the translation update counter stored in hidden names.
Public Sub ResetTranslationCounter()
End Sub

'@description Compose the time-series header displayed in analysis tables.
Public Function BuildTimeSeriesHeader(ByVal timeVar As String, _
                                      ByVal groupVar As String, _
                                      ByVal sumLabel As String) As String
End Function

'@description Retrieve a value from the analysis graph label table.
Public Function AnalysisGraphValue(ByVal graphTitle As String, _
                                   Optional ByVal graphCol As String = "Graph ID") As String
End Function

'@description Retrieve a value from the analysis time-series table.
Public Function AnalysisTimeSeriesValue(ByVal seriesTitle As String, _
                                        Optional ByVal tsColumn As String = "Series ID") As String
End Function

'@description Retrieve a value from the spatio-temporal specifications table.
Public Function SpatioTemporalSpecValue(ByVal sectionName As String, _
                                        Optional ByVal specColumn As String = "N geo max") As String
End Function

'@description Clear cached resources so they are lazily rebuilt on demand.
Public Sub ResetCaches()
End Sub
