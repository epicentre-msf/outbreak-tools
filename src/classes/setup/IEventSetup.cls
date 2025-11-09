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

'@section Row operations
'===============================================================================

'@description Manage setup worksheet rows using cached helpers.
'@param sheetName Worksheet name to process (Dictionary, Choices, Analysis, Exports, Translations).
'@param del Optional Boolean. When True the action removes rows instead of adding them.
Public Sub ManageRows(ByVal sheetName As String, Optional ByVal del As Boolean = False)
End Sub

'@description Insert rows relative to a selection for supported setup sheets.
'@param sheetName Worksheet name identifying which helper should process the insertion.
'@param targetCell Range describing the worksheet selection to mirror.
'@param insertShift Optional Boolean controlling whether worksheet rows should be inserted to preserve stacked tables.
Public Sub InsertRows(ByVal sheetName As String, _
                      ByVal targetCell As Range, _
                      Optional ByVal insertShift As Boolean = False)
End Sub

'@description Delete rows relative to a selection for supported setup sheets.
'@param sheetName Worksheet name identifying which helper should process the deletion.
'@param targetCell Range describing the worksheet selection to remove.
'@param forceShift Optional Boolean controlling whether worksheet rows should be deleted (Analysis always forces shift).
Public Sub DeleteRows(ByVal sheetName As String, _
                      ByVal targetCell As Range, _
                      Optional ByVal forceShift As Boolean = False)
End Sub

'@description Sort setup worksheets leveraging cached managers.
'@param sheetName Worksheet name to sort.
Public Sub SortTables(ByVal sheetName As String)
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
