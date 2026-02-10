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

'@Interface
'@Folder("Setup")
'@ModuleDescription("Defines the contract implemented by EventSetup to service workbook level events")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IEventSetup
'@description
'Defines the public contract for the EventSetup class, exposing methods
'to handle workbook-level setup events (open, activate, change), manage
'setup worksheet rows, sort tables, recalculate analysis helpers, and
'retrieve values from analysis-related ListObjects.
'@depends IPasswords, BetterArray, IUpdatedValues, IDropdownLists, ILLdictionary, ILLChoices, ILLExport, IAnalysis, ISetupTranslationsTable


'@section Lifecycle
'===============================================================================

'@jump:configure
'@sub-title Configure the service with its host workbook context.
'@param hostBook Workbook. The workbook providing setup worksheets and infrastructure.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@jump:on-workbook-open
'@sub-title React to workbook open notifications.
'@details
'Switches update watcher tags to "yes", resets the translation counter,
'disables stale value formatting, and applies the protection matrix.
'@export
Public Sub OnWorkbookOpen()
End Sub


'@section Workbook events
'===============================================================================

'@jump:on-sheet-activate
'@sub-title Handle worksheet activation events.
'@details
'Invalidates relevant ribbon controls and refreshes analysis dropdowns
'when the Analysis sheet becomes active.
'@param sh Worksheet. The worksheet that became active.
'@export
Public Sub OnSheetActivate(ByVal sh As Worksheet)
End Sub

'@jump:on-sheet-change
'@sub-title Handle worksheet change events fired by Excel.
'@details
'Forwards the change to the update watcher, processes choices label
'recalculation when the Choices sheet changes, and manages analysis
'dropdowns and calculations when the Analysis sheet changes.
'@param sh Worksheet. The worksheet raising the change event.
'@param target Range. The edited cells.
'@export
Public Sub OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
End Sub


'@section Auxiliary operations
'===============================================================================

'@jump:recalculate-analysis
'@sub-title Recalculate analysis tables in response to structural edits.
'@export
Public Sub RecalculateAnalysis()
End Sub

'@jump:update-analysis-dropdowns
'@sub-title Refresh cached dropdowns and related analysis helpers.
'@param forceUpdate Optional Boolean. When True, forces a full refresh regardless of update tags. Defaults to False.
'@export
Public Sub UpdateAnalysisDropdowns(Optional ByVal forceUpdate As Boolean = False)
End Sub

'@jump:reset-translation-counter
'@sub-title Reset the translation update counter stored in hidden names.
'@export
Public Sub ResetTranslationCounter()
End Sub


'@section Row operations
'===============================================================================

'@jump:manage-rows
'@sub-title Manage setup worksheet rows using cached helpers.
'@details
'Delegates to the appropriate domain manager (dictionary, choices, analysis,
'exports) to add or remove rows, handling sheet protection around each operation.
'@param sheetName String. Worksheet name to process.
'@param del Optional Boolean. When True, removes rows instead of adding them. Defaults to False.
'@export
Public Sub ManageRows(ByVal sheetName As String, Optional ByVal del As Boolean = False)
End Sub

'@jump:insert-rows
'@sub-title Insert rows relative to a selection for supported setup sheets.
'@param sheetName String. Worksheet name identifying which helper should process the insertion.
'@param targetCell Range. The worksheet selection to mirror.
'@param insertShift Optional Boolean. Controls whether worksheet rows should be inserted to preserve stacked tables. Defaults to False.
'@export
Public Sub InsertRows(ByVal sheetName As String, _
                      ByVal targetCell As Range, _
                      Optional ByVal insertShift As Boolean = False)
End Sub

'@jump:delete-rows
'@sub-title Delete rows relative to a selection for supported setup sheets.
'@param sheetName String. Worksheet name identifying which helper should process the deletion.
'@param targetCell Range. The worksheet selection to remove.
'@param forceShift Optional Boolean. Controls whether worksheet rows should be deleted. Defaults to False.
'@export
Public Sub DeleteRows(ByVal sheetName As String, _
                      ByVal targetCell As Range, _
                      Optional ByVal forceShift As Boolean = False)
End Sub

'@jump:sort-tables
'@sub-title Sort setup worksheets leveraging cached managers.
'@param sheetName String. Worksheet name to sort.
'@export
Public Sub SortTables(ByVal sheetName As String)
End Sub

'@jump:build-time-series-header
'@sub-title Compose the time-series header displayed in analysis tables.
'@param timeVar String. Time variable name.
'@param groupVar String. Group variable name.
'@param sumLabel String. Summary label prefix.
'@return String. The composed header text.
'@export
Public Function BuildTimeSeriesHeader(ByVal timeVar As String, _
                                      ByVal groupVar As String, _
                                      ByVal sumLabel As String) As String
End Function

'@jump:analysis-graph-value
'@sub-title Retrieve a value from the analysis graph label table.
'@param graphTitle String. Graph title used as the row key.
'@param graphCol Optional String. Column to read. Defaults to "Graph ID".
'@return String. The retrieved value, or empty.
'@export
Public Function AnalysisGraphValue(ByVal graphTitle As String, _
                                   Optional ByVal graphCol As String = "Graph ID") As String
End Function

'@jump:analysis-time-series-value
'@sub-title Retrieve a value from the analysis time-series table.
'@param seriesTitle String. Series title used as the row key.
'@param tsColumn Optional String. Column to read. Defaults to "Series ID".
'@return String. The retrieved value, or empty.
'@export
Public Function AnalysisTimeSeriesValue(ByVal seriesTitle As String, _
                                        Optional ByVal tsColumn As String = "Series ID") As String
End Function

'@jump:spatio-temporal-spec-value
'@sub-title Retrieve a value from the spatio-temporal specifications table.
'@param sectionName String. Section name used as the row key.
'@param specColumn Optional String. Column to read. Defaults to "N geo max".
'@return String. The retrieved value, or empty.
'@export
Public Function SpatioTemporalSpecValue(ByVal sectionName As String, _
                                        Optional ByVal specColumn As String = "N geo max") As String
End Function

'@jump:reset-caches
'@sub-title Clear cached resources so they are lazily rebuilt on demand.
'@export
Public Sub ResetCaches()
End Sub
