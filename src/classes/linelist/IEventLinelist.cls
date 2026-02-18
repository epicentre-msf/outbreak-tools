VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEventLinelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing workbook-level linelist event orchestration"

'@Interface
'@Folder("Linelist Events")
'@ModuleDescription("Defines the contract implemented by EventLinelist to service workbook-level events in the output linelist")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IEventLinelist
'@description
'Defines the public contract for the EventLinelist class, exposing methods
'to handle workbook-level events in the output linelist workbook: sheet
'change, deactivate, selection change, double-click, filter table updates,
'and list-auto dropdown maintenance.
'@depends ILLTranslation, ITranslationObject, IHiddenNames, IDropdownLists, ILLGeo, ILLdictionary, ILLVariables, IPasswords, ICustomTable


'@section Lifecycle
'===============================================================================

'@sub-title Configure the service with its host workbook context.
'@details
'Binds the service to a host workbook and resets all cached helpers. Must be
'called once before any event handler.
'@param hostBook Workbook. The linelist workbook providing data and infrastructure.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
End Sub


'@section Workbook events
'===============================================================================

'@sub-title React to workbook open notifications.
'@details
'Sets keyboard shortcuts, disables stale value formatting, and applies
'the initial workbook configuration (calculation mode, etc.).
'@export
Public Sub OnWorkbookOpen()
End Sub

'@sub-title Handle worksheet change events fired by Excel.
'@details
'Inspects the sheet type tag (cell C1) and dispatches to the appropriate
'handler: HList value change (geo cascades, editable labels, list-auto,
'go-to-section, header protection, multiple choices), VList value change,
'or analysis value change (go-to-section, spatial table dispatch).
'@param sh Worksheet. The worksheet raising the change event.
'@param target Range. The edited cells.
'@export
Public Sub OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
End Sub

'@sub-title Handle worksheet deactivation events.
'@details
'When the deactivated sheet had list-auto updates pending (flagged via
'RNG_UpdateListAuto hidden name), triggers the list-auto refresh and
'clears the flag.
'@param prevSheetName String. Name of the sheet that was deactivated.
'@export
Public Sub OnSheetDeactivate(ByVal prevSheetName As String)
End Sub

'@sub-title Handle selection change events on HList sheets.
'@details
'Updates geo dropdown validation lists based on the selected cell's geo
'control tag. When the selection is on the last row of the ListObject,
'automatically adds rows.
'@param sh Worksheet. The worksheet where the selection changed.
'@param target Range. The newly selected cell(s).
'@export
Public Sub OnSelectionChange(ByVal sh As Worksheet, ByVal target As Range)
End Sub

'@sub-title Handle double-click events on analysis sheets.
'@details
'When the double-clicked cell is on a spatio-temporal analysis sheet and
'falls within an INPUTSPTGEO or INPUTSPTHF named range, opens the geo
'form for the corresponding type.
'@param sh Worksheet. The worksheet where the double-click occurred.
'@param target Range. The double-clicked cell.
'@export
Public Sub OnDoubleClick(ByVal sh As Worksheet, ByVal target As Range)
End Sub


'@section Auxiliary operations
'===============================================================================

'@sub-title Update filter tables used by analysis sheets.
'@details
'Copies visible HList data to corresponding filtered sheets. Optionally
'recalculates the active analysis sheet afterwards.
'@param calculate Optional Boolean. When True (default), recalculates the active sheet after updating.
'@export
Public Sub UpdateFilterTables(Optional ByVal calculate As Boolean = True)
End Sub

'@sub-title Update all list-auto dropdowns across all HList sheets.
'@details
'Iterates all worksheets in the workbook, and for each HList sheet,
'scans columns tagged as list_auto_origin and refreshes their dropdown
'validation lists with unique values from the data column.
'@export
Public Sub UpdateAllListAuto()
End Sub

'@sub-title Clear cached resources so they are lazily rebuilt on demand.
'@details
'Clears all lazily-cached domain managers so they are rebuilt from
'scratch on the next access.
'@export
Public Sub ResetCaches()
End Sub
