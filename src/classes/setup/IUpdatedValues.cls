VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUpdatedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of updated values class"

'@Interface
'@Folder("Setup")
'@ModuleDescription("Interface describing the update watcher responsible for registry maintenance")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IUpdatedValues
'@description
'Defines the public contract for the UpdatedValues class, exposing methods
'to register, query, and maintain a column-level update watcher backed by
'a registry of ListObjects on a dedicated worksheet.


'@section Column registration
'===============================================================================

'@jump:add-columns
'@sub-title Register or refresh watched columns for the supplied ListObject.
'@details
'Scans the source ListObject header row for metadata tags and synchronises
'the registry table so named ranges remain aligned with the current layout.
'@param Lo ListObject. Table providing the column metadata.
'@export
Public Sub AddColumns(ByVal Lo As ListObject): End Sub

'@jump:add-sheet
'@sub-title Register tagged columns for every ListObject on the supplied worksheet.
'@param targetSheet Worksheet. Worksheet providing ListObjects to register.
'@export
Public Sub AddSheet(ByVal targetSheet As Worksheet): End Sub

'@jump:remove-lo
'@sub-title Remove registry entries associated with the provided ListObject.
'@param target ListObject. Table whose entries should be removed.
'@export
Public Sub RemoveLo(ByVal target As ListObject): End Sub


'@section Status tracking
'===============================================================================

'@jump:is-updated
'@prop-title Determine whether the tracked column has been marked as updated.
'@details
'Returns True when the column has been flagged as updated since the last
'reset. Searches across all registry ListObjects for a matching entry.
'@param colName String. Column name to inspect.
'@return Boolean. True when the column is flagged as updated.
'@export
Public Property Get IsUpdated(ByVal colName As String) As Boolean: End Property

'@jump:check-update
'@sub-title Evaluate a range change against the registered watchers and flag updates.
'@details
'Compares the changed range with named ranges tracked in the registry and
'marks the first matching entry as updated. The tagReference parameter
'may be a String, ListObject, or Worksheet used to scope the search.
'@param tagReference Variant. Identifier selecting the registry scope.
'@param Target Range. The changed cells to evaluate.
'@export
Public Sub CheckUpdate(ByVal tagReference As Variant, ByVal Target As Range): End Sub


'@section Clean-up
'===============================================================================

'@jump:clear-up
'@sub-title Reset all tracked columns to the default (not updated) status.
'@export
Public Sub ClearUp(): End Sub

'@jump:delete-up
'@sub-title Remove the watcher registry and associated workbook named ranges.
'@export
Public Sub DeleteUp(): End Sub

'@jump:switch-tags-to-no
'@sub-title Force every registry status flag to the default value.
'@export
Public Sub SwitchTagsToNo(): End Sub

'@jump:switch-tags-to-yes
'@sub-title Force every registry status flag to the updated value.
'@export
Public Sub SwitchTagsToYes(): End Sub
