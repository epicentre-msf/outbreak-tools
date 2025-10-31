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

'@Folder("Setup")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@ModuleDescription("Interface describing the update watcher responsible for registry maintenance")
'@Interface

Option Explicit

'@sub-title AddColumns
'@description Register or refresh watched columns for the supplied ListObject.
'@param Lo ListObject providing the column metadata.
Public Sub AddColumns(ByVal Lo As ListObject): End Sub

'@sub-title AddSheet
'@description Register tagged columns for every ListObject on the supplied worksheet.
'@param targetSheet Worksheet providing ListObjects to register.
Public Sub AddSheet(ByVal targetSheet As Worksheet): End Sub

'@sub-title RemoveLo
'@description Remove registry entries associated with the provided ListObject.
'@param target ListObject whose entries should be removed.
Public Sub RemoveLo(ByVal target As ListObject): End Sub

'@sub-title IsUpdated
'@description Determine whether the tracked column has been marked as updated since the last reset.
'@param colName String column name to inspect.
'@returns Boolean True when the column is flagged as updated.
Public Property Get IsUpdated(ByVal colName As String) As Boolean: End Property

'@sub-title CheckUpdate
'@description Evaluate the target range change against the registered watchers and flag updates.
'@param tagReference String, ListObject, or Worksheet used to select the registry scope.
'@param Target Range representing the changed cells.
Public Sub CheckUpdate(ByVal tagReference As Variant, ByVal Target As Range): End Sub

'@sub-title ClearUp
'@description Reset all tracked columns to the default (not updated) status.
Public Sub ClearUp(): End Sub

'@sub-title DeleteUp
'@description Remove the watcher registry and associated workbook named ranges.
Public Sub DeleteUp(): End Sub

'@sub-title SwitchTagsToNo
Public Sub SwitchTagsToNo(): End Sub

'@sub-title SwitchTagsToYes
Public Sub SwitchTagsToYes(): End Sub
