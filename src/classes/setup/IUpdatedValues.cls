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

'Exposed methods

'@method AddColumns
'@description Register or refresh watched columns for the supplied ListObject.
'@param Lo ListObject providing the column metadata.
Public Sub AddColumns(ByVal Lo As ListObject): End Sub

'@method IsUpdated
'@description Determine whether the tracked column has been marked as updated since the last reset.
'@param colName String column name to inspect.
'@returns Boolean True when the column is flagged as updated.
Public Property Get IsUpdated(ByVal colName As String) As Boolean: End Property

'@method CheckUpdate
'@description Evaluate the target range change against the registered watchers and flag updates.
'@param sh Worksheet containing the change.
'@param Target Range representing the changed cells.
Public Sub CheckUpdate(ByVal sh As Worksheet, ByVal Target As Range): End Sub

'@method ClearUp
'@description Reset all tracked columns to the default (not updated) status.
Public Sub ClearUp(): End Sub

'@method DeleteUp
'@description Remove the watcher registry and associated workbook named ranges.
Public Sub DeleteUp(): End Sub
