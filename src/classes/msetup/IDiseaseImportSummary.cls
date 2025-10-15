VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseImportSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the outcome of a disease import merge operation"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface describing the outcome of a disease import merge operation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section State
'===============================================================================

'@sub-title List of variables that were missing from the import data.
'@return BetterArray clone containing variable names.
Public Property Get MissingVariables() As BetterArray
End Property

'@sub-title List of variables updated from the import data.
'@return BetterArray clone containing variable names.
Public Property Get UpdatedVariables() As BetterArray
End Property

'@sub-title List of variables appended to the target table.
'@return BetterArray clone containing variable names.
Public Property Get AppendedVariables() As BetterArray
End Property

'@sub-title Indicates whether the merge requires a follow-up report.
Public Property Get RequiresReport() As Boolean
End Property

'@section Mutators
'===============================================================================

'@sub-title Register a variable that was missing from the import.
Public Sub RegisterMissing(ByVal varName As String)
End Sub

'@sub-title Register a variable whose details were updated.
Public Sub RegisterUpdated(ByVal varName As String)
End Sub

'@sub-title Register a variable appended from the import.
Public Sub RegisterAppended(ByVal varName As String)
End Sub
