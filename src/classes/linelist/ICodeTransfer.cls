VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICodeTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for transferring VBA code components between workbooks"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Contract for transferring VBA code components between workbooks")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@section Code transfer
'===============================================================================

'@description Transfer a class module (.cls) from source to target workbook via export/import.
Public Sub TransferClass(ByVal className As String)
End Sub

'@description Transfer a standard module (.bas) from source to target workbook via export/import.
Public Sub TransferModule(ByVal moduleName As String)
End Sub

'@description Transfer a UserForm (.frm/.frx) from source to target workbook via export/import.
Public Sub TransferForm(ByVal formName As String)
End Sub

'@description Copy code content from a source module into the target workbook's ThisWorkbook component.
Public Sub TransferWorkbookCode(ByVal moduleName As String)
End Sub

'@description Copy code text from a source module to an existing same-named module in the target workbook.
Public Sub CopyModuleText(ByVal moduleName As String)
End Sub

'@description Within the target workbook, copy code from one component to another, then remove the source.
Public Sub CopyModuleCode(ByVal sourceName As String, ByVal targetName As String)
End Sub
