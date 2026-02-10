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

'@class ICodeTransfer
'@description
'Defines the public contract for the CodeTransfer class, exposing methods to
'transfer VBA code components (classes, modules, forms, workbook code) between
'workbooks using temporary file export/import cycles.


'@section CodeTransfer
'===============================================================================
'@description Methods for moving VBA components between workbooks.

'@jump:transfer-class
'@sub-title Transfer a class module from source to target workbook
'@details
'Exports the named class module (.cls) from the source workbook to a temporary
'file, then imports it into the target workbook.
'@param className String. The VBA component name of the class to transfer.
Public Sub TransferClass(ByVal className As String)
End Sub

'@jump:transfer-module
'@sub-title Transfer a standard module from source to target workbook
'@details
'Exports the named standard module (.bas) from the source workbook to a
'temporary file, then imports it into the target workbook.
'@param moduleName String. The VBA component name of the module to transfer.
Public Sub TransferModule(ByVal moduleName As String)
End Sub

'@jump:transfer-form
'@sub-title Transfer a UserForm from source to target workbook
'@details
'Exports the named UserForm (.frm/.frx) from the source workbook to temporary
'files, then imports it into the target workbook. Both the .frm and .frx files
'are cleaned up after import.
'@param formName String. The VBA component name of the form to transfer.
Public Sub TransferForm(ByVal formName As String)
End Sub

'@jump:transfer-workbook-code
'@sub-title Copy code from a source module into the target's ThisWorkbook component
'@details
'Reads all code lines from the named module in the source workbook and writes
'them into the target workbook's ThisWorkbook code module, replacing any
'existing content.
'@param moduleName String. The source module whose code will be copied.
Public Sub TransferWorkbookCode(ByVal moduleName As String)
End Sub

'@jump:copy-module-text
'@sub-title Copy code text from a source module to an existing target module
'@details
'Reads all code from the named module in the source workbook and writes it
'into the same-named module in the target workbook, replacing existing content.
'Used for template workbooks where the target module already exists.
'@param moduleName String. The module name (must exist in both workbooks).
Public Sub CopyModuleText(ByVal moduleName As String)
End Sub

'@jump:copy-module-code
'@sub-title Copy code between components within the target workbook
'@details
'Within the target workbook, reads code from the source component, writes it
'into the target component (replacing any existing content), then removes the
'source component. Useful after importing a module that should be injected
'into a different component (e.g., a worksheet module or ThisWorkbook).
'@param sourceName String. The component to read code from and then remove.
'@param targetName String. The component to write code into.
Public Sub CopyModuleCode(ByVal sourceName As String, ByVal targetName As String)
End Sub
