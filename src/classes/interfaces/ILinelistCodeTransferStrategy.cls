VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistCodeTransferStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Transfers VBA components between workbooks using a specific mechanism"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Transfers VBA components between workbooks using a specific mechanism")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@section Component transfer
'===============================================================================
Public Sub TransferComponent(ByVal componentName As String, _
                             ByVal scope As LinelistCodeScope, _
                             ByVal source As Workbook, _
                             ByVal target As Workbook)
End Sub

Public Sub TransferWorkbookModule(ByVal moduleName As String, _
                                  ByVal source As Workbook, _
                                  ByVal target As Workbook)
End Sub

Public Sub TransferForm(ByVal formName As String, _
                        ByVal source As Workbook, _
                        ByVal target As Workbook)
End Sub


'@section Diagnostics
'===============================================================================
Public Property Get Name() As String
End Property
