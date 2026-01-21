VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesignerPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract preparing designer workbooks and ribbon settings"

'@Interface
'@Folder("Designer")
'@ModuleDescription("Prepares designer workbooks for ribbon interactions and exposes persisted settings.")
'@depends IDevelopment, IHiddenNames
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Lifecycle
'===============================================================================
'@Description("Bind the preparation helper to a designer workbook.")
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@Description("Execute preparation tasks using the supplied development manager when available.")
Public Sub Prepare(ByRef devManager As IDevelopment)
End Sub

'@section Settings
'===============================================================================
'@Description("Persist a ribbon flag (e.g. alert/instruct) as a workbook-level hidden name.")
Public Sub SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
End Sub

'@Description("Read a ribbon flag from workbook-level hidden names.")
Public Function GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
End Function

'@section State
'===============================================================================
'@Description("Expose the configured host workbook.")
Public Property Get HostWorkbook() As Workbook
End Property

'@Description("Expose the hidden names store used to persist designer settings.")
Public Property Get HiddenStore() As IHiddenNames
End Property
