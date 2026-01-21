VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesignerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for managing designer entry sheet interactions"

'@Interface
'@Folder("Designer")
'@ModuleDescription("Coordinates clearing and translation actions for designer entry sheets.")
'@depends IMain, IDesTranslation
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Lifecycle
'===============================================================================
'@Description("Bind the entry helper to the worksheet hosting designer inputs.")
Public Sub Configure(ByVal hostSheet As Worksheet)
End Sub

'@Description("Swap in a custom entry manager implementation (mainly for tests).")
Public Sub UseEntryManager(ByVal entryManager As IMain)
End Sub

'@Description("Swap in a custom translations provider (mainly for tests).")
Public Sub UseTranslator(ByVal translator As IDesTranslation)
End Sub

'@section Operations
'===============================================================================
'@Description("Clear user entry ranges on the configured worksheet.")
Public Sub Clear()
End Sub

'@Description("Translate the configured worksheet to the provided language code when supplied.")
Public Sub Translate(Optional ByVal languageCode As String = vbNullString)
End Sub

'@section State
'===============================================================================
'@Description("Expose the worksheet managed by this helper.")
Public Property Get HostSheet() As Worksheet
End Property
