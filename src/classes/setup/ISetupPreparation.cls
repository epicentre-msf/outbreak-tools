VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISetupPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface orchestrating setup preparation tasks"

'@Interface
'@Folder("Setup")
'@ModuleDescription("Describes the preparation workflow responsible for dropdown registration and update watcher initialisation")
'@depends IDevelopment
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Lifecycle
'===============================================================================
'@Description("Bind the preparation helper to a workbook hosting setup worksheets.")
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@Description("Execute the preparation workflow (dropdowns first, then updated values registry).")
Public Sub Prepare(ByRef manager As IDevelopment)
End Sub

'@section Operations
'===============================================================================
'@Description("Ensure dropdown list objects exist with the expected default entries.")
Public Sub EnsureDropdowns()
End Sub

'@Description("Initialise the updated values registry worksheet and register watched tables.")
Public Sub EnsureUpdatedRegistry()
End Sub

'@section State
'===============================================================================
'@Description("Expose the dropdown manager created during preparation.")
Public Property Get Dropdowns() As IDropdownLists
End Property

'@Description("Expose the updated values watcher created during preparation.")
Public Property Get UpdatedWatcher() As IUpdatedValues
End Property

'@Description("Expose the workbook hosting the setup worksheets.")
Public Property Get HostWorkbook() As Workbook
End Property
