VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMasterSetupPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Master Setup preparation and configuration"

'@Interface
'@Folder("MasterSetup")
'@ModuleDescription("Interface for Master Setup preparation and configuration")
'@depends IDropdownLists, IMasterSetupVariables
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Lifecycle
'===============================================================================
'@Description("Bind the preparation orchestrator to the workbook hosting master setup assets.")
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@Description("Run the master setup preparation workflow.")
'@param excelApplication Optional Excel Application reference used to guard busy state transitions.
Public Sub Prepare(Optional ByVal excelApplication As Application)
End Sub

'@section Operations
'===============================================================================
'@Description("Ensure core dropdown lists exist on the master setup dropdown worksheet.")
Public Sub EnsureDropdowns()
End Sub

'@Description("Ensure the master setup variables table exists and is initialised.")
Public Sub EnsureVariables()
End Sub

'@section Accessors
'===============================================================================
'@Description("Expose the dropdown manager configured during preparation.")
Public Property Get Dropdowns() As IDropdownLists
End Property

'@Description("Expose the variables manager bound to the master setup table.")
Public Property Get Variables() As IMasterSetupVariables
End Property

'@Description("Expose the workbook currently configured for preparation.")
Public Property Get HostWorkbook() As Workbook
End Property
