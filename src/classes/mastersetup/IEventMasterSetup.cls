VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEventMasterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for managing Master Setup event handling"

'@Interface
'@Folder("MasterSetup")
'@ModuleDescription("Interface for managing Master Setup event handling")
'@depends IDropdownLists, IMasterSetupVariables, ILLChoices, ITranslationObject
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Lifecycle
'===============================================================================
'@Description("Bind the orchestrator to the workbook hosting master setup assets.")
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@Description("Run initialisation logic when the workbook opens.")
Public Sub OnWorkbookOpen(Optional ByVal excelApplication As Application)
End Sub

'@Description("Capture the ribbon reference once callbacks load.")
Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
End Sub

'@Description("Refresh translation-backed caches after a language change.")
Public Sub RefreshTranslations()
End Sub

'@section Accessors
'===============================================================================
'@Description("Expose the workbook configured for the master setup session.")
Public Property Get HostWorkbook() As Workbook
End Property

'@Description("Expose the dropdown manager ready for row operations.")
Public Property Get Dropdowns() As IDropdownLists
End Property

'@Description("Expose the master setup variables manager.")
Public Property Get Variables() As IMasterSetupVariables
End Property

'@Description("Expose the choices helper associated with the master setup workbook.")
Public Property Get Choices() As ILLChoices
End Property

'@Description("Expose the workbook translations used for headers and content.")
Public Property Get Translations() As ITranslationObject
End Property

'@Description("Expose the ribbon translations table for UI captions.")
Public Property Get RibbonTranslations() As ITranslationObject
End Property

'@Description("Expose the cached ribbon interface for invalidation scenarios.")
Public Property Get Ribbon() As IRibbonUI
End Property
