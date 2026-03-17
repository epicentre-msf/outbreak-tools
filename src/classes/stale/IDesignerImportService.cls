VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDesignerImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for designer import/export service"

'@Interface
'@Folder("Designer")
'@ModuleDescription("Defines the contract for importing setup components into a designer workbook and exporting them to a linelist workbook.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class IDesignerImportService
'@description
'Defines the public contract for the DesignerImportService class, exposing
'methods to import setup components (dictionary, choices, analysis, exports,
'formatter, translations) from a setup workbook into the designer, and to
'export components (dictionary, choices, analysis, exports, geo, passwords)
'to a linelist workbook. After export, the service provides accessor
'properties for new domain manager instances targeting the exported workbook.


'@section Lifecycle
'===============================================================================

'@jump:configure
'@sub-title Bind the service to a designer workbook
'@details
'Sets the internal host workbook reference. Must be called before any
'other operation. Raises an error when the workbook is Nothing.
'@param hostBook Workbook. The designer workbook.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
End Sub


'@section Core Operations
'===============================================================================

'@jump:import-from-setup
'@sub-title Import setup components into the host workbook
'@details
'Opens the setup workbook at the supplied path, then imports dictionary,
'choices, exports, analysis, and formatter worksheets by exporting from
'setup domain managers into the host workbook. Translations are imported
'via LLTranslation.ImportDictionary using the setup Translations ListObject.
'The setup workbook is closed after import completes (or on error).
'@param setupPath String. Full path to the setup workbook (.xlsb or .xlsx).
'@export
Public Sub ImportFromSetup(ByVal setupPath As String)
End Sub

'@jump:export-to-linelist
'@sub-title Export designer components to a linelist workbook
'@details
'Exports dictionary, choices, analysis, exports, geo, and passwords
'worksheets from the host workbook into the provided target workbook.
'The target workbook is NOT closed after export. New domain manager
'instances targeting the exported workbook are stored and accessible
'via the Exported* properties.
'@param targetBook Workbook. The linelist workbook to receive exported data.
'@export
Public Sub ExportToLinelist(ByVal targetBook As Workbook)
End Sub


'@section Exported Instances
'===============================================================================

'@jump:exported-dictionary
'@prop-title Dictionary manager targeting the exported workbook
'@details
'Returns an ILLdictionary instance bound to the Dictionary sheet in the
'target workbook after ExportToLinelist completes. Returns Nothing before
'export or when the sheet was not exported.
'@return ILLdictionary. The exported dictionary manager, or Nothing.
'@export
Public Property Get ExportedDictionary() As ILLdictionary
End Property

'@jump:exported-choices
'@prop-title Choices manager targeting the exported workbook
'@return ILLChoices. The exported choices manager, or Nothing.
'@export
Public Property Get ExportedChoices() As ILLChoices
End Property

'@jump:exported-analysis
'@prop-title Analysis manager targeting the exported workbook
'@return IAnalysis. The exported analysis manager, or Nothing.
'@export
Public Property Get ExportedAnalysis() As IAnalysis
End Property

'@jump:exported-export
'@prop-title Export manager targeting the exported workbook
'@return ILLExport. The exported export manager, or Nothing.
'@export
Public Property Get ExportedExport() As ILLExport
End Property

'@jump:exported-geo
'@prop-title Geo manager targeting the exported workbook
'@return ILLGeo. The exported geo manager, or Nothing.
'@export
Public Property Get ExportedGeo() As ILLGeo
End Property

'@jump:exported-passwords
'@prop-title Passwords manager targeting the exported workbook
'@return IPasswords. The exported passwords manager, or Nothing.
'@export
Public Property Get ExportedPasswords() As IPasswords
End Property

'@jump:exported-translations
'@prop-title Translations manager targeting the exported workbook
'@return ILLTranslation. The exported translations manager, or Nothing.
'@export
Public Property Get ExportedTranslations() As ILLTranslation
End Property


'@section State
'===============================================================================

'@jump:host-workbook
'@prop-title The configured host workbook
'@return Workbook. The host workbook.
'@export
Public Property Get HostWorkbook() As Workbook
End Property
