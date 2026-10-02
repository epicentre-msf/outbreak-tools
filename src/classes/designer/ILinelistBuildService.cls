VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistBuildService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for direct setup-to-linelist build service"

'@Interface
'@Folder("Designer")
'@ModuleDescription("Defines the contract for building a linelist directly from setup and designer workbooks.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class ILinelistBuildService
'@description
'Defines the public contract for the LinelistBuildService class, which
'orchestrates building a linelist by exporting pass-through components
'(dictionary, choices, exports, analysis, translations, hidden names)
'directly from a setup workbook to the linelist, and exporting
'designer-sourced components (geo, passwords, formatter, language codes)
'from the designer workbook to the linelist. After all exports,
'domain manager instances targeting the linelist workbook are available
'via accessor properties.


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

'@jump:import-setup-to-linelist
'@sub-title Export pass-through components directly from setup to linelist
'@details
'Opens the setup workbook at the supplied path (read-only), then exports
'dictionary, choices, exports, analysis, and hidden names directly into
'the target linelist workbook. The formatter is also exported from setup
'when it was not already imported via the ribbon (TAG_FORMATTER_IMPORTED).
'Translations are created directly on the linelist sheet. The setup
'workbook is always closed on exit, even when errors occur.
'@param setupPath String. Full path to the setup workbook (.xlsb or .xlsx).
'@param targetBook Workbook. The linelist workbook to receive exported data.
'@export
Public Sub ImportSetupToLinelist(ByVal setupPath As String, ByVal targetBook As Workbook)
End Sub

'@jump:export-designer-components
'@sub-title Export designer-sourced components to the linelist
'@details
'Exports geo, passwords, formatter (when ribbon-imported), and language
'HiddenNames from the designer workbook into the target linelist workbook.
'These are components that the user configures via the designer ribbon
'and cannot come directly from the setup file.
'@param targetBook Workbook. The linelist workbook to receive exported data.
'@export
Public Sub ExportDesignerComponents(ByVal targetBook As Workbook)
End Sub

'@jump:build-exported-instances
'@sub-title Create domain manager instances from the exported linelist
'@details
'After ImportSetupToLinelist and ExportDesignerComponents have written
'sheets into the target workbook, this method creates new domain manager
'instances pointing to those sheets. Each manager is created only when
'its corresponding sheet exists in the target workbook.
'@param targetBook Workbook. The linelist workbook containing exported data.
'@export
Public Sub BuildExportedInstances(ByVal targetBook As Workbook)
End Sub


'@section Exported Instances
'===============================================================================

'@jump:exported-dictionary
'@prop-title Dictionary manager targeting the linelist workbook
'@return ILLdictionary. The exported dictionary manager, or Nothing.
'@export
Public Property Get ExportedDictionary() As ILLdictionary
End Property

'@jump:exported-choices
'@prop-title Choices manager targeting the linelist workbook
'@return ILLChoices. The exported choices manager, or Nothing.
'@export
Public Property Get ExportedChoices() As ILLChoices
End Property

'@jump:exported-analysis
'@prop-title Analysis manager targeting the linelist workbook
'@return IAnalysis. The exported analysis manager, or Nothing.
'@export
Public Property Get ExportedAnalysis() As IAnalysis
End Property

'@jump:exported-export
'@prop-title Export manager targeting the linelist workbook
'@return ILLExport. The exported export manager, or Nothing.
'@export
Public Property Get ExportedExport() As ILLExport
End Property

'@jump:exported-geo
'@prop-title Geo manager targeting the linelist workbook
'@return ILLGeo. The exported geo manager, or Nothing.
'@export
Public Property Get ExportedGeo() As ILLGeo
End Property

'@jump:exported-passwords
'@prop-title Passwords manager targeting the linelist workbook
'@return IPasswords. The exported passwords manager, or Nothing.
'@export
Public Property Get ExportedPasswords() As IPasswords
End Property

'@jump:exported-translations
'@prop-title Translations manager targeting the linelist workbook
'@return ILLTranslation. The exported translations manager, or Nothing.
'@export
Public Property Get ExportedTranslations() As ILLTranslation
End Property


'@section State
'===============================================================================

'@jump:host-workbook
'@prop-title The configured host (designer) workbook
'@return Workbook. The host workbook.
'@export
Public Property Get HostWorkbook() As Workbook
End Property
