VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISetupImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the improved setup import service"

'@Interface
'@Folder("Setup")
'@ModuleDescription("Interface describing the improved setup import workflow for setup synchronisation")
'@depends BetterArray, IPasswords
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class ISetupImportService
'@description
'Defines the public contract for the SetupImportService class, exposing
'methods to validate, import, export, and clean setup worksheets between
'workbooks. Supports both table-level import (via CustomTable) and
'domain-class-driven import (via LLdictionary, LLChoices, LLExport,
'Analysis, SetupTranslationsTable).


'@section Core operations
'===============================================================================

'@jump:check
'@sub-title Validate the requested import operation before execution
'@details
'Ensures at least one import option is selected, verifies the import
'path is configured and the file exists, then attempts to open and
'immediately close the source workbook to confirm it is readable.
'@param importDictionary Boolean. Flag requesting dictionary import.
'@param importChoices Boolean. Flag requesting choices import.
'@param importExports Boolean. Flag requesting exports import.
'@param importAnalysis Boolean. Flag requesting analysis import.
'@param importTranslations Boolean. Flag requesting translations import.
'@param cleanSetup Optional Boolean. When True, performs a clean-only run. Defaults to False.
'@export
Public Sub Check(ByVal importDictionary As Boolean, _
                 ByVal importChoices As Boolean, _
                 ByVal importExports As Boolean, _
                 ByVal importAnalysis As Boolean, _
                 ByVal importTranslations As Boolean, _
                 Optional ByVal cleanSetup As Boolean = False)
End Sub

'@jump:import
'@sub-title Perform the setup import from the configured workbook into the host workbook
'@details
'Opens the source workbook, iterates over each requested sheet, matches
'ListObjects between source and host, and delegates import to
'CustomTable.Import. Handles sheet protection/unprotection and export
'column alignment via PrepareImport.
'@param passwords IPasswords. Password handler for worksheet protection.
'@param sheetsList BetterArray. Worksheets participating in the import.
'@export
Public Sub Import(ByVal passwords As IPasswords, ByVal sheetsList As BetterArray)
End Sub

'@jump:import-from-workbook
'@sub-title Perform the setup import leveraging dedicated worksheet classes
'@details
'Uses domain-specific classes (LLdictionary, LLChoices, LLExport, Analysis,
'SetupTranslationsTable) to import each worksheet from the source workbook.
'Falls back to the default sheets list when none is provided.
'@param passwords IPasswords. Password handler for worksheet protection.
'@param sheetsList Optional BetterArray. Worksheets to import. Defaults to Dictionary, Choices, Exports, Analysis, and Translations.
'@export
Public Sub ImportFromWorkbook(ByVal passwords As IPasswords, Optional ByVal sheetsList As BetterArray = Nothing)
End Sub

'@jump:export-to-workbook
'@sub-title Export all setup worksheets to a new workbook
'@details
'Creates or reuses an export workbook, then exports each setup worksheet
'(Dictionary, Choices, Exports, Analysis, Translations) using domain-specific
'managers, along with hidden names and formatting data.
'@export
Public Sub Export():End Sub

'@jump:clean
'@sub-title Clean target worksheets ahead of a fresh import
'@details
'Iterates over the requested sheets, unprotects each one, clears all
'ListObject data and worksheet comments, then re-protects.
'@param passwords IPasswords. Password handler for worksheet protection.
'@param sheetsList BetterArray. Worksheets participating in the clean-up.
'@export
Public Sub Clean(ByVal passwords As IPasswords, ByVal sheetsList As BetterArray)
End Sub


'@section Configuration
'===============================================================================

'@jump:path
'@prop-title Path to the source setup workbook
'@return String. The configured import path.
'@export
Public Property Get Path() As String
End Property

'@jump:path-set
'@prop-title Store the path to the setup workbook
'@param value String. The path to assign.
Public Property Let Path(ByVal value As String)
End Property

'@jump:progress-object
'@prop-title UI object receiving informational messages while the import runs
'@return Object. The progress display object.
'@export
Public Property Get ProgressObject() As Object
End Property

'@jump:progress-object-set
'@prop-title Assign the progress display object
'@param value Object. The progress display to store.
Public Property Set ProgressObject(ByVal value As Object)
End Property

'@jump:display-prompts
'@prop-title Toggle UI prompts displayed during operations
'@details
'When False, suppresses folder-picker dialogs and other interactive prompts.
'Primarily used by automated tests.
'@return Boolean. True when prompts are enabled.
'@export
Public Property Get DisplayPrompts() As Boolean
End Property

'@jump:display-prompts-set
'@prop-title Assign the display prompts toggle
'@param value Boolean. True to enable prompts.
Public Property Let DisplayPrompts(ByVal value As Boolean)
End Property

'@jump:set-export-folder
'@sub-title Provide an explicit folder for exports, bypassing the folder picker
'@param folderPath String. The folder path to use for exports.
'@export
Public Sub SetExportFolder(ByVal folderPath As String)
End Sub


'@section Miscellaneous
'===============================================================================

'@jump:last-export-file
'@prop-title Path of the last export produced by the service
'@details
'Returns the full file path of the last export workbook saved by the
'service. Returns an empty string when no export has occurred.
'@return String. The last export file path, or empty.
'@export
Public Property Get LastExportFile() As String
End Property

'@jump:self
'@prop-title Current object instance
'@return ISetupImportService. The current instance cast to the interface.
Public Property Get Self() As ISetupImportService
End Property
