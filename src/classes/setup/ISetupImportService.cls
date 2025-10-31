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

'@section Core operations
'===============================================================================

'@jump:Check
'@sub-title Validate the requested import operation before execution.
'@param importDictionary Boolean flag requesting dictionary import.
'@param importChoices Boolean flag requesting choices import.
'@param importExports Boolean flag requesting exports import.
'@param importAnalysis Boolean flag requesting analysis import.
'@param importTranslations Boolean flag requesting translations import.
'@param cleanSetup Optional Boolean flag indicating a clean-only run.
Public Sub Check(ByVal importDictionary As Boolean, _
                 ByVal importChoices As Boolean, _
                 ByVal importExports As Boolean, _
                 ByVal importAnalysis As Boolean, _
                 ByVal importTranslations As Boolean, _
                 Optional ByVal cleanSetup As Boolean = False)
End Sub

'@jump:Import
'@sub-title Perform the setup import from the configured workbook into the host workbook.
'@param passwords IPasswords implementation handling worksheet protection.
'@param sheetsList BetterArray describing the worksheets participating in the import.
Public Sub Import(ByVal passwords As IPasswords, ByVal sheetsList As BetterArray)
End Sub

'@jump:ImportFromWorkbook
'@sub-title Perform the setup import leveraging dedicated worksheet classes.
'@param passwords IPasswords implementation handling worksheet protection.
'@param sheetsList Optional BetterArray describing worksheets to import; defaults to Dictionary, Exports, Analysis and Translations.
Public Sub ImportFromWorkbook(ByVal passwords As IPasswords, Optional ByVal sheetsList As BetterArray = Nothing)
End Sub

'@jump:ExportToWorkbook
Public Sub Export():End Sub

'@jump:Clean
'@sub-title Clean target worksheets ahead of a fresh import.
'@param passwords IPasswords implementation handling worksheet protection.
'@param sheetsList BetterArray describing the worksheets participating in the clean-up.
Public Sub Clean(ByVal passwords As IPasswords, ByVal sheetsList As BetterArray)
End Sub


'@section Configuration
'===============================================================================

'@jump:Path
'@sub-title Path to the source setup workbook.
Public Property Get Path() As String
End Property

Public Property Let Path(ByVal value As String)
End Property

'@jump:ProgressObject
'@sub-title UI object receiving informational messages while the import runs.
Public Property Get ProgressObject() As Object
End Property

Public Property Set ProgressObject(ByVal value As Object)
End Property

'@jump:DisplayPrompts
'@sub-title Toggle UI prompts displayed during operations (primarily for automated tests).
Public Property Get DisplayPrompts() As Boolean
End Property

Public Property Let DisplayPrompts(ByVal value As Boolean)
End Property

'@jump:SetExportFolder
'@sub-title Provide an explicit folder for exports, bypassing the folder picker.
Public Sub SetExportFolder(ByVal folderPath As String)
End Sub


'@section Miscellaneous
'===============================================================================

'@jump:LastExportFile
'@sub-title Path of the last export produced by the service (empty when no export occurred).
Public Property Get LastExportFile() As String
End Property

'@jump:Self
'@sub-title Retrieve the interface bound to the instance.
Public Property Get Self() As ISetupImportService
End Property
