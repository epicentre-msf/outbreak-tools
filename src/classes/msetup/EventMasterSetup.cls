VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventMasterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Workbook-level orchestrator for master setup lifecycle"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Coordinates workbook events and shared services for the master setup workbook.")
'@depends MasterSetupPreparation, MasterSetupVariables, DropdownLists, LLChoices, TranslationObject, MasterSetupLog, Passwords
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit


'@section State
'===============================================================================
Private Type TEventMasterSetup
    hostBook As Workbook                 'Workbook hosting master setup assets
    prep As MasterSetupPreparation      'Preparation helper ensuring tables and dropdowns exist
    dropdowns As DropdownLists          'Cached dropdown manager
    variables As MasterSetupVariables   'Cached variables manager
    choices As LLChoices                'Cached choices helper
    translations As TranslationObject   'Cached workbook translations
    ribbon As IRibbonUI                  'Cached ribbon interface for future invalidations
    logStore As MasterSetupLog           'Cached user log on the __log worksheet
    triedLog As Boolean                  'The log build ran once, whatever it answered
End Type

Private this As TEventMasterSetup


'@section Constants
'===============================================================================
Private Const CLASS_NAME As String = "EventMasterSetup"

Private Const VARIABLES_SHEET_NAME As String = "Variables"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const CHOICES_START_ROW As Long = 4
Private Const CHOICES_START_COLUMN As Long = 1
Private Const CHOICES_IDENTIFIER As String = "__master_choices"

Private Const DROPDOWNS_SHEET_NAME As String = "__dropdowns"
Private Const DROPDOWNS_PREFIX As String = "dropdown_"

Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"

Private Const PASSWORD_SHEET_NAME As String = "__pass"



'@section Factory helpers
'===============================================================================
'@sub-title Create a configured master setup event manager.
Public Function Create(ByVal hostBook As Workbook) As EventMasterSetup

    Dim instance As EventMasterSetup
    Set instance = New EventMasterSetup
    instance.Configure hostBook
    Set Create = instance
End Function


'@section Lifecycle
'===============================================================================
'@sub-title Bind the orchestrator to the workbook hosting master setup assets.
Public Sub Configure(ByVal hostBook As Workbook)
    ValidateWorkbook hostBook
    Set this.hostBook = hostBook
    ResetCaches
End Sub

'@sub-title Handle workbook open events by running preparation once.
Public Sub OnWorkbookOpen(Optional ByVal excelApplication As Application)
    Dim svc As MasterSetupPreparation

    Set svc = Preparation()
    svc.Prepare excelApplication

    CacheCoreManagers svc
    ResetTranslationCaches
End Sub

'@sub-title Store the ribbon reference for future invalidation support.
Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    Set this.ribbon = ribbon
End Sub

'@sub-title Refresh translation-backed caches after a language change.
Public Sub RefreshTranslations()
    ResetTranslationCaches
End Sub


'@section Accessors
'===============================================================================
'@sub-title Expose the currently configured workbook.
Public Property Get HostWorkbook() As Workbook
    Set HostWorkbook = EnsureWorkbook
End Property

'@sub-title Provide the dropdown manager, ensuring the preparation ran.
Public Property Get Dropdowns() As DropdownLists
    Set Dropdowns = EnsureDropdownManager
End Property

'@sub-title Provide the master setup variables manager.
Public Property Get Variables() As MasterSetupVariables
    Set Variables = EnsureVariablesManager
End Property

'@sub-title Provide the master setup choices helper.
Public Property Get Choices() As LLChoices
    Set Choices = EnsureChoicesHelper
End Property

'@sub-title Provide the workbook-wide translations object.
Public Property Get Translations() As TranslationObject
    Set Translations = EnsureWorkbookTranslations
End Property

'@sub-title Expose the cached ribbon interface.
Public Property Get Ribbon() As IRibbonUI
    Set Ribbon = this.ribbon
End Property

'@sub-title The user log of the workbook, built once and held.
'@details
'The dated record of the user actions, kept on the very hidden __log
'worksheet. The ribbon, the events manager and the exports module write
'their outcome lines through it. An old master setup grows the log sheet on
'the first logged event.
'@return MasterSetupLog. The held log, or Nothing when the build failed.
Public Function UserLog() As MasterSetupLog
    EnsureLog
    Set UserLog = this.logStore
End Function


'@section Internal helpers
'===============================================================================
Private Sub ResetCaches()
    Set this.prep = Nothing
    Set this.dropdowns = Nothing
    Set this.variables = Nothing
    Set this.choices = Nothing
    Set this.logStore = Nothing
    this.triedLog = False
    ResetTranslationCaches
End Sub

'@sub-title Build the user log once, growing the __log sheet when it is missing.
'@details
'Adding a worksheet needs the workbook structure open, so a workbook without
'the log sheet is unprotected through its passwords for the build and
'protected again after it. A workbook with no usable password sheet is
'built bare. The tried flag keeps a failed build from repeating on every
'logged line, and a failed build leaves the log at Nothing so every caller
'stays quiet.
Private Sub EnsureLog()
    Dim mustBuild As Boolean
    Dim passManager As Passwords

    If Not (this.logStore Is Nothing) Then Exit Sub
    If this.triedLog Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedLog = True

    mustBuild = (ResolveWorksheet(MasterSetupLog.SheetName) Is Nothing)
    If mustBuild Then Set passManager = ResolvePasswords()

    On Error Resume Next
    If Not (passManager Is Nothing) Then passManager.UnProtect this.hostBook
    Set this.logStore = MasterSetupLog.Create(this.hostBook)
    If Not (passManager Is Nothing) Then passManager.Protect this.hostBook
    On Error GoTo 0
End Sub

'@sub-title The password manager of the workbook, or Nothing.
'@details
'Passwords.Create raises on a sheet missing its tables or its names; a
'workbook without a usable password sheet answers Nothing.
Private Function ResolvePasswords() As Passwords
    Dim passSheet As Worksheet

    Set passSheet = ResolveWorksheet(PASSWORD_SHEET_NAME)
    If passSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolvePasswords = Passwords.Create(passSheet)
    On Error GoTo 0
End Function

Private Sub ResetTranslationCaches()
    Set this.translations = Nothing
    Set this.choices = Nothing    'The helper re-reads the sheet next time
End Sub

Private Sub CacheCoreManagers(ByVal svc As MasterSetupPreparation)
    If svc Is Nothing Then Exit Sub

    On Error Resume Next
        Set this.dropdowns = svc.Dropdowns
        Set this.variables = svc.Variables
    On Error GoTo 0
End Sub

Private Function EnsureWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Configure must run before accessing EventMasterSetup."
    End If

    Set EnsureWorkbook = this.hostBook
End Function

Private Function Preparation() As MasterSetupPreparation
    If this.prep Is Nothing Then
        Set this.prep = MasterSetupPreparation.Create(EnsureWorkbook)
    End If

    Set Preparation = this.prep
End Function

Private Sub EnsureCorePrepared()
    Dim svc As MasterSetupPreparation

    Set svc = Preparation()
    svc.EnsureDropdowns
    svc.EnsureVariables
End Sub

Private Function EnsureDropdownManager() As DropdownLists
    EnsureCorePrepared

    If this.dropdowns Is Nothing Then
        Set this.dropdowns = Preparation().Dropdowns
    End If

    Set EnsureDropdownManager = this.dropdowns
End Function

Private Function EnsureVariablesManager() As MasterSetupVariables
    EnsureCorePrepared

    If this.variables Is Nothing Then
        Set this.variables = Preparation().Variables
    End If

    Set EnsureVariablesManager = this.variables
End Function

Private Function EnsureChoicesHelper() As LLChoices
    Dim targetSheet As Worksheet

    If this.choices Is Nothing Then
        Set targetSheet = ResolveWorksheet(CHOICES_SHEET_NAME)
        If targetSheet Is Nothing Then Exit Function
        'The master setup stays in English: the choices carry no
        'translation object, and their headers keep their English names.
        Set this.choices = LLChoices.Create(targetSheet, CHOICES_START_ROW, CHOICES_START_COLUMN, CHOICES_IDENTIFIER)
    End If

    Set EnsureChoicesHelper = this.choices
End Function

Private Function EnsureWorkbookTranslations() As TranslationObject
    Dim table As ListObject
    Dim languageTag As String

    If this.translations Is Nothing Then
        Set table = ResolveTranslationTable(TRANSLATIONS_SHEET_NAME, TRANSLATIONS_TABLE_NAME)
        If table Is Nothing Then Exit Function

        'The file language went away with the ribbon dropdown; the first
        'language column of the table is the default.
        languageTag = Trim$(CStr(table.HeaderRowRange.Cells(1, 1).Value))

        Set this.translations = TranslationObject.Create(table, languageTag)
    End If

    Set EnsureWorkbookTranslations = this.translations
End Function

Private Function ResolveTranslationTable(ByVal sheetName As String, _
                                         ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    Set ws = ResolveWorksheet(sheetName)
    If ws Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolveTranslationTable = ws.ListObjects(tableName)
    On Error GoTo 0

    If ResolveTranslationTable Is Nothing Then
        If ws.ListObjects.Count = 1 Then
            Set ResolveTranslationTable = ws.ListObjects(1)
        End If
    End If
End Function

Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    Dim wb As Workbook

    Set wb = EnsureWorkbook
    On Error Resume Next
        Set ResolveWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub ValidateWorkbook(ByVal targetBook As Workbook)
    If targetBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook reference is required."
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASS_NAME, message
End Sub
