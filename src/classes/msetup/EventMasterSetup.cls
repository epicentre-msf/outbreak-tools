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
'@depends MasterSetupVariables, DropdownLists, LLChoices, UpdatedValues, MasterSetupLog, Passwords
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit


'@section State
'===============================================================================
Private Type TEventMasterSetup
    hostBook As Workbook                 'Workbook hosting master setup assets
    dropdowns As DropdownLists          'Cached dropdown manager
    variables As MasterSetupVariables   'Cached variables manager
    choices As LLChoices                'Cached choices helper
    updater As UpdatedValues            'Cached update watcher over the __updated registry
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

Private Const PASSWORD_SHEET_NAME As String = "__pass"
Private Const REGISTRY_SHEET_NAME As String = "__updated"



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

'@sub-title Handle the workbook open the way the setup does.
'@details
'The managers are dropped and built from the sheets on the first read after
'the open. The preparation of the workbook is a dev-time step: it writes on
'the sheets, and the Initialize button of the Dev tab is its one door.
'
'Two things the open does write, the same two EventSetup.OnWorkbookOpen
'writes. Every flag of the __updated registry goes to yes, so the next
'Update Translations reads every watched column. And a workbook saved during
'a debug session reopens with the debug flag still set, and that flag is
'what Passwords.Protect reads before it does anything: while it is set every
're-protect is ignored. LeaveDebugMode puts the matrix back and clears the
'flag, and returns untouched when the flag is off; the matrix is applied on
'its own after it because Excel drops UserInterfaceOnly when a file closes,
'and a sheet protected without it is closed to this code as much as to the
'user. Each step is skipped when its helper cannot be built. The application
'argument is kept for the callers that pass it.
Public Sub OnWorkbookOpen(Optional ByVal excelApplication As Application)
    Dim pass As Passwords

    Set this.dropdowns = Nothing
    Set this.variables = Nothing
    ResetTranslationCaches

    EnsureUpdater
    If Not (this.updater Is Nothing) Then
        On Error Resume Next
            this.updater.SwitchTagsToYes
        On Error GoTo 0
    End If

    Set pass = ResolvePasswords()
    If pass Is Nothing Then Exit Sub

    pass.LeaveDebugMode this.hostBook
    pass.ApplyProtectionMatrix this.hostBook
End Sub

'@sub-title Store the ribbon reference for future invalidation support.
Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    Set this.ribbon = ribbon
End Sub

'@sub-title Drop the choices helper and the variables manager after an edit.
'@details The master setup holds its content in English, so the service
'carries no translation object of its own. A choices edit cascades through
'RefreshChoices into the Variables table, and a Variables table that grew
'rows through the ribbon needs a manager over the rows it has today, so
'both caches drop. The manager is cheap to rebuild: the column indexes come
'off the hidden names.
Public Sub RefreshTranslations()
    Set this.variables = Nothing
    ResetTranslationCaches
End Sub


'@section Accessors
'===============================================================================
'@sub-title Expose the currently configured workbook.
Public Property Get HostWorkbook() As Workbook
    Set HostWorkbook = EnsureWorkbook
End Property

'@sub-title Provide the dropdown manager built over the __dropdowns sheet.
'@details Nothing when the workbook has no __dropdowns sheet.
Public Property Get Dropdowns() As DropdownLists
    Set Dropdowns = EnsureDropdownManager
End Property

'@sub-title Provide the master setup variables manager built over the Variables table.
'@details Nothing when the Variables sheet is missing or carries no table.
Public Property Get Variables() As MasterSetupVariables
    Set Variables = EnsureVariablesManager
End Property

'@sub-title Provide the master setup choices helper.
Public Property Get Choices() As LLChoices
    Set Choices = EnsureChoicesHelper
End Property

'@sub-title The update watcher over the __updated registry, or Nothing.
'@details Nothing when the workbook has no registry sheet.
Public Function Updater() As UpdatedValues
    EnsureUpdater
    Set Updater = this.updater
End Function

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
    Set this.dropdowns = Nothing
    Set this.variables = Nothing
    Set this.choices = Nothing
    Set this.updater = Nothing
    Set this.logStore = Nothing
    this.triedLog = False
    ResetTranslationCaches
End Sub

'@sub-title Build the update watcher over the __updated sheet, once.
'@details A workbook without the sheet answers Nothing. Create is guarded
'the way EventSetup guards its watcher.
Private Sub EnsureUpdater()
    Dim registrySheet As Worksheet

    If Not (this.updater Is Nothing) Then Exit Sub

    Set registrySheet = ResolveWorksheet(REGISTRY_SHEET_NAME)
    If registrySheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set this.updater = UpdatedValues.Create(registrySheet)
    On Error GoTo 0
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
    Set this.choices = Nothing    'The helper re-reads the sheet next time
End Sub

Private Function EnsureWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Configure must run before accessing EventMasterSetup."
    End If

    Set EnsureWorkbook = this.hostBook
End Function

'@sub-title Build the dropdown manager over the __dropdowns sheet, once.
'@details
'The manager is built from the sheet as it stands, the way the worksheet
'functions build theirs. A workbook without the sheet answers Nothing.
'Create is guarded the way EventSetup guards its dictionary: a half-built
'sheet answers Nothing and the caller decides.
Private Function EnsureDropdownManager() As DropdownLists
    Dim targetSheet As Worksheet

    If this.dropdowns Is Nothing Then
        Set targetSheet = ResolveWorksheet(DROPDOWNS_SHEET_NAME)
        If targetSheet Is Nothing Then Exit Function

        On Error Resume Next
            Set this.dropdowns = DropdownLists.Create(targetSheet, DROPDOWNS_PREFIX)
        On Error GoTo 0
    End If

    Set EnsureDropdownManager = this.dropdowns
End Function

'@sub-title Build the variables manager over the Variables table, once.
'@details
'The manager reads the first table of the Variables sheet. A sheet that is
'missing or carries no table answers Nothing. Create is guarded the way
'EventSetup guards its dictionary: a half-built table answers Nothing and
'the caller decides.
Private Function EnsureVariablesManager() As MasterSetupVariables
    Dim targetSheet As Worksheet

    If this.variables Is Nothing Then
        Set targetSheet = ResolveWorksheet(VARIABLES_SHEET_NAME)
        If targetSheet Is Nothing Then Exit Function
        If targetSheet.ListObjects.Count = 0 Then Exit Function

        On Error Resume Next
            Set this.variables = MasterSetupVariables.Create(targetSheet.ListObjects(1))
        On Error GoTo 0
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
