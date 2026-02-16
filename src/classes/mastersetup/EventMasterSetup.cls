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
'@Folder("MasterSetup")
'@ModuleDescription("Coordinates workbook events and shared services for the master setup workbook.")
'@depends MasterSetupPreparation, IMasterSetupPreparation, MasterSetupVariables, IMasterSetupVariables, DropdownLists, IDropdownLists, LLChoices, ILLChoices, Translation, ITranslationObject
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IEventMasterSetup


'@section State
'===============================================================================
Private Type TEventMasterSetup
    hostBook As Workbook                 'Workbook hosting master setup assets
    prep As IMasterSetupPreparation      'Preparation helper ensuring tables and dropdowns exist
    dropdowns As IDropdownLists          'Cached dropdown manager
    variables As IMasterSetupVariables   'Cached variables manager
    choices As ILLChoices                'Cached choices helper
    translations As ITranslationObject   'Cached workbook translations
    ribbonTranslations As ITranslationObject 'Cached ribbon translations
    ribbon As IRibbonUI                  'Cached ribbon interface for future invalidations
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

Private Const RIBBON_TRANSLATION_SHEET As String = "__ribbonTranslation"
Private Const RIBBON_TRANSLATION_TABLE As String = "TabTransId"
Private Const RIBBON_LANGUAGE_RANGE As String = "RNG_FileLang"


'@section Factory helpers
'===============================================================================
'@sub-title Create a configured master setup event manager.
Public Function Create(ByVal hostBook As Workbook) As IEventMasterSetup

    With New EventMasterSetup
        .Configure hostBook
        Set Create = .Self
    End With
End Function

'@sub-title Self reference used by factories returning the interface.
Public Property Get Self() As IEventMasterSetup
    Set Self = Me
End Property


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
    Dim svc As IMasterSetupPreparation

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
Public Property Get Dropdowns() As IDropdownLists
    Set Dropdowns = EnsureDropdownManager
End Property

'@sub-title Provide the master setup variables manager.
Public Property Get Variables() As IMasterSetupVariables
    Set Variables = EnsureVariablesManager
End Property

'@sub-title Provide the master setup choices helper.
Public Property Get Choices() As ILLChoices
    Set Choices = EnsureChoicesHelper
End Property

'@sub-title Provide the workbook-wide translations object.
Public Property Get Translations() As ITranslationObject
    Set Translations = EnsureWorkbookTranslations
End Property

'@sub-title Provide the ribbon translations object.
Public Property Get RibbonTranslations() As ITranslationObject
    Set RibbonTranslations = EnsureRibbonTranslations
End Property

'@sub-title Expose the cached ribbon interface.
Public Property Get Ribbon() As IRibbonUI
    Set Ribbon = this.ribbon
End Property


'@section Internal helpers
'===============================================================================
Private Sub ResetCaches()
    Set this.prep = Nothing
    Set this.dropdowns = Nothing
    Set this.variables = Nothing
    Set this.choices = Nothing
    ResetTranslationCaches
End Sub

Private Sub ResetTranslationCaches()
    Set this.translations = Nothing
    Set this.ribbonTranslations = Nothing
    Set this.choices = Nothing    'Choices rely on translated headers
End Sub

Private Sub CacheCoreManagers(ByVal svc As IMasterSetupPreparation)
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

Private Function Preparation() As IMasterSetupPreparation
    If this.prep Is Nothing Then
        Set this.prep = MasterSetupPreparation.Create(EnsureWorkbook)
    End If

    Set Preparation = this.prep
End Function

Private Sub EnsureCorePrepared()
    Dim svc As IMasterSetupPreparation

    Set svc = Preparation()
    svc.EnsureDropdowns
    svc.EnsureVariables
End Sub

Private Function EnsureDropdownManager() As IDropdownLists
    EnsureCorePrepared

    If this.dropdowns Is Nothing Then
        Set this.dropdowns = Preparation().Dropdowns
    End If

    Set EnsureDropdownManager = this.dropdowns
End Function

Private Function EnsureVariablesManager() As IMasterSetupVariables
    EnsureCorePrepared

    If this.variables Is Nothing Then
        Set this.variables = Preparation().Variables
    End If

    Set EnsureVariablesManager = this.variables
End Function

Private Function EnsureChoicesHelper() As ILLChoices
    Dim targetSheet As Worksheet

    If this.choices Is Nothing Then
        Set targetSheet = ResolveWorksheet(CHOICES_SHEET_NAME)
        If targetSheet Is Nothing Then Exit Function
        Set this.choices = LLChoices.Create(targetSheet, CHOICES_START_ROW, CHOICES_START_COLUMN, _
                                            EnsureWorkbookTranslations, CHOICES_IDENTIFIER)
    End If

    Set EnsureChoicesHelper = this.choices
End Function

Private Function EnsureWorkbookTranslations() As ITranslationObject
    Dim table As ListObject
    Dim languageTag As String

    If this.translations Is Nothing Then
        Set table = ResolveTranslationTable(TRANSLATIONS_SHEET_NAME, TRANSLATIONS_TABLE_NAME)
        If table Is Nothing Then Exit Function

        languageTag = ResolveLanguageTag()
        Set this.translations = Translation.Create(table, languageTag)
    End If

    Set EnsureWorkbookTranslations = this.translations
End Function

Private Function EnsureRibbonTranslations() As ITranslationObject
    Dim table As ListObject
    Dim languageTag As String

    If this.ribbonTranslations Is Nothing Then
        Set table = ResolveTranslationTable(RIBBON_TRANSLATION_SHEET, RIBBON_TRANSLATION_TABLE)
        If table Is Nothing Then Exit Function

        languageTag = ResolveLanguageTag()
        Set this.ribbonTranslations = Translation.Create(table, languageTag)
    End If

    Set EnsureRibbonTranslations = this.ribbonTranslations
End Function

Private Function ResolveLanguageTag() As String
    Dim ribbonSheet As Worksheet
    Dim candidate As Variant

    Set ribbonSheet = ResolveWorksheet(RIBBON_TRANSLATION_SHEET)
    If ribbonSheet Is Nothing Then Exit Function

    On Error Resume Next
        candidate = ribbonSheet.Range(RIBBON_LANGUAGE_RANGE).Value
    On Error GoTo 0

    If IsError(candidate) Then
        ResolveLanguageTag = vbNullString
    Else
        ResolveLanguageTag = Trim$(CStr(candidate))
    End If
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

Private Sub ValidateWorkbook(ByVal workbook As Workbook)
    If workbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook reference is required."
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASS_NAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub IEventMasterSetup_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Sub IEventMasterSetup_OnWorkbookOpen(Optional ByVal excelApplication As Application)
    OnWorkbookOpen excelApplication
End Sub

Private Sub IEventMasterSetup_OnRibbonLoad(ByVal ribbon As IRibbonUI)
    OnRibbonLoad ribbon
End Sub

Private Sub IEventMasterSetup_RefreshTranslations()
    RefreshTranslations
End Sub

Private Property Get IEventMasterSetup_HostWorkbook() As Workbook
    Set IEventMasterSetup_HostWorkbook = HostWorkbook
End Property

Private Property Get IEventMasterSetup_Dropdowns() As IDropdownLists
    Set IEventMasterSetup_Dropdowns = Dropdowns
End Property

Private Property Get IEventMasterSetup_Variables() As IMasterSetupVariables
    Set IEventMasterSetup_Variables = Variables
End Property

Private Property Get IEventMasterSetup_Choices() As ILLChoices
    Set IEventMasterSetup_Choices = Choices
End Property

Private Property Get IEventMasterSetup_Translations() As ITranslationObject
    Set IEventMasterSetup_Translations = Translations
End Property

Private Property Get IEventMasterSetup_RibbonTranslations() As ITranslationObject
    Set IEventMasterSetup_RibbonTranslations = RibbonTranslations
End Property

Private Property Get IEventMasterSetup_Ribbon() As IRibbonUI
    Set IEventMasterSetup_Ribbon = Ribbon
End Property
