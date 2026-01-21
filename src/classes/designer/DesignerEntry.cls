VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Coordinates clearing and translation behaviours for designer entry sheets"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Coordinates clearing and translation behaviours for designer entry sheets.")
'@depends ProjectError, Main, IMain, DesTranslation, IDesTranslation
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements IDesignerEntry

'@section State
'===============================================================================
Private Type TDesignerEntry
    hostSheet As Worksheet      'Worksheet hosting designer inputs
    entryManager As IMain       'Entry manager responsible for clearing values
    translator As IDesTranslation 'Translation handler for the designer workbook
End Type

Private this As TDesignerEntry

'@section Constants
'===============================================================================
Private Const CLASS_NAME As String = "DesignerEntry"
Private Const TRANSLATIONS_SHEET As String = "DesignerTranslation"
Private Const LANGUAGE_CODE_RANGE As String = "RNG_MainLangCode"


'@section Factory
'===============================================================================
'@Description("Create and configure a designer entry helper.")
Public Function Create(ByVal hostSheet As Worksheet) As IDesignerEntry
    Dim instance As DesignerEntry

    Set instance = New DesignerEntry
    instance.Configure hostSheet
    Set Create = instance.Self
End Function

'@Description("Expose the current instance as an interface reference.")
Public Property Get Self() As IDesignerEntry
    Set Self = Me
End Property


'@section Public API
'===============================================================================
'@Description("Bind the helper to the target worksheet.")
Public Sub Configure(ByVal hostSheet As Worksheet)
    If hostSheet Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host sheet is required."
    End If

    Set this.hostSheet = hostSheet
End Sub

'@Description("Use a provided entry manager instead of creating one (primarily for tests).")
Public Sub UseEntryManager(ByVal entryManager As IMain)
    Set this.entryManager = entryManager
End Sub

'@Description("Use a provided translations manager instead of creating one (primarily for tests).")
Public Sub UseTranslator(ByVal translator As IDesTranslation)
    Set this.translator = translator
End Sub

'@Description("Clear user inputs on the configured worksheet.")
Public Sub Clear()
    ResolveEntryManager.ClearInputRanges clearValues:=True
End Sub

'@Description("Translate the configured worksheet and optionally update the language code.")
Public Sub Translate(Optional ByVal languageCode As String = vbNullString)
    Dim translator As IDesTranslation
    Dim langRange As Range

    Set translator = ResolveTranslator()

    If LenB(languageCode) > 0 Then
        Set langRange = ResolveLanguageRange(translator)
        langRange.Value = languageCode
        langRange.Worksheet.Calculate
    End If

    translator.TranslateDesigner HostSheet
End Sub

'@Description("Expose the managed worksheet.")
Public Property Get HostSheet() As Worksheet
    If this.hostSheet Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "HostSheet requested before configuration."
    End If
    Set HostSheet = this.hostSheet
End Property


'@section Internal helpers
'===============================================================================
Private Function ResolveEntryManager() As IMain
    If this.entryManager Is Nothing Then
        On Error GoTo Handler
            Set this.entryManager = Main.Create(HostSheet)
        On Error GoTo 0
    End If

    Set ResolveEntryManager = this.entryManager
    Exit Function

Handler:
    RaiseError ProjectError.ElementNotFound, "Unable to initialise entry manager for '" & HostSheet.Name & "'."
End Function

Private Function ResolveTranslator() As IDesTranslation
    If this.translator Is Nothing Then
        On Error GoTo Handler
            Set this.translator = DesTranslation.Create(TranslationSheet())
        On Error GoTo 0
    End If

    Set ResolveTranslator = this.translator
    Exit Function

Handler:
    RaiseError ProjectError.ElementNotFound, "Designer translation sheet '" & TRANSLATIONS_SHEET & "' was not found."
End Function

Private Function TranslationSheet() As Worksheet
    Dim sh As Worksheet
    Dim wb As Workbook

    Set wb = HostSheet.Parent
    On Error Resume Next
        Set sh = wb.Worksheets(TRANSLATIONS_SHEET)
    On Error GoTo 0

    If sh Is Nothing Then
        RaiseError ProjectError.ElementNotFound, "Translations worksheet '" & TRANSLATIONS_SHEET & "' is missing."
    End If

    Set TranslationSheet = sh
End Function

Private Function ResolveLanguageRange(ByVal translator As IDesTranslation) As Range
    Dim candidate As Range
    Dim fallback As Range

    On Error Resume Next
        Set candidate = translator.LangListRng
    On Error GoTo 0

    If Not candidate Is Nothing Then
        Set ResolveLanguageRange = candidate
        Exit Function
    End If

    On Error Resume Next
        Set fallback = TranslationSheet.Range(LANGUAGE_CODE_RANGE)
    On Error GoTo 0

    If fallback Is Nothing Then
        RaiseError ProjectError.ElementNotFound, "Language code range '" & LANGUAGE_CODE_RANGE & "' was not found."
    End If

    Set ResolveLanguageRange = fallback
End Function

Private Sub RaiseError(ByVal errorCode As ProjectError, ByVal messageText As String)
    Err.Raise CLng(errorCode), CLASS_NAME, messageText
End Sub


'@section Interface implementation
'===============================================================================
Private Sub IDesignerEntry_Clear()
    Clear
End Sub

Private Sub IDesignerEntry_Configure(ByVal hostSheet As Worksheet)
    Configure hostSheet
End Sub

Private Property Get IDesignerEntry_HostSheet() As Worksheet
    Set IDesignerEntry_HostSheet = HostSheet
End Property

Private Sub IDesignerEntry_Translate(Optional ByVal languageCode As String = vbNullString)
    Translate languageCode
End Sub

Private Sub IDesignerEntry_UseEntryManager(ByVal entryManager As IMain)
    UseEntryManager entryManager
End Sub

Private Sub IDesignerEntry_UseTranslator(ByVal translator As IDesTranslation)
    UseTranslator translator
End Sub
