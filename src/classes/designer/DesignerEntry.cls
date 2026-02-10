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

'@class DesignerEntry
'@description
'Coordinates clearing and translation actions for a designer entry
'worksheet. Lazily resolves an IMain entry manager (for clearing input
'ranges) and an IDesTranslation translator (for applying language
'translations) from the host worksheet and its parent workbook. Both
'dependencies can be replaced via setter injection for testing.
'Consumers interact through the IDesignerEntry interface.
'@depends Main, IMain, DesTranslation, IDesTranslation


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

'@label:create
'@sub-title Create and configure a designer entry helper
'@details
'Factory method that creates a new DesignerEntry instance, binds it to
'the supplied worksheet via Configure, and returns the result through
'the IDesignerEntry interface. Called on the predeclared instance.
'@param hostSheet Worksheet. The worksheet hosting designer inputs.
'@return IDesignerEntry. A fully initialised instance ready for use.
'@export
Public Function Create(ByVal hostSheet As Worksheet) As IDesignerEntry
    Dim instance As DesignerEntry

    Set instance = New DesignerEntry
    instance.Configure hostSheet
    Set Create = instance.Self
End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the
'interface reference from the predeclared Create method.
'@return IDesignerEntry. The current instance cast to the interface.
Public Property Get Self() As IDesignerEntry
    Set Self = Me
End Property


'@section Public API
'===============================================================================

'@label:configure
'@sub-title Bind the helper to the target worksheet
'@details
'Stores a reference to the host worksheet. Raises an error when the
'worksheet is Nothing. Does not reset the entry manager or translator
'so that previously injected test doubles are preserved.
'@param hostSheet Worksheet. The worksheet hosting designer inputs.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@export
Public Sub Configure(ByVal hostSheet As Worksheet)
    If hostSheet Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host sheet is required."
    End If

    Set this.hostSheet = hostSheet
End Sub

'@label:use-entry-manager
'@sub-title Use a provided entry manager instead of creating one
'@details
'Replaces the internally resolved IMain entry manager with the supplied
'instance. Primarily used in tests to inject a stub or mock.
'@param entryManager IMain. The entry manager to use.
'@export
Public Sub UseEntryManager(ByVal entryManager As IMain)
    Set this.entryManager = entryManager
End Sub

'@label:use-translator
'@sub-title Use a provided translations manager instead of creating one
'@details
'Replaces the internally resolved IDesTranslation translator with the
'supplied instance. Primarily used in tests to inject a stub or mock.
'@param translator IDesTranslation. The translations provider to use.
'@export
Public Sub UseTranslator(ByVal translator As IDesTranslation)
    Set this.translator = translator
End Sub

'@label:clear
'@sub-title Clear user inputs on the configured worksheet
'@details
'Delegates to the lazily resolved entry manager to clear all input
'ranges on the configured worksheet, removing user-entered values.
'@export
Public Sub Clear()
    ResolveEntryManager.ClearInputRanges clearValues:=True
End Sub

'@label:translate
'@sub-title Translate the configured worksheet and optionally update the language code
'@details
'When a non-empty language code is supplied, resolves the language code
'range on the translation sheet and writes the new code, then
'recalculates the sheet so translation formulas update. Finally
'delegates to the translations provider to apply translated labels,
'shapes, and dropdowns to the host worksheet.
'@param languageCode Optional String. ISO language code to apply. Defaults to vbNullString.
'@export
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

'@label:host-sheet
'@prop-title The managed worksheet
'@details
'Returns the designer entry worksheet that was bound via Configure.
'Raises an error when accessed before configuration.
'@return Worksheet. The host worksheet.
'@throws ProjectError.ObjectNotInitialized When the worksheet was not configured.
'@export
Public Property Get HostSheet() As Worksheet
    If this.hostSheet Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "HostSheet requested before configuration."
    End If
    Set HostSheet = this.hostSheet
End Property


'@section Internal Helpers
'===============================================================================

'@label:resolve-entry-manager
'@sub-title Lazily resolve or return the entry manager
'@details
'Returns the cached IMain entry manager, or creates one via
'Main.Create(HostSheet) on first access. Raises an error when the entry
'manager cannot be initialised.
'@return IMain. The entry manager instance.
'@throws ProjectError.ElementNotFound When the entry manager cannot be initialised.
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

'@label:resolve-translator
'@sub-title Lazily resolve or return the translations provider
'@details
'Returns the cached IDesTranslation translator, or creates one via
'DesTranslation.Create with the DesignerTranslation worksheet on first
'access. Raises an error when the translation sheet cannot be found.
'@return IDesTranslation. The translations provider instance.
'@throws ProjectError.ElementNotFound When the translation sheet is missing.
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

'@label:translation-sheet
'@sub-title Locate the DesignerTranslation worksheet in the parent workbook
'@details
'Searches the parent workbook of the host worksheet for a sheet named
'"DesignerTranslation". Raises an error when not found.
'@return Worksheet. The translations worksheet.
'@throws ProjectError.ElementNotFound When the translations worksheet is missing.
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

'@label:resolve-language-range
'@sub-title Resolve the language code range from the translator or translation sheet
'@details
'First attempts to read the language range from the translator's
'LangListRng property. If that fails, falls back to the
'RNG_MainLangCode named range on the translation sheet. Raises an error
'when neither source yields a range.
'@param translator IDesTranslation. The translator whose LangListRng is tried first.
'@return Range. The language code range.
'@throws ProjectError.ElementNotFound When no language code range is found.
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

'@label:raise-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME.
'@param errorCode ProjectError. The error code to raise.
'@param messageText String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub RaiseError(ByVal errorCode As ProjectError, ByVal messageText As String)
    Err.Raise CLng(errorCode), CLASS_NAME, messageText
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IDesignerEntry contract.
'See the corresponding members above for full documentation.

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
