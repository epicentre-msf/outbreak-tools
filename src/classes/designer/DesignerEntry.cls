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
'@depends ProjectError, DesignerTranslation, IDesignerTranslation
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements IDesignerEntry

'@class DesignerEntry
'@description
'Coordinates clearing, translation, and info read/write actions for a
'designer entry worksheet. Lazily resolves an IDesignerTranslation
'translator (for applying language translations and translating message
'codes) from the host worksheet and its parent workbook. The translator
'can be replaced via setter injection for testing.
'Consumers interact through the IDesignerEntry interface.
'@depends DesignerTranslation, IDesignerTranslation


'@section State
'===============================================================================

Private Type TDesignerEntry
    hostSheet As Worksheet      'Worksheet hosting designer inputs
    translator As IDesignerTranslation 'Translation handler for the designer workbook
End Type

Private this As TDesignerEntry

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "DesignerEntry"
Private Const TRANSLATIONS_SHEET As String = "DesignerTranslation"

'Named ranges on the Main worksheet (matching the designer layout)
Private Const RNG_EDITION As String = "RNG_Edition"
Private Const RNG_PATH_GEO As String = "RNG_PathGeo"
Private Const RNG_PATH_DICO As String = "RNG_PathDico"
Private Const RNG_LL_DIR As String = "RNG_LLDir"
Private Const RNG_LL_NAME As String = "RNG_LLName"
Private Const RNG_LANG_SETUP As String = "RNG_LangSetup"
Private Const RNG_TEMP_PATH As String = "RNG_LLTemp"
Private Const RNG_LL_FORM As String = "RNG_LLForm"
Private Const RNG_WARNING As String = "RNG_Warning"


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
'worksheet is Nothing. Does not reset the translator so that a
'previously injected test double is preserved.
'@param hostSheet Worksheet. The worksheet hosting designer inputs.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@export
Public Sub Configure(ByVal hostSheet As Worksheet)
    If hostSheet Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host sheet is required."
    End If

    Set this.hostSheet = hostSheet
End Sub

'@label:use-translator
'@sub-title Use a provided translations manager instead of creating one
'@details
'Replaces the internally resolved IDesignerTranslation translator with the
'supplied instance. Primarily used in tests to inject a stub or mock.
'@param translator IDesignerTranslation. The translations provider to use.
'@export
Public Sub UseTranslator(ByVal translator As IDesignerTranslation)
    Set this.translator = translator
End Sub

'@label:clear
'@sub-title Clear user inputs on the configured worksheet
'@details
'Resets all input ranges on the configured worksheet by clearing their
'values and setting backgrounds to white. Also clears any warning
'messages below the RNG_Warning range. Silently skips any named range
'that does not exist on the worksheet.
'@export
Public Sub Clear()
    ClearRange RNG_PATH_GEO
    ClearRange RNG_PATH_DICO
    ClearRange RNG_LL_NAME
    ClearRange RNG_LL_DIR
    ClearRange RNG_EDITION
    ClearRange RNG_TEMP_PATH
    ClearRange RNG_LANG_SETUP
    ClearRange RNG_LL_FORM
    ClearWarnings
End Sub

'@label:translate
'@sub-title Translate the configured worksheet using the provided language
'@details
'Delegates to the translations provider to apply translated labels,
'shapes, and dropdowns to the host worksheet. When a non-empty language
'string is supplied (e.g. "FRA - Francais"), the translator extracts
'the language code and persists it to the worksheet-level HiddenName.
'@param language Optional String. Language string in "CODE - Label" format. Defaults to vbNullString.
'@export
Public Sub Translate(Optional ByVal language As String = vbNullString)
    Dim translator As IDesignerTranslation
    Set translator = ResolveTranslator()
    translator.TranslateDesigner HostSheet, language
End Sub

'@label:add-info
'@sub-title Write a value to a named info range on the host worksheet
'@details
'Maps the logical infoName to its corresponding named range on the host
'worksheet, writes the value, and sets the background to white. Silently
'exits when the named range does not exist on the worksheet.
'@param value String. The value to write (path, translated message, etc.).
'@param infoName String. The logical name identifying the target range.
'@export
Public Sub AddInfo(ByVal value As String, ByVal infoName As String)
    Dim targetRange As Range

    Set targetRange = ResolveInfoRange(infoName)
    If targetRange Is Nothing Then Exit Sub

    targetRange.value = value
    targetRange.Interior.Color = vbWhite
End Sub

'@label:value-of
'@prop-title Read a value from a named info range on the host worksheet
'@details
'Maps the logical infoName to its corresponding named range on the host
'worksheet and returns the current cell value. Returns vbNullString when
'the named range does not exist.
'@param infoName String. The logical name identifying the source range.
'@return String. The current value of the named range.
'@export
Public Property Get ValueOf(ByVal infoName As String) As String
    Dim targetRange As Range

    Set targetRange = ResolveInfoRange(infoName)
    If targetRange Is Nothing Then Exit Property

    On Error Resume Next
    ValueOf = CStr(targetRange.value)
    On Error GoTo 0
End Property

'@label:translate-message
'@sub-title Translate a message code using the designer translations
'@details
'Delegates to the lazily resolved translations provider to translate a
'message code into the current designer language. Returns the raw
'msgCode unchanged when no translation is available.
'@param msgCode String. The message code to translate.
'@return String. The translated message text.
'@export
Public Function TranslateMessage(ByVal msgCode As String) As String
    Dim translator As IDesignerTranslation
    Dim result As String

    Set translator = ResolveTranslator()
    result = translator.TranslatedValue(msgCode)

    'Fall back to raw code when translation returns empty
    If LenB(result) = 0 Then result = msgCode
    TranslateMessage = result
End Function

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

'@label:clear-range
'@sub-title Clear a single named range value and reset its background
'@details
'Resolves the named range from the host worksheet, sets its value to
'empty, and resets the background to white. Silently skips when the
'named range does not exist.
'@param rngName String. The named range identifier on the worksheet.
Private Sub ClearRange(ByVal rngName As String)
    Dim rng As Range

    On Error Resume Next
    Set rng = this.hostSheet.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    rng.value = vbNullString
    rng.Interior.Color = vbWhite
End Sub

'@label:clear-warnings
'@sub-title Clear the warning range and all contiguous non-empty cells below it
'@details
'Resolves RNG_Warning and walks downward, clearing the value and
'background of each non-empty cell until an empty cell is encountered.
'Silently exits when RNG_Warning does not exist.
Private Sub ClearWarnings()
    Dim rng As Range

    On Error Resume Next
    Set rng = this.hostSheet.Range(RNG_WARNING)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Do While Not IsEmpty(rng)
        rng.value = vbNullString
        rng.Interior.Color = vbWhite
        Set rng = rng.Offset(1)
    Loop
End Sub

'@label:resolve-info-range
'@sub-title Map a logical info name to its named range on the host worksheet
'@details
'Resolves an infoName string to the corresponding named range constant,
'then returns the Range object from the host worksheet. Returns Nothing
'when the infoName is unknown or the named range does not exist.
'@param infoName String. The logical info name to resolve.
'@return Range. The named range, or Nothing.
Private Function ResolveInfoRange(ByVal infoName As String) As Range
    Dim rngName As String

    Select Case infoName
    Case "edition"
        rngName = RNG_EDITION
    Case "setuppath"
        rngName = RNG_PATH_DICO
    Case "geopath"
        rngName = RNG_PATH_GEO
    Case "lldir"
        rngName = RNG_LL_DIR
    Case "llname"
        rngName = RNG_LL_NAME
    Case "setuplang"
        rngName = RNG_LANG_SETUP
    Case "temppath"
        rngName = RNG_TEMP_PATH
    Case "lllang"
        rngName = RNG_LL_FORM
    Case "warning"
        rngName = RNG_WARNING
    Case Else
        Exit Function
    End Select

    On Error Resume Next
    Set ResolveInfoRange = this.hostSheet.Range(rngName)
    On Error GoTo 0
End Function

'@label:resolve-translator
'@sub-title Lazily resolve or return the translations provider
'@details
'Returns the cached IDesignerTranslation translator, or creates one via
'DesignerTranslation.Create with the DesignerTranslation worksheet on first
'access. Raises an error when the translation sheet cannot be found.
'@return IDesignerTranslation. The translations provider instance.
'@throws ProjectError.ElementNotFound When the translation sheet is missing.
Private Function ResolveTranslator() As IDesignerTranslation
    If this.translator Is Nothing Then
        On Error GoTo Handler
            Set this.translator = DesignerTranslation.Create(TranslationSheet())
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

Private Sub IDesignerEntry_AddInfo(ByVal value As String, ByVal infoName As String)
    AddInfo value, infoName
End Sub

Private Sub IDesignerEntry_Clear()
    Clear
End Sub

Private Sub IDesignerEntry_Configure(ByVal hostSheet As Worksheet)
    Configure hostSheet
End Sub

Private Property Get IDesignerEntry_HostSheet() As Worksheet
    Set IDesignerEntry_HostSheet = HostSheet
End Property

Private Sub IDesignerEntry_Translate(Optional ByVal language As String = vbNullString)
    Translate language
End Sub

Private Function IDesignerEntry_TranslateMessage(ByVal msgCode As String) As String
    IDesignerEntry_TranslateMessage = TranslateMessage(msgCode)
End Function

Private Sub IDesignerEntry_UseTranslator(ByVal translator As IDesignerTranslation)
    UseTranslator translator
End Sub

Private Property Get IDesignerEntry_ValueOf(ByVal infoName As String) As String
    IDesignerEntry_ValueOf = ValueOf(infoName)
End Property
