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
'@depends ProjectError, DesignerTranslation, Checking
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit

'@class DesignerEntry
'@description
'Coordinates clearing, translation, validation, and info read/write
'actions for a designer entry worksheet. Lazily resolves a
'DesignerTranslation translator (for applying language translations and
'translating message codes) from the host worksheet and its parent
'workbook. A test injects its translator through UseTranslator before
'the first resolve.
'@depends DesignerTranslation, Checking


'@section State
'===============================================================================

Private Type TDesignerEntry
    hostSheet As Worksheet      'Worksheet hosting designer inputs
    translator As DesignerTranslation 'Translation handler for the designer workbook
    sealed As Boolean           'True once the factory has finished setup
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
Private Const RNG_DESIGN_LL As String = "RNG_DesignLL"
Private Const RNG_WARNING As String = "RNG_Warning"

'The two password ranges and the epiweek range LinelistSpecs.Value reads
'at generation. The multi driver writes each T_Multi row through these
'keys, so the row's passwords and epiweek reach the build the same way
'the typed entries do.
Private Const RNG_PWD_OPEN As String = "RNG_LLPwdOpen"
Private Const RNG_PWD_DEBUG As String = "RNG_LLPassword"
Private Const RNG_EPIWEEK As String = "RNG_DefaultEpiWeek"

'Leading shape of the internal tag columns of a setup file. A tag column
'header can reach the setup language dropdown through the fallback read,
'and Validate refuses it as a language.
Private Const INTERNAL_TAG_LEAD As String = "__"


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create and configure a designer entry helper
'@details
'Factory method that creates a new DesignerEntry instance, binds it to
'the supplied worksheet via Configure, and returns the result. Called
'on the predeclared instance.
'@param hostSheet Worksheet. The worksheet hosting designer inputs.
'@return DesignerEntry. A fully initialised instance ready for use.
'@export
Public Function Create(ByVal hostSheet As Worksheet) As DesignerEntry
    Dim instance As DesignerEntry

    Set instance = New DesignerEntry
    instance.Configure hostSheet
    instance.Seal
    Set Create = instance
End Function


'@section Instance Sealing
'===============================================================================

'@label:seal
'@sub-title Seal the instance against further setup writes
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup-only member against post-seal writes
'@param memberName String. The member name for the error message.
Private Sub GuardNotSealed(ByVal memberName As String)
    If this.sealed Then _
        RaiseError ProjectError.SomethingWentWrong, memberName & " is set at creation only"
End Sub


'@section Public API
'===============================================================================

'@label:configure
'@sub-title Bind the helper to the target worksheet
'@details
'Stores a reference to the host worksheet. Called once by the factory,
'which then seals the instance; a later call raises. Raises an error
'when the worksheet is Nothing.
'@param hostSheet Worksheet. The worksheet hosting designer inputs.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
'@export
Public Sub Configure(ByVal hostSheet As Worksheet)
    GuardNotSealed "Configure"

    If hostSheet Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host sheet is required."
    End If

    Set this.hostSheet = hostSheet
End Sub

'@label:use-translator
'@sub-title Inject the translations manager the entry will use
'@details
'Sets the DesignerTranslation translator the entry delegates to. A test
'injects its own translator here right after Create. The setter guards
'itself: it refuses Nothing, and it refuses to replace a translator that
'is already held (injected earlier, or lazily resolved by a translate
'call).
'@param translator DesignerTranslation. The translations provider to use.
'@throws ProjectError.ObjectNotInitialized When the translator is Nothing.
'@throws ProjectError.SomethingWentWrong When a translator is already held.
'@export
Public Sub UseTranslator(ByVal translator As DesignerTranslation)
    If translator Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Translator is required."
    End If

    If Not this.translator Is Nothing Then
        RaiseError ProjectError.SomethingWentWrong, "A translator is already held."
    End If

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
    Dim translator As DesignerTranslation
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

'@label:validate
'@sub-title Check the generation entries on the host worksheet
'@details
'Reads the generation entries from the Main named ranges and files one
'entry per fault into a Checking, every entry with the error scope. The
'checks: the setup path is set and on disk, the output folder is set
'and on disk, the linelist name is set, the setup language is set and
'is free of the internal double-underscore tag shape, the linelist
'language is set, the design value is set, and the template and geobase
'paths are on disk when they are set. An empty checking means every
'entry passes. A named range missing from the worksheet reads as an
'empty value and files the matching fault.
'@return Checking. One entry per faulty generation entry.
'@export
Public Function Validate() As Checking
    Dim checks As Checking
    Dim entryValue As String

    Set checks = Checking.Create("Generation entries", _
                                 "Checks run on the Main entries before the build")

    'Setup file: required, and the file has to be on disk
    entryValue = ValueOf("setuppath")
    If LenB(entryValue) = 0 Then
        checks.Add "setup path", _
                   "The setup file path is empty. Load a setup file first.", _
                   checkingError
    ElseIf LenB(Dir(entryValue)) = 0 Then
        checks.Add "setup path", _
                   "The setup file was not found: " & entryValue, _
                   checkingError
    End If

    'Output folder: required, and the folder has to be on disk
    entryValue = ValueOf("lldir")
    If LenB(entryValue) = 0 Then
        checks.Add "output folder", _
                   "The output folder is empty. Pick where the linelist will be written.", _
                   checkingError
    ElseIf LenB(Dir(entryValue, vbDirectory)) = 0 Then
        checks.Add "output folder", _
                   "The output folder was not found: " & entryValue, _
                   checkingError
    End If

    'Linelist name: required
    If LenB(ValueOf("llname")) = 0 Then
        checks.Add "linelist name", "The linelist name is empty.", checkingError
    End If

    'Setup language: required, and an internal tag column is refused
    entryValue = ValueOf("setuplang")
    If LenB(entryValue) = 0 Then
        checks.Add "setup language", _
                   "The setup language is empty. Load a setup file and pick a language.", _
                   checkingError
    ElseIf Left$(entryValue, Len(INTERNAL_TAG_LEAD)) = INTERNAL_TAG_LEAD Then
        checks.Add "setup language", _
                   "The setup language reads as an internal tag: " & entryValue & _
                   ". Pick a language from the dropdown.", _
                   checkingError
    End If

    'Linelist language: required
    If LenB(ValueOf("lllang")) = 0 Then
        checks.Add "linelist language", _
                   "The linelist language is empty. Pick the language of the generated file.", _
                   checkingError
    End If

    'Design value: required
    If LenB(ValueOf("design")) = 0 Then
        checks.Add "design", _
                   "The design value is empty. Pick a design for the linelist.", _
                   checkingError
    End If

    'Template file: optional, and the file has to be on disk when set
    entryValue = ValueOf("temppath")
    If LenB(entryValue) > 0 Then
        If LenB(Dir(entryValue)) = 0 Then
            checks.Add "template", _
                       "The template file was not found: " & entryValue, _
                       checkingError
        End If
    End If

    'Geobase file: optional, and the file has to be on disk when set
    entryValue = ValueOf("geopath")
    If LenB(entryValue) > 0 Then
        If LenB(Dir(entryValue)) = 0 Then
            checks.Add "geobase", _
                       "The geobase file was not found: " & entryValue, _
                       checkingError
        End If
    End If

    Set Validate = checks
End Function

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
    Dim translator As DesignerTranslation
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
    Case "design"
        rngName = RNG_DESIGN_LL
    Case "warning"
        rngName = RNG_WARNING
    Case "llpassword"
        rngName = RNG_PWD_OPEN
    Case "debugpassword"
        rngName = RNG_PWD_DEBUG
    Case "epiweekstart"
        rngName = RNG_EPIWEEK
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
'Returns the cached DesignerTranslation translator, or creates one via
'DesignerTranslation.Create with the DesignerTranslation worksheet on first
'access. Raises an error when the translation sheet cannot be found.
'@return DesignerTranslation. The translations provider instance.
'@throws ProjectError.ElementNotFound When the translation sheet is missing.
Private Function ResolveTranslator() As DesignerTranslation
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
