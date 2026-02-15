VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Prepares designer workbooks and ribbon settings"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Prepares designer workbooks and persists ribbon settings through hidden names.")
'@depends ProjectError, HiddenNames, IHiddenNames, IDevelopment, OSFiles, IOSFiles, CustomTable, ICustomTable, BetterArray
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements IDesignerPreparation

'@class DesignerPreparation
'@description
'Prepares a designer workbook for ribbon interactions. The full
'preparation sequence imports translation tables from a user-selected
'workbook, hides internal worksheets (VeryHidden for __pass/__formatter/
'__formula, Hidden for translation sheets), seeds workbook- and
'worksheet-level HiddenNames flags, and registers hidden sheets with
'the development manager. Also provides methods to read and write
'ribbon checkbox flags persisted through workbook-level hidden names.
'@depends HiddenNames, IHiddenNames, IDevelopment, OSFiles, IOSFiles, CustomTable, ICustomTable, BetterArray


'@section State
'===============================================================================

Private Type TDesignerPreparation
    hostBook As Workbook       'Designer workbook being prepared
    store As IHiddenNames      'Workbook-level hidden names manager
End Type

Private this As TDesignerPreparation

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "DesignerPreparation"
Private Const FLAG_ALERT As String = "chkAlert"
Private Const FLAG_INSTRUCT As String = "chkInstruct"
Private Const FLAG_ON As String = "Yes"
Private Const FLAG_OFF As String = "No"

'Sheet names
Private Const SHEET_PASS As String = "__pass"
Private Const SHEET_FORMATTER As String = "__formatter"
Private Const SHEET_FORMULA As String = "__formula"
Private Const SHEET_LL_TRANSLATION As String = "LinelistTranslation"
Private Const SHEET_DES_TRANSLATION As String = "DesignerTranslation"
Private Const SHEET_GEO As String = "Geo"

'Workbook-level HiddenNames
Private Const TAG_DES_LANG As String = "TAG_DES_LANG"
Private Const TAG_LL_LANG_CODE As String = "RNG_LLLanguageCode"
Private Const TAG_DICT_LANG As String = "RNG_DictionaryLanguage"
Private Const TAG_LAST_OPENED As String = "RNG_LastOpenedDate"

'Geo worksheet-level HiddenNames
Private Const GEO_LANG_CODE As String = "RNG_GeoLangCode"
Private Const GEO_NAME As String = "RNG_GeoName"
Private Const GEO_META_LANG As String = "RNG_MetaLang"
Private Const GEO_UPDATED As String = "RNG_GeoUpdated"
Private Const GEO_PASTING_COL As String = "RNG_PastingGeoCol"
Private Const GEO_FORM_LOADED As String = "RNG_FormLoaded"


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create and configure a designer preparation helper
'@details
'Factory method that creates a new DesignerPreparation instance, binds
'it to the supplied designer workbook via Configure, and returns the
'result through the IDesignerPreparation interface. Called on the
'predeclared instance.
'@param hostBook Workbook. The designer workbook to prepare.
'@return IDesignerPreparation. A fully initialised instance ready for use.
'@export
Public Function Create(ByVal hostBook As Workbook) As IDesignerPreparation
    Dim instance As DesignerPreparation

    Set instance = New DesignerPreparation
    instance.Configure hostBook
    Set Create = instance.Self
End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the
'interface reference from the predeclared Create method.
'@return IDesignerPreparation. The current instance cast to the interface.
Public Property Get Self() As IDesignerPreparation
    Set Self = Me
End Property


'@section Public API
'===============================================================================

'@label:configure
'@sub-title Bind the helper to the designer workbook
'@details
'Stores a reference to the host workbook and clears any previously
'cached hidden names store so it will be rebuilt on next access. Raises
'an error when the workbook is Nothing.
'@param hostBook Workbook. The designer workbook to bind to.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook is required."
    End If

    Set this.hostBook = hostBook
    Set this.store = Nothing
End Sub

'@label:prepare
'@sub-title Execute all preparation tasks on the designer workbook
'@details
'Runs the full preparation sequence:
'1. Prompts the user to select a translations workbook and imports
'   translation tables into LinelistTranslation and DesignerTranslation.
'2. Hides internal worksheets (VeryHidden for __pass, __formatter,
'   __formula; Hidden for LinelistTranslation, DesignerTranslation).
'3. Seeds workbook-level HiddenNames flags (ribbon checks, language
'   codes, last opened date).
'4. Seeds Geo worksheet-level HiddenNames flags (geo language, name,
'   update status).
'When devManager is provided, hidden sheets are also registered with
'the manager for deployment.
'@param devManager IDevelopment. ByRef development manager instance, or Nothing.
'@export
Public Sub Prepare(ByRef devManager As IDevelopment)
    ImportTranslations
    HideWorksheets devManager
    EnsureHiddenStore
    EnsureWorkbookFlags
    EnsureGeoFlags
End Sub

'@label:set-flag
'@sub-title Persist a ribbon checkbox flag as Yes/No in hidden names
'@details
'Writes a boolean value to the workbook-level hidden names store as
'"Yes" or "No". Ensures the store and default flags exist before writing.
'@param nameId String. The hidden name identifier for the flag (e.g. "chkAlert").
'@param enabled Boolean. True to store "Yes", False to store "No".
'@export
Public Sub SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
    EnsureHiddenStore
    EnsureDefaultFlags
    this.store.SetValue nameId, IIf(enabled, FLAG_ON, FLAG_OFF)
End Sub

'@label:get-flag
'@sub-title Read a ribbon checkbox flag as a boolean from hidden names
'@details
'Reads the named value from the hidden names store and returns True when
'the stored value equals "Yes" (case-insensitive comparison). Falls back
'to the defaultValue when the name is missing.
'@param nameId String. The hidden name identifier for the flag.
'@param defaultValue Optional Boolean. Value returned when the flag is missing. Defaults to False.
'@return Boolean. True when the stored value is "Yes".
'@export
Public Function GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim rawValue As String
    EnsureHiddenStore
    EnsureDefaultFlags

    rawValue = this.store.ValueAsString(nameId, IIf(defaultValue, FLAG_ON, FLAG_OFF))
    GetFlag = (StrComp(rawValue, FLAG_ON, vbTextCompare) = 0)
End Function

'@label:host-workbook
'@prop-title The configured host workbook
'@details
'Returns the designer workbook that was bound via Configure. Raises an
'error when accessed before configuration.
'@return Workbook. The host workbook.
'@throws ProjectError.ObjectNotInitialized When the workbook was not configured.
'@export
Public Property Get HostWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook was not configured."
    End If
    Set HostWorkbook = this.hostBook
End Property

'@label:hidden-store
'@prop-title The workbook-level hidden names manager
'@details
'Returns the IHiddenNames instance lazily created from the host workbook.
'Used internally to persist and retrieve ribbon flag values.
'@return IHiddenNames. The hidden names manager.
'@export
Public Property Get HiddenStore() As IHiddenNames
    EnsureHiddenStore
    Set HiddenStore = this.store
End Property


'@section Preparation Helpers
'===============================================================================

'@label:import-translations
'@sub-title Import translation tables from a user-selected workbook
'@details
'Prompts the user to select an xlsx file, then iterates through
'LinelistTranslation and DesignerTranslation sheets. For each sheet,
'imports matching ListObjects via CustomTable.Import. This is the same
'logic as the clickImpTrans ribbon callback, encapsulated here so it
'runs as part of the preparation sequence.
Private Sub ImportTranslations()
    Dim io As IOSFiles
    Dim importBook As Workbook
    Dim wb As Workbook
    Dim sheetNames As BetterArray
    Dim tableNames As BetterArray
    Dim idx As Long
    Dim sheetName As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim lo As ListObject
    Dim targetTable As ICustomTable
    Dim sourceTable As ICustomTable

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    Set wb = HostWorkbook

    On Error GoTo ImportCleanup
    Set importBook = Workbooks.Open(io.File(), ReadOnly:=True)

    Set sheetNames = New BetterArray
    sheetNames.Push SHEET_LL_TRANSLATION, SHEET_DES_TRANSLATION

    Set tableNames = New BetterArray
    tableNames.Push "t_tradllshapes", "t_tradllmsg", "t_tradllforms", "t_tradllribbon", _
                    "t_tradmsg", "t_tradrange", "t_tradshape"

    For idx = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(idx)

        On Error Resume Next
        Set targetSheet = wb.Worksheets(CStr(sheetName))
        Set sourceSheet = importBook.Worksheets(CStr(sheetName))
        On Error GoTo ImportCleanup

        If (Not targetSheet Is Nothing) And (Not sourceSheet Is Nothing) Then
            For Each lo In targetSheet.ListObjects
                If tableNames.Includes(LCase$(lo.Name)) Then
                    Set targetTable = CustomTable.Create(lo)

                    On Error Resume Next
                    Set sourceTable = Nothing
                    Set sourceTable = CustomTable.Create(sourceSheet.ListObjects(lo.Name))
                    On Error GoTo ImportCleanup

                    If Not sourceTable Is Nothing Then
                        targetTable.Import sourceTable
                    End If
                End If
            Next lo
        End If

        Set targetSheet = Nothing
        Set sourceSheet = Nothing
    Next idx

ImportCleanup:
    If Not importBook Is Nothing Then importBook.Close saveChanges:=False
    If Err.Number <> 0 Then
        RaiseError ProjectError.SomethingWentWrong, _
                   "Unable to import translations: " & Err.Description
    End If
End Sub

'@label:hide-worksheets
'@sub-title Hide internal worksheets and register them with the dev manager
'@details
'Sets __pass, __formatter, __formula to xlSheetVeryHidden. Sets
'LinelistTranslation and DesignerTranslation to xlSheetHidden (visible
'in the VBA editor for developers). When a development manager is
'provided, all five sheets are also registered via AddHiddenSheet for
'deployment. Silently skips any worksheet that does not exist.
'@param devManager IDevelopment. Development manager for deployment registration, or Nothing.
Private Sub HideWorksheets(ByRef devManager As IDevelopment)
    SetSheetVisibility SHEET_PASS, xlSheetVeryHidden
    SetSheetVisibility SHEET_FORMATTER, xlSheetVeryHidden
    SetSheetVisibility SHEET_FORMULA, xlSheetVeryHidden
    SetSheetVisibility SHEET_LL_TRANSLATION, xlSheetHidden
    SetSheetVisibility SHEET_DES_TRANSLATION, xlSheetHidden

    If Not devManager Is Nothing Then
        RegisterHiddenSheet devManager, SHEET_PASS
        RegisterHiddenSheet devManager, SHEET_FORMATTER
        RegisterHiddenSheet devManager, SHEET_FORMULA
        RegisterHiddenSheet devManager, SHEET_LL_TRANSLATION
        RegisterHiddenSheet devManager, SHEET_DES_TRANSLATION
    End If
End Sub

'@label:ensure-workbook-flags
'@sub-title Ensure all workbook-level HiddenNames flags exist
'@details
'Creates workbook-level HiddenNames for ribbon checkbox flags, language
'codes, and the last opened date. Checkbox flags default to "Yes" (on).
'Language codes default to empty strings. Last opened date is set to the
'current date.
Private Sub EnsureWorkbookFlags()
    EnsureDefaultFlags
    this.store.EnsureName TAG_DES_LANG, vbNullString, HiddenNameTypeString
    this.store.EnsureName TAG_LL_LANG_CODE, vbNullString, HiddenNameTypeString
    this.store.EnsureName TAG_DICT_LANG, vbNullString, HiddenNameTypeString
    this.store.EnsureName TAG_LAST_OPENED, Format$(Now, "yyyy-mm-dd"), HiddenNameTypeString
End Sub

'@label:ensure-geo-flags
'@sub-title Ensure Geo worksheet-level HiddenNames flags exist
'@details
'Creates worksheet-level HiddenNames on the Geo worksheet for language
'code, geobase name, metadata language, update status, pasting column,
'and form loaded state. Silently exits when the Geo worksheet does not
'exist in the host workbook.
Private Sub EnsureGeoFlags()
    Dim geoSheet As Worksheet
    Dim geoStore As IHiddenNames

    Set geoSheet = ResolveWorksheet(SHEET_GEO)
    If geoSheet Is Nothing Then Exit Sub

    Set geoStore = HiddenNames.Create(geoSheet)
    geoStore.EnsureName GEO_LANG_CODE, vbNullString, HiddenNameTypeString
    geoStore.EnsureName GEO_NAME, vbNullString, HiddenNameTypeString
    geoStore.EnsureName GEO_META_LANG, vbNullString, HiddenNameTypeString
    geoStore.EnsureName GEO_UPDATED, "empty", HiddenNameTypeString
    geoStore.EnsureName GEO_PASTING_COL, vbNullString, HiddenNameTypeString
    geoStore.EnsureName GEO_FORM_LOADED, vbNullString, HiddenNameTypeString
End Sub


'@section Internal Helpers
'===============================================================================

'@label:ensure-hidden-store
'@sub-title Lazily initialise the hidden names store from the host workbook
'@details
'Creates the IHiddenNames instance on first access by calling
'HiddenNames.Create with the host workbook. Subsequent calls are no-ops.
Private Sub EnsureHiddenStore()
    If this.store Is Nothing Then
        Set this.store = HiddenNames.Create(HostWorkbook)
    End If
End Sub

'@label:ensure-default-flags
'@sub-title Ensure both default ribbon flags exist in the hidden store
'@details
'Calls EnsureName for chkAlert and chkInstruct with "Yes" as default,
'creating the hidden names if they do not already exist. All checks
'default to enabled per specification.
Private Sub EnsureDefaultFlags()
    this.store.EnsureName FLAG_ALERT, FLAG_ON, HiddenNameTypeString
    this.store.EnsureName FLAG_INSTRUCT, FLAG_ON, HiddenNameTypeString
End Sub

'@label:resolve-worksheet
'@sub-title Resolve a worksheet by name from the host workbook
'@details
'Returns the worksheet with the given name from the host workbook, or
'Nothing when the sheet does not exist.
'@param sheetName String. The worksheet name to resolve.
'@return Worksheet. The resolved worksheet, or Nothing.
Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set ResolveWorksheet = HostWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:set-sheet-visibility
'@sub-title Set the visibility of a worksheet by name
'@details
'Resolves the worksheet from the host workbook and sets its Visible
'property. Silently skips when the worksheet does not exist.
'@param sheetName String. The worksheet name.
'@param visibility Long. The xlSheetVisibility constant to apply.
Private Sub SetSheetVisibility(ByVal sheetName As String, ByVal visibility As Long)
    Dim sh As Worksheet
    Set sh = ResolveWorksheet(sheetName)
    If Not sh Is Nothing Then sh.Visible = visibility
End Sub

'@label:register-hidden-sheet
'@sub-title Register a sheet with the development manager for deployment hiding
'@details
'Wraps manager.AddHiddenSheet in error handling so a missing sheet does
'not abort the preparation sequence.
'@param manager IDevelopment. The development manager to register with.
'@param sheetName String. The worksheet name to register.
Private Sub RegisterHiddenSheet(ByRef manager As IDevelopment, ByVal sheetName As String)
    On Error Resume Next
    manager.AddHiddenSheet sheetName
    On Error GoTo 0
End Sub

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
'@description Delegated members satisfying the IDesignerPreparation contract.
'See the corresponding members above for full documentation.

Private Sub IDesignerPreparation_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Function IDesignerPreparation_GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    IDesignerPreparation_GetFlag = GetFlag(nameId, defaultValue)
End Function

Private Property Get IDesignerPreparation_HiddenStore() As IHiddenNames
    Set IDesignerPreparation_HiddenStore = HiddenStore
End Property

Private Property Get IDesignerPreparation_HostWorkbook() As Workbook
    Set IDesignerPreparation_HostWorkbook = HostWorkbook
End Property

Private Sub IDesignerPreparation_Prepare(ByRef devManager As IDevelopment)
    Prepare devManager
End Sub

Private Sub IDesignerPreparation_SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
    SetFlag nameId, enabled
End Sub
