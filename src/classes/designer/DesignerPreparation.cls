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
'@depends ProjectError, HiddenNames, Development, OSFiles, CustomTable, BetterArray, DropdownLists, Messenger
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit

'@class DesignerPreparation
'@description
'Prepares a designer workbook for ribbon interactions. The full
'preparation sequence imports translation tables from a user-selected
'workbook, creates a dropdown worksheet with standard dropdown lists
'(setup languages, interface languages, epiweek start, design values),
'hides internal worksheets (VeryHidden for __pass/__formula/__dropdowns,
'Hidden for translation sheets), seeds workbook- and
'worksheet-level HiddenNames flags, carries the designer version from the
'Dev worksheet into a workbook-level name, applies data validation to the
'T_Multi table on the GenerateMultiple worksheet, and registers hidden
'sheets with the development manager. Also provides methods to read and
'write ribbon checkbox flags persisted through workbook-level hidden names.
'@depends HiddenNames, Development, OSFiles, CustomTable, BetterArray, DropdownLists, Messenger


'@section State
'===============================================================================

Private Type TDesignerPreparation
    hostBook As Workbook       'Designer workbook being prepared
    store As HiddenNames      'Workbook-level hidden names manager
    dropdowns As DropdownLists 'Dropdown manager bound to dropdown sheet
    dropSheet As Worksheet     'Worksheet hosting dropdown listobjects
    sealed As Boolean          'True once the factory has finished setup
End Type

Private this As TDesignerPreparation

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "DesignerPreparation"
Private Const FLAG_ALERT As String = "chkAlert"
Private Const FLAG_INSTRUCT As String = "chkInstruct"
'The build path switch. Off by default: the build runs in a hidden
'instance on Windows, and Mac ignores the flag.
Private Const FLAG_BUILD_IN_PLACE As String = "chkBuildInPlace"
'The silence switch has no constant here. Its name and its two values live on
'Messenger, which is the one place that string is written on the VBA side, and
'the setup and the generated linelist seed the same flag from the same three
'properties. See EnsureDefaultFlags.
Private Const FLAG_ON As String = "Yes"
Private Const FLAG_OFF As String = "No"

'Sheet names
Private Const SHEET_PASS As String = "__pass"
Private Const SHEET_FORMATTER As String = "__formatter"
Private Const SHEET_FORMULA As String = "__formula"
Private Const SHEET_CHECKING As String = "__check"
Private Const SHEET_LL_TRANSLATION As String = "LinelistTranslation"
Private Const SHEET_DES_TRANSLATION As String = "DesignerTranslation"
Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const SHEET_GENERATE_MULTIPLE As String = "GenerateMultiple"
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_DEV As String = "Dev"

'Main worksheet named ranges for dropdown validation
Private Const RNG_LANG_SETUP As String = "RNG_LangSetup"
Private Const RNG_LL_FORM As String = "RNG_LLForm"
Private Const RNG_DESIGN_LL As String = "RNG_DesignLL"

'Workbook-level HiddenNames
'TAG_DES_LANG has one scope: DesignerTranslation writes it on the worksheet.
Private Const TAG_LL_LANG_CODE As String = "RNG_LLLanguageCode"
Private Const TAG_DICT_LANG As String = "RNG_DictionaryLanguage"
Private Const TAG_LAST_OPENED As String = "RNG_LastOpenedDate"

'The formatter flag has one home, here. The styles import sets it, loading a
'setup file clears it, and the transfer reads it to pick the formatter source.
Private Const TAG_FORMATTER_IMPORTED As String = "TAG_FORMATTER_IMPORTED"

'The ribbon template file the multi generation gives to every row it builds.
'The single generation reads the path the user loaded on Main; the rows of
'GenerateMultiple read this one, so the two screens stop sharing a cell.
Private Const TAG_RIBBON_TEMPLATE As String = "TAG_RIBBON_TEMPLATE"

'The version of the designer. The owner writes it by hand on the Dev worksheet,
'as the value name RNG_Version, and preparation copies it into the workbook-
'scoped name below. LLGeo reads that one when it writes the metadata sheet of a
'generated linelist.
Private Const RNG_VERSION As String = "RNG_Version"
Private Const TAG_DESIGNER_VERSION As String = "RNG_DesignerVersion"

'Geo worksheet-level HiddenNames
Private Const GEO_LANG_CODE As String = "RNG_GeoLangCode"
Private Const GEO_NAME As String = "RNG_GeoName"
Private Const GEO_META_LANG As String = "RNG_MetaLang"
Private Const GEO_UPDATED As String = "RNG_GeoUpdated"
Private Const GEO_TRANSLATED As String = "RNG_GeoTranslated"
Private Const GEO_FORM_LOADED As String = "RNG_FormLoaded"

'Geo workbook-level HiddenNames holding the five level labels
Private Const GEO_ADM1_NAME As String = "RNG_ADM1NAME"
Private Const GEO_ADM2_NAME As String = "RNG_ADM2NAME"
Private Const GEO_ADM3_NAME As String = "RNG_ADM3NAME"
Private Const GEO_ADM4_NAME As String = "RNG_ADM4NAME"
Private Const GEO_HF_NAME As String = "RNG_HFNAME"

'Dropdown names
Private Const DROP_SETUP_LANGUAGES As String = "__setup_languages"
Private Const DROP_INTERFACE_LANGUAGES As String = "__interface_languages"
Private Const DROP_EPIWEEK_START As String = "__epiweek_start"
Private Const DROP_DESIGN_VALUES As String = "__design_values"
Private Const DROPDOWN_PREFIX As String = "dropdown_"

'T_Multi table
Private Const TABLE_MULTI As String = "T_Multi"


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create and configure a designer preparation helper
'@details
'Factory method that creates a new DesignerPreparation instance, binds
'it to the supplied designer workbook via Configure, and returns the
'result. Called on the predeclared instance.
'@param hostBook Workbook. The designer workbook to prepare.
'@return DesignerPreparation. A fully initialised instance ready for use.
'@export
Public Function Create(ByVal hostBook As Workbook) As DesignerPreparation
    Dim instance As DesignerPreparation

    Set instance = New DesignerPreparation
    instance.Configure hostBook
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
'@sub-title Bind the helper to the designer workbook
'@details
'Stores a reference to the host workbook. Called once by the factory,
'which then seals the instance; a later call raises. Raises an error
'when the workbook is Nothing.
'@param hostBook Workbook. The designer workbook to bind to.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    GuardNotSealed "Configure"

    If hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook is required."
    End If

    Set this.hostBook = hostBook
End Sub

'@label:prepare
'@sub-title Execute all preparation tasks on the designer workbook
'@details
'Runs the full preparation sequence:
'1. Prompts the user to select a translations workbook and imports
'   translation tables into LinelistTranslation and DesignerTranslation.
'2. Creates the dropdown worksheet and registers all standard dropdown
'   lists (setup languages, interface languages, epiweek start, design).
'3. Hides internal worksheets (VeryHidden for __pass/__formula/
'   __dropdowns, Hidden for LinelistTranslation, DesignerTranslation).
'   __formatter and __check are left alone: the designer works with both.
'4. Seeds workbook-level HiddenNames flags (ribbon checks, language
'   codes, last opened date).
'5. Copies the version from the Dev worksheet into the workbook-level
'   RNG_DesignerVersion name.
'6. Seeds Geo worksheet-level HiddenNames flags (geo language, name,
'   update status).
'7. Applies data validation to the T_Multi table on GenerateMultiple.
'8. Applies dropdown validation to the Main named ranges (RNG_LangSetup,
'   RNG_DesignLL, RNG_LLForm).
'When devManager is provided, hidden sheets are also registered with
'the manager for deployment.
'When translationsSource is provided, the translation tables are read
'from that open workbook. When it is Nothing, the user is prompted to
'select the translations file. Tests hand a fixture workbook here so
'the sequence runs without a dialog.
'@param devManager Development. ByRef development manager instance, or Nothing.
'@param translationsSource Optional Workbook. Open workbook holding the translation tables. Defaults to Nothing.
'@export
Public Sub Prepare(ByRef devManager As Development, _
                   Optional ByVal translationsSource As Workbook = Nothing)
    ImportTranslations translationsSource
    EnsureDropdowns
    HideWorksheets devManager
    EnsureHiddenStore
    EnsureWorkbookFlags
    EnsureVersionName
    EnsureGeoFlags
    ApplyMultiValidations
    ApplyMainValidations
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

'@label:formatter-imported
'@prop-title Whether the designer holds a formatter loaded through the ribbon
'@details
'Reads the workbook-level TAG_FORMATTER_IMPORTED flag. True means the
'styles import button was pressed for the setup file that is loaded now,
'so the generation ships the designer's __formatter sheet. False means
'the setup file owns the formatter.
'@return Boolean. True when the flag reads "Yes".
'@export
Public Property Get FormatterImported() As Boolean
    Dim rawValue As String

    EnsureHiddenStore
    rawValue = this.store.ValueAsString(TAG_FORMATTER_IMPORTED, FLAG_OFF)
    FormatterImported = (StrComp(rawValue, FLAG_ON, vbTextCompare) = 0)
End Property

'@label:formatter-imported-set
'@prop-title Set the formatter flag
'@details
'Writes "Yes" or "No" into the workbook-level TAG_FORMATTER_IMPORTED flag,
'creating the hidden name when the designer has none yet. The styles
'import writes True; loading a setup file writes False, so the designer's
'formatter wins only after an explicit styles import for the setup that
'is loaded now.
'@param imported Boolean. True to mark the designer's formatter as the live copy.
'@export
Public Property Let FormatterImported(ByVal imported As Boolean)
    Dim storedValue As String

    storedValue = IIf(imported, FLAG_ON, FLAG_OFF)
    EnsureHiddenStore
    this.store.EnsureName TAG_FORMATTER_IMPORTED, storedValue, HiddenNameTypeString
    this.store.SetValue TAG_FORMATTER_IMPORTED, storedValue
End Property

'@label:ribbon-template-path
'@prop-title The ribbon template file of the multi generation
'@details
'Reads the workbook-level TAG_RIBBON_TEMPLATE name. The multi generation
'gives this file to every row it builds, so the rows carry a ribbon
'without anybody loading a template on the Main worksheet first.
'
'An empty answer means the designer has never been given one. It also
'means the linelists carry buttons: that is what an empty template path
'has always meant to the generation.
'@return String. The stored path, empty when the designer holds none.
'@export
Public Property Get RibbonTemplatePath() As String
    EnsureHiddenStore
    RibbonTemplatePath = Trim$(this.store.ValueAsString(TAG_RIBBON_TEMPLATE, vbNullString))
End Property

'@label:ribbon-template-path-set
'@prop-title Store the ribbon template file of the multi generation
'@details
'Creates the hidden name when the designer has none yet. An empty path
'takes the name off the store, which puts the designer back in the state
'it had before it was ever given a file: the read answers empty and the
'generation builds the buttons.
'@param templatePath String. The full path of the template file.
'@export
Public Property Let RibbonTemplatePath(ByVal templatePath As String)
    EnsureHiddenStore

    If LenB(templatePath) = 0 Then
        If this.store.HasName(TAG_RIBBON_TEMPLATE) Then _
            this.store.RemoveName TAG_RIBBON_TEMPLATE
        Exit Property
    End If

    this.store.EnsureName TAG_RIBBON_TEMPLATE, templatePath, HiddenNameTypeString
    this.store.SetValue TAG_RIBBON_TEMPLATE, templatePath
End Property

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
'Returns the HiddenNames instance lazily created from the host workbook.
'Used internally to persist and retrieve ribbon flag values.
'@return HiddenNames. The hidden names manager.
'@export
Public Property Get HiddenStore() As HiddenNames
    EnsureHiddenStore
    Set HiddenStore = this.store
End Property

'@label:dropdowns
'@prop-title The dropdown manager created during preparation
'@details
'Returns the DropdownLists manager bound to the designer dropdown
'worksheet. Lazily initialises the manager on first access if it was
'not already created by EnsureDropdowns or Prepare.
'@return DropdownLists. The dropdown manager instance.
'@export
Public Property Get Dropdowns() As DropdownLists
    EnsureDropdownManager
    Set Dropdowns = this.dropdowns
End Property


'@section Preparation Helpers
'===============================================================================

'@label:import-translations
'@sub-title Import translation tables from a translations workbook
'@details
'Iterates through LinelistTranslation and DesignerTranslation sheets and
'imports matching ListObjects via CustomTable.Import. When source is
'Nothing, prompts the user to select an xlsx file and opens it read-only;
'the workbook opened here is also closed here. When source is provided,
'reads from that open workbook and leaves it open -- the clickImpTrans
'ribbon callback opens the file itself and hands it over that way, and
'tests hand a fixture workbook so the sequence runs without a dialog.
'
'This body is the only translations import in the project. The
'preparation sequence and the ribbon button both come here, so the eight
'table names are written down once.
'
'A sheet or table missing on either side is skipped through the resolve
'helpers, which probe and clear their own errors. The cleanup label only
'sees errors raised while the import itself is running, so a designer
'that lacks one sheet prepares without a raise.
'@param source Optional Workbook. Open workbook holding the translation tables, or Nothing.
'@export
Public Sub ImportTranslations(Optional ByVal source As Workbook = Nothing)
    Dim io As OSFiles
    Dim importBook As Workbook
    Dim openedHere As Boolean
    Dim wb As Workbook
    Dim sheetNames As BetterArray
    Dim tableNames As BetterArray
    Dim idx As Long
    Dim sheetName As String
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim lo As ListObject
    Dim targetTable As CustomTable
    Dim sourceTable As CustomTable

    Set wb = HostWorkbook

    If source Is Nothing Then
        Set io = OSFiles.Create()
        io.LoadFile "*.xlsx"
        If Not io.HasValidFile() Then Exit Sub

        On Error GoTo ImportCleanup
        Set importBook = wb.Application.Workbooks.Open(io.File(), ReadOnly:=True)
        openedHere = True
    Else
        Set importBook = source
        On Error GoTo ImportCleanup
    End If

    Set sheetNames = New BetterArray
    sheetNames.Push SHEET_LL_TRANSLATION, SHEET_DES_TRANSLATION

    'T_tradDrop is in the list. DesignerTranslation refuses to build without
    'that table, and neither importer carried it, so the dropdown labels of
    'the designer could never be refreshed from a translations workbook.
    Set tableNames = New BetterArray
    tableNames.Push "t_tradllshapes", "t_tradllmsg", "t_tradllforms", "t_tradllribbon", _
                    "t_tradmsg", "t_tradrange", "t_tradshape", "t_traddrop"

    For idx = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = CStr(sheetNames.Item(idx))

        Set targetSheet = ResolveSheetIn(wb, sheetName)
        Set sourceSheet = ResolveSheetIn(importBook, sheetName)

        If (Not targetSheet Is Nothing) And (Not sourceSheet Is Nothing) Then
            For Each lo In targetSheet.ListObjects
                If tableNames.Includes(LCase$(lo.Name)) Then
                    Set targetTable = CustomTable.Create(lo)
                    Set sourceTable = ResolveTableIn(sourceSheet, lo.Name)

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
    If openedHere And (Not importBook Is Nothing) Then importBook.Close saveChanges:=False
    If Err.Number <> 0 Then
        RaiseError ProjectError.SomethingWentWrong, _
                   "Unable to import translations: " & Err.Description
    End If
End Sub

'@label:ensure-dropdowns
'@sub-title Create the dropdown worksheet and register all standard dropdown lists
'@details
'Creates the dropdown manager if needed, then registers setup languages
'(empty placeholder), interface languages, epiweek start (1-7), and
'design values dropdown lists.
Private Sub EnsureDropdowns()
    EnsureDropdownManager
    RegisterAllDropdowns
End Sub

'@label:hide-worksheets
'@sub-title Hide internal worksheets and register them with the dev manager
'@details
'Sets __pass, __formula, __dropdowns and Geo to xlSheetVeryHidden. The Geo
'sheet of the designer is an empty template: the geobase is imported on the
'linelist, never here, so nobody edits it by hand. __formatter and __check
'are deliberately left as they are: __formatter carries the formats a user
'picks and __check carries the generation report, so a prepared designer
'keeps both on screen. Sets LinelistTranslation and DesignerTranslation to
'xlSheetHidden (visible in the VBA editor for developers). When a
'development manager is provided, all sheets are also registered via
'AddHiddenSheet for deployment. Silently skips any worksheet that does
'not exist.
'@param devManager Development. Development manager for deployment registration, or Nothing.
Private Sub HideWorksheets(ByRef devManager As Development)
    SetSheetVisibility SHEET_PASS, xlSheetVeryHidden
    SetSheetVisibility SHEET_FORMULA, xlSheetVeryHidden
    SetSheetVisibility SHEET_DROPDOWNS, xlSheetVeryHidden
    SetSheetVisibility SHEET_GEO, xlSheetVeryHidden
    SetSheetVisibility SHEET_LL_TRANSLATION, xlSheetHidden
    SetSheetVisibility SHEET_DES_TRANSLATION, xlSheetHidden

    If Not devManager Is Nothing Then
        RegisterHiddenSheet devManager, SHEET_PASS
        RegisterHiddenSheet devManager, SHEET_FORMULA
        RegisterHiddenSheet devManager, SHEET_DROPDOWNS
        RegisterHiddenSheet devManager, SHEET_GEO
        RegisterHiddenSheet devManager, SHEET_LL_TRANSLATION
        RegisterHiddenSheet devManager, SHEET_DES_TRANSLATION
        RegisterHiddenSheet devManager, SHEET_DEV
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
    this.store.EnsureName TAG_LL_LANG_CODE, vbNullString, HiddenNameTypeString
    this.store.EnsureName TAG_DICT_LANG, vbNullString, HiddenNameTypeString
    this.store.EnsureName TAG_LAST_OPENED, Format$(Now, "yyyy-mm-dd"), HiddenNameTypeString
    this.store.SetValue TAG_LAST_OPENED, Format$(Now, "yyyy-mm-dd")
End Sub

'@label:ensure-version-name
'@sub-title Copy the designer version into a workbook-level hidden name
'@details
'The version is owner hand work: it is written once on the Dev worksheet, as
'the value name RNG_Version scoped to that sheet. This copies it into the
'workbook-scoped RNG_DesignerVersion name, which is where LLGeo reads it for
'the metadata sheet of every generated linelist.
'
'The copy is what the metadata sees, so a version edited on the Dev sheet
'reaches a linelist once the designer is initialised again.
'
'Three states leave the workbook name as it is: no Dev worksheet, no
'RNG_Version on it, and an empty RNG_Version. A designer built before the name
'existed is in one of them, and it is an ordinary designer rather than a
'broken one: LLGeo answers "(not found)" and the build carries on.
Private Sub EnsureVersionName()
    Dim devSheet As Worksheet
    Dim versionText As String

    Set devSheet = ResolveWorksheet(SHEET_DEV)
    If devSheet Is Nothing Then Exit Sub

    versionText = Trim$(DevVersionText(devSheet))
    If LenB(versionText) = 0 Then Exit Sub

    EnsureHiddenStore
    this.store.EnsureName TAG_DESIGNER_VERSION, versionText, HiddenNameTypeString
    this.store.SetValue TAG_DESIGNER_VERSION, versionText
End Sub

'@label:dev-version-text
'@sub-title Read the version the Dev worksheet carries
'@details
'RNG_Version is a value name: the version text is in the definition itself, and
'QuickValue reads it without walking the whole Names collection of the sheet.
'
'A Dev worksheet that points the name at a cell instead is read from that cell,
'so a version typed in the grid answers here too. Without that, the name would
'be parsed as text and the metadata sheet would carry an address.
'@param devSheet Worksheet. The Dev worksheet.
'@return String. The version text, or an empty string when the sheet holds none.
Private Function DevVersionText(ByVal devSheet As Worksheet) As String
    Dim rng As Range
    Dim cellValue As Variant

    'A sheet holding no such name raises, and no name is an ordinary answer.
    On Error Resume Next
    Set rng = devSheet.Names(RNG_VERSION).RefersToRange
    On Error GoTo 0

    If rng Is Nothing Then
        DevVersionText = HiddenNames.QuickValue(devSheet, RNG_VERSION)
        Exit Function
    End If

    cellValue = rng.Cells(1, 1).Value
    If Not IsError(cellValue) Then DevVersionText = CStr(cellValue)
End Function

'@label:ensure-geo-flags
'@sub-title Ensure the Geo HiddenNames flags exist
'@details
'Creates worksheet-level HiddenNames on the Geo worksheet for language
'code, geobase name, metadata language, update status, translation state
'and form loaded state.
'
'The five level labels (RNG_ADM1NAME through RNG_ADM4NAME and RNG_HFNAME)
'are created at WORKBOOK scope. LLGeo.CheckRequirements asks for them there,
'and CrossTable resolves them from a worksheet formula on the spatial
'analysis sheet, where a worksheet-scoped name would give #NAME?.
'
'RNG_PastingGeoCol stays a real cell-based named range. It is a paste anchor
'at runtime and the paste needs a cell to land in.
'
'Exits with no work when the host workbook has no Geo worksheet.
'
'PUBLIC BECAUSE A HEADLESS BUILD NEEDS IT ON ITS OWN
'-------------------------------------------------------------------------------
'Prepare is the ordinary caller and it opens a file dialog on its first step, so
'a headless run cannot reach this through it. Every name below is what
'LLGeo.CheckRequirements asks for, and a designer that has never been prepared
'with the current code carries none of them: LLGeo.Create then fails, the
'dictionary logs "geolines not append", and every geo variable stays one column
'instead of the twelve it expands into. Calling this first is what a build does
'to make the geo sheet readable, and it is idempotent -- every write is an
'EnsureName, so a designer that already carries them is left alone.
'@export
Public Sub EnsureGeoFlags()
    Dim geoSheet As Worksheet
    Dim geoStore As HiddenNames
    Dim levelStore As HiddenNames

    Set geoSheet = ResolveWorksheet(SHEET_GEO)
    If geoSheet Is Nothing Then Exit Sub

    Set geoStore = HiddenNames.Create(geoSheet)
    geoStore.EnsureName GEO_LANG_CODE, vbNullString, HiddenNameTypeString
    geoStore.EnsureName GEO_NAME, vbNullString, HiddenNameTypeString
    geoStore.EnsureName GEO_META_LANG, vbNullString, HiddenNameTypeString
    geoStore.EnsureName GEO_UPDATED, "empty", HiddenNameTypeString
    geoStore.EnsureName GEO_TRANSLATED, False, HiddenNameTypeBoolean
    geoStore.EnsureName GEO_FORM_LOADED, vbNullString, HiddenNameTypeString

    Set levelStore = HiddenNames.Create(geoSheet.Parent)
    EnsureLevelName levelStore, geoSheet, GEO_ADM1_NAME
    EnsureLevelName levelStore, geoSheet, GEO_ADM2_NAME
    EnsureLevelName levelStore, geoSheet, GEO_ADM3_NAME
    EnsureLevelName levelStore, geoSheet, GEO_ADM4_NAME
    EnsureLevelName levelStore, geoSheet, GEO_HF_NAME

    EnsureGeoRange geoSheet, "RNG_PastingGeoCol"
End Sub

'@label:ensure-level-name
'@sub-title Ensure one workbook-level hidden name holding a geo level label
'@details
'A Geo worksheet written before the level labels moved into HiddenNames
'carries each label in a cell-based named range of the same identifier. The
'cell is read first and its value seeds the hidden name, so the label the
'designer already shows survives the move.
'@param levelStore HiddenNames. The store over the workbook.
'@param sh Worksheet. The Geo worksheet.
'@param nameId String. The hidden name identifier.
Private Sub EnsureLevelName(ByVal levelStore As HiddenNames, ByVal sh As Worksheet, _
                            ByVal nameId As String)
    Dim rng As Range
    Dim seedValue As String
    Dim cellValue As Variant

    If levelStore.HasName(nameId) Then Exit Sub

    On Error Resume Next
    Set rng = sh.Range(nameId)
    On Error GoTo 0

    If Not rng Is Nothing Then
        cellValue = rng.Cells(1, 1).Value
        If Not IsError(cellValue) Then seedValue = CStr(cellValue)
    End If

    levelStore.EnsureName nameId, seedValue, HiddenNameTypeString
End Sub


'@section Dropdown Helpers
'===============================================================================

'@label:ensure-dropdown-manager
'@sub-title Ensure the dropdown manager is initialised for the target workbook
'@details
'Exits immediately when the manager is already initialised. Otherwise
'resolves or creates the dropdown worksheet, then creates a DropdownLists
'instance bound to it with the standard dropdown prefix.
Private Sub EnsureDropdownManager()
    If Not this.dropdowns Is Nothing Then Exit Sub

    Dim targetSheet As Worksheet

    Set targetSheet = EnsureDropdownSheet
    Set this.dropdowns = DropdownLists.Create(targetSheet, DROPDOWN_PREFIX)
End Sub

'@label:ensure-dropdown-sheet
'@sub-title Ensure the dropdown worksheet exists and cache the reference
'@details
'Resolves the dropdown worksheet from the host workbook. If not found,
'creates a new worksheet at the end and names it. Caches the reference
'for subsequent calls.
'@return Worksheet. The dropdown worksheet.
Private Function EnsureDropdownSheet() As Worksheet
    If Not this.dropSheet Is Nothing Then
        Set EnsureDropdownSheet = this.dropSheet
        Exit Function
    End If

    Dim sh As Worksheet
    Dim wb As Workbook

    Set wb = HostWorkbook

    On Error Resume Next
    Set sh = wb.Worksheets(SHEET_DROPDOWNS)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = SHEET_DROPDOWNS
    End If

    Set this.dropSheet = sh
    Set EnsureDropdownSheet = sh
End Function

'@label:register-all-dropdowns
'@sub-title Register every standard dropdown list for designer worksheets
'@details
'Calls RegisterDropdown for each predefined dropdown: setup languages
'(empty placeholder for later population), interface languages
'(pre-populated with the five supported languages), epiweek start
'(numbers 1 through 7), and design values (matching LLFormat design
'constants).
Private Sub RegisterAllDropdowns()
    RegisterDropdown DROP_SETUP_LANGUAGES, Array(vbNullString, vbNullString)
    'The prefix of each entry is not decoration. InitTransfer splits the picked
    'value on the dash and writes the prefix as RNG_LLLanguageCode, and that is
    'the COLUMN NAME the four LinelistTranslation tables are keyed on: MSG_ID,
    'ENG, FRA, SPA, POR, ARA. Spanish read "ESP-" here and the column is SPA, so
    'picking Spanish resolved no column and every message in the generated
    'linelist came out as its own tag, with nothing said anywhere.
    RegisterDropdown DROP_INTERFACE_LANGUAGES, _
        Array("ARA-" & ChrW$(1593) & ChrW$(1585) & ChrW$(1576) & ChrW$(1609), _
              "ENG-English", _
              "FRA-Fran" & ChrW$(231) & "ais", _
              "POR-Portugu" & ChrW$(233) & "s", _
              "SPA-Espa" & ChrW$(241) & "ol")
    RegisterDropdown DROP_EPIWEEK_START, Array("1", "2", "3", "4", "5", "6", "0")
    RegisterDropdown DROP_DESIGN_VALUES, Array("design 1", "design 2", "user defined")
End Sub

'@label:register-dropdown
'@sub-title Register a single dropdown by converting the values into a BetterArray
'@details
'Ensures the dropdown manager is initialised, converts the supplied Variant
'array into a BetterArray (1-based), and adds it to the dropdown manager
'under the given name.
'@param dropdownName String. The name to register the dropdown under.
'@param values Variant. An array of values to populate the dropdown.
Private Sub RegisterDropdown(ByVal dropdownName As String, ByVal values As Variant)
    Dim listValues As BetterArray
    Dim idx As Long

    EnsureDropdownManager

    Set listValues = New BetterArray
    listValues.LowerBound = 1

    If IsArray(values) Then
        For idx = LBound(values) To UBound(values)
            listValues.Push values(idx)
        Next idx
    Else
        listValues.Push values
    End If

    this.dropdowns.Add listValues, dropdownName
End Sub


'@section Multi Validations
'===============================================================================

'@label:apply-multi-validations
'@sub-title Apply data validation to the T_Multi table on GenerateMultiple
'@details
'Resolves the GenerateMultiple worksheet and the T_Multi ListObject,
'then wires the "language of the interface", "epiweek start", and
'"design" columns to their corresponding dropdown lists via
'CustomTable.SetValidation. Silently exits when the worksheet or
'table does not exist in the host workbook.
Private Sub ApplyMultiValidations()
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim table As CustomTable
    Dim drop As DropdownLists

    Set sh = ResolveWorksheet(SHEET_GENERATE_MULTIPLE)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set lo = sh.ListObjects(TABLE_MULTI)
    On Error GoTo 0

    If lo Is Nothing Then Exit Sub

    Set table = CustomTable.Create(lo)
    EnsureDropdownManager
    Set drop = this.dropdowns

    table.SetValidation colName:="language of the interface", _
                        drop:=drop, _
                        dropName:=DROP_INTERFACE_LANGUAGES, _
                        alertType:="info"
    table.SetValidation colName:="epiweek start", _
                        drop:=drop, _
                        dropName:=DROP_EPIWEEK_START, _
                        alertType:="info"
    table.SetValidation colName:="design", _
                        drop:=drop, _
                        dropName:=DROP_DESIGN_VALUES, _
                        alertType:="info"
End Sub

'@section Main Validations
'===============================================================================

'@label:apply-main-validations
'@sub-title Apply dropdown validation to named ranges on the Main worksheet
'@details
'Resolves the Main worksheet and applies data validation from the
'registered dropdown lists to three named ranges: RNG_LangSetup
'(setup language), RNG_DesignLL (design selection), and RNG_LLForm
'(language of linelist interface). Silently exits when the Main
'worksheet does not exist or individual named ranges are missing.
Private Sub ApplyMainValidations()
    Dim sh As Worksheet
    Dim drop As DropdownLists

    Set sh = ResolveWorksheet(SHEET_MAIN)
    If sh Is Nothing Then Exit Sub

    EnsureDropdownManager
    Set drop = this.dropdowns

    ApplyRangeValidation sh, RNG_LANG_SETUP, drop, DROP_SETUP_LANGUAGES
    ApplyRangeValidation sh, RNG_DESIGN_LL, drop, DROP_DESIGN_VALUES
    ApplyRangeValidation sh, RNG_LL_FORM, drop, DROP_INTERFACE_LANGUAGES
End Sub

'@label:apply-range-validation
'@sub-title Apply dropdown validation to a single named range
'@details
'Resolves the named range from the worksheet and applies list
'validation using the specified dropdown. Silently exits when the
'named range does not exist on the worksheet.
'@param sh Worksheet. The worksheet containing the named range.
'@param rngName String. The named range identifier.
'@param drop DropdownLists. The dropdown manager with registered lists.
'@param dropName String. The dropdown list name to use as validation source.
Private Sub ApplyRangeValidation(ByVal sh As Worksheet, ByVal rngName As String, _
                                 ByVal drop As DropdownLists, ByVal dropName As String)
    Dim rng As Range

    Set rng = ResolveNamedRange(sh, rngName)
    If rng Is Nothing Then Exit Sub

    drop.SetValidation cellRng:=rng, listName:=dropName, alertType:="info"
End Sub


'@section Internal Helpers
'===============================================================================

'@label:ensure-hidden-store
'@sub-title Lazily initialise the hidden names store from the host workbook
'@details
'Creates the HiddenNames instance on first access by calling
'HiddenNames.Create with the host workbook. Subsequent calls are no-ops.
Private Sub EnsureHiddenStore()
    If this.store Is Nothing Then
        Set this.store = HiddenNames.Create(HostWorkbook)
    End If
End Sub

'@label:ensure-default-flags
'@sub-title Ensure the default workbook flags exist in the hidden store
'@details
'Calls EnsureName for chkAlert and chkInstruct with "Yes" as default and
'for chkBuildInPlace with "No", creating the hidden names if they do not
'already exist. SetFlag updates a name and raises on a missing one, so a
'designer whose names predate a flag gets it here, on the first read or
'write of any flag.
'
'The fourth flag is the silence switch, __OBT__SILENT_OPERATIONS__, seeded
'"No". "No" means the designer shows every message box it has always shown, so
'a person pressing buttons sees no change at all. The name and both values are
'read off Messenger rather than written again here. A designer that predates
'the switch holds no such name, and Messenger.ReadStoredSwitch answers False
'for a name that is not there, so it behaves as it did until this runs.
Private Sub EnsureDefaultFlags()
    this.store.EnsureName FLAG_ALERT, FLAG_ON, HiddenNameTypeString
    this.store.EnsureName FLAG_INSTRUCT, FLAG_ON, HiddenNameTypeString
    this.store.EnsureName FLAG_BUILD_IN_PLACE, FLAG_OFF, HiddenNameTypeString
    this.store.EnsureName Messenger.SwitchName, Messenger.SwitchOffValue, HiddenNameTypeString
End Sub

'@label:resolve-worksheet
'@sub-title Resolve a worksheet by name from the host workbook
'@details
'Returns the worksheet with the given name from the host workbook, or
'Nothing when the sheet does not exist.
'@param sheetName String. The worksheet name to resolve.
'@return Worksheet. The resolved worksheet, or Nothing.
Private Function ResolveWorksheet(ByVal sheetName As String) As Worksheet
    Set ResolveWorksheet = ResolveSheetIn(HostWorkbook, sheetName)
End Function

'@label:resolve-sheet-in
'@sub-title Resolve a worksheet by name from a given workbook
'@details
'Returns the worksheet with the given name from the supplied workbook,
'or Nothing when the sheet does not exist. Probes and clears its own
'error, so the caller's error state stays clean.
'@param book Workbook. The workbook to resolve the sheet from.
'@param sheetName String. The worksheet name to resolve.
'@return Worksheet. The resolved worksheet, or Nothing.
Private Function ResolveSheetIn(ByVal book As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set ResolveSheetIn = book.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:resolve-table-in
'@sub-title Resolve a ListObject by name into a CustomTable
'@details
'Returns a CustomTable over the named ListObject on the supplied
'worksheet, or Nothing when the table does not exist or cannot be
'wrapped. Probes and clears its own error, so the caller's error state
'stays clean.
'@param sh Worksheet. The worksheet holding the ListObject.
'@param tableName String. The ListObject name to resolve.
'@return CustomTable. The wrapped table, or Nothing.
Private Function ResolveTableIn(ByVal sh As Worksheet, ByVal tableName As String) As CustomTable
    Dim lo As ListObject

    On Error Resume Next
    Set lo = sh.ListObjects(tableName)
    If Not lo Is Nothing Then Set ResolveTableIn = CustomTable.Create(lo)
    On Error GoTo 0
End Function

'@label:resolve-named-range
'@sub-title Resolve a named range from a worksheet
'@details
'Returns the Range matching the given name on the supplied worksheet,
'or Nothing when the named range does not exist.
'@param sh Worksheet. The worksheet to resolve the range from.
'@param rngName String. The named range identifier.
'@return Range. The resolved range, or Nothing.
Private Function ResolveNamedRange(ByVal sh As Worksheet, ByVal rngName As String) As Range
    On Error Resume Next
    Set ResolveNamedRange = sh.Range(rngName)
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
'@param manager Development. The development manager to register with.
'@param sheetName String. The worksheet name to register.
Private Sub RegisterHiddenSheet(ByRef manager As Development, ByVal sheetName As String)
    On Error Resume Next
    manager.AddHiddenSheet sheetName
    On Error GoTo 0
End Sub

'@label:ensure-geo-range
'@sub-title Ensure a cell-based named range exists on the Geo worksheet
'@details
'Checks whether the named range already exists on the worksheet. If not,
'creates the range in a free cell below the used area so it does not
'overlap with existing ListObjects or other content.
'@param sh Worksheet. The Geo worksheet.
'@param rngName String. The named range identifier to create.
Private Sub EnsureGeoRange(ByVal sh As Worksheet, ByVal rngName As String)
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then
        Dim lastRow As Long
        lastRow = sh.UsedRange.Rows(sh.UsedRange.Rows.Count).Row
        sh.Cells(lastRow + 2, 1).Name = rngName
    End If
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
