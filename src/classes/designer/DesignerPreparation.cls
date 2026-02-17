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
'@depends ProjectError, HiddenNames, IHiddenNames, IDevelopment, OSFiles, IOSFiles, CustomTable, ICustomTable, BetterArray, DropdownLists, IDropdownLists
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements IDesignerPreparation

'@class DesignerPreparation
'@description
'Prepares a designer workbook for ribbon interactions. The full
'preparation sequence imports translation tables from a user-selected
'workbook, creates a dropdown worksheet with standard dropdown lists
'(setup languages, interface languages, epiweek start, design values),
'hides internal worksheets (VeryHidden for __pass/__formatter/__formula/
'__dropdowns, Hidden for translation sheets), seeds workbook- and
'worksheet-level HiddenNames flags, applies data validation to the
'T_Multi table on the GenerateMultiple worksheet, and registers hidden
'sheets with the development manager. Also provides methods to read and
'write ribbon checkbox flags persisted through workbook-level hidden names.
'@depends HiddenNames, IHiddenNames, IDevelopment, OSFiles, IOSFiles, CustomTable, ICustomTable, BetterArray, DropdownLists, IDropdownLists


'@section State
'===============================================================================

Private Type TDesignerPreparation
    hostBook As Workbook       'Designer workbook being prepared
    store As IHiddenNames      'Workbook-level hidden names manager
    dropdowns As IDropdownLists 'Dropdown manager bound to dropdown sheet
    dropSheet As Worksheet     'Worksheet hosting dropdown listobjects
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
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const SHEET_GENERATE_MULTIPLE As String = "GenerateMultiple"

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
Private Const GEO_FORM_LOADED As String = "RNG_FormLoaded"

'Dropdown names
Private Const DROP_SETUP_LANGUAGES As String = "__setup_languages"
Private Const DROP_INTERFACE_LANGUAGES As String = "__interface_languages"
Private Const DROP_EPIWEEK_START As String = "__epiweek_start"
Private Const DROP_DESIGN_VALUES As String = "__design_values"
Private Const DROPDOWN_PREFIX As String = "dropdown_"

'T_Multi table
Private Const TABLE_MULTI As String = "T_Multi"

'Main worksheet named ranges for dropdown validation
Private Const SHEET_MAIN As String = "Main"
Private Const RNG_LANG_SETUP As String = "RNG_LangSetup"
Private Const RNG_DESIGN_LL As String = "RNG_DesignLL"
Private Const RNG_LL_FORM As String = "RNG_LLForm"


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
'cached hidden names store and dropdown manager so they will be rebuilt
'on next access. Raises an error when the workbook is Nothing.
'@param hostBook Workbook. The designer workbook to bind to.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook is required."
    End If

    Set this.hostBook = hostBook
    Set this.store = Nothing
    Set this.dropdowns = Nothing
    Set this.dropSheet = Nothing
End Sub

'@label:prepare
'@sub-title Execute all preparation tasks on the designer workbook
'@details
'Runs the full preparation sequence:
'1. Prompts the user to select a translations workbook and imports
'   translation tables into LinelistTranslation and DesignerTranslation.
'2. Creates the dropdown worksheet and registers all standard dropdown
'   lists (setup languages, interface languages, epiweek start, design).
'3. Hides internal worksheets (VeryHidden for __pass/__formatter/
'   __formula/__dropdowns, Hidden for LinelistTranslation,
'   DesignerTranslation).
'4. Seeds workbook-level HiddenNames flags (ribbon checks, language
'   codes, last opened date).
'5. Seeds Geo worksheet-level HiddenNames flags (geo language, name,
'   update status).
'6. Applies data validation to the T_Multi table on GenerateMultiple.
'When devManager is provided, hidden sheets are also registered with
'the manager for deployment.
'@param devManager IDevelopment. ByRef development manager instance, or Nothing.
'@export
Public Sub Prepare(ByRef devManager As IDevelopment)
    ImportTranslations
    EnsureDropdowns
    HideWorksheets devManager
    EnsureHiddenStore
    EnsureWorkbookFlags
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

'@label:dropdowns
'@prop-title The dropdown manager created during preparation
'@details
'Returns the IDropdownLists manager bound to the designer dropdown
'worksheet. Lazily initialises the manager on first access if it was
'not already created by EnsureDropdowns or Prepare.
'@return IDropdownLists. The dropdown manager instance.
'@export
Public Property Get Dropdowns() As IDropdownLists
    EnsureDropdownManager
    Set Dropdowns = this.dropdowns
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
'Sets __pass, __formatter, __formula, __dropdowns to xlSheetVeryHidden.
'Sets LinelistTranslation and DesignerTranslation to xlSheetHidden
'(visible in the VBA editor for developers). When a development manager
'is provided, all sheets are also registered via AddHiddenSheet for
'deployment. Silently skips any worksheet that does not exist.
'@param devManager IDevelopment. Development manager for deployment registration, or Nothing.
Private Sub HideWorksheets(ByRef devManager As IDevelopment)
    SetSheetVisibility SHEET_PASS, xlSheetVeryHidden
    SetSheetVisibility SHEET_FORMATTER, xlSheetVeryHidden
    SetSheetVisibility SHEET_FORMULA, xlSheetVeryHidden
    SetSheetVisibility SHEET_DROPDOWNS, xlSheetVeryHidden
    SetSheetVisibility SHEET_LL_TRANSLATION, xlSheetHidden
    SetSheetVisibility SHEET_DES_TRANSLATION, xlSheetHidden

    If Not devManager Is Nothing Then
        RegisterHiddenSheet devManager, SHEET_PASS
        RegisterHiddenSheet devManager, SHEET_FORMATTER
        RegisterHiddenSheet devManager, SHEET_FORMULA
        RegisterHiddenSheet devManager, SHEET_DROPDOWNS
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
    this.store.SetValue TAG_LAST_OPENED, Format$(Now, "yyyy-mm-dd")
End Sub

'@label:ensure-geo-flags
'@sub-title Ensure Geo worksheet-level HiddenNames flags exist
'@details
'Creates worksheet-level HiddenNames on the Geo worksheet for language
'code, geobase name, metadata language, update status, and form loaded
'state. Also creates RNG_PastingGeoCol as a real cell-based named range
'(not a HiddenName) because it is used as a paste anchor at runtime.
'Silently exits when the Geo worksheet does not exist in the host
'workbook.
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
    geoStore.EnsureName GEO_FORM_LOADED, vbNullString, HiddenNameTypeString

    'RNG_PastingGeoCol is a real cell-based named range, not a HiddenName
    EnsureGeoRange geoSheet, "RNG_PastingGeoCol"
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
    RegisterDropdown DROP_INTERFACE_LANGUAGES, _
        Array("ARA-" & ChrW$(1593) & ChrW$(1585) & ChrW$(1576) & ChrW$(1609), _
              "ENG-English", _
              "FRA-Fran" & ChrW$(231) & "ais", _
              "POR-Portugu" & ChrW$(233) & "s", _
              "ESP-Espa" & ChrW$(241) & "ol")
    RegisterDropdown DROP_EPIWEEK_START, Array("1", "2", "3", "4", "5", "6", "7")
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
    Dim table As ICustomTable
    Dim drop As IDropdownLists

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
    Dim drop As IDropdownLists

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
'@param drop IDropdownLists. The dropdown manager with registered lists.
'@param dropName String. The dropdown list name to use as validation source.
Private Sub ApplyRangeValidation(ByVal sh As Worksheet, ByVal rngName As String, _
                                 ByVal drop As IDropdownLists, ByVal dropName As String)
    Dim rng As Range

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    drop.SetValidation cellRng:=rng, listName:=dropName, alertType:="info"
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


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IDesignerPreparation contract.
'See the corresponding members above for full documentation.

Private Sub IDesignerPreparation_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Property Get IDesignerPreparation_Dropdowns() As IDropdownLists
    Set IDesignerPreparation_Dropdowns = Dropdowns
End Property

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
