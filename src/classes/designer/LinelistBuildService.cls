VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistBuildService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Direct setup-to-linelist build service eliminating the designer intermediary"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Exports pass-through setup data directly to the linelist, bypassing the designer intermediary.")
'@depends ProjectError, LLdictionary, ILLdictionary, LLChoices, ILLChoices, LLExport, ILLExport, Analysis, IAnalysis, LLGeo, ILLGeo, IPasswords, Passwords, LLFormat, ILLFormat, LLTranslation, ILLTranslation, HiddenNames, IHiddenNames
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements ILinelistBuildService

'@class LinelistBuildService
'@description
'Orchestrates building a linelist by exporting pass-through components
'(dictionary, choices, exports, analysis, hidden names) directly from a
'setup workbook to the linelist workbook, eliminating the redundant
'designer intermediary copy. Designer-sourced components (geo, passwords,
'formatter when ribbon-imported, language HiddenNames, translations) are
'exported from the designer. After export, new domain manager instances
'targeting the linelist workbook are available via accessor properties.
'Replaces DesignerImportService for the generation flow.


'@section State
'===============================================================================

Private Type TLinelistBuildService
    hostBook As Workbook            'Designer workbook
    exportedDict As ILLdictionary   'Dictionary manager on linelist workbook
    exportedChoices As ILLChoices   'Choices manager on linelist workbook
    exportedAnalysis As IAnalysis   'Analysis manager on linelist workbook
    exportedExport As ILLExport     'Export manager on linelist workbook
    exportedGeo As ILLGeo           'Geo manager on linelist workbook
    exportedPass As IPasswords      'Passwords manager on linelist workbook
    exportedTrans As ILLTranslation 'Translations manager on linelist workbook
End Type

Private this As TLinelistBuildService

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "LinelistBuildService"

'Sheet names in the setup workbook
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_EXPORTS As String = "Exports"
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const SHEET_FORMATTER As String = "__formatter"

'Sheet names in the designer workbook
Private Const SHEET_LL_TRANSLATION As String = "LinelistTranslation"
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_PASS As String = "__pass"

'HiddenNames flags
Private Const TAG_FORMATTER_IMPORTED As String = "TAG_FORMATTER_IMPORTED"

'Named ranges on the Main sheet
Private Const RNG_LL_NAME As String = "RNG_LLName"
Private Const RNG_PATH_DICO As String = "RNG_PathDico"
Private Const RNG_LANG_SETUP As String = "RNG_LangSetup"
Private Const RNG_LL_FORM As String = "RNG_LLForm"

'Workbook-level HiddenName tags for language codes
Private Const TAG_DICT_LANG As String = "RNG_DictionaryLanguage"
Private Const TAG_LL_LANG_CODE As String = "RNG_LLLanguageCode"

'Anchor positions for setup workbook tables
Private Const SETUP_DICT_START_ROW As Long = 5
Private Const SETUP_DICT_START_COLUMN As Long = 1
Private Const SETUP_CHOICES_START_ROW As Long = 4
Private Const SETUP_CHOICES_START_COLUMN As Long = 1
Private Const SETUP_EXPORT_START_ROW As Long = 4
Private Const SETUP_EXPORT_START_COLUMN As Long = 1

'Anchor positions for exported data (row 1, col 1 in exported sheets)
Private Const EXPORTED_START_ROW As Long = 1
Private Const EXPORTED_START_COLUMN As Long = 1


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create and configure a linelist build service
'@details
'Factory method that creates a new LinelistBuildService instance, binds it
'to the supplied designer workbook, and returns the result through the
'ILinelistBuildService interface. When no workbook is provided, defaults
'to ThisWorkbook.
'@param hostBook Optional Workbook. The designer workbook. Defaults to ThisWorkbook.
'@return ILinelistBuildService. A configured instance ready for use.
'@export
Public Function Create(Optional ByVal hostBook As Workbook = Nothing) As ILinelistBuildService
    Dim instance As LinelistBuildService

    Set instance = New LinelistBuildService

    If hostBook Is Nothing Then
        instance.Configure ThisWorkbook
    Else
        instance.Configure hostBook
    End If

    Set Create = instance.Self
End Function

'@label:self
'@prop-title Current object instance
'@return ILinelistBuildService. The current instance cast to the interface.
Public Property Get Self() As ILinelistBuildService
    Set Self = Me
End Property


'@section Public API
'===============================================================================

'@label:configure
'@sub-title Bind the service to the designer workbook
'@param hostBook Workbook. The designer workbook to operate on.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook is required."
    End If

    Set this.hostBook = hostBook
    ClearExportedInstances
End Sub

'@label:host-workbook
'@prop-title The configured host workbook
'@return Workbook. The host workbook.
'@throws ProjectError.ObjectNotInitialized When the workbook was not configured.
'@export
Public Property Get HostWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook was not configured."
    End If
    Set HostWorkbook = this.hostBook
End Property

'@label:import-setup-to-linelist
'@sub-title Export pass-through components directly from setup to linelist
'@details
'Opens the setup workbook at the supplied path (read-only), then exports
'dictionary, choices, exports, analysis, and hidden names directly into
'the target linelist workbook. Translations dictionary data is imported
'from setup into the designer's LinelistTranslation sheet (the full
'translations sheet is exported later in ExportDesignerComponents).
'The formatter is exported from setup when TAG_FORMATTER_IMPORTED is not
'set on the designer. The setup workbook is always closed on exit.
'@param setupPath String. Full path to the setup workbook.
'@param targetBook Workbook. The linelist workbook to receive exported data.
'@throws ProjectError.SomethingWentWrong When setup cannot be opened or export fails.
'@export
Public Sub ImportSetupToLinelist(ByVal setupPath As String, ByVal targetBook As Workbook)
    Dim setupBook As Workbook
    Dim hostBook As Workbook
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    ValidateFilePath setupPath

    If targetBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target workbook is required."
    End If

    Set hostBook = HostWorkbook

    On Error Resume Next
    Set setupBook = Workbooks.Open(setupPath, ReadOnly:=True)
    On Error GoTo 0

    If setupBook Is Nothing Then
        ThrowError ProjectError.SomethingWentWrong, _
                   "Unable to open setup workbook at '" & setupPath & "'."
    End If

    On Error GoTo ImportFailed

    'Pass-through components: setup -> linelist (skip the designer)
    ExportDictionaryFromSetup setupBook, targetBook
    ExportChoicesFromSetup setupBook, targetBook
    ExportExportsFromSetup setupBook, targetBook
    ExportAnalysisFromSetup setupBook, targetBook
    ExportHiddenNamesFromSetup setupBook, targetBook

    'Formatter: export from setup when not ribbon-imported
    If Not FormatterAlreadyImported(hostBook) Then
        ExportFormatterFromSetup setupBook, targetBook
    End If

    'Translations: import dictionary from setup into designer sheet
    '(full translations sheet exported in ExportDesignerComponents)
    ImportTranslationsToDesigner setupBook, hostBook

ImportCleanup:
    CloseWorkbook setupBook
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

ImportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ImportCleanup
End Sub

'@label:export-designer-components
'@sub-title Export designer-sourced components to the linelist
'@details
'Syncs language codes from Main sheet into designer HiddenNames, then
'exports designer HiddenNames, geo, passwords, translations, and
'formatter (when ribbon-imported) to the target linelist workbook.
'@param targetBook Workbook. The linelist workbook to receive exported data.
'@throws ProjectError.ObjectNotInitialized When targetBook is Nothing.
'@export
Public Sub ExportDesignerComponents(ByVal targetBook As Workbook)
    Dim hostBook As Workbook

    If targetBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target workbook is required."
    End If

    Set hostBook = HostWorkbook

    'Sync language selections from Main sheet into designer HiddenNames,
    'then copy all designer HiddenNames to the target linelist workbook
    SyncDesignerLanguageNames hostBook
    ExportDesignerHiddenNames hostBook, targetBook

    'Designer-sourced components
    ExportGeoFromDesigner hostBook, targetBook
    ExportPasswordsFromDesigner hostBook, targetBook
    ExportTranslationsFromDesigner hostBook, targetBook

    'Formatter: export from designer when ribbon-imported
    If FormatterAlreadyImported(hostBook) Then
        ExportFormatterFromDesigner hostBook, targetBook
    End If
End Sub

'@label:build-exported-instances
'@sub-title Create domain manager instances from the exported linelist
'@details
'Creates new domain manager instances pointing to sheets in the target
'workbook. Each manager is created only when its corresponding sheet
'exists. Failures for individual managers are silently ignored so that
'remaining managers can still be created.
'@param targetBook Workbook. The linelist workbook containing exported data.
'@export
Public Sub BuildExportedInstances(ByVal targetBook As Workbook)
    Dim targetSheet As Worksheet

    ClearExportedInstances

    'Dictionary (export count is read from HiddenName __ll_exports_total__
    'persisted on the sheet by LLdictionary.Export)
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_DICTIONARY)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedDict = LLdictionary.Create(targetSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
        On Error GoTo 0
    End If

    'Exports
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_EXPORTS)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedExport = LLExport.Create(targetSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
        On Error GoTo 0
    End If

    'Choices
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_CHOICES)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedChoices = LLChoices.Create(targetSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
        On Error GoTo 0
    End If

    'Analysis
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_ANALYSIS)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedAnalysis = Analysis.Create(targetSheet)
        On Error GoTo 0
    End If

    'Geo
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_GEO)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedGeo = LLGeo.Create(targetSheet)
        On Error GoTo 0
    End If

    'Passwords
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_PASS)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedPass = Passwords.Create(targetSheet)
        On Error GoTo 0
    End If

    'Translations
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_LL_TRANSLATION)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedTrans = LLTranslation.Create(targetSheet)
        On Error GoTo 0
    End If
End Sub


'@section Exported Instance Properties
'===============================================================================

'@label:exported-dictionary
Public Property Get ExportedDictionary() As ILLdictionary
    Set ExportedDictionary = this.exportedDict
End Property

'@label:exported-choices
Public Property Get ExportedChoices() As ILLChoices
    Set ExportedChoices = this.exportedChoices
End Property

'@label:exported-analysis
Public Property Get ExportedAnalysis() As IAnalysis
    Set ExportedAnalysis = this.exportedAnalysis
End Property

'@label:exported-export
Public Property Get ExportedExport() As ILLExport
    Set ExportedExport = this.exportedExport
End Property

'@label:exported-geo
Public Property Get ExportedGeo() As ILLGeo
    Set ExportedGeo = this.exportedGeo
End Property

'@label:exported-passwords
Public Property Get ExportedPasswords() As IPasswords
    Set ExportedPasswords = this.exportedPass
End Property

'@label:exported-translations
Public Property Get ExportedTranslations() As ILLTranslation
    Set ExportedTranslations = this.exportedTrans
End Property


'@section Setup-to-Linelist Export Helpers (Direct Path)
'===============================================================================

'@label:export-dictionary-from-setup
'@sub-title Export dictionary directly from setup to linelist
'@details
'The data start row depends on the setup file format: .xlsb files
'have metadata rows above the data (row 5), while .xlsx files start
'at row 1. Exported sheet is set to very hidden.
Private Sub ExportDictionaryFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim dictManager As ILLdictionary

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_DICTIONARY)
    If setupSheet Is Nothing Then Exit Sub

    Set dictManager = LLdictionary.Create(setupSheet, _
                      SetupStartRow(setupBook, SETUP_DICT_START_ROW), SETUP_DICT_START_COLUMN)
    dictManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-choices-from-setup
'@sub-title Export choices directly from setup to linelist
Private Sub ExportChoicesFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim choicesManager As ILLChoices

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_CHOICES)
    If setupSheet Is Nothing Then Exit Sub

    Set choicesManager = LLChoices.Create(setupSheet, _
                         SetupStartRow(setupBook, SETUP_CHOICES_START_ROW), SETUP_CHOICES_START_COLUMN)
    choicesManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-exports-from-setup
'@sub-title Export export specs directly from setup to linelist
Private Sub ExportExportsFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim exportsManager As ILLExport

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_EXPORTS)
    If setupSheet Is Nothing Then Exit Sub

    Set exportsManager = LLExport.Create(setupSheet, _
                         SetupStartRow(setupBook, SETUP_EXPORT_START_ROW), SETUP_EXPORT_START_COLUMN)
    exportsManager.ExportSpecs targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-analysis-from-setup
'@sub-title Export analysis directly from setup to linelist
Private Sub ExportAnalysisFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim analysisManager As IAnalysis

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_ANALYSIS)
    If setupSheet Is Nothing Then Exit Sub

    Set analysisManager = Analysis.Create(setupSheet)
    analysisManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-hidden-names-from-setup
'@sub-title Export hidden names directly from setup to linelist
Private Sub ExportHiddenNamesFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim store As IHiddenNames

    Set store = HiddenNames.Create(setupBook)
    store.ExportNamesToWorkbook targetBook
End Sub

'@label:export-formatter-from-setup
'@sub-title Export formatter from setup to linelist (when not ribbon-imported)
Private Sub ExportFormatterFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim formatManager As ILLFormat

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_FORMATTER)
    If setupSheet Is Nothing Then Exit Sub

    Set formatManager = LLFormat.Create(setupSheet)
    formatManager.Export targetBook

    If WorksheetExists(targetBook, SHEET_FORMATTER) Then
        targetBook.Worksheets(SHEET_FORMATTER).Visible = xlSheetVeryHidden
    End If
End Sub

'@label:import-translations-to-designer
'@sub-title Import dictionary translations from setup into designer sheet
'@details
'The LinelistTranslation sheet template lives on the designer workbook
'with all its ListObjects pre-populated. This method imports the
'dictionary translations from the setup file into the designer sheet.
'The full translations sheet is exported to the linelist later in
'ExportDesignerComponents via ExportTranslationsFromDesigner.
Private Sub ImportTranslationsToDesigner(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
    Dim setupSheet As Worksheet
    Dim hostTransSheet As Worksheet
    Dim sourceTable As ListObject
    Dim translations As ILLTranslation

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_TRANSLATIONS)
    If setupSheet Is Nothing Then Exit Sub

    Set hostTransSheet = ResolveWorksheet(hostBook, SHEET_LL_TRANSLATION)
    If hostTransSheet Is Nothing Then Exit Sub

    If setupSheet.ListObjects.Count = 0 Then Exit Sub
    Set sourceTable = setupSheet.ListObjects(1)

    Set translations = LLTranslation.Create(hostTransSheet)
    translations.ImportDictionary sourceTable
End Sub


'@section Designer-to-Linelist Export Helpers
'===============================================================================

'@label:export-geo-from-designer
'@sub-title Export geo from designer to linelist
'@details
'Sets RNG_GeoLangCode and RNG_MetaLang on the designer's geo sheet
'before export so that UpdateLevelNames can resolve translations
'during ExportToWkb.
Private Sub ExportGeoFromDesigner(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim geoManager As ILLGeo
    Dim geoStore As IHiddenNames
    Dim dictLang As String
    Dim langCode As String

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_GEO)
    If hostSheet Is Nothing Then Exit Sub

    'Set language metadata on the designer's geo sheet before export
    dictLang = ReadMainRange(hostBook, RNG_LANG_SETUP)
    langCode = Split(ReadMainRange(hostBook, RNG_LL_FORM), "-")(0)
    Set geoStore = HiddenNames.Create(hostSheet)

    If LenB(dictLang) > 0 Then
        geoStore.EnsureName "RNG_MetaLang", dictLang, HiddenNameTypeString
        geoStore.SetValue "RNG_MetaLang", dictLang
    End If

    If LenB(langCode) > 0 Then
        geoStore.EnsureName "RNG_GeoLangCode", langCode, HiddenNameTypeString
        geoStore.SetValue "RNG_GeoLangCode", langCode
    End If

    Set geoManager = LLGeo.Create(hostSheet)
    geoManager.ExportToWkb targetBook, _
                           ReadMainRange(hostBook, RNG_LL_NAME), _
                           ReadMainRange(hostBook, RNG_PATH_DICO)
End Sub

'@label:export-passwords-from-designer
'@sub-title Export passwords from designer to linelist
Private Sub ExportPasswordsFromDesigner(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim passManager As IPasswords

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_PASS)
    If hostSheet Is Nothing Then Exit Sub

    Set passManager = Passwords.Create(hostSheet)
    passManager.ExportToWorkbook targetBook
End Sub

'@label:export-translations-from-designer
'@sub-title Export translations from designer to linelist
'@details
'Exports the full LinelistTranslation sheet (all ListObjects including
'the just-updated dictionary translations) from the designer to the
'linelist workbook.
Private Sub ExportTranslationsFromDesigner(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim translations As ILLTranslation

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_LL_TRANSLATION)
    If hostSheet Is Nothing Then Exit Sub

    Set translations = LLTranslation.Create(hostSheet)
    translations.Export targetBook
End Sub

'@label:export-formatter-from-designer
'@sub-title Export formatter from designer to linelist (when ribbon-imported)
Private Sub ExportFormatterFromDesigner(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim formatManager As ILLFormat

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_FORMATTER)
    If hostSheet Is Nothing Then Exit Sub

    Set formatManager = LLFormat.Create(hostSheet)
    formatManager.Export targetBook

    If WorksheetExists(targetBook, SHEET_FORMATTER) Then
        targetBook.Worksheets(SHEET_FORMATTER).Visible = xlSheetVeryHidden
    End If
End Sub


'@section Language and HiddenNames Helpers
'===============================================================================

'@label:sync-designer-language-names
'@sub-title Update language HiddenNames on the designer workbook from Main sheet values
'@details
'Reads the dictionary language from RNG_LangSetup and the interface
'language from RNG_LLForm on the Main worksheet. Writes these values
'as workbook-level HiddenNames (RNG_DictionaryLanguage and
'RNG_LLLanguageCode) on the host workbook so they are current before
'export. The interface language is split on "-" to extract the code
'prefix (e.g., "FRA-Francais" -> "FRA").
Private Sub SyncDesignerLanguageNames(ByVal hostBook As Workbook)
    Dim store As IHiddenNames
    Dim dictLang As String
    Dim llFormValue As String
    Dim llLang As String

    Set store = HiddenNames.Create(hostBook)

    dictLang = ReadMainRange(hostBook, RNG_LANG_SETUP)
    llFormValue = ReadMainRange(hostBook, RNG_LL_FORM)

    'Extract language code prefix (e.g. "FRA-Francais" -> "FRA")
    If InStr(1, llFormValue, "-", vbBinaryCompare) > 0 Then
        llLang = Split(llFormValue, "-")(0)
    Else
        llLang = llFormValue
    End If

    store.EnsureName TAG_DICT_LANG, dictLang, HiddenNameTypeString
    store.SetValue TAG_DICT_LANG, dictLang
    store.EnsureName TAG_LL_LANG_CODE, llLang, HiddenNameTypeString
    store.SetValue TAG_LL_LANG_CODE, llLang
End Sub

'@label:export-designer-hidden-names
'@sub-title Export designer workbook HiddenNames to the target linelist workbook
'@details
'Copies all workbook-level HiddenNames from the designer (host) workbook
'to the target linelist workbook, including language codes, ribbon flags,
'and other configuration values.
Private Sub ExportDesignerHiddenNames(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim store As IHiddenNames

    Set store = HiddenNames.Create(hostBook)
    store.ExportNamesToWorkbook targetBook
End Sub


'@section Internal Helpers
'===============================================================================

'@label:formatter-already-imported
'@sub-title Check whether the formatter was imported via the designer ribbon
'@return Boolean. True when TAG_FORMATTER_IMPORTED is "Yes" on the host workbook.
Private Function FormatterAlreadyImported(ByVal hostBook As Workbook) As Boolean
    Dim store As IHiddenNames

    Set store = HiddenNames.Create(hostBook)
    FormatterAlreadyImported = (store.ValueAsString(TAG_FORMATTER_IMPORTED) = "Yes")
End Function

'@label:resolve-worksheet
'@sub-title Resolve a worksheet by name from a workbook
'@return Worksheet. The resolved worksheet, or Nothing.
Private Function ResolveWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set ResolveWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:worksheet-exists
'@sub-title Check whether a worksheet exists in a workbook
Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    WorksheetExists = Not (ResolveWorksheet(wb, sheetName) Is Nothing)
End Function

'@label:setup-start-row
'@sub-title Return the data start row based on the setup file format
'@details
'Setup workbooks saved as .xlsb have metadata rows above the data
'tables (Dictionary at row 5, Choices/Exports at row 4). In .xlsx
'format the data always starts at row 1.
Private Function SetupStartRow(ByVal wb As Workbook, ByVal xlsbRow As Long) As Long
    If LCase$(Right$(wb.Name, 5)) = ".xlsb" Then
        SetupStartRow = xlsbRow
    Else
        SetupStartRow = 1
    End If
End Function

'@label:read-main-range
'@sub-title Read a named range value from the Main worksheet
'@return String. Cell value or vbNullString.
Private Function ReadMainRange(ByVal wb As Workbook, ByVal rngName As String) As String
    Dim sh As Worksheet
    Dim rng As Range

    Set sh = ResolveWorksheet(wb, SHEET_MAIN)
    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then Exit Function

    On Error Resume Next
    ReadMainRange = CStr(rng.Value)
    On Error GoTo 0
End Function

'@label:close-workbook
'@sub-title Close a workbook safely without raising errors
Private Sub CloseWorkbook(ByRef wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close saveChanges:=False
    End If
    On Error GoTo 0
    Set wb = Nothing
End Sub

'@label:validate-file-path
'@sub-title Validate that a file path is not empty and the file exists
Private Sub ValidateFilePath(ByVal filePath As String)
    Dim resolved As String

    If LenB(Trim$(filePath)) = 0 Then
        ThrowError ProjectError.ObjectNotInitialized, "Setup file path cannot be empty."
    End If

    On Error Resume Next
    resolved = Dir(filePath)
    On Error GoTo 0

    If LenB(resolved) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "Setup file not found at '" & filePath & "'."
    End If
End Sub

'@label:clear-exported-instances
'@sub-title Reset all exported domain manager references
Private Sub ClearExportedInstances()
    Set this.exportedDict = Nothing
    Set this.exportedChoices = Nothing
    Set this.exportedAnalysis = Nothing
    Set this.exportedExport = Nothing
    Set this.exportedGeo = Nothing
    Set this.exportedPass = Nothing
    Set this.exportedTrans = Nothing
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
Private Sub ThrowError(ByVal errorCode As ProjectError, ByVal messageText As String)
    Err.Raise CLng(errorCode), CLASS_NAME, messageText
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ILinelistBuildService contract.

Private Sub ILinelistBuildService_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Sub ILinelistBuildService_ImportSetupToLinelist(ByVal setupPath As String, ByVal targetBook As Workbook)
    ImportSetupToLinelist setupPath, targetBook
End Sub

Private Sub ILinelistBuildService_ExportDesignerComponents(ByVal targetBook As Workbook)
    ExportDesignerComponents targetBook
End Sub

Private Sub ILinelistBuildService_BuildExportedInstances(ByVal targetBook As Workbook)
    BuildExportedInstances targetBook
End Sub

Private Property Get ILinelistBuildService_ExportedDictionary() As ILLdictionary
    Set ILinelistBuildService_ExportedDictionary = ExportedDictionary
End Property

Private Property Get ILinelistBuildService_ExportedChoices() As ILLChoices
    Set ILinelistBuildService_ExportedChoices = ExportedChoices
End Property

Private Property Get ILinelistBuildService_ExportedAnalysis() As IAnalysis
    Set ILinelistBuildService_ExportedAnalysis = ExportedAnalysis
End Property

Private Property Get ILinelistBuildService_ExportedExport() As ILLExport
    Set ILinelistBuildService_ExportedExport = ExportedExport
End Property

Private Property Get ILinelistBuildService_ExportedGeo() As ILLGeo
    Set ILinelistBuildService_ExportedGeo = ExportedGeo
End Property

Private Property Get ILinelistBuildService_ExportedPasswords() As IPasswords
    Set ILinelistBuildService_ExportedPasswords = ExportedPasswords
End Property

Private Property Get ILinelistBuildService_ExportedTranslations() As ILLTranslation
    Set ILinelistBuildService_ExportedTranslations = ExportedTranslations
End Property

Private Property Get ILinelistBuildService_HostWorkbook() As Workbook
    Set ILinelistBuildService_HostWorkbook = HostWorkbook
End Property
