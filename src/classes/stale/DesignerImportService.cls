VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Designer import/export service for setup and linelist synchronisation"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Imports setup components into a designer workbook and exports them to a linelist workbook.")
'@depends ProjectError, LLdictionary, ILLdictionary, LLChoices, ILLChoices, LLExport, ILLExport, Analysis, IAnalysis, LLGeo, ILLGeo, IPasswords, Passwords, LLFormat, ILLFormat, LLTranslation, ILLTranslation, HiddenNames, IHiddenNames
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements IDesignerImportService

'@class DesignerImportService
'@description
'Orchestrates importing setup components (dictionary, choices, analysis,
'exports, formatter, translations) from a setup workbook into the designer
'workbook, and exporting components (dictionary, choices, analysis, exports,
'geo, passwords) to a linelist workbook. After ExportToLinelist, new domain
'manager instances targeting the exported workbook are available via
'accessor properties. This class is designed to replace the import/export
'step in LinelistSpecs in the long run.
'@depends LLdictionary, LLChoices, LLExport, Analysis, LLGeo, Passwords, LLFormat, LLTranslation, HiddenNames


'@section State
'===============================================================================

Private Type TDesignerImportService
    hostBook As Workbook            'Designer workbook
    exportedDict As ILLdictionary   'Dictionary manager on exported workbook
    exportedChoices As ILLChoices   'Choices manager on exported workbook
    exportedAnalysis As IAnalysis   'Analysis manager on exported workbook
    exportedExport As ILLExport    'Export manager on exported workbook
    exportedGeo As ILLGeo          'Geo manager on exported workbook
    exportedPass As IPasswords     'Passwords manager on exported workbook
    exportedTrans As ILLTranslation 'Translations manager on exported workbook
End Type

Private this As TDesignerImportService

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "DesignerImportService"

'Sheet names in the setup workbook
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_EXPORTS As String = "Exports"
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const SHEET_FORMATTER As String = "__formatter"

'Sheet names in the designer workbook (for translations import)
Private Const SHEET_LL_TRANSLATION As String = "LinelistTranslation"

'HiddenNames flags
Private Const TAG_FORMATTER_IMPORTED As String = "TAG_FORMATTER_IMPORTED"

'Sheet names in the designer workbook (for export to linelist)
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_PASS As String = "__pass"

'Named ranges on the Main sheet (geo metadata)
Private Const RNG_LL_NAME As String = "RNG_LLName"
Private Const RNG_PATH_DICO As String = "RNG_PathDico"

'Named ranges on the Main sheet (language selection)
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
'@sub-title Create and configure a designer import service
'@details
'Factory method that creates a new DesignerImportService instance, binds it
'to the supplied designer workbook, and returns the result through the
'IDesignerImportService interface. When no workbook is provided, defaults
'to ThisWorkbook.
'@param hostBook Optional Workbook. The designer workbook. Defaults to ThisWorkbook.
'@return IDesignerImportService. A configured instance ready for use.
'@export
Public Function Create(Optional ByVal hostBook As Workbook = Nothing) As IDesignerImportService
    Dim instance As DesignerImportService

    Set instance = New DesignerImportService

    If hostBook Is Nothing Then
        instance.Configure ThisWorkbook
    Else
        instance.Configure hostBook
    End If

    Set Create = instance.Self
End Function

'@label:self
'@prop-title Current object instance
'@return IDesignerImportService. The current instance cast to the interface.
Public Property Get Self() As IDesignerImportService
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

'@label:import-from-setup
'@sub-title Import setup components into the host workbook
'@details
'Opens the setup workbook at the supplied path (read-only), then imports
'each component by creating domain managers from the setup workbook and
'exporting their data into the host workbook. Translations are imported
'via LLTranslation.ImportDictionary. The setup workbook is always closed
'on exit, even when errors occur.
'@param setupPath String. Full path to the setup workbook.
'@throws ProjectError.SomethingWentWrong When the setup workbook cannot be opened or import fails.
'@export
Public Sub ImportFromSetup(ByVal setupPath As String)
    Dim setupBook As Workbook
    Dim hostBook As Workbook
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    ValidateFilePath setupPath
    Set hostBook = HostWorkbook

    On Error Resume Next
    Set setupBook = Workbooks.Open(setupPath, ReadOnly:=True)
    On Error GoTo 0

    If setupBook Is Nothing Then
        ThrowError ProjectError.SomethingWentWrong, _
                   "Unable to open setup workbook at '" & setupPath & "'."
    End If

    On Error GoTo ImportFailed

    ImportDictionary setupBook, hostBook
    ImportChoices setupBook, hostBook
    ImportExports setupBook, hostBook
    ImportAnalysis setupBook, hostBook
    ImportFormatter setupBook, hostBook
    ImportTranslations setupBook, hostBook
    ImportHiddenNames setupBook, hostBook

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

'@label:export-to-linelist
'@sub-title Export designer components to a linelist workbook
'@details
'Exports dictionary, choices, analysis, exports, geo, and passwords from
'the host workbook into the provided target workbook. The target workbook
'is NOT closed. After export, new domain manager instances targeting the
'exported workbook are stored on the service and accessible via Exported*
'properties. Any error during export is re-raised after cleanup.
'@param targetBook Workbook. The linelist workbook to receive exported data.
'@throws ProjectError.ObjectNotInitialized When targetBook is Nothing.
'@throws ProjectError.SomethingWentWrong When export fails.
'@export
Public Sub ExportToLinelist(ByVal targetBook As Workbook)
    Dim hostBook As Workbook
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If targetBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target workbook is required."
    End If

    Set hostBook = HostWorkbook
    ClearExportedInstances

    On Error GoTo ExportFailed

    'Sync language selections from Main sheet into designer HiddenNames,
    'then copy all designer HiddenNames to the target linelist workbook
    SyncDesignerLanguageNames hostBook
    ExportDesignerHiddenNames hostBook, targetBook

    ExportDictionary hostBook, targetBook
    ExportChoices hostBook, targetBook
    ExportExports hostBook, targetBook
    ExportAnalysis hostBook, targetBook
    ExportGeo hostBook, targetBook, ReadMainRange(hostBook, RNG_LL_NAME), _
                                     ReadMainRange(hostBook, RNG_PATH_DICO)
    ExportPasswords hostBook, targetBook
    ExportTranslations hostBook, targetBook
    ExportFormatter hostBook, targetBook

    BuildExportedInstances targetBook

    Exit Sub

ExportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    ClearExportedInstances
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
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


'@section Import Helpers (Setup -> Designer)
'===============================================================================

'@label:import-dictionary
'@sub-title Export dictionary from setup into the host workbook
'@details
'The data start row depends on the setup file format: .xlsb files
'have metadata rows above the data (row 5), while .xlsx files start
'at row 1.
Private Sub ImportDictionary(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
    Dim setupSheet As Worksheet
    Dim dictManager As ILLdictionary

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_DICTIONARY)
    If setupSheet Is Nothing Then Exit Sub

    Set dictManager = LLdictionary.Create(setupSheet, _
                      SetupStartRow(setupBook, SETUP_DICT_START_ROW), SETUP_DICT_START_COLUMN)
    dictManager.Export hostBook, Hide:=xlSheetHidden
End Sub

'@label:import-choices
'@sub-title Export choices from setup into the host workbook
Private Sub ImportChoices(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
    Dim setupSheet As Worksheet
    Dim choicesManager As ILLChoices

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_CHOICES)
    If setupSheet Is Nothing Then Exit Sub

    Set choicesManager = LLChoices.Create(setupSheet, _
                         SetupStartRow(setupBook, SETUP_CHOICES_START_ROW), SETUP_CHOICES_START_COLUMN)
    choicesManager.Export hostBook, Hide:=xlSheetHidden
End Sub

'@label:import-exports
'@sub-title Export export specs from setup into the host workbook
Private Sub ImportExports(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
    Dim setupSheet As Worksheet
    Dim exportsManager As ILLExport

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_EXPORTS)
    If setupSheet Is Nothing Then Exit Sub

    Set exportsManager = LLExport.Create(setupSheet, _
                         SetupStartRow(setupBook, SETUP_EXPORT_START_ROW), SETUP_EXPORT_START_COLUMN)
    exportsManager.ExportSpecs hostBook, Hide:=xlSheetHidden
End Sub

'@label:import-analysis
'@sub-title Export analysis from setup into the host workbook
Private Sub ImportAnalysis(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
    Dim setupSheet As Worksheet
    Dim analysisManager As IAnalysis

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_ANALYSIS)
    If setupSheet Is Nothing Then Exit Sub

    Set analysisManager = Analysis.Create(setupSheet)
    analysisManager.Export hostBook, Hide:=xlSheetHidden
End Sub

'@label:import-formatter
'@sub-title Export formatter from setup into the host workbook
'@details
'Skips import when the formatter was already imported via the ribbon
'(clickImpStyle sets TAG_FORMATTER_IMPORTED on the host workbook).
Private Sub ImportFormatter(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
    Dim setupSheet As Worksheet
    Dim formatManager As ILLFormat
    Dim store As IHiddenNames

    'Skip if the formatter was already imported via the ribbon
    Set store = HiddenNames.Create(hostBook)
    If store.ValueAsString(TAG_FORMATTER_IMPORTED) = "Yes" Then Exit Sub

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_FORMATTER)
    If setupSheet Is Nothing Then Exit Sub

    Set formatManager = LLFormat.Create(setupSheet)
    formatManager.Export hostBook

    'Set the exported worksheet to hidden (not very hidden)
    If WorksheetExists(hostBook, SHEET_FORMATTER) Then
        hostBook.Worksheets(SHEET_FORMATTER).Visible = xlSheetHidden
    End If
End Sub

'@label:import-translations
'@sub-title Import translations from setup via LLTranslation.ImportDictionary
'@details
'Gets the first ListObject from the setup Translations worksheet and
'imports it into the designer LinelistTranslation sheet using
'LLTranslation.ImportDictionary. Silently exits when either sheet or
'ListObject is missing.
Private Sub ImportTranslations(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
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

'@label:import-hidden-names
'@sub-title Export hidden names from setup into the host workbook
Private Sub ImportHiddenNames(ByVal setupBook As Workbook, ByVal hostBook As Workbook)
    Dim store As IHiddenNames

    Set store = HiddenNames.Create(setupBook)
    store.ExportNamesToWorkbook hostBook
End Sub


'@section Export Helpers (Designer -> Linelist)
'===============================================================================

'@label:export-dictionary
'@sub-title Export dictionary from host to target workbook
'@details
'Data in the designer workbook always starts at row 1, col 1 (written
'by DataSheet.Export during ImportFromSetup), regardless of the
'original setup file format.
Private Sub ExportDictionary(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim dictManager As ILLdictionary

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_DICTIONARY)
    If hostSheet Is Nothing Then Exit Sub

    Set dictManager = LLdictionary.Create(hostSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
    dictManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-choices
'@sub-title Export choices from host to target workbook
Private Sub ExportChoices(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim choicesManager As ILLChoices

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_CHOICES)
    If hostSheet Is Nothing Then Exit Sub

    Set choicesManager = LLChoices.Create(hostSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
    choicesManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-exports
'@sub-title Export export specs from host to target workbook
Private Sub ExportExports(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim exportsManager As ILLExport

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_EXPORTS)
    If hostSheet Is Nothing Then Exit Sub

    Set exportsManager = LLExport.Create(hostSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
    exportsManager.ExportSpecs targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-analysis
'@sub-title Export analysis from host to target workbook
Private Sub ExportAnalysis(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim analysisManager As IAnalysis

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_ANALYSIS)
    If hostSheet Is Nothing Then Exit Sub

    Set analysisManager = Analysis.Create(hostSheet)
    analysisManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@label:export-geo
'@sub-title Export geo from host to target workbook
'@param hostBook Workbook. Designer workbook containing the Geo sheet.
'@param targetBook Workbook. Linelist workbook receiving the geo data.
'@param llName String. Linelist name written as geo metadata.
'@param setupPath String. Setup file path written as geo metadata.
Private Sub ExportGeo(ByVal hostBook As Workbook, ByVal targetBook As Workbook, _
                      ByVal llName As String, ByVal setupPath As String)
    Dim hostSheet As Worksheet
    Dim geoManager As ILLGeo
    Dim geoStore As IHiddenNames
    Dim dictLang As String
    Dim langCode As String

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_GEO)
    If hostSheet Is Nothing Then Exit Sub

    'Set RNG_GeoLangCode on the designer's geo sheet before export
    'so that UpdateLevelNames can resolve translations during ExportToWkb
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
    geoManager.ExportToWkb targetBook, llName, setupPath
End Sub

'@label:export-passwords
'@sub-title Export passwords from host to target workbook
Private Sub ExportPasswords(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim passManager As IPasswords

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_PASS)
    If hostSheet Is Nothing Then Exit Sub

    Set passManager = Passwords.Create(hostSheet)
    passManager.ExportToWorkbook targetBook
End Sub

'@label:export-translations
'@sub-title Export translations from host to target workbook
Private Sub ExportTranslations(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim hostSheet As Worksheet
    Dim translations As ILLTranslation

    Set hostSheet = ResolveWorksheet(hostBook, SHEET_LL_TRANSLATION)
    If hostSheet Is Nothing Then Exit Sub

    Set translations = LLTranslation.Create(hostSheet)
    translations.Export targetBook
End Sub

'@label:export-formatter
'@sub-title Export formatter from host to target workbook
Private Sub ExportFormatter(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
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


'@section Post-Export Instance Building
'===============================================================================

'@label:build-exported-instances
'@sub-title Create domain manager instances from the exported workbook
'@details
'After ExportToLinelist writes sheets into the target workbook, this method
'creates new domain manager instances pointing to those sheets. Each manager
'is created only when its corresponding sheet exists in the target workbook.
'Failures for individual managers are silently ignored so that remaining
'managers can still be created.
Private Sub BuildExportedInstances(ByVal targetBook As Workbook)
    Dim targetSheet As Worksheet

    'Dictionary
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_DICTIONARY)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedDict = LLdictionary.Create(targetSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
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

    'Exports
    Set targetSheet = ResolveWorksheet(targetBook, SHEET_EXPORTS)
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        Set this.exportedExport = LLExport.Create(targetSheet, EXPORTED_START_ROW, EXPORTED_START_COLUMN)
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


'@section Internal Helpers
'===============================================================================

'@label:sync-designer-language-names
'@sub-title Update language HiddenNames on the designer workbook from Main sheet values
'@details
'Reads the dictionary language from RNG_LangSetup and the interface
'language from RNG_LLForm on the Main worksheet. Writes these values
'as workbook-level HiddenNames (RNG_DictionaryLanguage and
'RNG_LLLanguageCode) on the host workbook so they are current before
'export. The interface language is split on "-" to extract the code
'prefix (e.g., "FRA-Fran" & ChrW$(231) & "ais" -> "FRA").
'@param hostBook Workbook. The designer workbook to update.
Private Sub SyncDesignerLanguageNames(ByVal hostBook As Workbook)
    Dim store As IHiddenNames
    Dim dictLang As String
    Dim llFormValue As String
    Dim llLang As String

    Set store = HiddenNames.Create(hostBook)

    dictLang = ReadMainRange(hostBook, RNG_LANG_SETUP)
    llFormValue = ReadMainRange(hostBook, RNG_LL_FORM)

    'Extract language code prefix (e.g. "FRA-Français" -> "FRA")
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
'and other configuration values. This ensures the linelist workbook has
'RNG_DictionaryLanguage and RNG_LLLanguageCode available before any
'LLTranslation.TransObject call.
'@param hostBook Workbook. The designer workbook (source).
'@param targetBook Workbook. The linelist workbook (destination).
Private Sub ExportDesignerHiddenNames(ByVal hostBook As Workbook, ByVal targetBook As Workbook)
    Dim store As IHiddenNames

    Set store = HiddenNames.Create(hostBook)
    store.ExportNamesToWorkbook targetBook
End Sub

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
'tables (Dictionary at row 5, Choices/Exports at row 4).  In .xlsx
'format the data always starts at row 1.
'@param wb Workbook. The setup workbook to inspect.
'@param xlsbRow Long. Start row to use for .xlsb files.
'@return Long. The resolved start row (xlsbRow for .xlsb, 1 otherwise).
Private Function SetupStartRow(ByVal wb As Workbook, ByVal xlsbRow As Long) As Long
    If LCase$(Right$(wb.Name, 5)) = ".xlsb" Then
        SetupStartRow = xlsbRow
    Else
        SetupStartRow = 1
    End If
End Function

'@label:read-main-range
'@sub-title Read a named range value from the Main worksheet
'@details
'Returns the string value of a named range on the Main sheet of the
'supplied workbook.  Returns vbNullString when the sheet or range is
'missing.
'@param wb Workbook. Workbook containing the Main sheet.
'@param rngName String. Named range to read.
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

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
Private Sub ThrowError(ByVal errorCode As ProjectError, ByVal messageText As String)
    Err.Raise CLng(errorCode), CLASS_NAME, messageText
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IDesignerImportService contract.

Private Sub IDesignerImportService_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Sub IDesignerImportService_ImportFromSetup(ByVal setupPath As String)
    ImportFromSetup setupPath
End Sub

Private Sub IDesignerImportService_ExportToLinelist(ByVal targetBook As Workbook)
    ExportToLinelist targetBook
End Sub

Private Property Get IDesignerImportService_ExportedDictionary() As ILLdictionary
    Set IDesignerImportService_ExportedDictionary = ExportedDictionary
End Property

Private Property Get IDesignerImportService_ExportedChoices() As ILLChoices
    Set IDesignerImportService_ExportedChoices = ExportedChoices
End Property

Private Property Get IDesignerImportService_ExportedAnalysis() As IAnalysis
    Set IDesignerImportService_ExportedAnalysis = ExportedAnalysis
End Property

Private Property Get IDesignerImportService_ExportedExport() As ILLExport
    Set IDesignerImportService_ExportedExport = ExportedExport
End Property

Private Property Get IDesignerImportService_ExportedGeo() As ILLGeo
    Set IDesignerImportService_ExportedGeo = ExportedGeo
End Property

Private Property Get IDesignerImportService_ExportedPasswords() As IPasswords
    Set IDesignerImportService_ExportedPasswords = ExportedPasswords
End Property

Private Property Get IDesignerImportService_ExportedTranslations() As ILLTranslation
    Set IDesignerImportService_ExportedTranslations = ExportedTranslations
End Property

Private Property Get IDesignerImportService_HostWorkbook() As Workbook
    Set IDesignerImportService_HostWorkbook = HostWorkbook
End Property
