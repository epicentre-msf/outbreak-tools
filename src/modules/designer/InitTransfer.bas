Attribute VB_Name = "InitTransfer"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Fills a new linelist workbook straight from the setup file and the designer, with no intermediate copy.")
'@depends ProjectError, LLdictionary, LLChoices, LLExport, Analysis, LLGeo, LLFormat, LLTranslation, Passwords, HiddenNames
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'Puts everything a generation needs into a freshly created linelist workbook,
'then hands out one domain manager per exported sheet. LinelistSpecs.Prepare is
'the only caller.
'
'WHERE EACH COMPONENT COMES FROM
'-------------------------------------------------------------------------------
'The dictionary, the choices, the exports, the analyses and the setup's hidden
'names travel from the setup file straight into the linelist. They are read
'once, at their setup anchors, and written at row 1 of a very hidden sheet in
'the linelist. Nothing lands on the designer on the way.
'
'The geo sheet, the passwords, the five translation tables and the designer's
'own hidden names come from the designer. The formatter comes from whichever
'of the two holds the live copy: the designer when it was loaded through the
'ribbon, the setup otherwise.
'
'WHY THE SETUP TRANSLATION TABLE IS ITS OWN STEP
'-------------------------------------------------------------------------------
'Four of the five translation tables translate the linelist itself, and the
'designer owns them: they are edited by hand there, through the designer's own
'translation import button, and a generation only copies them out.
'
'The fifth, Tab_Translations, is the SETUP's table, and it is what translates
'the dictionary, the choices and the analyses. So it lands last, on the
'linelist's translation sheet, once the designer's four have arrived. A
'generation never writes on the designer's translation sheet.
'
'WHY THIS IS A MODULE
'-------------------------------------------------------------------------------
'It replaced the LLBuild class. That class held its seven exported managers in a
'UDT only to carry them from one method call to the next, and its one caller
'copied all seven out again the moment they were built. Nothing here outlives a
'single call, so nothing here needs an instance.


'@section Constants
'===============================================================================

Private Const MODULE_NAME As String = "InitTransfer"

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


'@section Public API
'===============================================================================

'@sub-title Fill a new linelist workbook from the setup file and the designer
'@details
'Runs the whole transfer in one pass:
'1. The pass-through components go from the setup straight into the linelist:
'   dictionary, choices, exports, analyses and the setup's hidden names.
'2. The designer components follow: its hidden names, the geo sheet, the
'   passwords and the five translation tables. The designer's hidden names are
'   written after the setup's on purpose, so a name both workbooks carry comes
'   out of the linelist holding the designer's answer.
'3. The setup translation table lands on the LINELIST's translation sheet.
'4. The formatter comes from the designer when it was loaded through the
'   ribbon, and from the setup otherwise.
'
'The setup workbook is opened read-only and is always closed on the way out,
'including when a step raises.
'@param designerBook Workbook. The designer workbook.
'@param setupPath String. Full path to the setup workbook.
'@param targetBook Workbook. The linelist workbook receiving everything.
'@throws ProjectError.ObjectNotInitialized When either workbook is Nothing or
'the setup path is empty.
'@throws ProjectError.ElementNotFound When no file sits at the setup path.
'@throws ProjectError.SomethingWentWrong When the setup workbook cannot be opened.
Public Sub TransferToLinelist(ByVal designerBook As Workbook, _
                              ByVal setupPath As String, _
                              ByVal targetBook As Workbook)
    Dim setupBook As Workbook
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If designerBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Designer workbook is required."
    End If

    If targetBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target workbook is required."
    End If

    ValidateFilePath setupPath

    On Error Resume Next
    Set setupBook = Workbooks.Open(setupPath, ReadOnly:=True)
    On Error GoTo 0

    If setupBook Is Nothing Then
        ThrowError ProjectError.SomethingWentWrong, _
                   "Unable to open setup workbook at '" & setupPath & "'."
    End If

    On Error GoTo TransferFailed

    'Pass-through components: setup -> linelist
    ExportDictionaryFromSetup setupBook, targetBook
    ExportChoicesFromSetup setupBook, targetBook
    ExportExportsFromSetup setupBook, targetBook
    ExportAnalysisFromSetup setupBook, targetBook
    ExportHiddenNamesFromSetup setupBook, targetBook

    'Designer components: designer -> linelist. The hidden names go first of
    'these five, after the setup's, because the designer's answer wins.
    SyncDesignerLanguageNames designerBook
    ExportDesignerHiddenNames designerBook, targetBook
    ExportGeoFromDesigner designerBook, targetBook
    ExportPasswordsFromDesigner designerBook, targetBook
    ExportTranslationsFromDesigner designerBook, targetBook

    'The setup's own translation table, onto the linelist's translation sheet.
    'It runs after the export because the sheet it writes on is the one the
    'export just created.
    ImportSetupTranslationTable setupBook, targetBook

    'The formatter lives on one of the two workbooks, never on both.
    If FormatterAlreadyImported(designerBook) Then
        ExportFormatterFromDesigner designerBook, targetBook
    Else
        ExportFormatterFromSetup setupBook, targetBook
    End If

TransferCleanup:
    'This label is reached with the handler still armed, so it opens with an
    'On Error line of its own.
    On Error Resume Next
    CloseWorkbook setupBook
    On Error GoTo 0

    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

TransferFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume TransferCleanup
End Sub


'@section The linelist domain managers
'===============================================================================
'@description
'One builder per exported sheet, each answering a manager that points at the
'LINELIST workbook. A builder answers Nothing when the sheet it needs is absent
'and raises when the sheet is there but malformed. The silent Nothing that the
'old class handed back on a failed Create is gone: it let a caller carry on
'against the designer's copy of a sheet with nothing said.

'@fun-title The dictionary manager on the linelist
'@details
'The export count is read from the HiddenName __ll_exports_total__ that
'LLdictionary.Export leaves on the sheet.
'@param targetBook Workbook. The linelist workbook.
'@return LLdictionary. The manager, or Nothing when the sheet is absent.
Public Function LinelistDictionary(ByVal targetBook As Workbook) As LLdictionary
    Dim targetSheet As Worksheet

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_DICTIONARY)
    If targetSheet Is Nothing Then Exit Function

    Set LinelistDictionary = LLdictionary.Create(targetSheet, _
                                                 EXPORTED_START_ROW, EXPORTED_START_COLUMN)
End Function

'@fun-title The choices manager on the linelist
'@param targetBook Workbook. The linelist workbook.
'@return LLChoices. The manager, or Nothing when the sheet is absent.
Public Function LinelistChoices(ByVal targetBook As Workbook) As LLChoices
    Dim targetSheet As Worksheet

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_CHOICES)
    If targetSheet Is Nothing Then Exit Function

    Set LinelistChoices = LLChoices.Create(targetSheet, _
                                           EXPORTED_START_ROW, EXPORTED_START_COLUMN)
End Function

'@fun-title The export specifications manager on the linelist
'@param targetBook Workbook. The linelist workbook.
'@return LLExport. The manager, or Nothing when the sheet is absent.
Public Function LinelistExports(ByVal targetBook As Workbook) As LLExport
    Dim targetSheet As Worksheet

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_EXPORTS)
    If targetSheet Is Nothing Then Exit Function

    Set LinelistExports = LLExport.Create(targetSheet, _
                                          EXPORTED_START_ROW, EXPORTED_START_COLUMN)
End Function

'@fun-title The analyses manager on the linelist
'@param targetBook Workbook. The linelist workbook.
'@return Analysis. The manager, or Nothing when the sheet is absent.
Public Function LinelistAnalysis(ByVal targetBook As Workbook) As Analysis
    Dim targetSheet As Worksheet

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_ANALYSIS)
    If targetSheet Is Nothing Then Exit Function

    Set LinelistAnalysis = Analysis.Create(targetSheet)
End Function

'@fun-title The geo manager on the linelist
'@param targetBook Workbook. The linelist workbook.
'@return LLGeo. The manager, or Nothing when the sheet is absent.
Public Function LinelistGeo(ByVal targetBook As Workbook) As LLGeo
    Dim targetSheet As Worksheet

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_GEO)
    If targetSheet Is Nothing Then Exit Function

    Set LinelistGeo = LLGeo.Create(targetSheet)
End Function

'@fun-title The passwords manager on the linelist
'@param targetBook Workbook. The linelist workbook.
'@return Passwords. The manager, or Nothing when the sheet is absent.
Public Function LinelistPasswords(ByVal targetBook As Workbook) As Passwords
    Dim targetSheet As Worksheet

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_PASS)
    If targetSheet Is Nothing Then Exit Function

    Set LinelistPasswords = Passwords.Create(targetSheet)
End Function

'@fun-title The translations manager on the linelist
'@param targetBook Workbook. The linelist workbook.
'@return LLTranslation. The manager, or Nothing when the sheet is absent.
Public Function LinelistTranslations(ByVal targetBook As Workbook) As LLTranslation
    Dim targetSheet As Worksheet

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_LL_TRANSLATION)
    If targetSheet Is Nothing Then Exit Function

    Set LinelistTranslations = LLTranslation.Create(targetSheet)
End Function


'@section Setup to linelist
'===============================================================================

'@sub-title Export the dictionary from the setup to the linelist
'@details
'The data start row depends on the setup file format: .xlsb files carry
'metadata rows above the data (row 5), .xlsx files start at row 1. The
'exported sheet is set very hidden.
Private Sub ExportDictionaryFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim dictManager As LLdictionary

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_DICTIONARY)
    If setupSheet Is Nothing Then Exit Sub

    Set dictManager = LLdictionary.Create(setupSheet, _
                      SetupStartRow(setupBook, SETUP_DICT_START_ROW), SETUP_DICT_START_COLUMN)
    dictManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@sub-title Export the choices from the setup to the linelist
Private Sub ExportChoicesFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim choicesManager As LLChoices

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_CHOICES)
    If setupSheet Is Nothing Then Exit Sub

    Set choicesManager = LLChoices.Create(setupSheet, _
                         SetupStartRow(setupBook, SETUP_CHOICES_START_ROW), SETUP_CHOICES_START_COLUMN)
    choicesManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@sub-title Export the export specifications from the setup to the linelist
Private Sub ExportExportsFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim exportsManager As LLExport

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_EXPORTS)
    If setupSheet Is Nothing Then Exit Sub

    Set exportsManager = LLExport.Create(setupSheet, _
                         SetupStartRow(setupBook, SETUP_EXPORT_START_ROW), SETUP_EXPORT_START_COLUMN)
    exportsManager.ExportSpecs targetBook, Hide:=xlSheetVeryHidden
End Sub

'@sub-title Export the analyses from the setup to the linelist
Private Sub ExportAnalysisFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim analysisManager As Analysis

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_ANALYSIS)
    If setupSheet Is Nothing Then Exit Sub

    Set analysisManager = Analysis.Create(setupSheet)
    analysisManager.Export targetBook, Hide:=xlSheetVeryHidden
End Sub

'@sub-title Export the setup's hidden names to the linelist
Private Sub ExportHiddenNamesFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(setupBook)
    store.ExportNamesToWorkbook targetBook
End Sub

'@sub-title Export the formatter from the setup to the linelist
Private Sub ExportFormatterFromSetup(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim formatManager As LLFormat

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_FORMATTER)
    If setupSheet Is Nothing Then Exit Sub

    Set formatManager = LLFormat.Create(setupSheet)
    formatManager.Export targetBook

    If WorksheetExists(targetBook, SHEET_FORMATTER) Then
        targetBook.Worksheets(SHEET_FORMATTER).Visible = xlSheetVeryHidden
    End If
End Sub

'@sub-title Put the setup's translation table on the linelist's translation sheet
'@details
'Tab_Translations is the setup's own table, and it is what translates the
'dictionary, the choices and the analyses. It overwrites the copy that arrived
'with the designer's translation sheet, and it keeps the setup's own headers,
'so the linelist comes out carrying the languages the setup carries.
'
'The designer's translation sheet is not touched. It holds the four tables that
'translate the linelist itself, it is maintained by hand through the designer's
'own import button, and a generation only reads it.
Private Sub ImportSetupTranslationTable(ByVal setupBook As Workbook, ByVal targetBook As Workbook)
    Dim setupSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceTable As ListObject
    Dim translations As LLTranslation

    Set setupSheet = ResolveWorksheet(setupBook, SHEET_TRANSLATIONS)
    If setupSheet Is Nothing Then Exit Sub
    If setupSheet.ListObjects.Count = 0 Then Exit Sub

    Set targetSheet = ResolveWorksheet(targetBook, SHEET_LL_TRANSLATION)
    If targetSheet Is Nothing Then Exit Sub

    Set sourceTable = setupSheet.ListObjects(1)
    Set translations = LLTranslation.Create(targetSheet)
    translations.ImportDictionary sourceTable
End Sub


'@section Designer to linelist
'===============================================================================

'@sub-title Export the geo sheet from the designer to the linelist
'@details
'RNG_MetaLang and RNG_GeoLangCode are written on the designer's geo sheet
'first. The meta language is the one picked on the Main sheet, out of the
'languages the setup translation table carries, and UpdateLevelNames needs both
'to resolve the level names while ExportToWkb runs. The geobase itself is
'imported later, on the linelist's own copy of the sheet.
Private Sub ExportGeoFromDesigner(ByVal designerBook As Workbook, ByVal targetBook As Workbook)
    Dim designerSheet As Worksheet
    Dim geoManager As LLGeo
    Dim geoStore As HiddenNames
    Dim dictLang As String
    Dim langCode As String

    Set designerSheet = ResolveWorksheet(designerBook, SHEET_GEO)
    If designerSheet Is Nothing Then Exit Sub

    dictLang = ReadMainRange(designerBook, RNG_LANG_SETUP)
    langCode = Split(ReadMainRange(designerBook, RNG_LL_FORM), "-")(0)
    Set geoStore = HiddenNames.Create(designerSheet)

    If LenB(dictLang) > 0 Then
        geoStore.EnsureName "RNG_MetaLang", dictLang, HiddenNameTypeString
        geoStore.SetValue "RNG_MetaLang", dictLang
    End If

    If LenB(langCode) > 0 Then
        geoStore.EnsureName "RNG_GeoLangCode", langCode, HiddenNameTypeString
        geoStore.SetValue "RNG_GeoLangCode", langCode
    End If

    Set geoManager = LLGeo.Create(designerSheet)
    geoManager.ExportToWkb targetBook, _
                           ReadMainRange(designerBook, RNG_LL_NAME), _
                           ReadMainRange(designerBook, RNG_PATH_DICO)
End Sub

'@sub-title Export the passwords from the designer to the linelist
Private Sub ExportPasswordsFromDesigner(ByVal designerBook As Workbook, ByVal targetBook As Workbook)
    Dim designerSheet As Worksheet
    Dim passManager As Passwords

    Set designerSheet = ResolveWorksheet(designerBook, SHEET_PASS)
    If designerSheet Is Nothing Then Exit Sub

    Set passManager = Passwords.Create(designerSheet)
    passManager.ExportToWorkbook targetBook
End Sub

'@sub-title Export the five translation tables from the designer to the linelist
'@details
'Copies the whole LinelistTranslation sheet across. Four of the five tables
'arrive as their final content; Tab_Translations is overwritten right after by
'ImportSetupTranslationTable.
Private Sub ExportTranslationsFromDesigner(ByVal designerBook As Workbook, ByVal targetBook As Workbook)
    Dim designerSheet As Worksheet
    Dim translations As LLTranslation

    Set designerSheet = ResolveWorksheet(designerBook, SHEET_LL_TRANSLATION)
    If designerSheet Is Nothing Then Exit Sub

    Set translations = LLTranslation.Create(designerSheet)
    translations.Export targetBook
End Sub

'@sub-title Export the formatter from the designer to the linelist
Private Sub ExportFormatterFromDesigner(ByVal designerBook As Workbook, ByVal targetBook As Workbook)
    Dim designerSheet As Worksheet
    Dim formatManager As LLFormat

    Set designerSheet = ResolveWorksheet(designerBook, SHEET_FORMATTER)
    If designerSheet Is Nothing Then Exit Sub

    Set formatManager = LLFormat.Create(designerSheet)
    formatManager.Export targetBook

    If WorksheetExists(targetBook, SHEET_FORMATTER) Then
        targetBook.Worksheets(SHEET_FORMATTER).Visible = xlSheetVeryHidden
    End If
End Sub


'@section Languages and hidden names
'===============================================================================

'@sub-title Write the two language codes onto the designer workbook
'@details
'Reads the dictionary language from RNG_LangSetup and the interface language
'from RNG_LLForm on the Main worksheet, and writes both as workbook-level
'hidden names on the designer so they are current before the export. The
'interface language is split on "-" to take the code prefix, so
'"FRA-Francais" comes out as "FRA".
Private Sub SyncDesignerLanguageNames(ByVal designerBook As Workbook)
    Dim store As HiddenNames
    Dim dictLang As String
    Dim llFormValue As String
    Dim llLang As String

    Set store = HiddenNames.Create(designerBook)

    dictLang = ReadMainRange(designerBook, RNG_LANG_SETUP)
    llFormValue = ReadMainRange(designerBook, RNG_LL_FORM)

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

'@sub-title Export the designer's workbook hidden names to the linelist
'@details
'Carries the language codes, the ribbon flags and the rest of the designer's
'configuration values. It runs after the setup's own export, so a name both
'workbooks carry comes out of the linelist holding the designer's answer.
Private Sub ExportDesignerHiddenNames(ByVal designerBook As Workbook, ByVal targetBook As Workbook)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(designerBook)
    store.ExportNamesToWorkbook targetBook
End Sub


'@section Internal helpers
'===============================================================================

'@fun-title Whether the formatter was loaded through the designer ribbon
'@return Boolean. True when TAG_FORMATTER_IMPORTED reads "Yes" on the designer.
Private Function FormatterAlreadyImported(ByVal designerBook As Workbook) As Boolean
    Dim store As HiddenNames

    Set store = HiddenNames.Create(designerBook)
    FormatterAlreadyImported = (store.ValueAsString(TAG_FORMATTER_IMPORTED) = "Yes")
End Function

'@fun-title Resolve a worksheet by name
'@return Worksheet. The worksheet, or Nothing.
Private Function ResolveWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set ResolveWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@fun-title Whether a workbook carries a worksheet of that name
Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    WorksheetExists = Not (ResolveWorksheet(wb, sheetName) Is Nothing)
End Function

'@fun-title The data start row for a setup table
'@details
'Setup workbooks saved as .xlsb carry metadata rows above the data tables, so
'the dictionary starts at row 5 and the choices and exports at row 4. In .xlsx
'format the data always starts at row 1.
Private Function SetupStartRow(ByVal wb As Workbook, ByVal xlsbRow As Long) As Long
    If LCase$(Right$(wb.Name, 5)) = ".xlsb" Then
        SetupStartRow = xlsbRow
    Else
        SetupStartRow = 1
    End If
End Function

'@fun-title Read a named range from the Main worksheet
'@return String. The cell value, or an empty string.
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

'@sub-title Close a workbook without raising
Private Sub CloseWorkbook(ByRef wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close saveChanges:=False
    End If
    On Error GoTo 0
    Set wb = Nothing
End Sub

'@sub-title Check that a file path is filled in and points at a file
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

'@sub-title Raise a ProjectError with this module as the source
Private Sub ThrowError(ByVal errorCode As Long, ByVal messageText As String)
    Err.Raise errorCode, MODULE_NAME, messageText
End Sub
