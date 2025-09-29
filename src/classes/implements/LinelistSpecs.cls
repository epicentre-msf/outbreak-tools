VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Modern facade exposing linelist specifications with caching and validation")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILinelistSpecs

Private Type TLinelistSpecs
    specsWkb As Workbook
    dictionary As ILLdictionary
    choices As ILLChoices
    translations As ILLTranslations
    designerTranslations As IDesTranslation
    geo As ILLGeo
    exports As ILLExport
    analysis As ILLAnalysis
    passwords As ILLPasswords
    format As ILLFormat
    formulaData As IFormulaData
    main As IMain
    variables As ILLVariables
    translationCache As Collection
    categoryCache As Collection
    sheetCache As Collection
End Type

Private this As TLinelistSpecs

Private Const CLASSNAME As String = "LinelistSpecs"
Private Const CATEGORY_CACHE_SEPARATOR As String = "|"
Private Const DEFAULT_DICT_START_ROW As Long = 1
Private Const DEFAULT_DICT_START_COLUMN As Long = 1
Private Const DEFAULT_CHOICES_START_ROW As Long = 1
Private Const DEFAULT_CHOICES_START_COLUMN As Long = 1

'@section Factory helpers
'===============================================================================
'@description Create a new specification facade while validating workbook structure upfront.
Public Function Create(ByVal specsWorkbook As Workbook) As ILinelistSpecs
    ValidateWorkbookReference specsWorkbook

    With New LinelistSpecs
        Set .InternalWorkbook = specsWorkbook
        .ResetCaches
        Set Create = .Self
    End With
End Function

'@description Return the current instance cast as its interface for fluent construction patterns.
Public Property Get Self() As ILinelistSpecs
    Set Self = Me
End Property

'@description Primary workbook backing all specification queries.
Public Property Get Workbook() As Workbook
    EnsureWorkbook
    Set Workbook = this.specsWkb
End Property

Friend Property Set InternalWorkbook(ByVal specsWorkbook As Workbook)
    Set this.specsWkb = specsWorkbook
End Property

'@section State management
'===============================================================================
'@description Reset every cached collaborator so subsequent lookups hit the workbook again.
Public Sub ResetCaches()
    ResetObjectCaches
    ResetSheetCache
    ResetTranslationCache
    ResetCategoryCache
End Sub

Private Sub ResetObjectCaches()
    Set this.dictionary = Nothing
    Set this.choices = Nothing
    Set this.translations = Nothing
    Set this.designerTranslations = Nothing
    Set this.geo = Nothing
    Set this.exports = Nothing
    Set this.analysis = Nothing
    Set this.passwords = Nothing
    Set this.format = Nothing
    Set this.formulaData = Nothing
    Set this.main = Nothing
    Set this.variables = Nothing
End Sub

Private Sub ResetSheetCache()
    Set this.sheetCache = Nothing
End Sub

Private Sub ResetTranslationCache()
    Set this.translationCache = Nothing
End Sub

Private Sub ResetCategoryCache()
    Set this.categoryCache = Nothing
End Sub

'@section Guard helpers
'===============================================================================
Private Sub EnsureWorkbook()
    If this.specsWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications workbook has not been assigned"
    End If
End Sub

Private Sub EnsureSheetCache()
    If this.sheetCache Is Nothing Then Set this.sheetCache = New Collection
End Sub

Private Sub EnsureTranslationCache()
    If this.translationCache Is Nothing Then Set this.translationCache = New Collection
End Sub

Private Sub EnsureCategoryCache()
    If this.categoryCache Is Nothing Then Set this.categoryCache = New Collection
End Sub

'@section Worksheet resolution
'===============================================================================
'@description Normalise sheet names for cache keys.
Private Function SheetCacheKey(ByVal sheetName As String) As String
    SheetCacheKey = LCase$(Trim$(sheetName))
End Function

'@description Resolve a worksheet by name with optional caching and error surfacing.
Private Function ResolveSheet(ByVal sheetName As String, _
                              Optional ByVal allowMissing As Boolean = False) As Worksheet

    Dim key As String
    Dim cachedSheet As Worksheet
    Dim target As Worksheet
    Dim normalisedName As String

    normalisedName = Trim$(sheetName)
    If LenB(normalisedName) = 0 Then
        If allowMissing Then Exit Function
        ThrowError ProjectError.InvalidArgument, "Worksheet name cannot be empty"
    End If

    EnsureWorkbook
    EnsureSheetCache

    key = SheetCacheKey(normalisedName)

    On Error Resume Next
        Set cachedSheet = this.sheetCache.Item(key)
    On Error GoTo 0

    If Not (cachedSheet Is Nothing) Then
        Set ResolveSheet = cachedSheet
        Exit Function
    End If

    On Error Resume Next
        Set target = this.specsWkb.Worksheets(normalisedName)
    On Error GoTo 0

    If target Is Nothing Then
        If allowMissing Then Exit Function
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & normalisedName & "' not found in specifications workbook '" & SafeWorkbookName(this.specsWkb) & "'"
    End If

    On Error Resume Next
        this.sheetCache.Add target, key
    On Error GoTo 0

    Set ResolveSheet = target
End Function

'@section External workbook helpers
'===============================================================================
Private Sub EnterBusyState(ByRef busyStateApplied As Boolean)
    If busyStateApplied Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False

    busyStateApplied = True
End Sub

Private Sub ExitBusyState(ByRef busyStateApplied As Boolean)
    If Not busyStateApplied Then Exit Sub

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True

    busyStateApplied = False
End Sub

Private Function OpenWorkbookInternal(ByVal workbookPath As String, _
                                      ByRef busyStateApplied As Boolean) As Workbook
    Dim target As Workbook
    Dim trimmedPath As String

    trimmedPath = Trim$(workbookPath)
    If LenB(trimmedPath) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Workbook path has not been configured"
    End If

    If LenB(Dir$(trimmedPath)) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Workbook '" & trimmedPath & "' cannot be located"
    End If

    EnterBusyState busyStateApplied

    On Error GoTo HandleOpen
        Set target = Workbooks.Open(trimmedPath)
    On Error GoTo 0

    Set OpenWorkbookInternal = target
    Exit Function

HandleOpen:
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Unable to open workbook '" & trimmedPath & "': " & Err.Description
End Function

Private Sub CloseWorkbookQuiet(ByRef target As Workbook)
    On Error Resume Next
        If Not target Is Nothing Then target.Close savechanges:=False
    On Error GoTo 0
    Set target = Nothing
End Sub

Private Function ResolveExternalSheet(ByVal hostBook As Workbook, _
                                      ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set ResolveExternalSheet = hostBook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub EnsureExternalSheetExists(ByVal hostBook As Workbook, _
                                      ByVal sheetName As String)
    If ResolveExternalSheet(hostBook, sheetName) Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & sheetName & "' is missing from workbook '" & SafeWorkbookName(hostBook) & "'"
    End If
End Sub

Private Sub EnsureSetupWorkbookSheets(ByVal setupBook As Workbook)
    Dim requiredSheets As Variant
    Dim idx As Long

    requiredSheets = Array( _
        LinelistSpecConstants.SHEET_DICTIONARY, _
        LinelistSpecConstants.SHEET_CHOICES, _
        LinelistSpecConstants.SHEET_TRANSLATIONS_DICT, _
        LinelistSpecConstants.SHEET_EXPORT, _
        LinelistSpecConstants.SHEET_ANALYSIS)

    For idx = LBound(requiredSheets) To UBound(requiredSheets)
        EnsureExternalSheetExists setupBook, CStr(requiredSheets(idx))
    Next idx
End Sub

Private Sub EnsureDictionaryNotEmpty(ByVal dictionarySheet As Worksheet)
    Dim importedDict As ILLdictionary
    Dim dataRange As Range

    Set importedDict = LLdictionary.Create(dictionarySheet, 5, 1)

    On Error Resume Next
        Set dataRange = importedDict.DataRange
    On Error GoTo 0

    If dataRange Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Dictionary worksheet '" & dictionarySheet.Name & "' does not expose any data"
    End If

    If Application.WorksheetFunction.CountA(dataRange) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Dictionary worksheet '" & dictionarySheet.Name & "' is empty"
    End If
End Sub

Private Function BuildPreservedSheetNames(ByVal messageTranslations As ITranslationObject) As BetterArray
    Dim preserved As BetterArray

    Set preserved = New BetterArray
    preserved.LowerBound = 1

    preserved.Push LinelistSpecConstants.SHEET_DICTIONARY, _
                   LinelistSpecConstants.SHEET_CHOICES, _
                   LinelistSpecConstants.SHEET_GEO, _
                   LinelistSpecConstants.SHEET_PASSWORDS, _
                   LinelistSpecConstants.SHEET_FORMULAS, _
                   LinelistSpecConstants.SHEET_TRANSLATIONS_DICT, _
                   LinelistSpecConstants.SHEET_TRANSLATIONS_LL, _
                   LinelistSpecConstants.SHEET_EXPORT, _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_SINGLE), _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_LIST), _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_UPDATES), _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_METADATA), _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_ANALYSIS), _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_IMPORT), _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_SPATIAL), _
                   TemporarySheetName(LinelistSpecConstants.TEMP_SHEET_SHOWHIDE)

    If Not messageTranslations Is Nothing Then
        preserved.Push SafeTranslatedValue(messageTranslations, "LLSHEET_Admin"), _
                       SafeTranslatedValue(messageTranslations, "LLSHEET_Analysis"), _
                       SafeTranslatedValue(messageTranslations, "LLSHEET_SpatialAnalysis"), _
                       SafeTranslatedValue(messageTranslations, "LLSHEET_TemporalAnalysis"), _
                       SafeTranslatedValue(messageTranslations, "LLSHEET_CustomChoice"), _
                       SafeTranslatedValue(messageTranslations, "LLSHEET_CustomPivotTable"), _
                       SafeTranslatedValue(messageTranslations, "LLSHEET_SpatioTemporalAnalysis")
    End If

    Set BuildPreservedSheetNames = preserved
End Function

Private Function SafeTranslatedValue(ByVal translations As ITranslationObject, ByVal key As String) As String
    On Error Resume Next
        SafeTranslatedValue = translations.TranslatedValue(key)
    On Error GoTo 0

    If LenB(SafeTranslatedValue) = 0 Then SafeTranslatedValue = key
End Function

Private Sub ResetPostImportCaches()
    Set this.variables = Nothing
    ResetCategoryCache
    ResetTranslationCache
End Sub

'@section Workbook validation
'===============================================================================
Private Sub ValidateWorkbookReference(ByVal specsWorkbook As Workbook)
    If specsWorkbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "A specifications workbook reference is required"
    End If

    ValidateWorkbookStructure specsWorkbook
End Sub

Private Sub ValidateWorkbookStructure(ByVal specsWorkbook As Workbook)
    Dim requiredSheets As Variant
    Dim missingSheets As BetterArray
    Dim idx As Long
    Dim sheetName As String

    requiredSheets = RequiredSheetNames()
    Set missingSheets = New BetterArray
    missingSheets.LowerBound = 1

    For idx = LBound(requiredSheets) To UBound(requiredSheets)
        sheetName = CStr(requiredSheets(idx))
        If Not WorksheetExistsInternal(specsWorkbook, sheetName) Then
            missingSheets.Push sheetName
        End If
    Next idx

    If missingSheets.Length > 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Specifications workbook '" & SafeWorkbookName(specsWorkbook) & "' is missing required worksheets: " & JoinNames(missingSheets)
    End If

    Dim formatSheet As Worksheet
    On Error Resume Next
        Set formatSheet = specsWorkbook.Worksheets(LinelistSpecConstants.SHEET_FORMAT)
    On Error GoTo 0

    If formatSheet Is Nothing Then Exit Sub

    If Not RangeExistsInternal(formatSheet, LinelistSpecConstants.RANGE_DESIGN_TYPE) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & LinelistSpecConstants.SHEET_FORMAT & "' must define the '" & LinelistSpecConstants.RANGE_DESIGN_TYPE & "' named range"
    End If
End Sub

Private Function RequiredSheetNames() As Variant
    RequiredSheetNames = Array( _
        LinelistSpecConstants.SHEET_DICTIONARY, _
        LinelistSpecConstants.SHEET_CHOICES, _
        LinelistSpecConstants.SHEET_GEO, _
        LinelistSpecConstants.SHEET_PASSWORDS, _
        LinelistSpecConstants.SHEET_FORMULAS, _
        LinelistSpecConstants.SHEET_TRANSLATIONS_DICT, _
        LinelistSpecConstants.SHEET_TRANSLATIONS_LL, _
        LinelistSpecConstants.SHEET_ANALYSIS, _
        LinelistSpecConstants.SHEET_EXPORT, _
        LinelistSpecConstants.SHEET_FORMAT, _
        LinelistSpecConstants.SHEET_MAIN, _
        LinelistSpecConstants.SHEET_DESIGNER_TRANSLATION)
End Function

Private Function WorksheetExistsInternal(ByVal specsWorkbook As Workbook, _
                                         ByVal sheetName As String) As Boolean
    Dim target As Worksheet
    On Error Resume Next
        Set target = specsWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExistsInternal = Not (target Is Nothing)
End Function

Private Function RangeExistsInternal(ByVal hostSheet As Worksheet, _
                                     ByVal rangeName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
        Set rng = hostSheet.Range(rangeName)
    On Error GoTo 0
    RangeExistsInternal = Not (rng Is Nothing)
End Function

Private Function JoinNames(ByVal values As BetterArray) As String
    Dim buffer As String
    Dim idx As Long

    For idx = values.LowerBound To values.UpperBound
        If LenB(buffer) > 0 Then buffer = buffer & ", "
        buffer = buffer & "'" & CStr(values.Item(idx)) & "'"
    Next idx

    JoinNames = buffer
End Function

Private Function SafeWorkbookName(ByVal specsWorkbook As Workbook) As String
    On Error Resume Next
        SafeWorkbookName = specsWorkbook.Name
    On Error GoTo 0
    If LenB(SafeWorkbookName) = 0 Then SafeWorkbookName = "<unknown>"
End Function

'@section Collaborator ensures
'===============================================================================
Private Sub EnsureExports()
    If this.exports Is Nothing Then
        Dim exportSheet As Worksheet
        Set exportSheet = ResolveSheet(LinelistSpecConstants.SHEET_EXPORT)
        Set this.exports = LLExport.Create(exportSheet)
    End If
End Sub

Private Sub EnsureDictionary()
    If this.dictionary Is Nothing Then
        EnsureExports
        Dim dictSheet As Worksheet
        Set dictSheet = ResolveSheet(LinelistSpecConstants.SHEET_DICTIONARY)
        Set this.dictionary = LLdictionary.Create(dictSheet, _
                                                  DEFAULT_DICT_START_ROW, _
                                                  DEFAULT_DICT_START_COLUMN, _
                                                  NumberOfExports())
    End If
End Sub

Private Sub EnsureChoices()
    If this.choices Is Nothing Then
        Dim choicesSheet As Worksheet
        Set choicesSheet = ResolveSheet(LinelistSpecConstants.SHEET_CHOICES)
        Set this.choices = LLChoices.Create(choicesSheet, _
                                            DEFAULT_CHOICES_START_ROW, _
                                            DEFAULT_CHOICES_START_COLUMN)
    End If
End Sub

Private Sub EnsureTranslations()
    If this.translations Is Nothing Then
        Dim linelistTransSheet As Worksheet
        Dim dictTransSheet As Worksheet
        Set linelistTransSheet = ResolveSheet(LinelistSpecConstants.SHEET_TRANSLATIONS_LL)
        Set dictTransSheet = ResolveSheet(LinelistSpecConstants.SHEET_TRANSLATIONS_DICT)
        Set this.translations = LLTranslations.Create(linelistTransSheet, dictTransSheet)
    End If
End Sub

Private Sub EnsureDesignerTranslations()
    If this.designerTranslations Is Nothing Then
        Dim designerSheet As Worksheet
        Set designerSheet = ResolveSheet(LinelistSpecConstants.SHEET_DESIGNER_TRANSLATION)
        Set this.designerTranslations = DesTranslation.Create(designerSheet)
    End If
End Sub

Private Sub EnsureGeo()
    If this.geo Is Nothing Then
        Dim geoSheet As Worksheet
        Set geoSheet = ResolveSheet(LinelistSpecConstants.SHEET_GEO)
        Set this.geo = LLGeo.Create(geoSheet)
    End If
End Sub

Private Sub EnsureAnalysis()
    If this.analysis Is Nothing Then
        Dim analysisSheet As Worksheet
        Set analysisSheet = ResolveSheet(LinelistSpecConstants.SHEET_ANALYSIS)
        Set this.analysis = Analysis.Create(analysisSheet)
    End If
End Sub

Private Sub EnsurePasswords()
    If this.passwords Is Nothing Then
        Dim passwordSheet As Worksheet
        Set passwordSheet = ResolveSheet(LinelistSpecConstants.SHEET_PASSWORDS)
        Set this.passwords = Passwords.Create(passwordSheet)
    End If
End Sub

Private Sub EnsureFormulaData()
    If this.formulaData Is Nothing Then
        Dim formulaSheet As Worksheet
        Set formulaSheet = ResolveSheet(LinelistSpecConstants.SHEET_FORMULAS)
        Set this.formulaData = FormulaData.Create(formulaSheet)
    End If
End Sub

Private Sub EnsureMain()
    If this.main Is Nothing Then
        Dim mainSheet As Worksheet
        Set mainSheet = ResolveSheet(LinelistSpecConstants.SHEET_MAIN)
        Set this.main = Main.Create(mainSheet)
    End If
End Sub

Private Sub EnsureVariables()
    EnsureDictionary
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(this.dictionary)
    End If
End Sub

Private Sub EnsureFormat()
    If this.format Is Nothing Then
        Dim formatSheet As Worksheet
        Dim designName As String

        Set formatSheet = ResolveSheet(LinelistSpecConstants.SHEET_FORMAT)
        designName = ResolveDesignName(formatSheet)

        Set this.format = LLFormat.Create(formatSheet, designName)
    End If
End Sub

Private Function ResolveDesignName(ByVal formatSheet As Worksheet) As String
    Dim designRange As Range

    On Error Resume Next
        Set designRange = formatSheet.Range(LinelistSpecConstants.RANGE_DESIGN_TYPE)
    On Error GoTo 0

    If designRange Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & LinelistSpecConstants.RANGE_DESIGN_TYPE & "' missing from format worksheet"
    End If

    ResolveDesignName = Trim$(CStr(designRange.Cells(1, 1).Value))
End Function

Private Sub EnsureTranslationObjectCache()
    EnsureTranslations
    EnsureTranslationCache
End Sub

Private Sub EnsureCategoryLookup()
    EnsureChoices
    EnsureVariables
    EnsureCategoryCache
End Sub

Private Function NumberOfExports() As Long
    EnsureExports
    NumberOfExports = this.exports.NumberOfExports()
End Function

'@section List auto helpers
'===============================================================================
Private Sub AddListAuto()
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim listVars As BetterArray
    Dim idx As Long
    Dim varName As String
    Dim controlType As String

    Set dict = Dictionary()
    EnsureVariables
    Set vars = this.variables

    dict.AddColumn "list auto"
    Set listVars = dict.Data.FilterData("control", "list_auto", "control details")

    If listVars Is Nothing Then Exit Sub

    For idx = listVars.LowerBound To listVars.UpperBound
        varName = CStr(listVars.Item(idx))

        vars.SetValue varName:=varName, colName:="list auto", newValue:="list_auto_origin", onEmpty:=True

        controlType = LCase$(Trim$(vars.Value(varName:=varName, colName:="control")))
        Select Case controlType
            Case "formula", "case_when", "choice_formula"
                RecursiveListAuto varName
        End Select
    Next idx
End Sub

Private Sub RecursiveListAuto(ByVal varName As String)
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim controlType As String
    Dim controlDetails As String
    Dim formulaEngine As IFormulas
    Dim referencedVars As BetterArray
    Dim idx As Long
    Dim referencedName As String
    Dim referencedControl As String

    Set dict = Dictionary()
    EnsureVariables
    Set vars = this.variables

    controlType = LCase$(Trim$(vars.Value(varName:=varName, colName:="control")))
    If controlType <> "case_when" And controlType <> "formula" And controlType <> "choice_formula" Then Exit Sub

    controlDetails = Trim$(vars.Value(varName:=varName, colName:="control details"))
    If LenB(controlDetails) = 0 Then Exit Sub

    Set formulaEngine = Formulas.Create(dict, FormulaDataObject(), controlDetails)
    If Not formulaEngine.Valid("linelist") Then Exit Sub

    Set referencedVars = formulaEngine.Varlists("linelist")
    If referencedVars Is Nothing Then Exit Sub

    For idx = referencedVars.LowerBound To referencedVars.UpperBound
        referencedName = CStr(referencedVars.Item(idx))

        vars.SetValue varName:=referencedName, colName:="list auto", newValue:="list_auto_origin", onEmpty:=True

        referencedControl = LCase$(Trim$(vars.Value(varName:=referencedName, colName:="control")))
        If referencedControl = "formula" Or referencedControl = "case_when" Then
            RecursiveListAuto referencedName
        End If
    Next idx
End Sub

'@section Category helpers
'===============================================================================
Public Function Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Dim cacheKey As String
    Dim cached As BetterArray
    Dim clone As BetterArray

    EnsureCategoryLookup

    cacheKey = BuildCategoryCacheKey(varName, useShortlabels)

    On Error Resume Next
        Set cached = this.categoryCache.Item(cacheKey)
    On Error GoTo 0

    If cached Is Nothing Then
        Set cached = ResolveCategoryList(varName, useShortlabels)
        On Error Resume Next
            this.categoryCache.Add cached, cacheKey
        On Error GoTo 0
    End If

    Set clone = cached.Clone
    Set Categories = clone
End Function

Private Function ResolveCategoryList(ByVal varName As String, _
                                     ByVal useShortlabels As Boolean) As BetterArray
    Dim normalizedName As String
    Dim controlValue As String
    Dim controlDetails As String
    Dim categoryName As String
    Dim result As BetterArray

    normalizedName = Trim$(varName)
    If LenB(normalizedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Variable name is required to resolve categories"
    End If

    controlValue = NormaliseControlType(this.variables.Value(varName:=normalizedName, colName:="control"))
    controlDetails = Trim$(this.variables.Value(varName:=normalizedName, colName:="control details"))

    Select Case controlValue
        Case "choice_formula"
            categoryName = ExtractChoiceFormula(controlDetails)
        Case "choice_manual", "choice_multiple", "choice_custom"
            categoryName = controlDetails
        Case Else
            ThrowError ProjectError.ElementNotFound, _
                      "Variable '" & normalizedName & "' does not expose selectable categories"
    End Select

    categoryName = Trim$(categoryName)
    If LenB(categoryName) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Variable '" & normalizedName & "' references an empty category definition"
    End If

    Set result = this.choices.Categories(categoryName, useShortlabels)
    If result Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Categories for '" & categoryName & "' are missing from the choices sheet"
    End If

    Set ResolveCategoryList = result
End Function

Private Function BuildCategoryCacheKey(ByVal varName As String, _
                                       ByVal useShortlabels As Boolean) As String
    BuildCategoryCacheKey = LCase$(Trim$(varName)) & CATEGORY_CACHE_SEPARATOR & CStr(useShortlabels)
End Function

Private Function NormaliseControlType(ByVal controlValue As String) As String
    Dim trimmed As String
    Dim delimiterPos As Long

    trimmed = LCase$(Trim$(controlValue))
    delimiterPos = InStr(1, trimmed, "(", vbBinaryCompare)
    If delimiterPos > 0 Then
        trimmed = Left$(trimmed, delimiterPos - 1)
    End If

    NormaliseControlType = trimmed
End Function

Private Function ExtractChoiceFormula(ByVal controlDetails As String) As String
    Dim working As String
    Dim commaPosition As Long

    working = controlDetails
    working = Replace$(working, "CHOICE_FORMULA(", vbNullString, 1, -1, vbTextCompare)

    commaPosition = InStr(1, working, ",", vbBinaryCompare)
    If commaPosition > 0 Then
        working = Left$(working, commaPosition - 1)
    End If

    working = Replace$(working, ")", vbNullString, 1, -1, vbBinaryCompare)
    ExtractChoiceFormula = Trim$(working)
End Function

'@section Core accessors
'===============================================================================
Public Property Get Dictionary() As ILLdictionary
    EnsureDictionary
    Set Dictionary = this.dictionary
End Property

Public Property Get Choices() As ILLChoices
    EnsureChoices
    Set Choices = this.choices
End Property

Public Property Get DesignFormat() As ILLFormat
    EnsureFormat
    Set DesignFormat = this.format
End Property

Public Property Get Analysis() As ILLAnalysis
    EnsureAnalysis
    Set Analysis = this.analysis
End Property

Public Property Get ExportObject() As ILLExport
    EnsureExports
    Set ExportObject = this.exports
End Property

Public Property Get Password() As ILLPasswords
    EnsurePasswords
    Set Password = this.passwords
End Property

Public Property Get GeoObject() As ILLGeo
    EnsureGeo
    Set GeoObject = this.geo
End Property

Public Property Get FormulaDataObject() As IFormulaData
    EnsureFormulaData
    Set FormulaDataObject = this.formulaData
End Property

Public Property Get MainObject() As IMain
    EnsureMain
    Set MainObject = this.main
End Property

Public Property Get Translations() As ILLTranslations
    EnsureTranslations
    Set Translations = this.translations
End Property

Public Property Get DesignerTranslations() As IDesTranslation
    EnsureDesignerTranslations
    Set DesignerTranslations = this.designerTranslations
End Property

Public Property Get TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As ITranslationObject
    Dim cacheKey As String
    Dim cached As ITranslationObject
    Dim resolved As ITranslationObject

    EnsureTranslationObjectCache

    cacheKey = CStr(translationScope)

    On Error Resume Next
        Set cached = this.translationCache.Item(cacheKey)
    On Error GoTo 0

    If cached Is Nothing Then
        Set resolved = this.translations.TransObject(translationScope)
        If resolved Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                      "Translation scope '" & CStr(translationScope) & "' is not available in the specifications workbook"
        End If
        On Error Resume Next
            this.translationCache.Add resolved, cacheKey
        On Error GoTo 0
        Set cached = resolved
    End If

    Set TransObject = cached
End Property

Public Property Get TemporarySheetName(ByVal scope As Byte) As String
    Select Case scope
        Case LinelistSpecConstants.TEMP_SHEET_ANALYSIS
            TemporarySheetName = "ana_tabnames__"
        Case LinelistSpecConstants.TEMP_SHEET_UPDATES
            TemporarySheetName = "updates__"
        Case LinelistSpecConstants.TEMP_SHEET_METADATA
            TemporarySheetName = "Metadata"
        Case LinelistSpecConstants.TEMP_SHEET_LIST
            TemporarySheetName = "dropdown_lists__"
        Case LinelistSpecConstants.TEMP_SHEET_IMPORT
            TemporarySheetName = "import_rep__"
        Case LinelistSpecConstants.TEMP_SHEET_SINGLE
            TemporarySheetName = "temp__"
        Case LinelistSpecConstants.TEMP_SHEET_SPATIAL
            TemporarySheetName = "spatial_tables__"
        Case LinelistSpecConstants.TEMP_SHEET_SHOWHIDE
            TemporarySheetName = "show_hide__"
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                      "Unknown temporary sheet scope '" & CStr(scope) & "'"
    End Select
End Property

Public Property Get Value(ByVal tagName As String) As String
    Dim normalizedTag As String

    normalizedTag = LCase$(Trim$(tagName))
    If LenB(normalizedTag) = 0 Then Exit Property

    Select Case normalizedTag
        Case "numberofexports"
            Value = CStr(NumberOfExports())
        Case Else
            EnsureMain
            Value = this.main.ValueOf(tagName)
    End Select
End Property

'@section Writer factories
'===============================================================================
Public Function CreateSectionContext(ByVal linelist As ILinelist, ByVal startRow As Long) As ILLSectionContext
    If linelist Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist reference is required to build a section context"
    End If
    If startRow <= 0 Then
        ThrowError ProjectError.InvalidArgument, "Section context start row must be greater than zero"
    End If

    Dim context As ILLSectionContext
    Set context = New LLSectionContext
    context.Initialise linelist, startRow
    Set CreateSectionContext = context
End Function

Public Function CreateSectionWriter() As ILLSectionWriter
    Dim writer As ILLSectionWriter
    Set writer = New LLSectionWriterBase
    Set CreateSectionWriter = writer
End Function

Public Function CreateSubSectionWriter() As ILLSubSectionWriter
    Dim writer As ILLSubSectionWriter
    Set writer = New LLSubSectionWriterBase
    Set CreateSubSectionWriter = writer
End Function

Public Function CreateVarContext(ByVal linelist As ILinelist, ByVal variableName As String) As ILLVarContext
    Dim trimmedName As String

    If linelist Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist reference is required to build a variable context"
    End If

    trimmedName = Trim$(variableName)
    If LenB(trimmedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Variable name is required to create a variable context"
    End If

    Set CreateVarContext = LLVarContext.Create(linelist, trimmedName)
End Function

Public Function CreateVarWriter() As ILLVarWriter
    Dim writer As ILLVarWriter
    Set writer = New LLVarWriterBase
    Set CreateVarWriter = writer
End Function

'@section Workbook orchestration
'===============================================================================
Public Sub Prepare()
    Dim mainObj As IMain
    Dim designerTrads As IDesTranslation
    Dim setupPath As String
    Dim geoPath As String
    Dim dict As ILLdictionary
    Dim choices As ILLChoices
    Dim exports As ILLExport
    Dim translations As ILLTranslations
    Dim analysisObj As ILLAnalysis
    Dim geoObj As ILLGeo
    Dim dictTrads As ITranslationObject
    Dim msgTrads As ITranslationObject
    Dim preservedNames As BetterArray
    Dim setupBook As Workbook
    Dim geoBook As Workbook
    Dim busyStateApplied As Boolean
    Dim prepared As Boolean
    Dim dictionarySheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim analysisSheet As Worksheet

    On Error GoTo CleanFail

    Set mainObj = MainObject()
    Set designerTrads = DesignerTranslations()

    mainObj.CheckReadiness designerTrads
    If Not mainObj.Ready Then GoTo CleanExit

    setupPath = Trim$(mainObj.ValueOf("setuppath"))
    If LenB(setupPath) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Setup workbook path is not configured on the Main sheet"
    End If

    geoPath = Trim$(mainObj.ValueOf("geopath"))

    Set dict = Dictionary()
    Set choices = Choices()
    Set exports = ExportObject()
    Set translations = Translations()
    Set analysisObj = Analysis()
    Set geoObj = GeoObject()

    Set setupBook = OpenWorkbookInternal(setupPath, busyStateApplied)
    EnsureSetupWorkbookSheets setupBook

    exports.ImportSpecs setupBook.Worksheets(LinelistSpecConstants.SHEET_EXPORT), 4, 1

    Set dictionarySheet = setupBook.Worksheets(LinelistSpecConstants.SHEET_DICTIONARY)
    EnsureDictionaryNotEmpty dictionarySheet
    dict.Import dictionarySheet, fromStartRow:=5, fromStartcol:=1

    Set choicesSheet = setupBook.Worksheets(LinelistSpecConstants.SHEET_CHOICES)
    choices.Import choicesSheet, fromStartRow:=4, fromStartcol:=1

    translations.ImportTranslation setupBook

    Set analysisSheet = setupBook.Worksheets(LinelistSpecConstants.SHEET_ANALYSIS)
    analysisObj.Import analysisSheet

    CloseWorkbookQuiet setupBook

    If LenB(geoPath) > 0 Then
        Set geoBook = OpenWorkbookInternal(geoPath, busyStateApplied)
        geoObj.Import geoBook
        CloseWorkbookQuiet geoBook
    End If

    Set dictTrads = TransObject(TranslationOfDictionary)
    Set msgTrads = TransObject()

    Set preservedNames = BuildPreservedSheetNames(msgTrads)
    dict.Prepare preservedNames, geoObj, msgTrads, "table"

    If Not dict.ColumnExists("list auto") Then
        AddListAuto
    End If

    dict.Translate dictTrads
    choices.Translate dictTrads
    choices.Sort
    exports.Translate dictTrads
    analysisObj.Translate dictTrads

    prepared = True

CleanExit:
    CloseWorkbookQuiet setupBook
    CloseWorkbookQuiet geoBook
    ExitBusyState busyStateApplied
    If prepared Then ResetPostImportCaches
    Exit Sub

CleanFail:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    CloseWorkbookQuiet setupBook
    CloseWorkbookQuiet geoBook
    ExitBusyState busyStateApplied

    If errNumber = 0 Then Exit Sub

    If errNumber >= ProjectError.ErrorNone And errNumber <= ProjectError.ElementNotFound Then
        Err.Raise errNumber, CLASSNAME, errDescription
    Else
        ThrowError ProjectError.ErrorUnexpectedState, errDescription
    End If
End Sub

Public Sub Export(ByVal toWkb As Workbook)
    Dim dict As ILLdictionary
    Dim choices As ILLChoices
    Dim translations As ILLTranslations
    Dim exports As ILLExport
    Dim geo As ILLGeo
    Dim passwords As ILLPasswords
    Dim mainObj As IMain

    On Error GoTo HandleError

    If toWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Destination workbook is required to export specifications"
    End If

    Set dict = Dictionary()
    Set choices = Choices()
    Set translations = Translations()
    Set exports = ExportObject()
    Set geo = GeoObject()
    Set passwords = Password()
    Set mainObj = MainObject()

    dict.Export toWkb, Hide:=xlSheetVeryHidden
    choices.Export toWkb, Hide:=xlSheetVeryHidden
    translations.ExportToWkb toWkb
    exports.ExportSpecs toWkb, Hide:=xlSheetVeryHidden
    geo.ExportToWkb toWkb, mainObj.ValueOf("llname"), mainObj.ValueOf("setuppath")
    passwords.Export toWkb

    Exit Sub

HandleError:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    If errNumber = 0 Then Exit Sub

    If errNumber >= ProjectError.ErrorNone And errNumber <= ProjectError.ElementNotFound Then
        Err.Raise errNumber, CLASSNAME, errDescription
    Else
        ThrowError ProjectError.ErrorUnexpectedState, errDescription
    End If
End Sub

Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
    Dim messageBody As String

    messageBody = "An error occurred while building the linelist specifications." & vbNewLine & _
                  "Details: " & textMessage & vbNewLine & _
                  "Aborting linelist creation process."

    MsgBox messageBody, vbExclamation + vbOKOnly, "Linelist Specifications"
End Sub

'@section Error handling
'===============================================================================
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Test support
'===============================================================================
Friend Sub TestAssignMain(ByVal mainObj As IMain)
    Set this.main = mainObj
End Sub

Friend Sub TestAssignDesignerTranslations(ByVal designerTrads As IDesTranslation)
    Set this.designerTranslations = designerTrads
End Sub

Friend Sub TestAssignTranslations(ByVal translations As ILLTranslations)
    Set this.translations = translations
    ResetTranslationCache
End Sub

'@section Interface implementation
'===============================================================================
Private Sub ILinelistSpecs_ResetCaches()
    ResetCaches
End Sub

Private Property Get ILinelistSpecs_Workbook() As Workbook
    Set ILinelistSpecs_Workbook = Me.Workbook
End Property

Private Property Get ILinelistSpecs_Dictionary() As ILLdictionary
    Set ILinelistSpecs_Dictionary = Me.Dictionary
End Property

Private Property Get ILinelistSpecs_Choices() As ILLChoices
    Set ILinelistSpecs_Choices = Me.Choices
End Property

Private Property Get ILinelistSpecs_DesignFormat() As ILLFormat
    Set ILinelistSpecs_DesignFormat = Me.DesignFormat
End Property

Private Property Get ILinelistSpecs_Analysis() As ILLAnalysis
    Set ILinelistSpecs_Analysis = Me.Analysis
End Property

Private Property Get ILinelistSpecs_ExportObject() As ILLExport
    Set ILinelistSpecs_ExportObject = Me.ExportObject
End Property

Private Property Get ILinelistSpecs_Password() As ILLPasswords
    Set ILinelistSpecs_Password = Me.Password
End Property

Private Property Get ILinelistSpecs_GeoObject() As ILLGeo
    Set ILinelistSpecs_GeoObject = Me.GeoObject
End Property

Private Property Get ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = Me.FormulaDataObject
End Property

Private Property Get ILinelistSpecs_MainObject() As IMain
    Set ILinelistSpecs_MainObject = Me.MainObject
End Property

Private Property Get ILinelistSpecs_Translations() As ILLTranslations
    Set ILinelistSpecs_Translations = Me.Translations
End Property

Private Property Get ILinelistSpecs_DesignerTranslations() As IDesTranslation
    Set ILinelistSpecs_DesignerTranslations = Me.DesignerTranslations
End Property

Private Property Get ILinelistSpecs_TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As ITranslationObject
    Set ILinelistSpecs_TransObject = Me.TransObject(translationScope)
End Property

Private Property Get ILinelistSpecs_TemporarySheetName(ByVal scope As Byte) As String
    ILinelistSpecs_TemporarySheetName = Me.TemporarySheetName(scope)
End Property

Private Property Get ILinelistSpecs_Value(ByVal tagName As String) As String
    ILinelistSpecs_Value = Me.Value(tagName)
End Property

Private Function ILinelistSpecs_Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Set ILinelistSpecs_Categories = Me.Categories(varName, useShortlabels)
End Function

Private Sub ILinelistSpecs_Prepare()
    Prepare
End Sub

Private Sub ILinelistSpecs_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub

Private Sub ILinelistSpecs_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    ErrorManage textMessage
End Sub

Private Function ILinelistSpecs_CreateSectionContext(ByVal linelist As ILinelist, ByVal startRow As Long) As ILLSectionContext
    Set ILinelistSpecs_CreateSectionContext = CreateSectionContext(linelist, startRow)
End Function

Private Function ILinelistSpecs_CreateSectionWriter() As ILLSectionWriter
    Set ILinelistSpecs_CreateSectionWriter = CreateSectionWriter()
End Function

Private Function ILinelistSpecs_CreateSubSectionWriter() As ILLSubSectionWriter
    Set ILinelistSpecs_CreateSubSectionWriter = CreateSubSectionWriter()
End Function

Private Function ILinelistSpecs_CreateVarContext(ByVal linelist As ILinelist, ByVal variableName As String) As ILLVarContext
    Set ILinelistSpecs_CreateVarContext = CreateVarContext(linelist, variableName)
End Function

Private Function ILinelistSpecs_CreateVarWriter() As ILLVarWriter
    Set ILinelistSpecs_CreateVarWriter = CreateVarWriter()
End Function

