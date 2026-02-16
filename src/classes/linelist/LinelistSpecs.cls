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

'@class LinelistSpecs
'@description
'Modern facade that wraps a specifications workbook and exposes lazy-loaded,
'cached accessors for every collaborator object (dictionary, choices,
'translations, geo, analysis, passwords, format, exports, formulas).
'Provides lifecycle orchestration methods for importing data from external
'setup workbooks, preparing the dictionary with preserved sheet names,
'exporting specifications to a destination linelist workbook, and surfacing
'user-facing error dialogs. Uses a PredeclaredId factory pattern so callers
'construct instances via LinelistSpecs.Create(workbook) and receive an
'ILinelistSpecs interface reference.
'@depends ILLdictionary, ILLChoices, ILLTranslation, IDesignerTranslation
'@depends ILLGeo, ILLExport, IAnalysis, IPasswords, ILLFormat
'@depends IDesignerImportService, ITemporaryRepos, TemporaryRepos
'@depends IFormulaData, ILLVariables, IFormulas
'@depends ITranslationObject, IApplicationState, ApplicationState
'@depends BetterArray, LLdictionary, LLChoices, LLTranslation
'@depends DesignerTranslation, LLGeo, LLExport, Analysis, Passwords
'@depends LLFormat, FormulaData, LLVariables, Formulas
'@version 1.1 (2026-02-16)



'@section State
'===============================================================================
'@description Backing UDT and module-level constants.

Private Type TLinelistSpecs
    specsWkb As Workbook
    dictionary As ILLdictionary
    choices As ILLChoices
    translations As ILLTranslation
    designerTranslations As IDesignerTranslation
    geo As ILLGeo
    exports As ILLExport
    analysis As IAnalysis
    passwords As IPasswords
    format As ILLFormat
    formulaData As IFormulaData
    variables As ILLVariables
    'Translation scope cache: direct fields for each TranslationScope value (0-4)
    transMessages As ITranslationObject
    transDictionary As ITranslationObject
    transShapes As ITranslationObject
    transForms As ITranslationObject
    transRibbon As ITranslationObject
    'Category cache: parallel arrays (same index = same entry)
    catKeys As BetterArray
    catValues As BetterArray
    outputWkb As Workbook
    tempRepos As ITemporaryRepos
End Type

Private this As TLinelistSpecs

Private Const CLASSNAME As String = "LinelistSpecs"
Private Const DEFAULT_DICT_START_ROW As Long = 1
Private Const DEFAULT_DICT_START_COLUMN As Long = 1
Private Const DEFAULT_CHOICES_START_ROW As Long = 1
Private Const DEFAULT_CHOICES_START_COLUMN As Long = 1

'Linelist Spec constants
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const SHEET_CHOICES As String = "Choices"
Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_PASSWORDS As String = "__pass"
Private Const SHEET_FORMULAS As String = "__formula"
Private Const SHEET_TRANSLATIONS_DICT As String = "Translations"
Private Const SHEET_TRANSLATIONS_LL As String = "LinelistTranslation"
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_EXPORT As String = "Exports"
Private Const SHEET_FORMAT As String = "__formatter"
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_DESIGNER_TRANSLATION As String = "DesignerTranslation"
Private Const RANGE_DESIGN_TYPE As String = "DESIGNTYPE"

'@section Factory helpers
'===============================================================================
'@description Construction and identity accessors.

'@label:create
'@sub-title Create a new specification facade from a workbook
'@details
'Validates the workbook structure upfront (checking for all required
'worksheets and named ranges), then constructs a new LinelistSpecs
'instance with caches initialised. Returns the ILinelistSpecs interface.
'@param specsWorkbook Workbook. The specifications workbook to wrap.
'@return ILinelistSpecs. A fully initialised facade ready for use.
'@throws ProjectError.ObjectNotInitialized When specsWorkbook is Nothing.
'@throws ProjectError.ElementNotFound When required worksheets are missing.
'@export
Public Function Create(ByVal specsWorkbook As Workbook) As ILinelistSpecs
    ValidateWorkbookReference specsWorkbook

    With New LinelistSpecs
        Set .InternalWorkbook = specsWorkbook
        .ResetCaches
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Interface-cast self reference
'@details
'Returns the current instance cast as ILinelistSpecs for fluent
'construction patterns within the factory method.
'@return ILinelistSpecs. This instance as its interface type.
Public Property Get Self() As ILinelistSpecs
    Set Self = Me
End Property

'@label:workbook
'@prop-title Primary workbook backing all specification queries
'@details
'Returns the specifications workbook reference. Raises an error if
'the workbook has not been assigned via the factory method.
'@return Workbook. The specifications workbook.
'@throws ProjectError.ObjectNotInitialized When the workbook is not set.
'@remarks
'The public property is named LLWorkbook to avoid colliding with the
'VBA Workbook keyword; the ILinelistSpecs interface exposes this as
'Workbook. The delegation stub ILinelistSpecs_Workbook calls Me.Workbook,
'which resolves through the interface dispatch mechanism.
'@export
Public Property Get LLWorkbook() As Workbook
    EnsureWorkbook
    Set LLWorkbook = this.specsWkb
End Property

'@label:internal-workbook
'@prop-title Setter for the backing workbook reference
'@details
'Used during construction to assign the workbook before caches are
'initialised. Not part of the public interface.
Public Property Set InternalWorkbook(ByVal specsWorkbook As Workbook)
    Set this.specsWkb = specsWorkbook
End Property

'@section State management
'===============================================================================
'@description Cache invalidation for collaborator objects, translations, and categories.

'@label:reset-caches
'@sub-title Reset every cached collaborator
'@details
'Clears all cached collaborator objects, the translation scope cache,
'and the category cache. Call after importing new data so subsequent
'reads query the workbook again.
'@export
Public Sub ResetCaches()
    ResetObjectCaches
    ResetTranslationCache
    ResetCategoryCache
End Sub

'@label:reset-object-caches
'@sub-title Clear all collaborator object references
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
    Set this.variables = Nothing
    Set this.outputWkb = Nothing
    Set this.tempRepos = Nothing
End Sub

'@label:reset-translation-cache
'@sub-title Clear the translation scope cache
Private Sub ResetTranslationCache()
    Set this.transMessages = Nothing
    Set this.transDictionary = Nothing
    Set this.transShapes = Nothing
    Set this.transForms = Nothing
    Set this.transRibbon = Nothing
End Sub

'@label:reset-category-cache
'@sub-title Clear the category list cache
Private Sub ResetCategoryCache()
    Set this.catKeys = Nothing
    Set this.catValues = Nothing
End Sub

'@section Guard helpers
'===============================================================================
'@description Lazy-initialisation guards for the workbook and category arrays.

'@label:ensure-workbook
'@sub-title Verify that the backing workbook is assigned
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
Private Sub EnsureWorkbook()
    If this.specsWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications workbook has not been assigned"
    End If
End Sub

'@label:ensure-category-arrays
'@sub-title Create the parallel category arrays if they do not exist
Private Sub EnsureCategoryArrays()
    If this.catKeys Is Nothing Then
        Set this.catKeys = New BetterArray
        this.catKeys.LowerBound = 1
        Set this.catValues = New BetterArray
        this.catValues.LowerBound = 1
    End If
End Sub

'@section Worksheet resolution
'===============================================================================
'@description Direct worksheet lookups by name within the specifications workbook.

'@label:resolve-sheet
'@sub-title Resolve a worksheet by name with optional error surfacing
'@details
'Looks up a worksheet in the specifications workbook by name. When the
'sheet is not found and allowMissing is False, raises a descriptive error.
'@param sheetName String. The worksheet name to find.
'@param allowMissing Boolean. When True, returns Nothing instead of raising.
'@return Worksheet. The resolved worksheet, or Nothing if allowMissing.
'@throws ProjectError.InvalidArgument When sheetName is empty and not allowMissing.
'@throws ProjectError.ElementNotFound When the sheet is missing and not allowMissing.
Private Function ResolveSheet(ByVal sheetName As String, _
                              Optional ByVal allowMissing As Boolean = False) As Worksheet

    Dim target As Worksheet
    Dim normalisedName As String

    normalisedName = Trim$(sheetName)
    If LenB(normalisedName) = 0 Then
        If allowMissing Then Exit Function
        ThrowError ProjectError.InvalidArgument, "Worksheet name cannot be empty"
    End If

    EnsureWorkbook

    On Error Resume Next
        Set target = this.specsWkb.Worksheets(normalisedName)
    On Error GoTo 0

    If target Is Nothing Then
        If allowMissing Then Exit Function
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & normalisedName & "' not found in specifications workbook '" & SafeWorkbookName(this.specsWkb) & "'"
    End If

    Set ResolveSheet = target
End Function

'@section External workbook helpers
'===============================================================================
'@description Utilities for opening, closing, and validating external workbooks
'during the Prepare workflow.

'@label:enter-busy-state
'@sub-title Transition the application into a busy state
'@details
'Creates or refreshes an ApplicationState snapshot and applies the busy
'state (suppressing screen updates and calculation) if not already busy.
'@param busyScope IApplicationState (ByRef). The state object to manage.
Private Sub EnterBusyState(ByRef busyScope As IApplicationState)
    If busyScope Is Nothing Then
        Set busyScope = ApplicationState.Create(Application)
    ElseIf Not busyScope.HasSnapshot Then
        busyScope.RefreshSnapshot
    End If

    If Not busyScope.IsBusy Then
        busyScope.ApplyBusyState _
            suppressEvents:=False, _
            calculateOnSave:=Application.CalculateBeforeSave, _
            busyCursor:=Application.Cursor
    End If
End Sub

'@label:exit-busy-state
'@sub-title Restore the application from a busy state
'@param busyScope IApplicationState (ByRef). The state object to restore and release.
Private Sub ExitBusyState(ByRef busyScope As IApplicationState)
    If busyScope Is Nothing Then Exit Sub

    On Error Resume Next
        busyScope.Restore
    On Error GoTo 0
    Set busyScope = Nothing
End Sub

'@label:build-preserved-sheet-names
'@sub-title Build the list of sheet names to preserve during dictionary preparation
'@details
'Collects all specification sheet names and temporary sheet names into
'a BetterArray, then appends translated linelist sheet names (Admin,
'Analysis, Spatial, etc.) when message translations are available.
'@param messageTranslations ITranslationObject. The message translations
'for resolving linelist sheet names, or Nothing to skip translated names.
'@return BetterArray. The list of preserved sheet names.
Private Function BuildPreservedSheetNames(ByVal messageTranslations As ITranslationObject) As BetterArray
    Dim preserved As BetterArray

    Set preserved = New BetterArray
    preserved.LowerBound = 1

    preserved.Push SHEET_DICTIONARY, _
                   SHEET_CHOICES, _
                   SHEET_GEO, _
                   SHEET_PASSWORDS, _
                   SHEET_FORMULAS, _
                   SHEET_TRANSLATIONS_DICT, _
                   SHEET_TRANSLATIONS_LL, _
                   SHEET_EXPORT, _
                   TemporarySheetName(TempSheetSingle), _
                   TemporarySheetName(TempSheetList), _
                   TemporarySheetName(TempSheetUpdates), _
                   TemporarySheetName(TempSheetMetadata), _
                   TemporarySheetName(TempSheetAnalysis), _
                   TemporarySheetName(TempSheetImport), _
                   TemporarySheetName(TempSheetSpatial), _
                   TemporarySheetName(TempSheetShowHide)

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

'@label:safe-translated-value
'@sub-title Look up a translation key, falling back to the key itself on failure
'@param translations ITranslationObject. The translation object to query.
'@param key String. The translation key.
'@return String. The translated value, or the key if translation fails.
Private Function SafeTranslatedValue(ByVal translations As ITranslationObject, ByVal key As String) As String
    On Error Resume Next
        SafeTranslatedValue = translations.TranslatedValue(key)
    On Error GoTo 0

    If LenB(SafeTranslatedValue) = 0 Then SafeTranslatedValue = key
End Function

'@label:reset-post-import-caches
'@sub-title Clear caches that become stale after importing new data
'@details
'Resets the variables object, category cache, and translation cache
'so they are rebuilt from the freshly imported data.
Private Sub ResetPostImportCaches()
    Set this.variables = Nothing
    ResetCategoryCache
    ResetTranslationCache
End Sub

'@section Workbook validation
'===============================================================================
'@description Structural validation of specifications workbooks.

'@label:validate-workbook-reference
'@sub-title Validate a workbook reference and its structure
'@throws ProjectError.ObjectNotInitialized When specsWorkbook is Nothing.
'@throws ProjectError.ElementNotFound When required sheets are missing.
Private Sub ValidateWorkbookReference(ByVal specsWorkbook As Workbook)
    If specsWorkbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "A specifications workbook reference is required"
    End If

    ValidateWorkbookStructure specsWorkbook
End Sub

'@label:validate-workbook-structure
'@sub-title Check that a workbook contains all required worksheets and named ranges
'@details
'Iterates the list of required sheet names and collects any that are
'missing. If the format sheet exists, also validates that the DESIGNTYPE
'named range is present.
'@param specsWorkbook Workbook. The workbook to validate.
'@throws ProjectError.ElementNotFound When required sheets or ranges are missing.
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
        Set formatSheet = specsWorkbook.Worksheets(SHEET_FORMAT)
    On Error GoTo 0

    If formatSheet Is Nothing Then Exit Sub

    If Not RangeExistsInternal(formatSheet, RANGE_DESIGN_TYPE) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & SHEET_FORMAT & "' must define the '" & RANGE_DESIGN_TYPE & "' named range"
    End If
End Sub

'@label:required-sheet-names
'@sub-title Return the array of required specification worksheet names
'@return Variant. Array of sheet name strings.
Private Function RequiredSheetNames() As Variant
    RequiredSheetNames = Array( _
        SHEET_DICTIONARY, _
        SHEET_CHOICES, _
        SHEET_GEO, _
        SHEET_PASSWORDS, _
        SHEET_FORMULAS, _
        SHEET_TRANSLATIONS_LL, _
        SHEET_ANALYSIS, _
        SHEET_EXPORT, _
        SHEET_FORMAT, _
        SHEET_MAIN, _
        SHEET_DESIGNER_TRANSLATION)
End Function

'@label:worksheet-exists-internal
'@sub-title Test whether a worksheet exists in a workbook
'@param specsWorkbook Workbook. The workbook to search.
'@param sheetName String. The worksheet name to look for.
'@return Boolean. True when the worksheet exists.
Private Function WorksheetExistsInternal(ByVal specsWorkbook As Workbook, _
                                         ByVal sheetName As String) As Boolean
    Dim target As Worksheet
    On Error Resume Next
        Set target = specsWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExistsInternal = Not (target Is Nothing)
End Function

'@label:range-exists-internal
'@sub-title Test whether a named range exists on a worksheet
'@param hostSheet Worksheet. The worksheet to search.
'@param rangeName String. The named range to look for.
'@return Boolean. True when the named range exists.
Private Function RangeExistsInternal(ByVal hostSheet As Worksheet, _
                                     ByVal rangeName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
        Set rng = hostSheet.Range(rangeName)
    On Error GoTo 0
    RangeExistsInternal = Not (rng Is Nothing)
End Function

'@label:join-names
'@sub-title Join a BetterArray of names into a comma-separated, quoted string
'@param values BetterArray. The names to join.
'@return String. Formatted string like "'Sheet1', 'Sheet2'".
Private Function JoinNames(ByVal values As BetterArray) As String
    Dim buffer As String
    Dim idx As Long

    For idx = values.LowerBound To values.UpperBound
        If LenB(buffer) > 0 Then buffer = buffer & ", "
        buffer = buffer & "'" & CStr(values.Item(idx)) & "'"
    Next idx

    JoinNames = buffer
End Function

'@label:safe-workbook-name
'@sub-title Safely retrieve a workbook's name, returning "<unknown>" on error
'@param specsWorkbook Workbook. The workbook whose name to retrieve.
'@return String. The workbook name or "<unknown>".
Private Function SafeWorkbookName(ByVal specsWorkbook As Workbook) As String
    On Error Resume Next
        SafeWorkbookName = specsWorkbook.Name
    On Error GoTo 0
    If LenB(SafeWorkbookName) = 0 Then SafeWorkbookName = "<unknown>"
End Function

'@section Collaborator ensures
'===============================================================================
'@description Lazy-initialisation guards for each collaborator object. Each
'EnsureXxx method creates the collaborator from its backing worksheet on first
'access, then caches it for subsequent calls.

'@label:ensure-exports
'@sub-title Lazy-load the export collaborator
Private Sub EnsureExports()
    If this.exports Is Nothing Then
        Dim exportSheet As Worksheet
        Set exportSheet = ResolveSheet(SHEET_EXPORT)
        Set this.exports = LLExport.Create(exportSheet)
    End If
End Sub

'@label:ensure-dictionary
'@sub-title Lazy-load the dictionary collaborator
'@details
'Requires the export collaborator to be available first (for the number
'of exports parameter), then creates the dictionary from the Dictionary
'worksheet.
Private Sub EnsureDictionary()
    If this.dictionary Is Nothing Then
        EnsureExports
        Dim dictSheet As Worksheet
        Set dictSheet = ResolveSheet(SHEET_DICTIONARY)
        Set this.dictionary = LLdictionary.Create(dictSheet, _
                                                  DEFAULT_DICT_START_ROW, _
                                                  DEFAULT_DICT_START_COLUMN, _
                                                  NumberOfExports())
    End If
End Sub

'@label:ensure-choices
'@sub-title Lazy-load the choices collaborator
Private Sub EnsureChoices()
    If this.choices Is Nothing Then
        Dim choicesSheet As Worksheet
        Set choicesSheet = ResolveSheet(SHEET_CHOICES)
        Set this.choices = LLChoices.Create(choicesSheet, _
                                            DEFAULT_CHOICES_START_ROW, _
                                            DEFAULT_CHOICES_START_COLUMN)
    End If
End Sub

'@label:ensure-translations
'@sub-title Lazy-load the translations collaborator
'@details
'Creates the LLTranslation instance from the linelist translation worksheet.
Private Sub EnsureTranslations()
    If this.translations Is Nothing Then
        Dim linelistTransSheet As Worksheet
        Set linelistTransSheet = ResolveSheet(SHEET_TRANSLATIONS_LL)
        Set this.translations = LLTranslation.Create(linelistTransSheet)
    End If
End Sub

'@label:ensure-designer-translations
'@sub-title Lazy-load the designer translations collaborator
Private Sub EnsureDesignerTranslations()
    If this.designerTranslations Is Nothing Then
        Dim designerSheet As Worksheet
        Set designerSheet = ResolveSheet(SHEET_DESIGNER_TRANSLATION)
        Set this.designerTranslations = DesignerTranslation.Create(designerSheet)
    End If
End Sub

'@label:ensure-geo
'@sub-title Lazy-load the geo collaborator
Private Sub EnsureGeo()
    If this.geo Is Nothing Then
        Dim geoSheet As Worksheet
        Set geoSheet = ResolveSheet(SHEET_GEO)
        Set this.geo = LLGeo.Create(geoSheet)
    End If
End Sub

'@label:ensure-analysis
'@sub-title Lazy-load the analysis collaborator
Private Sub EnsureAnalysis()
    If this.analysis Is Nothing Then
        Dim analysisSheet As Worksheet
        Set analysisSheet = ResolveSheet(SHEET_ANALYSIS)
        Set this.analysis = Analysis.Create(analysisSheet)
    End If
End Sub

'@label:ensure-passwords
'@sub-title Lazy-load the passwords collaborator
Private Sub EnsurePasswords()
    If this.passwords Is Nothing Then
        Dim passwordSheet As Worksheet
        Set passwordSheet = ResolveSheet(SHEET_PASSWORDS)
        Set this.passwords = Passwords.Create(passwordSheet)
    End If
End Sub

'@label:ensure-formula-data
'@sub-title Lazy-load the formula data collaborator
Private Sub EnsureFormulaData()
    If this.formulaData Is Nothing Then
        Dim formulaSheet As Worksheet
        Set formulaSheet = ResolveSheet(SHEET_FORMULAS)
        Set this.formulaData = FormulaData.Create(formulaSheet)
    End If
End Sub

'@label:ensure-variables
'@sub-title Lazy-load the variables collaborator
'@details
'Requires the dictionary to be available first, then creates the
'LLVariables instance from it.
Private Sub EnsureVariables()
    EnsureDictionary
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(this.dictionary)
    End If
End Sub

'@label:ensure-format
'@sub-title Lazy-load the format collaborator
'@details
'Resolves the design name from the DESIGNTYPE named range on the format
'sheet, then creates the LLFormat instance.
Private Sub EnsureFormat()
    If this.format Is Nothing Then
        Dim formatSheet As Worksheet
        Dim designName As String

        Set formatSheet = ResolveSheet(SHEET_FORMAT)
        designName = ResolveDesignName(formatSheet)

        Set this.format = LLFormat.Create(formatSheet, designName)
    End If
End Sub

'@label:resolve-design-name
'@sub-title Extract the design name from the format worksheet's DESIGNTYPE range
'@param formatSheet Worksheet. The format worksheet.
'@return String. The trimmed design type name.
'@throws ProjectError.ElementNotFound When the DESIGNTYPE range is missing.
Private Function ResolveDesignName(ByVal formatSheet As Worksheet) As String
    Dim designRange As Range

    On Error Resume Next
        Set designRange = formatSheet.Range(RANGE_DESIGN_TYPE)
    On Error GoTo 0

    If designRange Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & RANGE_DESIGN_TYPE & "' missing from format worksheet"
    End If

    ResolveDesignName = Trim$(CStr(designRange.Cells(1, 1).Value))
End Function

'@label:ensure-category-lookup
'@sub-title Ensure choices, variables, and category arrays are ready
Private Sub EnsureCategoryLookup()
    EnsureChoices
    EnsureVariables
    EnsureCategoryArrays
End Sub

'@label:number-of-exports
'@sub-title Return the number of exports defined in the export specifications
'@return Long. The export count.
Private Function NumberOfExports() As Long
    EnsureExports
    NumberOfExports = this.exports.NumberOfExports()
End Function

'@section List auto helpers
'===============================================================================
'@description Automatic list column generation for formula-based variables.

'@label:add-list-auto
'@sub-title Populate the "list auto" dictionary column for list_auto variables
'@details
'Adds a "list auto" column to the dictionary, then finds all variables
'with control type "list_auto" and marks their control details with
'"list_auto_origin". For formula-type variables (formula, case_when,
'choice_formula), recursively traces referenced variables and marks
'them as well.
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

'@label:recursive-list-auto
'@sub-title Recursively mark referenced variables in formula chains as list_auto
'@details
'Parses the control details of a formula-type variable using the Formulas
'engine, extracts the list of referenced variables, marks each with
'"list_auto_origin", and recurses into any that are themselves formulas
'or case_when types.
'@param varName String. The formula variable whose references to trace.
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
'@description Category resolution and caching for choice-type variables.

'@label:categories
'@sub-title Resolve category values for a variable with caching
'@details
'Returns the list of category values for a choice-type variable
'(choice_manual, choice_multiple, choice_custom, or choice_formula).
'Results are cached using parallel key/value BetterArrays for efficient
'repeated access. Always returns a cloned BetterArray to prevent callers
'from modifying the cached copy.
'@param varName String. The variable name whose categories to resolve.
'@param useShortlabels Boolean. When True, returns short labels. Defaults to False.
'@return BetterArray. Cloned list of category values.
'@throws ProjectError.InvalidArgument When varName is empty.
'@throws ProjectError.ElementNotFound When the variable has no selectable categories.
'@export
Public Function Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
    Dim cacheKey As String
    Dim cached As BetterArray
    Dim cacheIdx As Long

    EnsureCategoryLookup

    cacheKey = LCase$(Trim$(varName)) & "|" & CStr(useShortlabels)

    'Look up in parallel arrays
    cacheIdx = FindCategoryIndex(cacheKey)
    If cacheIdx > 0 Then
        Set cached = this.catValues.Item(cacheIdx)
    End If

    If cached Is Nothing Then
        Set cached = ResolveCategoryList(varName, useShortlabels)
        this.catKeys.Push cacheKey
        this.catValues.Push cached
    End If

    Set Categories = cached.Clone
End Function

'@label:find-category-index
'@sub-title Search the category keys array for a matching key
'@param cacheKey String. The key to search for.
'@return Long. The 1-based index if found, or 0 if not found.
Private Function FindCategoryIndex(ByVal cacheKey As String) As Long
    Dim idx As Long

    If this.catKeys Is Nothing Then Exit Function
    If this.catKeys.Length = 0 Then Exit Function

    For idx = this.catKeys.LowerBound To this.catKeys.UpperBound
        If StrComp(CStr(this.catKeys.Item(idx)), cacheKey, vbBinaryCompare) = 0 Then
            FindCategoryIndex = idx
            Exit Function
        End If
    Next idx
End Function

'@label:resolve-category-list
'@sub-title Build a category list from the choices sheet for a given variable
'@details
'Determines the category name from the variable's control type and
'control details, then queries the choices collaborator.
'@param varName String. The variable name.
'@param useShortlabels Boolean. Whether to use short labels.
'@return BetterArray. The resolved category list.
'@throws ProjectError.InvalidArgument When varName is empty.
'@throws ProjectError.ElementNotFound When the variable or category is not found.
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

'@label:normalise-control-type
'@sub-title Strip parenthesised suffixes from a control type string
'@param controlValue String. The raw control type value.
'@return String. The normalised, lower-cased control type.
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

'@label:extract-choice-formula
'@sub-title Extract the choice name from a CHOICE_FORMULA(...) control details string
'@param controlDetails String. The raw control details value.
'@return String. The extracted choice name.
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
'@description Public property accessors for each collaborator object, exposing
'the lazy-loaded instances through their respective interfaces.

'@label:dictionary
'@prop-title Linelist dictionary
'@details
'Returns the ILLdictionary instance providing column and row access to the
'Dictionary worksheet. Lazy-loaded on first access.
'@return ILLdictionary. The dictionary collaborator.
'@export
Public Property Get Dictionary() As ILLdictionary
    EnsureDictionary
    Set Dictionary = this.dictionary
End Property

'@label:choices
'@prop-title Choices catalogue
'@details
'Returns the ILLChoices instance providing category lookups from the
'Choices worksheet. Lazy-loaded on first access.
'@return ILLChoices. The choices collaborator.
'@export
Public Property Get Choices() As ILLChoices
    EnsureChoices
    Set Choices = this.choices
End Property

'@label:design-format
'@prop-title Visual formatting specifications
'@details
'Returns the ILLFormat instance providing sheet and cell formatting
'methods based on the __formatter worksheet. Lazy-loaded on first access.
'@return ILLFormat. The format collaborator.
'@export
Public Property Get DesignFormat() As ILLFormat
    EnsureFormat
    Set DesignFormat = this.format
End Property

'@label:analysis
'@prop-title Analysis setup specifications
'@details
'Returns the IAnalysis instance providing access to the Analysis
'worksheet configuration. Lazy-loaded on first access.
'@return IAnalysis. The analysis collaborator.
'@export
Public Property Get Analysis() As IAnalysis
    EnsureAnalysis
    Set Analysis = this.analysis
End Property

'@label:export-object
'@prop-title Export column specifications
'@details
'Returns the ILLExport instance providing access to export column
'definitions from the Exports worksheet. Lazy-loaded on first access.
'@return ILLExport. The export collaborator.
'@export
Public Property Get ExportObject() As ILLExport
    EnsureExports
    Set ExportObject = this.exports
End Property

'@label:password
'@prop-title Password manager
'@details
'Returns the IPasswords instance providing workbook and sheet protection
'utilities from the __pass worksheet. Lazy-loaded on first access.
'@return IPasswords. The password collaborator.
'@export
Public Property Get Password() As IPasswords
    EnsurePasswords
    Set Password = this.passwords
End Property

'@label:geo-object
'@prop-title Geographic data manager
'@details
'Returns the ILLGeo instance providing geographic hierarchy lookups from
'the Geo worksheet. Lazy-loaded on first access.
'@return ILLGeo. The geo collaborator.
'@export
Public Property Get GeoObject() As ILLGeo
    EnsureGeo
    Set GeoObject = this.geo
End Property

'@label:formula-data-object
'@prop-title Formula data provider
'@details
'Returns the IFormulaData instance providing formula parsing support from
'the __formula worksheet. Lazy-loaded on first access.
'@return IFormulaData. The formula data collaborator.
'@export
Public Property Get FormulaDataObject() As IFormulaData
    EnsureFormulaData
    Set FormulaDataObject = this.formulaData
End Property

'@label:has-template
'@prop-title Whether the workbook uses a ribbon template
'@details
'Reads the RNG_MainTemp named range on the Main worksheet and returns True
'when the value equals "has template".
'@return Boolean. True when a template is configured.
'@export
Public Property Get HasTemplate() As Boolean
    Dim mainSheet As Worksheet
    Set mainSheet = ResolveSheet(SHEET_MAIN)
    On Error Resume Next
    HasTemplate = (mainSheet.Range("RNG_MainTemp").Value = "has template")
    On Error GoTo 0
End Property

'@label:translations
'@prop-title Translation tables manager
'@details
'Returns the ILLTranslation instance providing access to the linelist
'translation worksheet. Lazy-loaded on first access.
'@return ILLTranslation. The translations collaborator.
'@export
Public Property Get Translations() As ILLTranslation
    EnsureTranslations
    Set Translations = this.translations
End Property

'@label:designer-translations
'@prop-title Designer translation accessor
'@details
'Returns the IDesignerTranslation instance providing designer UI translations
'from the DesignerTranslation worksheet. Lazy-loaded on first access.
'@return IDesignerTranslation. The designer translation collaborator.
'@export
Public Property Get DesignerTranslations() As IDesignerTranslation
    EnsureDesignerTranslations
    Set DesignerTranslations = this.designerTranslations
End Property

'@label:trans-object
'@prop-title Scoped translation object
'@details
'Returns an ITranslationObject for the requested translation scope.
'Results are cached per scope using direct fields for each TranslationScope
'value (0-4), avoiding generic container overhead.
'@param translationScope Byte. The translation scope constant
'(e.g., TranslationOfMessages, TranslationOfDictionary). Defaults to
'TranslationOfMessages.
'@return ITranslationObject. The scoped translation object.
'@throws ProjectError.ElementNotFound When the scope is not available.
'@export
Public Property Get TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As ITranslationObject
    Dim resolved As ITranslationObject

    EnsureTranslations

    Select Case translationScope
        Case TranslationOfMessages
            If this.transMessages Is Nothing Then
                Set this.transMessages = ResolveTranslationScope(translationScope)
            End If
            Set TransObject = this.transMessages

        Case TranslationOfDictionary
            If this.transDictionary Is Nothing Then
                Set this.transDictionary = ResolveTranslationScope(translationScope)
            End If
            Set TransObject = this.transDictionary

        Case TranslationOfShapes
            If this.transShapes Is Nothing Then
                Set this.transShapes = ResolveTranslationScope(translationScope)
            End If
            Set TransObject = this.transShapes

        Case TranslationOfForms
            If this.transForms Is Nothing Then
                Set this.transForms = ResolveTranslationScope(translationScope)
            End If
            Set TransObject = this.transForms

        Case TranslationOfRibbon
            If this.transRibbon Is Nothing Then
                Set this.transRibbon = ResolveTranslationScope(translationScope)
            End If
            Set TransObject = this.transRibbon

        Case Else
            ThrowError ProjectError.ElementNotFound, _
                      "Translation scope '" & CStr(translationScope) & "' is not available in the specifications workbook"
    End Select
End Property

'@label:resolve-translation-scope
'@sub-title Query the translations collaborator for a given scope
'@param translationScope Byte. The translation scope to resolve.
'@return ITranslationObject. The resolved translation object.
'@throws ProjectError.ElementNotFound When the scope returns Nothing.
Private Function ResolveTranslationScope(ByVal translationScope As Byte) As ITranslationObject
    Dim resolved As ITranslationObject

    Set resolved = this.translations.TransObject(translationScope)
    If resolved Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Translation scope '" & CStr(translationScope) & "' is not available in the specifications workbook"
    End If

    Set ResolveTranslationScope = resolved
End Function

'@label:temporary-sheet-name
'@prop-title Temporary sheet name for a given scope
'@details
'Returns the fixed worksheet name used for temporary internal sheets,
'mapped from the TempSheetScope enum (TempSheetAnalysis, TempSheetList,
'etc.) to its canonical name string.
'@param scope Byte. A TempSheetScope enum value.
'@return String. The worksheet name string.
'@throws ProjectError.InvalidArgument When the scope is not recognised.
'@export
Public Property Get TemporarySheetName(ByVal scope As Byte) As String
    Select Case scope
        Case TempSheetAnalysis
            TemporarySheetName = "ana_tabnames__"
        Case TempSheetUpdates
            TemporarySheetName = "updates__"
        Case TempSheetMetadata
            TemporarySheetName = "Metadata"
        Case TempSheetList
            TemporarySheetName = "dropdown_lists__"
        Case TempSheetImport
            TemporarySheetName = "import_rep__"
        Case TempSheetSingle
            TemporarySheetName = "temp__"
        Case TempSheetSpatial
            TemporarySheetName = "spatial_tables__"
        Case TempSheetShowHide
            TemporarySheetName = "show_hide__"
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                      "Unknown temporary sheet scope '" & CStr(scope) & "'"
    End Select
End Property

'@label:value
'@prop-title Configuration value by tag name
'@details
'Returns a configuration value from the Main worksheet for the given
'tag name. The special tag "numberofexports" returns the export count
'from the Exports worksheet; all other tags resolve the named range directly
'on the Main worksheet.
'@param tagName String. The configuration key to look up.
'@return String. The configuration value, or empty string if tagName is empty.
'@export
Public Property Get Value(ByVal tagName As String) As String
    Dim normalizedTag As String
    Dim rngName As String
    Dim mainSheet As Worksheet
    Dim retVal As String

    normalizedTag = LCase$(Trim$(tagName))
    If LenB(normalizedTag) = 0 Then Exit Property

    Select Case normalizedTag
        Case "numberofexports"
            Value = CStr(NumberOfExports())
            Exit Property
        Case "llpassword"
            rngName = "RNG_LLPwdOpen"
        Case "geopath"
            rngName = "RNG_PathGeo"
        Case "setuppath"
            rngName = "RNG_PathDico"
        Case "lldir"
            rngName = "RNG_LLDir"
        Case "llname"
            rngName = "RNG_LLName"
        Case "setuplang"
            rngName = "RNG_LangSetup"
        Case "lllang"
            rngName = "RNG_LLForm"
        Case "askopen"
            rngName = "RNG_AskOpen"
        Case "instruct"
            rngName = "RNG_MainInstruct"
        Case "epiweekstart"
            Set mainSheet = ResolveSheet(SHEET_MAIN)
            On Error Resume Next
            Value = "1"
            Value = Split(mainSheet.Range("RNG_DefaultEpiWeek").Value, "-")(0)
            On Error GoTo 0
            Exit Property
        Case "temppath"
            Set mainSheet = ResolveSheet(SHEET_MAIN)
            On Error Resume Next
            retVal = mainSheet.Range("RNG_LLTemp").Value
            On Error GoTo 0
            If retVal = "_default" Or retVal = "default" Then
                retVal = this.specsWkb.Path & Application.PathSeparator & _
                         "_ribbontemplate.xlsb"
            End If
            Value = retVal
            Exit Property
        Case Else
            Exit Property
    End Select

    Set mainSheet = ResolveSheet(SHEET_MAIN)
    On Error Resume Next
    Value = mainSheet.Range(rngName).Value
    On Error GoTo 0
End Property


'@section Workbook orchestration
'===============================================================================
'@description Lifecycle methods for importing, preparing, exporting, and
'error-managing the specifications workbook.

'@label:prepare
'@sub-title Prepare all specifications using an import service
'@details
'Orchestrates the full specification preparation workflow:
'1. Creates a temporary repository for working files.
'2. Creates a new output workbook (from template if configured).
'3. Exports designer components via the import service.
'4. Retrieves exported domain instances from the service.
'5. Prepares the dictionary with preserved sheet names and geo references.
'6. Adds the "list auto" column if not already present.
'7. Translates and sorts all collaborators using dictionary translations.
'Handles errors with cleanup and busy-state restoration.
'@param importService IDesignerImportService. The configured import service.
'@throws ProjectError.ObjectNotInitialized When importService is Nothing.
'@throws ProjectError.ErrorUnexpectedState When an unexpected error occurs.
'@export
Public Sub Prepare(ByVal importService As IDesignerImportService)
    Dim dict As ILLdictionary
    Dim choicesObj As ILLChoices
    Dim exportsObj As ILLExport
    Dim analysisObj As IAnalysis
    Dim geoObj As ILLGeo
    Dim dictTrads As ITranslationObject
    Dim msgTrads As ITranslationObject
    Dim preservedNames As BetterArray
    Dim busyScope As IApplicationState
    Dim prepared As Boolean
    Dim outWkb As Workbook
    Dim templatePath As String
    Dim tempPath As String

    If importService Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "An import service is required to prepare specifications"
    End If

    On Error GoTo CleanFail

    'Readiness is validated by the caller (e.g. clickGenerate) before Prepare
    EnterBusyState busyScope

    'Create the temporary repository for working files
    Set this.tempRepos = TemporaryRepos.Create(Value("lldir"))
    this.tempRepos.EnsureReady

    'Create the output workbook (from template when configured)
    If HasTemplate() Then
        templatePath = Trim$(Value("temppath"))
        Set outWkb = Workbooks.Open(templatePath)
        tempPath = this.tempRepos.CreateFilePath("__temp.xlsb")
        outWkb.SaveAs fileName:=tempPath, fileFormat:=xlExcel12
    Else
        Set outWkb = Workbooks.Add
    End If

    'Export all designer components into the output workbook
    importService.ExportToLinelist outWkb

    'Retrieve exported domain instances from the service
    Set this.dictionary = importService.ExportedDictionary
    Set this.choices = importService.ExportedChoices
    Set this.exports = importService.ExportedExport
    Set this.analysis = importService.ExportedAnalysis
    Set this.geo = importService.ExportedGeo
    Set this.passwords = importService.ExportedPasswords
    Set this.translations = importService.ExportedTranslations

    Set this.outputWkb = outWkb

    Set dict = this.dictionary
    Set choicesObj = this.choices
    Set exportsObj = this.exports
    Set analysisObj = this.analysis
    Set geoObj = this.geo

    Set dictTrads = TransObject(TranslationOfDictionary)
    Set msgTrads = TransObject()

    Set preservedNames = BuildPreservedSheetNames(msgTrads)
    dict.Prepare preservedNames, geoObj, "table"

    If Not dict.ColumnExists("list auto") Then
        AddListAuto
    End If

    dict.Translate dictTrads
    choicesObj.Translate dictTrads
    choicesObj.Sort
    exportsObj.Translate dictTrads
    analysisObj.Translate dictTrads

    prepared = True

CleanExit:
    ExitBusyState busyScope
    If prepared Then ResetPostImportCaches
    Exit Sub

CleanFail:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.Description

    ExitBusyState busyScope

    If errNumber = 0 Then Exit Sub

    If errNumber >= ProjectError.ErrorNone And errNumber <= ProjectError.ElementNotFound Then
        Err.Raise errNumber, CLASSNAME, errDescription
    Else
        ThrowError ProjectError.ErrorUnexpectedState, errDescription
    End If
End Sub

'@label:output-workbook
'@prop-title The output linelist workbook created during Prepare
'@details
'Returns the workbook created by Prepare that received exported data from
'the import service. Returns Nothing before Prepare is called.
'@return Workbook. The output workbook, or Nothing.
'@export
Public Property Get OutputWorkbook() As Workbook
    Set OutputWorkbook = this.outputWkb
End Property

'@label:temporary-repos
'@prop-title The temporary repository created during Prepare
'@details
'Returns the ITemporaryRepos instance created by Prepare for working files.
'Returns Nothing before Prepare is called.
'@return ITemporaryRepos. The temporary repository, or Nothing.
'@export
Public Property Get OutputTemporaryRepos() As ITemporaryRepos
    Set OutputTemporaryRepos = this.tempRepos
End Property

'@label:export
'@sub-title Export all specifications to an output workbook
'@details
'Exports dictionary, choices, translations, export specs, geo data, and
'passwords to the given workbook, creating hidden worksheets for each.
'@param toWkb Workbook. The destination workbook to export specifications into.
'@throws ProjectError.ObjectNotInitialized When toWkb is Nothing.
'@throws ProjectError.ErrorUnexpectedState When an unexpected error occurs.
'@export
Public Sub Export(ByVal toWkb As Workbook)
    Dim dict As ILLdictionary
    Dim choices As ILLChoices
    Dim translations As ILLTranslation
    Dim exports As ILLExport
    Dim geo As ILLGeo
    Dim passwords As IPasswords

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

    dict.Export toWkb, Hide:=xlSheetVeryHidden
    choices.Export toWkb, Hide:=xlSheetVeryHidden
    translations.Export toWkb
    exports.ExportSpecs toWkb, Hide:=xlSheetVeryHidden
    geo.ExportToWkb toWkb, Value("llname"), Value("setuppath")
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

'@label:error-manage
'@sub-title Display an error dialog for specification failures
'@details
'Shows a message box describing the error and informing the user that
'the linelist creation process is being aborted.
'@param textMessage String. Error description. Defaults to vbNullString.
'@export
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
    Dim messageBody As String

    messageBody = "An error occurred while building the linelist specifications." & vbNewLine & _
                  "Details: " & textMessage & vbNewLine & _
                  "Aborting linelist creation process."

    MsgBox messageBody, vbExclamation + vbOKOnly, "Linelist Specifications"
End Sub

'@section Error handling
'===============================================================================
'@description Internal error-raising utility.

'@label:throw-error
'@sub-title Raise a project error with the class name as source
'@param errNumber Long. The ProjectError constant.
'@param messageText String. The error description.
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Test support
'===============================================================================
'@description Backdoor setters for injecting mock collaborators during testing.

'@label:test-assign-designer-translations
'@sub-title Inject a mock IDesignerTranslation instance for testing
'@param designerTrads IDesignerTranslation. The mock designer translations object.
Public Sub TestAssignDesignerTranslations(ByVal designerTrads As IDesignerTranslation)
    Set this.designerTranslations = designerTrads
End Sub

'@label:test-assign-translations
'@sub-title Inject a mock ILLTranslation instance for testing
'@details
'Also resets the translation cache to ensure the mock object is used
'for subsequent TransObject lookups.
'@param translations ILLTranslation. The mock translations object.
Public Sub TestAssignTranslations(ByVal translations As ILLTranslation)
    Set this.translations = translations
    ResetTranslationCache
End Sub

'@section Interface implementation
'===============================================================================
'@description ILinelistSpecs delegation stubs. Each member forwards to its
'corresponding public member on the concrete class. No @label tags are
'needed on delegation stubs.

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

Private Property Get ILinelistSpecs_Analysis() As IAnalysis
    Set ILinelistSpecs_Analysis = Me.Analysis
End Property

Private Property Get ILinelistSpecs_ExportObject() As ILLExport
    Set ILinelistSpecs_ExportObject = Me.ExportObject
End Property

Private Property Get ILinelistSpecs_Password() As IPasswords
    Set ILinelistSpecs_Password = Me.Password
End Property

Private Property Get ILinelistSpecs_GeoObject() As ILLGeo
    Set ILinelistSpecs_GeoObject = Me.GeoObject
End Property

Private Property Get ILinelistSpecs_FormulaDataObject() As IFormulaData
    Set ILinelistSpecs_FormulaDataObject = Me.FormulaDataObject
End Property

Private Property Get ILinelistSpecs_HasTemplate() As Boolean
    ILinelistSpecs_HasTemplate = Me.HasTemplate
End Property

Private Property Get ILinelistSpecs_Translations() As ILLTranslation
    Set ILinelistSpecs_Translations = Me.Translations
End Property

Private Property Get ILinelistSpecs_DesignerTranslations() As IDesignerTranslation
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

Private Sub ILinelistSpecs_Prepare(ByVal importService As IDesignerImportService)
    Prepare importService
End Sub

Private Property Get ILinelistSpecs_OutputWorkbook() As Workbook
    Set ILinelistSpecs_OutputWorkbook = Me.OutputWorkbook
End Property

Private Property Get ILinelistSpecs_OutputTemporaryRepos() As ITemporaryRepos
    Set ILinelistSpecs_OutputTemporaryRepos = Me.OutputTemporaryRepos
End Property

Private Sub ILinelistSpecs_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub

Private Sub ILinelistSpecs_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    ErrorManage textMessage
End Sub
