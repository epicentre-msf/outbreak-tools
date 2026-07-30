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

'Value 3 was TempSheetUpdates and it is gone. The sheet it named, updates__,
'survives in no live code. The remaining values keep the numbers they had so a
'stored scope byte still means what it meant.
Public Enum TempSheetScope
    TempSheetSingle = 1
    TempSheetList = 2
    TempSheetMetadata = 4
    TempSheetAnalysis = 5
    TempSheetImport = 6
    TempSheetSpatial = 7
    TempSheetShowHide = 8
End Enum

'@class LinelistSpecs
'@description
'Facade over the designer workbook. Before Prepare it answers the designer's
'own specifications. After Prepare it is also the handle on the domain managers
'that live on the generated linelist. Callers build one through
'LinelistSpecs.Create(workbook).
'
'WHICH WORKBOOK EACH COLLABORATOR POINTS AT
'-------------------------------------------------------------------------------
'The type says nothing about this, and it is where most of this class's defects
'came from. The table is the contract.
'
'  dictionary, choices, exports, ana    Linelist. Prepare sets them from the
'                                       build service. Reading one earlier
'                                       raises and names Prepare.
'  geo, passwords, translations         Designer until Prepare runs, linelist
'                                       afterwards. Prepare overwrites them.
'  format, formulaData                  Designer, loaded on first read.
'  Value, HasTemplate, DesWorkbook      Designer.
'  TemporarySheetName                   A constant table, backed by no sheet.
'
'The four names on the first row had designer-side loaders until 2026-07-30.
'RequiredSheetNames() stopped asking for those four worksheets when the build
'service began importing them straight from the setup file, so a loader either
'built a manager over a leftover designer sheet at the wrong anchor row, or
'raised a message naming the designer when the caller's real mistake was
'calling before Prepare. They are guards now.
'@depends LLdictionary, LLChoices, LLTranslation, LLGeo, LLExport, Analysis
'@depends Passwords, LLFormat, FormulaData, LLVariables, Formulas
'@depends LinelistBuildService, TemporaryRepos, ApplicationState
'@depends TranslationObject, HiddenNames, Checking, BetterArray
'@version 1.2 (2026-07-30)



'@section State
'===============================================================================
'@description Backing UDT and module-level constants.

Private Type TLinelistSpecs
    specsWkb As Workbook
    dictionary As LLdictionary
    choices As LLChoices
    translations As LLTranslation
    geo As LLGeo
    exports As LLExport
    ana As Analysis
    passwords As Passwords
    format As LLFormat
    formulaData As FormulaData
    variables As LLVariables
    'Translation scope cache: direct fields for each TranslationScope value (0-4)
    transMessages As TranslationObject
    transDictionary As TranslationObject
    transShapes As TranslationObject
    transForms As TranslationObject
    transRibbon As TranslationObject
    'Category cache, keyed by variable name and label choice. Mac Excel has no
    'Scripting.Dictionary, so a keyed Collection carries the lookup.
    catCache As Collection
    'What Value used to rebuild on every call
    mainSheet As Worksheet
    designerHidden As HiddenNames
    checks As Checking
    outputWkb As Workbook
    tempRepos As TemporaryRepos
    prepared As Boolean
    sealed As Boolean
End Type

Private this As TLinelistSpecs

Private checkCounter As Long

Private Const CLASSNAME As String = "LinelistSpecs"

'Worksheets read from the DESIGNER workbook (this.specsWkb). No constant here
'names a worksheet of the output workbook. The one list that used to,
'BuildPreservedSheetNames, reads the real worksheets of that workbook now.
Private Const DESIGNER_SHEET_GEO As String = "Geo"
Private Const DESIGNER_SHEET_PASSWORDS As String = "__pass"
Private Const DESIGNER_SHEET_FORMULAS As String = "__formula"
Private Const DESIGNER_SHEET_TRANSLATIONS_LL As String = "LinelistTranslation"
Private Const DESIGNER_SHEET_FORMAT As String = "__formatter"
Private Const DESIGNER_SHEET_MAIN As String = "Main"
Private Const DESIGNER_SHEET_DESIGNER_TRANSLATION As String = "DesignerTranslation"
Private Const RANGE_DESIGN_TYPE As String = "DESIGNTYPE"

'The debugging password lives in this named range on the designer's Main
'worksheet. The tag table asked for RNG_RNG_LLPassword, a doubled prefix that
'matches no name in the workbook, so the lookup answered an empty string under
'On Error Resume Next and every generated linelist got an empty debugging
'password. Measured in .mock/designer_mock.xlsb, which defines RNG_LLPassword.
Private Const RANGE_DEBUG_PASSWORD As String = "RNG_LLPassword"

'@section Factory helpers
'===============================================================================
'@description Construction and identity accessors.

'@label:create
'@sub-title Create a new specification facade from a workbook
'@details
'Validates the workbook structure upfront (checking for all required
'worksheets and named ranges), then constructs a new LinelistSpecs
'instance with caches initialised. Returns the LinelistSpecs interface.
'@param specsWorkbook Workbook. The specifications workbook to wrap.
'@return LinelistSpecs. A fully initialised facade ready for use.
'@throws ProjectError.ObjectNotInitialized When specsWorkbook is Nothing.
'@throws ProjectError.ElementNotFound When required worksheets are missing.
'@export
Public Function Create(ByVal specsWorkbook As Workbook) As LinelistSpecs
    ValidateWorkbookReference specsWorkbook

    Dim newSpecs As LinelistSpecs
    Set newSpecs = New LinelistSpecs
    Set newSpecs.InternalWorkbook = specsWorkbook
    newSpecs.ResetCaches
    newSpecs.Seal
    Set Create = newSpecs
End Function

'@label:seal
'@sub-title Seal the instance against further creation-only writes
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a creation-only setter against post-seal writes
'@param propName String. The property name for the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:des-workbook
'@prop-title Designer workbook backing all specification queries
'@details
'Returns the designer/specifications workbook reference. Raises an error
'if the workbook has not been assigned via the factory method.
'@return Workbook. The designer workbook.
'@throws ProjectError.ObjectNotInitialized When the workbook is not set.
'@export
Public Property Get DesWorkbook() As Workbook
    EnsureWorkbook
    Set DesWorkbook = this.specsWkb
End Property

'@label:internal-workbook
'@prop-title Setter for the backing workbook reference
'@details
'Used during construction to assign the workbook before caches are
'initialised. Not part of the public interface.
Public Property Set InternalWorkbook(ByVal specsWorkbook As Workbook)
    GuardNotSealed "InternalWorkbook"
    Set this.specsWkb = specsWorkbook
End Property

'@label:prepared
'@prop-title Whether Prepare has run on this instance
'@details
'Prepare imports the setup workbook into the linelist and points the
'collaborators at it. Linelist refuses to build until that has happened, and
'this is the flag it reads.
'@return Boolean. True once Prepare has completed.
'@export
Public Property Get Prepared() As Boolean
    Prepared = this.prepared
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
    Set this.geo = Nothing
    Set this.exports = Nothing
    Set this.ana = Nothing
    Set this.passwords = Nothing
    Set this.format = Nothing
    Set this.formulaData = Nothing
    Set this.variables = Nothing
    Set this.outputWkb = Nothing
    Set this.tempRepos = Nothing
    Set this.mainSheet = Nothing
    Set this.designerHidden = Nothing
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
    Set this.catCache = Nothing
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

'@label:ensure-category-cache
'@sub-title Create the category lookup if it does not exist
Private Sub EnsureCategoryCache()
    If this.catCache Is Nothing Then Set this.catCache = New Collection
End Sub

'@section Worksheet resolution
'===============================================================================
'@description Direct worksheet lookups by name within the specifications workbook.

'@label:resolve-sheet
'@sub-title Resolve a worksheet of the designer workbook by name
'@details
'Looks a worksheet up in the designer workbook by name and raises a
'descriptive error when it is absent.
'@param sheetName String. The worksheet name to find.
'@return Worksheet. The resolved worksheet.
'@throws ProjectError.InvalidArgument When sheetName is empty.
'@throws ProjectError.ElementNotFound When the sheet is missing.
Private Function ResolveSheet(ByVal sheetName As String) As Worksheet

    Dim target As Worksheet
    Dim normalisedName As String

    normalisedName = Trim$(sheetName)
    If LenB(normalisedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Worksheet name cannot be empty"
    End If

    EnsureWorkbook

    On Error Resume Next
        Set target = this.specsWkb.Worksheets(normalisedName)
    On Error GoTo 0

    If target Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & normalisedName & "' not found in specifications workbook '" & SafeWorkbookName(this.specsWkb) & "'"
    End If

    Set ResolveSheet = target
End Function

'@label:main-sheet
'@sub-title The designer's Main worksheet, resolved once per instance
'@details
'Value used to call ResolveSheet on every read, and Prepare and Linelist
'together read a dozen tags. ResetCaches drops this reference.
'@return Worksheet. The Main worksheet of the designer workbook.
'@throws ProjectError.ElementNotFound When the Main worksheet is missing.
Private Function MainSheet() As Worksheet
    If this.mainSheet Is Nothing Then
        Set this.mainSheet = ResolveSheet(DESIGNER_SHEET_MAIN)
    End If
    Set MainSheet = this.mainSheet
End Function

'@label:designer-hidden-store
'@sub-title The designer workbook's hidden name store, built once per instance
'@details
'Building one runs EnsureReady and a full metadata rebuild, and Value used to
'build a fresh store for each of the two hidden-name tags. ResetCaches drops it.
'@return HiddenNames. The store over the designer workbook.
Private Function DesignerHiddenStore() As HiddenNames
    If this.designerHidden Is Nothing Then
        EnsureWorkbook
        Set this.designerHidden = HiddenNames.Create(this.specsWkb)
    End If
    Set DesignerHiddenStore = this.designerHidden
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
'@param busyScope ApplicationState (ByRef). The state object to manage.
Private Sub EnterBusyState(ByRef busyScope As ApplicationState)
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

'@label:import-geobase
'@sub-title Import geobase data from the path the user gave
'@details
'Reads the geobase file path from the designer's Main worksheet. An empty path
'is a choice the user made and the import is skipped. A path that is set and
'will not open is a failure the user has to hear about: the linelist comes out
'with no geographic data at all, so the reason is filed as a report entry
'naming the path.
'
'The file is opened read only, the way ImportSetupToLinelist opens the setup
'workbook. The geobase is read and never written.
'@param geoObj LLGeo. The geo instance on the output workbook.
Private Sub ImportGeobase(ByVal geoObj As LLGeo)
    Dim geoPath As String
    Dim geoWkb As Workbook
    Dim openError As String

    geoPath = Trim$(Value("geopath"))
    If LenB(geoPath) = 0 Then
        LogCheck "No geobase file was given, so the linelist carries no geographic data.", _
                 checkingInfo
        Exit Sub
    End If

    On Error Resume Next
        Set geoWkb = Workbooks.Open(fileName:=geoPath, ReadOnly:=True)
        openError = Err.Description
    On Error GoTo 0

    If geoWkb Is Nothing Then
        LogCheck "The geobase file " & geoPath & " could not be opened, so the " & _
                 "linelist carries no geographic data. " & openError, _
                 checkingError
        Exit Sub
    End If

    geoObj.Import geoWkb
    geoWkb.Close saveChanges:=False
    LogCheck "The geobase " & geoPath & " was imported.", checkingSuccess
End Sub

'@label:exit-busy-state
'@sub-title Restore the application from a busy state
'@param busyScope ApplicationState (ByRef). The state object to restore and release.
Private Sub ExitBusyState(ByRef busyScope As ApplicationState)
    If busyScope Is Nothing Then Exit Sub

    On Error Resume Next
        busyScope.Restore
    On Error GoTo 0
    Set busyScope = Nothing
End Sub

'@label:build-preserved-sheet-names
'@sub-title Build the list of sheet names the dictionary must avoid
'@details
'The dictionary generates one data-entry worksheet per sheet name it holds, and
'a generated name that collides with a worksheet already in the linelist breaks
'the build. This is the list of names it has to keep clear of.
'
'It is read off the output workbook itself. Prepare has already run every
'exporter against that workbook by the time it calls this, so the real
'worksheets are the honest answer, and a list of constants drifts away from
'them on every change to the architecture. The constant list this replaced
'still carried a Translations worksheet the designer gave up, and it named
'Metadata as a temporary sheet when LLGeo creates a real one.
'
'On top of the real worksheets go the temporary names that do not exist yet,
'and the translated names of the worksheets the build is about to create.
'@param messageTranslations TranslationObject. The message translations that
'resolve the linelist worksheet names.
'@return BetterArray. The list of preserved sheet names.
Private Function BuildPreservedSheetNames(ByVal messageTranslations As TranslationObject) As BetterArray
    Dim preserved As BetterArray
    Dim sh As Worksheet
    Dim scopes As Variant
    Dim labels As Variant
    Dim idx As Long

    Set preserved = New BetterArray
    preserved.LowerBound = 1

    If Not this.outputWkb Is Nothing Then
        For Each sh In this.outputWkb.Worksheets
            preserved.Push sh.Name
        Next sh
    End If

    scopes = Array(TempSheetSingle, TempSheetList, TempSheetMetadata, _
                   TempSheetAnalysis, TempSheetImport, TempSheetSpatial, _
                   TempSheetShowHide)

    For idx = LBound(scopes) To UBound(scopes)
        preserved.Push TemporarySheetName(CByte(scopes(idx)))
    Next idx

    labels = Array("LLSHEET_Admin", "LLSHEET_Analysis", "LLSHEET_SpatialAnalysis", _
                   "LLSHEET_TemporalAnalysis", "LLSHEET_CustomChoice", _
                   "LLSHEET_CustomPivotTable", "LLSHEET_SpatioTemporalAnalysis")

    For idx = LBound(labels) To UBound(labels)
        preserved.Push SafeTranslatedValue(messageTranslations, CStr(labels(idx)))
    Next idx

    Set BuildPreservedSheetNames = preserved
End Function

'@label:safe-translated-value
'@sub-title Look up a translation key, falling back to the key itself on failure
'@param translations TranslationObject. The translation object to query.
'@param key String. The translation key.
'@return String. The translated value, or the key if translation fails.
Private Function SafeTranslatedValue(ByVal translations As TranslationObject, ByVal key As String) As String
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
        Set formatSheet = specsWorkbook.Worksheets(DESIGNER_SHEET_FORMAT)
    On Error GoTo 0

    If formatSheet Is Nothing Then Exit Sub

    If Not RangeExistsInternal(formatSheet, RANGE_DESIGN_TYPE) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & DESIGNER_SHEET_FORMAT & "' must define the '" & RANGE_DESIGN_TYPE & "' named range"
    End If
End Sub

'@label:required-sheet-names
'@sub-title Return the array of required designer worksheet names
'@details
'Only includes sheets that LinelistSpecs reads directly from the designer
'workbook. Pass-through sheets (Dictionary, Choices, Analysis, Export) are
'excluded because the build service imports them from the setup file
'directly into the linelist.
'@return Variant. Array of sheet name strings.
Private Function RequiredSheetNames() As Variant
    RequiredSheetNames = Array( _
        DESIGNER_SHEET_GEO, _
        DESIGNER_SHEET_PASSWORDS, _
        DESIGNER_SHEET_FORMULAS, _
        DESIGNER_SHEET_TRANSLATIONS_LL, _
        DESIGNER_SHEET_FORMAT, _
        DESIGNER_SHEET_MAIN, _
        DESIGNER_SHEET_DESIGNER_TRANSLATION)
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
'@description How each collaborator is obtained. Read the residency table in
'the class description first: it says which workbook each one is backed by.
'
'The collaborators that live on the DESIGNER are built here on first read, from
'a worksheet RequiredSheetNames() guarantees is present.
'
'The four that live on the LINELIST have no loader. Prepare sets them from the
'build service, and their accessors raise before that. A loader for them used
'to resolve a designer worksheet the designer no longer carries.

'@label:require-prepared
'@sub-title Refuse a read of a collaborator Prepare has to supply
'@details
'The four pass-through collaborators are imported from the setup file straight
'into the linelist. Reading one before Prepare has run is a caller mistake, and
'this is the message that says so.
'@param memberName String. The accessor the caller asked for.
'@throws ProjectError.ErrorUnexpectedState Always.
Private Sub RequirePrepared(ByVal memberName As String)
    ThrowError ProjectError.ErrorUnexpectedState, _
              memberName & " is available only after Prepare has imported the " & _
              "setup file into the linelist workbook"
End Sub

'@label:ensure-translations
'@sub-title Load the translations collaborator on first read
'@details
'Creates the LLTranslation instance from the designer's linelist translation
'worksheet. Prepare replaces it with the copy on the linelist.
Private Sub EnsureTranslations()
    If this.translations Is Nothing Then
        Dim linelistTransSheet As Worksheet
        Set linelistTransSheet = ResolveSheet(DESIGNER_SHEET_TRANSLATIONS_LL)
        Set this.translations = LLTranslation.Create(linelistTransSheet)
    End If
End Sub

'@label:ensure-geo
'@sub-title Load the geo collaborator on first read
'@details
'Built over the designer's Geo worksheet. Prepare replaces it with the copy on
'the linelist.
Private Sub EnsureGeo()
    If this.geo Is Nothing Then
        Dim geoSheet As Worksheet
        Set geoSheet = ResolveSheet(DESIGNER_SHEET_GEO)
        Set this.geo = LLGeo.Create(geoSheet)
    End If
End Sub

'@label:ensure-passwords
'@sub-title Load the passwords collaborator on first read
'@details
'Built over the designer's password worksheet. Prepare replaces it with the
'copy on the linelist.
Private Sub EnsurePasswords()
    If this.passwords Is Nothing Then
        Dim passwordSheet As Worksheet
        Set passwordSheet = ResolveSheet(DESIGNER_SHEET_PASSWORDS)
        Set this.passwords = Passwords.Create(passwordSheet)
    End If
End Sub

'@label:ensure-formula-data
'@sub-title Load the formula data collaborator on first read
Private Sub EnsureFormulaData()
    If this.formulaData Is Nothing Then
        Dim formulaSheet As Worksheet
        Set formulaSheet = ResolveSheet(DESIGNER_SHEET_FORMULAS)
        Set this.formulaData = FormulaData.Create(formulaSheet)
    End If
End Sub

'@label:ensure-variables
'@sub-title Build the variables collaborator over the dictionary
'@details
'Reads the dictionary through its accessor, so a call before Prepare raises
'there and names Prepare.
Private Sub EnsureVariables()
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(Dictionary())
    End If
End Sub

'@label:ensure-format
'@sub-title Load the format collaborator on first read
'@details
'Resolves the design name from the DESIGNTYPE named range on the designer's
'format worksheet, then creates the LLFormat instance.
Private Sub EnsureFormat()
    If this.format Is Nothing Then
        Dim formatSheet As Worksheet
        Dim designName As String

        Set formatSheet = ResolveSheet(DESIGNER_SHEET_FORMAT)
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
'@sub-title Ensure choices, variables and the category lookup are ready
'@details
'Choices reads through its accessor, so a call before Prepare raises there.
Private Sub EnsureCategoryLookup()
    Choices
    EnsureVariables
    EnsureCategoryCache
End Sub

'@label:number-of-exports
'@sub-title Return the number of exports defined in the export specifications
'@return Long. The export count.
Private Function NumberOfExports() As Long
    NumberOfExports = ExportObject().NumberOfExports()
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
    Dim dict As LLdictionary
    Dim vars As LLVariables
    Dim listVars As BetterArray
    Dim idx As Long
    Dim varName As String
    Dim controlType As String

    Set dict = Dictionary()
    EnsureVariables
    Set vars = this.variables

    dict.AddColumn "list auto"
    Set listVars = dict.Data.FilterData("control", "list_auto", "control details", matchCase:=False)

    If listVars Is Nothing Then Exit Sub

    For idx = listVars.LowerBound To listVars.UpperBound
        varName = CStr(listVars.Item(idx))

        vars.SetValue varName:=varName, colName:="List Auto", newValue:="list_auto_origin", onEmpty:=True

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
    Dim dict As LLdictionary
    Dim vars As LLVariables
    Dim controlType As String
    Dim controlDetails As String
    Dim formulaEngine As Formulas
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

    EnsureCategoryLookup

    cacheKey = LCase$(Trim$(varName)) & "|" & CStr(useShortlabels)

    Set cached = CachedCategories(cacheKey)

    If cached Is Nothing Then
        Set cached = ResolveCategoryList(varName, useShortlabels)
        this.catCache.Add cached, cacheKey
    End If

    'The clone stays: callers do change what they are handed.
    Set Categories = cached.Clone
End Function

'@label:cached-categories
'@sub-title Read one category list out of the lookup
'@details
'A Collection raises on a key it does not hold, which is how a miss is read.
'The lookup this replaced was two parallel BetterArrays walked one entry at a
'time, and every walk opened with a full copy of the internal array. A
'500-variable dictionary spent about 125,000 item reads on cache misses alone.
'@param cacheKey String. The key to look up.
'@return BetterArray. The cached list, or Nothing when the key is absent.
Private Function CachedCategories(ByVal cacheKey As String) As BetterArray
    On Error Resume Next
        Set CachedCategories = this.catCache.Item(cacheKey)
    On Error GoTo 0
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
    If (LenB(categoryName) = 0)  And (controlValue = "choice_manual" Or controlValue = "choice_multiple") Then
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
'The LLdictionary instance over the Dictionary worksheet of the LINELIST
'workbook. Prepare sets it from the build service.
'@return LLdictionary. The dictionary collaborator.
'@throws ProjectError.ErrorUnexpectedState When Prepare has not run.
'@export
Public Property Get Dictionary() As LLdictionary
    If this.dictionary Is Nothing Then RequirePrepared "Dictionary"
    Set Dictionary = this.dictionary
End Property

'@label:choices
'@prop-title Choices catalogue
'@details
'The LLChoices instance over the Choices worksheet of the LINELIST workbook.
'Prepare sets it from the build service.
'@return LLChoices. The choices collaborator.
'@throws ProjectError.ErrorUnexpectedState When Prepare has not run.
'@export
Public Property Get Choices() As LLChoices
    If this.choices Is Nothing Then RequirePrepared "Choices"
    Set Choices = this.choices
End Property

'@label:design-format
'@prop-title Visual formatting specifications
'@details
'Returns the LLFormat instance providing sheet and cell formatting
'methods based on the __formatter worksheet. Lazy-loaded on first access.
'@return LLFormat. The format collaborator.
'@export
Public Property Get DesignFormat() As LLFormat
    EnsureFormat
    Set DesignFormat = this.format
End Property

'@label:analysis
'@prop-title Analysis setup specifications
'@details
'The Analysis instance over the Analysis worksheet of the LINELIST workbook.
'Prepare sets it from the build service.
'@return Analysis. The analysis collaborator.
'@throws ProjectError.ErrorUnexpectedState When Prepare has not run.
'@export
Public Property Get AnalysisObject() As Analysis
    If this.ana Is Nothing Then RequirePrepared "AnalysisObject"
    Set AnalysisObject = this.ana
End Property

'@label:export-object
'@prop-title Export column specifications
'@details
'The LLExport instance over the Exports worksheet of the LINELIST workbook.
'Prepare sets it from the build service.
'@return LLExport. The export collaborator.
'@throws ProjectError.ErrorUnexpectedState When Prepare has not run.
'@export
Public Property Get ExportObject() As LLExport
    If this.exports Is Nothing Then RequirePrepared "ExportObject"
    Set ExportObject = this.exports
End Property

'@label:password
'@prop-title Password manager
'@details
'Returns the Passwords instance providing workbook and sheet protection
'utilities from the __pass worksheet. Lazy-loaded on first access.
'@return Passwords. The password collaborator.
'@export
Public Property Get Password() As Passwords
    EnsurePasswords
    Set Password = this.passwords
End Property

'@label:geo-object
'@prop-title Geographic data manager
'@details
'Returns the LLGeo instance providing geographic hierarchy lookups from
'the Geo worksheet. Lazy-loaded on first access.
'@return LLGeo. The geo collaborator.
'@export
Public Property Get GeoObject() As LLGeo
    EnsureGeo
    Set GeoObject = this.geo
End Property

'@label:formula-data-object
'@prop-title Formula data provider
'@details
'Returns the FormulaData instance providing formula parsing support from
'the __formula worksheet. Lazy-loaded on first access.
'@return FormulaData. The formula data collaborator.
'@export
Public Property Get FormulaDataObject() As FormulaData
    EnsureFormulaData
    Set FormulaDataObject = this.formulaData
End Property

'@label:has-template
'@prop-title Whether the workbook uses a ribbon template
'@details
'Reads the RNG_LLTemp named range on the designer's Main worksheet and answers
'True when it holds anything. An absent range answers False.
'@return Boolean. True when a template is configured.
'@export
Public Property Get HasTemplate() As Boolean
    On Error Resume Next
        HasTemplate = (MainSheet().Range("RNG_LLTemp").Value <> vbNullString)
    On Error GoTo 0
End Property

'@label:translations
'@prop-title Translation tables manager
'@details
'Returns the LLTranslation instance providing access to the linelist
'translation worksheet. Lazy-loaded on first access.
'@return LLTranslation. The translations collaborator.
'@export
Public Property Get Translations() As LLTranslation
    EnsureTranslations
    Set Translations = this.translations
End Property

'@label:trans-object
'@prop-title Scoped translation object
'@details
'Returns an TranslationObject for the requested translation scope.
'Results are cached per scope using direct fields for each TranslationScope
'value (0-4), avoiding generic container overhead.
'@param translationScope Byte. The translation scope constant
'(e.g., TranslationOfMessages, TranslationOfDictionary). Defaults to
'TranslationOfMessages.
'@return TranslationObject. The scoped translation object.
'@throws ProjectError.ElementNotFound When the scope is not available.
'@export
Public Property Get TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As TranslationObject
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
'@return TranslationObject. The resolved translation object.
'@throws ProjectError.ElementNotFound When the scope returns Nothing.
Private Function ResolveTranslationScope(ByVal translationScope As Byte) As TranslationObject
    Dim resolved As TranslationObject

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
'Answers one configuration value the user typed into the designer's Main
'worksheet. Most tags name a range on that worksheet; two read a hidden name
'on the designer workbook, and one counts the exports.
'
'A tag this class does not know raises. It used to answer an empty string, and
'that is what hid the debugging password fault for weeks: the tag table asked
'for a range name that exists nowhere, the read failed under On Error Resume
'Next, and every generated linelist got an empty password with nothing said.
'
'A tag the class knows whose RANGE is absent still answers an empty string.
'That is a designer the user has yet to fill in, and Prepare's caller validates
'the two values that have to be there.
'@param tagName String. The configuration key to look up.
'@return String. The configuration value, or an empty string for an empty tag.
'@throws ProjectError.InvalidArgument When the tag is unknown to this class.
'@export
Public Property Get Value(ByVal tagName As String) As String
    Dim normalizedTag As String
    Dim rngName As String
    Dim retVal As String
    Dim useHidden As Boolean

    normalizedTag = LCase$(Trim$(tagName))
    If LenB(normalizedTag) = 0 Then Exit Property

    Select Case normalizedTag
        Case "numberofexports"
            Value = CStr(NumberOfExports())
            Exit Property
        Case "llpassword"
            rngName = "RNG_LLPwdOpen"
        Case "debugpassword"
            rngName = RANGE_DEBUG_PASSWORD
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
        Case "instruct"
            'Hidden name at workbook level carrying the instructions choice
            rngName = "chkInstruct"
            useHidden = True
        Case "epiweekstart"
            'Week 1 is the default when the designer holds no choice
            On Error Resume Next
                Value = "1"
                Value = Split(MainSheet().Range("RNG_DefaultEpiWeek").Value, "-")(0)
            On Error GoTo 0
            Exit Property
        Case "temppath"
            On Error Resume Next
                retVal = MainSheet().Range("RNG_LLTemp").Value
            On Error GoTo 0
            If retVal = "_default" Or retVal = "default" Then
                retVal = this.specsWkb.Path & Application.PathSeparator & _
                         "_ribbontemplate.xlsb"
            End If
            Value = retVal
            Exit Property
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                      "This class knows no specification tag called '" & tagName & "'"
    End Select

    If useHidden Then
        Value = DesignerHiddenStore().ValueAsString(rngName)
    Else
        On Error Resume Next
            Value = MainSheet().Range(rngName).Value
        On Error GoTo 0
    End If
End Property


'@section Workbook orchestration
'===============================================================================
'@description Lifecycle methods for importing, preparing, exporting, and
'error-managing the specifications workbook.

'@label:prepare
'@sub-title Prepare all specifications using a build service
'@details
'Runs the whole preparation:
'1. Creates a temporary repository for working files.
'2. Creates the output workbook, from the template when one is configured.
'3. Imports the pass-through components straight from setup into the linelist.
'4. Exports the designer-sourced components (geo, passwords, translations,
'   formatter, language hidden names) into the linelist.
'5. Builds the domain managers on the linelist workbook.
'6. Takes those instances from the service, which is what points the four
'   pass-through collaborators at the linelist.
'7. Imports the geobase the user asked for.
'8. Prepares the dictionary with the preserved sheet names.
'9. Adds the "list auto" column when the dictionary has none.
'10. Translates and sorts the collaborators.
'11. Writes the debugging password into the linelist.
'
'A failure anywhere closes the output workbook, drops the temporary
'repository and restores the application state, then re-raises. The caller
'creates its Linelist after this returns, so a Prepare that leaked left an
'orphaned workbook open and temporary files on disk every time.
'@param buildService LinelistBuildService. The configured build service.
'@param setupPath String. Full path to the setup workbook.
'@throws ProjectError.ObjectNotInitialized When buildService is Nothing.
'@throws ProjectError.ErrorUnexpectedState When an unexpected error occurs.
'@export
Public Sub Prepare(ByVal buildService As LinelistBuildService, ByVal setupPath As String)
    Dim dict As LLdictionary
    Dim choicesObj As LLChoices
    Dim exportsObj As LLExport
    Dim analysisObj As Analysis
    Dim geoObj As LLGeo
    Dim dictTrads As TranslationObject
    Dim msgTrads As TranslationObject
    Dim preservedNames As BetterArray
    Dim busyScope As ApplicationState
    Dim outWkb As Workbook
    Dim templatePath As String
    Dim tempPath As String
    Dim pass As Passwords
    Dim errNumber As Long
    Dim errDescription As String

    If buildService Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "A build service is required to prepare specifications"
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

    'Export pass-through data directly from setup to linelist (skip designer)
    buildService.ImportSetupToLinelist setupPath, outWkb

    'Export designer-sourced components (geo, passwords, translations, etc.)
    buildService.ExportDesignerComponents outWkb

    'Build domain manager instances on the linelist workbook
    buildService.BuildExportedInstances outWkb

    'Retrieve exported domain instances from the service
    Set this.dictionary = buildService.ExportedDictionary
    Set this.choices = buildService.ExportedChoices
    Set this.exports = buildService.ExportedExport
    Set this.ana = buildService.ExportedAnalysis
    Set this.geo = buildService.ExportedGeo
    Set this.passwords = buildService.ExportedPasswords
    Set this.translations = buildService.ExportedTranslations

    Set this.outputWkb = outWkb

    Set dict = this.dictionary
    Set choicesObj = this.choices
    Set exportsObj = this.exports
    Set analysisObj = this.ana
    Set geoObj = this.geo
    Set pass = this.passwords

    Set dictTrads = TransObject(TranslationOfDictionary)
    Set msgTrads = TransObject()

    'Import the geobase the user asked for. An empty path is a valid choice.
    ImportGeobase geoObj

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

    'Update the debugging Password
    pass.NewDebugingPassword Value("debugpassword")
    this.prepared = True

CleanExit:
    ExitBusyState busyScope
    If this.prepared Then ResetPostImportCaches
    Exit Sub

'The cleanup label runs with the handler disarmed, so the work below can raise
'without ending the routine where it stands. A handler that does its own
'cleanup throws away everything after the first raise, because VBA cannot trap
'an error raised while a handler is active.
CleanUpAfterFail:
    On Error Resume Next
        ExitBusyState busyScope
        If Not outWkb Is Nothing Then outWkb.Close saveChanges:=False
        Set this.outputWkb = Nothing
        Set this.tempRepos = Nothing
    On Error GoTo 0

    If errNumber = 0 Then Exit Sub

    'ProjectError runs 1001, 1002 and 1004 to 1008. The test this replaced was
    'a range from ErrorNone to ElementNotFound, which left SomethingWentWrong
    'above the upper bound and rewrote it as ErrorUnexpectedState. The build
    'service raises exactly that code when it cannot open the setup workbook,
    'which is the most likely failure of the direct import path, and its
    'identity was destroyed. The old window also swallowed the whole native VBA
    'range, where Excel's own 1004 collides with ErrorUnexpectedState.
    Select Case errNumber
        Case ProjectError.InvalidArgument, ProjectError.ObjectNotInitialized, _
             ProjectError.ErrorUnexpectedState To ProjectError.SomethingWentWrong
            Err.Raise errNumber, CLASSNAME, errDescription
        Case Else
            ThrowError ProjectError.ErrorUnexpectedState, errDescription
    End Select
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanUpAfterFail
End Sub

'@label:ll-workbook
'@prop-title The output linelist workbook created during Prepare
'@details
'Returns the linelist workbook created by Prepare that received exported
'data from the import service. Returns Nothing before Prepare is called.
'@return Workbook. The linelist workbook, or Nothing.
'@export
Public Property Get LLWorkbook() As Workbook
    Set LLWorkbook = this.outputWkb
End Property

'@label:temporary-repos
'@prop-title The temporary repository created during Prepare
'@details
'Returns the TemporaryRepos instance created by Prepare for working files.
'Returns Nothing before Prepare is called.
'@return TemporaryRepos. The temporary repository, or Nothing.
'@export
Public Property Get OutputTemporaryRepos() As TemporaryRepos
    Set OutputTemporaryRepos = this.tempRepos
End Property

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

'@section Checkings
'===============================================================================
'@description What this class reports into the generation report.
'GenerationReport.HarvestSpecsCheckings collects these after Prepare, alongside
'the entries of the dictionary, the choices, the exports, the analysis, the
'passwords and the format.

'@label:checks
'@prop-title The report entries of this instance
'@return Checking. The entries, created on first use.
Private Property Get Checks() As Checking
    If this.checks Is Nothing Then Set this.checks = Checking.Create(CLASSNAME)
    Set Checks = this.checks
End Property

'@label:log-check
'@sub-title File one report entry
'@details
'Checking.Add raises on a duplicate key, and one report carries the entries of
'several classes, so the key names this class and a counter.
'@param message String. The message text.
'@param scope Optional Byte. The checking scope. Defaults to checkingNote.
Private Sub LogCheck(ByVal message As String, _
                     Optional ByVal scope As Byte = checkingNote)
    Dim entries As Checking

    Set entries = Checks
    checkCounter = checkCounter + 1
    entries.Add CLASSNAME & "-" & CStr(checkCounter), message, scope
End Sub

'@label:has-checkings
'@prop-title Whether this instance has anything to report
'@return Boolean. True once one entry has been filed.
'@export
Public Property Get HasCheckings() As Boolean
    If this.checks Is Nothing Then Exit Property
    HasCheckings = (this.checks.Length() > 0)
End Property

'@label:checking-values
'@prop-title The entries filed while the specifications were prepared
'@return Checking. The entries, or Nothing when none were filed.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = this.checks
End Property

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

'@label:test-assign-dictionary
'@sub-title Inject a dictionary collaborator for testing
'@param dict LLdictionary. The dictionary to expose through Dictionary().
Public Sub TestAssignDictionary(ByVal dict As LLdictionary)
    Set this.dictionary = dict
End Sub

'@label:test-assign-design-format
'@sub-title Inject a design-format collaborator for testing
'@param fmt LLFormat. The format object to expose through DesignFormat().
Public Sub TestAssignDesignFormat(ByVal fmt As LLFormat)
    Set this.format = fmt
End Sub

'@label:test-assign-formula-data
'@sub-title Inject a formula-data collaborator for testing
'@param data FormulaData. The formula data to expose through FormulaDataObject().
Public Sub TestAssignFormulaData(ByVal data As FormulaData)
    Set this.formulaData = data
End Sub

'@label:test-assign-trans-object
'@sub-title Inject one TranslationObject for every translation scope
'@details Sets all five scope-cache fields so TransObject returns the same
'object regardless of scope, mirroring the former single-object test stubs.
'@param trans TranslationObject. The translation object to expose.
Public Sub TestAssignTransObject(ByVal trans As TranslationObject)
    Set this.transMessages = trans
    Set this.transDictionary = trans
    Set this.transShapes = trans
    Set this.transForms = trans
    Set this.transRibbon = trans
End Sub
