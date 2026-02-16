VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract describing reusable linelist specifications accessors"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Contract describing reusable linelist specifications accessors")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ILinelistSpecs
'@description
'Defines the public contract for the LinelistSpecs class, exposing lazy-loaded
'accessors for all specification collaborators (dictionary, choices, translations,
'geo, analysis, passwords, format, exports, formulas, main), metadata helpers
'for temporary sheet names and category lookups, and lifecycle methods for
'preparing, exporting, and error-managing the specifications workbook.


'@section Lifecycle
'===============================================================================
'@description Cache management and workbook reference.

'@jump:reset-caches
'@sub-title Refresh all cached collaborators
'@details
'Resets every cached collaborator object, sheet reference, translation cache,
'and category cache so subsequent reads query the workbook again. Call after
'importing new data to ensure stale caches are cleared.
Public Sub ResetCaches()
End Sub

'@jump:workbook
'@prop-title Workbook hosting the specifications
'@details
'Returns the workbook reference that backs all specification queries. Raises
'an error if the workbook has not been assigned.
'@return Workbook. The specifications workbook.
Public Property Get Workbook() As Workbook
End Property


'@section CoreSpecifications
'===============================================================================
'@description Lazy-loaded accessors for specification collaborator objects.

'@jump:dictionary
'@prop-title Linelist dictionary
'@details
'Returns the ILLdictionary instance providing column and row access to the
'Dictionary worksheet. Lazy-loaded on first access.
'@return ILLdictionary. The dictionary collaborator.
Public Property Get Dictionary() As ILLdictionary
End Property

'@jump:choices
'@prop-title Choices catalogue
'@details
'Returns the ILLChoices instance providing category lookups from the Choices
'worksheet. Lazy-loaded on first access.
'@return ILLChoices. The choices collaborator.
Public Property Get Choices() As ILLChoices
End Property

'@jump:design-format
'@prop-title Visual formatting specifications
'@details
'Returns the ILLFormat instance providing sheet and cell formatting methods
'based on the LinelistStyle worksheet. Lazy-loaded on first access.
'@return ILLFormat. The format collaborator.
Public Property Get DesignFormat() As ILLFormat
End Property

'@jump:analysis
'@prop-title Analysis setup specifications
'@details
'Returns the ILLAnalysis instance providing access to the Analysis worksheet
'configuration. Lazy-loaded on first access.
'@return ILLAnalysis. The analysis collaborator.
Public Property Get Analysis() As ILLAnalysis
End Property

'@jump:export-object
'@prop-title Export column specifications
'@details
'Returns the ILLExport instance providing access to export column definitions
'from the Exports worksheet. Lazy-loaded on first access.
'@return ILLExport. The export collaborator.
Public Property Get ExportObject() As ILLExport
End Property

'@jump:password
'@prop-title Password manager
'@details
'Returns the IPasswords instance providing workbook and sheet protection
'utilities from the __pass worksheet. Lazy-loaded on first access.
'@return IPasswords. The password collaborator.
Public Property Get Password() As IPasswords
End Property

'@jump:geo-object
'@prop-title Geographic data manager
'@details
'Returns the ILLGeo instance providing geographic hierarchy lookups from the
'Geo worksheet. Lazy-loaded on first access.
'@return ILLGeo. The geo collaborator.
Public Property Get GeoObject() As ILLGeo
End Property

'@jump:formula-data-object
'@prop-title Formula data provider
'@details
'Returns the IFormulaData instance providing formula parsing support from the
'__formula worksheet. Lazy-loaded on first access.
'@return IFormulaData. The formula data collaborator.
Public Property Get FormulaDataObject() As IFormulaData
End Property

'@jump:main-object
'@prop-title Main configuration sheet accessor
'@details
'Returns the IMain instance providing key-value configuration from the Main
'worksheet. Lazy-loaded on first access.
'@return IMain. The main configuration collaborator.
Public Property Get MainObject() As IMain
End Property

'@jump:translations
'@prop-title Translation tables manager
'@details
'Returns the ILLTranslation instance providing access to all translation
'ListObjects on the consolidated translation worksheet. Lazy-loaded on
'first access.
'@return ILLTranslation. The translations collaborator.
Public Property Get Translations() As ILLTranslation
End Property

'@jump:designer-translations
'@prop-title Designer translation accessor
'@details
'Returns the IDesignerTranslation instance providing designer UI translations from
'the DesignerTranslation worksheet. Lazy-loaded on first access.
'@return IDesignerTranslation. The designer translation collaborator.
Public Property Get DesignerTranslations() As IDesignerTranslation
End Property

'@jump:trans-object
'@prop-title Scoped translation object
'@details
'Returns an ITranslationObject for the requested translation scope. Results
'are cached per scope value for efficient repeated access.
'@param translationScope Optional Byte. The translation scope constant
'(e.g., TranslationOfMessages, TranslationOfDictionary, TranslationOfShapes).
'Defaults to TranslationOfMessages.
'@return ITranslationObject. The scoped translation object.
Public Property Get TransObject(Optional ByVal translationScope As Byte = TranslationOfMessages) As ITranslationObject
End Property


'@section MetadataHelpers
'===============================================================================
'@description Derived accessors for temporary sheet names, tag values, and categories.

'@jump:temporary-sheet-name
'@prop-title Temporary sheet name for a given scope
'@details
'Returns the fixed worksheet name used for temporary internal sheets, based
'on the scope constant (TEMP_SHEET_ANALYSIS, TEMP_SHEET_LIST, etc.).
'@param scope Byte. The temporary sheet scope constant.
'@return String. The worksheet name string.
Public Property Get TemporarySheetName(ByVal scope As Byte) As String
End Property

'@jump:value
'@prop-title Configuration value by tag name
'@details
'Returns a configuration value from the Main worksheet for the given tag name.
'The special tag "numberofexports" returns the export count from the Exports
'worksheet; all other tags delegate to MainObject.ValueOf.
'@param tagName String. The configuration key to look up.
'@return String. The configuration value.
Public Property Get Value(ByVal tagName As String) As String
End Property

'@jump:categories
'@sub-title Resolve category values for a variable
'@details
'Returns the list of category values for a choice-type variable (choice_manual,
'choice_multiple, choice_custom, or choice_formula). Results are cached per
'variable name and short-label flag for efficient repeated access.
'@param varName String. The variable name whose categories to resolve.
'@param useShortlabels Optional Boolean. When True, returns short labels
'instead of full labels. Defaults to False.
'@return BetterArray. Cloned list of category values.
Public Function Categories(ByVal varName As String, Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Function


'@section WorkbookManagement
'===============================================================================
'@description Preparation, export, and error management operations.

'@jump:prepare
'@sub-title Import and prepare all specifications from external workbooks
'@details
'Orchestrates the full specification preparation workflow: validates the Main
'sheet readiness, opens the setup workbook, imports dictionary, choices,
'translations, analysis, and optionally geo data, then prepares the dictionary
'with preserved names, adds list auto columns, translates and sorts all
'collaborators. Handles errors with cleanup of open workbooks.
Public Sub Prepare()
End Sub

'@jump:export
'@sub-title Export all specifications to an output workbook
'@details
'Exports dictionary, choices, translations, export specs, geo data, and
'passwords to the given workbook, creating hidden worksheets for each.
'@param toWkb Workbook. The destination workbook to export specifications into.
Public Sub Export(ByVal toWkb As Workbook)
End Sub

'@jump:error-manage
'@sub-title Display an error dialog for specification failures
'@details
'Shows a message box describing the error and aborts the linelist creation
'process.
'@param textMessage Optional String. Error description. Defaults to vbNullString.
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub
