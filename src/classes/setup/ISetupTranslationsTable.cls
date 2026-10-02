VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISetupTranslationsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the improved translations table manager"

'@Interface
'@Folder("Setup")
'@ModuleDescription("Interface describing the improved translations table manager with caching and guarded updates")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class ISetupTranslationsTable
'@description
'Defines the public contract for the SetupTranslationsTable class, exposing
'methods to initialise, configure, and synchronise a managed translations
'ListObject against a watcher registry, handle language columns, report
'missing or duplicate labels, and export translations to external workbooks.
'@depends IApplicationState


'@section Initialisation
'===============================================================================

'@jump:initialise
'@sub-title Provide the translations table ListObject and optional application scope.
'@param translationTable ListObject. The managed translations table.
'@param stateScope Optional IApplicationState. Guard for Excel settings during updates.
'@export
Public Sub Initialise(ByVal translationTable As ListObject, Optional ByVal stateScope As IApplicationState)
End Sub


'@section Configuration
'===============================================================================

'@jump:host-worksheet
'@prop-title Retrieve the worksheet hosting the translations table.
'@return Worksheet. The parent worksheet of the managed ListObject.
'@export
Public Property Get HostWorksheet() As Worksheet
End Property

'@jump:table-name
'@prop-title Retrieve the managed ListObject name.
'@return String. The name of the translations table.
'@export
Public Property Get TableName() As String
End Property


'@section Update pipeline
'===============================================================================

'@jump:ensure-languages
'@sub-title Ensure that language columns exist within the table.
'@param languages Optional String. Semicolon-separated list of language headers to append.
'@export
Public Sub EnsureLanguages(Optional ByVal languages As String = vbNullString)
End Sub

'@jump:switch-default-language
'@sub-title Promote an existing language column to become the default language.
'@param languageName String. Name of the language column to promote.
'@export
Public Sub SwitchDefaultLanguage(ByVal languageName As String)
End Sub

'@jump:switchable-languages
'@sub-title Retrieve languages that can be promoted to default.
'@return BetterArray. Array of non-default language names.
'@export
Public Function SwitchableLanguages() As BetterArray
End Function

'@jump:languages
'@prop-title Retrieve translation languages excluding the default column unless requested.
'@param includeDefault Optional Boolean. When True, includes the default language. Defaults to False.
'@return BetterArray. Array of language names.
'@export
Public Property Get Languages(Optional ByVal includeDefault As Boolean = False) As BetterArray
End Property

'@jump:update-from-registry
'@sub-title Synchronise translations against the supplied registry sheet.
'@param registrySheet Worksheet. Worksheet hosting the registry ListObjects with named ranges and statuses.
'@param languages Optional String. Semicolon-separated set of languages to ensure before syncing.
'@export
Public Sub UpdateFromRegistry(ByVal registrySheet As Worksheet, Optional ByVal languages As String = vbNullString)
End Sub

'@jump:reset-sequence
'@sub-title Reset the hidden update sequence counter.
'@param registrySheet Worksheet. The worksheet whose counter should be reset.
'@export
Public Sub ResetSequence(ByVal registrySheet As Worksheet)
End Sub

'@jump:number-of-missing
'@sub-title Summarise and optionally display missing translation counts.
'@return String. Summary message listing missing counts per language.
'@export
Public Function NumberOfMissing() As String
End Function

'@jump:duplicate-labels
'@sub-title Summarise duplicate labels for the specified language or across all languages.
'@param duplicateMessage String (ByRef). Receives the formatted duplicate report.
'@param languageName Optional String. Scope the check to a single language. Defaults to all.
'@return Boolean. True when duplicates were detected.
'@export
Public Function DuplicateLabels(ByRef duplicateMessage As String, Optional ByVal languageName As String = vbNullString) As Boolean
End Function

'@jump:internal-missing-summary
'@prop-title Retrieve the cached missing summary across languages.
'@return BetterArray. Cloned copy of the internal missing-labels summary.
'@export
Public Property Get InternalMissingSummary() As BetterArray
End Property

'@jump:internal-duplicates-summary
'@prop-title Retrieve the cached duplicate summary across languages.
'@return BetterArray. Cloned copy of the internal duplicates summary.
'@export
Public Property Get InternalDuplicatesSummary() As BetterArray
End Property

'@jump:set-display-prompts
'@sub-title Enable or disable prompts when reporting missing translations.
'@param state Boolean. True enables prompts; False suppresses them.
'@export
Public Sub SetDisplayPrompts(ByVal state As Boolean)
End Sub

'@jump:missing-labels
'@sub-title Count missing translation labels for a specific language.
'@param lang String. Language name to inspect.
'@return Long. Number of blank cells in the language column.
'@export
Public Function MissingLabels(ByVal lang As String) As Long
End Function


'@section Data exchange
'===============================================================================

'@jump:export
'@sub-title Export translations into a destination workbook.
'@param exportWorkbook Workbook. Destination workbook receiving the translations sheet.
'@param Hide Optional Long. Visibility flag applied to the exported sheet. Defaults to xlSheetVisible.
'@export
Public Sub Export(ByVal exportWorkbook As Workbook, Optional ByVal Hide As Long = xlSheetVisible)
End Sub


'@section Identity
'===============================================================================

'@jump:self
'@prop-title Retrieve the interface bound to the current instance.
'@return ISetupTranslationsTable. The typed self-reference.
Public Property Get Self() As ISetupTranslationsTable
End Property
