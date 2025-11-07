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
'@depends IApplicationState
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Initialisation
'===============================================================================
'@sub-title Provide the translations table listobject and optional application scope.
'@param translationTable ListObject representing the translations table.
'@param stateScope Optional IApplicationState used to guard Excel settings.
Public Sub Initialise(ByVal translationTable As ListObject, Optional ByVal stateScope As IApplicationState)
End Sub

'@section Configuration
'===============================================================================
'@sub-title Retrieve the worksheet hosting the translations table.
Public Property Get HostWorksheet() As Worksheet
End Property

'@sub-title Retrieve the managed listobject name.
Public Property Get TableName() As String
End Property

'@section Update pipeline
'===============================================================================
'@sub-title Ensure that language columns exist within the table.
'@param languages Optional semicolon separated list of language headers to append.
Public Sub EnsureLanguages(Optional ByVal languages As String = vbNullString)
End Sub

'@sub-title Promote an existing language column to become the default language.
'@param languageName Name of the language column to promote.
Public Sub SwitchDefaultLanguage(ByVal languageName As String)
End Sub

'@sub-title Retrieve languages that can be promoted to default.
Public Function SwitchableLanguages() As BetterArray
End Function

'@sub-title Retrieve translation languages excluding the default column.
Public Property Get Languages() As BetterArray
End Property

'@sub-title Synchronise translations against the supplied registry sheet.
'@param registrySheet Worksheet hosting the registry listobjects with named ranges and statuses.
'@param languages Optional semicolon separated set of languages to ensure before syncing.
Public Sub UpdateFromRegistry(ByVal registrySheet As Worksheet, Optional ByVal languages As String = vbNullString)
End Sub
'@sub-title Reset the hidden update sequence counter.
Public Sub ResetSequence(ByVal registrySheet As Worksheet)
End Sub

'@sub-title Summarise and optionally display missing translation counts.
Public Function NumberOfMissing() As String
End Function
'@sub-title Enable or disable prompts when reporting missing translations.
Public Sub SetDisplayPrompts(ByVal state As Boolean)
End Sub

Public Function MissingLabels(ByVal lang As String) As Long
End Function

'@section Identity
'===============================================================================
'@sub-title Retrieve the interface bound to the current instance.
Public Property Get Self() As ISetupTranslationsTable
End Property
