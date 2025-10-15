VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseTranslationCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface providing cached translation lookups for disease exports"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface providing cached translation lookups for disease exports")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Configuration
'===============================================================================

'@sub-title Prime the cache with the translation source.
'@param translations Optional ITranslationObject providing translated values.
Public Sub Initialise(Optional ByVal translations As ITranslationObject)
End Sub

'@section Queries
'===============================================================================

'@sub-title Retrieve a translated string using cache-aware lookup.
'@param key String translation key to resolve.
'@param fallback String returned when translation is missing.
'@return String translated value or fallback when unavailable.
Public Function Resolve(ByVal key As String, ByVal fallback As String) As String
End Function

'@section Maintenance
'===============================================================================

'@sub-title Clear cached values and detach the translation object.
Public Sub Reset()
End Sub
