VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelistSheetNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides safe access and activation helpers for linelist worksheets"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Provides safe access and activation helpers for linelist worksheets")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit

'@section Core navigation
'===============================================================================
'@sub-title Activate the admin worksheet using cached translation metadata.
Public Sub ActivateAdminSheet()
End Sub

'@sub-title Activate the instruction worksheet using cached translation metadata.
Public Sub ActivateInstructionSheet()
End Sub

'@sub-title Activate a translated sheet by base name and scope.
Public Function ActivateSheet(ByVal baseName As String, _
                              Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
End Function

'@sub-title Resolve a worksheet by translated base name without activation.
Public Function ResolveSheet(ByVal baseName As String, _
                             Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Worksheet
End Function

'@sub-title Check whether a translated sheet exists for the provided scope.
Public Function SheetExists(ByVal baseName As String, _
                            Optional ByVal scope As LinelistSheetScope = sheetScopeStandard) As Boolean
End Function

'@section Diagnostics
'===============================================================================
'@sub-title Returns the last sheet name activated by the navigator (scoped name).
Public Property Get LastActivatedSheetName() As String
End Property

'@sub-title Returns the number of cached translation lookups performed.
Public Property Get TranslationCacheHits() As Long
End Property

