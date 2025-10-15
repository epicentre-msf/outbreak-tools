VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface coordinating disease worksheet import and merge operations"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface coordinating disease worksheet import and merge operations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================

Public Enum DiseaseImportPriority
    DiseaseImportPriority_Existing = 0
    DiseaseImportPriority_Foreign = 1
End Enum

'@section Operations
'===============================================================================

'@sub-title Import a disease table, optionally merging with the existing one.
'@param targetTable ListObject representing the existing disease worksheet data.
'@param sourceTable ListObject containing the imported disease data.
'@param mergeValues Optional Boolean. When False the target table is replaced.
'@param priority Optional DiseaseImportPriority controlling conflict resolution. Defaults to foreign.
'@return IDiseaseImportSummary containing the merge outcome.
Public Function MergeDisease(ByVal targetTable As ListObject, ByVal sourceTable As ListObject, _
                             Optional ByVal mergeValues As Boolean = True, _
                             Optional ByVal priority As DiseaseImportPriority = DiseaseImportPriority_Foreign, _
                             Optional ByVal logger As IDiseaseLogger) As IDiseaseImportSummary
End Function
