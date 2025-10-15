VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseSheetBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface constructing disease worksheets with validations and translations"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface constructing disease worksheets with validations and translations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Configuration
'===============================================================================

'@sub-title Configure builder dependencies.
'@param hostWorkbook Workbook where disease sheets are created.
'@param dropdowns IDropdownLists providing data validation helpers.
'@param translations ITranslationObject supplying translated labels.
Public Sub Initialise(ByVal hostWorkbook As Workbook, _
                      ByVal dropdowns As IDropdownLists, _
                      ByVal translations As ITranslationObject)
End Sub

'@section Worksheet Construction
'===============================================================================

'@sub-title Create a disease worksheet populated with metadata, headers, and validations.
'@param diseaseName String worksheet name to create.
'@param diseaseIndex Long sequential index used for names and counters.
'@param languageTag Optional String initial language value; defaults to translated prompt.
'@param options Optional Dictionary-like BetterArray containing overrides (languageList, statusList, etc.).
'@return Worksheet freshly created and formatted.
Public Function Build(ByVal diseaseName As String, _
                      ByVal diseaseIndex As Long, _
                      Optional ByVal languageTag As String = vbNullString, _
                      Optional ByVal options As BetterArray) As Worksheet
End Function

'@section Maintenance
'===============================================================================

'@sub-title Reset cached state so the builder can be reused safely.
Public Sub Reset()
End Sub
