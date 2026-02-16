VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface orchestrating disease export workflows"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface orchestrating disease export workflows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Configuration
'===============================================================================

'@sub-title Configure dependencies required for the exporter.
'@param workbookManager IDiseaseExportWorkbook responsible for workbook lifecycle.
'@param appGuard IDiseaseApplicationState guarding Application state toggles.
Public Sub Initialise(ByVal workbookManager As IDiseaseExportWorkbook, ByVal appGuard As IDiseaseApplicationState)
End Sub

'@section Disease Export
'===============================================================================

'@sub-title Build an export workbook in-memory for the supplied disease worksheet.
'@param diseaseSheet Worksheet containing the disease table and metadata.
'@param translationTable ListObject providing translation values to copy.
'@param ribbonTranslations ITranslationObject translating choice headers (may be Nothing).
'@param diseaseName String disease identifier used in metadata.
'@param languageTag String language captured in metadata.
'@param diseaseCode Optional String disease tag/code for metadata.
'@return Workbook populated with export worksheets. Caller owns lifecycle.
Public Function BuildDiseaseWorkbook(ByVal diseaseSheet As Worksheet, _
                                     ByVal translationTable As ListObject, _
                                     ByVal ribbonTranslations As ITranslationObject, _
                                     ByVal diseaseName As String, _
                                     ByVal languageTag As String, _
                                     Optional ByVal diseaseCode As String = vbNullString) As Workbook
End Function

'@sub-title Export a disease workbook to disk.
'@param exportFolder String folder path where the workbook is saved.
'@param diseaseSheet Worksheet containing the disease table.
'@param translationTable ListObject providing translation values.
'@param ribbonTranslations ITranslationObject translating choice headers.
'@param diseaseName String disease identifier used in metadata and file name.
'@param languageTag String disease language stored in metadata.
'@param diseaseCode Optional String disease tag/code for metadata.
'@return String file path of the saved workbook.
Public Function ExportDisease(ByVal exportFolder As String, _
                              ByVal diseaseSheet As Worksheet, _
                              ByVal translationTable As ListObject, _
                              ByVal ribbonTranslations As ITranslationObject, _
                              ByVal diseaseName As String, _
                              ByVal languageTag As String, _
                              Optional ByVal diseaseCode As String = vbNullString) As String
End Function

'@section Migration Export
'===============================================================================

'@sub-title Build a migration workbook aggregating multiple diseases.
'@param sourceWorkbook Workbook hosting dictionary/choices/translation sheets.
'@param diseaseNames BetterArray listing disease worksheet names to export.
'@return Workbook populated with migration data. Caller owns lifecycle.
Public Function BuildMigrationWorkbook(ByVal sourceWorkbook As Workbook, _
                                       ByVal diseaseNames As BetterArray) As Workbook
End Function

'@sub-title Export the migration workbook to disk.
'@param exportFolder String folder path where the workbook is saved.
'@param sourceWorkbook Workbook hosting dictionary resources.
'@param diseaseNames BetterArray listing disease worksheet names to export.
'@return String file path of the saved workbook.
Public Function ExportForMigration(ByVal exportFolder As String, _
                                   ByVal sourceWorkbook As Workbook, _
                                   ByVal diseaseNames As BetterArray) As String
End Function
