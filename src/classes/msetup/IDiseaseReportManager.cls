VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseReportManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface providing helpers for disease import report maintenance"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface providing helpers for disease import report maintenance")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Queries
'===============================================================================

'@sub-title Check whether a disease requires a follow-up report.
'@param reportTable ListObject hosting the import report entries.
'@param diseaseName String disease identifier to search for.
'@param expectedValue Optional String value indicating a pending report (defaults to ""yes"").
'@return Boolean True when the disease row exists and matches the expected flag.
Public Function HasPendingReport(ByVal reportTable As ListObject, ByVal diseaseName As String, Optional ByVal expectedValue As String = "yes") As Boolean
End Function

'@section Commands
'===============================================================================

'@sub-title Remove all report rows related to the disease from every table on the worksheet.
'@param reportSheet Worksheet hosting import report tables.
'@param diseaseName String disease identifier to purge.
Public Sub ClearReportStatus(ByVal reportSheet As Worksheet, ByVal diseaseName As String)
End Sub
