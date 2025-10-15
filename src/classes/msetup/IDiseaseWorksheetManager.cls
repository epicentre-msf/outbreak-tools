VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseWorksheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface providing safe worksheet operations for disease setup workflows"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface providing safe worksheet operations for disease setup workflows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Queries
'===============================================================================

'@sub-title Determine whether the specified worksheet exists within the workbook.
'@param workbook Workbook hosting the worksheet collection.
'@param sheetName String worksheet name to locate.
'@return Boolean True when the worksheet exists.
Public Function WorksheetExists(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
End Function

'@section Commands
'===============================================================================

'@sub-title Remove a worksheet using guarded Application settings.
'@param workbook Workbook hosting the worksheet.
'@param sheetName String worksheet name to delete.
'@return Boolean True when the worksheet was removed.
Public Function RemoveWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
End Function
