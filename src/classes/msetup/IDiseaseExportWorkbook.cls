VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseExportWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface providing a safe lifecycle for export workbooks used by disease workflows"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface providing a safe lifecycle for export workbooks used by disease workflows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Lifecycle
'===============================================================================

'@sub-title Ensure a workbook instance is available for exports.
'@param excelApp Application host used to create the workbook.
'@return Workbook ready for export operations.
Public Function EnsureWorkbook(ByVal excelApp As Application) As Workbook
End Function

'@sub-title Persist the workbook to disk.
'@param filePath String destination path.
'@param fileFormat Optional XlFileFormat. Defaults to xlOpenXMLWorkbook.
'@param createBackup Optional Boolean. Defaults to False.
Public Sub SaveAs(ByVal filePath As String, _
                  Optional ByVal fileFormat As XlFileFormat = xlOpenXMLWorkbook, _
                  Optional ByVal createBackup As Boolean = False)
End Sub

'@sub-title Close the workbook and clear references.
'@param saveChanges Optional Boolean controlling SaveChanges flag.
Public Sub Close(Optional ByVal saveChanges As Boolean = False)
End Sub

'@sub-title Release the workbook without saving, ensuring references are cleared.
Public Sub Release()
End Sub

'@section State
'===============================================================================

'@sub-title Indicates whether a workbook is currently managed.
Public Property Get HasWorkbook() As Boolean
End Property

'@sub-title Retrieve the managed workbook reference.
Public Property Get Workbook() As Workbook
End Property
