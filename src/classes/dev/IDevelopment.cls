VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the development workflow coordinator"

'@Interface
'@Folder("Dev")
'@ModuleDescription("Interface describing the development workflow coordinator")
'@depends IPasswords
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Lifecycle
'===============================================================================
'@sub-title Bind the development helper to the Dev worksheet and optional Codes sheet.
'@param devSheet Worksheet hosting development named ranges.
'@param codeSheet Optional worksheet hosting code listobjects.
Public Sub Initialise(ByVal devSheet As Worksheet, Optional ByVal codeSheet As Worksheet)
End Sub

'@section Configuration
'===============================================================================
'@sub-title Retrieve the worksheet holding development named ranges.
Public Property Get DevWorksheet() As Worksheet
End Property

'@sub-title Retrieve the worksheet hosting code listobjects (if any).
Public Property Get CodeWorksheet() As Worksheet
End Property

'@sub-title Control UI prompts when mutating sheets or listobjects.
Public Property Get DisplayPrompts() As Boolean
End Property

Public Property Let DisplayPrompts(ByVal value As Boolean)
End Property

'@section Worksheet preparation
'===============================================================================
'@sub-title Register an existing worksheet that will host code tables.
'@param sheetName Name of the worksheet to register as code worksheet.
'@return Worksheet reference that will host future development tables.
Public Function AddCodeSheets(ByVal sheetName As String) As Worksheet
End Function

'@sub-title Insert a class listobject at the next available anchor.
'@param includeTests When True, create a tests classes table instead of a general one.
Public Function AddClassTable(Optional ByVal includeTests As Boolean = False) As ListObject
End Function

'@sub-title Insert a module listobject at the next available anchor.
'@param includeTests When True, create a tests modules table instead of a general one.
Public Function AddModuleTable(Optional ByVal includeTests As Boolean = False) As ListObject
End Function

'@sub-title Insert a forms mapping listobject at the next available anchor.
Public Function AddFormsTable() As ListObject
End Function

'@sub-title Insert a test listobject at the next available anchor.
Public Function AddTestTable() As ListObject
End Function

'@section Deployment
'===============================================================================
'@sub-title Register a worksheet that should be hidden during deployment.
'@param sheetName Name of the worksheet to hide during deployment.
Public Sub AddHiddenSheet(ByVal sheetName As String)
End Sub

'@sub-title Register a worksheet that should be protected during deployment.
'@param sheetName Name of the worksheet to protect during deployment.
Public Sub AddProtectedSheet(ByVal sheetName As String, _ 
                             Optional ByVal allowShapes As Boolean = True, _ 
                             Optional ByVal allowDeletingRows As Boolean = True)
End Sub

'@sub-title Import all code artefacts from configured folders.
Public Sub ImportAll()
End Sub

'@sub-title Export all code artefacts to configured folders.
Public Sub ExportAll()
End Sub

'@sub-title Move code modules into associated forms according to forms tables.
Public Sub AddFormsCodes()
End Sub

'@sub-title Protect and hide registered worksheets, optionally locking the workbook.
'@param passwords Password handler used to apply protection settings.
'@param lockWorkbook When True, locks the parent workbook after protecting sheets.
Public Sub Deploy(ByVal passwords As IPasswords, Optional ByVal lockWorkbook As Boolean = False)
End Sub

'@section Status
'===============================================================================
'@sub-title Evaluate whether the workbook is already marked as being in deployment.
'@return True when deployment flag equals "Yes".
Public Function InDeployment() As Boolean
End Function