VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract describing the linelist workbook facade"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Contract describing the linelist workbook facade")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Worksheets
'===============================================================================

'@description Return a worksheet from the output workbook by name.
Public Property Get Wksh(ByVal sheetName As String) As Worksheet
End Property

'@description Return the printed companion worksheet (print_ prefix) by base name.
Public Property Get PrintWksh(ByVal sheetName As String) As Worksheet
End Property

'@description Return the CRF companion worksheet (crf_ prefix) by base name.
Public Property Get CRFWksh(ByVal sheetName As String) As Worksheet
End Property

'@description Check whether a sheet exists in the output workbook.
Public Function SheetExists(ByVal sheetName As String) As Boolean
End Function

'@description Add a worksheet to the output workbook with optional visibility and scope.
Public Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal sheetScope As Byte = 1)
End Sub

'@section Collaborators
'===============================================================================

'@description Return the dropdown lists manager for the output workbook.
Public Property Get Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
End Property

'@description Return the pivot tables manager for the output workbook.
Public Property Get Pivots() As ICustomPivotTables
End Property

'@description Return the specifications backing this linelist.
Public Property Get LinelistData() As ILinelistSpecs
End Property

'@description Return the dictionary from the specifications.
Public Property Get Dictionary() As ILLdictionary
End Property

'@section Lifecycle
'===============================================================================

'@description Build the full linelist output workbook from specifications.
Public Sub Prepare()
End Sub

'@description Save the output workbook as .xlsb with password protection.
Public Sub SaveLL()
End Sub

'@description Display an error dialog for incomplete linelist operations.
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub
