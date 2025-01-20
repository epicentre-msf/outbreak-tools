VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of linelist sheets"

'@Folder("Dictionary")
'@ModuleDescription("Interface of linelist sheets")
'@Interface
'@IgnoreModule ProcedureNotUsed

Option Explicit

Public Enum SheetBound
    RowSart = 1
    RowEnd = 2
    ColStart = 3
    ColEnd = 4
End Enum

Public Enum SheetInfoType
    SheetInfoSheetType = 1
    SheetInfoSheetTable = 2
End Enum

'Properties

Public Property Get Dictionary() As ILLdictionary: End Property

'Exposed methods (Seems Handy to use sheetype and table instead of one function)
Public Function SheetInfo(ByVal sheetName As String, Optional ByVal info As Byte = SheetInfoType.SheetInfoSheetType) As String: End Function

Public Function Contains(ByVal sheetName As String) As Boolean: End Function

Public Function DataBounds(ByVal sheetName As String, Optional ByVal Bound As Byte = SheetBound.RowSart) As Long: End Function

Public Function ContainsControl(ByVal sheetName As String, Optional ByVal control As String = "list_auto") As Boolean: End Function
Public Function VariableAddress(ByVal varName As String, Optional ByVal onSheet As String = vbNullString) As String: End Function

Public Function RowIndex(ByVal sheetName As String) As Long: End Function
Public Function NumberOfVars(ByVal sheetName As String) As Long: End Function

