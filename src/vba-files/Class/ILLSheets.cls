VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of Linelist sheets")
'@IgnoreModule

Option Explicit

Public Enum SheetBound
  RowSart = 1
  RowEnd = 2
  ColStart = 3
  ColEnd = 4
End Enum

'Properties

Public Property Get Dictionary As ILLdictionary: End Property
Public Property Get AllSheets As BetterArray : End Property

'Exposed methods (Seems Handy to use sheetype and table instead of one function)
Public Function SheetType(ByVal sheetName As String) As String: End Function
Public Function Contains(ByVal sheetName As String) As Boolean: End Function
Public Function Table(ByVal sheetName As String) As String : End Function
Public Function DataBounds(ByVal sheetName As String, Optional ByVal Bound As SheetBound = 1) As Long: End Function
Public Function VarList(ByVal sheetName As String) As BetterArray : End Function
Public Function ContainsListAuto(ByVal sheetName As String) As Boolean: End Function
Public Function VariableAddress(Byval sheetName As String, ByVal varName As String, Optional ByVal useTable As Boolean = False) As String : End Function

