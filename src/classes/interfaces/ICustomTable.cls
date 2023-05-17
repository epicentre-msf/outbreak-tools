VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit


Public Property Get DataRange(ByVal colName As String, Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = False) As Range
End Property
Public Property Get HeaderRange() As Range: End Property
Public Property Get Name() As String: End Property
Public Property Get CellRange(ByVal colName As String, ByVal lineNum As Long) As Range: End Property
Public Property Get Value(ByVal colName As String, ByVal keyName As String) As String: End Property


Public Sub Import(ByVal impTab As Object): End Sub
Public Sub Clean(): End Sub
Public Sub Sort(ByVal els As Variant): End Sub
Public Sub AddRows(Optional ByVal nbRows As Long = 100): End Sub
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0): End Sub
Public Sub SetValidation(ByVal colName As String, _
                           ByVal drop As IDropdownLists, _
                           ByVal dropName As String, _
                           Optional ByVal alertType As String = "info")
End Sub
