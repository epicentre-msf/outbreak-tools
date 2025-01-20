VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of CustomTable"

'@Folder("General")
'@ModuleDescription("Interface of CustomTable")
'@Interface

Option Explicit

Private Const NBROWSADD As Long = 5 'Default Number of Rows to add when resizing


Public Property Get DataRange(ByVal colName As String, _ 
                              Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = False) _ 
                              As Range
End Property
Public Property Get HeaderRange() As Range: End Property
Public Property Get Name() As String: End Property
Public Property Get CellRange(ByVal colName As String, _ 
                              ByVal lineNum As Long) As Range
End Property
Public Property Get Value(ByVal colName As String, _
                          ByVal keyName As String) As String
End Property


Public Sub Import(ByVal impTab As ICustomTable, _ 
                  Optional ByVal keepImportHeaders As Boolean = False)
End Sub
Public Sub Clean(ByVal excludeNames As BetterArray): End Sub
Public Sub Sort(ByVal els As Variant): End Sub
Public Sub AddRows(Optional ByVal nbRows As Long = NBROWSADD): End Sub
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0): End Sub
Public Sub SetValidation(ByVal colName As String, _
                           ByVal drop As IDropdownLists, _
                           ByVal dropName As String, _
                           Optional ByVal alertType As String = "info")
End Sub
