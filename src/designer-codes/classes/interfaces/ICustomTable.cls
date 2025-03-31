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
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit


Public Property Get DataRange(Optional ByVal colName As String = "__all__", _
                              Optional ByVal includeHeaders As Boolean = False, _
                              Optional ByVal strictSearch As Boolean = False) _
                              As Range
End Property
Public Property Get HeaderRange() As Range: End Property
Public Property Get Name() As String: End Property
Public Property Get CellRange(ByVal colName As String, ByVal lineNum As Long) As Range: End Property
Public Property Get Value(ByVal colName As String, ByVal keyName As String) As String: End Property


Public Sub Import(ByVal impTab As Object, _
                  Optional ByVal pasteAtBottom As Boolean = False, _
                  Optional ByVal strictColumnSearch As Boolean = False, _ 
                  Optional ByVal insertShift As Boolean = True)
End Sub
Public Sub Clean(): End Sub
Public Sub Sort(ByVal els As Variant, Optional ByVal onFirst As Boolean = True): End Sub
Public Sub AddRows(Optional ByVal nbRows As Long = 100, _ 
                   Optional ByVal insertShift As Boolean = False)
End Sub
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0): End Sub
Public Sub SetValidation(ByVal colName As String, _
                         ByVal drop As IDropdownLists, _
                         ByVal dropName As String, _
                         Optional ByVal alertType As String = "info")
End Sub

'Use very wisely, could modify the custom table values
Public Sub SetValue(ByVal colName As String, ByVal keyName As String, ByVal newValue As String)
End Sub

Public Property Get ImportColumnsNotFound() As BetterArray
End Property

'Test if there are imports in the custom table
Public Property Get HasColumnsNotImported() As Boolean
End Property
  
Public Sub Export(ByVal sh As Worksheet, ByVal headersList As BetterArray, Optional ByVal startLine As Long = 1)
End Sub
