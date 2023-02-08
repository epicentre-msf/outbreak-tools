VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDropdownLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get Exists(ByVal listName As String) As Boolean: End Property
Public Sub Add(ByVal listValues As BetterArray, ByVal listName As String): End Sub

Public Sub Remove(ByVal listName As String): End Sub
Public Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                         Optional ByVal alertType As String = "info", _
                         Optional ByVal message As String = vbNullString)
End Sub


Public Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String)
End Sub
