VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDropdownLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of dropdowns"


'@Folder("Backend")
'@ModuleDescription("Interface of dropdowns")
'@IgnoreModule UseMeaningfulName, HungarianNotation, ProcedureNotUsed

Option Explicit

'Exposed methods
Public Property Get Exists(ByVal listName As String) As Boolean: End Property
Public Property Get LabelRange(ByVal listName As String) As Range: End Property
Public Property Get Items(ByVal listName As String) As BetterArray: End Property

Public Sub Add(ByVal listValues As BetterArray, ByVal listName As String, _ 
               Optional ByVal addLabel As Boolean = False, _
               Optional ByVal counterPrefix As String = vbNullString, _ 
               Optional ByVal tableStyle As String = "TableStyleLight13")
End Sub

Public Sub Remove(ByVal listName As String): End Sub

Public Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                         Optional ByVal alertType As String = "info", _
                         Optional ByVal message As String = vbNullString, _ 
                         Optional ByVal showError As Boolean = True)
End Sub
Public Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String, _ 
                  Optional ByVal pasteAtBottom As Boolean = False)
End Sub
Public Sub ClearList(ByVal listName As String)
End Sub
Public Sub Sort(ByVal listName As String, Optional ByVal sortOrder As Long = xlAscending)
End Sub
Public Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range): End Sub
