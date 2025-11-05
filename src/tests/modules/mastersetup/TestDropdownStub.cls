VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDropdownStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Lightweight IDropdownLists stub for unit testing"

Option Explicit
Implements IDropdownLists

Private Const STUB_NAME As String = "TestDropdowns"

Private Function NotSupported() As Variant
    Err.Raise vbObjectError + 2048, STUB_NAME, "Method not supported in test stub"
End Function

Private Property Get IDropdownLists_Exists(ByVal listName As String) As Boolean
    IDropdownLists_Exists = True
End Property

Private Property Get IDropdownLists_Name() As String
    IDropdownLists_Name = STUB_NAME
End Property

Private Property Get IDropdownLists_LabelRange(ByVal listName As String) As Range
    NotSupported
End Property

Private Property Get IDropdownLists_Values(ByVal listName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
    NotSupported
End Property

Private Property Get IDropdownLists_Length() As Long
    IDropdownLists_Length = 0
End Property

Private Sub IDropdownLists_Add(ByVal listValues As BetterArray, ByVal listName As String, Optional ByVal addLabel As Boolean = False, Optional ByVal counterPrefix As String = vbNullString, Optional ByVal tableStyle As String = "TableStyleLight13")
    NotSupported
End Sub

Private Sub IDropdownLists_Remove(ByVal listName As String)
    NotSupported
End Sub

Private Sub IDropdownLists_ClearList(ByVal listName As String)
    NotSupported
End Sub

Private Sub IDropdownLists_Update(ByVal UpdateData As BetterArray, ByVal listName As String, Optional ByVal pasteAtBottom As Boolean = False, Optional ByVal removeDuplicates As Boolean = True)
    NotSupported
End Sub

Private Sub IDropdownLists_Sort(ByVal listName As String, Optional ByVal sortOrder As Long = xlAscending)
    NotSupported
End Sub

Private Sub IDropdownLists_SetValidation(ByVal cellRng As Range, ByVal listName As String, Optional ByVal alertType As String = "info", Optional ByVal message As String = vbNullString, Optional ByVal showError As Boolean = True)
    With cellRng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="= " & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = message
        .ShowInput = False
        .ShowError = showError
    End With
End Sub

Private Sub IDropdownLists_AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    NotSupported
End Sub

Private Sub IDropdownLists_AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
    NotSupported
End Sub

Private Property Get IDropdownLists_HasCheckings() As Boolean
    IDropdownLists_HasCheckings = False
End Property

Private Property Get IDropdownLists_CheckingValues() As Object
    Set IDropdownLists_CheckingValues = Nothing
End Property
