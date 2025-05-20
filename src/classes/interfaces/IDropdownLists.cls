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

'@Folder("GeneralInterface")
'@ModuleDescription("Interface of dropdowns")
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

'@jump:Create
Option Explicit

'@section Dropdown Elements

'@jump:Exists
Public Property Get Exists(ByVal listName As String) As Boolean
End Property

'@jump:Name
Public Property Get Name() As String
End Property

'@jump:LabelRange
Public Property Get LabelRange(ByVal listName As String) As Range
End Property

'@section Modify the dropdown list

'@jump:Add
'Add a new dropdown
Public Sub Add(ByVal listValues As BetterArray, ByVal listName As String, _ 
               Optional ByVal addLabel As Boolean = False, _
               Optional ByVal counterPrefix As String = vbNullString, _ 
               Optional ByVal tableStyle As String = "TableStyleLight13")
End Sub

'@jump:Remove
'Remove a dropdown from the worksheet, including references
Public Sub Remove(ByVal listName As String): End Sub

'@jump:ClearList
'Clear entries of a dropdow
Public Sub ClearList(ByVal listName As String)
End Sub

'@jump:Update
'Update a dropdown with new values
Public Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String, _ 
                  Optional ByVal pasteAtBottom As Boolean = False)
End Sub

'@jump:Sort
'Sort a dropdown in the worksheet.
Public Sub Sort(ByVal listName As String, Optional ByVal sortOrder As Long = xlAscending)
End Sub

'@section Operate with the outside world

'@jump:SetValidation
Public Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                         Optional ByVal alertType As String = "info", _
                         Optional ByVal message As String = vbNullString, _ 
                         Optional ByVal showError As Boolean = True)
End Sub

'@jump:AddReturnLink
Public Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
End Sub

'@jump:AddHyperLink
Public Sub AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
End Sub

'@section checkings

'@jump:HasCheckings
'Test if the object has checkings enable
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'Get the checking values
Public Property Get CheckingValues() As Object: End Property