VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLchoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Choice class"


'@Folder("Dictionary")
'@ModuleDescription("Interface of Choice class")
'@IgnoreModule ProcedureNotUsed
'@Interface

Option Explicit

'Properties
Public Property Get Wksh() As Worksheet: End Property
Public Property Get StartRow() As Long: End Property

Public Property Get StartColumn() As Long: End Property
Public Property Get ChoiceData() As IDataSheet: End Property

Public Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = True) As Range: End Property
Public Property Get Categories(ByVal choiceName As String) As BetterArray: End Property
Public Property Get ChoiceExists(ByVal choiceName As String) As Boolean: End Property
Public Property Get AllChoices() As BetterArray: End Property

'Exposed methods

'Void methods
'Sort the choices sheet
Public Sub Sort(): End Sub

'Add choice
Public Sub AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray): End Sub

'Add and remove rows
Public Sub AddRows()
End Sub

Public Sub RemoveRows()
End Sub
