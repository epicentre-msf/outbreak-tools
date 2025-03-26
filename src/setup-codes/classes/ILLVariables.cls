VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of variable class"

'@Folder("Dictionary")
'@ModuleDescription("Interface of variable class")
'@Interface

Option Explicit

'Properties
Public Property Get Dictionary() As ILLdictionary: End Property

'Exposed methods
Public Function Value(ByVal colName As String, ByVal varName As String) As String
End Function

Public Function Contains(ByVal varName As String) As Boolean
End Function

'Change the value of one column for one variable in the dictionary (use carefully it modifies setup data)
Public Sub SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)
End Sub

