VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of Variable class")
'@IgnoreModule

Option Explicit

'Properties
Private Property Get Dictionary() As ILLdictionary: End Property


'Exposed methods
Public Function Value(ByVal colName As String, ByVal varName As String) As String
End Function

'column or row index in the sheet of one variable
Public Function index(ByVal varName As String) As Long
End Function

Public Function Contains(ByVal varName As String) As Boolean
End Function
