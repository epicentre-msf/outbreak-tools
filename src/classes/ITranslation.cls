VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods

Public Sub TranslateForm(ByVal UserFrm As UserForm)
End Sub

Public Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
End Sub

Public Function TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
End Function

