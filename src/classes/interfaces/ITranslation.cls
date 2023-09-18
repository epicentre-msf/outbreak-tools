VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods

Public Sub TranslateForm(ByVal UserFrm As Object)
End Sub

Public Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
End Sub

Public Function TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
End Function
