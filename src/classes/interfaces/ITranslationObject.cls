VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITranslationObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of translation object"

'@Interface
'@Folder("GeneralInterface")
'@ModuleDescription("Interface of translation object")
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods

Public Sub TranslateForm(ByVal UserFrm As Object)
End Sub

Public Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
End Sub

Public Function TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
End Function
