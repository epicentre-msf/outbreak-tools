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

'TranslationObject is a class for translation of elements such 
'as ranges, values including formulas. Do not confuse with the 
'translation class of the setup or other translation elements. This
'one is only for translating chunks.

'@jump:Name
Public Property Get Name() As String
End Property

'@jump:ValueExists
Public Property Get ValueExists(ByVal text As String) As Boolean
End Property


'@jump:TranslateForm
Public Sub TranslateForm(ByVal UserFrm As Object)
End Sub

'@jump:TranslateRange
Public Sub TranslateRange(ByVal colRng As Range, Optional ByVal containsFormula As Boolean = False)
End Sub

'@jump:TranslatedValue
Public Function TranslatedValue(ByVal Text As String, Optional ByVal containsFormula As Boolean = False) As String
End Function


'@section checkings

'@jump:HasCheckings
'Test if the object has checkings enable
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'Get the checking values
Public Property Get CheckingValues() As Object: End Property