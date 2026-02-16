VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSpecsDesignerTranslationStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Test double exposing a minimal IDesignerTranslation implementation")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDesignerTranslation

Private translationObj As ITranslationObject

'@description Supply the translation object returned for every scope.
Public Sub UseTranslationObject(ByVal translation As ITranslationObject)
    Set translationObj = translation
End Sub

Private Function ResolveTranslationObject() As ITranslationObject
    If translationObj Is Nothing Then
        Dim stub As LinelistSpecsTranslationStub
        Set stub = New LinelistSpecsTranslationStub
        stub.Initialise
        Set translationObj = stub
    End If

    Set ResolveTranslationObject = translationObj
End Function

'@section IDesignerTranslation implementation
'===============================================================================
Private Property Get IDesignerTranslation_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set IDesignerTranslation_TransObject = ResolveTranslationObject()
End Property

Private Sub IDesignerTranslation_TranslateDesigner(ByVal mainsh As Worksheet, ByVal language As String)
    'No-op for stub
End Sub

Private Function IDesignerTranslation_TranslatedValue(ByVal msgCode As String) As String
    IDesignerTranslation_TranslatedValue = ResolveTranslationObject().TranslatedValue(msgCode)
End Function
