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
'@ModuleDescription("Test double exposing a minimal IDesTranslation implementation")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDesTranslation

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

'@section IDesTranslation implementation
'===============================================================================
Private Property Get IDesTranslation_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set IDesTranslation_TransObject = ResolveTranslationObject()
End Property

Private Sub IDesTranslation_TranslateDesigner(ByVal mainsh As Worksheet)
    'No-op for stub
End Sub

Private Property Get IDesTranslation_LangListRng() As Range
    'No language list involved for the stub implementation
End Property

Private Function IDesTranslation_TranslationMsg(ByVal msgCode As String) As String
    IDesTranslation_TranslationMsg = ResolveTranslationObject().TranslatedValue(msgCode)
End Function

