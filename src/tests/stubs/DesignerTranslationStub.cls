VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerTranslationStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Stub implementation of IDesignerTranslation for designer ribbon tests"

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IDesignerTranslation for designer ribbon tests.")
'@depends IDesignerTranslation
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit
Implements IDesignerTranslation

Private Type TDesignerTranslationStub
    TranslateCalled As Boolean
    TargetSheet As Worksheet
    Language As String
    Messages As Collection
End Type

Private this As TDesignerTranslationStub

Private Sub Class_Initialize()
    Set this.Messages = New Collection
End Sub

Public Sub SetMessage(ByVal msgCode As String, ByVal messageText As String)
    On Error Resume Next
        this.Messages.Remove msgCode
    On Error GoTo 0
    this.Messages.Add messageText, msgCode
End Sub

Public Property Get TranslateRequested() As Boolean
    TranslateRequested = this.TranslateCalled
End Property

Public Property Get TargetSheet() As Worksheet
    Set TargetSheet = this.TargetSheet
End Property

Public Property Get LastLanguage() As String
    LastLanguage = this.Language
End Property

'@section IDesignerTranslation implementation
'===============================================================================
Private Property Get IDesignerTranslation_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set IDesignerTranslation_TransObject = Nothing
End Property

Private Sub IDesignerTranslation_TranslateDesigner(ByVal mainsh As Worksheet, ByVal language As String)
    this.TranslateCalled = True
    Set this.TargetSheet = mainsh
    this.Language = language
End Sub

Private Function IDesignerTranslation_TranslatedValue(ByVal msgCode As String) As String
    On Error Resume Next
        IDesignerTranslation_TranslatedValue = CStr(this.Messages.Item(msgCode))
    On Error GoTo 0

    If LenB(IDesignerTranslation_TranslatedValue) = 0 Then
        IDesignerTranslation_TranslatedValue = msgCode
    End If
End Function
