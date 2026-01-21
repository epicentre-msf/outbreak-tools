VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerTranslationStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Stub implementation of IDesTranslation for designer ribbon tests"

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IDesTranslation for designer ribbon tests.")
'@depends IDesTranslation
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit
Implements IDesTranslation

Private Type TDesignerTranslationStub
    LangRange As Range
    TranslateCalled As Boolean
    TargetSheet As Worksheet
    Messages As Collection
End Type

Private this As TDesignerTranslationStub

Private Sub Class_Initialize()
    Set this.Messages = New Collection
End Sub

Public Sub SetLanguageRange(ByVal langRange As Range)
    Set this.LangRange = langRange
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

'@section IDesTranslation implementation
'===============================================================================
Private Property Get IDesTranslation_TransObject(Optional ByVal trans As Byte = TranslationOfMessages) As ITranslationObject
    Set IDesTranslation_TransObject = Nothing
End Property

Private Sub IDesTranslation_TranslateDesigner(ByVal mainsh As Worksheet)
    this.TranslateCalled = True
    Set this.TargetSheet = mainsh
End Sub

Private Property Get IDesTranslation_LangListRng() As Range
    Set IDesTranslation_LangListRng = this.LangRange
End Property

Private Function IDesTranslation_TranslationMsg(ByVal msgCode As String) As String
    On Error Resume Next
        IDesTranslation_TranslationMsg = CStr(this.Messages.Item(msgCode))
    On Error GoTo 0

    If LenB(IDesTranslation_TranslationMsg) = 0 Then
        IDesTranslation_TranslationMsg = msgCode
    End If
End Function
