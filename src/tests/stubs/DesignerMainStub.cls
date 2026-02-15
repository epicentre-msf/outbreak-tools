VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerMainStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Stub implementation of IMain for designer ribbon tests"

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IMain for designer ribbon tests.")
'@depends IMain
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, UseMeaningfulName, SuperfluousAnnotationArgument

Option Explicit
Implements IMain

Private Type TDesignerMainStub
    Values As Collection
    ReadyState As Boolean
    Password As String
    ClearCalled As Boolean
    ClearWithValues As Boolean
    TemplateAvailable As Boolean
    LastStatus As Integer
End Type

Private this As TDesignerMainStub

Private Sub Class_Initialize()
    Set this.Values = New Collection
    this.ReadyState = True
End Sub

Public Sub SetValue(ByVal tagName As String, ByVal value As String)
    On Error Resume Next
        this.Values.Remove tagName
    On Error GoTo 0
    this.Values.Add value, tagName
End Sub

Public Sub SetReady(ByVal isReady As Boolean)
    this.ReadyState = isReady
End Sub

Public Sub SetPassword(ByVal passwordValue As String)
    this.Password = passwordValue
End Sub

Public Sub SetTemplateAvailable(ByVal hasTemplate As Boolean)
    this.TemplateAvailable = hasTemplate
End Sub

Public Property Get ClearRequested() As Boolean
    ClearRequested = this.ClearCalled
End Property

Public Property Get ClearedWithValues() As Boolean
    ClearedWithValues = this.ClearWithValues
End Property

'@section IMain implementation
'===============================================================================
Private Property Get IMain_ValueOf(ByVal tagName As String) As String
    On Error Resume Next
        IMain_ValueOf = CStr(this.Values.Item(tagName))
    On Error GoTo 0
End Property

Private Property Get IMain_Ready() As Boolean
    IMain_Ready = this.ReadyState
End Property

Private Property Get IMain_LinelistPassWord() As String
    IMain_LinelistPassWord = this.Password
End Property

Private Sub IMain_AddInfo(ByVal trads As IDesignerTranslation, ByVal msgCode As String, Optional ByVal infoName As String = "edition")
    'No-op for stub
End Sub

Private Sub IMain_CheckReadiness(ByVal trads As IDesignerTranslation)
    'No-op for stub
End Sub

Private Sub IMain_CheckWarnings(ByVal trads As IDesignerTranslation, Optional ByVal passversion As String = vbNullString)
    'No-op for stub
End Sub

Private Sub IMain_CheckFileExistence(ByVal trads As IDesignerTranslation)
    'No-op for stub
End Sub

Private Sub IMain_UpdateStatus(ByVal numValue As Integer)
    this.LastStatus = numValue
End Sub

Private Sub IMain_ClearInputRanges(Optional ByVal clearValues As Boolean = False)
    this.ClearCalled = True
    this.ClearWithValues = clearValues
End Sub

Private Sub IMain_OpenLL()
    'No-op for stub
End Sub

Private Property Get IMain_HasTemplate() As Boolean
    IMain_HasTemplate = this.TemplateAvailable
End Property

Private Sub IMain_CheckRibbonExistence(ByVal trads As IDesignerTranslation)
    'No-op for stub
End Sub
