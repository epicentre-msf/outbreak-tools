VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSpecsMainStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Test double providing a minimal IMain implementation for LinelistSpecs tests")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements IMain

Private Type TMainStub
    ready As Boolean
    values As Collection
    password As String
    hasTemplate As Boolean
End Type

Private this As TMainStub

'@section Configuration
'===============================================================================
'@description Initialise the stub with an optional readiness flag.
Public Sub Initialise(Optional ByVal readyFlag As Boolean = True)
    this.ready = readyFlag
    Set this.values = New Collection
End Sub

'@description Override the readiness flag exposed through the interface.
Public Sub SetReady(ByVal readyFlag As Boolean)
    this.ready = readyFlag
End Sub

'@description Configure the password returned by the stub.
Public Sub SetPassword(ByVal passwordText As String)
    this.password = passwordText
End Sub

'@description Configure the template availability flag.
Public Sub SetTemplateAvailable(ByVal templateAvailable As Boolean)
    this.hasTemplate = templateAvailable
End Sub

'@description Store a value that can later be retrieved via ValueOf.
Public Sub SetValue(ByVal tagName As String, ByVal valueText As String)
    Dim key As String

    key = NormaliseKey(tagName)
    EnsureStorage

    On Error Resume Next
        this.values.Remove key
    On Error GoTo 0

    this.values.Add valueText, key
End Sub

Private Sub EnsureStorage()
    If this.values Is Nothing Then Set this.values = New Collection
End Sub

Private Function NormaliseKey(ByVal tagName As String) As String
    NormaliseKey = LCase$(Trim$(tagName))
End Function

Private Function FetchValue(ByVal tagName As String) As String
    Dim key As String
    Dim result As String

    key = NormaliseKey(tagName)
    EnsureStorage

    On Error Resume Next
        result = CStr(this.values.Item(key))
    On Error GoTo 0

    FetchValue = result
End Function

'@section IMain implementation
'===============================================================================
Private Property Get IMain_ValueOf(ByVal tagName As String) As String
    IMain_ValueOf = FetchValue(tagName)
End Property

Private Property Get IMain_Ready() As Boolean
    IMain_Ready = this.ready
End Property

Private Property Get IMain_LinelistPassWord() As String
    IMain_LinelistPassWord = this.password
End Property

Private Sub IMain_AddInfo(ByVal trads As IDesTranslation, ByVal msgCode As String, Optional ByVal infoName As String = "edition")
    'No-op for stub
End Sub

Private Sub IMain_CheckReadiness(ByVal trads As IDesTranslation)
    'No-op for stub
End Sub

Private Sub IMain_CheckWarnings(ByVal trads As IDesTranslation, Optional ByVal passversion As String = vbNullString)
    'No-op for stub
End Sub

Private Sub IMain_CheckFileExistence(ByVal trads As IDesTranslation)
    'No-op for stub
End Sub

Private Sub IMain_UpdateStatus(ByVal numValue As Integer)
    'No-op for stub
End Sub

Private Sub IMain_ClearInputRanges(Optional ByVal clearValues As Boolean = False)
    'No-op for stub
End Sub

Private Sub IMain_OpenLL()
    'No-op for stub
End Sub

Private Property Get IMain_HasTemplate() As Boolean
    IMain_HasTemplate = this.hasTemplate
End Property

Private Sub IMain_CheckRibbonExistence(ByVal trads As IDesTranslation)
    'No-op for stub
End Sub

