VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistMainStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Simple IMain stub for save workflow tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements IMain

Private Type TMain
    values As Collection
End Type

Private this As TMain

Public Sub Initialise()
    Set this.values = New Collection
End Sub

Public Sub SetValue(ByVal key As String, ByVal value As String)
    Dim normalized As String

    normalized = LCase$(Trim$(key))
    EnsureStorage

    On Error Resume Next
        this.values.Remove normalized
    On Error GoTo 0

    this.values.Add value, normalized
End Sub

Private Sub EnsureStorage()
    If this.values Is Nothing Then Set this.values = New Collection
End Sub

Private Function IMain_ValueExists(ByVal Key As String) As Boolean
    Dim normalized As String
    Dim dummy As String

    normalized = LCase$(Trim$(Key))
    EnsureStorage

    On Error Resume Next
        dummy = CStr(this.values.Item(normalized))
    On Error GoTo 0

    IMain_ValueExists = LenB(dummy) > 0
End Function

Private Function IMain_ValueOf(ByVal Key As String) As String
    Dim normalized As String
    Dim value As String

    normalized = LCase$(Trim$(Key))
    EnsureStorage

    On Error Resume Next
        value = CStr(this.values.Item(normalized))
    On Error GoTo 0

    IMain_ValueOf = value
End Function

Private Sub IMain_SetValue(ByVal Key As String, ByVal Value As String)
    SetValue Key, Value
End Sub

Private Sub IMain_Refresh()
End Sub

Private Sub IMain_Save()
End Sub

Private Sub IMain_Reset()
End Sub

Private Sub IMain_ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub

Private Property Get IMain_Workbook() As Workbook
    Set IMain_Workbook = Nothing
End Property

Private Property Get IMain_TextExists(ByVal text As String) As Boolean
    IMain_TextExists = IMain_ValueExists(text)
End Property

Private Property Get IMain_ValueOfScope(ByVal Key As String, Optional ByVal scope As Byte = 0) As String
    IMain_ValueOfScope = IMain_ValueOf(Key)
End Property

Private Property Get IMain_ValueExistsScope(ByVal Key As String, Optional ByVal scope As Byte = 0) As Boolean
    IMain_ValueExistsScope = IMain_ValueExists(Key)
End Property

Private Property Get IMain_Values() As Collection
    EnsureStorage
    Set IMain_Values = this.values
End Property

Private Property Get IMain_StateExists(ByVal stateKey As String) As Boolean
    IMain_StateExists = False
End Property

Private Property Get IMain_StateValue(ByVal stateKey As String) As Variant
    IMain_StateValue = Empty
End Property

Private Property Get IMain_SettingsValue(ByVal Key As String) As Variant
    IMain_SettingsValue = Empty
End Property

Private Sub IMain_SetStateValue(ByVal stateKey As String, ByVal Value As Variant)
End Sub

Private Sub IMain_Flush()
End Sub

Private Property Get IMain_ValueOrDefault(ByVal Key As String, Optional ByVal defaultValue As String = "") As String
    Dim value As String
    value = IMain_ValueOf(Key)
    If LenB(value) = 0 Then
        IMain_ValueOrDefault = defaultValue
    Else
        IMain_ValueOrDefault = value
    End If
End Property

Private Property Get IMain_ValueOfPointer(ByVal Key As String) As String
    IMain_ValueOfPointer = vbNullString
End Property

Private Sub IMain_SetValuePointer(ByVal Key As String, ByVal Value As String)
End Sub

Private Property Get IMain_TranslatedValue(ByVal Key As String) As String
    IMain_TranslatedValue = IMain_ValueOf(Key)
End Property

Private Sub IMain_SetTranslatedValue(ByVal Key As String, ByVal Value As String)
End Sub

Private Property Get IMain_DictOf(ByVal Key As String) As Object
    Set IMain_DictOf = Nothing
End Property

Private Sub IMain_SetDictOf(ByVal Key As String, ByVal Value As Object)
End Sub

Private Property Get IMain_ArrayOf(ByVal Key As String) As Variant
    IMain_ArrayOf = Empty
End Property

Private Sub IMain_SetArrayOf(ByVal Key As String, ByVal Value As Variant)
End Sub

