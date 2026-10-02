VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistPasswordStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub IPasswords implementation for save workflow tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements IPasswords

Private Type TPassword
    lastProtectedWorkbook As Workbook
    protectCount As Long
End Type

Private this As TPassword

Public Property Get LastProtectedWorkbook() As Workbook
    Set LastProtectedWorkbook = this.lastProtectedWorkbook
End Property

Public Property Get ProtectCount() As Long
    ProtectCount = this.protectCount
End Property

Private Sub RecordProtection(ByVal target As Variant)
    this.protectCount = this.protectCount + 1

    If IsObject(target) Then
        Select Case TypeName(target)
            Case "Workbook"
                Set this.lastProtectedWorkbook = target
            Case "Worksheet"
                Set this.lastProtectedWorkbook = target.Parent
        End Select
    End If
End Sub



'@prop-title Password worksheet implementation
Private Property Get IPasswords_PasswordSheet() As Worksheet
End Property

'@prop-title Named ranges values implementation
Private Property Get IPasswords_Value(ByVal tagName As String) As String
    If tagName = "debugginpassword" Then IPasswords_Value = "1234"
End Property

'@prop-title Prompt display flag implementation
Private Property Get IPasswords_DisplayPrompts() As Boolean
End Property

Private Property Let IPasswords_DisplayPrompts(ByVal value As Boolean)
End Property

'@prop-title Last private prompt text implementation
Private Property Get IPasswords_LastPrivatePrompt() As String
End Property

Private Property Get IPasswords_LastPrivatePromptTitle() As String
End Property

'@prop-title Table Range
Private Property Get IPasswords_TableRange(ByVal tableName As String, Optional ByVal includeHeaders As Boolean = False) As Range
End Property

'@prop-title Checkings implementation
Private Property Get IPasswords_HasCheckings() As Boolean
End Property

'@prop-title Debug Mode?
Private Property Get IPasswords_IsInDebugMode() As Boolean
    IPasswords_IsInDebugMode = False
End Property

'@prop-title Checkings values implementation
Private Property Get IPasswords_CheckingValues() As IChecking
End Property

'@sub-title Enter debug mode impl.
Private Sub IPasswords_EnterDebugMode(Optional ByVal targetWorkbook As Workbook)
End Sub

'@sub-title Leave debug mode impl.
Private Sub IPasswords_LeaveDebugMode(Optional ByVal targetWorkbook As Workbook)
End Sub

'@sub-title Inject before-close impl.
Private Sub IPasswords_EnsureDebugExitHandler(Optional ByVal targetWorkbook As Workbook)
End Sub

'@sub-title Protection impl.
Private Sub IPasswords_Protect(Optional targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True)
    RecordProtection targetObject
End Sub

'@sub-title UnProtection impl.
Private Sub IPasswords_UnProtect(Optional targetObject As Variant = "_active")
End Sub

'@sub-title Ensure protected sheet names impl.
Private Sub IPasswords_EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)
End Sub

'@sub-title Apply protection matrix impl.
Private Sub IPasswords_ApplyProtectionMatrix(ByVal wb As Workbook)
End Sub

'@sub-title Generate key impl.
Private Sub IPasswords_GenerateKey(ByVal trads As ITranslationObject)
End Sub

'@sub-title Display key impl.
Private Sub IPasswords_DisplayPrivateKey(ByVal trads As ITranslationObject)
End Sub

'@sub-title Export to workbook impl
Private Sub IPasswords_ExportToWorkbook(ByVal wb As Workbook)
End Sub

'@sub-title Import from pass impl.
Private Sub IPasswords_ImportFrom(ByVal source As IPasswords)
End Sub

'@fun-title Clone to worksheet impl.
Private Function IPasswords_CloneToWorksheet(ByVal targetSheet As Worksheet) As IPasswords
End Function

'@fun-title Clone to workbook impl.
Private Function IPasswords_CloneToWorkbook(ByVal wb As Workbook) As IPasswords
End Function

'@sub-title New debugging password impl.
Private Sub IPasswords_NewDebugingPassword(ByVal newValue As String)
End Sub
