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

Private Sub RecordProtection(ByVal workbook As Workbook)
    Set this.lastProtectedWorkbook = workbook
    this.protectCount = this.protectCount + 1
End Sub

Private Sub IPasswords_Protect(ByVal workbook As Workbook)
    RecordProtection workbook
End Sub

Private Sub IPasswords_ProtectWkb(ByVal workbook As Workbook)
    RecordProtection workbook
End Sub

Private Sub IPasswords_Unprotect(ByVal workbook As Workbook)
End Sub

Private Sub IPasswords_UnprotectWkb(ByVal workbook As Workbook)
End Sub

Private Function IPasswords_IsProtected(ByVal workbook As Workbook) As Boolean
    IPasswords_IsProtected = False
End Function

Private Function IPasswords_IsWorkbookProtected(ByVal workbook As Workbook) As Boolean
    IPasswords_IsWorkbookProtected = False
End Function

Private Property Get IPasswords_Password() As String
    IPasswords_Password = ""
End Property

Private Property Let IPasswords_Password(ByVal value As String)
End Property

Private Property Get IPasswords_RestrictionPassword() As String
    IPasswords_RestrictionPassword = ""
End Property

Private Property Let IPasswords_RestrictionPassword(ByVal value As String)
End Property

Private Property Get IPasswords_UnprotectPassword() As String
    IPasswords_UnprotectPassword = ""
End Property

Private Property Let IPasswords_UnprotectPassword(ByVal value As String)
End Property

