VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of passwords"

'@Folder("setup file")
'@ModuleDescription("Interface of passwords")
'@Interface

Option Explicit

'Exposed methods
Public Property Get Password() As String: End Property

Public Sub UnProtect(Optional ByVal sheetName As String = "_active"): End Sub

Public Sub Protect(Optional ByVal sheetName As String = "_active", _
                   Optional ByVal strictProtect As Boolean = False, _
                   Optional ByVal deleteColumns As Boolean = False)
End Sub

'Public Sub EnterDebugMode(): End Sub

'Public Sub LeaveDebugMode(): End Sub


'Protect and unprotect a given workbook to avoid sheet name changes.
Public Sub ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
End Sub

Public Sub UnProtectWkb(ByVal wb As Workbook)
End Sub

