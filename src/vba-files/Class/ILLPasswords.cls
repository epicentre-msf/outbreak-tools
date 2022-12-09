VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Interface of Password")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get Password() As String: End Property

Public Sub UnProtect(Optional ByVal sheetName As String = "_active"): End Sub

Public Sub Protect(Optional ByVal sheetName As String = "_active", Optional Byval allowShapes As Boolean = False)
End Sub

'Public Sub EnterDebugMode(): End Sub

'Public Sub LeaveDebugMode(): End Sub

'Export password
Public Sub Export(ByVal toWkb As Workbook): End Sub

