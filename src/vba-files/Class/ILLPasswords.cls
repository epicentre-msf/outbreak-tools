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
Public Property Get PassWord() As String: End Property

Public Sub UnProtect(ByVal sh As Worksheet): End Sub
Public Sub Protect(ByVal sh As Worksheet): End Sub
Public Sub EnterDebugMode(): End Sub
Public Sub LeaveDebugMode(): End Sub

'Export password