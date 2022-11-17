VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist Passwords")
'@ModuleDescription("Manage passwords in the linelist")
'@IgnoreModule

Option Explicit

'Exposed methods
Implements ILLPasswords
Private Const CLASSNAME As String = "LLPasswords"


Public Function Create(ByVal sh As Worksheet)
  CheckRequirements sh
End Function



'Errors

Private Sub CheckRequirements(ByVal sh As Worksheet)
  
End Sub
