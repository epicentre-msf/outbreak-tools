VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IUpVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of updated values"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Interface of updated values")

Option Explicit

'Set and Get updated values
Public Property Get Value(ByVal rngName As String) As String
End Property

Public Sub SetValue(ByVal rngName As String, ByVal newValue As String)
End Sub
