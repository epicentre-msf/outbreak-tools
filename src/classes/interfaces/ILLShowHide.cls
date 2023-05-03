VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ILLShowHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Linelist")
'@ModuleDescription("Interface of Show/Hide form")
'@Interface

Option Explicit

Public Enum ShowHideScope
  ShowHideScopeLinelist = 1
  ShowHideScopePrint = 2
End Enum

'Exposed methods

'Load the show/hide form with values
Public Sub Load(): End Sub
Public Sub UpdateVisibilityStatus(ByVal index As Long): End Sub
Public Sub ShowHideLogic(ByVal index As Long): End Sub