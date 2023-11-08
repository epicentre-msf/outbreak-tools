VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLShowHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Show/Hide form"


'@Folder("Linelist")
'@ModuleDescription("Interface of Show/Hide form")
'@Interface

Option Explicit

Public Enum ShowHideScope
  ShowHideScopeHList = 1
  ShowHideScopePrint = 2
  ShowHideScopeVList = 3
End Enum

'Exposed methods

'Load the show/hide form with values
Public Sub Load(ByVal tradsform As ITranslation, Optional ByVal showForm As Boolean = True, Optional ByVal showOptional As Boolean = False)
End Sub
'Retrieve the ShowHide table Range
Public Property Get ShowHideTable() As ICustomTable
End Property
'Initialize show/hide table, without loading the form
Public Sub InitializeTable():End Sub
Public Sub UpdateVisibilityStatus(ByVal Index As Long): End Sub
Public Sub ShowHideLogic(ByVal Index As Long): End Sub
Public Sub ChangeColWidth(ByVal Index As Long): End Sub
