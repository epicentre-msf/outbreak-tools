VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHListVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Interface of HListVars")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods

Public Sub WriteInfo(): End Sub
Public Property Get ValueOf(ByVal colName As String) As String: End Property
