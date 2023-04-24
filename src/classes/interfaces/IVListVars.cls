VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVListVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of VlistVars")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Sub WriteInfo(): End Sub
Public Property Get ValueOf(ByVal colName As String) As String: End Property

