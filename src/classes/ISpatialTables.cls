VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISpatialTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get Exists(ByVal varName As String) As Boolean: End Property
Public Sub Add(ByVal formObject As IFormulas): End Sub
