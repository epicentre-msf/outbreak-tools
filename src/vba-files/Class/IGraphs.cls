VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Sub Add(): End Sub
Public Sub AddSeries(ByVal rngName As String, ByVal chrtType As String, Optional axisPos As String = vbNullString): End Sub
Public Sub AddLabels(ByVal catName As String, ByVal lblName As String): End Sub
'Maybe format
Public Sub Format(): End Sub
