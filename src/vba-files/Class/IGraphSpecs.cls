VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGraphSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get NumberOfSeries() As Long
End Property

Public Property Get NumberOfGraphs() As Long
End Property

Public Property Get SeriesName(ByVal index As Long) As String
End Property

Public Property Get SeriesType(ByVal index As Long) As String
End Property

Public Property Get SeriesPos(ByVal index As Long) As String
End Property

Public Property Get SeriesLabel(ByVal index As Long) As String
End Property

Public Property Get SeriesColumnLabel(ByVal index As Long) As String
End Property

Public Sub CreateSeries()
End Sub

