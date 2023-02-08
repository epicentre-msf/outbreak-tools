VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of Spatial analysis maangement")
'@IgnoreModule

Option Explicit

'Exposed methods
Public Property Get Exists(ByVal varName As String) As Boolean: End Property
Public Sub Update(): End Sub
Public Property Get TopGeoValue(ByVal adminName As String, ByVal levelValue As Integer, ByVal varName As String, _
   Optional ByVal tabId As String = vbNullString) As String
End Property

Public Property Get TopHFValue(ByVal orderValue As Integer, ByVal inputVarName As String, Optional ByVal tabId As String = vbNullString) As String
End Property
