VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of Spatial analysis maangement")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
Public Property Get Exists(ByVal varName As String) As Boolean: End Property
Public Sub Update(): End Sub
Public Property Get TopGeoValue(ByVal adminName As String, ByVal levelValue As Integer, ByVal varName As String, _
   Optional ByVal tabId As String = vbNullString, Optional ByVal concatValue As Byte = 1) As String
End Property

Public Property Get TopHFValue(ByVal orderValue As Integer, _
                               ByVal inputVarName As String, _
                               Optional ByVal tabId As String = vbNullString) As String
End Property

Public Sub Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False)
End Sub
