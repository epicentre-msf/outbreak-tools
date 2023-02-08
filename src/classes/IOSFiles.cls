VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOSFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule

Option Explicit

'Exposed methods

Public Property Get File() As String: End Property
Public Property Get Folder() As String: End Property

Public Sub LoadFile(ByVal filters As String): End Sub

Public Sub LoadFolder(): End Sub

Public Function HasValidFile() As Boolean: End Function

Public Function HasValidFolder() As Boolean: End Function

