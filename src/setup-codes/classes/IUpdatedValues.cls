VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUpdatedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of updated values class"

'@Folder("setup file")
'@ModuleDescription("Interface of updated values class")
'@Interface

Option Explicit

'Exposed methods

Public Sub AddColumns(ByVal Lo As ListObject): End Sub
Public Property Get IsUpdated(ByVal colName As String) As Boolean: End Property
Public Sub CheckUpdate(ByVal sh As Worksheet, ByVal Target As Range): End Sub
Public Sub ClearUp(): End Sub
Public Sub DeleteUp(): End Sub
