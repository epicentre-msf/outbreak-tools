VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ICustomPivotTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of CustomPivotTable"

'@Interface
'@Folder("Linelist Building")
'@ModuleDescription("Interface of CustomPivotTable")

Option Explicit

'Exposed methods
Public Sub Add(ByVal title As String, _ 
               ByVal tablename As String, _ 
               Optional ByVal pivotName As String = "Pivot Table")
End Sub