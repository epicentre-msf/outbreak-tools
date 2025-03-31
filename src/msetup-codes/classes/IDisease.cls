VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Disease class"

'@Folder("Dictionary")
'@ModuleDescription("Interface of Disease class")
'@IgnoreModule ProcedureNotUsed
'@Interface

Option Explicit

'Exposed methods
Public Property Get Exists(ByVal disName As String) As Boolean
End Property

Public Property Get OutputWkb() As Workbook
End Property

Public Sub Add(ByVal disName As String,  _ 
               Optional ByVal disLang As String = vbNullString)
End Sub

Public Sub Remove(ByVal disName As String)
End Sub

Public Sub ExportDisease(ByVal disName As String)
End Sub

Public Sub ExportForMigration()
End Sub
