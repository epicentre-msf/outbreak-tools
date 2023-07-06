VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IImpRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Import Report"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Interface of Import Report")

Option Explicit

'Exposed methods

'Show the import report
Public Sub Show()
End Sub

'Add some sheets to the report
Public Sub AddSheetsToReport(ByVal scope As Byte, ByVal sheetsList As BetterArray)
End Sub

'Add some variables to the report
Public Sub AddVariablesToReport(ByVal scope As Byte, _ 
                                ByVal sheetName As String, _ 
                                ByVal varsList As BetterArray)
End Sub

Public Sub Clear()
End Sub

Public Property Get ContainsSheet(ByVal scope As Byte, ByVal sheetName As String) As Boolean
End Property
  
Public Property Get ContainsVariable(ByVal scope As Byte, _ 
                                     ByVal sheetName As String, _ 
                                     ByVal varName As String) As Boolean
End Property