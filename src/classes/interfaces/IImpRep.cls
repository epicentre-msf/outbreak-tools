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
Public Sub AddSheetsToReport(ByVal scope As Byte, ByVal sheetName As String)
End Sub

'Add some variables to the report
Public Sub AddVariablesToReport(ByVal scope As Byte, _ 
                                ByVal sheetName As String, _
                                ByVal varName As String)
End Sub

Public Sub AddVariablesListToReport(ByVal scope As Byte, _ 
                                    ByVal sheetName As String, _ 
                                    ByVal varLists As BetterArray)
End Sub

'Clear the import report
Public Sub Clear()
End Sub

'Return true or False if the report sheet contains one sheet or one variable
'scope can take 2 values:
'1- sheets or variables not imported
'2- sheets or variables not used. 
Public Property Get ContainsSheet(ByVal scope As Byte, ByVal sheetName As String) As Boolean
End Property
  
Public Property Get ContainsVariable(ByVal scope As Byte, _ 
                                     ByVal sheetName As String, _ 
                                     ByVal varName As String) As Boolean
End Property