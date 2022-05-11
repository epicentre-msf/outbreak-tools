VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

BeginWork xlsapp:=Application
    'If does not have a name property do nothing
'    Dim sIsoCountry As String, sCountry As String, sSubCounty As String, sWard As String, sPlace As String, sFacility As String
    
    On Error GoTo no
        If Target.Name.Name = "RNG_LangDesigner" Then
          Call TranslateDesigner
        ElseIf Target.Name.Name = "RNG_LLForm" Then
        
            Call translateHeadGeo

        End If
        Exit Sub
no:

EndWork xlsapp:=Application
End Sub
