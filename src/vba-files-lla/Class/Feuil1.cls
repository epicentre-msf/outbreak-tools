VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    'If does not have a name property do nothing
    
    Dim sIsoCountry As String, sCountry As String, sSubCounty As String, sWard As String, sPlace As String, sFacility As String
    
    On Error GoTo no
        If Target.Name.Name = "RNG_LangDesigner" Then
            StartTranslate
        ElseIf Target.Name.Name = "RNG_LLForm" Then 'lla
            Select Case [RNG_LLForm].value
                Case "English"
                    sIsoCountry = "ENG"
                Case "Français"
                    sIsoCountry = "FRA"
                Case "Español"
                    sIsoCountry = "SPA"
                Case "Portugués"
                    sIsoCountry = "POR"
                Case Else
                    sIsoCountry = "ARA"
            End Select

            sCountry = Application.WorksheetFunction.HLookup(sIsoCountry, Sheets("GEO").[T_NAMES_GEO], 2, False)
            sSubCounty = Application.WorksheetFunction.HLookup(sIsoCountry, Sheets("GEO").[T_NAMES_GEO], 3, False)
            sWard = Application.WorksheetFunction.HLookup(sIsoCountry, Sheets("GEO").[T_NAMES_GEO], 4, False)
            sPlace = Application.WorksheetFunction.HLookup(sIsoCountry, Sheets("GEO").[T_NAMES_GEO], 5, False)
            sFacility = Application.WorksheetFunction.HLookup(sIsoCountry, Sheets("GEO").[T_NAMES_GEO], 6, False)
                
            Sheets("GEO").Range("A1,E1,J1,P1,Z1").value = sCountry
            Sheets("GEO").Range("F1,K1,Q1,Y1").value = sSubCounty
            Sheets("GEO").Range("L1,R1,X1").value = sWard
            Sheets("GEO").Range("S1").value = sPlace
            Sheets("GEO").Range("W1").value = sFacility
            
            
            Sheets("linelist-translation").[RNG_Language].value = [RNG_LLForm].value 'check Language of linelist's forms
            
        End If
        Exit Sub
no:
End Sub


