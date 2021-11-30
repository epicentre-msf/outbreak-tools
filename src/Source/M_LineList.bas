Attribute VB_Name = "M_LineList"
Option Explicit

Public T_typeDico
Const C_ColTitre As Byte = 5

Function CreerEnteteDico() As Scripting.Dictionary

Dim i As Integer
Dim D_temp As New Scripting.Dictionary

i = 1
While Sheets("dico").Cells(1, i).Value <> ""
    D_temp.Add Sheets("dico").Cells(1, i).Value, i
    i = i + 1
Wend
Set CreerEnteteDico = D_temp
Set D_temp = Nothing

End Function

Function CreerDonneeDico(D_Entete As Scripting.Dictionary)

Dim i As Integer
Dim T_temp

ReDim T_temp(Sheets("dico").Cells(1, 1).End(xlDown).Row, D_Entete.Count)
i = 1
While i < UBound(T_temp, 1)
    T_temp(D_Entete("name"), i) = Sheets("dico").Cells(i, D_Entete("name")).Value
    T_temp(D_Entete("label_1"), i) = Sheets("dico").Cells(i, D_Entete("label_1")).Value
    T_temp(D_Entete("control"), i) = Sheets("dico").Cells(i, D_Entete("control")).Value

    i = i + 1
Wend
CreerDonneeDico = T_temp

End Function

Function retourneControl(sNom As String) As String  '

Dim i As Integer
Dim D_Entete As New Scripting.Dictionary
Dim T_data

Set D_Entete = CreerEnteteDico
T_data = CreerDonneeDico(D_Entete)

If Not TabEstVide(T_data) Then
    i = 1
    While i < UBound(T_data) And T_data(D_Entete("name"), i) <> sNom
        i = i + 1
    Wend
    If T_data(D_Entete("name"), i) = sNom Then
        retourneControl = T_data(D_Entete("control"), i)
    End If
End If
Set D_Entete = Nothing

End Function

Sub clicBtnGeoApps()

Dim iNumCol As Integer
Dim sType As String

iNumCol = ActiveCell.Column

On Error GoTo fin
If ActiveCell.Row > C_ColTitre Then
    sType = retourneControl(ActiveSheet.Cells(C_ColTitre, iNumCol).Name.Name)       'parce qu'un seul .Name ne suffit pas...
    Select Case LCase(sType)
    Case "geo"
        iTypeGeo = 0
        Call chargerGeo(iTypeGeo)
    
    Case "hf"
        iTypeGeo = 1
        Call chargerGeo(iTypeGeo)
    
    Case Else
        MsgBox "Vous n'etes pas sur la bonne cellule"
    
    End Select
End If

Exit Sub
        
fin:
MsgBox "Vous n'etes pas sur la bonne cellule"

End Sub
