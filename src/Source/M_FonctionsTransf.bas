Attribute VB_Name = "M_FonctionsTransf"
Option Explicit
'M_FonctionsTransf

Public Function TabEstVide(T_aTest) As Boolean

Dim test As Variant

TabEstVide = False
On Error GoTo crash
test = UBound(T_aTest)
On Error GoTo 0
Exit Function

crash:
TabEstVide = True

End Function

Public Function TriBulle(T_aTrier)

Dim sInterverti As String
Dim iTest As Long
Dim bInversion As Boolean
Dim iMinTab As Long
Dim iMaxTab As Long

iMinTab = LBound(T_aTrier)
iMaxTab = UBound(T_aTrier)

bInversion = True
While bInversion = True
    bInversion = False
    For iTest = (iMinTab + 1) To iMaxTab
        If T_aTrier(iTest - 1) > T_aTrier(iTest) Then
            sInterverti = T_aTrier(iTest - 1)
            T_aTrier(iTest - 1) = T_aTrier(iTest)
            T_aTrier(iTest) = sInterverti
            bInversion = True
        End If
    Next iTest
Wend
TriBulle = T_aTrier

End Function

Public Function chargerChemin() As String

Dim fDialog As Office.FileDialog

chargerChemin = ""
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
With fDialog
    .AllowMultiSelect = False
    .Title = "Choix du fichier"
    .Filters.Clear
    .Filters.Add "Feuille de calcul Excel", "*.xlsx, *.xlsm, *.xlsb,  *.xls"

    If .Show = True Then
        chargerChemin = .SelectedItems(1)
    End If
End With
Set fDialog = Nothing

End Function

