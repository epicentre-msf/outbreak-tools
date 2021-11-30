VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Geo 
   Caption         =   "GEO Apps"
   ClientHeight    =   8415.001
   ClientLeft      =   45
   ClientTop       =   -165
   ClientWidth     =   14415
   OleObjectBlob   =   "F_Geo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_GEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CMD_ChoixFac_Click()

If Not Sheets("geo").[T_facility].ListObject.DataBodyRange Is Nothing Then
    FRM_Facility.Visible = True
    LBL_Fac1.Visible = True
    FRM_Geo.Visible = False
    LBL_Geo1.Visible = False
    iTypeGeo = 1
Else
    [TXT_Msg].Value = "Liste des facility vide"
End If

End Sub

Private Sub CMD_ChoixGeo_Click()

FRM_Facility.Visible = False
LBL_Fac1.Visible = False
FRM_Geo.Visible = True
LBL_Geo1.Visible = True
iTypeGeo = 0

End Sub

Private Sub CMD_Copier_Click()

Dim i As Long
Dim T_temp
Dim sChaine As String

Select Case iTypeGeo
Case 0
    'Creation de l'histo
    'If LST_Histo.Value <> "" Then
        If Not Sheets("geo").[T_HistoGeo].ListObject.DataBodyRange Is Nothing Then
            If [T_HistoGeo].Count > 1 Then
                
                T_temp = Sheets("GEO").[T_HistoGeo]
                i = 1
                While T_temp(i, 1) <> inverserChaine(TXT_Msg.Value) And i < UBound(T_temp, 1)
                    i = i + 1
                Wend
                If T_temp(i, 1) <> inverserChaine(TXT_Msg.Value) Then
                    i = 0
                    ReDim T_Histo(UBound(T_temp, 1))
                    While i < UBound(T_temp, 1)
                        T_Histo(i) = T_temp(i + 1, 1)
                        i = i + 1
                    Wend
                    
                    ReDim Preserve T_Histo(UBound(T_Histo))
                    T_Histo(UBound(T_Histo)) = inverserChaine(TXT_Msg.Value)
                
                    T_Histo = TriBulle(T_Histo)
                    
                    Sheets("GEO").ListObjects("T_HistoGeo").Resize Range(Cells(1, Sheets("GEO").[T_HistoGeo].Column), Cells(UBound(T_Histo) + 2, Sheets("GEO").[T_HistoGeo].Column))
                    'Sheets("GEO").[T_HistoGeo] = T_Histo
                    i = 0
                    While i <= UBound(T_Histo)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_HistoGeo].Column).Value = T_Histo(i)
                        i = i + 1
                    Wend
                    
                End If
            Else
                ReDim T_temp(1)
                T_temp(0) = [T_HistoGeo]
                T_temp(1) = inverserChaine(TXT_Msg.Value)
                If T_temp(0) <> T_temp(1) Then
                    T_Histo = TriBulle(T_temp)
                        
                    'Sheets("GEO").[T_HistoGeo].Delete
                    Sheets("GEO").ListObjects("T_HistoGeo").Resize Range(Cells(1, Sheets("GEO").[T_HistoGeo].Column), Cells(UBound(T_Histo) + 2, Sheets("GEO").[T_HistoGeo].Column))
                    
                    'Mais PQ ne veut il pas ecrire ce fichu tableau ??!!
                    'Sheets("GEO").ListObjects("T_HistoGeo").DataBodyRange = T_Histo
                    i = 0
                    While i <= UBound(T_Histo)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_HistoGeo].Column).Value = T_Histo(i)
                        i = i + 1
                    Wend
                End If
             End If
        Else
            If sLieuSelection <> "" Then
                Sheets("GEO").[T_HistoGeo].Value = inverserChaine(sLieuSelection)
            End If
        End If
    'End If
    
    'ecriture a la bonne place
    i = 0
    T_temp = Split([TXT_Msg].Value, " | ")
    While i <= UBound(T_temp)
        ActiveCell.Offset(, i).Value = T_temp(i)
        i = i + 1
    Wend
    
    
Case 1
    'If LST_Histo.Value <> "" Then
        If Not Sheets("geo").[T_Histofacil].ListObject.DataBodyRange Is Nothing Then
            If [T_Histofacil].Count > 1 Then
                
                T_temp = Sheets("GEO").[T_Histofacil]
                i = 1
                While T_temp(i, 1) <> inverserChaine(TXT_Msg.Value) And i < UBound(T_temp, 1)
                    i = i + 1
                Wend
                If T_temp(i, 1) <> inverserChaine(TXT_Msg.Value) Then
                    i = 0
                    ReDim T_Histo(UBound(T_temp, 1))
                    While i < UBound(T_temp, 1)
                        T_Histo(i) = T_temp(i + 1, 1)
                        i = i + 1
                    Wend
                    
                    ReDim Preserve T_Histo(UBound(T_Histo))
                    T_Histo(UBound(T_Histo)) = inverserChaine(TXT_Msg.Value)
                
                    T_Histo = TriBulle(T_Histo)
                    
                    Sheets("GEO").ListObjects("T_HistoFacil").Resize Range(Cells(1, Sheets("GEO").[T_Histofacil].Column), Cells(UBound(T_Histo) + 2, Sheets("GEO").[T_Histofacil].Column))
                    i = 0
                    While i <= UBound(T_Histo)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_Histofacil].Column).Value = T_Histo(i)
                        i = i + 1
                    Wend
                End If
            Else
                ReDim T_temp(1)
                T_temp(0) = [T_Histofacil]
                T_temp(1) = inverserChaine(TXT_Msg.Value)
                If T_temp(0) <> T_temp(1) Then
                    T_Histo = TriBulle(T_temp)
        
                    Sheets("GEO").ListObjects("T_HistoFacil").Resize Range(Cells(1, Sheets("GEO").[T_Histofacil].Column), Cells(UBound(T_Histo) + 2, Sheets("GEO").[T_Histofacil].Column))
                    i = 0
                    While i <= UBound(T_Histo)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_Histofacil].Column).Value = T_Histo(i)
                        i = i + 1
                    Wend
                End If
            End If
        Else
            If sLieuSelection <> "" Then
                Sheets("GEO").[T_Histofacil].Value = inverserChaine(sLieuSelection)
            End If
        End If
    'End If
    
    'ecriture a la bonne place
    Selection.Value = [TXT_Msg].Value
        
End Select

[F_GEO].Hide

End Sub

Private Sub CMD_Retour_Click()

Me.Hide

End Sub

Private Sub LST_Adm1_Click()

Call afficherLst2(LST_Adm1.Value)

sLieuSelection = TXT_Msg.Value

End Sub

Private Sub LST_Adm2_Click()

Call afficherLst3(LST_Adm2.Value)

sLieuSelection = TXT_Msg.Value

End Sub

Private Sub LST_Adm3_Click()

Call afficherLst4(LST_Adm3.Value)

sLieuSelection = TXT_Msg.Value

End Sub

Private Sub LST_Adm4_Click()

sLieuSelection = [F_GEO].LST_Adm1.Value & " | " & [F_GEO].LST_Adm2.Value & " | " & [F_GEO].LST_Adm3.Value & " | " & [F_GEO].LST_Adm4.Value

TXT_Msg.Value = sLieuSelection

End Sub

Private Sub LST_AdmF1_Click()

Call afficherLstF2(LST_AdmF1.Value)

sLieuSelection = TXT_Msg.Value

End Sub

Private Sub LST_AdmF2_Click()

Call afficherLstF3(LST_AdmF2.Value)

sLieuSelection = TXT_Msg.Value

End Sub

Private Sub LST_AdmF3_Click()

Call afficherLstF4(LST_AdmF3.Value)

sLieuSelection = TXT_Msg.Value

End Sub

Private Sub LST_AdmF4_Click()

sLieuSelection = [F_GEO].LST_AdmF1.Value & " | " & [F_GEO].LST_AdmF2.Value & " | " & [F_GEO].LST_AdmF3.Value & " | " & [F_GEO].LST_AdmF4.Value

TXT_Msg.Value = sLieuSelection

End Sub

Private Sub LST_Histo_Click()

TXT_Msg.Value = inverserChaine(LST_Histo.Value)
sLieuSelection = LST_Histo.Value

End Sub

Private Sub LST_HistoF_Click()

If LST_HistoF.Value <> "" Then
    TXT_Msg.Value = LST_HistoF.Value
    sLieuSelection = LST_HistoF.Value
End If

End Sub

Private Sub LST_ListeAgre_Click()

TXT_Msg.Value = LST_ListeAgre.Value
sLieuSelection = LST_ListeAgre.Value

End Sub

Private Sub LST_ListeAgreF_Click()

TXT_Msg.Value = LST_ListeAgreF.Value
sLieuSelection = LST_ListeAgreF.Value


End Sub

Private Sub TXT_Recherche_Change()

Call ChercherValeurTab(T_concat, F_GEO.TXT_Recherche.Value)

End Sub

Private Sub TXT_RechercheF_Change()

Call ChercherValeurTabF(T_concatF, F_GEO.TXT_RechercheF.Value)

End Sub

Private Sub TXT_RechercheHisto_Change()

Call ChercherValeurHisto(T_Histo, F_GEO.TXT_RechercheHisto.Value)

End Sub

Private Sub TXT_RechercheHistoF_Change()

Call ChercherValeurHistoF(T_HistoF, F_GEO.TXT_RechercheHisto.Value)

End Sub

