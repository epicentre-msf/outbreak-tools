VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Geo 
   Caption         =   "GEO Apps"
   ClientHeight    =   9576.001
   ClientLeft      =   48
   ClientTop       =   -348
   ClientWidth     =   10200
   OleObjectBlob   =   "F_Geo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Geo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CMD_ChoicesFac_Click()

    'If Not Sheets("geo").[T_facility].ListObject.DataBodyRange Is Nothing Then
    '    FRM_Facility.Visible = True
    '    LBL_Fac1.Visible = True
    '    FRM_Geo.Visible = False
    '    LBL_Geo1.Visible = False
    '    iGeoType = 1
    'Else
    '    [TXT_Msg].Value = "Liste des facility vide"
    'End If

End Sub

Private Sub CMD_ChoicesGeo_Click()

    'FRM_Facility.Visible = False
    'LBL_Fac1.Visible = False
    'FRM_Geo.Visible = True
    'LBL_Geo1.Visible = True
    'iGeoType = 0

End Sub

Private Sub CMD_Copier_Click()

    Dim i As Long
    Dim T_temp
    Dim sChaine As String

    Select Case iGeoType
    Case 0
        'Creation de l'histo
        'If LST_Histo.Value <> "" Then
        If Not Sheets("geo").[T_HistoGeo].ListObject.DataBodyRange Is Nothing Then
            If [T_HistoGeo].Count > 1 Then
                
                T_temp = Sheets("GEO").[T_HistoGeo]
                i = 1
                While T_temp(i, 1) <> ReverseString(TXT_Msg.value) And i < UBound(T_temp, 1)
                    i = i + 1
                Wend
                If T_temp(i, 1) <> ReverseString(TXT_Msg.value) Then
                    i = 0
                    ReDim T_histo(UBound(T_temp, 1))
                    While i < UBound(T_temp, 1)
                        T_histo(i) = T_temp(i + 1, 1)
                        i = i + 1
                    Wend
                    
                    ReDim Preserve T_histo(UBound(T_histo))
                    T_histo(UBound(T_histo)) = ReverseString(TXT_Msg.value)
                
                    'T_Histo = TriBulle(T_Histo)
                    Call QuickSort(T_histo, LBound(T_histo), UBound(T_histo))
                                        
                    Sheets("GEO").ListObjects("T_HistoGeo").Resize Range(Cells(1, Sheets("GEO").[T_HistoGeo].Column), Cells(UBound(T_histo) + 2, Sheets("GEO").[T_HistoGeo].Column))
                    'Sheets("GEO").[T_HistoGeo] = T_Histo
                    i = 0
                    While i <= UBound(T_histo)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_HistoGeo].Column).value = T_histo(i)
                        i = i + 1
                    Wend
                    
                End If
            Else
                ReDim T_temp(1)
                T_temp(0) = [T_HistoGeo]
                T_temp(1) = ReverseString(TXT_Msg.value)
                If T_temp(0) <> T_temp(1) Then
                    'T_Histo = TriBulle(T_temp)
                    T_histo = T_temp
                    Call QuickSort(T_histo, LBound(T_histo), UBound(T_histo))
                        
                    'Sheets("GEO").[T_HistoGeo].Delete
                    Sheets("GEO").ListObjects("T_HistoGeo").Resize Range(Cells(1, Sheets("GEO").[T_HistoGeo].Column), Cells(UBound(T_histo) + 2, Sheets("GEO").[T_HistoGeo].Column))
                    
                    'Mais PQ ne veut il pas ecrire ce fichu tableau ??!!
                    'Sheets("GEO").ListObjects("T_HistoGeo").DataBodyRange = T_Histo
                    i = 0
                    While i <= UBound(T_histo)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_HistoGeo].Column).value = T_histo(i)
                        i = i + 1
                    Wend
                End If
            End If
        Else
            If sPlaceSelection <> "" Then
                Sheets("GEO").[T_HistoGeo].value = ReverseString(sPlaceSelection)
            End If
        End If
        'End If
    
        'ecriture a la bonne place
        i = 0
        T_temp = Split([TXT_Msg].value, " | ")
        While i <= UBound(T_temp)
            ActiveCell.Offset(, i).value = T_temp(i)
            i = i + 1
        Wend
    
    
    Case 1
        'If LST_Histo.Value <> "" Then
        If Not Sheets("geo").[T_HistoFacil].ListObject.DataBodyRange Is Nothing Then
            If [T_HistoFacil].Count > 1 Then
                
                T_temp = Sheets("GEO").[T_HistoFacil]
                i = 1
                While T_temp(i, 1) <> TXT_Msg.value And i < UBound(T_temp, 1)
                    i = i + 1
                Wend
                If T_temp(i, 1) <> TXT_Msg.value Then
                    i = 0
                    ReDim T_histoF(UBound(T_temp, 1))
                    While i < UBound(T_temp, 1)
                        T_histoF(i) = T_temp(i + 1, 1)
                        i = i + 1
                    Wend
                    
                    ReDim Preserve T_histoF(UBound(T_histoF))
                    T_histoF(UBound(T_histoF)) = TXT_Msg.value
                
                    'T_Histo = TriBulle(T_Histo)
                    Call QuickSort(T_histoF, LBound(T_histoF), UBound(T_histoF))
                    
                    Sheets("GEO").ListObjects("T_HistoFacil").Resize Range(Cells(1, Sheets("GEO").[T_HistoFacil].Column), Cells(UBound(T_histoF) + 2, Sheets("GEO").[T_HistoFacil].Column))
                    i = 0
                    While i <= UBound(T_histoF)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_HistoFacil].Column).value = T_histoF(i)
                        i = i + 1
                    Wend
                End If
            Else
                ReDim T_temp(1)
                T_temp(0) = [T_HistoFacil]
                T_temp(1) = TXT_Msg.value
                If T_temp(0) <> T_temp(1) Then
                    'T_Histo = TriBulle(T_temp)
                    T_histoF = T_temp
                    Call QuickSort(T_histoF, LBound(T_histoF), UBound(T_histoF))
        
                    Sheets("GEO").ListObjects("T_HistoFacil").Resize Range(Cells(1, Sheets("GEO").[T_HistoFacil].Column), Cells(UBound(T_histoF) + 2, Sheets("GEO").[T_HistoFacil].Column))
                    i = 0
                    While i <= UBound(T_histoF)
                        Sheets("GEO").Cells(i + 2, Sheets("GEO").[T_HistoFacil].Column).value = T_histoF(i)
                        i = i + 1
                    Wend
                End If
            End If
        Else
            If sPlaceSelection <> "" Then
                Sheets("GEO").[T_HistoFacil].value = sPlaceSelection
            End If
        End If
        'End If
    
        'ecriture a la bonne place
        Selection.value = [TXT_Msg].value
        
    End Select

    [F_GEO].Hide

End Sub

Private Sub CMD_Retour_Click()

    'iCountLineAdm1 = UBound(T_geo0, 2)
    'iCountLineAdm2 = UBound(T_geo1, 2)
    'iCountLineAdm3 = UBound(T_geo2, 2)
    'iCountLineAdm4 = UBound(T_geo3, 2)

    Me.Hide

End Sub

Private Sub LST_Adm1_Click()

    Call ShowLst2(LST_Adm1.value)

    sPlaceSelection = TXT_Msg.value

End Sub

Private Sub LST_Adm2_Click()

    Call ShowLst3(LST_Adm2.value)

    sPlaceSelection = TXT_Msg.value

End Sub

Private Sub LST_Adm3_Click()

    Call ShowLst4(LST_Adm3.value)

    sPlaceSelection = TXT_Msg.value

End Sub

Private Sub LST_Adm4_Click()

    sPlaceSelection = [F_GEO].LST_Adm1.value & " | " & [F_GEO].LST_Adm2.value & " | " & [F_GEO].LST_Adm3.value & " | " & [F_GEO].LST_Adm4.value

    TXT_Msg.value = sPlaceSelection

End Sub

Private Sub LST_AdmF1_Click()

    Call ShowLstF2(LST_AdmF1.value)

    sPlaceSelection = TXT_Msg.value

End Sub

Private Sub LST_AdmF2_Click()

    Call ShowLstF3(LST_AdmF2.value)

    sPlaceSelection = TXT_Msg.value

End Sub

Private Sub LST_AdmF3_Click()

    Call ShowLstF4(LST_AdmF3.value)

    sPlaceSelection = TXT_Msg.value

End Sub

Private Sub LST_AdmF4_Click()

    sPlaceSelection = ReverseString([F_GEO].LST_AdmF1.value & " | " & [F_GEO].LST_AdmF2.value & " | " & [F_GEO].LST_AdmF3.value & " | " & [F_GEO].LST_AdmF4.value)

    TXT_Msg.value = sPlaceSelection

End Sub

Private Sub LST_Histo_Click()

    TXT_Msg.value = ReverseString(LST_Histo.value)
    sPlaceSelection = LST_Histo.value

End Sub

Private Sub LST_HistoF_Click()

    If LST_HistoF.value <> "" Then
        TXT_Msg.value = LST_HistoF.value
        sPlaceSelection = LST_HistoF.value
    End If

End Sub

Private Sub LST_ListeAgre_Click()

    TXT_Msg.value = LST_ListeAgre.value
    sPlaceSelection = LST_ListeAgre.value

End Sub

Private Sub LST_ListeAgreF_Click()

    TXT_Msg.value = LST_ListeAgreF.value
    sPlaceSelection = LST_ListeAgreF.value


End Sub

Private Sub TXT_Recherche_Change()

    Call SearchValue(T_concat, F_GEO.TXT_Recherche.value)

End Sub

Private Sub TXT_RechercheF_Change()

    Call SearchValueF(T_concatF, F_GEO.TXT_RechercheF.value)

End Sub

Private Sub TXT_RechercheHisto_Change()

    Call SeachHistoValue(T_histo, F_GEO.TXT_RechercheHisto.value)

End Sub

Private Sub TXT_RechercheHistoF_Change()

    Call SeachHistoValueF(T_histoF, F_GEO.TXT_RechercheHisto.value)

End Sub

