Attribute VB_Name = "M_Geo"
Option Explicit
'M_Geo
Public T_geo0
Public T_aff0
Public T_geo1
Public T_aff1
Public T_geo2
Public T_aff2
Public T_geo3
Public T_aff3
Public T_concat
Public T_Histo

Public T_fac
Public T_concatF
Public T_HistoF

Public sLieuSelection As String

Public iTypeGeo As Byte 'geo 0 ou facility 1 ?

Sub chargerGeo(iTypeGeo As Byte)

Dim i As Long
Dim iNbMax As Long
Dim iNbMaxF As Long
Dim iDerNiveau As Long
Dim j As Integer
Dim k As Integer
Dim l As Integer

Application.ScreenUpdating = False
Call vidangeGeo

[F_GEO].Height = 360
[F_GEO].Width = 606

If TabEstVide(T_geo0) Then
    If Not Sheets("geo").[T_adm0].ListObject.DataBodyRange Is Nothing Then
        T_geo0 = Sheets("geo").[T_adm0]
        iNbMax = UBound(T_geo0, 1)
        iDerNiveau = 0
        
        'remplit LST_Adm
        ReDim T_aff0(iNbMax - 1)
        i = 1
        j = 0
        While i <= iNbMax
            T_aff0(j) = T_geo0(i, 2)
            i = i + 1
            j = j + 1
        Wend
        [F_GEO].[LST_Adm1].List = T_aff0
        [F_GEO].[LST_AdmF1].List = T_aff0
    End If
Else
    [F_GEO].[LST_Adm1].List = T_aff0
    [F_GEO].[LST_AdmF1].List = T_aff0
    iNbMax = UBound(T_geo0, 1)
    iDerNiveau = 0
End If

If TabEstVide(T_geo1) Then
    If Not Sheets("geo").[T_adm1].ListObject.DataBodyRange Is Nothing Then
        T_geo1 = Sheets("geo").[T_adm1]
        iNbMax = UBound(T_geo1, 1)
        iDerNiveau = 1
        
        F_GEO.LBL_Adm1.Caption = Sheets("geo").[T_adm0].ListObject.HeaderRowRange.Item(2).Value
        F_GEO.LBL_Adm1F.Caption = Sheets("geo").[T_adm0].ListObject.HeaderRowRange.Item(2).Value
    End If
Else
    iNbMax = UBound(T_geo1, 1)
    iDerNiveau = 1
    
    F_GEO.LBL_Adm1.Caption = Sheets("geo").[T_adm0].ListObject.HeaderRowRange.Item(2).Value
    F_GEO.LBL_Adm1F.Caption = Sheets("geo").[T_adm0].ListObject.HeaderRowRange.Item(2).Value
End If

If TabEstVide(T_geo2) Then
    If Not Sheets("geo").[T_adm2].ListObject.DataBodyRange Is Nothing Then
        T_geo2 = Sheets("geo").[T_adm2]
        iNbMax = UBound(T_geo2, 1)
        iDerNiveau = 2
        
        F_GEO.LBL_Adm2.Caption = Sheets("geo").[T_adm1].ListObject.HeaderRowRange.Item(2).Value
        F_GEO.LBL_Adm2F.Caption = Sheets("geo").[T_adm1].ListObject.HeaderRowRange.Item(2).Value
    End If
Else
    iNbMax = UBound(T_geo2, 1)
    iDerNiveau = 2
    
    F_GEO.LBL_Adm2.Caption = Sheets("geo").[T_adm1].ListObject.HeaderRowRange.Item(2).Value
    F_GEO.LBL_Adm2F.Caption = Sheets("geo").[T_adm1].ListObject.HeaderRowRange.Item(2).Value
End If

If TabEstVide(T_geo3) Then
    If Not Sheets("geo").[T_adm3].ListObject.DataBodyRange Is Nothing Then
        T_geo3 = Sheets("geo").[T_adm3].ListObject.DataBodyRange
        iNbMax = UBound(T_geo3, 1)
        iDerNiveau = 3
        
        F_GEO.LBL_Adm3.Caption = Sheets("geo").[T_adm2].ListObject.HeaderRowRange.Item(2).Value
        F_GEO.LBL_Adm3F.Caption = Sheets("geo").[T_adm2].ListObject.HeaderRowRange.Item(2).Value
    End If
Else
    iNbMax = UBound(T_geo3, 1)
    iDerNiveau = 3
    
    F_GEO.LBL_Adm3.Caption = Sheets("geo").[T_adm2].ListObject.HeaderRowRange.Item(2).Value
    F_GEO.LBL_Adm3F.Caption = Sheets("geo").[T_adm2].ListObject.HeaderRowRange.Item(2).Value
End If

If TabEstVide(T_fac) Then
    If Not Sheets("geo").[T_facility].ListObject.DataBodyRange Is Nothing Then
        T_fac = Sheets("geo").[T_facility]
        iNbMaxF = UBound(T_fac, 1)
        
        F_GEO.LBL_Adm4.Caption = Sheets("geo").[T_adm3].ListObject.HeaderRowRange.Item(2).Value
        F_GEO.LBL_Adm4F.Caption = Sheets("geo").[T_adm3].ListObject.HeaderRowRange.Item(2).Value
    End If
Else
    iNbMaxF = UBound(T_fac, 1)
    
    F_GEO.LBL_Adm4.Caption = Sheets("geo").[T_adm3].ListObject.HeaderRowRange.Item(2).Value
    F_GEO.LBL_Adm4F.Caption = Sheets("geo").[T_adm3].ListObject.HeaderRowRange.Item(2).Value
End If

'creation du tableau concat
If TabEstVide(T_concat) Then
    ReDim T_concat(iNbMax)  'attention a la desynchro de -1
    i = 1
    While i <= iNbMax
        Select Case iDerNiveau
        Case 0
            T_concat(i - 1) = T_geo0(i, 2)
    
        Case 1
            T_concat(i - 1) = T_geo1(i, 1) & " | " & T_geo1(i, 2)
    
        Case 2
            T_concat(i - 1) = T_geo2(i, 1) & " | " & T_geo2(i, 2)
            
            j = 1
            While j <= UBound(T_geo1, 1)
                If T_geo1(j, 2) = T_geo2(i, 1) Then
                    T_concat(i - 1) = T_geo1(j, 1) & " | " & T_concat(i - 1)
                End If
                j = j + 1
            Wend
            
        Case 3
            T_concat(i - 1) = T_geo3(i, 1) & " | " & T_geo3(i, 2)
            
            j = 1
            While j <= UBound(T_geo2, 1)
                If T_geo2(j, 2) = T_geo3(i, 1) Then
                    T_concat(i - 1) = T_geo2(j, 1) & " | " & T_concat(i - 1)
                    
                    k = 1
                    While k <= UBound(T_geo1, 1)
                        If T_geo1(k, 2) = T_geo2(j, 1) Then
                            T_concat(i - 1) = T_geo1(k, 1) & " | " & T_concat(i - 1)
                        End If
                        k = k + 1
                    Wend
                End If
                j = j + 1
            Wend
        End Select
        i = i + 1
    Wend
    If Not TabEstVide(T_concat) Then
        [F_GEO].LST_ListeAgre.List = TriBulle(T_concat)
    End If
Else
    [F_GEO].LST_ListeAgre.List = T_concat
End If

If TabEstVide(T_concatF) Then
    ReDim T_concatF(iNbMaxF)     'meme manip pour facility
    i = 1
    While i <= iNbMaxF
        Select Case iDerNiveau
        Case 0
            T_concatF(i - 1) = T_geo0(i, 1)
        
        Case 1
            T_concatF(i - 1) = T_geo1(i, 1) & " | " & T_geo1(i, 2)
    
        Case 2
            T_concatF(i - 1) = T_fac(i, 1) & " | " & T_fac(i, 2)
            
            j = 1
            While j <= UBound(T_geo2, 1)
                If T_fac(i, 1) = T_geo2(j, 2) Then
                    T_concatF(i - 1) = T_geo2(j, 1) & " | " & T_concatF(i - 1)
                End If
                j = j + 1
            Wend
        
        Case 3
            T_concatF(i - 1) = T_fac(i, 1) & " | " & T_fac(i, 2)
            
            j = 1
            While j <= UBound(T_geo2, 1)
                If T_geo2(j, 2) = T_fac(i, 1) Then
                    T_concatF(i - 1) = T_geo2(j, 1) & " | " & T_concatF(i - 1)
                    
                    k = 1
                    While k <= UBound(T_geo1, 1)
                        If T_geo1(k, 2) = T_geo2(j, 1) Then
                            T_concatF(i - 1) = T_geo1(k, 1) & " | " & T_concatF(i - 1)
                        End If
                        k = k + 1
                    Wend
                End If
                j = j + 1
            Wend
        End Select
        i = i + 1
    Wend
    If Not TabEstVide(T_concatF) Then
        [F_GEO].LST_ListeAgreF.List = TriBulle(T_concatF)
    End If
Else
    [F_GEO].LST_ListeAgreF.List = T_concatF
End If

'Creation de l'histo
If Not Sheets("geo").[T_HistoGeo].ListObject.DataBodyRange Is Nothing Then
    i = 1
    ReDim T_Histo(i)
    If [T_HistoGeo].Count > 1 Then
        T_Histo = [T_HistoGeo]
    Else
        T_Histo(0) = [T_HistoGeo]
    End If
    [F_GEO].LST_Histo.List = T_Histo
End If

If Not Sheets("geo").[T_Histofacil].ListObject.DataBodyRange Is Nothing Then
    i = 1
    ReDim T_HistoF(i)
    If [T_Histofacil].Count > 1 Then
        T_HistoF = [T_Histofacil]
    Else
        T_HistoF(0) = [T_Histofacil]
    End If
    [F_GEO].LST_HistoF.List = T_HistoF
End If

'
Select Case iTypeGeo
Case 0
    [F_GEO].FRM_Facility.Visible = False
    [F_GEO].FRM_Geo.Visible = True
    [F_GEO].LBL_Fac1.Visible = False
    [F_GEO].LBL_Geo1.Visible = True
Case 1
    [F_GEO].FRM_Facility.Visible = True
    [F_GEO].FRM_Geo.Visible = False
    [F_GEO].LBL_Fac1.Visible = True
    [F_GEO].LBL_Geo1.Visible = False
End Select
Application.ScreenUpdating = True
'the show must go on
[F_GEO].Show

End Sub

Sub afficherLst2(sLieu As String)

Dim i As Integer
Dim j As Integer

[F_GEO].LST_Adm2.Clear
[F_GEO].LST_Adm3.Clear
[F_GEO].LST_Adm4.Clear

If Not TabEstVide(T_geo1) Then
    i = 1
    j = 0
    ReDim T_aff2(0)
    While i <= UBound(T_geo1)
        If T_geo1(i, 1) = sLieu Then
            ReDim Preserve T_aff2(j)
            T_aff2(j) = T_geo1(i, 2)
            j = j + 1
        End If
        i = i + 1
    Wend
End If
    
If Not TabEstVide(T_aff2) Then
    [F_GEO].LST_Adm2.List = T_aff2
    [F_GEO].TXT_Msg.Value = sLieu
Else
    [F_GEO].TXT_Msg.Value = sLieu '& " : Pas de niveau2"
End If

End Sub

Sub afficherLstF2(sLieu As String)

Dim i As Integer
Dim j As Integer
Dim bTrouve As Boolean

[F_GEO].LST_AdmF2.Clear
[F_GEO].LST_AdmF3.Clear
[F_GEO].LST_AdmF4.Clear

bTrouve = False
If Not TabEstVide(T_fac) Then
    i = 1
    j = 0
    ReDim T_aff2(0)
    While i <= UBound(T_fac)
        If T_fac(i, 1) = sLieu Then
            ReDim Preserve T_aff2(j)
            T_aff2(j) = T_fac(i, 2)
            bTrouve = True
            j = j + 1
        End If
        i = i + 1
    Wend
    If Not bTrouve Then
        i = 1
        j = 0
        ReDim T_aff2(0)
        While i <= UBound(T_geo1)
            If T_geo1(i, 1) = sLieu Then
                ReDim Preserve T_aff2(j)
                T_aff2(j) = T_geo1(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
End If
    
If Not TabEstVide(T_aff2) Then
    [F_GEO].LST_AdmF2.List = T_aff2
    [F_GEO].TXT_Msg.Value = sLieu
Else
    [F_GEO].TXT_Msg.Value = sLieu '& " : Pas de niveau2"
End If

End Sub

Sub afficherLst3(sLieu As String)

Dim i As Integer
Dim j As Integer

[F_GEO].LST_Adm3.Clear
[F_GEO].LST_Adm4.Clear

If Not TabEstVide(T_geo2) Then
    i = 1
    j = 0
    ReDim T_aff3(0)
    While i <= UBound(T_geo2)
        If T_geo2(i, 1) = sLieu Then
            ReDim Preserve T_aff3(j)
            T_aff3(j) = T_geo2(i, 2)
            j = j + 1
        End If
        i = i + 1
    Wend
End If
    
If Not TabEstVide(T_aff3) Then
    [F_GEO].LST_Adm3.List = T_aff3
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_Adm1.Value & " | " & [F_GEO].LST_Adm2.Value
Else
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_Adm1.Value & " | " & [F_GEO].LST_Adm2.Value '& " : Pas de niveau 3"
End If

End Sub

Sub afficherLstF3(sLieu As String)

Dim i As Integer
Dim j As Integer
Dim bTrouve As Boolean

[F_GEO].LST_AdmF3.Clear
[F_GEO].LST_AdmF4.Clear

bTrouve = False
If Not TabEstVide(T_fac) Then
    i = 1
    j = 0
    ReDim T_aff3(0)
    While i <= UBound(T_fac)
        If T_fac(i, 1) = sLieu Then
            ReDim Preserve T_aff3(j)
            T_aff3(j) = T_fac(i, 2)
            bTrouve = True
            j = j + 1
        End If
        i = i + 1
    Wend
    If Not bTrouve Then
        i = 1
        j = 0
        ReDim T_aff2(0)
        While i <= UBound(T_geo2)
            If T_geo2(i, 1) = sLieu Then
                ReDim Preserve T_aff3(j)
                T_aff3(j) = T_geo2(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
End If
    
If Not TabEstVide(T_aff3) Then
    [F_GEO].LST_AdmF3.List = T_aff3
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_AdmF2.Value & " | " & [F_GEO].LST_AdmF1.Value
Else
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_AdmF2.Value & " | " & [F_GEO].LST_AdmF1.Value '& " : Pas de niveau 3"
End If

End Sub

Sub afficherLst4(sLieu As String)

Dim i As Integer
Dim j As Integer

[F_GEO].LST_Adm4.Clear

If Not TabEstVide(T_geo3) Then
    i = 1
    j = 0
    ReDim T_aff4(0)
    While i <= UBound(T_geo3)
        If T_geo3(i, 1) = sLieu Then
            ReDim Preserve T_aff4(j)
            T_aff4(j) = T_geo3(i, 2)
            j = j + 1
        End If
        i = i + 1
    Wend
End If
    
If Not TabEstVide(T_aff4) Then
    [F_GEO].LST_Adm4.List = T_aff4
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_Adm1.Value & " | " & [F_GEO].LST_Adm2.Value & " | " & [F_GEO].LST_Adm3.Value
    
Else
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_Adm1.Value & " | " & [F_GEO].LST_Adm2.Value & " | " & [F_GEO].LST_Adm3.Value '& " : Pas de niveau 4"
End If

End Sub

Sub afficherLstF4(sLieu As String)

Dim i As Integer
Dim j As Integer
Dim bTrouve As Boolean

[F_GEO].LST_AdmF4.Clear

bTrouve = False
If Not TabEstVide(T_fac) Then
    i = 1
    j = 0
    ReDim T_aff4(0)
    While i <= UBound(T_fac)
        If T_fac(i, 1) = sLieu Then
            ReDim Preserve T_aff4(j)
            T_aff4(j) = T_fac(i, 2)
            bTrouve = True
            j = j + 1
        End If
        i = i + 1
    Wend
    If Not bTrouve Then
        i = 1
        j = 0
        ReDim T_aff4(0)
        While i <= UBound(T_geo3)
            If T_geo3(i, 1) = sLieu Then
                ReDim Preserve T_aff4(j)
                T_aff4(j) = T_geo3(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
End If

If Not TabEstVide(T_aff4) Then
    [F_GEO].LST_AdmF4.List = T_aff4
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_AdmF3.Value & " | " & [F_GEO].LST_AdmF2.Value & " | " & [F_GEO].LST_AdmF1.Value
Else
    [F_GEO].TXT_Msg.Value = [F_GEO].LST_AdmF3.Value & " | " & [F_GEO].LST_AdmF2.Value & " | " & [F_GEO].LST_AdmF1.Value '& " : Pas de niveau 4"
End If

End Sub

Sub vidangeGeo()

[F_GEO].LST_Adm1.Clear
[F_GEO].LST_Adm2.Clear
[F_GEO].LST_Adm3.Clear
[F_GEO].LST_Adm4.Clear
[F_GEO].LST_ListeAgre.Clear
[F_GEO].LST_AdmF1.Clear
[F_GEO].LST_AdmF2.Clear
[F_GEO].LST_AdmF3.Clear
[F_GEO].LST_AdmF4.Clear
[F_GEO].LST_ListeAgreF.Clear

[F_GEO].TXT_Msg.Value = ""

End Sub

Sub ChercherValeurTab(T_concat, sValeurCherche As String)

Dim T_result
Dim i As Long
Dim j As Long

'[F_Geo].LST_ListeAgre.Clear

If Len(sValeurCherche) >= 3 Then
    i = 0
    j = 0
    ReDim T_result(j)
    While i <= UBound(T_concat)
        If InStr(1, LCase(T_concat(i)), LCase(sValeurCherche)) > 0 Then
            ReDim Preserve T_result(j)
            T_result(j) = T_concat(i)
            j = j + 1
        End If
        i = i + 1
    Wend
    
    If Not TabEstVide(T_result) Then
        [F_GEO].LST_ListeAgre.List = TriBulle(T_result)
    Else
        If [F_GEO].LST_ListeAgre.ListCount - 1 <> UBound(T_concat) Then
            [F_GEO].LST_ListeAgre.List = T_concat
        End If
    End If
Else
    If [F_GEO].LST_ListeAgre.ListCount - 1 <> UBound(T_concat) Then
        [F_GEO].LST_ListeAgre.List = T_concat
    End If
End If

End Sub

Sub ChercherValeurHisto(T_Histo, sValeurCherche As String)

Dim T_result
Dim i As Long
Dim j As Long

'[F_Geo].LST_ListeAgre.Clear

If Len(sValeurCherche) >= 3 Then
    i = 0
    j = 0
    ReDim T_result(j)
    While i <= UBound(T_Histo)
        If InStr(1, LCase(T_Histo(i)), LCase(sValeurCherche)) > 0 Then
            ReDim Preserve T_result(j)
            T_result(j) = T_Histo(i)
            j = j + 1
        End If
        i = i + 1
    Wend
    
    If Not TabEstVide(T_result) Then
        [F_GEO].LST_Histo.List = TriBulle(T_result)
    Else
        If [F_GEO].LST_Histo.ListCount - 1 <> UBound(T_Histo) Then
            [F_GEO].LST_Histo.List = T_Histo
        End If
    End If
Else
    If [F_GEO].LST_Histo.ListCount - 1 <> UBound(T_Histo) Then
        [F_GEO].LST_Histo.List = T_Histo
    End If
End If

End Sub

Sub ChercherValeurTabF(T_concatF, sValeurCherche As String)

Dim T_result
Dim i As Long
Dim j As Long

If Len(sValeurCherche) >= 3 Then
    i = 0
    j = 0
    ReDim T_result(j)
    While i <= UBound(T_concatF)
        If InStr(1, LCase(T_concatF(i)), LCase(sValeurCherche)) > 0 Then
            ReDim Preserve T_result(j)
            T_result(j) = T_concatF(i)
            j = j + 1
        End If
        i = i + 1
    Wend
    
    If Not TabEstVide(T_result) Then
        [F_GEO].LST_ListeAgreF.List = TriBulle(T_result)
    Else
        If [F_GEO].LST_ListeAgreF.ListCount - 1 <> UBound(T_concatF) Then
            [F_GEO].LST_ListeAgreF.List = T_concatF
        End If
    End If
Else
    If [F_GEO].LST_ListeAgreF.ListCount - 1 <> UBound(T_concatF) Then
        [F_GEO].LST_ListeAgreF.List = T_concatF
    End If
End If

End Sub

Sub ChercherValeurHistoF(T_HistoF, sValeurCherche As String)

Dim T_result
Dim i As Long
Dim j As Long

If Len(sValeurCherche) >= 3 Then
    i = 0
    j = 0
    ReDim T_result(j)
    While i <= UBound(T_HistoF)
        If InStr(1, LCase(T_HistoF(i)), LCase(sValeurCherche)) > 0 Then
            ReDim Preserve T_result(j)
            T_result(j) = T_HistoF(i)
            j = j + 1
        End If
        i = i + 1
    Wend
    
    If Not TabEstVide(T_result) Then
        [F_GEO].LST_HistoF.List = TriBulle(T_result)
    Else
        If [F_GEO].LST_HistoF.ListCount - 1 <> UBound(T_HistoF) Then
            [F_GEO].LST_HistoF.List = T_Histo
        End If
    End If
Else
    If [F_GEO].LST_HistoF.ListCount - 1 <> UBound(T_HistoF) Then
        [F_GEO].LST_HistoF.List = T_Histo
    End If
End If

End Sub

Function inverserChaine(sChaine As String)

Dim i As Integer
Dim T_temp
Dim sRes As String

inverserChaine = ""
T_temp = Split(sChaine, " | ")
i = UBound(T_temp)
While i >= 0
    If i = UBound(T_temp) Then
        sRes = T_temp(i)
    Else
        sRes = sRes & " | " & T_temp(i)
    End If
    i = i - 1
Wend
inverserChaine = sRes

End Function

