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
Public T_histo

Public T_fac
Public T_concatF
Public T_histoF

Public sPlaceSelection As String

Public iCountLineAdm1 As Long
Public iCountLineAdm2 As Long
Public iCountLineAdm3 As Long
Public iCountLineAdm4 As Long
Public iCountLineFac As Long
Public bHaveToDo As Boolean

Public iGeoType As Byte                          'geo 0 ou facility 1 ?

Sub chargerGeo(iGeoType As Byte)

    Dim i As Long
    Dim iNbMax As Long
    Dim iNbMaxF As Long
    Dim iLastLevel As Long
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer

    Application.ScreenUpdating = False
    Call ClearGeo
    bHaveToDo = False
    [F_GEO].Height = 360
    [F_GEO].Width = 606

    If Not Sheets("geo").[T_adm0].ListObject.DataBodyRange Is Nothing Then
        If IsEmptyTable(T_geo0) Or iCountLineAdm1 <> Sheets("geo").[T_adm0].Rows.Count Then
            T_geo0 = Sheets("geo").[T_adm0]
            iNbMax = UBound(T_geo0, 1)
            'remplit LST_Adm
            ReDim T_aff0(iNbMax - 1)
            i = 1
            j = 0
            While i <= iNbMax
                T_aff0(j) = T_geo0(i, 2)
                i = i + 1
                j = j + 1
            Wend
            iCountLineAdm1 = UBound(T_aff0) + 1
            bHaveToDo = True
        End If
    
        [F_GEO].[LST_Adm1].List = T_aff0
        [F_GEO].[LST_AdmF1].List = T_aff0
        iLastLevel = 0
    End If

    If Not Sheets("geo").[T_adm1].ListObject.DataBodyRange Is Nothing Then
        If IsEmptyTable(T_geo1) Or iCountLineAdm2 <> Sheets("geo").[T_adm1].Rows.Count Then
            T_geo1 = Sheets("geo").[T_adm1]
            iCountLineAdm2 = UBound(T_geo1, 1)
            bHaveToDo = True
        End If
        iNbMax = UBound(T_geo1, 1)
        iLastLevel = 1
        F_GEO.LBL_Adm1.Caption = Sheets("geo").[T_adm0].ListObject.HeaderRowRange.Item(2).value
        F_GEO.LBL_Adm1F.Caption = Sheets("geo").[T_adm0].ListObject.HeaderRowRange.Item(2).value
    End If

    If Not Sheets("geo").[T_adm2].ListObject.DataBodyRange Is Nothing Then
        If IsEmptyTable(T_geo2) Or Sheets("geo").[T_adm2].Rows.Count <> iCountLineAdm3 Then
            T_geo2 = Sheets("geo").[T_adm2]
            iCountLineAdm3 = UBound(T_geo2, 1)
            bHaveToDo = True
        End If
        iNbMax = UBound(T_geo2, 1)
        iLastLevel = 2
        F_GEO.LBL_Adm2.Caption = Sheets("geo").[T_adm1].ListObject.HeaderRowRange.Item(2).value
        F_GEO.LBL_Adm2F.Caption = Sheets("geo").[T_adm1].ListObject.HeaderRowRange.Item(2).value
    End If

    If Not Sheets("geo").[T_adm3].ListObject.DataBodyRange Is Nothing Then
        If IsEmptyTable(T_geo3) Or iCountLineAdm4 <> Sheets("geo").[T_adm3].Rows.Count Then
            T_geo3 = Sheets("geo").[T_adm3].ListObject.DataBodyRange
            iCountLineAdm4 = UBound(T_geo3, 1)
            bHaveToDo = True
        End If
        iNbMax = UBound(T_geo3, 1)
        iLastLevel = 3
        F_GEO.LBL_Adm3.Caption = Sheets("geo").[T_adm2].ListObject.HeaderRowRange.Item(2).value
        F_GEO.LBL_Adm3F.Caption = Sheets("geo").[T_adm2].ListObject.HeaderRowRange.Item(2).value
    End If

    If Not Sheets("geo").[T_facility].ListObject.DataBodyRange Is Nothing Then
        If IsEmptyTable(T_fac) Or iCountLineFac <> Sheets("geo").[T_facility].Rows.Count Then
            T_fac = Sheets("geo").[T_facility]
            iCountLineFac = UBound(T_fac, 1)
            bHaveToDo = True
        End If
        iNbMaxF = UBound(T_fac, 1)
        F_GEO.LBL_Adm4.Caption = Sheets("geo").[T_adm3].ListObject.HeaderRowRange.Item(2).value
        F_GEO.LBL_Adm4F.Caption = Sheets("geo").[T_facility].ListObject.HeaderRowRange.Item(2).value
    End If

    'creation du tableau concat
    If IsEmptyTable(T_concat) Or bHaveToDo Then
        ReDim T_concat(iNbMax)                   'attention a la desynchro de -1
        i = 1
        While i <= iNbMax
            Select Case iLastLevel
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
        If Not IsEmptyTable(T_concat) Then
            '[F_GEO].LST_ListeAgre.list = TriBulle(T_concat)
            Call QuickSort(T_concat, LBound(T_concat), UBound(T_concat))
            [F_GEO].LST_ListeAgre.List = T_concat
        End If
    Else
        [F_GEO].LST_ListeAgre.List = T_concat
    End If

    If IsEmptyTable(T_concatF) Or bHaveToDo Then
        ReDim T_concatF(iNbMaxF)                 'meme manip pour facility
        i = 1
        While i <= iNbMaxF
            Select Case iLastLevel
            Case 0
                T_concatF(i - 1) = T_geo0(i, 1)
        
            Case 1
                'T_concatF(i - 1) = T_geo1(i, 1) & " | " & T_geo1(i, 2)
                T_concatF(i - 1) = T_geo1(i, 2) & " | " & T_geo1(i, 1)
    
            Case 2
                'T_concatF(i - 1) = T_fac(i, 1) & " | " & T_fac(i, 2)
                T_concatF(i - 1) = T_fac(i, 2) & " | " & T_fac(i, 1)
                j = 1
                While j <= UBound(T_geo2, 1)
                    If T_fac(i, 1) = T_geo2(j, 2) Then
                        'T_concatF(i - 1) = T_geo2(j, 1) & " | " & T_concatF(i - 1)
                        T_concatF(i - 1) = T_concatF(i - 1) & " | " & T_geo2(j, 1)
                    End If
                    j = j + 1
                Wend
        
            Case 3
                'T_concatF(i - 1) = T_fac(i, 1) & " | " & T_fac(i, 2)
                T_concatF(i - 1) = T_fac(i, 2) & " | " & T_fac(i, 1)
            
                j = 1
                While j <= UBound(T_geo2, 1)
                    If T_geo2(j, 2) = T_fac(i, 1) Then
                        'T_concatF(i - 1) = T_geo2(j, 1) & " | " & T_concatF(i - 1)
                        T_concatF(i - 1) = T_concatF(i - 1) & " | " & T_geo2(j, 1)
                        k = 1
                        While k <= UBound(T_geo1, 1)
                            If T_geo1(k, 2) = T_geo2(j, 1) Then
                                'T_concatF(i - 1) = T_geo1(k, 1) & " | " & T_concatF(i - 1)
                                T_concatF(i - 1) = T_concatF(i - 1) & " | " & T_geo1(k, 1)
                            End If
                            k = k + 1
                        Wend
                    End If
                    j = j + 1
                Wend
            End Select
            i = i + 1
        Wend
        If Not IsEmptyTable(T_concatF) Then
            '[F_GEO].LST_ListeAgreF.list = TriBulle(T_concatF)
            Call QuickSort(T_concatF, LBound(T_concatF), UBound(T_concatF))
            [F_GEO].LST_ListeAgreF.List = T_concatF
        End If
    Else
        [F_GEO].LST_ListeAgreF.List = T_concatF
    End If

    'Creation de l'histo
    If Not Sheets("geo").[T_HistoGeo].ListObject.DataBodyRange Is Nothing Then
        i = 1
        ReDim T_histo(i)
        If [T_HistoGeo].Count > 1 Then
            T_histo = [T_HistoGeo]
        Else
            T_histo(0) = [T_HistoGeo]
        End If
        [F_GEO].LST_Histo.List = T_histo
    End If

    If Not Sheets("geo").[T_HistoFacil].ListObject.DataBodyRange Is Nothing Then
        i = 1
        ReDim T_histoF(i)
        If [T_HistoFacil].Count > 1 Then
            T_histoF = [T_HistoFacil]
        Else
            T_histoF(0) = [T_HistoFacil]
        End If
        [F_GEO].LST_HistoF.List = T_histoF
    End If

    '
    Select Case iGeoType
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

    'Call TranslateForm("F_Geo")
    'the show must go on
    [F_GEO].Show

End Sub

Sub ShowLst2(sPlace As String)

    Dim i As Integer
    Dim j As Integer

    [F_GEO].LST_Adm2.Clear
    [F_GEO].LST_Adm3.Clear
    [F_GEO].LST_Adm4.Clear

    If Not IsEmptyTable(T_geo1) Then
        i = 1
        j = 0
        ReDim T_aff2(0)
        While i <= UBound(T_geo1)
            If T_geo1(i, 1) = sPlace Then
                ReDim Preserve T_aff2(j)
                T_aff2(j) = T_geo1(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
    
    If Not IsEmptyTable(T_aff2) Then
        [F_GEO].LST_Adm2.List = T_aff2
        [F_GEO].TXT_Msg.value = sPlace
    Else
        [F_GEO].TXT_Msg.value = sPlace           '& " : Pas de niveau2"
    End If

End Sub

Sub ShowLstF2(sPlace As String)

    Dim i As Integer
    Dim j As Integer
    Dim bFound As Boolean

    [F_GEO].LST_AdmF2.Clear
    [F_GEO].LST_AdmF3.Clear
    [F_GEO].LST_AdmF4.Clear

    bFound = False
    If Not IsEmptyTable(T_fac) Then
        i = 1
        j = 0
        ReDim T_aff2(0)
        While i <= UBound(T_fac)
            If T_fac(i, 1) = sPlace Then
                ReDim Preserve T_aff2(j)
                T_aff2(j) = T_fac(i, 2)
                bFound = True
                j = j + 1
            End If
            i = i + 1
        Wend
        If Not bFound Then
            i = 1
            j = 0
            ReDim T_aff2(0)
            While i <= UBound(T_geo1)
                If T_geo1(i, 1) = sPlace Then
                    ReDim Preserve T_aff2(j)
                    T_aff2(j) = T_geo1(i, 2)
                    j = j + 1
                End If
                i = i + 1
            Wend
        End If
    Else
        i = 1
        j = 0
        ReDim T_aff2(0)
        While i <= UBound(T_geo1)
            If T_geo1(i, 1) = sPlace Then
                ReDim Preserve T_aff2(j)
                T_aff2(j) = T_geo1(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
    
    If Not IsEmptyTable(T_aff2) Then
        [F_GEO].LST_AdmF2.List = T_aff2
        [F_GEO].TXT_Msg.value = sPlace
    Else
        [F_GEO].TXT_Msg.value = sPlace           '& " : Pas de niveau2"
    End If

End Sub

Sub ShowLst3(sPlace As String)

    Dim i As Integer
    Dim j As Integer

    [F_GEO].LST_Adm3.Clear
    [F_GEO].LST_Adm4.Clear

    If Not IsEmptyTable(T_geo2) Then
        i = 1
        j = 0
        ReDim T_aff3(0)
        While i <= UBound(T_geo2)
            If T_geo2(i, 1) = sPlace Then
                ReDim Preserve T_aff3(j)
                T_aff3(j) = T_geo2(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
    
    If Not IsEmptyTable(T_aff3) Then
        [F_GEO].LST_Adm3.List = T_aff3
        [F_GEO].TXT_Msg.value = [F_GEO].LST_Adm1.value & " | " & [F_GEO].LST_Adm2.value
    Else
        [F_GEO].TXT_Msg.value = [F_GEO].LST_Adm1.value & " | " & [F_GEO].LST_Adm2.value '& " : Pas de niveau 3"
    End If

End Sub

Sub ShowLstF3(sPlace As String)

    Dim i As Integer
    Dim j As Integer
    Dim bFound As Boolean

    [F_GEO].LST_AdmF3.Clear
    [F_GEO].LST_AdmF4.Clear

    bFound = False
    If Not IsEmptyTable(T_fac) Then
        i = 1
        j = 0
        ReDim T_aff3(0)
        While i <= UBound(T_fac)
            If T_fac(i, 1) = sPlace Then
                ReDim Preserve T_aff3(j)
                T_aff3(j) = T_fac(i, 2)
                bFound = True
                j = j + 1
            End If
            i = i + 1
        Wend
        If Not bFound Then
            i = 1
            j = 0
            ReDim T_aff3(0)
            While i <= UBound(T_geo2)
                If T_geo2(i, 1) = sPlace Then
                    ReDim Preserve T_aff3(j)
                    T_aff3(j) = T_geo2(i, 2)
                    j = j + 1
                End If
                i = i + 1
            Wend
        End If
    Else
        i = 1
        j = 0
        ReDim T_aff3(0)
        While i <= UBound(T_geo2)
            If T_geo2(i, 1) = sPlace Then
                ReDim Preserve T_aff3(j)
                T_aff3(j) = T_geo2(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
    
    If Not IsEmptyTable(T_aff3) Then
        [F_GEO].LST_AdmF3.List = T_aff3
        [F_GEO].TXT_Msg.value = [F_GEO].LST_AdmF2.value & " | " & [F_GEO].LST_AdmF1.value
    Else
        [F_GEO].TXT_Msg.value = [F_GEO].LST_AdmF2.value & " | " & [F_GEO].LST_AdmF1.value '& " : Pas de niveau 3"
    End If

End Sub

Sub ShowLst4(sPlace As String)

    Dim i As Integer
    Dim j As Integer

    [F_GEO].LST_Adm4.Clear

    If Not IsEmptyTable(T_geo3) Then
        i = 1
        j = 0
        ReDim T_aff4(0)
        While i <= UBound(T_geo3)
            If T_geo3(i, 1) = sPlace Then
                ReDim Preserve T_aff4(j)
                T_aff4(j) = T_geo3(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If
    
    If Not IsEmptyTable(T_aff4) Then
        [F_GEO].LST_Adm4.List = T_aff4
        [F_GEO].TXT_Msg.value = [F_GEO].LST_Adm1.value & " | " & [F_GEO].LST_Adm2.value & " | " & [F_GEO].LST_Adm3.value
    
    Else
        [F_GEO].TXT_Msg.value = [F_GEO].LST_Adm1.value & " | " & [F_GEO].LST_Adm2.value & " | " & [F_GEO].LST_Adm3.value '& " : Pas de niveau 4"
    End If

End Sub

Sub ShowLstF4(sPlace As String)

    Dim i As Integer
    Dim j As Integer
    Dim bFound As Boolean

    [F_GEO].LST_AdmF4.Clear

    bFound = False
    If Not IsEmptyTable(T_fac) Then
        i = 1
        j = 0
        ReDim T_aff4(0)
        While i <= UBound(T_fac)
            If T_fac(i, 1) = sPlace Then
                ReDim Preserve T_aff4(j)
                T_aff4(j) = T_fac(i, 2)
                bFound = True
                j = j + 1
            End If
            i = i + 1
        Wend
        If Not bFound Then
            i = 1
            j = 0
            ReDim T_aff4(0)
            While i <= UBound(T_geo3)
                If T_geo3(i, 1) = sPlace Then
                    ReDim Preserve T_aff4(j)
                    T_aff4(j) = T_geo3(i, 2)
                    j = j + 1
                End If
                i = i + 1
            Wend
        End If
    Else
        i = 1
        j = 0
        ReDim T_aff4(0)
        While i <= UBound(T_geo3)
            If T_geo3(i, 1) = sPlace Then
                ReDim Preserve T_aff4(j)
                T_aff4(j) = T_geo3(i, 2)
                j = j + 1
            End If
            i = i + 1
        Wend
    End If

    If Not IsEmptyTable(T_aff4) Then
        [F_GEO].LST_AdmF4.List = T_aff4
        [F_GEO].TXT_Msg.value = [F_GEO].LST_AdmF3.value & " | " & [F_GEO].LST_AdmF2.value & " | " & [F_GEO].LST_AdmF1.value
    Else
        [F_GEO].TXT_Msg.value = [F_GEO].LST_AdmF3.value & " | " & [F_GEO].LST_AdmF2.value & " | " & [F_GEO].LST_AdmF1.value '& " : Pas de niveau 4"
    End If

End Sub

Sub ClearGeo()

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

    [F_GEO].TXT_Msg.value = ""

End Sub

Sub SearchValue(T_concat, sSearchedValue As String)

    Dim T_result
    Dim i As Long
    Dim j As Long

    '[F_Geo].LST_ListeAgre.Clear

    If Len(sSearchedValue) >= 3 Then
        i = 0
        j = 0
        ReDim T_result(j)
        While i <= UBound(T_concat)
            If InStr(1, LCase(T_concat(i)), LCase(sSearchedValue)) > 0 Then
                ReDim Preserve T_result(j)
                T_result(j) = T_concat(i)
                j = j + 1
            End If
            i = i + 1
        Wend
    
        If Not IsEmptyTable(T_result) Then
            '[F_GEO].LST_ListeAgre.list = TriBulle(T_result)
            Call QuickSort(T_result, LBound(T_result), UBound(T_result))
            [F_GEO].LST_ListeAgre.List = T_result
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

Sub SeachHistoValue(T_histo, sSearchedValue As String)

    Dim T_result
    Dim i As Long
    Dim j As Long

    '[F_Geo].LST_ListeAgre.Clear

    If Len(sSearchedValue) >= 3 Then
        i = 0
        j = 0
        ReDim T_result(j)
        While i <= UBound(T_histo)
            If InStr(1, LCase(T_histo(i)), LCase(sSearchedValue)) > 0 Then
                ReDim Preserve T_result(j)
                T_result(j) = T_histo(i)
                j = j + 1
            End If
            i = i + 1
        Wend
    
        If Not IsEmptyTable(T_result) Then
            '[F_GEO].LST_Histo.list = TriBulle(T_result)
            Call QuickSort(T_result, LBound(T_result), UBound(T_result))
            [F_GEO].LST_Histo.List = T_result
        Else
            If [F_GEO].LST_Histo.ListCount - 1 <> UBound(T_histo) Then
                [F_GEO].LST_Histo.List = T_histo
            End If
        End If
    Else
        If [F_GEO].LST_Histo.ListCount - 1 <> UBound(T_histo) Then
            [F_GEO].LST_Histo.List = T_histo
        End If
    End If

End Sub

Sub SearchValueF(T_concatF, sSearchedValue As String)

    Dim T_result
    Dim i As Long
    Dim j As Long

    If Len(sSearchedValue) >= 3 Then
        i = 0
        j = 0
        ReDim T_result(j)
        While i <= UBound(T_concatF)
            If InStr(1, LCase(T_concatF(i)), LCase(sSearchedValue)) > 0 Then
                ReDim Preserve T_result(j)
                T_result(j) = T_concatF(i)
                j = j + 1
            End If
            i = i + 1
        Wend
    
        If Not IsEmptyTable(T_result) Then
            '[F_GEO].LST_ListeAgreF.list = TriBulle(T_result)
            Call QuickSort(T_result, LBound(T_result), UBound(T_result))
            [F_GEO].LST_ListeAgreF.List = T_result
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

Sub SeachHistoValueF(T_histoF, sSearchedValue As String)

    Dim T_result
    Dim i As Long
    Dim j As Long

    If Len(sSearchedValue) >= 3 Then
        i = 0
        j = 0
        ReDim T_result(j)
        While i <= UBound(T_histoF)
            If InStr(1, LCase(T_histoF(i)), LCase(sSearchedValue)) > 0 Then
                ReDim Preserve T_result(j)
                T_result(j) = T_histoF(i)
                j = j + 1
            End If
            i = i + 1
        Wend
    
        If Not IsEmptyTable(T_result) Then
            '[F_GEO].LST_HistoF.list = TriBulle(T_result)
            Call QuickSort(T_result, LBound(T_result), UBound(T_result))
            [F_GEO].LST_HistoF.List = T_result
        Else
            If [F_GEO].LST_HistoF.ListCount - 1 <> UBound(T_histoF) Then
                [F_GEO].LST_HistoF.List = T_histo
            End If
        End If
    Else
        If [F_GEO].LST_HistoF.ListCount - 1 <> UBound(T_histoF) Then
            [F_GEO].LST_HistoF.List = T_histo
        End If
    End If

End Sub

Function ReverseString(sChaine As String)

    Dim i As Integer
    Dim T_temp
    Dim sRes As String

    ReverseString = ""
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
    ReverseString = sRes

End Function

