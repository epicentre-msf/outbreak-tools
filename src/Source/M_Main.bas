Attribute VB_Name = "M_Main"
Option Explicit

Const C_NomFeuilleSource As String = "Variables"
Const C_NomFeuilleChoix = "choices"

Sub ChargerFichierDico()

Dim sCheminFichier As String

sCheminFichier = chargerChemin

If sCheminFichier <> "" Then
    [RNG_Dico].Value = sCheminFichier

    [RNG_Msg].Value = TraduireMSG("MSG_ChemFich")
    [RNG_Dico].Interior.Color = vbWhite
Else
    [RNG_Msg].Value = TraduireMSG("MSG_FichNonTr")
End If

End Sub

Sub ChargerFichierGeo()

Dim sCheminFichier As String
Dim xlsApp As New Excel.Application
Dim oFeuille As Object
Dim T_Adm
Dim iDerLigne As Long
Dim i As Long
Dim j As Long

sCheminFichier = chargerChemin

If sCheminFichier <> "" Then
    With xlsApp
        .ScreenUpdating = False
        .Workbooks.Open sCheminFichier
        
        'un coup de menage sur les precedentes data
        [RNG_Msg].Value = TraduireMSG("MSG_NetoPrec")
        i = 0
        While i <= 3
            If Not Sheets("GEO").ListObjects("T_adm" & i).DataBodyRange Is Nothing Then
                Sheets("GEO").ListObjects("T_adm" & i).DataBodyRange.Delete
            End If
        
            i = i + 1
        Wend
        If Not Sheets("GEO").ListObjects("T_facility").DataBodyRange Is Nothing Then
            Sheets("GEO").ListObjects("T_facility").DataBodyRange.Delete
        End If
        
        'et on repompe le tout...
        For Each oFeuille In xlsApp.Worksheets
            [RNG_Msg].Value = TraduireMSG("MSG_EnCours") & oFeuille.Name
            iDerLigne = oFeuille.Cells(1, 1).End(xlDown).Row
            ReDim T_Adm(iDerLigne, 1)
            j = 0
            i = 2
            While i <= iDerLigne
                T_Adm(j, 0) = oFeuille.Cells(i, 1).Value
                T_Adm(j, 1) = oFeuille.Cells(i, 2).Value
                i = i + 1
                j = j + 1
            Wend
            
            If InStr(1, oFeuille.Name, "FACILITY") = 0 Then
                If Not Sheets("GEO").ListObjects("T_" & Left(oFeuille.Name, 4)).DataBodyRange Is Nothing Then
                    Sheets("GEO").ListObjects("T_" & Left(oFeuille.Name, 4)).DataBodyRange.Delete
                End If

                Sheets("GEO").ListObjects("T_" & Left(oFeuille.Name, 4)).Resize Range(Cells(LBound(T_Adm, 2) + 1, Sheets("GEO").ListObjects("T_" & Left(oFeuille.Name, 4)).Range.Column), Cells(UBound(T_Adm, 1), Sheets("GEO").ListObjects("T_" & Left(oFeuille.Name, 4)).Range.Column + 1))
                Sheets("GEO").ListObjects("T_" & Left(oFeuille.Name, 4)).DataBodyRange = T_Adm
            Else
                If Not Sheets("GEO").ListObjects("T_facility").DataBodyRange Is Nothing Then
                    Sheets("GEO").ListObjects("T_facility").DataBodyRange.Delete
                End If
            
                Sheets("GEO").ListObjects("T_facility").Resize Range(Cells(LBound(T_Adm, 2) + 1, Sheets("GEO").ListObjects("T_facility").Range.Column), Cells(UBound(T_Adm, 1), Sheets("GEO").ListObjects("T_facility").Range.Column + 1))
                Sheets("GEO").ListObjects("T_facility").DataBodyRange = T_Adm
                Sheets("GEO").ListObjects("T_facility").HeaderRowRange(1) = oFeuille.Cells(1, 1).Value 'pour savoir a quel niveau est rattaché le facility
            End If
        Next
        
        Sheets("MAIN").Range("RNG_GEO").Value = .ActiveWorkbook.Name
        
        .ScreenUpdating = True
        .Workbooks.Close
        xlsApp.Quit
        Set xlsApp = Nothing
            
        [RNG_Msg].Value = TraduireMSG("MSG_Fini")
        
    End With
Else
    [RNG_Msg].Value = TraduireMSG("MSG_OpeAnnule")

End If

If Not Sheets("GEO").ListObjects("T_HistoGeo").DataBodyRange Is Nothing Then
    Sheets("GEO").ListObjects("T_HistoGeo").DataBodyRange.Delete
End If
If Not Sheets("GEO").ListObjects("T_HistoFacil").DataBodyRange Is Nothing Then
    Sheets("GEO").ListObjects("T_HistoFacil").DataBodyRange.Delete
End If

ReDim T_geo0(0)
ReDim T_aff0(0)
ReDim T_geo1(0)
ReDim T_aff1(0)
ReDim T_geo2(0)
ReDim T_aff2(0)
ReDim T_geo3(0)
ReDim T_aff3(0)
ReDim T_concat(0)
ReDim T_Histo(0)
ReDim T_fac(0)
ReDim T_HistoF(0)
ReDim T_concatF(0)

End Sub

Sub GenererData()

Dim xlsApp As New Excel.Application
Dim D_enteteDic As Scripting.Dictionary
Dim T_dataDic
Dim D_Choix As Scripting.Dictionary
Dim T_Choix
    
    Application.DisplayAlerts = False
    Sheets("Main").Range("a1").Select
    Call AffMasquerBtnValidation(False)
    
    'On Error GoTo ErrLectureFichier
    '
    'Set xlsApp = New Excel.Application
    xlsApp.Workbooks.Open [RNG_Dico].Value
    xlsApp.ScreenUpdating = False
    xlsApp.Visible = False
    [RNG_Msg].Value = TraduireMSG("MSG_LectDico")
    Set D_enteteDic = CreateDicoColVar(xlsApp, C_NomFeuilleSource, 2)
    T_dataDic = CreateTabDataVar(xlsApp, C_NomFeuilleSource, D_enteteDic, 3)
    
    [RNG_Msg].Value = TraduireMSG("MSG_LectListe")
    Set D_Choix = CreateDicoColChoi(xlsApp, C_NomFeuilleChoix)
    T_Choix = CreateTabDataChoi(xlsApp, C_NomFeuilleChoix)

    xlsApp.ActiveWorkbook.Close
    xlsApp.Quit
    Set xlsApp = Nothing
    
    'On Error GoTo errCreatLL
    '
    [RNG_Msg].Value = TraduireMSG("MSG_CreationLL")
    Call BuildListe(D_enteteDic, T_dataDic, D_Choix, T_Choix)
    
    [RNG_Msg].Value = TraduireMSG("MSG_toutFbie")
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrLectureFichier:
    '[RNG_Msg].Value = "Une erreur s'est produite à la lecture du dico"
    'Exit Sub
    
errCreatLL:
    '[RNG_Msg].Value = "Une erreur s'est produite à la création de la LineList"
    'Exit Sub
    
End Sub

Sub AnnulerGenerer()

Sheets("Main").Shapes("SHP_CtrlNouv").Visible = True

Sheets("Main").Range("a1").Select
Call AffMasquerBtnValidation(False)

End Sub

Sub CtrlNouveau()

Call AffMasquerBtnValidation(False)
If [RNG_Dico].Value <> "" Then
    If Dir([RNG_Dico].Value) <> "" Then
        If [RNG_Geo].Value <> "" Then
            If Not ClasseurEstOuvert([RNG_Dico].Value) Then
                [RNG_Msg].Value = TraduireMSG("MSG_ToutEstBon")
                Call AffMasquerBtnValidation(True)  '
                [RNG_Geo].Interior.Color = vbWhite
                [RNG_Dico].Interior.Color = vbWhite
            Else
                [RNG_Msg].Value = TraduireMSG("MSG_FermerDico")
            End If
        Else
            [RNG_Msg].Value = TraduireMSG("MSG_VeriFichGeo")
            [RNG_Geo].Interior.Color = retourneCouleur("RougeEpi")
        End If
    Else
        [RNG_Msg].Value = TraduireMSG("MSG_VeriChemDico")
        [RNG_Dico].Interior.Color = retourneCouleur("RougeEpi")
    
    End If
Else
    [RNG_Msg].Value = TraduireMSG("MSG_VeriChemDico")
    [RNG_Dico].Interior.Color = retourneCouleur("RougeEpi")
End If

End Sub

Private Sub AffMasquerBtnValidation(EstVisible As Boolean)

Sheets("Main").Shapes("SHP_Generer").Visible = EstVisible
Sheets("Main").Shapes("SHP_Annuler").Visible = EstVisible
Sheets("Main").Shapes("SHP_validation").Visible = EstVisible

End Sub

Private Function ClasseurEstOuvert(sNomClasseur As String) As Boolean
       
Dim oWks As Object
Dim i As Byte

ClasseurEstOuvert = False
i = 1
While i <= Application.Workbooks.Count
    Set oWks = Application.Workbooks(i)
    If oWks.Path = sNomClasseur Then
        ClasseurEstOuvert = True
        Exit Function
    End If
    Set oWks = Nothing
    i = i + 1
Wend

End Function

Sub AfficherBtnValidation()

Call AffMasquerBtnValidation(False)

End Sub
