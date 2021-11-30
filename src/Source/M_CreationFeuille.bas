Attribute VB_Name = "M_CreationFeuille"
Option Explicit

Const C_ligneDebEntete1 As Byte = 3
Const C_ligneDebEntete2 As Byte = 4
Const C_ligneTitre As Byte = 5
Const C_ligneDeb As Byte = 6

Sub BuildListe(D_enteteDic As Scripting.Dictionary, T_dataDic, D_Choix As Scripting.Dictionary, T_Choix)
'c'est parti pour le fichier de sortie !

Dim xlsApp As New Excel.Application
Dim sCheminSortie As String
Dim i As Integer    'cpt result
Dim j As Integer    'cpt source
Dim l As Integer    'cpt nbcolonne
Dim T_NbCol 'tab nb colonne dans une feuille determinée
Dim sPrecNomFeuille As String

Dim oCellule As Object  'pour colorier entre les titres 2
Dim iColPrecS1 As Integer
Dim iColPrecS2 As Integer
Dim sTitre1 As String
Dim sTitre2 As String

Dim iNbDeci As Integer  'pour le calcul de décimal
Dim k As Integer
Dim sNbDeci As String

Dim sListeValidation As String  'renvoi la liste de validation

Dim bBtnGeoExist As Boolean 'controle lexistence de bouton geo pour sa creation

Dim oCle As Variant 'cpt dico attention typage special !
Dim oLstobj As Object
Dim oFeuille As Object

Dim iDebPrecS1 As Integer

With xlsApp
    .DisplayAlerts = False
    .ScreenUpdating = True
    .Visible = True
    .Workbooks.Add
    .ActiveWorkbook.VBProject.References.AddFromFile ("C:\windows\system32\scrrun.dll")     'on coche le scripting runtime pour acceder au dico
    
    DoEvents
    MkDir ("C:\LineListeApp\")
    DoEvents
    Call balanceTonCode(xlsApp, "M_LineList")
    
    Call balanceTonFrm(xlsApp, "F_Geo")
    Call balanceTonFrm(xlsApp, "F_NomVisible")
    Call balanceTonFrm(xlsApp, "F_Export")
    
    Call balanceTonCode(xlsApp, "M_Geo")
    Call balanceTonCode(xlsApp, "M_FonctionsTransf")
    DoEvents
    'Call EcrireEventOuverture(xlsApp)
    'DoEvents
    
    'deplacement de la geo
    ThisWorkbook.Sheets("GEO").Copy
    DoEvents
    ActiveWorkbook.SaveAs "C:\LineListeApp\tampon.xlsx"   'puisqu'on ne peut pas balancer une feuille vers une autre instance, on crée un fichier tampon
    ActiveWorkbook.Close
    DoEvents
    .Workbooks.Open Filename:="C:\LineListeApp\tampon.xlsx", UpdateLinks:=False
    
    .Sheets("GEO").Select
    .Sheets("GEO").Copy After:=.Workbooks(1).Sheets(1)
    DoEvents
    .Workbooks("tampon.xlsx").Close
    DoEvents
    
    Kill "C:\LineListeApp\tampon.xlsx"
    RmDir ("C:\LineListeApp\")
    
    'bascule du dico
    .Sheets.Add.Name = "Dico"
    i = 1
    For Each oCle In D_enteteDic.Keys
        .Sheets("dico").Cells(1, i).Value = oCle
        i = i + 1
    Next oCle
    .Sheets("dico").Range("A2").Resize(D_enteteDic.Count, UBound(T_dataDic, 2) + 1) = .WorksheetFunction.Transpose(T_dataDic)
    .Sheets("dico").Visible = False

    'premier tour pour la création des feuilles
    i = 1
    j = 0
    ReDim T_NbCol(j)
    sPrecNomFeuille = ""
    While i <= UBound(T_dataDic, 2)
        If sPrecNomFeuille <> T_dataDic(D_enteteDic("form_name") - 1, i) Then
            j = j + 1
            ReDim Preserve T_NbCol(j)
            T_NbCol(j) = 1
            .Worksheets.Add.Name = T_dataDic(D_enteteDic("form_name") - 1, i)
            .Worksheets(T_dataDic(D_enteteDic("form_name") - 1, i)).Rows("1:2").RowHeight = 22.5
                
            .ActiveWindow.SplitColumn = 2
            .ActiveWindow.SplitRow = 5
            .ActiveWindow.FreezePanes = True
            
            sPrecNomFeuille = T_dataDic(D_enteteDic("form_name") - 1, i)
            ReDim Preserve T_NbCol(UBound(T_NbCol) + 1)
        Else
            T_NbCol(j) = T_NbCol(j) + 1
        End If
        i = i + 1
    Wend

End With

sPrecNomFeuille = ""
sTitre1 = ""
sTitre2 = ""

iColPrecS1 = 1
iColPrecS2 = 1

j = 1   'cpt result
i = 0   'cpt dic
l = 0   'cpt nb col
While i <= UBound(T_dataDic, 2)
    With xlsApp.Sheets(T_dataDic(D_enteteDic("form_name") - 1, i))
        'creation du listobject
        If sPrecNomFeuille <> T_dataDic(D_enteteDic("form_name") - 1, i) Then
            If sPrecNomFeuille <> "" Then   'on conclue les titres de la feuille prec
                xlsApp.Sheets(sPrecNomFeuille).Range(xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete1, iColPrecS1), xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete1, j - 1)).Merge 'titre1
                xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete1, iColPrecS1).MergeArea.HorizontalAlignment = xlCenter
                xlsApp.Sheets(sPrecNomFeuille).Range(xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete1, iColPrecS1), xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete1 + 1, j - 1)).Interior.Color = retourneCouleur("BleuFonceTitre")
                Call TraceLigne(xlsApp.Sheets(sPrecNomFeuille).Range(xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete1, iColPrecS1), xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, j - 1)))
                
                If xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, iColPrecS2) <> "" Then
                    xlsApp.Sheets(sPrecNomFeuille).Range(xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, iColPrecS2), xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, j - 1)).Merge 'titre2
                    xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, iColPrecS1).MergeArea.HorizontalAlignment = xlCenter
                    xlsApp.Sheets(sPrecNomFeuille).Range(xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, iColPrecS2), xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, j - 1)).Interior.Color = retourneCouleur("BleuClairTitre")
                    Call TraceLigne(xlsApp.Sheets(sPrecNomFeuille).Range(xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, iColPrecS2), xlsApp.Sheets(sPrecNomFeuille).Cells(C_ligneDebEntete2, j - 1)))
                End If
            End If
            
            j = 1
            l = l + 1
            sPrecNomFeuille = T_dataDic(D_enteteDic("form_name") - 1, i)
            .ListObjects.Add(xlSrcRange, .Range(.Cells(C_ligneTitre, 1), .Cells(C_ligneTitre, T_NbCol(l))), , xlYes).Name = "o" & T_dataDic(D_enteteDic("form_name") - 1, i)
            .ListObjects("o" & T_dataDic(D_enteteDic("form_name") - 1, i)).TableStyle = "TableStyleLight16"
            
            .Cells.Font.Size = 9
                        
            iDebPrecS1 = 1
            sTitre1 = T_dataDic(D_enteteDic("section_1") - 1, i)    'prem titre
            .Cells(C_ligneDebEntete1, j).Value = T_dataDic(D_enteteDic("section_1") - 1, i)
            
            bBtnGeoExist = False
        End If
        
        'définition d'entete
        .Cells(C_ligneTitre, j).Name = T_dataDic(D_enteteDic("name") - 1, i)
        .Cells(C_ligneTitre, j).Value = T_dataDic(D_enteteDic("label_1") - 1, i)
        .Cells(C_ligneTitre, j).VerticalAlignment = xlTop
        If T_dataDic(D_enteteDic("label_2") - 1, i) <> "" Then
            .Cells(C_ligneTitre, j).Value = .Cells(C_ligneTitre, j).Value & Chr(10) & T_dataDic(D_enteteDic("label_2") - 1, i)
            .Cells(C_ligneTitre, j).Characters(Start:=Len(T_dataDic(D_enteteDic("label_1") - 1, i)) + 1, Length:=Len(T_dataDic(D_enteteDic("label_2") - 1, i)) + 1).Font.Size = 8
            .Cells(C_ligneTitre, j).Characters(Start:=Len(T_dataDic(D_enteteDic("label_1") - 1, i)) + 1, Length:=Len(T_dataDic(D_enteteDic("label_2") - 1, i)) + 1).Font.Color = retourneCouleur("Gris")
        End If
        
        If T_dataDic(D_enteteDic("note") - 1, i) <> "" Then
            .Cells(C_ligneTitre, j).AddComment
            .Cells(C_ligneTitre, j).Comment.Text Text:=T_dataDic(D_enteteDic("note") - 1, i)
            .Cells(C_ligneTitre, j).Comment.Visible = False
        End If
        
        
        'titres
        'cas particulier géo
        If LCase(T_dataDic(D_enteteDic("control") - 1, i)) = "geo" Then
            If T_dataDic(D_enteteDic("section_2") - 1, i) = "" Then
                T_dataDic(D_enteteDic("section_2") - 1, i) = T_dataDic(D_enteteDic("label_1") - 1, i)
            End If
        End If
        
        If sTitre1 <> T_dataDic(D_enteteDic("section_1") - 1, i) Then
            'si le titre change, on fusionne les prec cellules
            .Cells(C_ligneDebEntete1, j).Value = T_dataDic(D_enteteDic("section_1") - 1, i)
            sTitre1 = T_dataDic(D_enteteDic("section_1") - 1, i)
            
            .Range(.Cells(C_ligneDebEntete1, iColPrecS1), .Cells(C_ligneDebEntete1, j - 1)).Merge
            .Cells(C_ligneDebEntete1, iColPrecS1).MergeArea.HorizontalAlignment = xlCenter
            .Range(.Cells(C_ligneDebEntete1, iColPrecS1), .Cells(C_ligneDebEntete1, j - 1)).Interior.Color = retourneCouleur("BleuFonceTitre")
            For Each oCellule In .Range(.Cells(C_ligneDebEntete2, iColPrecS1), .Cells(C_ligneDebEntete2, j - 1))   'coloriage
                'If oCellule.Interior.Color <> retourneCouleur("BleuClairTitre") Then
                If oCellule.Value = "" Then
                    oCellule.Interior.Color = retourneCouleur("BleuFonceTitre")
                End If
            Next
            Set oCellule = Nothing
            Call TraceLigne(.Range(.Cells(C_ligneDebEntete1, iColPrecS1), .Cells(C_ligneDebEntete2, j - 1)))
            
            iColPrecS1 = j
            
        End If
        
        If sTitre2 <> T_dataDic(D_enteteDic("section_2") - 1, i) Then
            'si le titre change, on fusionne les prec cellules
            .Cells(C_ligneDebEntete2, j).Value = T_dataDic(D_enteteDic("section_2") - 1, i)
            
            sTitre2 = T_dataDic(D_enteteDic("section_2") - 1, i)
            If j > 1 Then
                .Range(.Cells(C_ligneDebEntete2, iColPrecS2), .Cells(C_ligneDebEntete2, j - 1)).Merge
                .Cells(C_ligneDebEntete2, iColPrecS2).MergeArea.HorizontalAlignment = xlCenter
                .Range(.Cells(C_ligneDebEntete2, iColPrecS2), .Cells(C_ligneDebEntete2, j - 1)).Interior.Color = retourneCouleur("BleuClairTitre")
                If .Cells(C_ligneDebEntete2, iColPrecS2) <> "" Then
                    Call TraceLigne(.Range(.Cells(C_ligneDebEntete2, iColPrecS2), .Cells(C_ligneDebEntete2, j - 1)))
                End If
            Else
                .Cells(C_ligneDebEntete2, iColPrecS2).HorizontalAlignment = xlCenter
                .Cells(C_ligneDebEntete2, iColPrecS2).Interior.Color = retourneCouleur("BleuClairTitre")
                If .Cells(C_ligneDebEntete2, iColPrecS2) <> "" Then
                    Call TraceLigne(.Cells(C_ligneDebEntete2, iColPrecS2))
                End If
            End If
            iColPrecS2 = j
        End If
        
        'champ obligatoire
        If T_dataDic(D_enteteDic("mandatory") - 1, i) = "yes" Then
            If T_dataDic(D_enteteDic("note") - 1, i) <> "" Then
                .Cells(C_ligneTitre, j).Comment.Text Text:="Mandatory data" & Chr(10) & T_dataDic(D_enteteDic("note") - 1, i)
            Else
                .Cells(C_ligneTitre, j).AddComment
                .Cells(C_ligneTitre, j).Comment.Text Text:="Mandatory data"
                .Cells(C_ligneTitre, j).Comment.Visible = False
            End If
        End If
        
        'typage
        If T_dataDic(D_enteteDic("type") - 1, i) <> "" Then
            Select Case LCase(T_dataDic(D_enteteDic("type") - 1, i))
            Case "text"
                .Cells(6, j).NumberFormat = "@"
            Case "date"
                .Cells(6, j).NumberFormat = "d/m/yyyy"
            Case "integer"
                .Cells(6, j).NumberFormat = "0"
            Case Else
                If InStr(1, LCase(T_dataDic(D_enteteDic("type") - 1, i)), "decimal") > 0 Then   'decimal
                    iNbDeci = Right(T_dataDic(D_enteteDic("type") - 1, i), 1)
                    k = 0
                    While k < iNbDeci
                        k = k + 1
                    Wend
                    .Cells(6, j).NumberFormat = "0." & RenvoiChaineDecimal(Right(T_dataDic(D_enteteDic("type") - 1, i), 1))   'sur un carac donc neuf decimal max... et c'est deja pas mal
                End If
            End Select
        End If
        
        'control & liste de validation (validation_alert)
        If T_dataDic(D_enteteDic("control") - 1, i) <> "" Then
            Select Case LCase(T_dataDic(D_enteteDic("control") - 1, i))
            Case "choices"

                If T_dataDic(D_enteteDic("choices") - 1, i) <> "" Then
                    sListeValidation = GetChaineValidation(T_Choix, D_Choix, T_dataDic(D_enteteDic("choices") - 1, i))
                    If sListeValidation <> "" Then
                        Call creerListeValid(.Cells(6, j), sListeValidation, renvoieTypeBlocageValidation(T_dataDic(D_enteteDic("validation_alert") - 1, i)))
                    End If
                End If
                
            Case "formula"
                'la ca va piquer !!!
        
                'un controle de formule ?
        
            Case "geo", "hf"
                'ajouter 3 colonnes plus tard pour geo
                If LCase(T_dataDic(D_enteteDic("control") - 1, i)) = "geo" Then
                    Call Ajouter4ColGeo(xlsApp, T_dataDic(D_enteteDic("form_name") - 1, i), T_dataDic(D_enteteDic("label_1") - 1, i), j)
                    j = j + 3
                End If
                
                If Not bBtnGeoExist Then
                    Call AjoutBtn(xlsApp, CStr(T_dataDic(D_enteteDic("form_name") - 1, i)), .Cells(1, 1).Left, .Cells(1, 1).Top, "SHP_GeoApps", "Geo Apps")
                    .Shapes("SHP_GeoApps").OnAction = "clicBtnGeoApps"
                    bBtnGeoExist = True
                End If
            Case Else
            
            End Select
        
        End If
    
        'min max
        If T_dataDic(D_enteteDic("min") - 1, i) <> "" And T_dataDic(D_enteteDic("max") - 1, i) <> "" Then
            If IsNumeric(T_dataDic(D_enteteDic("min") - 1, i)) And IsNumeric(T_dataDic(D_enteteDic("max") - 1, i)) Then
                Call creerMinMaxValid(.Cells(6, j), T_dataDic(D_enteteDic("min") - 1, i), T_dataDic(D_enteteDic("max") - 1, i), renvoieTypeBlocageValidation(T_dataDic(D_enteteDic("validation_alert") - 1, i)))
            End If
        End If
        
        'visible
        If T_dataDic(D_enteteDic("visible") - 1, i) = "No" Or T_dataDic(D_enteteDic("visible") - 1, i) = "Non" Then
            .Columns(j).EntireColumn.Hidden = True
        End If
        
        .Cells.EntireColumn.AutoFit
    
        j = j + 1
        i = i + 1

    End With
Wend

Application.ActiveWindow.WindowState = xlMinimized

With xlsApp
    For Each oFeuille In .ActiveWorkbook.Sheets   'on se crée les 30 premieres lignes
        If oFeuille.Name <> "GEO" And oFeuille.Name <> "TRANSLATION" And oFeuille.Name <> "Dico" Then
            For Each oLstobj In oFeuille.ListObjects
                oLstobj.Resize oFeuille.Range(oFeuille.Cells(C_ligneTitre, 1), oFeuille.Cells(35, oFeuille.Cells(C_ligneTitre, 1).End(xlToRight).Column))
            Next
        End If
    Next oFeuille

    .ActiveWindow.SplitRow = C_ligneTitre
    .ActiveWindow.FreezePanes = True
    .DisplayAlerts = True
    .ScreenUpdating = True
    .Visible = True
    .ActiveWindow.WindowState = xlMaximized
End With

End Sub

Function renvoieTypeBlocageValidation(sTypeBlocageValidation As Variant) As Byte

    renvoieTypeBlocageValidation = 3 'liste de validation info, warning ou erreur
    If sTypeBlocageValidation <> "" Then
        Select Case LCase(sTypeBlocageValidation)
        Case "warning"
            renvoieTypeBlocageValidation = 2
        Case "error"
            renvoieTypeBlocageValidation = 1
        End Select
    End If
    
End Function

Sub creerListeValid(oRange As Range, sListeValid As Variant, sTypeAlert As Byte)

    With oRange.Validation
        .Delete
        Select Case LCase(sTypeAlert)
        Case 1 '"error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sListeValid
        Case 2 '"warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=sListeValid
        Case Else
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=sListeValid
        End Select
        
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub creerMinMaxValid(oRange As Range, iMin As Variant, iMax As Variant, sTypeAlert As Byte)

    With oRange.Validation
        .Delete
        Select Case LCase(sTypeAlert)
        Case 1 '"error"
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
        Case 2 '"warning"
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
        Case Else
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
        End Select
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub TraceLigne(oRange As Range)

Dim i As Integer

For i = 7 To 10
    With oRange.Borders(i)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
Next

End Sub

Sub AjoutBtn(xlsApp As Excel.Application, sFeuille, iGauche As Integer, iTop As Integer, sNom As String, sTexte As String)

Dim oShape As Object
Dim bShapeExist As Boolean

bShapeExist = False
For Each oShape In xlsApp.Sheets(sFeuille).Shapes
    If oShape.Name = sNom Then
        bShapeExist = True
        Exit For
    End If
Next

If Not bShapeExist Then
    With xlsApp.Sheets(sFeuille)
        .Shapes.AddShape(msoShapeRectangle, iGauche + 3, iTop + 3, 60, 20).Name = sNom
        .Shapes(sNom).Placement = xlMove
        .Shapes(sNom).TextFrame2.TextRange.Characters.Text = sTexte
        .Shapes(sNom).ShapeStyle = msoShapeStylePreset30
    End With
End If

End Sub

Sub Ajouter4ColGeo(xlsApp As Application, sNomFeuille, sLib, iCol As Integer)

Dim i As Byte

With xlsApp.Sheets(sNomFeuille)
    i = 1
    While i <= 3
        .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        'xlsApp.Sheets(sNomFeuille).Cells(C_ligneTitre, iCol + 1).Value = sLib & 4 - i + 1
        .Cells(C_ligneTitre, iCol + 1).Value = Sheets("geo").ListObjects("T_adm" & 4 - i).HeaderRowRange.Item(2).Value
        i = i + 1
    Wend
    .Cells(C_ligneTitre, iCol).Value = Sheets("geo").ListObjects("T_adm0").HeaderRowRange.Item(2).Value
    .Range(.Cells(C_ligneTitre, iCol).Value, .Cells(C_ligneTitre, iCol + 3).Value).Merge
End With

End Sub
