Attribute VB_Name = "M_Main"
Option Explicit

Const C_SheetNameDic As String = "Dictionary"
Const C_SheetNameChoices As String = "Choices"
Const C_SheetNameExport As String = "Exports"

Sub LoadFileDic()

    Dim sFilePath As String

    sFilePath = LoadPathWindow

    If sFilePath <> "" Then
        [RNG_Dico].value = sFilePath

        [RNG_Msg].value = TranslateMsg("MSG_ChemFich")
        [RNG_Dico].Interior.Color = vbWhite
    Else
        [RNG_Msg].value = TranslateMsg("MSG_OpeAnnule")
    End If

End Sub

' Adding a new load geo for the Geo file, in a new sheet called Geo2
' we have two functions for loading the geodatabase, but the second one
' will be in use instead of the first one
Sub LoadGeoFile()
    Dim geoSheet As String
    geoSheet = "GEO"
    Dim sFilePath As String                      'File path to the geo file
    Dim xlsApp As Excel.Application
    Dim oSheet As Object
    Dim T_Adm  As BetterArray                    'Table for admin loading this is a variant
    Dim T_header As BetterArray                  'Table for the headers of the listobjects
    Dim outputRange As Range
    Dim objectName As String 'name of the listobject in the geo sheet
    
    'Defining the adm and headers array
    Set T_Adm = New BetterArray
    Set T_header = New BetterArray
    set xlsApp = New Excel.Application
    
    sFilePath = LoadPathWindow
    
    If sFilePath <> "" Then
        With xlsApp
            .ScreenUpdating = False
            .Workbooks.Open sFilePath
            
            'Cleaning the previous Data in case the ranges are not Empty
            [RNG_Msg].value = TranslateMsg("MSG_NetoPrec")
            
            'Adm
            If Not Sheets(geoSheet).ListObjects("T_Adm").DataBodyRange Is Nothing Then
                Sheets(geoSheet).ListObjects("T_adm").DataBodyRange.Delete
            End If
            'Facility
            If Not Sheets(geoSheet).ListObjects("T_Facility").DataBodyRange Is Nothing Then
                Sheets(geoSheet).ListObjects("T_Facility").DataBodyRange.Delete
            End If
            'Translations
            If Not Sheets(geoSheet).ListObjects("T_GeoTrad").DataBodyRange Is Nothing Then
                Sheets(geoSheet).ListObjects("T_GeoTrad").DataBodyRange.Delete
            End If
            
            'Reloading the data from the Geobase
            For Each oSheet In xlsApp.Worksheets
                [RNG_Msg].value = TranslateMsg("MSG_EnCours") & oSheet.Name
                T_Adm.Clear
                'loading the data in memory
                T_Adm.FromExcelRange oSheet.Range("A2"), True, True
                'The headers
                T_header.FromExcelRange oSheet.Range("A1"), False, True
                
                'keeping the object names for writing the data
                Select Case oSheet.Name
                Case "ADM"
                    objectName = "T_Adm"
                Case "HF"
                    objectName = "T_Facility"
                Case "NAMES"
                    objectName = "T_GeoTrad"
                Case Else
                    [RNG_Msg].value = TranslateMsg("MSG_Error_Sheet") & oSheet.Name
                    Exit Sub
                End Select
                
                ' Check if the sheet is the admin exists sheet before writing in the adm table
                With Sheets(geoSheet).ListObjects(objectName)
                    'writing the data body
                    Set outputRange = Range(Cells(T_Adm.LowerBound + 1, .Range.Column), Cells(T_Adm.UpperBound + 1, .Range.Column + T_Adm.Length))
                    T_Adm.ToExcelRange (outputRange)
                    .Resize outputRange
                    'Writing the headers
                    Set outputRange = .HeaderRowRange
                    T_header.ToExcelRange outputRange
                End With
            Next
            Sheets("MAIN").Range("RNG_GEO").value = .ActiveWorkbook.Name
            
            .ScreenUpdating = True
            .Workbooks.Close
            xlsApp.Quit
            Set xlsApp = Nothing
            [RNG_Msg].value = TranslateMsg("MSG_Fini")
        End With
        
        'Remove the historic of the Geo and the facility if not empty
        If Not Sheets(geoSheet).ListObjects("T_HistoGeo").DataBodyRange Is Nothing Then
            Sheets(geoSheet).ListObjects("T_HistoGeo").DataBodyRange.Delete
        End If
        If Not Sheets(geoSheet).ListObjects("T_HistoHF").DataBodyRange Is Nothing Then
            Sheets(geoSheet).ListObjects("T_HistoHF").DataBodyRange.Delete
        End If
    Else
        [RNG_Msg].value = TranslateMsg("MSG_OpeAnnule")
    End If

End Sub
'
'Sub LoadGeoFile()
'
'    Dim sFilePath As String
'    Dim xlsApp As New Excel.Application
'    Dim oSheet As Object
'    Dim T_Adm
'    Dim iLastLine As Long
'    Dim i As Long
'    Dim j As Long
'    Dim oListO As Object
'
'    sFilePath = LoadPathWindow
'
'    If sFilePath <> "" Then
'        With xlsApp
'            .ScreenUpdating = False
'            .Workbooks.Open sFilePath
'        
'            'un coup de menage sur les precedentes data
'            Sheets("main").Range("RNG_Msg").value = TranslateMsg("MSG_NetoPrec")
'            i = 0
'            While i <= 3
'                If Not Sheets("GEO").ListObjects("T_adm" & i).DataBodyRange Is Nothing Then
'                    Sheets("GEO").ListObjects("T_adm" & i).DataBodyRange.Delete
'                End If
'        
'                i = i + 1
'            Wend
'            If Not Sheets("GEO").ListObjects("T_facility").DataBodyRange Is Nothing Then
'                Sheets("GEO").ListObjects("T_facility").DataBodyRange.Delete
'            End If
'        
'            'et on repompe le tout...
'            For Each oSheet In xlsApp.Worksheets
'                Sheets("main").Range("RNG_Msg").value = TranslateMsg("MSG_EnCours") & oSheet.Name
'                iLastLine = oSheet.Cells(1, 1).End(xlDown).Row
'                ReDim T_Adm(iLastLine, 1)
'                j = 0
'                i = 2
'                While i <= iLastLine
'                    T_Adm(j, 0) = oSheet.Cells(i, 1).value
'                    T_Adm(j, 1) = oSheet.Cells(i, 2).value
'                    i = i + 1
'                    j = j + 1
'                Wend
'            
'                If InStr(1, oSheet.Name, "HF") = 0 And InStr(1, oSheet.Name, "NAMES") = 0 Then
'                    If Not Sheets("GEO").ListObjects("T_" & Left(oSheet.Name, 4)).DataBodyRange Is Nothing Then
'                        Sheets("GEO").ListObjects("T_" & Left(oSheet.Name, 4)).DataBodyRange.Delete
'                    End If
'
'                    Sheets("GEO").ListObjects("T_" & Left(oSheet.Name, 4)).Resize Range(Cells(LBound(T_Adm, 2) + 1, Sheets("GEO").ListObjects("T_" & Left(oSheet.Name, 4)).Range.Column), Cells(UBound(T_Adm, 1), Sheets("GEO").ListObjects("T_" & Left(oSheet.Name, 4)).Range.Column + 1))
'                    Sheets("GEO").ListObjects("T_" & Left(oSheet.Name, 4)).DataBodyRange = T_Adm
'                    Sheets("GEO").ListObjects("T_" & Left(oSheet.Name, 4)).HeaderRowRange(2) = oSheet.Cells(1, 2).value
'                
'                ElseIf InStr(1, oSheet.Name, "HF") > 0 Then
'                    If Not Sheets("GEO").ListObjects("T_facility").DataBodyRange Is Nothing Then
'                        Sheets("GEO").ListObjects("T_facility").DataBodyRange.Delete
'                    End If
'            
'                    Sheets("GEO").ListObjects("T_facility").Resize Range(Cells(LBound(T_Adm, 2) + 1, Sheets("GEO").ListObjects("T_facility").Range.Column), Cells(UBound(T_Adm, 1), Sheets("GEO").ListObjects("T_facility").Range.Column + 1))
'                    Sheets("GEO").ListObjects("T_facility").DataBodyRange = T_Adm
'                    Sheets("GEO").ListObjects("T_facility").HeaderRowRange(1) = oSheet.Cells(1, 1).value 'pour savoir a quel niveau est rattach� le facility
'                    Sheets("GEO").ListObjects("T_facility").HeaderRowRange(2) = oSheet.Cells(1, 2).value
'                ElseIf InStr(1, oSheet.Name, "NAMES") > 0 Then
'                    ReDim T_Adm(iLastLine, 2)
'                    i = 1
'                    While i <= Sheets(oSheet.Name).Cells(1, 1).End(xlDown).Row
'                        T_Adm(j, 0) = oSheet.Cells(i, 1).value
'                        T_Adm(j, 1) = oSheet.Cells(i, 2).value
'                        T_Adm(j, 2) = oSheet.Cells(i, 3).value
'                
'                        For Each oListO In Sheets("GEO").ListObjects
'                            If InStr(1, oListO.HeaderRowRange(2), Sheets(oSheet.Name).Cells(i, 1).value) > 0 Then
'                                oListO.HeaderRowRange(2) = Sheets(oSheet.Name).Cells(i, 1).value 'a revoir avec la Translation
'                
'                            End If
'                        Next
'                        i = i + 1
'                    Wend
'                    Sheets("GEO").ListObjects("T_geoTrad").DataBodyRange = T_Adm
'                End If
'            Next
'        
'            Sheets("MAIN").Range("RNG_GEO").value = .ActiveWorkbook.Name
'        
'            .ScreenUpdating = True
'            .Workbooks.Close
'            xlsApp.Quit
'            Set xlsApp = Nothing
'            
'            Sheets("main").Range("RNG_Msg").value = TranslateMsg("MSG_Fini")
'        
'        End With
'    Else
'        Sheets("main").Range("RNG_Msg").value = TranslateMsg("MSG_OpeAnnule")
'
'    End If
'
'    If Not Sheets("GEO").ListObjects("T_HistoGeo").DataBodyRange Is Nothing Then
'        Sheets("GEO").ListObjects("T_HistoGeo").DataBodyRange.Delete
'    End If
'    If Not Sheets("GEO").ListObjects("T_HistoFacil").DataBodyRange Is Nothing Then
'        Sheets("GEO").ListObjects("T_HistoFacil").DataBodyRange.Delete
'    End If
'
'    ReDim T_geo0(0)
'    ReDim T_aff0(0)
'    ReDim T_geo1(0)
'    ReDim T_aff1(0)
'    ReDim T_geo2(0)
'    ReDim T_aff2(0)
'    ReDim T_geo3(0)
'    ReDim T_aff3(0)
'    ReDim T_concat(0)
'    ReDim T_histo(0)
'    ReDim T_fac(0)
'    ReDim T_histoF(0)
'    ReDim T_concatF(0)
'
'End Sub
'
Sub GenerateData()

    Dim xlsApp As New Excel.Application
    Dim D_TitleDic As Scripting.Dictionary
    Dim T_dataDic
    Dim D_Choices As Scripting.Dictionary
    Dim T_Choices
    Dim T_Export
    
    Application.DisplayAlerts = False
    Sheets("Main").Range("a1").Select
    Call ShowHideCmdValidation(False)
    
    'On Error GoTo ErrLectureFichier
    '
    With ThisWorkbook.Sheets("Main").Range("RNG_Msg")
        xlsApp.Workbooks.Open [RNG_Dico].value
        xlsApp.ScreenUpdating = False
        xlsApp.Visible = False
        .value = TranslateMsg("MSG_LectDico")
        Set D_TitleDic = CreateDicoColVar(xlsApp, C_SheetNameDic, 2)
        T_dataDic = CreateTabDataVar(xlsApp, C_SheetNameDic, D_TitleDic, 3)
    
        .value = TranslateMsg("MSG_LectListe")
        Set D_Choices = CreateDicoColChoi(xlsApp, C_SheetNameChoices)
        T_Choices = CreateTabDataChoi(xlsApp, C_SheetNameChoices)
    
        .value = TranslateMsg("MSG_LectExport")
        T_Export = CreateParamExport(xlsApp)

        xlsApp.ActiveWorkbook.Close
        xlsApp.Quit
        Set xlsApp = Nothing
    
        'On Error GoTo errCreatLL
        '
        .value = TranslateMsg("MSG_CreationLL")
        Call BuildList(D_TitleDic, T_dataDic, D_Choices, T_Choices, T_Export)
        DoEvents
    
        .value = TranslateMsg("MSG_toutFbie")
        Application.DisplayAlerts = True
    
        Exit Sub
    End With
    
ErrLectureFichier:
    '[RNG_Msg].Value = "Une erreur s'est produite � la lecture du dico"
    'Exit Sub
    
errCreatLL:
    '[RNG_Msg].Value = "Une erreur s'est produite � la cr�ation de la LineList"
    'Exit Sub
    
End Sub

Sub CancelGenerate()

    Sheets("Main").Shapes("SHP_CtrlNouv").Visible = True

    Sheets("Main").Range("a1").Select
    Call ShowHideCmdValidation(False)

End Sub

Sub CtrlNew()

    Call ShowHideCmdValidation(False)
    If [RNG_Dico].value <> "" Then
        If Dir([RNG_Dico].value) <> "" Then
            If [RNG_Geo].value <> "" Then
                If Not IsWksOpened([RNG_Dico].value) Then
                    [RNG_Msg].value = TranslateMsg("MSG_ToutEstBon")
                    Call ShowHideCmdValidation(True) '
                    [RNG_Geo].Interior.Color = vbWhite
                    [RNG_Dico].Interior.Color = vbWhite
                Else
                    [RNG_Msg].value = TranslateMsg("MSG_FermerDico")
                End If
            Else
                [RNG_Msg].value = TranslateMsg("MSG_VeriFichGeo")
                [RNG_Geo].Interior.Color = LetColor("RedEpi")
            End If
        Else
            [RNG_Msg].value = TranslateMsg("MSG_VeriChemDico")
            [RNG_Dico].Interior.Color = LetColor("RedEpi")
    
        End If
    Else
        [RNG_Msg].value = TranslateMsg("MSG_VeriChemDico")
        [RNG_Dico].Interior.Color = LetColor("RedEpi")
    End If

End Sub

Private Sub ShowHideCmdValidation(EstVisible As Boolean)

    Sheets("Main").Shapes("SHP_Generer").Visible = EstVisible
    Sheets("Main").Shapes("SHP_Annuler").Visible = EstVisible
    Sheets("Main").Shapes("SHP_validation").Visible = EstVisible

End Sub

Private Function IsWksOpened(sNameClasseur As String) As Boolean
       
    Dim oWks As Object
    Dim i As Byte

    IsWksOpened = False
    i = 1
    While i <= Application.Workbooks.Count
        Set oWks = Application.Workbooks(i)
        If InStr(1, oWks.Name, Split(sNameClasseur, "\")(UBound(Split(sNameClasseur, "\")))) > 0 Then
            IsWksOpened = True
            Exit Function
        End If
        Set oWks = Nothing
        i = i + 1
    Wend

End Function

Public Sub ShowCmdValidation()

    Call ShowHideCmdValidation(False)

End Sub

Private Function CreateParamExport(xlsApp As Object)

    Dim i As Byte
    Dim j As Byte
    Dim T_temp

    With xlsApp.Sheets(C_SheetNameExport)
        i = 1
        j = 0
        ReDim T_temp(4, 0)
        While i <= .Cells(1, 1).End(xlDown).Row
            If LCase(.Cells(i, 4).value) = "active" Then
                ReDim Preserve T_temp(4, j)
                T_temp(0, j) = .Cells(i, 1).value
                T_temp(1, j) = .Cells(i, 2).value
                T_temp(2, j) = .Cells(i, 3).value
                T_temp(3, j) = .Cells(i, 4).value
                T_temp(4, j) = .Cells(i, 5).value
                j = j + 1
            End If
            i = i + 1
        Wend
    End With
    CreateParamExport = T_temp

End Function

