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
    Dim outputAdress As String
    Dim outputHeaderAdress As String
    Dim objectName As String 'name of the listobject in the geo sheet
    
    'Defining the adm and headers array
    Set T_Adm = New BetterArray
    Set T_header = New BetterArray
    Set xlsApp = New Excel.Application
    
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
                T_header.Clear
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
                    outputAdress = Cells(T_Adm.LowerBound + 1, .Range.Column).Address
                    outputHeaderAdress = Cells(T_Adm.LowerBound, .Range.Column).Address
               End With

                T_header.ToExcelRange Destination:=Sheets(geoSheet).Range(outputHeaderAdress), TransposeValues:=True
                T_Adm.ToExcelRange Destination:=Sheets(geoSheet).Range(outputAdress)
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

