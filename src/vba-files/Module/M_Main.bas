Attribute VB_Name = "M_Main"
Option Explicit
Option Base 1

Const C_SheetNameDic As String = "Dictionary"
Const C_SheetNameChoices As String = "Choices"
Const C_SheetNameExport As String = "Exports"

'Logic behind Loading the Dictionnary file
Sub LoadFileDic()

    Dim sFilePath As String                      'Path to the dictionnary
    
    'LoadFile is the procedure for loading the path to the dictionnary
    sFilePath = LoadFile
    
    'Update messages if the file path is correct
    If sFilePath <> "" Then
        [RNG_PathDico].value = sFilePath
        [RNG_Edition].value = TranslateMsg("MSG_ChemFich")
        [RNG_PathDico].Interior.Color = vbWhite
    Else
        [RNG_Edition].value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

'Logic behind specifying the linelist directory
Sub LinelistDir()
    Dim sfolder As String
    sfolder = LoadFolder
    [RNG_LLDir] = ""
    If (sfolder <> "") Then
        [RNG_LLDir].value = sfolder
    Else
        [RNG_Edition].value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

' Adding a new load geo for the Geo file, in a new sheet called Geo2
' we have two functions for loading the geodatabase, but the second one
' will be in use instead of the first one
Sub LoadGeoFile()
    Dim geoSheet As String
    geoSheet = "GEO"                             'geoSheet in the designer
    Dim sFilePath As String                      'File path to the geo file
    Dim xlsapp As Excel.Application
    Dim oSheet As Object
    Dim T_Adm  As BetterArray                    'Table for admin levels
    Dim T_header As BetterArray                  'Table for the headers of the listobjects
    Dim outputAdress As String
    Dim outputHeaderAdress As String
    Dim AdmNames As BetterArray                  'Array of the sheetnames
    Dim i As Integer                             'iterator
    
    'Sheet names
    Set AdmNames = New BetterArray
    AdmNames.LowerBound = 1
    AdmNames.Push "ADM1", "ADM2", "ADM3", "ADM4", "HF", "NAMES" 'Names of each sheet
    
    'Defining the adm and headers array
    Set T_Adm = New BetterArray
    Set T_header = New BetterArray
    Set xlsapp = New Excel.Application
    
    sFilePath = LoadFile
    
    If sFilePath <> "" Then
        With xlsapp
            .ScreenUpdating = False
            .Workbooks.Open sFilePath
            
            'Cleaning the previous Data in case the ranges are not Empty
            [RNG_Edition].value = TranslateMsg("MSG_NetoPrec")
            For i = 1 To AdmNames.Length
                'Adms
                If Not Sheets(geoSheet).ListObjects("T" & "_" & AdmNames.Items(i)).DataBodyRange Is Nothing Then
                    Sheets(geoSheet).ListObjects("T" & "_" & AdmNames.Items(i)).DataBodyRange.Delete
                End If
            Next

            'Reloading the data from the Geobase
            For Each oSheet In xlsapp.Worksheets
                [RNG_Edition].value = TranslateMsg("MSG_EnCours") & oSheet.Name
                T_Adm.Clear
                T_header.Clear
                'loading the data in memory
                T_Adm.FromExcelRange oSheet.Range("A2"), True, True
                'The headers
                T_header.FromExcelRange oSheet.Range("A1"), False, True
                
                'Be sure my sheetnames are correct
                If Not AdmNames.Includes(oSheet.Name) Then
                    [RNG_Edition].value = TranslateMsg("MSG_Error_Sheet") & oSheet.Name
                    Exit Sub
                End If
                
                'Check if the sheet is the admin exists sheet before writing in the adm table
                With Sheets(geoSheet).ListObjects("T" & "_" & oSheet.Name)
                    outputAdress = Cells(T_Adm.LowerBound + 1, .Range.Column).Address
                    outputHeaderAdress = Cells(T_Adm.LowerBound, .Range.Column).Address

                    T_header.ToExcelRange Destination:=Sheets(geoSheet).Range(outputHeaderAdress), TransposeValues:=True
                    T_Adm.ToExcelRange Destination:=Sheets(geoSheet).Range(outputAdress)
                    
                    'resizing the Table
                    .Resize .Range.CurrentRegion
                End With
            Next
            
            Sheets("MAIN").Range("RNG_PathGeo").value = sFilePath
            .ScreenUpdating = True
            .Workbooks.Close
            xlsapp.Quit
            Set xlsapp = Nothing
            Set T_Adm = Nothing
            Set T_header = Nothing
            Set AdmNames = Nothing
            
            [RNG_Edition].value = TranslateMsg("MSG_Fini")
        End With
        
        'Remove the historic of the Geo and the facility if not empty
        If Not Sheets(geoSheet).ListObjects("T_HistoGeo").DataBodyRange Is Nothing Then
            Sheets(geoSheet).ListObjects("T_HistoGeo").DataBodyRange.Delete
        End If
        
        If Not Sheets(geoSheet).ListObjects("T_HistoHF").DataBodyRange Is Nothing Then
            Sheets(geoSheet).ListObjects("T_HistoHF").DataBodyRange.Delete
        End If
    Else
        [RNG_Edition].value = TranslateMsg("MSG_OpeAnnule")
    End If

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This is the Sub for generating the data of the linelist using the input in the designer
' The main entry point is the BuildList function which creates the Linelist-patient sheet as
' well as all the forms in the linelist
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub GenerateData()

    Dim xlsapp As Excel.Application
    Dim D_TitleDic As Scripting.Dictionary
    Dim T_dataDic
    Dim D_Choices As Scripting.Dictionary
    Dim T_Choices
    Dim T_Export
    
    Set xlsapp = New Excel.Application
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'return to the create button
    Call ShowHideCmdValidation(False)
    
    'On Error GoTo ErrLectureFichier
    With ThisWorkbook.Sheets("Main").Range("RNG_Msg")
        xlsapp.ScreenUpdating = False
        xlsapp.Visible = False
        xlsapp.Workbooks.Open [RNG_Dico].value
        
        .value = TranslateMsg("MSG_LectDico")
        'create the Dictionnary of the linelist patient sheet
        Set D_TitleDic = CreateDicoColVar(xlsapp, C_SheetNameDic, 2)
        'create the data table of linelist patient using the dictionnary
        T_dataDic = CreateTabDataVar(xlsapp, C_SheetNameDic, D_TitleDic, 3)
    
        .value = TranslateMsg("MSG_LectListe")
        'Create the dictionnary for the choices sheet
        Set D_Choices = CreateDicoColChoi(xlsapp, C_SheetNameChoices)
        'Create the table for the choices
        T_Choices = CreateTabDataChoi(xlsapp, C_SheetNameChoices)
    
        .value = TranslateMsg("MSG_LectExport")
        'create parameters for export
        T_Export = CreateParamExport(xlsapp)

        xlsapp.ActiveWorkbook.Close
        xlsapp.Quit
        Set xlsapp = Nothing
        
        .value = TranslateMsg("MSG_CreationLL")
        
        'Creating the linelist using the dictionnary and choices data as well as export data
        ' The BuildList procedure is in the
        Call BuildList(D_TitleDic, T_dataDic, D_Choices, T_Choices, T_Export)
        DoEvents
    
        .value = TranslateMsg("MSG_toutFbie")
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub CancelGenerate()
    ShowHideCmdValidation show:=False
End Sub

'A control Sub to be sure that everything is fine for linelist Generation
Sub ControlForGenerate()
   
    'Hide the shapes for linelist generation
    ShowHideCmdValidation False
    
    'Be sure the dictionary path is not empty
    If [RNG_PathDico].value = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_PathDic")
        [RNG_PathDico].Interior.Color = LetColor("RedEpi")
        Exit Sub
    End If
    
    'Now check if the file exists
    If Dir([RNG_PathDico].value) = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_PathDic")
        [RNG_PathDico].Interior.Color = LetColor("RedEpi")
        Exit Sub
    End If
     
    'Be sure the dictionnary is not opened
    If IsWkbOpened([RNG_PathDico].value) Then
        [RNG_Edition].value = TranslateMsg("MSG_CloseDic")
        Exit Sub
    End If
    
    'Now that everything is fine, continue to generation process
    [RNG_Edition].value = TranslateMsg("MSG_Correct")
    ShowHideCmdValidation True
    [RNG_PathGeo].Interior.Color = vbWhite
    [RNG_PathDico].Interior.Color = vbWhite

End Sub

'Show or hide the generate the linelist shape
Private Sub ShowHideCmdValidation(show As Boolean)

    Sheets("Main").Shapes("SHP_Generer").Visible = show
    Sheets("Main").Shapes("SHP_Annuler").Visible = show
    Sheets("Main").Shapes("SHP_CtrlNouv").Visible = Not show

End Sub

'Check if a workbook is Opened
Private Function IsWkbOpened(sName As String) As Boolean
    Dim oWkb As Workbook                         'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set oWkb = Application.Workbooks.Item(sName)
    IsWkbOpened = (Not oWkb Is Nothing)
    On Error GoTo 0
End Function

Public Sub ShowCmdValidation()

    ShowHideCmdValidation show:=False

End Sub

Private Function CreateParamExport(xlsapp As Object)

    Dim i As Byte
    Dim j As Byte
    Dim T_temp

    With xlsapp.Sheets(C_SheetNameExport)
        i = 1
        j = 1
        ReDim T_temp(5, 1)
        While i <= .Cells(1, 1).End(xlDown).Row
            If LCase(.Cells(i, 4).value) = "active" Then
                ReDim Preserve T_temp(5, j)
                T_temp(1, j) = .Cells(i, 1).value
                T_temp(2, j) = .Cells(i, 2).value
                T_temp(3, j) = .Cells(i, 3).value
                T_temp(4, j) = .Cells(i, 4).value
                T_temp(5, j) = .Cells(i, 5).value
                j = j + 1
            End If
            i = i + 1
        Wend
    End With
    CreateParamExport = T_temp

End Function

