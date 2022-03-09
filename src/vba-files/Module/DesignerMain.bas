Attribute VB_Name = "DesignerMain"
Option Explicit

'Logic behind Loading the Dictionnary file
Sub DesLoadFileDic()

    Dim sFilePath As String                      'Path to the dictionnary
    
    'LoadFile is the procedure for loading the path to the dictionnary
    sFilePath = DesLoadFile("*.xlsb") 'lla
    
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
Sub DesLinelistDir()
    Dim sfolder As String
    sfolder = LoadFolder
    [RNG_LLDir] = ""
    If (sfolder <> "") Then
        [RNG_LLDir].value = sfolder
        [RNG_LLDir].Interior.Color = vbWhite
    Else
        [RNG_Edition].value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

' Adding a new load geo for the Geo file, in a new sheet called Geo2
' we have two functions for loading the geodatabase, but the second one
' will be in use instead of the first one
Sub DesLoadGeoFile()
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
    
    sFilePath = LoadFile("*.xlsx")
    
    If sFilePath <> "" Then
        With xlsapp
            .ScreenUpdating = False
            .Workbooks.Open sFilePath
            
            'Cleaning the previous Data in case the ranges are not Empty
            [RNG_Edition].value = TranslateMsg("MSG_NetoPrec")
            For i = 1 To AdmNames.Length
                'Adms
                If Not SheetGeo.ListObjects("T" & "_" & AdmNames.Items(i)).DataBodyRange Is Nothing Then
                    SheetGeo.ListObjects("T" & "_" & AdmNames.Items(i)).DataBodyRange.Delete
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
                With SheetGeo.ListObjects("T" & "_" & oSheet.Name)
                    outputAdress = Cells(T_Adm.LowerBound + 1, .Range.Column).Address
                    outputHeaderAdress = Cells(T_Adm.LowerBound, .Range.Column).Address

                    T_header.ToExcelRange Destination:=SheetGeo.Range(outputHeaderAdress), TransposeValues:=True
                    T_Adm.ToExcelRange Destination:=SheetGeo.Range(outputAdress)
                    
                    'resizing the Table
                    .Resize .Range.CurrentRegion
                End With
            Next
            
            SheetMain.Range("RNG_PathGeo").value = sFilePath
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
Sub DesGenerateData()
    Dim bGood As Boolean
    bGood = ControlForGenerate()
    
    If Not bGood Then
        ShowHideCmdValidation show:=False
        Exit Sub
    End If
    
    Dim xlsapp As Excel.Application 'Actual excel application
    Dim DictHeaders As BetterArray 'Dictionary headers
    Dim DictData As BetterArray 'Dictionary data
    Dim ChoicesHeaders As BetterArray 'Choices headers
    Dim ChoicesData As BetterArray 'Choices data
    Dim ExportData As BetterArray 'Export data
    Dim sPath As String
    
    Set xlsapp = New Excel.Application
    
    Application.DisplayAlerts = False
    'Application.ScreenUpdating = False
    
    'Be sure the actual Workbook is not opened
    If IsWkbOpened([RNG_LLName].value & ".xlsb") Then
        [RNG_Edition].value = TranslateMsg("MSG_CloseLL")
        [RNG_LLName].Interior.Color = LetColor("RedEpi")
        Exit Sub
    End If
        
    xlsapp.ScreenUpdating = False
    xlsapp.Visible = False
    xlsapp.Workbooks.Open SheetMain.Range(C_sRngLLName).value
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadDic")
    
    'create the Dictionnary data
    Set DictHeaders = GetHeaders(xlsapp, C_sParamSheetDict, eStartLinesDictHeaders)
    'create the data table of linelist patient using the dictionnary
    Set DictData = GetData(xlsapp, C_sSheetNameDic, eStartLinesDictData)
    
    'Create the choices data
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadList")
    'Create the dictionnary for the choices sheet
    Set ChoicesHeaders = GetHeaders(xlsapp, C_sParamSheetChoices, C_eStartLinesChoicesHeaders)
    'Create the table for the choices
    Set ChoicesData = GetData(xlsapp, C_sParamSheetChoices, C_eStartLinesChoicesData)
        
    'Reading the export sheet
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadExport")
    'create parameters for export
     ExportData = GetData(xlsapp, C_sParamSheetExport, C_eStartLinesExportData)

     xlsapp.ActiveWorkbook.Close
     xlsapp.Quit
     Set xlsapp = Nothing
        
     SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_BuildLL")
        
    'Creating the linelist using the dictionnary and choices data as well as export data
     'The BuildList procedure is in the linelist
     sPath = [RNG_LLDir].value & Application.PathSeparator & [RNG_LLName] & ".xlsb"
     Call BuildList(DictHeaders, DictData, ChoicesHeaders, ChoicesData, ExportData, sPath)
    DoEvents
    
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLCreated")
        [RNG_LLName].Interior.Color = vbWhite
        Sheets("Main").Shapes("SHP_OpenLL").Visible = msoTrue
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub CancelGenerate()
    Dim answer As Integer
    answer = MsgBox(TranslateMsg("MSG_ConfCancel"), vbYesNo)
    
    If answer = vbYes Then
        ShowHideCmdValidation show:=False
        Sheets("Main").Shapes("SHP_OpenLL").Visible = msoFalse
        End
    End If
    
    MsgBox TranslateMsg("MSG_Continue")
    
End Sub

Sub Control()
    Dim bGood As Boolean
    'Control to be sure we can generate a linelist
    bGood = ControlForGenerate()
    If Not bGood Then
        Exit Sub
    Else
        'Now that everything is fine, continue to generation process
        [RNG_Edition].value = TranslateMsg("MSG_Correct")
        ShowHideCmdValidation True
        [RNG_PathGeo].Interior.Color = vbWhite
        [RNG_PathDico].Interior.Color = vbWhite
        [RNG_LLName].Interior.Color = vbWhite
        [RNG_LLDir].Interior.Color = vbWhite
    End If
End Sub

'A control function to be sure that everything is fine for linelist Generation
Private Function DesControlForGenerate() As Boolean
   ControlForGenerate = False
    'Hide the shapes for linelist generation
    ShowHideCmdValidation False
    
    '****** dictionary
    
    'Be sure the dictionary path is not empty
    If [RNG_PathDico].value = "" Then
       [RNG_Edition].value = TranslateMsg("MSG_PathDic")
       [RNG_PathDico].Interior.Color = LetColor("RedEpi")
       Exit Function
    End If
    
    'Now check if the file exists
    If Dir([RNG_PathDico].value) = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_PathDic")
        [RNG_PathDico].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If
     
    'Be sure the dictionnary is not opened
    If IsWkbOpened(Dir([RNG_PathDico].value)) Then
        [RNG_Edition].value = TranslateMsg("MSG_CloseDic")
        [RNG_PathDico].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If
    
    [RNG_PathDico].Interior.Color = LetColor("White") 'if path is OK
    
    '****** geo
    
    'Be sure the geo path is not empty
    If [RNG_PathGeo].value = "" Then
       [RNG_Edition].value = TranslateMsg("MSG_PathDic")
       [RNG_PathGeo].Interior.Color = LetColor("RedEpi")
       Exit Function
    End If
    
    'Now check if the file exists
    If Dir([RNG_PathGeo].value) = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_PathDic")
        [RNG_PathGeo].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If
     
    'Be sure the geo is not opened
    If IsWkbOpened(Dir([RNG_PathGeo].value)) Then
        [RNG_Edition].value = TranslateMsg("MSG_CloseDic")
        [RNG_PathGeo].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If

    [RNG_PathGeo].Interior.Color = LetColor("White") 'if path is OK
    
    '****** linelist
    
    'Test the linelist directory is not empty
    If [RNG_LLDir].value = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_PathLL")
        [RNG_LLDir].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If

    'be sure the directory for the linelist exists
    If Dir([RNG_LLDir].value, vbDirectory) = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_PathLL")
        [RNG_LLDir].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If
    
    [RNG_LLDir].Interior.Color = LetColor("White") 'if path is OK

    '****** linelist name
    
    'be sure the linelist name is not empty
    If [RNG_LLName] = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_LLName")
        [RNG_LLName].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If
    
    'Be sure the linelist workbook is not already opened
    If IsWkbOpened([RNG_LLName].value & ".xlsb") Then
        [RNG_Edition].value = TranslateMsg("MSG_CloseLL")
        [RNG_LLName].Interior.Color = LetColor("RedEpi")
        Exit Function
    End If
    
    [RNG_LLName].Interior.Color = LetColor("White") 'if path is OK
    
    'Be sure the linelist does not exits
    'If Dir([RNG_LLDir].value & Application.PathSeparator & [RNG_LLName].value & ".xlsb") <> "" Then
    '    [RNG_Edition].value = TranslateMsg("MSG_exists")
     '   [RNG_LLName].Interior.Color = LetColor("RedEpi")
     '   Exit Function
    'End If
    
    ControlForGenerate = True

End Function

'Show or hide the generate the linelist shape
Private Sub ShowHideCmdValidation(show As Boolean)

    Sheets("Main").Shapes("SHP_Generer").Visible = show
    Sheets("Main").Shapes("SHP_Annuler").Visible = show
    Sheets("Main").Shapes("SHP_CtrlNouv").Visible = Not show

End Sub
Sub OpenLL()
    'Be sure that the directory and the linelist name are not empty
    If [RNG_LLDir].value = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_PathLL")
        [RNG_LLDir].Interior.Color = LetColor("RedEpi")
        Exit Sub
    End If
    
    If [RNG_LLName].value = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_LLName")
        [RNG_LLName].Interior.Color = LetColor("RedEpi")
        Exit Sub
    End If
    
    'Be sure the workbook is not already opened
    If IsWkbOpened([RNG_LLName].value & ".xlsb") Then
        [RNG_Edition].value = TranslateMsg("MSG_CloseLL")
        [RNG_LLName].Interior.Color = LetColor("RedEpi")
        Exit Sub
    End If
    
    'Be sure the workbook exits
    If Dir([RNG_LLDir].value & Application.PathSeparator & [RNG_LLName].value & ".xlsb") = "" Then
        [RNG_Edition].value = TranslateMsg("MSG_CheckLL")
        [RNG_LLName].Interior.Color = LetColor("RedEpi")
        [RNG_LLDir].Interior.Color = LetColor("RedEpi")
        ShowHideCmdValidation show:=False
        Sheets("Main").Shapes("SHP_OpenLL").Visible = msoFalse
        Exit Sub
    End If
    
    'Then open it
    Application.Workbooks.Open Filename:=[RNG_LLDir].value & Application.PathSeparator & [RNG_LLName].value & ".xlsb", ReadOnly:=False
End Sub

'Check if a workbook is Opened
Private Function IsWkbOpened(sName As String) As Boolean
    Dim oWkb As Workbook                         'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set oWkb = Application.Workbooks.Item(sName)
    IsWkbOpened = (Not oWkb Is Nothing)
    On Error GoTo 0
End Function


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

