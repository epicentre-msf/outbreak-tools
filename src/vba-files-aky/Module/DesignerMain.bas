Attribute VB_Name = "DesignerMain"
Option Explicit

Dim bGeoLoaded As Boolean
'Logic for loading files and folders and checking if workbooks are opened ===============================================================

'Logic behind Loading the Dictionnary file
Sub DesLoadFileDic()

    Dim sFilePath As String                      'Path to the dictionnary
    
    'LoadFile is the procedure for loading the path to the dictionnary
    sFilePath = DesLoadFile("*.xlsb") 'The
    
    'Update messages if the file path is correct
    If sFilePath <> "" Then
        SheetMain.Range(C_sRngPathDic).value = sFilePath
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ChemFich")
        SheetMain.Range(C_sRngPathDic).Interior.Color = vbWhite
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

'Logic behind specifying the linelist directory
Sub DesLinelistDir()
    Dim sfolder As String
    sfolder = LoadFolder
    SheetMain.Range(C_sRngLLDir) = ""
    If (sfolder <> "") Then
        SheetMain.Range(C_sRngLLDir).value = sfolder
        SheetMain.Range(C_sRngLLDir).Interior.Color = vbWhite
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

'Check if a workbook is Opened
Private Function IsWkbOpened(sName As String) As Boolean
    Dim oWkb As Workbook                         'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set oWkb = Application.Workbooks.Item(sName)
    IsWkbOpened = (Not oWkb Is Nothing)
    On Error GoTo 0
End Function


' Adding a new load geo for the Geo file, in a new sheet called Geo2
' we have two functions for loading the geodatabase, but the second one
' will be in use instead of the first one
Sub DesLoadGeoFile()
    Application.ScreenUpdating = False
    bGeoLoaded = False
    
    Dim sFilePath As String                      'File path to the geo file
    Dim oSheet As Object
    Dim AdmData  As BetterArray                  'Table for admin levels
    Dim AdmHeader As BetterArray                 'Table for the headers of the listobjects
    Dim AdmNames As BetterArray                  'Array of the sheetnames
    Dim i As Integer                             'iterator
    Dim Wkb As Workbook
    'Sheet names
    Set AdmNames = New BetterArray
    AdmNames.LowerBound = 1
    AdmNames.Push "ADM1", "ADM2", "ADM3", "ADM4", "HF", "NAMES" 'Names of each sheet
    
    'Defining the adm and headers array
    Set AdmData = New BetterArray
    AdmData.LowerBound = 1
    Set AdmHeader = New BetterArray
    AdmHeader.LowerBound = 1
    'Set xlsapp = New Excel.Application
    sFilePath = DesLoadFile("*.xlsx")
    
    If sFilePath <> "" Then
        'Open the geo workbook and hide the windows
        Set Wkb = Workbooks.Open(sFilePath)
        'Windows(Wkb.Name).Visible = False
            
        'Cleaning the previous Data in case the ranges are not Empty
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_NetoPrec")
        For i = 1 To AdmNames.Length
            'Adms (Maybe come back to work on the names?)
            If Not SheetGeo.ListObjects("T" & "_" & AdmNames.Items(i)).DataBodyRange Is Nothing Then
                SheetGeo.ListObjects("T" & "_" & AdmNames.Items(i)).DataBodyRange.Delete
            End If
        Next
            
        'Reloading the data from the Geobase
        For Each oSheet In Wkb.Worksheets
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_EnCours") & oSheet.Name
            AdmData.Clear
            AdmHeader.Clear
            'loading the data in memory
            AdmData.FromExcelRange oSheet.Range("A2"), DetectLastRow:=True, DetectLastColumn:=True
            'The headers
            AdmHeader.FromExcelRange oSheet.Range("A1"), DetectLastRow:=False, DetectLastColumn:=True
                
            'Be sure my sheetnames are correct
            If Not AdmNames.Includes(oSheet.Name) Then
                SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Error_Sheet") & oSheet.Name
                
                Application.ScreenUpdating = True
                Exit Sub
            End If
                
            'Check if the sheet is the admin exists sheet before writing in the adm table
            With SheetGeo.ListObjects("T" & "_" & oSheet.Name)
                AdmHeader.ToExcelRange Destination:=SheetGeo.Cells(1, .Range.Column), TransposeValues:=True
                AdmData.ToExcelRange Destination:=SheetGeo.Cells(2, .Range.Column)
                'resizing the Table
                .Resize .Range.CurrentRegion
            End With
        Next
            
        SheetMain.Range(C_sRngPathGeo).value = sFilePath
        Wkb.Close SaveChanges:=False
       
        Set AdmData = Nothing
        Set AdmHeader = Nothing
        Set AdmNames = Nothing
        Set Wkb = Nothing
            
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Fini")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = DesLetColor("White")
        bGeoLoaded = True
        
        'Remove the historic of the Geo and the facility if not empty
        If Not SheetGeo.ListObjects(C_sTabHistoGeo).DataBodyRange Is Nothing Then
            SheetGeo.ListObjects(C_sTabHistoGeo).DataBodyRange.Delete
        End If
        
        If Not SheetGeo.ListObjects(C_sTabHistoHF).DataBodyRange Is Nothing Then
            SheetGeo.ListObjects(C_sTabHistoHF).DataBodyRange.Delete
        End If
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If
    
    Application.ScreenUpdating = True
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This is the Sub for generating the data of the linelist using the input in the designer
' The main entry point is the BuildList function which creates the Linelist-patient sheet as
' well as all the forms in the linelist
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub DesGenerateData()
    Dim bGood As Boolean
    bGood = DesControlForGenerate()
   
    
    If Not bGood Then
        DesShowHideCmdValidation show:=False
        Exit Sub
    End If
    
    Dim DictHeaders As BetterArray 'Dictionary headers
    Dim DictData As BetterArray 'Dictionary data
    Dim ChoicesHeaders As BetterArray 'Choices headers
    Dim ChoicesData As BetterArray 'Choices data
    Dim ExportData As BetterArray 'Export data
    Dim sPath As String
    
    Set xlsapp = New Excel.Application
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Be sure the actual Workbook is not opened
    If IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = DesLetColor("RedEpi")
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
    
    'Create parameters for export
    ExportData = GetData(xlsapp, C_sParamSheetExport, C_eStartLinesExportData)

    xlsapp.ActiveWorkbook.Close
    xlsapp.Quit
    Set xlsapp = Nothing
        
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_BuildLL")
        
    'Creating the linelist using the dictionnary and choices data as well as export data
    'The BuildList procedure is in the linelist
    sPath = SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName) & ".xlsb"
    
    Call BuildList(DictHeaders, DictData, ChoicesHeaders, ChoicesData, ExportData, sPath)
    DoEvents
    
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLCreated")
    SheetMain.Range(C_sRngLLName).Interior.Color = vbWhite
    SheetMain.Shapes("SHP_OpenLL ").Visible = msoTrue
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub DesCancelGenerate()
    Dim answer As Integer
    answer = MsgBox(TranslateMsg("MSG_ConfCancel"), vbYesNo)
    
    If answer = vbYes Then
        DesShowHideCmdValidation show:=False
        SheetMain.Shapes("SHP_OpenLL ").Visible = msoFalse
        End
    End If
    
    MsgBox TranslateMsg("MSG_Continue")
    
End Sub

Sub SetInputRangesToWhite()
    
    SheetMain.Range(C_sRngPathGeo).Interior.Color = vbWhite
    SheetMain.Range(C_sRngPathDic).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLName).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLDir).Interior.Color = vbWhite
    SheetMain.Range(C_sRngEdition).Interior.Color = vbWhite

End Sub

Sub DesControl()
    'Put every range in white before the control
    
    Call SetInputRangesToWhite
    
    Dim bGood As Boolean
    'Control to be sure we can generate a linelist
    bGood = DesControlForGenerate()
    If Not bGood Then
        Exit Sub
    Else
        'Now that everything is fine, continue to generation process but issue a warning in case it will
        'replace the previous existing file
        Call SetInputRangesToWhite
        
        If Dir(SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb") <> "" Then
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Correct") & ": " & SheetMain.Range(C_sRngLLName).value & ".xlsb " & TranslateMsg("MSG_Exists")
            SheetMain.Range(C_sRngEdition).Interior.Color = DesLetColor("Grey")
        Else
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Correct")
        End If
        
        DesShowHideCmdValidation True
        
    End If
End Sub

'A control function to be sure that everything is fine for linelist Generation
Private Function DesControlForGenerate() As Boolean
    DesControlForGenerate = False
    'Hide the shapes for linelist generation
    DesShowHideCmdValidation False
    
    'Checking coherence of the Dictionnary ----------------------------
    
    'Be sure the dictionary path is not empty
    If SheetMain.Range(C_sRngPathDic).value = "" Then
       SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathDic")
       SheetMain.Range(C_sRngPathDic).Interior.Color = DesLetColor("RedEpi")
       Exit Function
    End If
    
    'Now check if the file exists
    If Dir(SheetMain.Range(C_sRngPathDic).value) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathDic")
        SheetMain.Range(C_sRngPathDic).Interior.Color = DesLetColor("RedEpi")
        Exit Function
    End If
     
    'Be sure the dictionnary is not opened
    If IsWkbOpened(Dir(SheetMain.Range(C_sRngPathDic).value)) Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseDic")
        SheetMain.Range(C_sRngPathDic).Interior.Color = DesLetColor("RedEpi")
        Exit Function
    End If
    
    SheetMain.Range(C_sRngPathDic).Interior.Color = DesLetColor("White") 'if path is OK
    
    'Checking coherence of the GEO (maybe remove?) ------------------------------------------
    
    'Be sure the geo path is not empty
    If SheetMain.Range(C_sRngPathGeo).value = "" Then
       SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathDic")
       SheetMain.Range(C_sRngPathGeo).Interior.Color = DesLetColor("RedEpi")
       Exit Function
    End If
    
    'Now check if the file exists
    If Dir(SheetMain.Range(C_sRngPathGeo).value) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathGeo")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = DesLetColor("RedEpi")
        Exit Function
    End If
     
    'Be sure the geo has been loaded correctly
    If Not bGeoLoaded Then 'bGeoLoaded is a global variable resticted to this module only
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LoadGeo")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = DesLetColor("RedEpi")
    End If

    SheetMain.Range(C_sRngPathGeo).Interior.Color = DesLetColor("White") 'if path is OK
    
    'Checking coherence of the Linelist File --------------------------------------------------
    
    'Be sure the linelist directory is not empty
    If SheetMain.Range(C_sRngLLDir).value = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = DesLetColor("RedEpi")
        Exit Function
    End If

    'Be sure the directory for the linelist exists
    If Dir(SheetMain.Range(C_sRngLLDir).value, vbDirectory) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = DesLetColor("RedEpi")
        Exit Function
    End If
    
    SheetMain.Range(C_sRngLLDir).Interior.Color = DesLetColor("White") 'if path is OK

    'Checking coherence of the linelist name ----------------------------------------------
    
    'be sure the linelist name is not empty
    If SheetMain.Range(C_sRngLLName) = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLName")
        SheetMain.Range(C_sRngLLName).Interior.Color = DesLetColor("RedEpi")
        Exit Function
    End If
    
    'Be sure the linelist workbook is not already opened
    If IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = DesLetColor("RedEpi")
        Exit Function
    End If
    
    SheetMain.Range(C_sRngLLName).Interior.Color = DesLetColor("White") 'if path is OK
    
    DesControlForGenerate = True

End Function

'Show or hide the generate the linelist shape
Private Sub DesShowHideCmdValidation(show As Boolean)

    SheetMain.Shapes("SHP_Generer").Visible = show
    SheetMain.Shapes("SHP_Annuler").Visible = show
    SheetMain.Shapes("SHP_CtrlNouv").Visible = Not show

End Sub
Sub DesOpenLL()
    'Be sure that the directory and the linelist name are not empty
    If SheetMain.Range(C_sRngLLDir).value = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = DesLetColor("RedEpi")
        Exit Sub
    End If
    
    If SheetMain.Range(C_sRngLLName).value = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLName")
        SheetMain.Range(C_sRngLLName).Interior.Color = DesLetColor("RedEpi")
        Exit Sub
    End If
    
    'Be sure the workbook is not already opened
    If IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = DesLetColor("RedEpi")
        Exit Sub
    End If
    
    'Be sure the workbook exits
    If Dir(SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb") = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CheckLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = DesLetColor("RedEpi")
        SheetMain.Range(C_sRngLLDir).Interior.Color = DesLetColor("RedEpi")
        DesShowHideCmdValidation show:=False
        Sheets("Main").Shapes("SHP_OpenLL ").Visible = msoFalse
        Exit Sub
    End If
    
    'Then open it
    Application.Workbooks.Open Filename:=SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb", ReadOnly:=False
End Sub



