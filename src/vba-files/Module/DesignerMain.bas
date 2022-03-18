Attribute VB_Name = "DesignerMain"
Option Explicit
Dim bGeoLoaded As Boolean 'This will check if the Geodata is loaded or Not: The user have to load a Geobase
 
'LOADING FILES AND FOLDERS ==============================================================================================================================================

'Loading the Dictionnary file ----------------------------------------------------------------------------------------------------------------------------------------------------
Sub LoadFileDic()

    Dim sFilePath As String                      'Path to the dictionnary
    
    'LoadFile is the procedure for loading the path to the dictionnary
    sFilePath = Helpers.LoadFile("*.xlsb") 'The
    
    'Update messages if the file path is correct
    If sFilePath <> "" Then
        SheetMain.Range(C_sRngPathDic).value = sFilePath
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ChemFich")
        SheetMain.Range(C_sRngPathDic).Interior.Color = vbWhite
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

'Loading the Lineist directory ---------------------------------------------------------------------------------------------------------------------------------------------------
Sub LinelistDir()
    Dim sfolder As String

    sfolder = Helpers.LoadFolder
    SheetMain.Range(C_sRngLLDir) = ""
    If (sfolder <> "") Then
        SheetMain.Range(C_sRngLLDir).value = sfolder
        SheetMain.Range(C_sRngLLDir).Interior.Color = vbWhite
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

'Loading the Geobase --------------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Adding a new load geo for the Geo file, in a new sheet called Geo2
' we have two functions for loading the geodatabase, but the second one
' will be in use instead of the first one.
'
Sub LoadGeoFile()
    
    Call Helpers.BeginWork(Application)

    bGeoLoaded = False
    
    Dim sFilePath   As String                      'File path to the geo file
    Dim oSheet      As Object
    Dim AdmData     As BetterArray                  'Table for admin levels
    Dim AdmHeader   As BetterArray                 'Table for the headers of the listobjects
    Dim AdmNames    As BetterArray                  'Array of the sheetnames
    Dim i           As Integer                             'iterator
    Dim Wkb         As Workbook
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
    sFilePath = Helpers.LoadFile("*.xlsx")
    
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
                
                Call EndWork(Application)
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
        Wkb.Close savechanges:=False
       
        Set AdmData = Nothing
        Set AdmHeader = Nothing
        Set AdmNames = Nothing
        Set Wkb = Nothing
            
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Fini")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("White")
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
    
    Call EndWork(Application)
End Sub

'GENERATE THE LINELIST DATA =========================================================================================

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This is the Sub for generating the data of the linelist using the input in the designer
' The main entry point is the BuildList function which creates the Linelist-patient sheet as
' well as all the forms in the linelist
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub GenerateData()
    Dim bGood As Boolean
    bGood = DesignerMainHelpers.ControlForGenerate(bGeoLoaded)

   BeginWork xlsapp:=Application
   
    If Not bGood Then
        DesignerMainHelpers.ShowHideCmdValidation show:=False
        Exit Sub
    End If
    
    Dim DictHeaders     As BetterArray          'Dictionary headers
    Dim DictData        As BetterArray          'Dictionary data
    Dim ChoicesHeaders  As BetterArray          'Choices headers
    Dim ChoicesData     As BetterArray          'Choices data
    Dim ExportData      As BetterArray          'Export data
    Dim sPath           As String
    Dim Wkb             As Workbook
    
    'Be sure the actual Workbook is not opened
    
    If IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If
                                          
    Set Wkb = Workbooks.Open(SheetMain.Range(C_sRngPathDic).value)
    
    
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadDic")
    
    'Create the Dictionnary data
    Set DictHeaders = Helpers.GetHeaders(Wkb, C_sParamSheetDict, C_eStartLinesDictHeaders)
    
    'Create the data table of linelist patient using the dictionnary
    Set DictData = Helpers.GetData(Wkb, C_sParamSheetDict, C_eStartLinesDictData)
    
    'Create the choices data
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadList")
    
    'Create the dictionnary for the choices sheet
    Set ChoicesHeaders = Helpers.GetHeaders(Wkb, C_sParamSheetChoices, C_eStartLinesChoicesHeaders)
    
    'Create the table for the choices
    Set ChoicesData = Helpers.GetData(Wkb, C_sParamSheetChoices, C_eStartLinesChoicesData)
       
    'Reading the export sheet
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadExport")
    
    'Create parameters for export
    Set ExportData = Helpers.GetData(Wkb, C_sParamSheetExport, C_eStartLinesExportData)

    Wkb.Close savechanges:=False
    
    Set Wkb = Nothing
    
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_BuildLL")
        
    'Creating the linelist using the dictionnary and choices data as well as export data
    'The BuildList procedure is in the linelist
    sPath = SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName) & ".xlsb"

    
    Call DesignerBuildList.BuildList(DictHeaders, DictData, ChoicesHeaders, ChoicesData, ExportData, sPath)
    
    DoEvents

    EndWork xlsapp:=Application

    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLCreated")
    
    Call SetInputRangesToWhite
    
    SheetMain.Shapes("SHP_OpenLL").Visible = msoTrue

    
End Sub


'Adding some controls before generating the linelist  =============================================================================================================================
Public Sub Control()
    'Put every range in white before the control
    
    Call SetInputRangesToWhite
    
    Dim bGood As Boolean
    'Control to be sure we can generate a linelist
    bGood = DesignerMainHelpers.ControlForGenerate(bGeoLoaded)
    If Not bGood Then
        Exit Sub
    Else
        'Now that everything is fine, continue to generation process but issue a warning in case it will
        'replace the previous existing file
        Call SetInputRangesToWhite
        
        If Dir(SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb") <> "" Then
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Correct") & ": " & SheetMain.Range(C_sRngLLName).value & ".xlsb " & TranslateMsg("MSG_Exists")
            SheetMain.Range(C_sRngEdition).Interior.Color = Helpers.GetColor("Grey")
        Else
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Correct")
        End If
        
        DesignerMainHelpers.ShowHideCmdValidation True
        
    End If
End Sub


'OPEN THE GENERATED LINELIST ==========================================================================================

Sub OpenLL()
    'Be sure that the directory and the linelist name are not empty
    If SheetMain.Range(C_sRngLLDir).value = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_PathLL")
        SheetMain.Range(C_sRngLLDir).Interior.Color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If
    
    If SheetMain.Range(C_sRngLLName).value = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLName")
        SheetMain.Range(C_sRngLLName).Interior.Color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If
    
    'Be sure the workbook is not already opened
    If IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If
    
    'Be sure the workbook exits
    If Dir(SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb") = "" Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CheckLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = Helpers.GetColor("RedEpi")
        SheetMain.Range(C_sRngLLDir).Interior.Color = Helpers.GetColor("RedEpi")
        ShowHideCmdValidation show:=False
        Sheets("Main").Shapes("SHP_OpenLL ").Visible = msoFalse
        Exit Sub
    End If
    
    'Then open it
    Application.Workbooks.Open Filename:=SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb", ReadOnly:=False
End Sub







