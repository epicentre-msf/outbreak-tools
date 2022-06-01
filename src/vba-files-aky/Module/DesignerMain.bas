Attribute VB_Name = "DesignerMain"
Option Explicit

Public iUpdateCpt As Integer

'LOADING FILES AND FOLDERS ==============================================================================================================================================

'Loading the Dictionnary file ----------------------------------------------------------------------------------------------------------------------------------------------------
Sub LoadFileDic()

    Dim sFilePath As String                      'Path to the dictionnary

    BeginWork xlsapp:=Application

    'LoadFile is the procedure for loading the path to the dictionnary
    sFilePath = Helpers.LoadFile("*.xlsb") 'The

    'Update messages if the file path is correct
    If sFilePath <> "" Then
        SheetMain.Range(C_sRngPathDic).value = sFilePath
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ChemFich")
        SheetMain.Range(C_sRngPathDic).Interior.Color = vbWhite

        'Import the languages after loading the setup file
        Call ImportLangAnalysis(sFilePath)
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If

    EndWork xlsapp:=Application
End Sub

'Loading a linelist file ----------------------------------------------------------------------------------------------------------------------------------------------------
Sub LoadFileLL()

    Dim sFilePath As String                      'Path to the linelist

    'LoadFile is the procedure for loading the path to the linelist
    sFilePath = Helpers.LoadFile("*.xlsb") 'The

    If sFilePath = "" Then Exit Sub

    On Error GoTo ErrorManage
    Application.Workbooks.Open Filename:=sFilePath, ReadOnly:=False

    Exit Sub
ErrorManage:
    MsgBox TranslateMsg("MSG_TitlePassWord"), vbCritical, TranslateMsg("MSG_PassWord")
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

    BeginWork xlsapp:=Application

    Dim sFilePath   As String                      'File path to the geo file
    Dim oSheet      As Object
    Dim AdmData     As BetterArray                  'Table for admin levels
    Dim AdmHeader   As BetterArray                 'Table for the headers of the listobjects
    Dim AdmNames    As BetterArray                  'Array of the sheetnames
    Dim i           As Integer                             'iterator
    Dim Wkb         As Workbook
    'Sheet names
    Set AdmNames = New BetterArray
    Set AdmData = New BetterArray
    Set AdmHeader = New BetterArray

    AdmNames.LowerBound = 1
    AdmNames.Push C_sAdm1, C_sAdm2, C_sAdm3, C_sAdm4, C_sHF, C_sNames, C_sHistoHF, C_sHistoGeo, C_sGeoMetadata

    sFilePath = Helpers.LoadFile("*.xlsx")

    If sFilePath <> "" Then
        'Open the geo workbook and hide the windows
        Set Wkb = Workbooks.Open(sFilePath)
        'Write the filename of the geobase somewhere for the export
        SheetGeo.Range(C_sRngGeoName).value = Dir(sFilePath)

        'Cleaning the previous Data in case the ranges are not Empty
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_NetoPrec")
        For i = 1 To AdmNames.Length
            'Adms (Maybe come back to work on the names?)
            If Not SheetGeo.ListObjects("T_" & AdmNames.Items(i)).DataBodyRange Is Nothing Then
                SheetGeo.ListObjects("T_" & AdmNames.Items(i)).DataBodyRange.Delete
            End If
        Next

        'Reloading the data from the Geobase
        For Each oSheet In Wkb.Worksheets
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_EnCours") & oSheet.Name
            AdmData.Clear
            AdmHeader.Clear

            'Be sure my sheetnames are correct before loading the data
            If AdmNames.Includes(oSheet.Name) Then

                'loading the data in memory
                AdmData.FromExcelRange oSheet.Range("A2"), DetectLastRow:=True, DetectLastColumn:=True
                'The headers
                AdmHeader.FromExcelRange oSheet.Range("A1"), DetectLastRow:=False, DetectLastColumn:=True

                'Check if the sheet is the admin exists sheet before writing in the adm table
                With SheetGeo.ListObjects("T_" & oSheet.Name)
                    AdmHeader.ToExcelRange Destination:=SheetGeo.Cells(1, .Range.Column), TransposeValues:=True
                    AdmData.ToExcelRange Destination:=SheetGeo.Cells(2, .Range.Column)

                    'Resizing the Table
                    .Resize .Range.CurrentRegion
                End With

            End If


        Next

        SheetMain.Range(C_sRngPathGeo).value = sFilePath
        Wkb.Close savechanges:=False

        Set AdmData = Nothing
        Set AdmHeader = Nothing
        Set AdmNames = Nothing
        Set Wkb = Nothing

        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Fini")
        SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("White")

    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If

    Call TranslateHeadGeo

    EndWork xlsapp:=Application
End Sub

'GENERATE THE LINELIST DATA =================================================================================================

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This is the Sub for generating the data of the linelist using the input in the designer
' The main entry point is the BuildList function which creates the Linelist-patient sheet as
' well as all the forms in the linelist
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub GenerateData()
    Dim bGood As Boolean
    bGood = DesignerMainHelpers.ControlForGenerate()

    If Not bGood Then
        Exit Sub
    End If

    Dim DictHeaders     As BetterArray          'Dictionary headers
    Dim DictData        As BetterArray          'Dictionary data
    Dim ChoicesHeaders  As BetterArray          'Choices headers
    Dim ChoicesData     As BetterArray          'Choices data
    Dim ExportData      As BetterArray          'Export data
    Dim TransData       As BetterArray          'Translation data
    Dim GlobalSumData   As BetterArray
    Dim sPath           As String
    Dim SetupWkb        As Workbook
    Dim DesWkb          As Workbook
    Dim iOpenLL         As Integer
    Dim i               As Integer

    iUpdateCpt = 0
    
    BeginWork xlsapp:=Application
    SheetMain.Range(C_sRngUpdate).value = vbNullString

    'Be sure the actual Workbook is not opened

    If IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If
    
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_MovingData")

    Set DesWkb = DesignerWorkbook
    Set SetupWkb = Workbooks.Open(SheetMain.Range(C_sRngPathDic).value)
    
    'Move the dictionary data
    Call Helpers.MoveData(SetupWkb, DesWkb, C_sParamSheetDict, C_eStartLinesDictHeaders)
    'Move the Choices data
    Call Helpers.MoveData(SetupWkb, DesWkb, C_sParamSheetChoices, C_eStartLinesChoicesHeaders)
    'Move the Export data
    Call Helpers.MoveData(SetupWkb, DesWkb, C_sParamSheetExport, C_eStartLinesExportTitle)

    SetupWkb.Close savechanges:=False
    Set SetupWkb = Nothing

    iUpdateCpt = iUpdateCpt + 5

    StatusBar_Updater (iUpdateCpt)

    'translation of the Export, Dictionary and Choice sheets for the linelist
    'Call Translate_Manage

    '--------------- Getting all required the Data

    'Create the Dictionnary data
    Set DictHeaders = Helpers.GetHeaders(DesWkb, C_sParamSheetDict, 1)
    'Create the data table of linelist patient using the dictionnary
    Set DictData = Helpers.GetData(DesWkb, C_sParamSheetDict, 2)
    'Create the choices data
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadList")
    'Create the dictionnary for the choices sheet
    Set ChoicesHeaders = Helpers.GetHeaders(DesWkb, C_sParamSheetChoices, 1)
    'Create the table for the choices
    Set ChoicesData = Helpers.GetData(DesWkb, C_sParamSheetChoices, 2)
    'Reading the export sheet
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadExport")
    'Create parameters for export
    Set ExportData = Helpers.GetData(DesWkb, C_sParamSheetExport, 1)
    'Create the translation Data
    Set TransData = New BetterArray
    TransData.FromExcelRange DesWkb.Worksheets(C_sParamSheetTranslation).Cells(C_eStartlinestransdata, 2), DetectLastRow:=True, DetectLastColumn:=True
    DoEvents
    Set DesWkb = Nothing

    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_BuildLL")

    'Creating the linelist using the dictionnary and choices data as well as export data
    sPath = SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb"

    'required temporary folder for analysis
    On Error Resume Next
        RmDir SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_"
        MkDir SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_" 'create a folder for sending all the data from designer
    On Error GoTo 0

    Call DesignerBuildList.BuildList(DictHeaders, DictData, ExportData, ChoicesHeaders, ChoicesData, TransData, sPath)
    DoEvents

    EndWork xlsapp:=Application
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLCreated")

    On Error Resume Next
        RmDir SheetMain.Range(C_sRngLLDir) & Application.PathSeparator & "LinelistApp_"
    On Error GoTo 0

    Call SetInputRangesToWhite

    StatusBar_Updater (100)

    iOpenLL = MsgBox(TranslateMsg("MSG_OpenLL") & " " & sPath & " ?", vbQuestion + vbYesNo, "Linelist")

    If iOpenLL = vbYes Then
        Call OpenLL
    End If

    'Setting the memory data to nothing
    Set DictHeaders = Nothing
    Set DictData = Nothing
    Set ChoicesHeaders = Nothing
    Set ChoicesData = Nothing
    Set ExportData = Nothing
    Set TransData = Nothing

End Sub

'Adding some controls before generating the linelist  ==================================================================================================================================

'Adding some controls before generating the linelist  =============================================================================================================================
Public Sub Control()
    'Put every range in white before the control

    'Put every range in white before the control
    Call SetInputRangesToWhite

    Dim bGood As Boolean

    'Control to be sure we can generate a linelist
    bGood = DesignerMainHelpers.ControlForGenerate()
    If Not bGood Then
        Exit Sub
    Else

        'Now that everything is fine, continue to generation process but issue a warning in case it will
        'replace the previous existing file

        Call SetInputRangesToWhite

        SheetMain.Range(C_sRngLLName).value = FileNameControl(SheetMain.Range(C_sRngLLName).value)

        If Dir(SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb") <> "" Then
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Correct") & ": " & SheetMain.Range(C_sRngLLName).value & ".xlsb " & TranslateMsg("MSG_Exists")
            SheetMain.Range(C_sRngEdition).Interior.Color = Helpers.GetColor("Grey")
            If MsgBox(SheetMain.Range(C_sRngLLName).value & ".xlsb " & TranslateMsg("MSG_Exists") & Chr(10) & TranslateMsg("MSG_Question"), vbYesNo, _
            TranslateMsg("MSG_Title")) = vbNo Then
                SheetMain.Range(C_sRngLLName).value = ""
                SheetMain.Range(C_sRngLLName).Interior.Color = GetColor("RedEpi")
                Exit Sub
            End If
        Else
            SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Correct")
        End If

        Call GenerateData

    End If
End Sub

'OPEN THE GENERATED LINELIST =========================================================================================================================================================

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
        Exit Sub
    End If

    On Error GoTo no
    'Then open it
    Application.Workbooks.Open Filename:=SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb", ReadOnly:=False
no:
    Exit Sub

End Sub



Sub ResetField()

    SheetMain.Range(C_sRngPathDic).value = ""
    SheetMain.Range(C_sRngPathGeo).value = ""
    SheetMain.Range(C_sRngLLName).value = ""
    SheetMain.Range(C_sRngLLDir).value = ""
    SheetMain.Range(C_sRngEdition).value = ""
    SheetMain.Range(C_sRngUpdate).value = ""

    SheetMain.Range(C_sRngPathGeo).Interior.Color = vbWhite
    SheetMain.Range(C_sRngPathDic).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLName).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLDir).Interior.Color = vbWhite
    SheetMain.Range(C_sRngEdition).Interior.Color = vbWhite
    SheetMain.Range(C_sRngUpdate).Interior.Color = vbWhite

End Sub



