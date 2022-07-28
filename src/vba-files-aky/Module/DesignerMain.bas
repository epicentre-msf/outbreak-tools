Attribute VB_Name = "DesignerMain"
Option Explicit

Public iUpdateCpt As Integer
Public bGeobaseIsImported As Boolean

'LOADING FILES AND FOLDERS ============================================================================================================================================================================

'Loading the Dictionnary File _________________________________________________________________________________________________________________________________________________________________________
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
        Call ImportLang
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If

    EndWork xlsapp:=Application
End Sub

'Loading a linelist File ______________________________________________________________________________________________________________________________________________________________________________
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

'Loading the Lineist Directory ________________________________________________________________________________________________________________________________________________________________________
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

'Loading the Geobase  _________________________________________________________________________________________________________________________________________________________________________________
'
' Adding a new load geo for the Geo file, in a new sheet called Geo2
' we have two functions for loading the geodatabase, but the second one
' will be in use instead of the first one.
'
Sub LoadGeoFile()

    Dim sFilePath   As String                      'File path to the geo file
    sFilePath = Helpers.LoadFile("*.xlsx")

    If sFilePath <> vbNullString Then
        'Open the geo workbook and hide the windows
        SheetMain.Range(C_sRngPathGeo).value = sFilePath
        Call ImportGeobase
        bGeobaseIsImported = True
    Else
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub


'Function to import the geobase

Public Sub ImportGeobase()

    Dim sFilePath As String
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

    BeginWork xlsapp:=Application

    AdmNames.LowerBound = 1
    AdmNames.Push C_sAdm1, C_sAdm2, C_sAdm3, C_sAdm4, C_sHF, C_sNames, C_sHistoHF, C_sHistoGeo, C_sGeoMetadata

    'Path to the geobase
    sFilePath = SheetMain.Range(C_sRngPathGeo).value

    'Be sure there is a geobase before proceeding, otherwhise, build the linelist without a geobase
    If sFilePath = vbNullString Then Exit Sub


    SheetGeo.Range(C_sRngGeoName).value = Dir(sFilePath)

    Set Wkb = Workbooks.Open(sFilePath)
    'Write the filename of the geobase somewhere for the export
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

    Wkb.Close SaveChanges:=False
    Set AdmData = Nothing
    Set AdmHeader = Nothing
    Set AdmNames = Nothing
    Set Wkb = Nothing
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Fini")
    SheetMain.Range(C_sRngPathGeo).Interior.Color = GetColor("White")

    Call TranslateHeadGeo
    EndWork xlsapp:=Application
End Sub

'GENERATE THE LINELIST DATA ===========================================================================================================================================================================

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This is the Sub for generating the data of the linelist using the input in the designer. The main entry point is the BuildList function which creates the Linelist-patient sheet as well as all
'the forms in the linelist
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub GenerateData(Optional iAsk As Byte = 0)

    Dim bGood As Boolean
    bGood = DesignerMainHelpers.ControlForGenerate()

    If Not bGood Then
        Exit Sub
    End If

    'Import the geobase if it is not imported
    If Not bGeobaseIsImported Then Call ImportGeobase

    Dim DictHeaders     As BetterArray          'Dictionary headers
    Dim DictData        As BetterArray          'Dictionary data
    Dim ChoicesHeaders  As BetterArray          'Choices headers
    Dim ChoicesData     As BetterArray          'Choices data
    Dim ExportData      As BetterArray          'Export data
    Dim TransData       As BetterArray          'Translation data
    Dim GSData          As BetterArray          'Global Summary Data
    Dim UAData          As BetterArray          'Univariate Analaysis Data
    Dim BAData          As BetterArray          'Bivariate Analysis DAta
    Dim sPath           As String
    Dim SetupWkb        As Workbook
    Dim DesWkb          As Workbook
    Dim iOpenLL         As Integer
    Dim i               As Integer
    Dim previousSecurity As Byte

    'Be sure the actual Workbook is not opened


    SheetMain.Range(C_sRngUpdate).value = vbNullString

    If IsWkbOpened(SheetMain.Range(C_sRngLLName).value & ".xlsb") Then
        SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range(C_sRngLLName).Interior.Color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If

    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_MovingData")

    iUpdateCpt = 0
    StatusBar_Updater (iUpdateCpt)

    BeginWork xlsapp:=Application

    Set DesWkb = DesignerWorkbook

    previousSecurity = Application.AutomationSecurity
    'Set security before opening  the setup workbook
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set SetupWkb = Workbooks.Open(SheetMain.Range(C_sRngPathDic).value)

    'Move the dictionary data
    Call Helpers.MoveData(SetupWkb, DesWkb, C_sParamSheetDict, C_eStartLinesDictHeaders)
    'Move the Choices data
    Call Helpers.MoveData(SetupWkb, DesWkb, C_sParamSheetChoices, C_eStartLinesChoicesHeaders)
    'Move the Export data
    Call Helpers.MoveData(SetupWkb, DesWkb, C_sParamSheetExport, C_eStartLinesExportTitle)

    Call DesignerMainHelpers.MoveAnalysis(SetupWkb)

    SetupWkb.Close SaveChanges:=False
    Set SetupWkb = Nothing

    Application.AutomationSecurity = previousSecurity

    iUpdateCpt = iUpdateCpt + 5
    StatusBar_Updater (iUpdateCpt)

    'Translating the linelist
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_Translating")

    'Translate
    Call TranslateLinelistData

    'Add the tables for every Sheets
    Call DesignerMainHelpers.AddTableNames

    ' Getting all required the Data ___________________________________________________________________________________________________________________________________________________________________

    'Create the Dictionnary data
    Set DictHeaders = Helpers.GetHeaders(DesWkb, C_sParamSheetDict, 1)
    'Create the data table of linelist patient using the dictionnary
    Set DictData = Helpers.GetData(DesWkb, C_sParamSheetDict, 2, DictHeaders.Length)
    'Create the choices data
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadList")
    'Create the dictionnary for the choices sheet
    Set ChoicesHeaders = Helpers.GetHeaders(DesWkb, C_sParamSheetChoices, 1)
    'Create the table for the choices
    Set ChoicesData = Helpers.GetData(DesWkb, C_sParamSheetChoices, 2, ChoicesHeaders.Length)
    'Reading the export sheet
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_ReadExport")
    'Create parameters for export
    Set ExportData = Helpers.GetData(DesWkb, C_sParamSheetExport, 1)
    'Create the translation Data
    Set TransData = New BetterArray
    TransData.FromExcelRange DesWkb.Worksheets(C_sParamSheetTranslation).ListObjects(C_sTabTranslation).Range

    'Filters data for analysis

    'Global summary data for analysis
    Set GSData = New BetterArray
    GSData.LowerBound = 1
    GSData.FromExcelRange DesWkb.Worksheets(C_sParamSheetAnalysis).ListObjects(C_sTabGS).Range

    'Bivariate and Univariate Analysis data for Analysis
    Set UAData = New BetterArray
    UAData.LowerBound = 1
    UAData.FromExcelRange DesWkb.Worksheets(C_sParamSheetAnalysis).ListObjects(C_sTabUA).Range

    Set BAData = New BetterArray
    BAData.LowerBound = 1
    BAData.FromExcelRange DesWkb.Worksheets(C_sParamSheetAnalysis).ListObjects(C_sTabBA).Range


    DoEvents

    Set DesWkb = Nothing

    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_BuildLL")

    'Creating the linelist using the dictionnary and choices data as well as export data
    sPath = SheetMain.Range(C_sRngLLDir).value & Application.PathSeparator & SheetMain.Range(C_sRngLLName).value & ".xlsb"

    Call PrepareTemporaryFolder
    Call SetUserDefineConstants

    Call BuildList(DictHeaders, DictData, ExportData, ChoicesHeaders, ChoicesData, TransData, GSData, UAData, BAData, sPath)

    DoEvents

    EndWork xlsapp:=Application
    SheetMain.Range(C_sRngEdition).value = TranslateMsg("MSG_LLCreated")

    Call PrepareTemporaryFolder(Create:=False)

    Call SetInputRangesToWhite

    StatusBar_Updater (100)

    If iAsk = 1 Then
        iOpenLL = MsgBox(TranslateMsg("MSG_OpenLL") & " " & sPath & " ?", vbQuestion + vbYesNo, "Linelist")

        If iOpenLL = vbYes Then
            Call OpenLL
        End If

    End If

    'Setting the memory data to nothing
    Set DictHeaders = Nothing
    Set DictData = Nothing
    Set ChoicesHeaders = Nothing
    Set ChoicesData = Nothing
    Set ExportData = Nothing
    Set TransData = Nothing
    Set GSData = Nothing
    Set BAData = Nothing

End Sub

'Adding some controls before generating the linelist  =================================================================================================================================================

'Adding some controls before generating the linelist  =================================================================================================================================================
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

        Call GenerateData(1)

    End If
End Sub

'OPEN THE GENERATED LINELIST ==========================================================================================================================================================================

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

    SheetMain.Range(C_sRngPathDic).value = vbNullString
    SheetMain.Range(C_sRngPathGeo).value = vbNullString
    SheetMain.Range(C_sRngLLName).value = vbNullString
    SheetMain.Range(C_sRngLLDir).value = vbNullString
    SheetMain.Range(C_sRngEdition).value = vbNullString
    SheetMain.Range(C_sRngUpdate).value = vbNullString
    SheetMain.Range(C_sRngLangSetup).value = vbNullString

    SheetMain.Range(C_sRngPathGeo).Interior.Color = vbWhite
    SheetMain.Range(C_sRngPathDic).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLName).Interior.Color = vbWhite
    SheetMain.Range(C_sRngLLDir).Interior.Color = vbWhite
    SheetMain.Range(C_sRngEdition).Interior.Color = vbWhite
    SheetMain.Range(C_sRngUpdate).Interior.Color = vbWhite

End Sub



