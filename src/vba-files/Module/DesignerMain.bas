Attribute VB_Name = "DesignerMain"
Option Explicit
Option Private Module

Public iUpdateCpt As Integer
Public bGeobaseIsImported As Boolean

'LOADING FILES AND FOLDERS ============================================================================================================================================================================

'Loading the Dictionnary File _________________________________________________________________________________________________________________________________________________________________________
Sub LoadFileDic()

    BeginWork xlsapp:=Application
    Dim io As IOSFiles
    Set io = OSFiles.Create()

    io.LoadFile "*.xlsb"

    'Update messages if the file path is correct
    If io.HasValidFile Then
        SheetMain.Range("RNG_PathDico").Value = io.File
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_ChemFich")
        SheetMain.Range("RNG_PathDico").Interior.color = vbWhite
        'Import the languages after loading the setup file
        ImportLang
    Else
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_OpeAnnule")
    End If
    EndWork xlsapp:=Application
End Sub

'Loading a linelist File ______________________________________________________________________________________________________________________________________________________________________________
Sub LoadFileLL()

    Dim io As IOSFiles
    Set io = OSFiles.Create()

    io.LoadFile "*.xlsb"       '
    If Not io.HasValidFile Then Exit Sub

    On Error GoTo ErrorManage
    Application.Workbooks.Open FileName:=io.File(), ReadOnly:=False
    Exit Sub
ErrorManage:
    MsgBox TranslateMsg("MSG_TitlePassWord"), vbCritical, TranslateMsg("MSG_PassWord")
End Sub

'Loading the Lineist Directory ________________________________________________________________________________________________________________________________________________________________________
Sub LinelistDir()
    Dim io As IOSFiles
    Set io = OSFiles.Create()
    io.LoadFolder

    SheetMain.Range("RNG_LLDir") = vbNullString

    If (io.HasValidFolder) Then
        SheetMain.Range("RNG_LLDir").Value = io.Folder()
        SheetMain.Range("RNG_LLDir").Interior.color = vbWhite
    Else
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

'Loading the Geobase  _________________________________________________________________________________________________________________________________________________________________________________
Sub LoadGeoFile()
    Dim io As IOSFiles
    Set io = OSFiles.Create()
    
    io.LoadFile "*.xlsx"
    
    If io.HasValidFile Then
        SheetMain.Range("RNG_PathGeo").Value = io.File()
    Else
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_OpeAnnule")
    End If
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub GenerateData(Optional iAsk As Byte = 0)

    Dim bGood As Boolean
    bGood = DesignerMainHelpers.ControlForGenerate()

    If Not bGood Then
        Exit Sub
    End If

    'Import the geobase if it is not imported
    If Not bGeobaseIsImported Then Call ImportGeobase

    Dim sPath           As String
    Dim DesWkb          As Workbook
    Dim iOpenLL         As Integer
    Dim previousSecurity As Byte

    'Be sure the actual Workbook is not opened
    SheetMain.Range("RNG_Update").Value = vbNullString

    If IsWkbOpened(SheetMain.Range("RNG_LLName").Value & ".xlsb") Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range("RNG_LLName").Interior.color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If

    SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_MovingData")

    iUpdateCpt = 0
    StatusBar_Updater (iUpdateCpt)

    BeginWork xlsapp:=Application

    Set DesWkb = DesignerWorkbook

    previousSecurity = Application.AutomationSecurity
    'Set security before opening  the setup workbook
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set setupWkb = Workbooks.Open(SheetMain.Range("RNG_PathDico").Value)

    'Move the dictionary data
    Call Helpers.MoveData(setupWkb, DesWkb, C_sParamSheetDict, C_eStartLinesDictHeaders)
    'Move the Choices data
    Call Helpers.MoveData(setupWkb, DesWkb, C_sParamSheetChoices, C_eStartLinesChoicesHeaders)
    'Move the Export data
    Call Helpers.MoveData(setupWkb, DesWkb, C_sParamSheetExport, C_eStartLinesExportTitle)

    Call DesignerMainHelpers.MoveAnalysis(setupWkb)

    setupWkb.Close SaveChanges:=False

    Application.AutomationSecurity = previousSecurity

    iUpdateCpt = iUpdateCpt + 5
    StatusBar_Updater (iUpdateCpt)

    'Translating the linelist
    SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_Translating")

    'Translate
    Call TranslateLinelistData

    'Add the tables for every Sheets
    Call DesignerMainHelpers.AddTableNames

    ' Getting all required the Data ___________________________________________________________________________________________________________________________________________________________________

    'Create the Dictionnary data
    Set DictHeaders = Helpers.GetHeaders(DesWkb, C_sParamSheetDict, 1)
    'Create the data table of linelist patient using the dictionnary
    Set dictData = Helpers.GetData(DesWkb, C_sParamSheetDict, 2, DictHeaders.Length)
    'Create the choices data
    SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_ReadList")
    'Create the dictionnary for the choices sheet
    Set ChoicesHeaders = Helpers.GetHeaders(DesWkb, C_sParamSheetChoices, 1)
    'Create the table for the choices
    Set ChoicesData = Helpers.GetData(DesWkb, C_sParamSheetChoices, 2, ChoicesHeaders.Length)
    'Reading the export sheet
    SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_ReadExport")
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

    'Time series and Spatial Analysis data
    Set TAData = New BetterArray                 'Time series analysis
    TAData.LowerBound = 1
    TAData.FromExcelRange DesWkb.Worksheets(C_sParamSheetAnalysis).ListObjects(C_sTabTA).Range

    Set SAData = New BetterArray                 'Spatial analysis
    SAData.LowerBound = 1
    SAData.FromExcelRange DesWkb.Worksheets(C_sParamSheetAnalysis).ListObjects(C_sTabTA).Range

    DoEvents

    SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_BuildLL")

    'Creating the linelist using the dictionnary and choices data as well as export data
    sPath = SheetMain.Range("RNG_LLDir").Value & Application.PathSeparator & SheetMain.Range("RNG_LLName").Value & ".xlsb"

    'Prepare the temporary folder for the linelist
    Call PrepareTemporaryFolder

    'Add some user define constants
    Call SetUserDefineConstants

    'Add the preprocessing step for the designer

    Call BuildList(DictHeaders, dictData, ExportData, ChoicesHeaders, ChoicesData, TransData, GSData, UAData, BAData, TAData, SAData, sPath)

    DoEvents

    EndWork xlsapp:=Application
    SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_LLCreated")

    Call PrepareTemporaryFolder(Create:=False)

    Call SetInputRangesToWhite

    StatusBar_Updater (100)

    If iAsk = 1 Then
        iOpenLL = MsgBox(TranslateMsg("MSG_OpenLL") & " " & sPath & " ?", vbQuestion + vbYesNo, "Linelist")

        If iOpenLL = vbYes Then
            Call OpenLL
        End If

    End If

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

        SheetMain.Range("RNG_LLName").Value = FileNameControl(SheetMain.Range("RNG_LLName").Value)

        If Dir(SheetMain.Range("RNG_LLDir").Value & Application.PathSeparator & SheetMain.Range("RNG_LLName").Value & ".xlsb") <> "" Then
            SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_Correct") & ": " & SheetMain.Range("RNG_LLName").Value & ".xlsb " & TranslateMsg("MSG_Exists")
            SheetMain.Range("RNG_Edition").Interior.color = Helpers.GetColor("Grey")
            If MsgBox(SheetMain.Range("RNG_LLName").Value & ".xlsb " & TranslateMsg("MSG_Exists") & Chr(10) & TranslateMsg("MSG_Question"), vbYesNo, _
                      TranslateMsg("MSG_Title")) = vbNo Then
                SheetMain.Range("RNG_LLName").Value = ""
                SheetMain.Range("RNG_LLName").Interior.color = GetColor("RedEpi")
                Exit Sub
            End If
        Else
            SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_Correct")
        End If

        Call GenerateData(1)

    End If
End Sub

'OPEN THE GENERATED LINELIST ==========================================================================================================================================================================

Sub OpenLL()
    'Be sure that the directory and the linelist name are not empty
    If SheetMain.Range("RNG_LLDir").Value = "" Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_PathLL")
        SheetMain.Range("RNG_LLDir").Interior.color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If

    If SheetMain.Range("RNG_LLName").Value = "" Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_LLName")
        SheetMain.Range("RNG_LLName").Interior.color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If

    'Be sure the workbook is not already opened
    If IsWkbOpened(SheetMain.Range("RNG_LLName").Value & ".xlsb") Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range("RNG_LLName").Interior.color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If

    'Be sure the workbook exits
    If Dir(SheetMain.Range("RNG_LLDir").Value & Application.PathSeparator & SheetMain.Range("RNG_LLName").Value & ".xlsb") = "" Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_CheckLL")
        SheetMain.Range("RNG_LLName").Interior.color = Helpers.GetColor("RedEpi")
        SheetMain.Range("RNG_LLDir").Interior.color = Helpers.GetColor("RedEpi")
        Exit Sub
    End If

    On Error GoTo no
    'Then open it
    Application.Workbooks.Open FileName:=SheetMain.Range("RNG_LLDir").Value & Application.PathSeparator & SheetMain.Range("RNG_LLName").Value & ".xlsb", ReadOnly:=False
no:
    Exit Sub

End Sub

Sub ResetField()

    SheetMain.Range("RNG_PathDico").Value = vbNullString
    SheetMain.Range("RNG_PathGeo").Value = vbNullString
    SheetMain.Range("RNG_LLName").Value = vbNullString
    SheetMain.Range("RNG_LLDir").Value = vbNullString
    SheetMain.Range("RNG_Edition").Value = vbNullString
    SheetMain.Range("RNG_Update").Value = vbNullString
    SheetMain.Range("RNG_LangSetup").Value = vbNullString

    SheetMain.Range("RNG_PathGeo").Interior.color = vbWhite
    SheetMain.Range("RNG_PathDico").Interior.color = vbWhite
    SheetMain.Range("RNG_LLName").Interior.color = vbWhite
    SheetMain.Range("RNG_LLDir").Interior.color = vbWhite
    SheetMain.Range("RNG_Edition").Interior.color = vbWhite
    SheetMain.Range("RNG_Update").Interior.color = vbWhite

End Sub


