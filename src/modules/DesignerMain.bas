Attribute VB_Name = "DesignerMain"
Option Explicit
Option Private Module

'Designer Translation sheet name
Private Const DESIGNERTRADSHEET As String = "DesignerTranslation"
'Setup translation sheet name
Private Const SETUPTRADSHEET As String = "Translations"
'Linelist translation sheet name
Private Const LINELISTTRADSHEET As String = "LinelistTranslation"
'Designer main sheet name
Private Const DESIGNERMAINSHEET As String = "Main"
'Range to the dictionary Path in the main sheet
Private Const RNGPATHDICO As String = "RNG_PathDico"
'Range for informations to user in the main sheet
Private Const RNGEDITION As String = "RNG_Edition"
'Linelist directory range name
Private Const RNGLLDIR As String = "RNG_LLDir"
'Linelist name range
Private Const RNGLLNAME As String = "RNG_LLName"
'Geobase path
Private Const RNGGEOPATH As String = "RNG_PathGeo"


'speed app
Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
End Sub

'LOADING FILES AND FOLDERS =====================================================
'Translate code messages in the designer
Public Function TranslateDesMsg(ByVal msgCode As String)
    'Translate a message in the designer
    Dim destrans As IDesTranslation
    Dim trads As ITranslation
    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets(DESIGNERTRADSHEET)
    Set destrans = DesTranslation.Create(sh)
    Set trads = destrans.TransObject()
    TranslateDesMsg = trads.TranslatedValue(msgCode)
End Function

'Import the language of the setup
Private Sub ImportLang()

    Const RNGLANGSETUP As String = "RNG_LangSetup" 'select the setup lang (mainsh)
    Const RNGDICTLANG As String = "RNG_DictionaryLanguage" 'selected lang (lltradsh)

    Dim inPath As String 'Path to the setup file, input path
    Dim actwb As Workbook 'actual workbook
    Dim impwb As Workbook 'imported setup workbook
    Dim tradLo As ListObject 'Translation listObject
    Dim langTable As BetterArray 'List of languages in the translation sheet
    Dim destradsh As Worksheet 'Worksheet for designer translation
    Dim lltradsh As Worksheet 'Worksheet for linelist translation
    Dim mainsh As Worksheet 'main worksheet
    Dim LangDictRng As Range 'in destradsh, range of languages in the setup

    Set actwb = ThisWorkbook
    Set mainsh = actwb.Worksheets(DESIGNERMAINSHEET)
    Set destradsh = actwb.Worksheets(DESIGNERTRADSHEET)
    Set lltradsh = actwb.Worksheets(LINELISTTRADSHEET)
    Set LangDictRng = destradsh.Range("LangDictList")
    inPath = mainsh.Range(RNGPATHDICO).Value

    On Error GoTo ExitImportLang
    Set impwb = Workbooks.Open(inPath)
    Set tradLo = impwb.Worksheets(SETUPTRADSHEET).ListObjects(1)
    Set langTable = New BetterArray
    langTable.FromExcelRange tradLo.HeaderRowRange
    langTable.ToExcelRange LangDictRng.Cells(1, 1)
    mainsh.Range(RNGLANGSETUP).Value = LangDictRng.Cells(1, 1).Value

    'Add the language to LLTranslations
    lltradsh.Range(RNGDICTLANG).Value = mainsh.Range(RNGLANGSETUP).Value

ExitImporLang:
    On Error Resume Next
    impwb.Close savechanges:=False
    On Error GoTo 0
End Sub

'@Description("Load the dictionary file")
'@EntryPoint
Public Sub LoadFileDic()

    BusyApp

    Dim io As IOSFiles
    Dim mainsh As Worksheet
    Dim wb As Workbook

    Set io = OSFiles.Create()
    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)

    io.LoadFile "*.xlsb"
    'Update messages if the file path is incorrect
    If io.HasValidFile Then
        mainsh.Range(RNGPATHDICO).Value = io.File
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_ChemFich")
        mainsh.Range(RNGPATHDICO).Interior.color = vbWhite
        'Import the languages after loading the setup file
        ImportLang
    Else
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_OpeAnnule")
    End If

    NotBusyApp
End Sub

'@Description("Path to future Lineist Directory")
'@EntryPoint
Sub LinelistDir()
    Dim wb As workbook
    Dim mainsh As Worksheet
    Dim io As IOSFiles

    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)

    Set io = OSFiles.Create()
    io.LoadFolder

    mainsh.Range(RNGLLDIR) = vbNullString

    If (io.HasValidFolder) Then
        mainsh.Range(RNGLLDIR).Value = io.Folder()
        mainsh.Range(RNGLLDIR).Interior.color = vbWhite
    Else
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_OpeAnnule")
    End If
End Sub

'@Description("Load the geobase")
'@EntryPoint
Public Sub LoadGeoFile()
    'Geobase path range name

    Dim wb As workbook
    Dim mainsh As Worksheet
    Dim io As IOSFiles

    Set io = OSFiles.Create()
    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)

    io.LoadFile "*.xlsx"
    If io.HasValidFile Then
        mainsh.Range(RNGGEOPATH).Value = io.File()
    Else
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_OpeAnnule")
    End If
End Sub

'GENERATE THE LINELIST =========================================================

'Generate the linelist after control
Private Sub GenerateData()

    Dim ll As ILinelist
    Dim lData As ILinelistSpecs
    Dim currSheetName As String
    Dim buildingSheet As Object
    Dim wb As Workbook
    Dim dict As ILLdictionary
    Dim llshs As ILLSheets
    Dim llana As ILLAnalysis
    Dim mainobj As IMain
    Dim outPath As String
    Dim nbOfSheets As Long
    Dim increment As Integer
    Dim statusValue As Integer
    Dim desTrads As IDesTranslation

   BusyApp

    Application.Cursor = xlWait

    Set wb = ThisWorkbook
    Set lData = LinelistSpecs.Create(wb)
    Set dict = lData.Dictionary() 'Dictionary
    Set mainobj = lData.MainObject() 'The main object is an object for dealing with the main sheet interface
    Set llana = lData.Analysis() 'Linelist analysis object

    'Create the designer translation object
    Set desTrads = DesTranslation.Create(wb.Worksheets(DESIGNERTRADSHEET))

    'After preparation steps, update the status
    mainobj.UpdateStatus (5) '5% after preparation steps are done

    'Add informations on the preparing step to the end user
    mainobj.AddInfo desTrads, "MSG_ReadSetup"

    'Preparing the setup and specification files
    lData.Prepare

    'Preparing the linelist file
    Set ll = Linelist.Create(lData)
    Set llshs = LLSheets.Create(dict) 'The worksheets object of the dictionary

    mainobj.AddInfo desTrads, "MSG_PreparLL"

    'If you want to change the behavior of the linelist, please go to the
    'linelist class instead of using functions here.

    ll.Prepare

    mainobj.UpdateStatus (10)

    'Should add Error management when something goes wrong
    mainobj.AddInfo desTrads, "MSG_HListVList"

    'On Error GoTo ErrorBuildingLLManage

    currSheetName = dict.DataRange("sheet name").Cells(1, 1).Value
    If llshs.sheetInfo(currSheetName) = "vlist1D" Then
        Set buildingSheet = Vlist.Create(currSheetName, ll)
    ElseIf llshs.sheetInfo(currSheetName) = "hlist2D" Then
        Set buildingSheet = Hlist.Create(currSheetName, ll)
    End If

    If buildingSheet Is Nothing Then Exit Sub

    mainobj.UpdateStatus (15)
    statusValue = 15
    nbOfSheets = dict.UniqueValues("sheet name").Length
    increment = CInt((80 - 15) / nbOfSheets)

    'Build the first sheet
    buildingSheet.Build
    statusValue = statusValue + increment
    mainobj.UpdateStatus statusValue


    'Loop through the other sheets and build them also
    Do While (buildingSheet.NextSheet() <> vbNullString)

        currSheetName = buildingSheet.NextSheet()

        If llshs.sheetInfo(currSheetName) = "vlist1D" Then
            Set buildingSheet = Vlist.Create(currSheetName, ll)
        ElseIf llshs.sheetInfo(currSheetName) = "hlist2D" Then
            Set buildingSheet = Hlist.Create(currSheetName, ll)
        End If

        'If you still remain on the same sheet exit (something weird happened)
        If currSheetName = buildingSheet.NextSheet() Then Exit Do
        buildingSheet.Build

        statusValue = statusValue + increment
        mainobj.UpdateStatus statusValue
    Loop

    'Save the linelist
    mainobj.AddInfo desTrads, "MSG_BuildAna"

    llana.Build ll
    ll.SaveLL

    'Update the status to 100%
    mainobj.UpdateStatus (100)

    Application.Cursor = xlDefault
    Application.EnableEvents = True

    'Open the linelist
    outPath = mainobj.OutputPath & Application.PathSeparator & mainobj.LinelistName & ".xlsb"
    If MsgBox(TranslateDesMsg("MSG_OpenLL") & " " & outPath & " ?", _
             vbQuestion + vbYesNo, "Linelist") = vbYes _
    Then OpenLL

    Exit Sub

ErrorBuildingLLManage:
        Application.Cursor = xlDefault
        Application.EnableEvents = True

        ll.ErrorManage
        Exit Sub

ErrorLinelistSpecsManage:
        Application.Cursor = xlDefault
        Application.EnableEvents = True

        lData.ErrorManage
        Exit Sub
End Sub

'@Description("Check everything is fine and generate the linelist")
'@EntryPoint
Public Sub Control()

    Dim mainobj As IMain
    Dim desTrads As IDesTranslation
    Dim trads As ITranslation
    Dim wb As Workbook
    Dim mainsh As Worksheet
    Dim tradsh As Worksheet

    'Put every range in white before the control
    SetInputRangesToWhite

    'Create Main object
    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set mainobj = Main.Create(mainsh)

    'Create the designer translation object
    Set tradsh = wb.Worksheets(DESIGNERTRADSHEET)
    Set desTrads = DesTranslation.Create(tradsh)
    Set trads = desTrads.TransObject(TranslationOfMessages)

    'Check readiness of the linelist
    mainobj.CheckReadiness desTrads

    'If the main sheet is not ready exit the sub
    If Not mainobj.Ready Then Exit Sub

    'Generate all the data in the other case
    GenerateData
End Sub

'OPEN THE GENERATED LINELIST ===================================================

Private Sub OpenLL()
    Dim wb As Workbook
    Dim mainsh As Worksheet

    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)

    'Be sure that the directory and the linelist name are not empty
    If mainsh.Range(RNGLLDIR).Value = "" Then
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_PathLL")
        mainsh.Range(RNGLLDIR).Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    If mainsh.Range(RNGLLNAME).Value = "" Then
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_LLName")
        mainsh.Range(RNGLLNAME).Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    'Be sure the workbook is not already opened
    If IsWkbOpened(mainsh.Range(RNGLLNAME).Value & ".xlsb") Then
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_CloseLL")
        mainsh.Range(RNGLLNAME).Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    'Be sure the workbook exits
    If Dir(mainsh.Range(RNGLLDIR).Value & Application.PathSeparator & mainsh.Range(RNGLLNAME).Value & ".xlsb") = "" Then
        mainsh.Range(RNGEDITION).Value = TranslateDesMsg("MSG_CheckLL")
        mainsh.Range(RNGLLNAME).Interior.color = RGB(252, 228, 214)
        mainsh.Range(RNGLLDIR).Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    On Error GoTo no
    'Then open it
    Application.Workbooks.Open FileName:=mainsh.Range(RNGLLDIR).Value & Application.PathSeparator & mainsh.Range(RNGLLNAME).Value & ".xlsb"
    Exit Sub
no:
    Exit Sub

End Sub

'Set All the Input ranges to white
Public Sub SetInputRangesToWhite()

End Sub


