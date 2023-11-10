Attribute VB_Name = "DesignerMain"
Option Explicit

'Designer Translation sheet name
Private Const DESIGNERTRADSHEET As String = "DesignerTranslation"
'Setup translation sheet name
Private Const SETUPTRADSHEET As String = "Translations"
'Linelist translation sheet name
Private Const LINELISTTRADSHEET As String = "LinelistTranslation"
'Designer main sheet name
Private Const DESIGNERMAINSHEET As String = "Main"

'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault, _
                    Optional ByVal changeSeparator As Boolean = False)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.cursor = cursor

    'The default is to change the separator back to English
    'And to return it back after all
    If changeSeparator Then
       Application.DecimalSeparator = "."
       Application.useSystemSeparators = False
    End If
End Sub

'Return back to previous state
Private Sub NotBusyApp(Optional ByVal returnSeparator As String = vbNullString, _
                       Optional ByVal useSystemSeparators As Boolean = False)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault

    If returnSeparator <> vbNullString Then
        Application.DecimalSeparator = returnSeparator
        Application.useSystemSeparators = useSystemSeparators
    End If

End Sub

'LOADING FILES AND FOLDERS =====================================================
'@Description("Import the language of the setup")
Private Sub ImportLang()
    Attribute ImportLang.VB_Description = "Import the language of the setup"

    Const RNGDICTLANG As String = "RNG_DictionaryLanguage" 'selected lang (lltradsh)

    Dim inPath As String 'Path to the setup file, input path
    Dim actwb As Workbook 'actual workbook
    Dim impwb As Workbook 'imported setup workbook
    Dim tradLo As listObject 'Translation listObject
    Dim langTable As BetterArray 'List of languages in the translation sheet
    Dim LangDictRng As Range 'in destradsh, range of languages in the setup
    Dim mainobj As IMain
    Dim trads As IDesTranslation

    Set actwb = ThisWorkbook
    Set mainobj = Main.Create(actwb.Worksheets(DESIGNERMAINSHEET))
    Set trads = DesTranslation.Create(actwb.Worksheets(DESIGNERTRADSHEET))
    Set LangDictRng = trads.LangListRng() 'This is one cell

    inPath = mainobj.ValueOf("setuppath")

    On Error GoTo ExitImportLang
    Set impwb = Workbooks.Open(inPath)
    Set tradLo = impwb.Worksheets(SETUPTRADSHEET).ListObjects(1)

    'add the list of languages
    Set langTable = New BetterArray
    langTable.FromExcelRange tradLo.HeaderRowRange
    langTable.ToExcelRange LangDictRng

    mainobj.AddInfo trads, LangDictRng.Value, "setuplang"

    'Add the language to LLTranslations
    actwb.Worksheets(LINELISTTRADSHEET).Range(RNGDICTLANG).Value = _
    LangDictRng.Value

ExitImportLang:
    On Error Resume Next
    impwb.Close savechanges:=False
    On Error GoTo 0
End Sub

'@Description("Load the dictionary file")
'@EntryPoint
Public Sub LoadFileDic()
    Attribute LoadFileDic.VB_Description = "Load the dictionary file"

    BusyApp

    Dim io As IOSFiles
    Dim mainsh As Worksheet
    Dim wb As Workbook
    Dim tradsh As Worksheet
    Dim trads As IDesTranslation
    Dim mainobj As IMain

    Set io = OSFiles.Create()
    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set tradsh = wb.Worksheets(DESIGNERTRADSHEET)
    Set trads = DesTranslation.Create(tradsh)
    Set mainobj = Main.Create(mainsh)

    io.LoadFile "*.xlsb"
    'Update messages if the file path is incorrect
    If io.HasValidFile Then
        mainobj.AddInfo trads, io.File, "setuppath"
        mainobj.AddInfo trads, "MSG_ChemFich", "edition"
        'Import the languages after loading the setup file
        ImportLang
    Else
        mainobj.AddInfo trads, "MSG_OpeAnnule", "edition"
    End If

    NotBusyApp
End Sub

'@Description("Load the template file")
'@EntryPoint
Public Sub LoadTemplateFile()
    Attribute LoadTemplateFile.VB_Description = "Load the template file"

    BusyApp

    Dim io As IOSFiles
    Dim mainsh As Worksheet
    Dim wb As Workbook
    Dim tradsh As Worksheet
    Dim trads As IDesTranslation
    Dim mainobj As IMain

    Set io = OSFiles.Create()
    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set tradsh = wb.Worksheets(DESIGNERTRADSHEET)
    Set trads = DesTranslation.Create(tradsh)
    Set mainobj = Main.Create(mainsh)

    io.LoadFile "*.xlsb"
    'Update messages if the file path is incorrect
    If io.HasValidFile Then
        mainobj.AddInfo trads, io.File, "temppath"
        'The user can write _default in the cell, it will not change the path
        mainobj.AddInfo trads, "MSG_ChemFich", "edition"
    Else
        mainobj.AddInfo trads, "MSG_OpeAnnule", "edition"
    End If

    NotBusyApp


    NotBusyApp
End Sub

'@Description("Path to future Lineist Directory")
'@EntryPoint
Sub LinelistDir()
    Attribute LinelistDir.VB_Description = "Path to future Lineist Directory"

    Dim wb As Workbook
    Dim mainsh As Worksheet
    Dim io As IOSFiles
    Dim trads As IDesTranslation
    Dim mainobj As IMain

    BusyApp

    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set trads = DesTranslation.Create(wb.Worksheets(DESIGNERTRADSHEET))
    Set mainobj = Main.Create(mainsh)

    Set io = OSFiles.Create()
    io.LoadFolder

    If (io.HasValidFolder) Then
        mainobj.AddInfo trads, io.Folder, "lldir"
    Else
        mainobj.AddInfo trads, "MSG_OpeAnnule", "edition"
    End If

    NotBusyApp
End Sub

'@Description("Load the geobase")
'@EntryPoint
Public Sub LoadGeoFile()
    Attribute LoadGeoFile.VB_Description = "Load the geobase"
    'Geobase path range name

    Dim wb As Workbook
    Dim mainsh As Worksheet
    Dim io As IOSFiles
    Dim trads As IDesTranslation
    Dim mainobj As IMain

    BusyApp

    Set io = OSFiles.Create()
    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set mainobj = Main.Create(mainsh)
    Set trads = DesTranslation.Create(wb.Worksheets(DESIGNERTRADSHEET))

    io.LoadFile "*.xlsx"
    If io.HasValidFile Then
        mainobj.AddInfo trads, io.File, "geopath"
    Else
        mainobj.AddInfo trads, "MSG_OpeAnnule", "edition"
    End If

    NotBusyApp
End Sub

'GENERATE THE LINELIST =========================================================

'Generate the linelist after passing through all the control and checkings
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
    Dim savedSeparator As String
    Dim savedUseSep As Boolean

    'Change decimal separators for building process
    savedSeparator = Application.DecimalSeparator
    savedUseSep = Application.useSystemSeparators

    BusyApp cursor:=xlWait, changeSeparator:=True
    
    
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

    On Error GoTo ErrorBuildingLLManage

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
    
    NotBusyApp returnSeparator:=savedSeparator, useSystemSeparators:=savedUseSep
    
    If (mainobj.ValueOf("askopen") = "yes") Then
        'Open the linelist
        outPath = mainobj.ValueOf("lldir") & Application.PathSeparator & mainobj.ValueOf("llname") & ".xlsb"
        If MsgBox(desTrads.TranslationMsg("MSG_OpenLL") & " " & outPath & " ?", _
                 vbQuestion + vbYesNo, "Linelist") = vbYes _
        Then mainobj.OpenLL
    End If


    NotBusyApp
    Exit Sub

ErrorBuildingLLManage:
        NotBusyApp returnSeparator:=savedSeparator
        ll.ErrorManage
        Exit Sub

ErrorLinelistSpecsManage:
        NotBusyApp
        lData.ErrorManage
        Exit Sub
End Sub

'@Description("Check everything is fine and generate the linelist")
'@EntryPoint
Public Sub Control()
    Attribute Control.VB_Description = "Check everything is fine and generate the linelist"

    Dim mainobj As IMain
    Dim desTrads As IDesTranslation
    Dim wb As Workbook
    Dim mainsh As Worksheet
    Dim tradsh As Worksheet


    'Create Main object
    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set mainobj = Main.Create(mainsh)
    
    'Put every range in white before the control
    mainobj.ClearInputRanges

    'Create the designer translation object
    Set tradsh = wb.Worksheets(DESIGNERTRADSHEET)
    Set desTrads = DesTranslation.Create(tradsh)

    'Check readiness of the linelist
    mainobj.CheckReadiness desTrads
    
    'If the main sheet is not ready exit the sub
    If Not mainobj.Ready Then Exit Sub

    mainobj.CheckFileExistence desTrads
    'If the main sheet is not ready exit the sub
    If Not mainobj.Ready Then Exit Sub

    mainobj.CheckRibbonExistence desTrads

    'If the main sheet is not ready exit the sub
    If Not mainobj.Ready Then Exit Sub

    'Generate all the data in the other case
    GenerateData
End Sub
