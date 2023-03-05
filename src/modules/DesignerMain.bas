Attribute VB_Name = "DesignerMain"
Option Explicit
Option Private Module

Public iUpdateCpt As Integer
Public bGeobaseIsImported As Boolean

'LOADING FILES AND FOLDERS ============================================================================================================================================================================
Private Function TranslateMsg(ByVal msgCode As String)
    
    'Translate a message in the designer
    Dim destrans As IDesTranslation
    Dim trads As ITranslation
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("DesignerTranslation")
    Set destrans = DesTranslation.Create(sh)
    Set trads = destrans.TransObject()
    TranslateMsg = trads.TranslatedValue(msgCode)
End Function

'Import the language of the setup
Private Sub ImportLang()
    Dim inPath As String
    Dim Wkb As Workbook
    Dim Lo As ListObject
    Dim langTable As BetterArray
 
    inPath = SheetMain.Range("RNG_PathDico").Value

    On Error Resume Next
    BeginWork xlsapp:=Application
    
    'Set Events to false to avoid
    Application.EnableEvents = False
    Set Wkb = Workbooks.Open(inPath)
    
    On Error GoTo 0
    
    If Wkb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set Lo = Wkb.Worksheets("Translations").ListObjects(1)
    On Error GoTo 0
 
    If Lo Is Nothing Then Exit Sub
 
    Set langTable = New BetterArray
    langTable.FromExcelRange Lo.HeaderRowRange
    langTable.ToExcelRange ThisWorkbook.Worksheets("DesignerTranslation").Range("T_LanguageDictionary").Cells(1, 1)
    SheetMain.Range("RNG_LangSetup").Value = langTable.Item(langTable.LowerBound)

    'Add the language to LLTranslations
    SheetLLTranslation.Range("RNG_DictionaryLanguage").Value = SheetMain.Range("RNG_LangSetup").Value
 
    Wkb.Close savechanges:=False
    
    Application.EnableEvents = True
 
End Sub

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

    io.LoadFile "*.xlsb"                         '
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

Sub GenerateData()
    
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

    BeginWork xlsapp:=Application

    Application.Cursor = xlWait
    Application.EnableEvents = False
    
    Set wb = ThisWorkbook
    Set lData = LinelistSpecs.Create(wb)
    Set dict = lData.Dictionary() 'Dictionary
    Set llshs = LLSheets.Create(dict) 'The worksheets object of the dictionary
    Set mainobj = lData.MainObject() 'The main object is an object for dealing with the main sheet interface
    Set llana = lData.Analysis() 'Linelist analysis object

    'Create the designer translation object
    Set desTrads = DesTranslation.Create(wb.Worksheets("DesignerTranslation"))

    'After preparation steps, update the status
    mainobj.UpdateStatus (5) '5% after preparation steps are done

    'Add informations on the preparing step to the end user
    mainobj.AddInfo desTrads, "MSG_ReadSetup"

    'Preparing the setup and specification files
    lData.Prepare

    'Preparing the linelist file
    Set ll = Linelist.Create(lData)

    mainobj.AddInfo desTrads, "MSG_PreparLL"

    'If you want to change the behavior of the linelist, please go to the
    'linelist class instead of using functions here.

    ll.Prepare
    
    mainobj.UpdateStatus (10)

    'Should add Error management when something goes wrong
    mainobj.AddInfo desTrads, "MSG_HListVList"

    'On Error GoTo ErrorBuildingLLManage

    currSheetName = dict.DataRange("sheet name").Cells(1, 1).Value
    If llshs.SheetInfo(currSheetName) = "vlist1D" Then
        Set buildingSheet = Vlist.Create(currSheetName, ll)
    ElseIf llshs.SheetInfo(currSheetName) = "hlist2D" Then
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

        If llshs.SheetInfo(currSheetName) = "vlist1D" Then
            Set buildingSheet = Vlist.Create(currSheetName, ll)
        ElseIf llshs.SheetInfo(currSheetName) = "hlist2D" Then
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
    If MsgBox(TranslateMsg("MSG_OpenLL") & " " & outPath & " ?", vbQuestion + vbYesNo, "Linelist") = vbYes Then OpenLL

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


'Adding some controls before generating the linelist  =================================================================================================================================================
Public Sub Control()
    
  
    Dim mainobj As IMain
    Dim desTrads As IDesTranslation
    Dim trads As ITranslation
    Dim wb As Workbook
    Dim sh As Worksheet

    'Put every range in white before the control
    Call SetInputRangesToWhite
    
    'Create Main object
    Set wb = ThisWorkbook
    Set sh = wb.Worksheets("Main")
    Set mainobj = Main.Create(sh)
    
    'Create the designer translation object
    Set sh = wb.Worksheets("DesignerTranslation")
    Set desTrads = DesTranslation.Create(sh)
    Set trads = desTrads.TransObject(TranslationOfMessages)
    
    'Check readiness of the linelist
    mainobj.CheckReadiness desTrads
    
    'If the main sheet is not ready exit the sub
    If Not mainobj.Ready Then Exit Sub

    If Dir(SheetMain.Range("RNG_LLDir").Value & _
           Application.PathSeparator & _
           SheetMain.Range("RNG_LLName").Value & ".xlsb") <> "" Then
       
        SheetMain.Range("RNG_Edition").Value = trads.TranslatedValue("MSG_Correct") & ": " _
                                                                                  & SheetMain.Range("RNG_LLName").Value & ".xlsb " _
                                                                                  & trads.TranslatedValue("MSG_Exists")
                                                
        SheetMain.Range("RNG_Edition").Interior.color = RGB(235, 232, 232)
        
        If MsgBox(SheetMain.Range("RNG_LLName").Value & ".xlsb " & _
                  trads.TranslatedValue("MSG_Exists") & chr(10) & _
                  trads.TranslatedValue("MSG_Question"), vbYesNo, _
                  trads.TranslatedValue("MSG_Title")) = vbNo Then
            
            SheetMain.Range("RNG_LLName").Value = ""
            SheetMain.Range("RNG_LLName").Interior.color = RGB(252, 228, 214)
            Exit Sub
        End If
        
    Else
        SheetMain.Range("RNG_Edition").Value = trads.TranslatedValue("MSG_Correct")
    End If
    
    'Generate all the data
    GenerateData
    
End Sub

'OPEN THE GENERATED LINELIST ==========================================================================================================================================================================

Sub OpenLL()
    'Be sure that the directory and the linelist name are not empty
    If SheetMain.Range("RNG_LLDir").Value = "" Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_PathLL")
        SheetMain.Range("RNG_LLDir").Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    If SheetMain.Range("RNG_LLName").Value = "" Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_LLName")
        SheetMain.Range("RNG_LLName").Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    'Be sure the workbook is not already opened
    If IsWkbOpened(SheetMain.Range("RNG_LLName").Value & ".xlsb") Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_CloseLL")
        SheetMain.Range("RNG_LLName").Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    'Be sure the workbook exits
    If Dir(SheetMain.Range("RNG_LLDir").Value & Application.PathSeparator & SheetMain.Range("RNG_LLName").Value & ".xlsb") = "" Then
        SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_CheckLL")
        SheetMain.Range("RNG_LLName").Interior.color = RGB(252, 228, 214)
        SheetMain.Range("RNG_LLDir").Interior.color = RGB(252, 228, 214)
        Exit Sub
    End If

    On Error GoTo no
    'Then open it
    Application.Workbooks.Open FileName:=SheetMain.Range("RNG_LLDir").Value & Application.PathSeparator & SheetMain.Range("RNG_LLName").Value & ".xlsb"
    Exit Sub
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

'Set All the Input ranges to white
Sub SetInputRangesToWhite()

    SheetMain.Range("RNG_PathGeo").Interior.color = vbWhite
    SheetMain.Range("RNG_PathDico").Interior.color = vbWhite
    SheetMain.Range("RNG_LLName").Interior.color = vbWhite
    SheetMain.Range("RNG_LLDir").Interior.color = vbWhite
    SheetMain.Range("RNG_Edition").Interior.color = vbWhite

End Sub


