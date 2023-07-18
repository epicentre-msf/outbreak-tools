VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ImpSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Specifications for import in the linelist"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Specifications for import in the linelist")

Option Explicit
Implements IImpSpecs

Private Type TImp
  currwb As Workbook
  repfobj As Object
  adfobj As Object
  llFilePath As String
  geoFilePath As String
  geohistFilePath As String
  needReport As Boolean
End Type

Private Const CLASSNAME                 As String = "ImpSpecs"
Private Const DICTIONARYSHEET           As String = "Dictionary"
Private Const CHOICESSHEET              As String = "Choices"
Private Const DICTTRANSLATIONSHEET      As String = "Translations"
Private Const MESSTRANSLATIONSHEET      As String = "LinelistTranslation"
Private Const METADATASHEET             As String = "Metadata"
Private Const GEOSHEET                  As String = "Geo"
Private Const PASSWORDSHEET             As String = "__pass" 'password sheets
Private Const IMPREPORTSHEET               As String = "imp_rep__" 'temporary boggy sheet for report
Private Const SHOWHIDESHEET             As String = "show_hide__"
Private Const UPDATESHEET               As String = "updates__"

Private RESERVEDSHEETS                  As BetterArray
Private this                            As TImp
Private IMPORTEDVARIABLES               As BetterArray
Private IMPORTEDSHEETS                  As BetterArray

'Exposed methods
'repfobj is the form of the import Report
'adfobj is the advanced form Object
Public Function Create(ByVal repfobj As Object, ByVal adfobj As Object, ByVal currwb As Workbook) AS IImpSpecs

  CheckRequirements currwb, repfobj, adfobj
  
  'Initialize some objects and elements of the class,
  'mainly IMPORTED SHEETS/VARIABLES, RESERVED SHEETS and 
  'import Paths
  InitializeFileElements

  With New ImpSpecs
    Set .ReportFormObject = repfobj
    Set .AdvanceFormObject = adfobj
    Set .CurrentWorkbook = currwb
    Set Create = .Self()
  End With

End Function

Public Property Get Self() As IImpSpecs
  Set Self = Me
End Property

Public Property Get CurrentWorkbook() As Workbook
  Set CurrentWorkbook = this.currwb
End Property

Public Property Get ReportFormObject() As Object
  Set ReportFormObject = this.repfobj
End Property

Public Property Set ReportFormObject(ByVal fobj As Object)
  Set this.repfobj = fobj
End Property

Public Property Get AdvanceFormObject() As Object
  Set AdvanceFormObject = this.adobj
End Property

Public Property Set AdvanceFormObject(ByVal fobj As Object)
  Set this.adfobj = fobj
End Property

Public Property Set CurrentWorkbook(ByVal currwb As Workbook)
    Set this.currwb = currwb
End Property

Private Sub InitializeFileElements()
  this.llFilePath = "__notset"
  this.geoFilePath = "__notset"
  this.geohistFilePath = "__notset"

  Set RESERVEDSHEETS = New BetterArray
  RESERVEDSHEETS.Push DICTIONARYSHEET, DICTTRANSLATIONSHEET, CHOICESSHEET, METADATASHEET

  Set IMPORTEDVARIABLES = New BetterArray
  Set IMPORTEDSHEETS = New BetterArray
End Sub

'Speed Up the application
Private Sub BusyApp(Optional ByVal cursor As Long = xlNorthwestArrow)
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Application.Calculation = xlCalculationManual
  Application.EnableAnimations = False
  Application.EnableEvents = False
  Application.Cursor = cursor
End Sub

Private Sub NotBusyApp()
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Application.EnableAnimations = True
  Application.EnableEvents = True
  Application.Cursor = xlDefault
End Sub

'Tell if the current import Need a report
Private Property Get RequireReport() As Boolean
  RequireReport = this.needReport
End Property

'Set need Report to true
Private Sub SetNeedReportToTrue()
  If Not this.needReport Then this.needReport = True 
End Sub

'Get the file path. Scope can take up to 3 different values
'linelist -- for the linelist import path
'geobase -- for the geobase file path
'geobase-historic -- for the geobase historic file path

Private Property Get ImportFilePath(Optional Byval scope As String)

  Dim filePath As String 'will save the file attribute in one variable
  Dim io As IOSFiles

  filePath = Switch( _
    scope = "linelist", this.llFilePath, _ 
    scope = "geobase", this.geoFilePath, _ 
    scope = "geobase-historic", this.geohistFilePath)

  '__notset is a tag to inform about the path not initialized.
  'The path name is initialized once.
  If filePath = "__notset" Then

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    
    If io.HasValidFile Then 
    
      filePath = io.File()
      
      Select Case scope
      Case "linelist"
        this.llFilePath = filePath
      Case "geobase"
        this.geoFilePath = filePath
      Case "geobase-historic"
        this.geohistFilePath = filePath
      End Select

    End If

  End If

  ImportFilePath = filePath
End Property

Private Property Get ReportSheet() As Worksheet
  Dim wb As Workbook

  Set wb = CurrentWorkbook()
  Set ReportSheet = wb.Worksheets(IMPREPORTSHEET)
End Property

Private Property Get Dictionary() As ILLdictionary
  Dim wb As Workbook

  Set wb = CurrentWorkbook()
  Set Dictionary = LLdictionary.Create(wb.Worksheets(DICTIONARYSHEET), 1, 1)
End Property

Private Property Get ReportObject() As IImpRep
  Dim sh As Worksheet
  Dim fobj As Object

  Set fobj = ReportFormObject()
  Set sh = ReportSheet()
  
  Set ReportObject = ImpRep.Create(fobj, sh)
End Property

Private Property Get PasswordsObject() As ILLPasswords
  Dim wb As Workbook

  Set wb = CurrentWorkbook()
  Set PasswordsObject = LLPasswords.Create(wb.Worksheets(PASSWORDSHEET)) 
End Property

'Geo Object for geobase manibulations
Private Property Get GeoObject() As ILLGeo
  Dim currwb As Workbook
  Set currwb = CurrentWorkbook()
  Set GeoObject = LLGeo.Create(currwb.Worksheets(GEOSHEET))
End Property

'Get the linelist Translation object, with all elements related
'to translations at the linelist level
Private Property Get LLTradsObject() As ILLTranslations
  Dim currwb As Workbook

  Set currwb = CurrentWorkbook()
  Set LLTradsObject() = LLTranslations.Create(currwb.Worksheets(MESSTRANSLATIONSHEET), _
                                     currwb.Worksheets(DICTTRANSLATIONSHEET))
End Property

'Get the translation of messages object
Private Property Get MessagesTrads() As ITranslation
  Dim lltrads As ILLTranslations
  Set lltrads = LLTradsObject()
  Set MessagesTrads = lltrads.TransObject()
End Property

'Get the update object
Private Property Get UpdateObject() As IUpVal()
  Dim currwb As Workbook
  
  Set currwb = CurrentWorkbook()
  Set UpdateObject = UpVal.Create(currwb.Worksheets(UPDATESHEET))
End Property


'Add sheets not touched to import report
Private Sub AddSheetsNotTouchedToReport()
  
  Dim dict As ILLdictionary
  Dim sheetsList As BetterArray
  Dim counter As Long
  Dim repObj As IImpRep
  Dim sheetName As String

  Set dict = Dictionary()
  Set sheetsList = dict.UniqueValues("sheet name")
  Set repObj = ReportObject()

  For counter = sheetsList.LowerBound to sheetsList.UpperBound
    sheetName = sheetsList.Item(counter)
    If Not IMPORTEDSHEETS.Includes(sheetName) Then 
      SetNeedReportToTrue
      repObj.AddSheetsToReport 2, sheetName
    End If
  Next
End Sub


'Add variables not touched to import report
Private Sub AddVariablesNotTouchedToReport()

  Dim dict As ILLdictionary
  Dim repObj As IImpRep
  Dim varName As String
  Dim counter As Long 
  Dim varCounter As Long
  Dim sheetName As String
  Dim vars As ILLVariables  
  Dim currSheetsList As BetterArray 'List of sheet names in current workbook
  Dim currVarsList As BetterArray 'For touched sheets, list of variables
  Dim varControl As String

  Set repObj = ReportObject()
  Set dict = Dictionary()
  Set vars = LLVariables.Create(dict)
  Set currVarsList = New BetterArray
  Set currSheetsList = New BetterArray

  'Get the list of sheet names in the current worksheet
  Set currSheetsList = dict.UniqueValues("sheet name")

  For counter = currSheetsList.LowerBound To currSheetsList.UpperBound
  
    sheetName = currVarsList.Item(counter)
    'Test variables of only sheets that have been touched.
    If Not repObj.ContainsSheet(2, sheetName) Then

      Set currVarsList = dict.Data.FilterData("sheet name", sheetName, "variable name")
      
      'For each variable in the list of the variable of one sheet, test if
      'the variable has been imported. If not, report it.
      For varCounter = currVarsList.LowerBound To currVarsList.UpperBound
        varName = currVarsList.Item(varCounter)
        varControl = vars.Value(colName:="control", varName:=varName)
        
        'report only variables that are not formulas, formulas could be ignored
        'because they are recomputed
        If (varControl <> "case_when") And (varControl <> "formula") And _ 
           (varControl <> "choice_formula") Then

          If Not IMPORTEDVARIABLES.Includes(varName) Then 
            SetNeedReportToTrue
            repObj.AddVariablesToReport 2, sheetName, varName
          End If

        End If
      Next

    End If
  Next
End Sub


'Import data for migration from another workbook
Private Sub ImportData(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)

  Dim currwb As Workbook
  Dim impsh As worksheet
  Dim currsh As Worksheet
  Dim dict As ILLdictionary
  Dim pass As ILLPasswords
  Dim currTab As ICustomTable
  Dim impTab As IDataSheet
  Dim impHeadersList As BetterArray
  Dim varsNotImpList As BetterArray
  Dim usedImpRng As Range
  Dim counter As Long
  Dim varName As String
  Dim vlistRng As Range
  Dim repObj As IImpRep
  Dim nbBlank As Long
  Dim sheetsList As BetterArray

  Set dict = Dictionary()
  Set sheetsList = dict.UniqueValues("sheet name")
  Set currwb = CurrentWorkbook()
  Set pass = PasswordsObject()
  Set impHeadersList = New BetterArray 'on HList sheet, this will be the list of variables
  Set repObj = ReportObject()
  Set varsNotImpList = New BetterArray


  For Each impsh in impwb.Worksheets
    
    'Save imported worksheet name
    If Not RESERVEDSHEETS.Includes(impsh.Name) Then IMPORTEDSHEETS.Push impsh.Name
    
    If sheetsList.Includes(impsh.Name) Then
      Set currsh = currwb.Worksheets(impsh.Name)

      Select Case currsh.Cells(1, 3).Value
      'On Vlist Sheets
      Case "VList"
        pass.UnProtect currsh.Name
        Set usedImpRng = impsh.UsedRange

        For counter = 2 To usedImpRng.Rows.Count
          
          Set vlistRng = Nothing
          'The first column is for variable name
          varName = usedImpRng.Cells(counter, 1).Value
          
          'Test to see if the range is in the current VList sheet
          On Error Resume Next
          Set vlistRng = currsh.Range(varName)
          On Error GoTo 0
          
          'If the range is Not found, then you need to report
          If vlistRng Is Nothing Then
            repObj.AddVariablesToReport 1, impsh.Name, varName
            SetNeedReportToTrue
          Else
            vlistRng.Value = usedImpRng.Cells(counter, 2).Value
          End If
          
          IMPORTEDVARIABLES.Push varName
        Next

        pass.Protect currsh.Name
      'On HList sheets, use custom table to import data from datasheet

      Case "HList"
        'Number of blanks to resize the custom table after import
        nbBlank = currsh.Cells(1, 6).Value
        pass.UnProtect currsh.Name
        Set currTab = CustomTable.Create(currsh.ListObjects(1))
        'When instanciating the datasheet, 
        'the endrow here is strictly True, so this will be the last Row
        'of the imported worksheet. The function will find the last used rows
        'and this will be the last non empty row of the worksheet.
        Set impTab = DataSheet.Create(impsh, 1, 1, True)
        currTab.Import impTab, pasteAtBottom

        'Test if variables are not imported, add them to the report
        If currTab.HasColumnsNotImported() Then
          Set varsNotImpList = currTab.ImportColumnsNotFound()
          'variables in the imported sheet not imported
          repObj.AddVariablesListToReport 1, currsh.Name, varsNotImpList
          SetNeedReportToTrue
        End If

        'Resize the listObject
        currTab.RemoveRows totalCount:=nbBlank
        
        'Add the list of variables from the imported Datasheet
        impHeadersList.Clear
        impHeadersList.FromExcelRange impTab.HeaderRange()
        'Concat is for pasting two tables:
        'paste the list of variables on a Hslist with the previous list of variables
        IMPORTEDVARIABLES.Concat impHeadersList
        
        pass.Protect currsh.Name
      End Select

    Else
      'If the sheet is not found, ask for report
      repObj.AddSheetsToReport 1, impsh.Name
      SetNeedReportToTrue
    End If
    
  Next
End Sub


'Clear data in the workbook
Private Sub ClearData()

  Dim sh As Worksheet
  Dim sheetTag As String
  Dim pass As ILLPasswords
  Dim currwb As Workbook
  Dim Lo As ListObject
  Dim actsh As Worksheet 'current active sheet

  Set pass = PasswordsObject()
  Set currwb = CurrentWorkbook()

  'Deletion process can change the active sheet
  'so save it and return it back after completion.
  Set actsh = ActiveSheet

  For Each sh in currwb.Worksheets
    Select Case sh.Cells(1, 3).Value

    Case "HList"

      pass.UnProtect sh.Name
      Set Lo = sh.ListObjects(1)
      If Not Lo.DataBodyRange Is Nothing Then Lo.DataBodyRange.Delete
      pass.Protect sh.Name

    Case "VList"
      
      pass.UnProtect sh.Name
      'The fourth column is where we write infos on vlist
      sh.Columns(4).ClearContents
      pass.Protect sh.Name

    End Select
  Next

  actsh.Activate
End Sub


Private Sub ControlClearData(Optional ByVal forImports As Boolean = False)

  Dim proceed As Byte
  Dim inputwbName As String
  Dim currwb As Workbook
  Dim goodName As Boolean
  Dim trads As ITranslation

  On Error GoTo ErrManage

  Set currwb = CurrentWorkbook()
  Set trads = MessagesTrads()

  'There are two different questions: one for import another to just delete the data.

  proceed = vbYes
  
  If Not forImports Then
    proceed = MsgBox(trads.TranslatedValue("MSG_DeleteAllData"), _ 
                       vbExclamation + vbYesNo, _ 
                        trads.TranslatedValue("MSG_Delete"))

  End If

  If proceed = vbYes Then    
    goodName = False

    'ask the user to correct the worksheet name until the name
    'is good. The user can choose to leave
    'If he doesn't want to delete the whole workbook's data.

    'The data is deleted upon validation of the workbook name
    Do While Not goodName

      'Ask the user about the current workbook name for validation before deleting
      inputwbName = InputBox(trads.TranslatedValue("MSG_LLName"), _ 
                        trads.TranslatedValue("MSG_Delete"), _ 
                        trads.TranslatedValue("MSG_EnterWkbName"))

      If inputwbName = Replace(currwb.Name, ".xlsb", vbNullString) Then

          'Proceed only if the user is able to provide the name of the actual workbook, we can delete
          goodName = True
          ClearData

      Else
        'Bad workbook name: Ask again if the user is sure about deleting  the data before leaving
        'In case the user is not sure, exit the do...while loop
        If MsgBox(trads.TranslatedValue("MSG_BadLLNameQ"), _ 
                  vbExclamation + vbYesNo, _ 
                  trads.TranslatedValue("MSG_Delete")) = vbNo _ 
        Then Exit Do             
      End If
    Loop

  Else 'The user wishes to cancel deletion process
      MsgBox trads.TranslatedValue("MSG_DelCancel"), _ 
            vbOKOnly, trads.TranslatedValue("MSG_Delete")
  End If

  Exit Sub

ErrManage:
  On Error Resume Next
  MsgBox trads.TranslatedValue("MSG_ErrClearData")
  On Error GoTo 0
End Sub


'Test if there is data in the worksheet
Private Property Get HasData() As Boolean

  Dim testData As Boolean
  Dim currwb As Workbook
  Dim sh As Worksheet
  Dim nbBlanks As Long
  Dim Lo As ListObject
  Dim counter As Long

  Dim proceed As Byte
  Dim inputwbName As String
  Dim NotGood As Boolean

  testData = False

  Set currwb = CurrentWorkbook()

  'Loop throughtout all the worksheets
  For Each sh In currwb.Worksheets

    'On HList sheet, loop through each row of 
    'the listObject
    If sh.Cells(1, 3).Value = "HList" Then
      nbBlanks = sh.Cells(1, 6).Value
      Set Lo = sh.ListObjects(1)
      If Not (Lo.DataBodyRange Is Nothing) Then
        counter = 1
        Do While (Not testData) And (counter <= Lo.DataBodyRange.Rows.Count)
          'Once you find a row with data in it, just exit
          If (Application.WorksheetFunction.COUNTA(Lo.ListRows(counter).Range) > nbBlanks) Then
            testData = True
            Exit For
          End If
          counter = counter + 1
        Loop
      End If  
    End If
  Next

  HasData = testData
End Property

Private Sub ControlHasData()

  Dim trads As ITranslation
  Dim testData As Boolean 'Test if the workbook contains data
  Dim proceed As Byte
  Dim continueAsking As Boolean
  
  Set trads = MessagesTrads()

  testData = HasData()
  'While the workbook contains data, ask if the user wants to delete
  'in that case delete the data and test if there are data left.
  'If the user doesn't want to delete the data, exit.
  Do While (testData) 
    proceed = MsgBox(trads.TranslatedValue("MSG_DeleteForImport"), _ 
                      vbExclamation + vbYesNo, _ 
                      trads.TranslatedValue("MSG_Delete"))
      
    If proceed = vbYes Then 
      ControlClearData forImports:=True
      testData = HasData()
    Else
      Exit Do
    End If
  Loop

  ControlHasData = testData
End Sub


Private Function ControlLanguage(ByVal impwb As Workbook) As Boolean

  'metadata sheet is where you store informations about
  'language in the linelist
  Const METADATASHEET As String = "Metadata"

  Dim metaTab As IDataSheet
  Dim lltrads As ILLTranslations
  Dim varRng As Range
  Dim trads As ITranslation
  Dim abortImport As Boolean
  Dim cellRng As Range
  'current and imported langauges
  Dim impLang As String
  Dim currLang As String

  Set trads = MessagesTrads()
  Set lltrads = LLTradsObject()

  abortImport = False

  If Not sheetExists(impwb, METADATASHEET) Then

    'If the metadata sheet is not found, ask the user if
    'he really wants to proceed to import. In that case
    'abort the imports.
    abortImport = (MsgBox(trads.TranslatedValue("MSG_NoMetadata"), _
                  vbExclamation + vbYesNo, _
                  trads.TranslatedValue("MSG_Imports")) = vbYes)
  Else

    'The metadata sheet is found, I will just look for the language
    Set metaTab = DataSheet.Create(impwb.Worksheets(METADATASHEET), 1, 1)
    Set varRng = metaTab.DataRange("variable")
    'Find the cellRange with the language value
    impLang = vbNullString
    currLang = lltrads.Value("lllanguage")
    'If there are errors, default back to empty string as import language
    On Error Resume Next
      Set cellRng = varRng.Find(What:="language", lookAt:=xlWhole, MatchCase:=False)
      impLang = cellRng.Offset(,1).Value
    On Error GoTo 0

    'If the import language is different from the current language, as if the user
    'is sure to proceed
    If (impLang <> currLang) And (impLang <> vbNullString) Then
      abortImport =  (   MsgBox(trads.TranslatedValue("MSG_ActualLanguage") & " " & currLang & _
                            trads.TranslatedValue("MSG_ImportLanguage") & " " & impLang & _
                            trads.TranslatedValue("MSG_QuitImports"), vbExclamation + vbYesNo, _
                            trads.TranslatedValue("MSG_LanguageDifferent")) = vbYes)
    'If the import language is not found, tell it to the user                        
    ElseIf (impLang = vbNullString) Then
      abortImport = (MsgBox(trads.TranslatedValue("MSG_NoLanguage"), _
                     vbExclamation + vbYesNo, _
                    trads.TranslatedValue("MSG_Imports")) = vbYes)
    End If
  End If

  ControlLanguage = abortImport
End Function


Private Sub ImportMigration()

  Const RNGLISTAUTOUPDATED As String = "RNG_UpdateListAuto"

  Dim impwb As Workbook
  Dim inPath As String
  Dim pasteAtBottom As Boolean
  Dim actsh As Worksheet
  Dim trads As ITranslation
  Dim adfobj As Object
  Dim upobj As IUpVal
  Dim repObj As IImpRep
  
  BusyApp
  Set actsh = ActiveSheet

  'Error Management on Import
  On Error GoTo ErrImport

  'Get the import Path
  inPath = ImportFilePath("linelist")

  'input the import Path: check if the file is valid, and exit if not
  'the case. Leave the advanced form as is to give the opportunity to re-do imports
  If (inPath = vbNullString) Then 
    NotBusyApp
    Exit Sub
  End If

  Set trads = MessagesTrads()

  'Test if the linelist has Data
  'If the linelist has data in it, I will paste values at the end of tables
  pasteAtBottom = ControlHasData()
  
  'Load the workbook
  BusyApp
  Set impwb = Workbooks.Open(inPath)
  ActiveWindow.WindowState = xlMinimized

  'Test if the language of Imports is the same as language in the workbook 
  'The user chooses to abort the import because the language is not correct
  If ControlLanguage(impwb) Then GoTo EndImport

  BusyApp
  ImportData impwb, pasteAtBottom

  'Report Sheets or variables not found in importReport
  AddSheetsNotTouchedToReport
  AddVariablesNotTouchedToReport

  'advance frame object (could be hidden depending on what the user wants)
  Set adfobj = AdvanceFormObject()

  'If some variables were not imported or touched, then
  'ask if the user wants to see a report, in that case
  'show it. Otherwise if the user wants to close the advance form,
  'close it.

  If RequireReport() Then
    If MsgBox(trads.TranslatedValue("MSG_FinishImportRep"), _
              vbQuestion + vbYesNo, _
              trads.TranslatedValue("MSG_Imports")) = vbYes Then
      adfobj.Hide
      Set repObj = ReportObject()
      repObj.Show
    End If
  Else
    'The user can hide the import form here if there is no report to show
    If MsgBox(trads.TranslatedValue("MSG_FinishImport"), _ 
              vbQuestion + vbYesNo, _
              trads.TranslatedValue("MSG_Imports")) = vbYes Then
      adfobj.Hide
    End If
  End If

  Set upobj = UpdateObject()

  'Set all list auto to yes update
  upobj.SetValue RNGLISTAUTOUPDATED, "yes"
  actsh.Activate
  NotBusyApp
  Exit Sub

'Something ended the imports
EndImport:
  On Error Resume Next
  MsgBox trads.TranslatedValue("MSG_AbortImport"), vbOkOnly, trads.TranslatedValue("MSG_Imports")
  impwb.Close saveChanges:=False
  actsh.Activate
  NotBusyApp
  Exit Sub

'There are errors in the imports
ErrImport:
  On Error Resume Next
  MsgBox trads.TranslatedValue("MSG_ErrorImport"), _
         vbCritical + vbOKOnly, _
         trads.TranslatedValue("MSG_Imports")
  actsh.Activate
  NotBusyApp
  On Error GoTo 0
End Sub


'Import the Geobase (including historic or Not)
Private Sub ImportGeobase(Optional ByVal histoOnly As Boolean = False)

  Dim geo As ILLGeo
  Dim dict As ILLdictionary
  Dim pass As ILLPasswords
  Dim impwb As Workbook
  Dim showhidesh As Worksheet
  Dim filePath As String
  Dim pathTag As String
  Dim currwb As Workbook
  Dim trads As ITranslation
  Dim adfobj As Object

  pathTag = IIf(histoOnly, "geobase-historic", "geobase")
  Set trads = MessagesTrads()

    On Error GoTo ErrImpGeo

    'Get the geobase file Path
    filePath = ImportFilePath(pathTag)


    If filePath <> vbNullString Then
      BusyApp
      Set geo = GeoObject()
      Set currwb = CurrentWorkbook()

      'Open the geobase file
      Set impwb = Workbooks.Open(filePath)
      ActiveWindow.WindowState = xlMinimized

      'Import elements of the file
      geo.Import impwb, histoOnly

      'Update the geobase names in the workbook and the dictionary, and show/hide
      'If not on historic
      If Not histoOnly Then
        Set showhidesh = currwb.Worksheets(SHOWHIDESHEET)
        Set pass = PasswordsObject()
        Set dict = Dictionary()
        'Update geobase headers in each HList worksheet
        geo.Update pass
        'Update dictionary geo variables
        geo.UpdateDict dict
        'Update show hide geo variables
        geo.UpdateShowHide showhidesh, dict
      End If

      impwb.Close savechanges:=False

      'Leaving the import process
      If MsgBox( _
            trads.TranslatedValue("MSG_FinishImportGeo"), _
            vbQuestion + vbYesNo, trads.TranslatedValue("MSG_Imports") _ 
            ) = vbYes Then
        Set adfobj = AdvanceFormObject()
        adfobj.Hide
      End If

      NotBusyApp
    End If

    Exit Sub

ErrImpGeo:
  On Error Resume Next
  MsgBox trads.TranslatedValue("MSG_ErrImportGeo"), _ 
         vbCritical + vbOKOnly, _
        trads.TranslatedValue("MSG_Imports")
  NotBusyApp
  On Error GoTo 0
End Sub


'Errors checking at instanciation
Private Sub CheckRequirements(ByVal currwb As Workbook, ByVal repfobj As Object, ByVal adfobj As Object)

    Dim counter As Long
    Dim sheetName As String
    Dim sheetsArray As BetterArray

    Set sheetsArray = New BetterArray

    If repfobj Is Nothing Then ThrowErrorUnableInstanciate "Form for import report object not defined"
    If adfobj Is Nothing Then ThrowErrorUnableInstanciate "Advanced form object not defined"

    If TypeName(repfobj) <> "F_ImportRep" Then _ 
     ThrowErrorUnableInstanciate "The provided form is not F_ImportRep"

    If TypeName(adfobj) <> "F_Advanced" Then _ 
    ThrowErrorUnableInstanciate "The provided form for advanced is not F_Advanced"
  
    sheetsArray.Push DICTIONARYSHEET, CHOICESSHEET, _
                     PASSWORDSHEET, IMPREPORTSHEET, UPDATESHEET, _
                     MESSTRANSLATIONSHEET, _
                     GEOSHEET, SHOWHIDESHEET
  
    For counter = sheetsArray.LowerBound To sheetsArray.UpperBound
      sheetName = sheetsArray.Item(counter)
      If Not sheetExists(currwb, sheetName) Then _ 
       ThrowErrorUnableInstanciate " Worksheet " & sheetName & " Not found!"
    Next
  End Sub

  'Check if a sheet Exists in a workbook
  Private Function sheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    sheetExists = (Not sh Is Nothing)

  End Function

  
  'Errors managements
  Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, _
     "Unable to create an instance of the class " & CLASSNAME & ": " & errorMessage
  End Sub
  
  Private Sub ThrowError(ByVal errorMessage As String)
      Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
  End Sub

  Private Sub IImpSpecs_ImportGeobase(Optional ByVal histoOnly As Boolean = False)
    ImportGeobase histoOnly
  End Sub

  Private Sub IImpSpecs_ImportMigration()
    ImportMigration
  End Sub