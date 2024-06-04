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
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements IImpSpecs

Private Type TImp
  currwb As Workbook
  repfobj As Object
  adfobj As Object
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
Private Const IMPREPORTSHEET            As String = "import_rep__" 'temporary boggy sheet for report
Private Const SHOWHIDESHEET             As String = "show_hide__"
Private Const UPDATESHEET               As String = "updates__"

Private RESERVEDSHEETS                  As BetterArray
Private this                            As TImp
Private IMPORTEDVARIABLES               As BetterArray
Private IMPORTEDSHEETS                  As BetterArray
Private HASSAMELANGUAGE                 As Boolean

'Exposed methods
'repfobj is the form of the import Report
'adfobj is the advanced form Object
Public Function Create(ByVal repfobj As Object, ByVal adfobj As Object, _
                       ByVal currwb As Workbook) As IImpSpecs

  CheckRequirements currwb, repfobj, adfobj

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
  Set AdvanceFormObject = this.adfobj
End Property

Public Property Set AdvanceFormObject(ByVal fobj As Object)
  Set this.adfobj = fobj
End Property

Public Property Set CurrentWorkbook(ByVal currwb As Workbook)
    Set this.currwb = currwb
End Property

Private Sub InitializeFileElements()

  Set RESERVEDSHEETS = New BetterArray
  RESERVEDSHEETS.Push DICTIONARYSHEET, DICTTRANSLATIONSHEET, CHOICESSHEET, METADATASHEET, SHOWHIDESHEET

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
  Application.cursor = cursor
End Sub

Private Sub NotBusyApp()
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  Application.EnableAnimations = True
  Application.EnableEvents = True
  Application.cursor = xlDefault
End Sub

'Tell if the current import Need a report
Private Property Get RequireReport() As Boolean
  RequireReport = this.needReport
End Property

'Set need Report to true
Private Sub SetNeedReportToTrue()
  If Not this.needReport Then this.needReport = True
End Sub

'Get the file path.
'File path for either geo, linelist or geobase historic
Private Function ImportFilePath() As String
  Dim io As IOSFiles
  Set io = OSFiles.Create()
  io.LoadFile "*.xlsx"
  If io.HasValidFile Then ImportFilePath = io.File()
End Function

Private Property Get ReportSheet() As Worksheet
  Dim wb As Workbook

  Set wb = CurrentWorkbook()
  Set ReportSheet = wb.Worksheets(IMPREPORTSHEET)
End Property

Private Property Get Dictionary() As ILLdictionary
  Dim wb As Workbook
  Dim upobj As IUpVal
  Dim nbExp As Long

  Set wb = CurrentWorkbook()
  Set upObj = UpdateObject()
  On Error Resume Next
  nbExp = CLng(upObj.Value("RNG_NumberExports"))
  On Error GoTo 0

  Set Dictionary = LLdictionary.Create(wb.Worksheets(DICTIONARYSHEET), 1, 1, nbExp)
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
Private Property Get geoObject() As ILLGeo
  Dim currwb As Workbook
  Set currwb = CurrentWorkbook()
  Set geoObject = LLGeo.Create(currwb.Worksheets(GEOSHEET))
End Property

'Get the linelist Translation object, with all elements related
'to translations at the linelist level
Private Property Get LLTradsObject() As ILLTranslations
  Dim currwb As Workbook

  Set currwb = CurrentWorkbook()
  Set LLTradsObject = LLTranslations.Create(currwb.Worksheets(MESSTRANSLATIONSHEET), _
                                     currwb.Worksheets(DICTTRANSLATIONSHEET))
End Property

'Get the translation of messages object
Private Property Get MessagesTrads() As ITranslation
  Dim lltrads As ILLTranslations
  Set lltrads = LLTradsObject()
  Set MessagesTrads = lltrads.TransObject()
End Property

'Get the update object
Private Property Get UpdateObject() As IUpVal
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

  For counter = sheetsList.LowerBound To sheetsList.UpperBound
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
  
    sheetName = currSheetsList.Item(counter)
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
  Dim impsh As Worksheet
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
  Dim hasFormula As Boolean

  Set dict = Dictionary()
  Set sheetsList = dict.UniqueValues("sheet name")
  Set currwb = CurrentWorkbook()
  Set pass = PasswordsObject()
  Set impHeadersList = New BetterArray 'on HList sheet, this will be the list of variables
  Set repObj = ReportObject()
  Set varsNotImpList = New BetterArray

  'Initialize IMPORTED SHEETS/VARIABLES, RESERVED SHEETS
  InitializeFileElements
  'Clear the report object before adding new elements.
  repObj.Clear

  For Each impsh In impwb.Worksheets
    
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
          hasFormula = True
          'The first column is for variable name
          varName = usedImpRng.Cells(counter, 1).Value
          
          'Test to see if the range is in the current VList sheet
          On Error Resume Next
          Set vlistRng = currsh.Range(varName)
          hasFormula = (CStr(vlistRng.formula) <> CStr(vlistRng.Value))
          On Error GoTo 0
          
          'If the range is Not found, then you need to report
          If vlistRng Is Nothing Then
            repObj.AddVariablesToReport 1, impsh.Name, varName
            SetNeedReportToTrue
          Else
            If Not hasFormula Then _
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
        'The last True value is for strict search for column names.
        'Speed up without adding shift
        Set impTab = DataSheet.Create(impsh, 1, 1, True)
        currTab.Import impTab, pasteAtBottom, True, False

        'Test if variables are not imported, add them to the report
        If currTab.HasColumnsNotImported() Then
          Set varsNotImpList = currTab.ImportColumnsNotFound()
          'variables in the imported sheet not imported
          repObj.AddVariablesListToReport 1, currsh.Name, varsNotImpList
          SetNeedReportToTrue
        End If

        'Resize the listObject
        'currTab.RemoveRows totalCount:=nbBlank
        
        'Add the list of variables from the imported Datasheet
        impHeadersList.Clear
        impHeadersList.FromExcelRange impTab.HeaderRange()
        
        'Concat is for pasting two tables:
        'paste the list of variables on a Hslist with the previous list of variables
        IMPORTEDVARIABLES.CONCAT impHeadersList.Items
        
        pass.Protect currsh.Name
      End Select

      'Calculate the used Range on the imported sheet
      'Calculation could result in a lot of unexpected behaviours.
      On Error Resume Next
      currsh.UsedRange.calculate
      On Error GoTo 0

    ElseIf Not RESERVEDSHEETS.Includes(impsh.Name) Then
      'If the sheet is not found, ask for report
      repObj.AddSheetsToReport 1, impsh.Name
      SetNeedReportToTrue
    End If
    
  Next
End Sub


'Clear data in the workbook
Private Sub ClearData()

  Dim sh As Worksheet
  Dim pass As ILLPasswords
  Dim currwb As Workbook
  Dim Lo As listObject
  Dim actsh As Worksheet 'current active sheet
  Dim labelRng As Range
  Dim valueRng As Range
  Dim counter As Long

  Set pass = PasswordsObject()
  Set currwb = CurrentWorkbook()

  'Deletion process can change the active sheet
  'so save it and return it back after completion.
  Set actsh = ActiveSheet

  For Each sh In currwb.Worksheets
    Select Case sh.Cells(1, 3).Value

    Case "HList"

      pass.UnProtect sh.Name
      Set Lo = sh.ListObjects(1)
      'Deleting the databody range of a listObject removes all formulas
      'we need to make sure we could delete values in a range before proceeding
      If Not Lo.DataBodyRange Is Nothing Then

        For counter = 1 To Lo.DataBodyRange.Columns.Count
          Set valueRng = Lo.ListColumns(counter).DataBodyRange
          If Not valueRng.Cells(1, 1).HasFormula Then valueRng.ClearContents
        Next

        Lo.Range.Calculate
      End If

      pass.Protect sh.Name
      
    Case "VList"
      
      pass.UnProtect sh.Name
      'The fourth column is the column for labels in Vlist
      'We need to take care of formulas when deleting
      Set labelRng = sh.Cells(4, 4)
      Do While Not IsEmpty(labelRng)
        Set valueRng = labelRng.Offset(, 1)
        If Not valueRng.HasFormula Then valueRng.ClearContents
        Set labelRng = labelRng.Offset(1)
      Loop

      sh.UsedRange.Calculate
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

      If (inputwbName = vbNullString) And (Not forImports) Then
        Exit Do
      ElseIf inputwbName = Replace(currwb.Name, ".xlsb", vbNullString) Then
        
        'Proceed only if the user is able to provide the name of the actual
        'workbook, we can delete
        goodName = True
        ClearData

      Else
        'Bad workbook name: Ask again if the user is sure about deleting
        'the data before leaving
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
  Dim Lo As listObject
  Dim counter As Long

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

Private Function ControlHasData() As Boolean

  Dim trads As ITranslation
  Dim testData As Boolean 'Test if the workbook contains data
  Dim proceed As Byte
  
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
End Function


Private Function ControlLanguage(ByVal impwb As Workbook) As Boolean

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
    HASSAMELANGUAGE = False
    abortImport = (MsgBox(trads.TranslatedValue("MSG_NoMetadata"), _
                  vbExclamation + vbYesNo, _
                  trads.TranslatedValue("MSG_Imports")) = vbYes)
  Else

    'The metadata sheet is found, I will just look for the language
    Set metaTab = DataSheet.Create(impwb.Worksheets(METADATASHEET), 1, 1)
    Set varRng = metaTab.DataRange("variable")
    'Find the cellRange with the language value
    impLang = vbNullString
    currLang = lltrads.Value("dictionarylanguage")
    'If there are errors, default back to empty string as import language
    On Error Resume Next
      Set cellRng = varRng.Find(What:="language", lookAt:=xlWhole, MatchCase:=False)
      impLang = cellRng.Offset(, 1).Value
    On Error GoTo 0

    'If the import language is different from the current language, as if the user
    'is sure to proceed
    If (impLang <> currLang) And (impLang <> vbNullString) Then
      abortImport = (MsgBox(trads.TranslatedValue("MSG_ActualLanguage") & " " & currLang & _
                            ". " & trads.TranslatedValue("MSG_ImportLanguage") & " " & impLang & _
                            ". " & trads.TranslatedValue("MSG_QuitImports"), vbExclamation + vbYesNo, _
                            trads.TranslatedValue("MSG_LanguageDifferent")) = vbNo)
    HASSAMELANGUAGE = False
    'If the import language is not found, tell it to the user
    ElseIf (impLang = vbNullString) Then
      abortImport = (MsgBox(trads.TranslatedValue("MSG_NoLanguage"), _
                     vbExclamation + vbYesNo, _
                    trads.TranslatedValue("MSG_Imports")) = vbNo)
      HASSAMELANGUAGE = False
    ElseIf (impLang = currLang) Then
      HASSAMELANGUAGE = True
    End If
  End If

  ControlLanguage = abortImport
End Function

'Import show/hide on import
'It is crucial that imports has the same language as the linelist. It it is not the case
'safely avoid importing show/hide objects
Private Sub ImportShowHide(ByVal inpwb As Workbook)

  Const SHOWHIDEPREF As String = "ShowHideTable_"

  Dim metaTab As IDataSheet
  Dim varRng As Range
  Dim tradsmess As ITranslation
  Dim cellRng As Range 'used for eventual stuffs
  Dim updateShowHide As Boolean
  Dim currwb As Workbook
  Dim sh As Worksheet
  Dim currDict As ILLdictionary 'current Dictionary in the current workbook
  Dim showHideObject As ILLShowHide 'Show/Hide Object
  Dim currLo As ListObject
  'Custom tables for import and current worksheet
  Dim inpCsTab As ICustomTable
  Dim currCsTab As ICustomTable
  'Input and current show/hide
  Dim inpsh As Worksheet
  Dim currsh As Worksheet
  Dim varName As String
  Dim sheetName As String
  Dim colIndex As Long
  Dim colWidth As Long
  Dim varStatus As String
  Dim inpDict As ILLDictionary 'Imported dictionary
  Dim inpVarRng As Range
  Dim currVarRng As Range
  Dim currVarObj As ILLVariables
  Dim counter As Long
  Dim tabName As String
  Dim sheetTag As String
  
  'If you import with different language, do not even attempt to update show/hide
  If Not HASSAMELANGUAGE Then Exit Sub
  'If there is no metadata, exit the sub (should have a warning or a report)
  If Not sheetExists(inpwb, METADATASHEET) Then Exit Sub
  'if There is no show/hide worksheet in the imported workbook, exits
  If Not sheetExists(inpwb, SHOWHIDESHEET) Then Exit Sub
  If Not sheetExists(inpwb, DICTIONARYSHEET) Then Exit Sub
  
  'The imported dictionary should be prepared
  Set inpDict = LLdictionary.Create(inpwb.Worksheets(DICTIONARYSHEET), 1, 1)
  If Not inpDict.Prepared Then Exit Sub

  'The metadata sheet is found, I will just look for the import show hide tag
  Set metaTab = DataSheet.Create(inpwb.Worksheets(METADATASHEET), 1, 1)
  Set varRng = metaTab.DataRange("variable")
  updateShowHide = False
  On Error Resume Next
      Set cellRng = varRng.Find(What:="show_hide", lookAt:=xlWhole, MatchCase:=False)
      updateShowHide = (cellRng.Offset(, 1).Value = "update on import")
  On Error GoTo 0

  'If there is no need to update, because update show/hide != update on import: Exit
  If Not updateShowHide Then Exit Sub

  'Step1: Initialize all the show/hide tables of the workbbok
  Set currwb = CurrentWorkbook()

  'Translation message
  Set tradsmess = MessagesTrads()
  Set currDict = Dictionary()
  'Initialize show/hide tables
  For Each sh In currwb.Worksheets
    If (sh.Cells(1, 3).Value = "HList") Or _ 
       (sh.Cells(1, 3).Value = "HList Print") Or _ 
       (sh.Cells(1, 3).Value = "VList") Then
      Set showHideObject = LLShowHide.Create(tradsmess, currDict, sh)
      showHideObject.InitializeTable
    End If
  Next

  'Step2: For each ListObjects of the current show/hide worksheet, initialize values of the
  Set currsh = currwb.Worksheets(SHOWHIDESHEET) 'show hide worksheet in the current workbook
  Set inpsh = inpwb.Worksheets(SHOWHIDESHEET) 'show hide worksheet in the imported workbook
  Set currVarObj = LLVariables.Create(currDict)

  For Each currLo in currsh.ListObjects

    If LoExists(inpsh, currLo.Name) Then
      Set inpCsTab = CustomTable.Create(inpsh.ListObjects(currLo.Name), idCol:="variable name")
      Set currCsTab = CustomTable.Create(currLo, idCol:="variable name")

      Set inpVarRng = inpCsTab.DataRange("variable name", strictSearch:=True)
      Set currVarRng = currCsTab.DataRange("variable name", strictSearch:=True)
     
      For counter = 2 To inpVarRng.Rows.Count
        varName = inpVarRng.Cells(counter, 1).Value

        'Test if variable is present in the current variable name range 
        If Not (currVarRng.Find(What:=varName, LookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
          
          'Change the status
          varStatus = inpVarRng.Cells(counter, 2).Value
          currCsTab.SetValue keyName:=varName, colName:="status", newValue:=varStatus
          
          'Change the column width
          colWidth = 22
          On Error Resume Next
            colWidth = CLng(inpVarRng.Cells(counter, 3).Value)
          On Error GoTo 0

          'Change the column width of the corresponding variable
          colIndex = 0
          On Error Resume Next
          colIndex = CLng(currVarObj.Value(colName:="column index", varName:=varName))
          On Error GoTo 0

          'For hidden variables, the column width is 0, so they will be hidden
          If colIndex <> 0 Then
            
            sheetName = currVarObj.Value(colName:="sheet name", varName:=varName)
            tabName = Replace(currLo.Name, SHOWHIDEPREF, vbNullString)
            
            'could be on printed or data worksheet
            If (InStr(1, tabName, "pr") > 0 ) Then sheetName = "print_" & sheetName
            Set sh = currwb.Worksheets(sheetName)
            sheetTag = sh.Cells(1, 3).Value

            On Error Resume Next
            If (sheetTag = "VList") Then
              sh.Rows(colIndex).RowHeight = colWidth
            Else
              sh.Columns(colIndex).ColumnWidth = colWidth
            End If
            On Error GoTo 0

            'On printed worksheet, if orientation is vertical, rotate the header
            If (varStatus = tradsmess.TranslatedValue("MSG_ShowVerti")) Then
              On Error Resume Next
              Set cellRng = sh.Range(Replace(tabName, "pr", vbNullString) & "_" & "PRINTSTART")
              Set cellRng = sh.Cells(cellRng.Offset(-2).Row, colIndex)
              cellRng.EntireRow.RowHeight = 100
              cellRng.Orientation = 90
              On Error GoTo 0
            End If
          End If
        End If
      Next
    End If
  Next
End Sub

'Update Editable Labels on Import
Private Sub UpdateEditableLabels(ByVal inpwb As Workbook)

  Dim metaTab As IDataSheet
  Dim varRng As Range
  Dim cellRng As Range 'used for eventual stuffs
  Dim updateLabels As Boolean
  Dim currwb As Workbook
  Dim sheetTag As String
  Dim inpDict As ILLdictionary
  Dim currDict As ILLdictionary
  Dim currVars As ILLVariables
  Dim inpVars As ILLVariables
  Dim counter As Long
  Dim varName As String
  Dim newLabel As String
  Dim sheetName As String
  Dim sheetInfo As String
  Dim sh As Worksheet
  Dim colIndex As Long

  
  'If you import with different language, do not even attempt to update show/hide
  If Not HASSAMELANGUAGE Then Exit Sub
  'If there is no metadata, exit the sub (should have a warning or a report)
  If Not sheetExists(inpwb, METADATASHEET) Then Exit Sub
  'if There is no dictionary worksheet in the imported workbook, exits
  If Not sheetExists(inpwb, DICTIONARYSHEET) Then Exit Sub
  'The imported dictionary should be prepared, otherwise exit
  Set inpDict = LLdictionary.Create(inpwb.Worksheets(DICTIONARYSHEET), 1, 1)
  If Not inpDict.Prepared Then Exit Sub

  'The metadata sheet is found, I will just look for the import show hide tag
  Set metaTab = DataSheet.Create(inpwb.Worksheets(METADATASHEET), 1, 1)
  Set varRng = metaTab.DataRange("variable")
  updateLabels = False
  On Error Resume Next
      Set cellRng = varRng.Find(What:="editable_labels", lookAt:=xlWhole, MatchCase:=False)
      updateLabels = (cellRng.Offset(, 1).Value = "update on import")
  On Error GoTo 0

  'If there is no need to update, because update show/hide != update on import: Exit
  If Not updateLabels Then Exit Sub
  Set currwb = CurrentWorkbook()
  Set currDict = Dictionary()
  Set currVars = LLVariables.Create(currDict)
  Set inpVars = LLVariables.Create(inpDict)
  Set varRng = currDict.DataRange("variable name")

  For counter = 1 To varRng.Rows.Count
    
    varName = varRng.Cells(counter, 1).Value
    If Not inpVars.Contains(varName) Then GoTo ContinueLoop
    'If not editable label, change the values
    If (currVars.Value(varName:=varName, colName:="editable label") <> "yes") Then GoTo ContinueLoop

    'Change the corresponding label in the worksheet ==============================================
    newLabel = inpVars.Value(varName:=varName, colName:="main label")
    'Retrieve the column index of the variable
    colIndex = 0
    On Error Resume Next
    colIndex = CLng(currVars.Value(varName:=varName, colName:="column index"))
    On Error GoTo 0
    'On incorrect column indexes, skip
    If colIndex = 0 Then GoTo ContinueLoop
    'On correct column indexes, change the label
    'Change the label in the worksheet (the cell is not protected)
    sheetName = currVars.Value(varName:=varName, colName:="sheet name")
    sheetInfo = currVars.Value(varName:=varName, colName:="sheet type")
    Set sh = currwb.Worksheets(sheetName)

    'update Label (in case it seems impossible, skip and continue)
    On Error GoTo ContinueLoop
    Set cellRng = sh.Range(sh.Cells(1, 4).Value & "_START")
    If (sheetInfo = "hlist2D") Then  
      Set cellRng = sh.Cells(cellRng.Offset(-1).Row, colIndex)
      cellRng.Value = newLabel
    ElseIf (sheetInfo = "vlist1D") Then
      Set cellRng = sh.Cells(colIndex, cellRng.Offset( ,-1).Column)
      cellRng.Value = newLabel
    End If
    On Error GoTo 0

    'Change the label in the dictionary ======================================================
    '(only if you are able to change the label on the worksheet)
    currVars.SetValue varName:=varName, colName:="main label", newValue:=newLabel

  ContinueLoop:
  Next
End Sub

'import custom dropdown
Private Sub ImportCustomDropdown(ByVal inpwb As Workbook, ByVal pasteAtBottom As Boolean)

  Dim currwb As Workbook
  Dim lltrads AS ILLTranslations
  Dim drop As IDropdownLists 'Dropdown on custom choices
  Dim Lo As ListObject
  Dim listName As String
  Dim choiceName As String
  Dim dropsh As Worksheet 'Dropdown worksheet
  Dim choiobj As ILLchoice 'Choices worksheet in the imported workbook
  Dim updateData As BetterArray


  If Not sheetExists(inpwb, "Choices") Then Exit Sub

  Set lltrads = LLTradsObject()
  Set currwb = CurrentWorkbook()
  Set choiobj = LLchoice.Create(inpwb.Worksheets("Choices"), 1, 1)
  Set dropsh = currwb.Worksheets(lltrads.Value("customdropdown"))
  'There is no prefix for dropdowns on customdropdown
  Set drop = DropdownLists.Create(dropsh, vbNullString)
  Set updateData = New BetterArray

  For Each Lo in dropsh.ListObjects
    
    updateData.Clear
    listName = lo.Name
    listName = Right(listName, (Len(listName) - 5))
    choiceName = "__choice_custom_" & listName

    On Error Resume Next
      Set updateData = choiobj.Categories(choiceName)
    On Error GoTo 0

    'Update the dropdown only if it is necessary
    If (updateData.Length > 0) Then drop.Update updateData, listName, pasteAtBottom
  Next
End Sub


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
  inPath = ImportFilePath()

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

  'Import all data from HList and VList worksheets
  ImportData impwb, pasteAtBottom

  'Import the custom dropdowns
  ImportCustomDropdown impwb, pasteAtBottom

  'Report Sheets or variables not found in importReport
  AddSheetsNotTouchedToReport
  AddVariablesNotTouchedToReport

  'Import show/hide elements
  ImportShowHide impwb

  'Update the values of editable labels
  UpdateEditableLabels impwb

  'Set all list auto to yes update
  impwb.Close savechanges:=False
  BusyApp
  actsh.Activate
  BusyApp

  'Reset updated values
  Set upobj = UpdateObject()
  upobj.SetValue RNGLISTAUTOUPDATED, "yes"

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
    MsgBox trads.TranslatedValue("MSG_FinishImport"), _
              vbOkOnly,  trads.TranslatedValue("MSG_Imports")
  End If

  NotBusyApp
  Exit Sub

'Something ended the imports
EndImport:
  On Error Resume Next
  MsgBox trads.TranslatedValue("MSG_AbortImport"), _
         vbOKOnly, trads.TranslatedValue("MSG_Imports")
  impwb.Close savechanges:=False
  actsh.Activate
  NotBusyApp
  On Error GoTo 0
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
  Dim currwb As Workbook
  Dim trads As ITranslation
  Dim adfobj As Object
  
  Set trads = MessagesTrads()

  On Error GoTo ErrImpGeo

    'Get the geobase file Path
  filePath = ImportFilePath()
  If filePath <> vbNullString Then

    BusyApp

    Set geo = geoObject()
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
    MsgBox trads.TranslatedValue("MSG_FinishImportGeo"), _
            vbOkOnly, trads.TranslatedValue("MSG_Imports")
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

  'Check if a listObject Exists in a workbook
  Private Function LoExists(ByVal sh As Worksheet, ByVal loName As String) As Boolean
    Dim Lo As ListObject
    On Error Resume Next
    Set Lo = sh.ListObjects(loName)
    On Error GoTo 0
    LoExists = (Not Lo Is Nothing)
  End Function

  
  'Errors managements
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, _
     "Unable to create an instance of the class " & CLASSNAME & ": " & errorMessage
End Sub

Private Sub IImpSpecs_ImportGeobase(Optional ByVal histoOnly As Boolean = False)
  ImportGeobase histoOnly
End Sub

Private Sub IImpSpecs_ImportMigration()
  ImportMigration
End Sub

Private Sub IImpSpecs_ClearData()
  BusyApp
  ControlClearData
  NotBusyApp
End Sub

Private Sub IImpSpecs_ShowReport()
  Dim repObj As IImpRep
  BusyApp
  Set repObj = ReportObject()
  repObj.Show
  NotBusyApp
End Sub
