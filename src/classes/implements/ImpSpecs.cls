VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "InpSpecs"
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
  fobj As Object
  llFilePath As String
  geoFilePath As String
  geohistFilePath As String
  needReport As Boolean
End Type

Private Const CLASSNAME                 As String = "ImpSpecs"
Private Const DICTIONARYSHEET           As String = "Dictionary"
Private Const CHOICESSHEET              As String = "Choices"
Private Const DICTTRANSLATIONSHEET      As String = "Translations"
Private Const METADATASHEET             As String = "Metadata"
Private Const PASSWORDSHEET             As String = "__pass" 'password sheets
Private Const REPORTSHEET               As String = "imp_rep__" 'temporary boggy sheet for report
Private Const UPDATESHEET               As String = ""

Private RESERVEDSHEETS                  As BetterArray
Private this                            As TImp
Private IMPORTEDVARIABLES               As BetterArray

'Exposed methods
'formObj is the form of the import Report
Public Function Create(ByVal formObj As Object, ByVal currwb As Workbook) AS IImpSpecs

  CheckRequirements currwb, fobj
  
  InitializeFileElements

  With New ImpSpecs
    Set .FormObject = formobj
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

Public Property Get FormObject() As Object
  Set FormObject = this.impForm
End Property

Public Property Set FormObject(ByVal fobj As Object)
  Set this.impForm = fobj
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
  Set ReportSheet = wb.Worksheets(REPORTSHEET)
End Property

Private Property Get Dictionary() As ILLdictionary
  Dim wb As Workbook

  Set wb = CurrentWorkbook()
  Set Dictionary = LLdictionary.Create(wb.Worksheets(DICTIONARYSHEET), 1, 1)
End Property

Private Property Get ReportObject() As IImpRep
  Dim sh As Worksheet
  Dim fobj As Object

  Set fobj = FormObject()
  Set sh = ReportSheet()
  
  Set ReportObject = ImpRep.Create(fobj, sh)
End Property

Private Property Get PasswordsObject() As ILLPasswords
  Dim wb As Workbook

  Set wb = CurrentWorkbook()
  Set PasswordsObject = LLPasswords.Create(wb.Worksheets(PASSWORDSHEET)) 
End Property

'Add sheets not imported/not touched to import report
Private Sub AddReportOnSheets(ByVal outwb As Workbook)
  
  Dim dict As ILLdictionary
  Dim sheetsList As BetterArray
  Dim unfoundSheetsList As BetterArray 
  Dim sh As Worksheet 'Used to iterate on worksheets in imported worbook
  Dim counter As Long
  Dim repObj As IImpRep

  Set dict = Dictionary()
  Set sheetsList = dict.UniqueValues("sheet name")
  Set unfoundSheetsList = New BetterArray
  Set repObj = ReportObject()

  'The scope for add sheets to report can take 2 values:
  '1- for sheets not imported from the import 
  '2- for sheets not touched in the current workbook

  'Sheets not imported
  For Each sh In outwb.Worksheets
    If (Not sheetsList.Includes(sh.Name)) Then
      unfoundSheetsList.Push sh.Name
    End If
  Next

  repObj.AddSheetsToReport(1, unfoundSheetsList)

  'Sheets in current workbook Not touched
  unfoundSheetsList.Clear
  For counter = sheetsList.LowerBound to sheetsList.UpperBound
    sheetName = sheetsList.Item(counter)
    If Not sheetExists(outwb, sheetName) Then unfoundSheetsList.Push sheetName
  Next

  repObj.AddSheetsToReport(2, unfoundSheetsList)
End Sub


'Add variables not imported/not touched to import report
Private Sub AddReportOnVariables(ByVal outwb As Workbook)

  Dim dict As ILLdictionary
  Dim sh As Worksheet
  Dim outData As IDatasheet
  Dim repObj As IImpRep
  Dim impData As IDataSheet 'Data sheet on imported worksheets
  Dim hRng As Range
  Dim vars As ILLVariables
  Dim varName As String
  Dim varsList As BetterArray
  Dim counter As Long 
  Dim varCounter As Long
  Dim sheetName As String  
  Dim currSheetsList As BetterArray 'List of sheet names in current workbook
  Dim currVarsList As BetterArray 'For touched sheets, list of variables

  Set repObj = ReportObject()
  Set dict = Dictionary()
  Set vars = LLvariables.Create(dict)
  Set varsList = New BetterArray
  Set currVarsList = New BetterArray
  Set currSheetsList = New BetterArray

  'Variables in imported workbook / sheets and not imported
  For Each sh In outwb.Worksheets

    varsList.Clear
    'If the sheet has been imported
    If (Not repObj.Contains(1, sh.Name)) And (Not RESERVEDSHEETS.Includes(sh.Name)) Then
      'Get the list of variables of the worksheet header
      Set impData = DataSheet.Create(sh, 1, 1, True)
      Set hRng = impData.HeaderRange()
      
      For counter = 1 To hRng.Columns.Count
        varName = hRng.Cells(counter, 1).Value
        'If the variable is not imported
        If (Not vars.Contains(varName)) Then varsList.Push varName
      Next
    End If

    If varsList.Length > 0 Then repObj.AddVariablesToReport(1, sh.Name, varsList)

  Next

  'Variables in the actual workbook / sheets not touched (not found in imported sheet)

  'Get the list of sheet names not found in the current sheet
  Set currSheetsList = dict.UniqueValues("sheet name")

  For counter = currSheetsList.LowerBound To currSheetsList.UpperBound
      
    'Get all variables in the current sheet
    sheetName = currVarsList.Item(counter)

    'Test variables of only sheets that have been touched.
    If Not repObj.Contains(2, sheetName) Then

      varsList.Clear
      Set currVarsList = dict.Data.FilterData("sheet name", sheetName, "variable name")

      'For each variable in the list of the variable of one sheet, test if
      'the variable has been imported. If not, report it.
      For varCounter = currVarsList.LowerBound To currVarsList.UpperBound
        varName = currVarsList.Item(varCounter)
        If Not IMPORTEDVARIABLES.Includes(varName) Then varsList.Push varName
      Next

      'If the varsList is not empty, add them to the sheets.
      If varsList.Length > 0 Then repObj.AddVariablesToReport(2, sheetName, varsList)
    End If
  Next
End Sub


'Import data for migration from another workbook
Private Sub ImportData(ByVal impwb As Workbook)

  Dim currwb As Workbook
  Dim impsh As worksheet
  Dim currsh As Worksheet
  Dim dict As ILLdictionary
  Dim pass As ILLPasswords
  Dim currTab As ICustomTable
  Dim impTab As IDataSheet
  Dim impHeadersList As BetterArray
  Dim usedImpRng As Range
  Dim counter As Long
  Dim rngName As String
  Dim vlistRng As Range

  Set sheetsList = dict.UniqueValues("sheet name")
  Set currwb = CurrentWorkbook()
  Set pass = PasswordsObject()
  Set impHeadersList = New BetterArray


  For Each impsh in impwb.Worksheets

    If sheetsList.Includes(sh.Name) Then
      Set currsh = currwb.Worksheets(sh.Name)

      Select Case currsh.Cells(1, 3).Value
      'On Vlist Sheets
      Case "VList"
        pass.UnProtect currsh.Name
        Set usedImpRng = impsh.UsedRange

        For counter = 2 To usedImpRng.Rows.Count
          
          Set vlistRng = Nothing
          rngName = usedImpRng.Cells(counter, 1).Value
          
          
          'Test to see if the range is in the current VList sheet
          On Error Resume Next
          Set vlistRng = currsh.Range(rngName)
          On Error GoTo 0
          
          'If the range is Not found, then you need to report
          If vlistRng Is Nothing Then
            SetNeedReportToTrue
          Else
            vlistRng.Value = usedImpRng.Cells(counter, 2).Value
          End If
          
          IMPORTEDVARIABLES.Push rngName
        Next

        pass.Protect currsh.Name
      'On HList sheets, use custom table to import data from datasheet
      Case "HList"
        pass.UnProtect currsh.Name
        Set currTab = CustomTable.Create(currsh.ListObjects(1))
        'The endrow here is strictly True, so this will be the last Range
        'of the worksheet
        Set impTab = DataShet.Create(impsh, 1, 1, True)
        currTab.Import impTab
        'Add the list of variables from the imported Datasheet
        impHeadersList.Clear
        impHeadersList.FromExcelRange impTab.HeaderRange()
        'Concat is for pasting two tables
        IMPORTEDVARIABLES.Concat impHeadersList
        pass.Protect currsh.Name
      End Select
    Else
      'If the sheet is not found, ask for report
      SetNeedReportToTrue
    End If
    
  Next







End Sub











































  'Errors checking at instanciation
  Private Sub CheckRequirements(ByVal currwb As Workbook, ByVal fobj As Object)

    Dim counter As Long
    Dim sheetName As String
    Dim sheetsArray As BetterArray

    Set sheetsArray = New BetterArray

    If formObj Is Nothing Then ThrowErrorUnableInstanciate "Form object not defined"
    If TypeName(formObj) <> "F_ImportRep" Then _ 
     ThrowErrorUnableInstanciate "The provided form is not F_ImportRep"
  
    sheetsArray.Push DICTIONARYSHEET, CHOICESSHEET, _
                     PASSWORDSHEET, REPORTSHEET, UPDATESHEET
  
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
  
  'Speed up application
  Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
  End Sub
  
  
  'Errors managements
  Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, _
     "Unable to create an instance of the class " & CLASSNAME & ": " & errorMessage
  End Sub
  
  Private Sub ThrowError(ByVal errorMessage As String)
      Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
  End Sub