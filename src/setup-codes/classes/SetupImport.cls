VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Imports")
'@ModuleDescription("Import another setup to this new one")
'@IgnoreModule

Option Explicit

'Exposed methods
Implements ISetupImport

Private Const CLASSNAME As String = "SetupImport"
Private Const EXPORTSHEETNAME As String = "Exports"
Private Const DICTSHEETNAME As String = "Dictionary"

Private Type TSetupImport
  inPath As String
  progress As Object
  inWb As Workbook
End Type

Private this As TSetupImport

Public Function Create(ByVal importPath As String, ByVal progrObj As Object)

  'Error Management when creating the imports.
  If (progrObj Is Nothing) Then ThrowError "Unable to initialize the class, information Object not found"

  With New SetupImport
    .Path = importPath
    Set .ProgressObject = progrObj
    Set Create = .Self()
  End With

End Function

'Get the path of the import of the setup
Public Property Get Path() As String
  Path = this.inPath
End Property

Public Property Let Path(ByVal importPath As String)
  this.inPath = importPath
End Property

Public Property Get ProgressObject() As Object
  Set ProgressObject = this.progress
End Property

Public Property Set ProgressObject(ByVal progrObj As Object)
  Set this.progress = progrObj
End Property

Public Property Get Self() As ISetupImport
  Set Self = Me
End Property

Private Sub WriteInfo(ByVal charValue As String)
  Dim progrObj As Object
  Set progrObj = ProgressObject()
  progrObj.Caption = charValue
End Sub

Private Sub BusyApp()
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Application.EnableAnimations = False
  Application.Calculation = xlCalculationManual
End Sub

Private Property Get ImportWkb() As Workbook
  BusyApp
  If this.inWb Is Nothing Then Set this.inWb = Workbooks.Open(Path())
  Application.WindowState = xlMinimized
  BusyApp
  Set ImportWkb = this.inWb
End Property

'Write a progress when importing the data
Private Sub WriteProgress(ByVal numValue As Integer)

  Const NUMBEROFBARS As Integer = 30
  Dim CurrentStatus As Integer
  Dim progrObj As Object
  Set progrObj = ProgressObject()
  CurrentStatus = (NUMBEROFBARS) * Round(Application.WorksheetFunction.Min(numValue, 100) / 100, 1)
  progrObj.Caption = "[" & String(CurrentStatus, "|") & SPACE(NUMBEROFBARS - CurrentStatus) & "]" & _
                        " " & CInt(numValue) & "%"
End Sub

Private Sub PrepareImport(ByVal pass As IPasswords, ByVal sheetLists As BetterArray)

  'curr prefix is for current ie in the current workbook
  'imp is for imported, ie in the imported workbook

  Dim currExpObj As ILLExport 'current export object,
  Dim impExpObj As ILLExport
  Dim currDict As ILLdictionary 'current dictionary
  Dim currExpsh As Worksheet 'current Export sheet
  Dim impExpSh As Worksheet 
  Dim currwb As Workbook 'current workbook
  Dim impwb As Workbook
  Dim counter As Long
  Dim prevNbExp As Long
  Dim actNbExp As Long

  'Exit if the import does not includes Exports
  If Not sheetLists.Includes(EXPORTSHEETNAME) Then Exit Sub

  Set impwb = ImportWkb()
  Set currwb = ThisWorkbook

  'Preparing the export worksheet for conformity with the imported one
  If sheetExists(impwb, EXPORTSHEETNAME) And sheetExists(impwb, DICTSHEETNAME) Then

    Set currExpsh = currwb.Worksheets(EXPORTSHEETNAME)
    Set impExpSh = impwb.Worksheets(EXPORTSHEETNAME)

    Set currExpObj = LLExport.Create(currExpsh, 4, 1)
    Set impExpObj = LLExport.Create(impExpsh, 4, 1)


    If currExpObj.NumberOfExports() < impExpObj.NumberOfExports() Then

      'Add the exports in the current setup
      pass.UnProtect DICTSHEETNAME
      pass.UnProtect EXPORTSHEETNAME

      Set currDict = LLdictionary.Create(currwb.Worksheets(DICTSHEETNAME), 5, 1)

      Do While currExpObj.NumberOfExports() < impExpObj.NumberOfExports()

        prevNbExp = currExpObj.NumberOfExports()
        currExpObj.AddRows
        actNbExp =  currExpObj.NumberOfExports()

        For counter = (prevNbExp + 1) To actNbExp
          currDict.InsertColumn "Export " & counter, After:= "Export " & (counter - 1)
        Next

      Loop

      pass.Protect DICTSHEETNAME
      pass.Protect EXPORTSHEETNAME
    End If
    
  End If
End Sub


'Check if the setup file to import is a valid setup file
Private Sub check(ByVal importDict As Boolean, _
                  ByVal importChoi As Boolean, _
                  ByVal importExp As Boolean, _
                  ByVal importAna As Boolean, _
                  ByVal importTrans As Boolean, _
                  Optional ByVal cleanSetup As Boolean = False)

  Dim wb As Workbook
  Dim filePath As String
  Dim testFilePath As String

  Const NO_FILE_WARNING As String = "Error : no setup file was founded at the provided path"
  Const NO_FILE_WB As String = "Error : the specified file is not an excel workbook, or is broken, unable to open it."
  Const IMPORT_OPTIONS_WARNING As String = "Error: Please select at least one option (Dictionary, Analysis or Translations)"

  If (Not importDict) And (Not importAna) And (Not importTrans) And (Not importChoi) And (Not importExp) Then
    MsgBox IMPORT_OPTIONS_WARNING, vbExclamation + vbOkOnly
    WriteInfo IMPORT_OPTIONS_WARNING
    Exit Sub
  End If

  If cleanSetup Then Exit Sub

  filePath = Path()

  On Error Resume Next
  testFilePath = Dir(filePath)
  On Error GoTo 0

  'Test if the file exists
  If testFilePath = vbNullString Then
    MsgBox NO_FILE_WARNING, vbExclamation + vbOkOnly
    WriteInfo NO_FILE_WARNING
    Exit Sub
  End If

  'Test if the workbook could be loaded without problems
  BusyApp
  On Error Resume Next
  Set wb = Workbooks.Open(filePath)
  On Error GoTo 0

  If wb Is Nothing Then
    MsgBox NO_FILE_WB, vbExclamation + vbOkOnly
    WriteInfo NO_FILE_WB
    Exit Sub
  End If

  On Error Resume Next
  wb.Close savechanges:=False
  On Error GoTo 0
End Sub

Private Sub Clean(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
  Dim wb As Workbook
  Dim Lo As ListObject
  Dim sh As Worksheet
  Dim counter As Long
  Dim sheetName As String
  Dim csTab As ICustomTable
  Dim excludeNames As BetterArray

  Set wb = ThisWorkbook
  Set excludeNames = New BetterArray
  For counter = sheetsList.LowerBound To sheetsList.UpperBound
    sheetName = sheetsList.Item(counter)
    Set sh = wb.Worksheets(sheetName)
    pass.UnProtect sheetName
    For Each Lo In sh.ListObjects
      Set csTab = CustomTable.Create(Lo)
      csTab.Clean excludeNames:=excludeNames
      excludeNames.Clear
    Next
    pass.Protect sheetName, (sheetName = "Translations" Or sheetName = "Analysis"), _
    (sheetName = "Translations")

    'Delete all eventual comments/commentsthreaded in the worksheets
    sh.UsedRange.ClearComments
  Next
End Sub

Private Sub Import(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)

  'curr is for current (meaning in this workbook)
  'imp is for imported (meaning in the imported workbook)

  Dim impwb As Workbook
  Dim currwb As Workbook
  Dim sheetName As String
  Dim impsh As Worksheet
  Dim currsh As Worksheet
  Dim impLo As ListObject
  Dim currLo As ListObject
  Dim impTab As ICustomTable
  Dim currTab As ICustomTable
  Dim counter As Long

  BusyApp

  PrepareImport pass, sheetsList

  Set impwb = ImportWkb()
  Set currwb = ThisWorkbook


  For counter = sheetsList.LowerBound To sheetsList.UpperBound
    sheetName = sheetsList.Item(counter)
    If sheetExists(impwb, sheetName) And sheetExists(currwb, sheetName) Then

      pass.UnProtect sheetName

      Set impsh = impwb.Worksheets(sheetName)
      Set currsh = currwb.Worksheets(sheetName)

      For Each impLo In impsh.ListObjects
        'Test if the listObject is Present in current Sheet before proceeding
        If LoExists(currsh, impLo.Name) Then
          Set currLo = currsh.ListObjects(impLo.Name)
          Set impTab = CustomTable.Create(impLo)
          Set currTab = CustomTable.Create(currLo)

          'Import the custom table in the current sheet
          'For translations, import all the headers of imported sheet
          currTab.Import impTab, keepImportHeaders := (sheetName = "Translations")

        End If
      Next

      'Translations And Analysis need strict protection (avoid deleting rows)
      pass.Protect sheetName, (sheetName = "Translations" Or sheetName = "Analysis"), _
                  (sheetName = "Translations")
    End If
  Next

  impwb.Close savechanges:=False
End Sub



'Check if a sheet Exists in a workbook
Private Function sheetExists(ByVal wb As Workbook, sheetName As String) As Boolean
  If wb Is Nothing Then Exit Function
  Dim sh As Worksheet
  On Error Resume Next
  Set sh = wb.Worksheets(sheetName)
  On Error GoTo 0
  sheetExists = (Not (sh Is Nothing))
End Function

Private Function LoExists(ByVal sh As Worksheet, ByVal loname As String) As Boolean
  If sh Is Nothing Then Exit Function
  Dim Lo As ListObject
  On Error Resume Next
    Set Lo = sh.ListObjects(loname)
  On Error GoTo 0
  LoExists = (Not (Lo Is Nothing))
End Function


'Errors
'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'-------- INTERFACE
Public Sub ISetupImport_Check(ByVal impDict As Boolean, _
                              ByVal impChoi As Boolean, _
                              ByVal impExp As Boolean, _
                              ByVal impAna As Boolean, _
                              ByVal impTrans As Boolean, _
                              Optional ByVal cleanSetup As Boolean = False)
  check impDict, impChoi, impExp, impAna, impTrans, cleanSetup
End Sub

Public Sub ISetupImport_Import(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
  Import pass, sheetsList
End Sub

Public Sub ISetupImport_Clean(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
  Clean pass, sheetsList
End Sub
