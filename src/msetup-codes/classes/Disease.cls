VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "Disease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Add, create and manipulate new diseases"

'@Folder("Diseases")
'@PredeclaredId
'@ModuleDescription("Add, create and manipulate new diseases")
'@IgnoreModule ConstantNotUsed, ProcedureNotUsed

Implements IDisease

Option Explicit

Private Const LSTNAME As String = "__diseases_list"
Private Const CLASSNAME As String = "Disease"
Private Const VARIABLESHEET As String = "Variables"
Private Const TABTAG As String = "disTab_"
Private Const LANGTAG As String = "disLang_"
Private Const TRANSLATIONSHEET As String = "Translations"
Private Const RIBBONTRADSHEET As String = "__ribbonTranslation"
Private Const IMPORTREPSHEET As String = "__impRep"
Private Const CHOICESHEET As String = "Choices"

'Default table columns in the parameter variable sheet 
Private Const PARAMVARNAME As String = "PARAMVARNAME"
Private Const PARAMVARLABEL As String = "PARAMVARLABEL"
Private Const PARAMDEFSTATUS As String = "PARAMDEFSTATUS"
Private Const PARAMCOMMENTS As String = "PARAMCOMMENTS"
Private Const PARAMDEFCHOICE As String = "PARAMDEFCHOICE"
Private Const PARAMCHOICESLIST As String = "PARAMCHOICESLIST"
Private Const PARAMVARTABLE As String = "Tab_Variables"

'Import Report Tables
Private Const TABIMPREP As String = "TabImpRep"
Private Const TABIMPVAR As String = "TabImpVars"
Private Const TABNOTFOUND As String = "TabNotFoundVars"


Private Type TDis
  wb As Workbook
  dropObj As IDropdownLists
  outwb As Workbook
  disFolderPath As String
  migFolderPath As String
End Type

Private this As TDis

'Interface
Public Function Create(ByVal wb As Workbook, ByVal dropObj As IDropdownLists) As IDisease

  CheckRequirements wb, dropObj

  With New Disease
    Set .CurrentWorkbook = wb
    Set .DropdownObject = dropObj
    Set Create = .Self
  End With

End Function

'Self
Public Property Get Self() As IDisease
    Set Self = Me
End Property

'wb is where to create the disease worksheet 
Public Property Get CurrentWorkbook() As Workbook
  Set CurrentWorkbook = this.wb
End Property

Public Property Set CurrentWorkbook(ByVal wb As Workbook)
  Set this.wb = wb
End Property

'dropObj is the list of all dropdowns (including the disease one)
Public Property Get DropdownObject() As IDropdownLists
  Set DropdownObject = this.dropObj
End Property

Public Property Set DropdownObject(ByVal dropObj As IDropdownLists)
  Set this.dropObj = dropObj
End Property

Private Property Get RibTradObject() As ITranslation

  Dim ribLo As ListObject
  Dim fileLang As String
  Dim ribSh As Worksheet
  Dim wb As Workbook


  Set wb = CurrentWorkbook()
  Set ribSh = wb.Worksheets(RIBBONTRADSHEET)
  Set ribLo = ribSh.ListObjects(1)
  fileLang = ribSh.Range("RNG_FileLang").Value
  Set RibTradObject = Translation.Create(ribLo, fileLang)
End Property

Private Sub AddName(ByVal diseaseName As String)
  Dim dropObj As IDropdownLists
  Dim lstElmts As BetterArray

  Set dropObj = DropdownObject()
  Set lstElmts = dropObj.Items(LSTNAME)

  If Not lstElmts.Includes(diseaseName) Then
    'The first element could be empty, in case there is no disease in the
    'worksheet
    If lstElmts.Item(lstElmts.LowerBound) = vbNullString Then 
      lstElmts.Item(lstElmts.LowerBound) = diseaseName
    Else 
      lstElmts.Push diseaseName
    End If
    dropObj.Update lstElmts, LSTNAME
  End If
End Sub

Private Sub RemoveName(ByVal rmdisName As String) 
  Dim dropObj As IDropdownLists
  Dim lstElmts As BetterArray
  Dim newlstElmts As BetterArray
  Dim counter As Long
  Dim disName As String

  Set dropObj = DropdownObject()
  Set lstElmts = dropObj.Items(LSTNAME)

  If lstElmts.Includes(rmdisName) Then
    Set newlstElmts = New BetterArray
    For counter = lstElmts.LowerBound To lstElmts.UpperBound
      disName = lstElmts.Item(counter) 
      If (disName <> rmdisName) Then newlstElmts.Push disName
    Next

    dropObj.Update newlstElmts, LSTNAME
  End If
End Sub

Private Property Get Exists(ByVal disName As String) As Boolean 
  Dim dropObj As IDropdownLists
  Dim lstElmts As BetterArray

  Set dropObj = DropdownObject()
  Set lstElmts = dropObj.Items(LSTNAME)
  Exists = lstElmts.Includes(disName)
End Property

Private Property Get ExportFolder(Optional ByVal forMigration As Boolean = False) As String
  Dim io As IOSFiles
  Dim folderPath As String

  folderPath = IIf(forMigration, this.migFolderPath, this.disFolderPath)

  'Test if the folder has been set, in other case set it.
  If folderPath = vbNullString Then

    Set io = OSFiles.Create()
    
    io.LoadFolder
    If Not io.HasValidFolder() Then
      folderPath = vbNullString
    Else
      folderPath = io.Folder()

      'In case there is a \ or / at the end of the folder path, remove it because
      'we will need to paste the file name further. On local disks like D:\ this should
      'avoid ending up with D:\\... as file Path.
      If InStrRev(folderPath, Application.PathSeparator, _
        Len(folderPath)) = Len(folderPath) Then
        folderPath = Left(folderPath, (Len(folderPath) - 1))
      End If

      If forMigration Then
        this.migFolderPath = folderPath
      Else
        this.disFolderPath = folderPath
      End If

    End If
  End If

  ExportFolder = folderPath
End Property

Private Property Get OutputWkb() As Workbook

  If this.outwb Is Nothing Then
    BusyApp
    Set this.outwb = Workbooks.Add
    ActiveWindow.WindowState = xlMinimized
  End If

  Set OutputWkb = this.outwb
End Property

'Add a new disease Worksheet
Private Sub Add(ByVal disName As String, _ 
                Optional ByVal disLang As String = vbNullString)
  
  Dim wb As Workbook
  Dim disWksh As Worksheet
  Dim lstElmts As BetterArray
  Dim dropObj As IDropdownLists
  Dim nbDis As Long
  Dim loRng As Range
  Dim selectValue As String
  Dim trads As ITranslation
  

  'If the disease is present, exit
  If Exists(disName) Then Exit Sub

  Set wb = CurrentWorkbook()

  'The workbook should contain the "Variables" worksheet, worksheets are added
  'before this one
  If Not sheetExists(wb, VARIABLESHEET) Then Exit Sub

  Set trads = RibTradObject()
  'message error for the dropdowns in the disease worksheet
  selectValue = trads.TranslatedValue("selectValue")
  
  'Create the worksheet of the disease
  With wb
    .Worksheets.Add(before:=.Worksheets(VARIABLESHEET)).Name = disName
    Set disWksh = .Worksheets(disName)
  End With

  'Add the disease to the list of diseases
  AddName disName

  Set dropObj = DropdownObject()
  Set lstElmts = dropObj.Items(LSTNAME)
  nbDis = lstElmts.Length()


  'Add the disease table to the worksheet
  With disWksh
    
    'Add the validation for the language
    If disLang = vbNullString Then
      .Cells(2, 2).Value = trads.TranslatedValue("infoSelectLang")
    Else
      .Cells(2, 2).Value = disLang
    End If

    .Cells(2, 2).Name = LANGTAG & nbDis

    'write the number of diseases in the worksheet
    .Cells(2, 3).Value = nbDis
    .Cells(2, 4).Value = "DISSHEET"

    dropObj.SetValidation cellRng:=.Cells(2, 2), _ 
                          listName:="__languages", _ 
                          alertType:= "error", _
                          message:=trads.TranslatedValue("errLang")

    'Add the listObject for the disease

    .Cells(4, 2).Value = trads.TranslatedValue("varName")
    .Cells(4, 3).Value = trads.TranslatedValue("varLabel")
    .Cells(4, 4).Value = trads.TranslatedValue("varChoice")
    .Cells(4, 5).Value = trads.TranslatedValue("choiceVal")
    .Cells(4, 6).Value = trads.TranslatedValue("varStatus")
    '.Cells(4, 7).Value = trads.TranslatedValue("varVis")
    
    'Validation for variable Name
    With .Cells(5, 2).Validation
      .Delete

      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="= " & PARAMVARNAME
      
      'Messages for the validation
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = vbNullString
      .errorTitle = vbNullString
      .InputMessage = vbNullString
      .errorMessage = selectValue
      .ShowInput = True
      .ShowError = True
    End With

    'Variable choice
    With .Cells(5, 4).Validation
      .Delete

      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="= " & PARAMCHOICESLIST
      
      'Messages for the validation
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = vbNullString
      .errorTitle = vbNullString
      .InputMessage = vbNullString
      .errorMessage = vbNullString
      .ShowInput = True
      .ShowError = True
    End With

    'Variable status
     dropObj.SetValidation cellRng:=.Cells(5, 6), _ 
                          listName:="__var_status", _ 
                          alertType:= "error", _
                          message:=selectValue

    'Variable visibility
    '  dropObj.SetValidation cellRng:=.Cells(5, 7), _ 
    '                       listName:="__var_visibility", _ 
    '                       alertType:= "error", _
    '                       message:=selectValue
    ' 'Creating the listObject
    Set loRng = .Range(.Cells(4, 2), .Cells(5, 6))
    .ListObjects.Add(xlSrcRange, loRng, , xlYes).Name = TABTAG & nbDis
  End With

  'Format the disease worksheet
  ApplyFormat disWksh
End Sub


'Format the disease Worksheet
Private Sub ApplyFormat(ByVal sh As Worksheet)

  Dim Lo As ListObject
  Dim colorBlueLight As Long
  Dim colorBlueMedium As Long
  Dim colorBlueDark As Long
  Dim styleValue As Long
  Dim rng As Range

  'Remove the gridlines
  RemoveGridLines sh
  sh.Cells.EntireColumn.ColumnWidth = 25

  'For the first listObject, format the lines of the table
  Set Lo = sh.ListObjects(1)
  colorBlueLight = RGB(217, 225, 242)
  colorBlueMedium = RGB(47, 117, 181)
  colorBlueDark = RGB(15, 26, 48)
  styleValue = 9

  Lo.TableStyle = ""
  DrawLines Lo.Range, color:=colorBlueDark

  'Unlock first column of disease
  On Error Resume Next
  Lo.ListColumns(1).DataBodyRange.Locked = False
  Lo.ListColumns(2).DataBodyRange.Locked = True
  Lo.ListColumns(2).DataBodyRange.Font.Color = RGB(59, 8, 145)
  Lo.ListColumns(2).DataBodyRange.Font.Italic = True 
  Lo.ListColumns(3).DataBodyRange.Locked = False
  Lo.ListColumns(4).DataBodyRange.Locked = True
  Lo.ListColumns(4).DataBodyRange.Font.Color = RGB(59, 8, 145)
  Lo.ListColumns(4).DataBodyRange.Font.Italic = True
  'Unlock status and visibility
  Lo.ListColumns(5).DataBodyRange.Locked = False
  Lo.ListColumns(6).DataBodyRange.Locked = False
  On Error GoTo 0

  With sh
    'Unlock the dropdown for the language
    .Cells(2, 2).Locked = False
    'Number of disease column
    Set rng = .Range(.Cells(2, 3), .Cells(2, 4))
    FormatRange rng, interiorColor:=vbWhite, fontColor:=vbWhite
    rng.Locked = True
    rng.FormulaHidden = True

    'Dropdown label
    FormatRange .Cells(2, 2), fontColor:=vbWhite, _
                      interiorColor:=colorBlueMedium, _
                      isBold:=False, FontSize:=styleValue

    'Header of the the listObject in the disease worksheet
    Set rng = Lo.HeaderRowRange
    FormatRange rng, interiorColor:=colorBlueLight, isBold:=True, _
                    fontColor:=colorBlueDark, _
                    FontSize:=styleValue + 2, Verti:=xlVAlignCenter, _
                    Horiz:=xlHAlignCenter

  End With
End Sub

'Remove a disease.
Private Sub Remove(ByVal disName As String)

  Dim disWksh As Worksheet
  Dim wb As Workbook
  Dim nbDis As Long
  Dim prevAlert As Long
  Dim counter As Long

  'Remove the disase name from the dropdown
  RemoveName disName
  'Clear the current worksheet
  Set wb = CurrentWorkbook()
  Set disWksh = wb.Worksheets(disName)
  nbDis = disWksh.Cells(2, 3).Value
  On Error Resume Next
  'Unlist the listObject of the disease
  disWksh.ListObjects(1).Unlist
  'Remove the name of corresponding to the number of disease
  wb.Names(LANGTAG & nbDis).Delete
  'Delete the worksheet
  prevAlert = Application.DisplayAlerts
  Application.DisplayAlerts = False
  For counter = 1 To 4
    wb.Worksheets(disName).Delete
  Next
  Application.DisplayAlerts = prevAlert
  On Error GoTo 0
End Sub

Private Sub ExportDisease(ByVal disName As String)

  Const DEFAULTCHOICESEPARATOR As String = " | "

  Dim exportedFileName As String
  Dim outwb As Workbook
  Dim wb As Workbook
  Dim sh As Worksheet
  Dim disWksh As Worksheet
  Dim cellRng As Range
  Dim counter As Long
  Dim listCols As BetterArray
  Dim choiObj As IChoices
  Dim choiName As String
  Dim filePath As String
  Dim folderPath As String
  Dim prevAlert As Long
  Dim ribTrads As ITranslation
  Dim ribSh As Worksheet


  exportedFileName = "setup_dictionary_" & disName & ".xlsx"

  'Initialize the elements
  folderPath = ExportFolder()
  If folderPath = vbNullString Then Exit Sub

  Set outwb = outputWkb()
  Set wb = CurrentWorkbook()
  Set listCols = New BetterArray
  Set disWksh = wb.Worksheets(disName)
  Set ribsh = wb.Worksheets(RIBBONTRADSHEET)

  'Choices is converted in English by Default when exporting "ENG"
  'is the language Tag for English.
  Set ribTrads = Translation.Create(ribSh.ListObjects(1), "ENG")

  'Add the different worksheets to the export
  outwb.Worksheets(1).Name = "Translations"
  outwb.Worksheets.Add.Name = "Choices"
  outwb.Worksheets.Add.Name = "Dictionary"
  outwb.Worksheets.Add.Name = "Metadata"

  'Add Metadata elements
  'sh is the Metadata worksheet
  Set sh = outwb.Worksheets("Metadata")
  FormatSheet sh
  sh.Cells(1, 1).Value = "Label"
  sh.Cells(1, 2).Value = "Value"
  'Add the Date of the export
  sh.Cells(2, 1).Value = "export_date"
  sh.Cells(2, 2).Value = Format(Now, "yyyy-mm-dd Hh:Nn")
  'Add the disease of the export
  sh.Cells(3, 1).Value = "export_disease"
  sh.Cells(3, 2).Value = disName
  'Add the language of the worksheet
  sh.Cells(4, 1).Value = "dictonary_language"
  sh.Cells(4, 2).Value = disWksh.Cells(2, 2).Value

  'Add choices and dictionary columns

  'sh is now the choices worksheet
  FormatSheet outwb.Worksheets("Choices")
  Set sh = outwb.Worksheets("Choices")
  Set choiObj = Choices.Create(sh, 1, 1, ribTrads)
  choiObj.AddHeaders

  'sh is now the dictionary worksheet
  FormatSheet outwb.Worksheets("Dictionary")
  Set sh = outwb.Worksheets("Dictionary")

  sh.Cells(1, 1).Value = "Variable Name"
  sh.Cells(1, 2).Value = "Main Label"
  sh.Cells(1, 3).Value = "Control"
  sh.Cells(1, 4).Value = "Status"
  sh.Cells(1, 5).Value = "Visibility"

  Set cellRng = disWksh.Cells(5, 2)
  counter = 2

  Do While (Not IsEmpty(cellRng))

    sh.Cells(counter, 1).Value = cellRng.Value
    sh.Cells(counter, 2).Value = cellRng.Offset(, 1).Value
    sh.Cells(counter, 3).Value = cellRng.Offset(, 2).Value
    'The third offset is for choices values separated with a |
    'Status and visibility (status has only two values: Core, non core)
    'visibility has only three values (always show, optional show, optional hidden)
    sh.Cells(counter, 4).Value = cellRng.Offset(, 4).Value
    sh.Cells(counter, 5).Value = cellRng.Offset(, 5).Value

    'The third column is the control column in the dictionary worksheet (choice Name)
    choiName = sh.Cells(counter, 3).Value
    listCols.Clear
    'Those are the translated values of the choices
    listCols.Items = Split(cellRng.Offset(, 4).Value, DEFAULTCHOICESEPARATOR)
    choiObj.AddChoice choiName, listCols

    'Loop over
    Set cellRng = cellRng.Offset(1)
    counter = counter + 1
  Loop

  'Add Translation
  Set sh = outwb.Worksheets("Translations")
  FormatSheet sh
  listCols.FromExcelRange wb.Worksheets(TRANSLATIONSHEET).ListObjects(1).Range
  listCols.ToExcelRange sh.Cells(1, 1)

  'Save the file
  filePath = folderPath & Application.PathSeparator & exportedFileName

  prevAlert = Application.DisplayAlerts
  Application.DisplayAlerts = False
  outwb.SaveAs fileName:=filePath, _
  CreateBackup:=False, _
  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
  Application.DisplayAlerts = prevAlert
  outwb.Close saveChanges:=False
End Sub

Private Property Get HasReport(ByVal disName As String) As Boolean

  Dim wb As Workbook
  Dim repSh As Worksheet
  Dim repLo As ListObject
  Dim needRep As Boolean
  Dim cellRng As Range

  needRep = False
  Set wb = CurrentWorkbook()
  Set repSh = wb.Worksheets(IMPORTREPSHEET)
  Set repLo = repSh.ListObjects(TABIMPREP)

  Set cellRng = repLo.ListColumns(1).Range.Find(What:=disName, lookAt:=xlWhole, MatchCase:=True)

  If Not (cellRng Is Nothing) Then 
    needReport = (cellRng.Offset(, 1).Value = "yes")
  End If
End Property


'Remove the report status
Private Sub RemoveReportStatus(ByVal disName As String)

  Dim wb As Workbook
  Dim repSh As Worksheet
  Dim repLo As ListObject
  Dim counter As Long
  Dim delRowsTab As BetterArray
  
  'Remove all the rows of the disease
  Set wb = CurrentWorkbook()
  Set repSh = wb.Worksheets(IMPORTREPSHEET)
  Set delRowsTab = New BetterArray

  For Each repLo in repSh.ListObjects
    
    'Target all the rows to remove
    For counter = 1 To repLo.ListRows.Count
      If (repLo.ListRows(counter).Range.Cells(1, 1).Value = disName) Then
        delRows.Push counter
      End If
    Next

    If (delRowsTab.Length > 0) Then
      For counter = delRowsTab.LowerBound To delRowsTab.UpperBound
        repLo.ListRows(delRowsTab.Item(counter)).Delete
      Next
    End If

    delRowsTab.Clear
  Next
End Sub


Private Sub AddReportVars(ByVal disName As String, _ 
                          Optional ByVal varName As String = "yes", _ 
                          Optional ByVal tableName As String = TABIMPREP)

  Dim wb As Workbook
  Dim repSh As Worksheet
  Dim repLo As ListObject
  Dim cellRng As Range

  Set wb = CurrentWorkbook()
  Set repSh = wb.Worksheets(IMPORTREPSHEET)
  Set repLo = repSh.ListObjects(tableName)
  

  Set cellRng = repLo.Range.Cells(2, 1)

  Do While Not IsEmpty(cellRng)
    Set cellRng = cellRng.Offset(1)
  Loop

  cellRng.Value = disName
  cellRng.Offset(, 1).Value = varName

  repLo.Resize repSh.Range(repLo.Range.Cells(1, 1), cellRng.Offset(, 1))
End Sub


Private Property Get NotImportedVars(ByVal disName As String, _ 
                                     Optional ByVal tableName As String = TABIMPVAR) As BetterArray

  Dim lstVars As BetterArray
  Dim wb As Workbook
  Dim repSh As Worksheet
  Dim repLo As ListObject
  Dim counter As Long
  Dim rowRng As Range


  Set wb = CurrentWorkbook()
  Set repSh = wb.Worksheets(IMPORTREPSHEET)
  Set repLo = repSh.ListObjects(tableName)
  Set lstVars = New BetterArray

  For counter = 1 To repLo.ListRows.Count

    'Range of the listRow
    Set rowRng = repLo.ListRows(counter).Range

    If (rowRng.Cells(1, 1).Value = disName) Then
      lstVars.Push rowRng.Cells(1, 2).Value
    End If

  Next

  Set NotImportedVars = lstVars.Clone()
End Property

'Import values in the disease worksheet
'importPriority could be foreign or actual, actual means keep the variables
'in the actual worksheet, don't change or update them. If foreign, update with
'imported values from the foreign worksheet.

Private Sub ImportElements(ByVal Lo As ListObject, ByVal disName As String, _ 
                           Optional ByVal mergeValues As Boolean = True, _
                           Optional ByVal importPriority As String = "foreign")

  Dim wb As Workbook
  Dim disWksh As Worksheet

  'Actual and input custom table
  Dim actCsTab As ICustomTable
  Dim inCsTab As ICustomTable

  Dim varRng As Range
  Dim inpVarRng As Range

  Dim cellRng As Range
  Dim counter As Range
  Dim actLo As ListObject
  Dim needReport As Boolean



  'If the disease does not exits, then exit the sub

  If Not Exists(disName) Then Exit Sub
  Set wb = CurrentWorkbook()
  Set disWksh = wb.Worksheets(disName)

  'Import values without merging them (just simple import from the previous to the new one)

  If Not mergeValues Then
    
    Set actCsTab = CustomTable.Create(disWksh.ListObjects(1))
    Set inCsTab = CustomTable.Create(Lo)
    'Resize the table before import
    actCsTab.Resize
    actCsTab.Import inCsTab
    'Change the language
    disWksh.Cells(2, 2).Value = Lo.Range.Cells(0, 2).Value

  Else
  'Merge values (take in account variables in the the actual disease worksheet)

    Set inpVarRng = Lo.ListColumns(1).Range
    Set actLo = disWksh.ListObjects(1)
    Set varRng = actLo.ListColumns(1).Range

    'Look for variables in the current disease worksheet
  
    For counter = 2 To varRng.Rows.Count
      varName = varRng.Cells(counter, 1).Value

      'Look for the variable name
      Set cellRng = inpVarRng.Find(What:=varName, LookAt:=xlWhole, MatchCase:=True)
      
      'Variable not found in imports, don't touch it
      If (cellRng Is Nothing) Then

        needReport = True
        AddReportVars varName:=varName, disName:=disName, tableName:=TABNOTFOUND
      
      Else
      
        'Variable found in import, default is not to touch, unless importPriority is foreign
        If (importPriority = "foreign") Then
          varRng.Cells(counter, 2).Value = cellRng.Offset(, 1).Value
          varRng.Cells(counter, 3).Value = cellRng.Offset(, 2).Value
          varRng.Cells(counter, 4).Value = cellRng.Offset(, 3).Value
          varRng.Cells(counter, 5).Value = cellRng.Offset(, 4).Value
          varRng.Cells(counter, 6).Value = cellRng.Offset(, 5).Value
        End If

        'Reset the cell Range
        Set cellRng = Nothing
      
      End If
    Next

    'Look for variables from the imported disease worksheet
    For counter = 2 To inpVarRng.Rows.Count


    Next



    If needReport Then AddReportVars disName:=disName 
  End If 


End Sub


'Import all variables of a listObject into a disease worksheet
Private Sub Import(ByVal Lo As ListObject, ByVal disName As String, _ 
                   Optional ByVal mergeValues As Boolean = True, _ 
                   Optional ByVal importPriority As String = "foreign")
  
  'If the disease does not exists, create it, and import It

  If Not Exists(disName) Then
    Add disName
    ImportElements Lo:=Lo, disName:=disName, mergeValues:=False
  Else

    'If the disease already exists, make import using given attributes
    ImportElements Lo:=Lo, disName:=disName, _ 
                   mergeValues:=mergeValues, _ 
                   importPriority:=importPriority

  End If
End Sub


'Export diseases worksheets for Migration
Private Sub ExportForMigration()


  Dim outwb As Workbook
  Dim wb As Workbook
  Dim sh As Worksheet
  Dim dropObj As IDropdownLists
  Dim lstElmts As BetterArray
  Dim folderPath As String
  Dim filePath As String
  Dim counter As Long
  Dim lstCols As BetterArray
  Dim lstSheets As BetterArray
  Dim sheetName As String
  Dim disWksh As Worksheet
  Dim Lo As ListObject
  Dim LoRng As Range
  Dim outLoRng As Range
  Dim lastColumn As Long
  Dim prevAlert As Long


  folderPath = ExportFolder(forMigration:=True)
  If folderPath = vbNullString Then Exit Sub

  Set lstElmts = New BetterArray
  Set lstCols = New BetterArray
  Set lstSheets = New BetterArray
  Set dropObj = DropdownObject()
  Set wb = CurrentWorkbook()
  Set outwb = OutputWkb()

  'Add Worksheets
  outwb.Worksheets(1).Name = "Diseases"

  'Transfert data to Choices, translation and variable worksheets
  lstSheets.Push  TRANSLATIONSHEET, CHOICESHEET, VARIABLESHEET

  For counter = lstSheets.LowerBound To lstSheets.UpperBound
    sheetName = lstSheets.Item(counter)
    outwb.Worksheets.Add.Name = sheetName
    Set sh = outwb.Worksheets(sheetName)
    FormatSheet sh
    lstCols.FromExcelRange wb.Worksheets(sheetName).ListObjects(1).Range
    lstCols.ToExcelRange sh.Cells(1, 1)
  Next

  'Get the list of Diseases
  Set lstElmts = dropObj.Items(LSTNAME)
  Set sh = outwb.Worksheets("Diseases")
  RemoveGridLines sh

  'Transfert Data to the Disease worksheet

  For counter = lstElmts.LowerBound To lstElmts.UpperBound
    Set disWksh = wb.Worksheets(lstElmts.Item(counter))

    Set Lo = disWksh.ListObjects(1)
    Set LoRng = Lo.Range

    With sh
      lastColumn = .Cells(3, .Columns.Count).End(xlToLeft).Column + 2
      'Remove 1 when adding to take in account the whole plage of the range
      Set outLoRng = .Range(.Cells(3, lastColumn),  _ 
                        .Cells(2 + LoRng.Rows.Count, lastColumn + LoRng.Columns.Count - 1))
      outLoRng.Value = LoRng.Value
      .ListObjects.Add(xlSrcRange, outLoRng, , xlYes).Name = Lo.Name

      'Add the language, the disease tag
      .Cells(2, lastColumn).Value = disWksh.Name
      .Cells(2, lastColumn + 1).Value = disWksh.Cells(2, 2).Value
      .Cells(2, lastColumn + 2).Value = disWksh.Cells(2, 3).Value

      FormatRange .Range(.Cells(2, lastColumn), .Cells(2, lastColumn + 2)), _ 
                  fontColor:=RGB(178, 190, 181)
    End With
  Next

  'Save the workbook to the corresponding filePath
  filePath = folderPath & Application.PathSeparator & _ 
             "flat_disease_file_" & Format(Now, "yyyy-mm-dd") & ".xlsx"

  prevAlert = Application.DisplayAlerts
  Application.DisplayAlerts = False
  outwb.SaveAs fileName:=filePath, _
  CreateBackup:=False, _
  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
  Application.DisplayAlerts = prevAlert
  outwb.Close saveChanges:=False
  
End Sub


'Checkings and errors, and transitory subs for work

Private Sub CheckRequirements(ByVal wb As Workbook, ByVal dropObj As IDropdownLists)

  If wb Is Nothing Then ThrowError "Unable to initialize the class, workbook object not initialized"
  If dropObj Is Nothing Then ThrowError "Unable to initialize the class, dropdown object not found"
  If Not dropObj.Exists(LSTNAME) Then ThrowError "Unable to initialize, there is no disease list in dropdowns"
  If Not sheetExists(wb, VARIABLESHEET) Then ThrowError "Unable to find the variables worksheet in the current workbook"
  If Not sheetExists(wb, TRANSLATIONSHEET) Then ThrowError "Unable to find the translation worksheet in the current workbook"
  If Not sheetExists(wb, CHOICESHEET) Then ThrowError "Unable to find the choice worksheet in the current workbook"
  If Not sheetExists(wb, IMPORTREPSHEET) Then ThrowError "Unable to find import report worksheet in the current workbook"
  If Not sheetExists(wb, RIBBONTRADSHEET) Then ThrowError "Unable to find ribbon translation worksheet in the current workbook"
End Sub


Private Function sheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
  Dim sh As Worksheet
  On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
  On Error GoTo 0
  sheetExists = Not (sh Is Nothing)
End Function

'Remove Gridlines of a worksheet
Private Sub RemoveGridLines(ByVal sh As Worksheet, Optional DisplayZeros As Boolean = False)
    Dim View As WorksheetView
    For Each View In sh.Parent.Windows(1).SheetViews
        If View.Sheet.Name = sh.Name Then
            View.DisplayGridlines = False
            View.DisplayZeros = DisplayZeros
            Exit Sub
        End If
    Next
End Sub


'Format a Range
Private Sub FormatRange(ByVal rng As Range, _
                        Optional ByVal fillValue As String = "", _
                        Optional ByVal interiorColor As Variant = "", _
                        Optional ByVal fontColor As Variant = "", _
                        Optional ByVal isBold As Boolean = False, _
                        Optional ByVal Horiz As Integer = xlHAlignCenter, _
                        Optional ByVal Verti As Integer = xlVAlignCenter, _
                        Optional ByVal FontSize As Double = 0, _
                        Optional ByVal NumFormat As String = "")

    With rng
        If interiorColor <> vbNullString Then .Interior.color = interiorColor
        If fontColor <> vbNullString Then .Font.color = fontColor
        If isBold Then .Font.Bold = isBold
        If FontSize <> 0 Then .Font.Size = FontSize
        If NumFormat <> vbNullString Then .NumberFormat = NumFormat
        .HorizontalAlignment = Horiz
        .VerticalAlignment = Verti
        If fillValue <> vbNullString Then .Value = fillValue
    End With
End Sub

'Format Export Worksheet
Private Sub FormatSheet(ByVal sh As Worksheet)
  RemoveGridLines sh
  With sh
    On Error Resume Next
    .Cells.EntireColumn.ColumnWidth = 25
    .Cells.Font.Size = 9
    .Rows(1).Font.Size = 10
    .Rows(1).RowHeight = 20
    .Rows(1).Interior.Color = RGB(240, 240, 244)
    .Rows(1).Font.Bold = True
    .Cells.Font.Name = "Consolas"
    .Cells.WrapText = True
    .Cells.HorizontalAlignment = xlHAlignLeft
    .Cells.VerticalAlignment = xlVAlignCenter
    .Cells.EntireRow.AutoFit
    On Error GoTo 0
  End With
End Sub

Private Sub DrawLines(ByVal rng As Range, _
                      Optional At As String = "All", _
                      Optional weight As Integer = xlHairline, _
                      Optional line As Integer = xlContinuous, _
                      Optional color As Variant = vbBlack)
    Dim borderpos As Byte
    Dim bordersList As BetterArray
    Dim counter As Long

    Set bordersList = New BetterArray

    Select Case At

    Case "All"
        bordersList.Push xlEdgeLeft, xlEdgeRight, xlEdgeBottom, xlEdgeTop, _ 
                         xlInsideHorizontal, xlInsideVertical
    Case "AvoidLeft", "AvoidRight"
        bordersList.Push xlEdgeBottom, xlEdgeTop, xlInsideHorizontal
    Case "AvoidTop", "AvoidBottom"
        bordersList.Push xlEdgeLeft, xlEdgeRight, xlInsideVertical
    Case "Left"
        bordersList.Push xlEdgeLeft
    Case "Right"
        bordersList.Push xlEdgeRight
    Case "Bottom"
        bordersList.Push xlEdgeBottom
    Case "Top"
        bordersList.Push xlEdgeTop
    Case Else
        bordersList.Push xlEdgeBottom
    End Select

    With rng
      For counter = bordersList.LowerBound To bordersList.UpperBound
          borderpos = bordersList.Item(counter)
          With .Borders(borderpos)
            .weight = weight
            .LineStyle = line
            .color = color
            .TintAndShade = 0.2
          End With
      Next
    End With
End Sub

Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.CalculateBeforeSave = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Interfaces -------------------------------

Private Sub IDisease_Add(ByVal disName As String, _ 
                        Optional ByVal disLang As String = vbNullString)
  Add disName, disLang
End Sub

Private Sub IDisease_Remove(ByVal disName As String)
  Remove disName
End Sub

Private Sub IDisease_ExportDisease(ByVal disName As String)
  ExportDisease disName
End Sub

Private Property Get IDisease_Exists(ByVal disName As String) As Boolean
  IDisease_Exists = Exists(disName)
End Property

Private Property Get IDisease_OutputWkb() As Workbook
  Set IDisease_OutputWkb = OutputWkb()
End Property

Private Sub IDisease_ExportForMigration()
  ExportForMigration
End Sub

Private Sub IDisease_Import(ByVal Lo As ListObject, ByVal disName As String)
  Import Lo, disName
End Sub