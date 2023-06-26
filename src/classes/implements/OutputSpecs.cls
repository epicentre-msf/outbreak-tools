VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "OutputSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Specifications for exports in the linelist"

'@PredeclaredId
'@Folder("Linelist Exports")
'@ModuleDescription("Specifications for exports in the linelist")

Option Explicit
Implements IOutputSpecs

Private Type TExpOut
  outwb As Workbook 'Output workbook
  inpwb As Workbook 'Input workbok
  scope As Byte
End Type

Private Const CLASSNAME                 As String = "OutputSpecs"
Private Const DICTIONARYSHEET           As String = "Dictionary"
Private Const DICTTRANSLATIONSHEET      As String = "Translations"
Private Const CHOICESSHEET              As String = "Choices"
Private Const METADATASHEET             As String = "Metadata"
Private Const PASSWORDSHEET             As String = "__pass" 'password sheets
Private Const TEMPSHEET                 As String = "temp__" 'temporary boggy sheet
Private Const EXPORTSHEET               As String = "Exports"

Private this As TExpOut


Public Function Create(currwb As Workbook,  ByVal scope As Byte) As IOutputSpecs

  CheckRequirements currwb
  
  With New OutputSpecs
    Set .InputWkb = currwb
    Set .ActualScope = scope
    Set Create = .Self()
  End With

End Function

Public Property Get Self() As IOutputSpecs
  Set Self = Me
End Property

Public Property Get InputWkb() As Workbook
  Set InputWkb = this.inpwb
End Property

Public Property Set InputWkb(ByVal currwb As Workbook)
    Set this.inpwb = currwb 
End Property

Public Property Get OutputWkb() As Workbook

  If this.outwb is Nothing Then 
    BusyApp
    Set this.outwb = Workbooks.Add
  End If
  
  Set OutputWkb = this.outwb
End Property

Public Property Get ActualScope() As Byte
  ActualScope = this.scope
End Property

Public Property Let ActualScope(ByVal scope As Byte)
    this.scope = scope 
End Property

Private Property Get ExportDictionary() As ILLdictionary
  
  Dim dict As ILLdictionary
  Dim wb As Workbook
  Dim tempsh As Worksheet

  Set wb = InputWkb()
  Set tempsh = wb.Worksheets(TEMPSHEET)
  Set dict = LLdictionary.Create(tempsh, 1, 1)
  Set ExportDictionary = dict

End Property

Private Property Get ExportObject() As ILLExport
  Dim sh As Worksheet
  Dim wb As Workbook
  Dim exp As ILLExport

  Set wb = InputWkb()
  Set sh = wb.Worksheets(EXPORTSHEET)
  Set exp = LLExport.Create(sh)

  Set ExportObject = exp
End Property


'Prepare the output workbook
Private Sub Prepare()
  
  Dim inpwb As Workbook
  Dim tempsh As Worksheet
  Dim dictData As BetterArray
  Dim dictsh As Worksheet
  Dim expName As String
  Dim dict As ILLdictionary
  Dim expdict As ILLdictionary
  Dim scope As Byte
  Dim sheetNamesList As BetterArray
  Dim outwb As Workbook
  Dim firstSheetName As String

  Dim counter As Long
  Dim prevSheetName As String
  Dim expDictRng As Range
  
  Set inpwb = InputWkb()
  Set tempsh = inpwb.Worksheets(TEMPSHEET)
  Set dictsh = inpwb.Worksheets(DICTIONARYSHEET)
  Set dict = LLdictionary.Create(dictsh, 1, 1)
  Set dictData = New BetterArray
  Set sheetNamesList = New BetterArray
  Set outwb = OutputWkb()

  firstSheetName = outwb.Worksheets(1).Name
  scope = ActualScope()

  If scope <> ExportAll Then 
    expName = "export" & " " & CStr(scope)  
    'Filter the dictionary
    Set dictData = dict.Data.FilterData(varName:=expName, criteriaName:="<>", returnedColumnName:="__all__")
    'The temporary sheet contains the dictionary for the exports
  Else
    dictData.FromExcelRange dict.DataRange(includeHeaders:=True)
  End If

  BusyApp 'Filetering can return back calculations
  tempsh.Cells.Clear
  dictData.ToExcelRange tempsh.Cells(1, 1)
  
  'Sort the dictionary on exports and also on sheet names
  Set expdict = ExportDictionary()

  'the export dictionary datarange is used for sorting, and also for preparing the dictionary
  Set expDictRng = expdict.DataRange(includeHeaders := True)

  'Sort on table names and then on export name
  If scope <> ExportAll Then

    expDictRng.Sort key1:= expdict.DataRange("table name", includeHeaders:=True), _ 
                    key2:=expdict.DataRange(expName, includeHeaders:=True), Header:=xlYes
  End If
  
  'Add the tags for preparation of the dictionary (the dictionary is already prepared in the linelist)
  expDictRng.Cells(expDictRng.Rows.Count + 1, 1).Font.Color = vbBlue

  'Add the required worksheets after sorting
  Set sheetNamesList = expdict.UniqueValues("sheet name")

  'Add the metadata and other worksheets
  prevSheetName = firstSheetName

  If (exp.Value("export metadata", CInt(scope)) = "yes") Or (scope = ExportAll) Then 
    AddMetaData scope
    prevSheetName = METADATASHEET
  End If

  'Add the worksheets
  For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound
    With outwb
      .Worksheets.Add(after:=.Worksheets(prevSheetName)).Name = _ 
         sheetNamesList.Item(counter)
         prevSheetName = sheetNamesList.Item(counter)
    End With
  Next

  'Delete the first sheet of the workbook
  outwb.Worksheets(firstSheetName).Delete
End Sub

Private Sub AddMetaData(ByVal scope As Byte)

  Dim outwb As Workbook
  Dim inpwb As Workbook
  Dim csData As IDataSheet
  Dim sheetNamesList As BetterArray
  Dim counter As Long
  Dim sh As Worksheet
  Dim varRng As Range
  Dim expName As String


  Set sheetNamesList = New BetterArray
  Set inpwb = InputWkb()
  Set outwb = OutputWkb()

  sheetNamesList.Push METADATASHEET, DICTIONARYSHEET, CHOICESSHEET, DICTTRANSLATIONSHEET

  For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound
    Set sh = inpwb.Worksheets(sheetNamesList.Item(counter))
    Set csData = DataSheet.Create(sh, 1, 1)

    'Then you can just use the export sub from the datasheet
    csData.Export Wkb:=outwb, Hide:=xlSheetVisible
  Next

  'Add informations on metadatasheet
  Set sh = inpwb.Worksheets(METADATASHEET)
  Set csData = DataSheet.Create(sh, 1, 1)
  Set varRng = csData.DataRange("variable")

  'Add export name/type
  If scope = ExportAll Then 
    expName = "export for migrations"
  Else
    expName = "custom export: export" & " " & CStr(scope)
  End If

  varRng.Cells(varRng.Rows.Count + 1, 1).Value = "export_type"
  varRng.Cells(varRng.Rows.Count + 1, 2).Value = expName

  'Add the Date of the export
  varRng.Cells(varRng.Rows.Count + 2, 1).Value = "export_date"
  varRng.Cells(varRng.Rows.Count + 2, 2).Value = Format(Now, "yyyy-mm-dd Hh:Nn")
End Sub

'Add Data to export
Private Sub AddData(Optional ByVal useFilters As Boolean = False)

  Dim expdict As ILLdictionary
  Dim headersList As BetterArray
  Dim sheetNamesList As BetterArray
  Dim counter As Long
  Dim outwb As workbook
  Dim csTab As ICustomTable
  Dim Lo As ListObject
  Dim sheetName As String
  Dim useLabels As Boolean
  Dim expobj As ILLExport
  Dim scope As Byte
  Dim inpsh As Worksheet
  Dim inpwb As Workbook
  Dim outsh As Worksheet

  Set expdict = ExportDictionary()
  Set outwb =  OutputWkb()
  Set inpwb = InputWkb()
  Set sheetNamesList = expdict.UniqueValues("sheet name")
  Set expobj = ExportObject()
  scope = ActualScope()

  useLabels = (expobj.Value("export header", CInt(scope)) = "variable labels")

  For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound

    'Find the sheet
    sheetName = sheetNamesList.Item(counter)
    Set inpsh = inpwb.Worksheets(sheetName)
    Set headersList = expdict.Data.FilterData("sheet name", sheetName, "variable name")
    
    'Test if the sheet is hlist or vlist
    If (inpsh.Cells(1, 3).Value = "HList") Then

      Set outsh = outwb.Worksheets(sheetName)
      
      'Test if I should use filtered Data or actual data
      If useFilters Then
        'ListObject on printed sheet the 
        Set Lo = inpwb.Worksheets(inpsh.Cells(1, 5).Value).ListObjects(1)
      Else
        Set Lo = inpsh.ListObjects(1)
      End If

      'Define the custom table to export
      Set csTab = CustomTable.Create(Lo)

      'Export the custom table to the worksheet
      csTab.Export outsh, headersList, useLabels

    ElseIf (inpsh.Cells(1, 3).Value = "VList") Then
      
      'Add VlistData
      AddVlistData sheetName,  headersList, useLabels
    
    End If
  Next

End Sub

'Add data to vlist sheet
Private Sub AddVlistData(ByVal sheetName As String, ByVal headersList As BetterArray, _ 
                         ByVal useLabels As Boolean)

  Dim outsh As Worksheet
  Dim inpsh As Worksheet
  Dim outwb As Workbook
  Dim inpwb As Workbook
  Dim varValue As String
  Dim varName As String
  Dim counter As Long
  Dim headerCounter As Long

  Set inpwb = InputWkb()
  Set outwb = OutputWkb()

  Set inpsh = inpwb.Worksheets(sheetName)
  Set outsh = outwb.Worksheets(sheetName)

  outsh.Cells.Clear
  outsh.Cells(1, 1).Value = "variable"
  outsh.Cells(1, 2).Value = "value"

  'Counter on variables added
  headerCounter = 2

  For counter = headersList.LowerBound To headersList.UpperBound

    'Get values from input sheet
    varName = headersList.Item(counter)

    On Error Resume Next
      varValue = inpsh.Range(varName).Value
      If uselabels Then varName = inpsh.Range(varName).Offset(, -1).Value
    On Error GoTo 0

    'Write values to output sheet. I don't know if it suceeded in getting the value
    'from the input sheet, that is why I test it before proceeding,
    'otherwise do nothing

    If (varValue <> vbNullString) Then
      outsh.Cells(headerCounter, 1).Value = varName
      outsh.Cells(headerCounter, 2).Vaue = varValue
      varValue = vbNullString
      headerCounter = headerCounter + 1
    End If
  Next

End Sub

Private Sub Save()

End Sub


  'Checking requirements before instanciating the class
Private Sub CheckRequirements(ByVal currwb As Workbook)

  Dim counter As Long
  Dim sheetName As String
  Dim sheetsArray As BetterArray
  Set sheetsArray = New BetterArray

  sheetsArray.Push DICTIONARYSHEET, CHOICESSHEET, DICTTRANSLATIONSHEET, _ 
                   METADATASHEET, PASSWORDSHEET, TEMPSHEET, EXPORTSHEET

  For counter = sheetsArray.LowerBound To sheetsArray.UpperBound
    sheetName = sheetsArray.Item(counter)
    sheetExists currwb, sheetName
  Next
End Sub



'Returns true or false if a sheet or a range exists
Private Sub sheetExists(ByVal specswkb As Workbook, ByVal sheetName As String)
  Dim sh As Worksheet                          'Just try to set the workbook if it fails it is closed
  Dim checkExists As Boolean

  On Error Resume Next
  Set sh = specswkb.Worksheets(sheetName)
  checkExists = (Not sh Is Nothing)
  On Error GoTo 0

  If Not checkExists Then ThrowErrorUnableInstanciate " Worksheet " & sheetName & " Not found!"

End Sub

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

'Exposed Methods -------------------------------------------------


Private Sub IOutputSpecs_Prepare() 
  Prepare
End Sub

'Add Dictionary, Translations or Analysis
Private Sub IOutputSpecs_AddData()
  AddData
End Sub

Private Sub IOutputSpecs_Save()
  Save 
End Sub