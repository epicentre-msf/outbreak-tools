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

'@IgnoreModule UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Specifications for exports in the linelist")

Option Explicit
Implements IOutputSpecs

Private Type TExpOut
  outwb As Workbook 'Output workbook
  inpwb As Workbook 'Input workbook
  scope As Integer 'The scope of the export: export 1 to 5 or more, or export for migration
  folderPath As String 'where to save the exports
  folderSet As Boolean 'Is the path to save the exports set?
  outgeowb As Workbook 'Output geo workbook
  outgeohistwb As Workbook 'Output for the geobase historic file
End Type

Private Const CLASSNAME                 As String = "OutputSpecs"
Private Const DICTIONARYSHEET           As String = "Dictionary"
Private Const DICTTRANSLATIONSHEET      As String = "Translations"
Private Const CHOICESSHEET              As String = "Choices"
Private Const METADATASHEET             As String = "Metadata"
Private Const PASSWORDSHEET             As String = "__pass" 'password sheets
Private Const TEMPSHEET                 As String = "temp__" 'temporary boggy sheet
Private Const EXPORTSHEET               As String = "Exports"
Private Const ANANAMESSHEET             As String = "ana_tabnames__"
Private Const MESSTRANSLATIONSHEET      As String = "LinelistTranslation"
Private Const SHOWHIDESHEET             As String = "show_hide__"

Private this As TExpOut


Public Function Create(currwb As Workbook, ByVal scope As Integer) As IOutputSpecs

  CheckRequirements currwb

  With New OutputSpecs
    Set .InputWkb = currwb
     .ActualScope = scope
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

Private Property Get outputWkb() As Workbook

  If this.outwb Is Nothing Then
    BusyApp
    Set this.outwb = Workbooks.Add
    ActiveWindow.WindowState = xlMinimized
  End If

  Set outputWkb = this.outwb
End Property

'Get the linelist Translation object, with all elements related
'to translations at the linelist level
Private Property Get LLTradsObject() As ILLTranslations
  Dim currwb As Workbook

  Set currwb = InputWkb()
  Set LLTradsObject = LLTranslations.Create( _
                          currwb.Worksheets(MESSTRANSLATIONSHEET), _
                          currwb.Worksheets(DICTTRANSLATIONSHEET) _
                        )
End Property


'Get the translation of messages object
Private Property Get MessagesTrads() As ITranslation
  Dim lltrads As ILLTranslations
  Set lltrads = LLTradsObject()
  Set MessagesTrads = lltrads.TransObject()
End Property

'Geobase export workbook
Private Property Get OutputGeoWkb() As Workbook

  If this.outgeowb Is Nothing Then
    BusyApp
    Set this.outgeowb = Workbooks.Add
    ActiveWindow.WindowState = xlMinimized
  End If

  Set OutputGeoWkb = this.outgeowb
End Property

'Geobase export workbook
Private Property Get OutputGeoHistoricWkb() As Workbook

  If this.outgeohistwb Is Nothing Then
    BusyApp
    Set this.outgeohistwb = Workbooks.Add
    ActiveWindow.WindowState = xlMinimized
  End If

  Set OutputGeoHistoricWkb = this.outgeohistwb
End Property


Public Property Get ActualScope() As Integer
  ActualScope = this.scope
End Property

Public Property Let ActualScope(ByVal scope As Integer)
    this.scope = scope
End Property

Private Property Get ExportDictionary() As ILLdictionary

  Dim dict As ILLdictionary
  Dim wb As Workbook
  Dim tempsh As Worksheet
  Dim nbExp As Long

  Set wb = InputWkb()
  Set tempsh = wb.Worksheets(TEMPSHEET)
  nbExp = NumberOfExports()
  Set dict = LLdictionary.Create(tempsh, 1, 1, nbExp)
  Set ExportDictionary = dict

End Property

Private Property Get ExportObject() As ILLExport
  Dim sh As Worksheet
  Dim wb As Workbook
  Dim expobj As ILLExport

  Set wb = InputWkb()
  Set sh = wb.Worksheets(EXPORTSHEET)
  Set expobj = LLExport.Create(sh)

  Set ExportObject = expobj
End Property

Private Property Get NumberOfExports() As Long
  Dim expObj As ILLExport
  Set expObj = ExportObject()
  NumberOfExports = expObj.NumberOfExports()
End Property

'Prepare the output workbook
'withAnalysis: Add analysis worksheets
'includeShowHide: Add show/hide in the output
'keepLabels: Add a tag for updating editableLabel variables on import
Private Sub Prepare(Optional ByVal withAnalysis As Boolean = False, _ 
                    Optional ByVal includeShowHide As Boolean = False, _ 
                    Optional ByVal keepLabels As Boolean = False)

  Dim inpwb As Workbook
  Dim tempsh As Worksheet
  Dim dictData As BetterArray
  Dim dictsh As Worksheet
  Dim expName As String
  Dim dict As ILLdictionary
  Dim expdict As ILLdictionary
  Dim scope As Integer
  Dim sheetNamesList As BetterArray
  Dim outwb As Workbook
  Dim firstSheetName As String
  Dim expobj As ILLExport
  Dim counter As Long
  Dim prevSheetName As String
  Dim expDictRng As Range
  Dim nbExp As Long
  
  
  'Variables for p-codes removal from the exports
  Dim expNumberRng As Range
  Dim expMainLabRng As Range
  Dim pcodesTagList As BetterArray
  Dim exportpcodes As Boolean

  'Variables for personal identifiers removal from the export
  Dim expIdentiferRng As Range
  Dim exportPI As Boolean

  Set inpwb = InputWkb()
  Set tempsh = inpwb.Worksheets(TEMPSHEET)
  Set dictsh = inpwb.Worksheets(DICTIONARYSHEET)
  nbExp = NumberOfExports()
  Set dict = LLdictionary.Create(dictsh, 1, 1, nbExp)
  Set dictData = New BetterArray
  Set sheetNamesList = New BetterArray
  Set outwb = outputWkb()
  Set expobj = ExportObject()
  Set pcodesTagList = New BetterArray

  firstSheetName = outwb.Worksheets(1).Name
  scope = ActualScope()

  If (scope <> ExportAll) Then 'Here export is export 1, 2, 3, 4 or 5 or ...
    expName = "export" & " " & CStr(scope)

    'Export Or Not the p-Codes.
    'I will add pcodes informations on the dictionary if export pcodes == yes
    'otherwise pcodes will be removed as default behavior. If you want pcodes
    'to be exported, you need to explicitly add export pcodes == Yes in the
    'export sheet of the setup file. Otherwise they will be removed from exports.
    exportpcodes = (expobj.Value("include p-codes", scope) = "yes")

    If Not exportpcodes Then
      pcodesTagList.Push "geo_pcode_adm1", "geo_pcode_adm2", "geo_pcode_adm3", _
                         "geo_pcode_adm4", "geo_concat_adm1", "geo_concat_adm2", _
                         "geo_concat_adm3", "geo_concat_adm4"

      Set expNumberRng = dict.DataRange(expName)
      Set expMainLabRng = dict.DataRange("main label")

      'remove the export number for those pcodes elements (including concat elements)
      For counter = 1 To expMainLabRng.Rows.Count
        If pcodesTagList.Includes(expMainLabRng.Cells(counter, 1).Value) Then
          expNumberRng.Cells(counter, 1).ClearContents
        End If
      Next
    End If

    'Export or Not the Personal identifiers
    'personal identifiers informations are added only if required by the user
    exportPI = (expobj.Value("include personal identifiers", scope) = "yes")

    'Exclude personal indentifiers from the exports
    If Not exportPI Then
      Set expNumberRng = dict.DataRange(expName)
      Set expIdentiferRng = dict.DataRange("personal identifier")

      For counter = 1 To expIdentiferRng.Rows.Count
          If (expIdentiferRng.Cells(counter, 1).Value = "yes") Then
              expNumberRng.Cells(counter, 1).ClearContents
          End If
      Next
    End If

    'It is important to remove pcodes elements before filtering the linelist.
    'Filter the dictionary on only variables to includes in exports
    Set dictData = dict.Data.FilterData(varName:=expName, _
                                        criteriaName:="<>", _
                                        returnedColumnName:="__all__", _
                                        includeHeaders:=True)
    'The temporary sheet contains the dictionary for the exports
  ElseIf (scope = ExportAll) Then 'Here export is export for migration
    dictData.FromExcelRange dict.DataRange(includeHeaders:=True)
  End If

  BusyApp 'Filetering can return back calculations, so important to kill them again
  tempsh.Cells.Clear
  dictData.ToExcelRange tempsh.Cells(1, 1)

  'Sort the dictionary on exports and also on sheet names
  Set expdict = ExportDictionary()

  'the export dictionary datarange is used for sorting, and also for preparing the dictionary
  Set expDictRng = expdict.DataRange(includeHeaders:=True)

  If scope <> ExportAll Then

    'Sort the dictionary range on table names and then on export name
    expDictRng.Sort key1:=expdict.DataRange("table name", includeHeaders:=True), _
                    key2:=expdict.DataRange(expName, includeHeaders:=True), Header:=xlYes
  End If

  'Add the tags for preparation of the dictionary (the dictionary is already prepared in the linelist)
  expDictRng.Cells(expDictRng.Rows.Count + 1, 1).Font.color = vbBlue

  'Add the required worksheets after sorting
  Set sheetNamesList = expdict.UniqueValues("sheet name")

  'Add the metadata and other worksheets
  prevSheetName = firstSheetName
  
  If (expobj.Value("export metadata sheets", scope) = "yes") Then
    'In the metadata sheet add tags for show/hide or for 
    AddMetaData scope:=scope, _ 
                includeShowHide:=True, _ 
                keepLabels:=True
    prevSheetName = METADATASHEET
  ElseIf (scope = ExportAll) Then
    AddMetaData scope:=scope, _ 
                includeShowHide:=includeShowHide, _ 
                keepLabels:=keepLabels
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

  'Add analysis if required
  If withAnalysis Then AddAnalysis prevSheetName, outwb, inpwb

  'Delete the first sheet of the workbook
  outwb.Worksheets(firstSheetName).Delete
End Sub

'Export Analysis worksheet
Private Sub AddAnalysis(ByVal prevSheet As String, _
                        ByVal outwb As Workbook, ByVal inpwb As Workbook)

  Dim sheetNamesList As BetterArray
  Dim scopeNamesList As BetterArray
  Dim scope As Byte
  Dim counter As Long
  Dim lltrads As ILLTranslations
  Dim prevSheetName As String
  Dim namObj As IAnaTabIds
  Dim outsh As Worksheet
  Dim sheetName As String

  Set sheetNamesList = New BetterArray
  Set scopeNamesList = New BetterArray
  Set lltrads = LLTradsObject()
  prevSheetName = prevSheet

  'Add Analysis Worksheets for analysis Exports
  sheetNamesList.Push lltrads.Value("uasheet"), lltrads.Value("tssheet"), _
                     lltrads.Value("spsheet"), lltrads.Value("sptsheet")
  
  scopeNamesList.Push AnalysisIdsScopeNormal,  AnalysisIdsScopeTimeSeries, _ 
                      AnalysisIdsScopeSpatial, AnalysisIdsScopeSpatioTemporal

  Set namObj = AnaTabIds.Create(inpwb.Worksheets(ANANAMESSHEET))
  
  For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound
    sheetName =  sheetNamesList.Item(counter)
    scope = scopeNamesList.Item(counter)
    With outwb
      .Worksheets.Add(after:=.Worksheets(prevSheetName)).Name = sheetName
      prevSheetName = sheetName
      Set outsh = .Worksheets(prevSheetName)
    End With

    On Error Resume Next
      Application.CutCopyMode = False
      inpwb.Worksheets(prevSheetName).Cells.Copy
      outsh.Cells.PasteSpecial xlPasteValues
      DoEvents
      Application.CutCopyMode = False
      inpwb.Worksheets(prevSheetName).Cells.Copy
      outsh.Cells.PasteSpecial xlPasteFormats
      DoEvents
      RemoveGridLines outsh
      DoEvents
      namObj.WriteGraphs outsh, scope
    On Error GoTo 0
  Next
End Sub

Private Sub AddMetaData(ByVal scope As Integer, ByVal includeShowHide As Boolean, ByVal keepLabels As Boolean)

  Dim outwb As Workbook
  Dim inpwb As Workbook
  Dim csData As IDataSheet
  Dim sheetNamesList As BetterArray
  Dim counter As Long
  Dim sh As Worksheet
  Dim varRng As Range
  Dim expName As String
  Dim dict As ILLdictionary
  Dim outdict As ILLdictionary
  Dim outDictRng As Range
  Dim sheetName As String
  Dim exportColumn As String
  'Elements for adding custom choices to exports
  Dim choi As ILLchoice
  Dim Lo As listObject
  Dim LoValues As BetterArray
  Dim lltrads As ILLTranslations
  Dim choiName As String
  
  Set sheetNamesList = New BetterArray
  Set inpwb = InputWkb()
  Set outwb = outputWkb()

  'Export all the used worksheets choices, metadata and translations
  sheetNamesList.Push METADATASHEET, CHOICESSHEET, DICTTRANSLATIONSHEET

  For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound
    sheetName = sheetNamesList.Item(counter)
    Set sh = inpwb.Worksheets(sheetName)
    Set csData = DataSheet.Create(sh, 1, 1)
    'Then you can just use the export sub from the datasheet
    csData.Export toWkb:=outwb, Hide:=xlSheetVisible
  Next

  'Export dictionary of scope of the export
  Set sh = inpwb.Worksheets(DICTIONARYSHEET)
  Set dict = LLdictionary.Create(sh, 1, 1, NumberOfExports())

  exportColumn = IIf(scope <> ExportAll, "export " & scope, "__all__")
  dict.Export toWkb:=outwb, exportType:=exportColumn, _
              addListObject:=False, Hide:=xlSheetVisible

  'This Tag shows that the exported dictionary is prepared (important for eventual imports process)
  Set outdict = LLdictionary.Create(outwb.Worksheets(DICTIONARYSHEET), 1, 1, NumberOfExports())
  Set outDictRng = outdict.DataRange()
   'Add preparation steps to output dictionary (to be able to use variables)
  outDictRng.Cells(outDictRng.Rows.Count + 1, 1).Font.color = vbBlue

  'Add the dropdown choices for custom choices to export
  Set choi = LLchoice.Create(outwb.Worksheets(CHOICESSHEET), 1, 1)
  Set lltrads = LLTradsObject()
  Set LoValues = New BetterArray
  'Add each custom dropdown to the choice sheet
  On Error Resume Next
  Set sh = inpwb.Worksheets(lltrads.Value("customdropdown"))
  For Each Lo In sh.ListObjects()
    'Get the categories of the choice
    LoValues.FromExcelRange Lo.Range.Cells(2, 1), DetectLastRow:=True, _
            DetectLastColumn:=False
    choiName = "__choice_custom_" & Lo.Range.Cells(1, 1).Value
    choi.AddChoice choiName, LoValues
  Next
  On Error GoTo 0

  'Add informations on metadatasheet
  Set sh = outwb.Worksheets(METADATASHEET)
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
  
  'Keep editable labels on import (default is no)
  varRng.Cells(varRng.Rows.Count + 3, 1).Value = "editable_labels"
  varRng.Cells(varRng.Rows.Count + 3, 2).Value = "no update on import"

  'keep show/hide on import (default is no)
  varRng.Cells(varRng.Rows.Count + 4, 1).Value = "show_hide"
  varRng.Cells(varRng.Rows.Count + 4, 2).Value = "no update on import"

  'keep Labels when importing the file
  If keepLabels Then  varRng.Cells(varRng.Rows.Count + 3, 2).Value = "update on import"

  'include the show/hide on import
  If includeShowHide Then
    'Add the tab on metadata sheet
    varRng.Cells(varRng.Rows.Count + 4, 2).Value = "update on import"
    AddShowHide outwb:=outwb, inpwb:=inpwb
  End If

End Sub

Private Sub AddShowHide(ByVal outwb As Workbook, ByVal inpwb As Workbook)

  Const SHOWHIDEPREF As String = "ShowHideTable_"

  'Input/output worksheets used when moving ListObjects from showhide worksheet
  Dim inpsh As Worksheet
  Dim outsh As Worksheet
  Dim sh As Worksheet
  Dim inpRng As Range
  Dim outRng As Range
  Dim Lo As ListObject
  Dim colWidth As Long
  Dim outdict As ILLdictionary
  Dim counter As Long
  Dim cellRng As Range 'move between headers of the show/hide output
  Dim varName As String 'variable Name on the show/hide worksheet
  Dim sheetName As String 'sheet Name on the show/hide worksheet
  Dim outvars As ILLVariables
  Dim statusValue As String
  Dim colIndex As Long
  Dim tabName As String
  Dim sheetInfo As String
  Dim currDict As ILLdictionary
  Dim tradsmess As ITranslation
  Dim showHideObj As ILLShowHide


  'Step 1: Initialize all the show/hide tables in  the current workbook -------------------------

  'Initialize all show/hide elements in the current workbook
  Set tradsmess = MessagesTrads()
  Set currDict = LLdictionary.Create(inpwb.Worksheets(DICTIONARYSHEET), 1, 1, NumberOfExports())
  
  'Initialize show/hide tables
  For Each sh In inpwb.Worksheets
    If (sh.Cells(1, 3).Value = "HList") Or _ 
       (sh.Cells(1, 3).Value = "HList Print") Or _ 
       (sh.Cells(1, 3).Value = "VList") Then
      Set showHideObj = LLShowHide.Create(tradsmess, currDict, sh)
      showHideObj.InitializeTable
    End If
  Next

  'Step 2: Add all the show/hide elements to the export ------------------------------------

  'Add the show/hide worksheet, and initialize elements
  With outwb
    .Worksheets.Add(.Worksheets(.Worksheets.Count)).Name = SHOWHIDESHEET
    Set inpsh = inpwb.Worksheets(SHOWHIDESHEET)
    Set outsh = .Worksheets(SHOWHIDESHEET)
    Set outdict = LLdictionary.Create(.Worksheets(DICTIONARYSHEET), 1, 1, NumberOfExports())
  End With
  
  'output variables (operate on dictionary in output workbook)
  Set outvars = LLVariables.Create(outdict)  
  
  'This is the first cell of the show/hide worksheet outsh is the show/hide worksheet
  Set cellRng = outsh.Cells(1, 1)
  
  'Import all existing variables of all ListObjects from the show/hide worksheet
  'in the input workbook inpsh is the show/hide worksheet
  For Each Lo In inpsh.ListObjects

    Set inpRng = Lo.Range
    cellRng.Value = "variable name"
    cellRng.Cells(1, 2).Value = "status"
    cellRng.Cells(1, 3).Value = "column width"

    'First cell of the listObject  
    Set cellRng = cellRng.Offset(1)

    For counter = 2 To inpRng.Rows.Count

      varName = inpRng.Cells(counter, 2).Value
      statusValue = inpRng.Cells(counter, 3).Value
      
      If outvars.Contains(varName) Then
        tabName = Replace(Lo.Name, SHOWHIDEPREF, vbNullString)
        sheetName = outvars.Value(colName:="sheet name", varName:=varName)
        'could be either on print or non data worksheet
        If (InStr(1, tabName, "pr") = 1) Then sheetName = "print_" & sheetName
        
        On Error Resume Next
        colWidth = 0
        colIndex = CLng(outvars.Value(colName:="column index", varName:=varName))
        sheetInfo = outvars.Value(colName:="sheet type", varName:=varName)
        If sheetInfo = "vlist1D" Then
          colWidth = inpwb.Worksheets(sheetName).Rows(colIndex).RowHeight
        Else
          colWidth = inpwb.Worksheets(sheetName).Columns(colIndex).ColumnWidth
        End If

        On Error GoTo 0
        'add variable Name
        cellRng.Value = varName
        'add status
        cellRng.Cells(1, 2).Value = statusValue 
        'write column width
        cellRng.Cells(1, 3).Value = colWidth
        'move to next line
        Set cellRng = cellRng.Offset(1)
      End If 
    Next

    'outRng is the range of the ListObject on output worksheet, this is done only when
    'you are done with the variables.
    With outsh
      Set outRng = .Range(.Cells(1, cellRng.Column), .Cells((cellRng.Row - 1), cellRng.Column + 2))
      'Add the listObject Name
      .ListObjects.Add(xlSrcRange, outRng, , xlYes).Name = Lo.Name
      .ListObjects(Lo.Name).TableStyle = "TableStyleLight4"
      Set cellRng = .Cells(1, cellRng.Column + 5)
    End With
  Next

  'Auto fit columns for the show/hide worksheet
  On Error Resume Next
  outsh.Cells.EntireColumn.AutoFit
  On Error GoTo 0
End Sub

'Add Data to export
Private Sub AddData(Optional ByVal useFilters As Boolean = False)

  Dim expdict As ILLdictionary
  Dim headersList As BetterArray
  Dim labelsList As BetterArray
  Dim sheetNamesList As BetterArray
  Dim counter As Long
  Dim outwb As Workbook
  Dim csTab As ICustomTable
  Dim Lo As listObject
  Dim sheetName As String
  Dim useLabels As Boolean
  Dim expobj As ILLExport
  Dim scope As Integer
  Dim inpsh As Worksheet
  Dim inpwb As Workbook
  Dim outsh As Worksheet
  Dim pass As ILLPasswords

  Set expdict = ExportDictionary()
  Set outwb = outputWkb()
  Set inpwb = InputWkb()
  Set sheetNamesList = expdict.UniqueValues("sheet name")
  Set pass = LLPasswords.Create(inpwb.Worksheets(PASSWORDSHEET))
  Set expobj = ExportObject()
  scope = ActualScope()

  useLabels = (expobj.Value("header format", scope) = "variables labels")

  For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound

    'Find the sheet
    sheetName = sheetNamesList.Item(counter)
    Set inpsh = inpwb.Worksheets(sheetName)
    Set headersList = expdict.Data.FilterData("sheet name", sheetName, "variable name")
    Set labelsList = expdict.Data.FilterData("sheet name", sheetName, "main label")
    Set outsh = outwb.Worksheets(sheetName)

    'Test if the sheet is hlist or vlist
    If (inpsh.Cells(1, 3).Value = "HList") Then

      'Test if I should use filtered Data or actual data
      If useFilters Then
        'ListObject on printed sheet the
        Set Lo = inpwb.Worksheets(inpsh.Cells(1, 5).Value).ListObjects(1)
      Else
        Set Lo = inpsh.ListObjects(1)
      End If

      'Define the custom table to export
      Set csTab = CustomTable.Create(Lo)

      'Unprotect the worksheet for export (important to unhide columns)
      pass.UnProtect inpsh.Name

      'Export the custom table to the worksheet
      csTab.Export outsh, headersList

      'Protect back the worksheet
      pass.Protect inpsh.Name

      'Add labels if required
      If useLabels Then labelsList.ToExcelRange outsh.Cells(1, 1), TransposeValues:=True

    ElseIf (inpsh.Cells(1, 3).Value = "VList") Then

      'Add VlistData
      AddVlistData sheetName, headersList, useLabels

    End If

    'Format the cells of the output worksheet
    With outsh
      On Error Resume Next
      .Cells.EntireRow.RowHeight = 14
      .Cells.EntireColumn.ColumnWidth = 35
      .Cells.Font.Name = "Consolas"
      .Cells.Font.Size = 9
      .Rows(1).Font.Size = 10
      .Rows(1).RowHeight = 25
      .Rows(1).Interior.Color = RGB(188, 212, 230)
      .Rows(1).VerticalAlignment = xlVAlignCenter
      .Rows(1).Font.Bold = True
      .Cells.EntireColumn.AutoFit
      On Error GoTo 0
    End With
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
  Set outwb = outputWkb()

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
      If useLabels Then varName = inpsh.Range(varName).Offset(, -1).Value
    On Error GoTo 0

    'Write values to output sheet. I don't know if it suceeded in getting
    'the value from the input sheet, that is why I test it before proceeding,
    'otherwise do nothing

    outsh.Cells(headerCounter, 1).Value = varName
    outsh.Cells(headerCounter, 2).Value = varValue
    varValue = vbNullString
    headerCounter = headerCounter + 1
    
  Next

End Sub

'--- Folder management for exports

Private Property Get HasFolder() As Boolean
  HasFolder = this.folderSet
End Property

Private Property Let HasFolder(ByVal folderSet As Boolean)
  If Not this.folderSet Then this.folderSet = folderSet
End Property


Private Property Get ExportFolder() As String
  Dim io As IOSFiles

  'Test if the folder has been set, in other case set it.
  If Not HasFolder() Then
    Set io = OSFiles.Create()
    
    io.LoadFolder
    HasFolder = True
    If Not io.HasValidFolder() Then
      this.folderPath = vbNullString
    Else
      this.folderPath = io.Folder()
    End If
  
  End If

  ExportFolder = this.folderPath
End Property

'Save the output workbook with corresponding elements to export
Private Sub Save(ByVal trads As ITranslation, _
                Optional ByVal useFilters As Boolean = False, _ 
                Optional ByVal includeShowHide As Boolean = False, _ 
                Optional ByVal keepLabels As Boolean = False)

  Dim outwb As Workbook
  Dim inpwb As Workbook
  Dim dict As ILLdictionary
  Dim pass As ILLPasswords
  Dim scope As Integer
  Dim fileExtension As String
  Dim fileName As String
  Dim filePath As String
  Dim expobj As ILLExport
  Dim fileFormat As Long
  Dim addPassword As Boolean
  Dim filePassword As String
  Dim folderPath As String
  Dim withAnalysis As Boolean
  Dim prevSheetName As String

  BusyApp
  folderPath = ExportFolder()
  If folderPath = vbNullString Then Exit Sub
  
  'In case there is a \ or / at the end of the folder path, remove it because
  'we will need to paste the file name further. On local disks like D:\ this should
  'avoid ending up with D:\\... as file Path.
  If InStrRev(folderPath, Application.PathSeparator, _
     Len(folderPath)) = Len(folderPath) Then
    folderPath = Left(folderPath, (Len(folderPath) - 1))
  End If
  

  scope = ActualScope()
  Set expobj = ExportObject()
  Set inpwb = InputWkb()
  Set outwb = outputWkb()
  
  Set dict = LLdictionary.Create(inpwb.Worksheets(DICTIONARYSHEET), 1, 1, NumberOfExports())
  Set pass = LLPasswords.Create(inpwb.Worksheets(PASSWORDSHEET))
  
  'Get the fileName, filePath and fileExtension of the export
  fileName = expobj.ExportFileName(scope, dict, pass)
  filePath = folderPath & Application.PathSeparator & fileName
  fileExtension = expobj.Value(colName:="file format", exportNumber:=scope)
  
  If fileExtension = vbNullString Then fileExtension = "xlsx" 'Default extension is xlsx
  
  addPassword = (expobj.Value(colName:="password", exportNumber:=scope) = "yes")

  If addPassword Then filePassword = pass.Value("privatekey")

  'The file format when you save
  fileFormat = IIf(fileExtension = "xlsb", xlExcel12, xlOpenXMLWorkbook)
  filePath = filePath & "." & fileExtension
  
  BusyApp

  'Prepare the output workbook (create required sheets, add metadata)
  If (scope <> ExportAna) Then
    withAnalysis = (expobj.Value("export analyses sheets", scope) = "yes")
    Prepare withAnalysis, includeShowHide, keepLabels 
    'Add the Data to export worksheet
    AddData useFilters
  Else
    prevSheetName = outwb.Worksheets(1).Name
    AddAnalysis prevSheetName, outwb, inpwb
  End If
  
  'Save the file to the corresponding path using corresponding password

  If addPassword Then
    'save with password
    outwb.SaveAs fileName:=filePath, _
                 fileFormat:=fileFormat, _
                 CreateBackup:=False, _
                 Password:=filePassword, _
                ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    'Display the private password!
    MsgBox trads.TranslatedValue("MSG_FileSaved") & chr(10) & _
           trads.TranslatedValue("MSG_Password") & filePassword
  Else
    'save with password
    outwb.SaveAs fileName:=filePath, _
    fileFormat:=fileFormat, _
    CreateBackup:=False, _
    ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges

    'Show message for only the other exports (not one for migration)
    If scope <> ExportAll Then
      'Display message that there is no password!
      MsgBox trads.TranslatedValue("MSG_FileSaved") & chr(10) & _
            trads.TranslatedValue("MSG_NoPassword")
    End If

  End If

  outwb.Close savechanges:=False
End Sub

'Save the geobase
Private Sub SaveGeo(ByVal geoObj As ILLGeo, _
                    Optional ByVal onlyHistoric As Boolean = False)

  Dim outwb As Workbook
  Dim folderPath As String
  Dim filePath As String
  Dim geobaseName As String

  folderPath = ExportFolder()
  If folderPath = vbNullString Then Exit Sub
  
  BusyApp

  If onlyHistoric Then
    Set outwb = OutputGeoHistoricWkb()
  Else
    Set outwb = OutputGeoWkb()
  End If

  geobaseName = geoObj.GeoNames("geobase")
  
  'Exported geobase name
  If (geobaseName = vbNullString Or geobaseName = "geobase") Then geobaseName = "OUTBREAK-TOOLS-GEOBASE-UNKNOWN"

  geobaseName = geobaseName & "_" & "export" & Format(Now, "yyyymmdd")
  If onlyHistoric Then geobaseName = geobaseName & "_historic"

  filePath = folderPath & Application.PathSeparator & geobaseName & ".xlsx"

  'Export geo data to the workbook
  geoObj.ExportToFile outwb, onlyHistoric

  'Save the data exported
  outwb.SaveAs fileName:=filePath, _
  CreateBackup:=False, _
  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges

  outwb.Close savechanges:=False
End Sub

'Close everything in case of error
Private Sub CloseAll()
  On Error Resume Next
  this.outwb.Close savechanges:=False
  this.outgeowb.Close savechanges:=False
  this.outgeohistwb.Close savechanges:=False
  On Error GoTo 0
End Sub

'Checking requirements before instanciating the class
Private Sub CheckRequirements(ByVal currwb As Workbook)

  Dim counter As Long
  Dim sheetName As String
  Dim sheetsArray As BetterArray
  Set sheetsArray = New BetterArray

  sheetsArray.Push DICTIONARYSHEET, CHOICESSHEET, DICTTRANSLATIONSHEET, _
                   METADATASHEET, PASSWORDSHEET, TEMPSHEET, EXPORTSHEET, _
                   MESSTRANSLATIONSHEET, SHOWHIDESHEET

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

'Remove gridlines on exported analysis sheets
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

'Errors managements
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, _
   "Unable to create an instance of the class " & CLASSNAME & ": " & errorMessage
End Sub

'Exposed Methods -------------------------------------------------


Private Sub IOutputSpecs_Prepare()
  Prepare
End Sub

'Add Dictionary, Translations or Analysis
Private Sub IOutputSpecs_AddData()
  AddData
End Sub

Private Sub IOutputSpecs_Save(ByVal trads As ITranslation, _ 
                              Optional ByVal useFilters As Boolean = False, _ 
                              Optional ByVal includeShowHide As Boolean = False, _ 
                              Optional ByVal keepLabels As Boolean = False)
  Save trads, useFilters, includeShowHide, keepLabels
End Sub

'Get the exports folderPath
Private Property Get IOutputSpecs_ExportFolder() As String
  IOutputSpecs_ExportFolder = ExportFolder()
End Property

Private Sub IOutputSpecs_SaveGeo(ByVal geoObj As ILLGeo, Optional ByVal onlyHistoric As Boolean = False)
  SaveGeo geoObj, onlyHistoric
End Sub

Private Sub IOutputSpecs_CloseAll()
  CloseAll
End Sub
