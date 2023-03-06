VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "AnalysisOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Module to put all together for the analysis"

'@Folder("Analysis")
'@ModuleDescription("Build And format the analysis")
'@IgnoreModule

Option Explicit
Implements IAnalysisOutput

'Name of differents 
Private Const UPDATEFILTERCOMMAND As String = "UpdateFilterTables"
Private Const SHPFILTERNAME As String = "SHP_Filter"
Private Const GOTOSECCODE As String = "go_to_section" 'Goto section code
Private Const UPDATESPCOMMAND As String = "UpdateSpTables" 'Spatial tables update
Private Const CLASSNAME As String = "AnalysisOutput"



Private Type TAnaOutput
  anaLL As ILinelist
  sh As Worksheet
  sectionsList As BetterArray
  headersList As BetterArray
  graphsList As BetterArray
  goToSection As String
  goToHeader As String
  goToGraph As String
End Type

Private this As TAnaOutput

'specSh is the worksheet where the specifications of the analyses are
'll is the linelist Object
Public Function Create(specSh As Worksheet, ByVal ll As ILinelist) As IAnalysisOutpout
  With New AnalysisOutput
    Set .OutputLL = ll
    Set .Wksh = specSh
    Set Create = .Self()
  End With

End Function

Public Function Self() As IAnalysisOutpout
  Set Self = Me
End Function

Public Property Get OutputLL() As ILinelist
  Set OutputLL = this.anaLL
End Property

Public Property Set OutputLL(ByVal ll As ILinelist)
  Set this.anaLL = ll
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal specSh As Worksheet)
  Set this.sh = specSh
End Property


'Get the linelist and linelistData attributes used when building the linelist

'Translation of messages / other translations
Private Property Get Trads(Optional ByVal scope As TradScopes = TranslationOfMessages) As ITranslation
  Dim ll As ILinelist
  Dim lData As LinelistSpecs
  Set ll = OutputLL()
  Set lData = ll.LinelistData()
  Set Trads = lData.TransObject(scope)
End Property

'Linelist Specs
Private Property Get LinelistData() As ILinelistSpecs
  Dim ll As ILinelist
  Set ll = OutputLL()
  Set LinelistData = ll.LinelistData()
End Property

'Design
Private Property Get DesignFormat() As ILLFormat
  Dim lData As ILinelistSpecs
  Set lData = LinelistData()
  Set DesignFormat = lData.DesignFormat()
End Property

'Translation of shapes and messages
Public Property Get TransObject(ByVal scope As  _ 
                                TradScopes = TranslationOfMessages) As ITranslation
  Dim lData As ILinelistSpecs
  Set lData = LinelistData()
  Set TransObject = lData.TransObject(scope)
End Property

'Output Worksheets (where to write the analysis)
Private Property Get OutputSh(Optional ByVal scope _
                             As AnalysisScope = AnalysisScopeNormal) As Worksheet
  Dim trads As ITranslation
  Dim ll As ILinelist

  Set ll = OutputLL()
  Set trads = TransObject(TranslationOfMessages)

  Select Case scope
  Case AnalysisScopeTimeSeries
    Set OutputSh = trads.TranslatedValue("LLSHEET_TemporalAnalysis")
  Case AnalysisScopeSpatial
    Set OutputSh = trads.TranslatedValue("LLSHEET_SpatialAnalysis")
  Case AnalysisScopeNormal
    Set OutputSh = trads.TranslatedValue("LLSHEET_Analysis")
  Case Else
    Set OutputSh = trads.TranslatedValue("LLSHEET_Analysis")
  End Select
End Property

'Speed Up the application
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Return previous state
Private Sub ReturnState()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
End Sub

'Split the output sheet

Private Sub SplitOutputSheet(Optional ByVal scope As AnalysisScope = AnalysisScopeNormal)
  Dim sh As Worksheet
  Dim rowSplit As Long
  'OUtput worksheet
  Set sh = OutputSh(scope)
  rowSplit = IIf(scope = AnalysisScopeTimeSeries, 3, 2)
  On Error Resume Next
  'Activating the worksheet restores all the application
  'status to true, remember to recall the busy app just
  'after and before
  BusyApp
  sh.Activate
  BusyApp

  With Application.ActiveWindow
    .SplitRow = rowSplit
    'Split on the third column for Time series only
    If scope = AnalysisScopeTimeSeries Then .SplitColumn = 3
    'For spatial or time series, reduce the zoom for the worksheeet
    If scope <> AnalysisScopeNormal Then .Zoom = 65
    .FreezePanes = True
  End With
  On Error GoTo 0
End Sub



'Add the command buttons to a worksheet
Private Sub AddCommands(Optional ByVal scope _ 
                       As AnalysisScope = AnalysisScopeNormal)
  Dim cellRng As Range
  Dim commandName As String
  Dim shapeLabel As String
  Dim design As ILLFormat
  Dim but As IButtons
  Dim sh As Worksheet
  Dim trads As ITranslation

  Set design = DesignFormat()
  Set sh = OutputSh(scope)
  Set trads = TransObject(TranslationOfShapes)
  Set cellRng = sh.Cells(2, 2)

  'Update shapeLabel and commandName depending on the scope of the butto
  commandName = IIf(scope = AnalysisScopeSpatial, UPDATESPCOMMAND, UPDATEFILTERCOMMAND)
  shapeLabel = trads.TranslatedValue(SHPFILTERNAME)
  Set but = Buttons.Create(outRng:=cellRng, codeName:=SHPFILTERNAME, _ 
                           scope:=ButtonScopeSmall)
  but.Add actionCommand:=commandName, shapeLabel:=shapeLabel
  but.Format design
End Sub

'Initialize the output worksheet (format, values)
Private Sub PrepareSheet(Optional ByVal scope As AnalysisScope _ 
                         = AnalysisScopeNormal)

  Dim sh As Worksheet
  Dim cellRng As Range
  Dim sheetInfo As String
  Dim design As ILLFormat

  Select Case scope
  Case AnalysisScopeTimeSeries
    sheetInfo = "TS-Analysis"
  Case AnalysisScopeSpatial
    sheetInfo = "SP-Analysis"
  Case AnalysisScopeNormal
    sheetInfo = "Uni-Bi-Analysis"
  End Select

  Set sh = OutputSh(scope)
  Set design = DesignFormat()
  Set cellRng = sh.Cells(1, 3)

  'Add the info of a worksheet
  cellRng.Value = sheetInfo
  design.ApplyFormat cellRng, scope:=LinelistHiddenCell

  'Add warning if the cells are filtered
  Set cellRng = outshCells(2, 5)
  cellRng.Formula = "= ComputedOnFiltered()"
  design.ApplyFormat cellRng, scope:=LinelistWarning
End Sub


'Push section, graphs or headers to the tables
Private Sub Push(pushedValue As String, Optional ByVal scope As String = "section")
  'Initialize arrays of sections, headers and graphs
  InitializeArrays
  Select case scope
  Case "section"
    this.sectiionsList.Push pushedValue
  Case "header"
    this.headersList.Push pushedValue
  Case "graph"
    this.graphsList.Push pushedValue
  End Select
End Sub

'Add a GoTo (could be section or Header, or Graphs)

Private Sub InitializeArrays()
    Dim trads As ITranslation

    Set trads = TransObject()
    this.goToSection = trads.TranslatedValue("MSG_GoToSec")
    this.goToHeader = trads.TranslatedValue("MSG_GoToHead")
    this.goToGraph = trads.TranslatedValue("MSG_GoToGraph")

    'initialize section list, graphList and headersList
    If this.sectionsList Is Nothing Then Set this.sectionsList = New BetterArray
    If this.headersList Is Nothing Then Set this.headersList = New BetterArray
    If this.graphsList Is Nothing Then Set this.graphsList = New BetterArray
End Sub

'Get the goToSection, header, graph value
Private Function GoToLabel(optional ByVal scope As String = "section") As String

  GoToLabel = Switch(scope = "section", this.goToSection, _ 
                     scope = "graph", this.goToGraph, _ 
                     scope = "header", this.goToHeader, _ 
                     True, this.goToSection)
End Function

Private Sub AddGoTo(cellRng As Range, _ 
                    Optional ByVal goToPrefix As String = "ua", _ 
                    Optional ByVal scope As String = "section")
  
  Dim drop As IDropdownLists
  Dim ll As ILinelist
  Dim dropArray As BetterArray
  Dim trads As ITranslation
  Dim dropName As String
  Dim design As ILLFormat

  Set dropArray = New BetterArray
  Select Case scope
  Case "section"
    Set dropArray = this.sectionsList.Clone()
  Case "graph"
    Set dropArray = this.headersList.Clone()
  Case "header"
    Set dropArray = this.graphsList.Clone()
  End Select

  'Add the goTo Section to the range
  Set ll = OutputLL()
  Set drop = ll.Dropdown()
  Set design = DesignFormat()
  Set trads = TransObject()

  'dropdown name
  dropName = goToPrefix & "gotosection"
  drop.Add dropArray, dropName
  drop.SetValidation cellrng, dropName

  'Apply the dropdown format to the cell
  design.ApplyFormat cellRng, scope:=LinelistSelectionDropdown
  
  'Add a Name and value to the cell Range
  cellRng.Value = GoToLabel(scope)
  cellRng.Name = goToPrefix & GOTOSECCODE

  'Clear the 
  Select Case scope
  Case "section"
    this.sectionsList.Clear()
  Case "graph"
    this.headersList.Clear()
  Case "header"
    this.graphsList.Clear()
  End Select
End Sub


'Write tables of one section of univariate or bivariate
Private Sub WriteTable(ByVal LoName As String,  _
                       Optional ByVal addGraphs As Boolean = False)
  Dim sh As Worksheet
  Dim outsh As Worksheet
  Dim tabRng As Range
  Dim hRng As Range
  Dim counter As Long
  Dim graphCounter As Long
  Dim tabl As ICrossTable
  Dim tabform As ICrossTableFormul
  Dim gr As IGraphs
  Dim Lo As ListObjects
  Dim specs As ITablesSpecs
  Dim lData As LinelistSpecs
  Dim goToName As String

  'Initialize the arrays
  InitializeArrays

  'Worksheet of the analyses
  Set sh = Wksh()
  'Output Worksheet
  Set outsh = OutputSh(scope)
  Set Lo = sh.ListObjects(LoName)
  Set lData = LinelistData()

  If Not (Lo.DataBodyRange Is Nothing) Then
    Set hRng = Lo.HeaderRowRange
      For counter = 1 To Lo.ListRows.Count
      Set tabRng = Lo.ListRows(counter).Range
      Set specs = TablesSpecs.Create(hRng, tabRng, lData)
      
      If specs.ValidTable() Then
        If specs.isNewSection Then Push 
      
      
      
      End If


      Next

  End If
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, className, "Error : " & errorMessage
End Sub
