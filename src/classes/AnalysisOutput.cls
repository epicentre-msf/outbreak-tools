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
  Set sh = OUtputSh()
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
Private Sub AddCommands()

  

End Sub


'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, className, "Error : " & errorMessage
End Sub
