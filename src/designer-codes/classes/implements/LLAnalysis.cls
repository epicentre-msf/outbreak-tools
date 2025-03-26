VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Analysis Class"

'@Folder("Analysis")
'@ModuleDescription("Analysis Class")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLAnalysis

Private Type TLLAnalysis
    anash As Worksheet
    lData As ILinelistSpecs
End Type

Private this As TLLAnalysis
Private Const CLASSNAME As String = "LLAnalysis"
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const SPATIOTEMPSPECS As String = "Tab_SpatioTemporal_Specs"

'Exposed methods
Public Function Create(ByVal anaWksh As Worksheet) As ILLAnalysis
    If anaWksh Is Nothing Then ThrowErrorUnableInstanciate "Worksheet object not defined"
    'There is no cheking at instanciation since the worksheet could be empty
    With New LLAnalysis
        Set .Wksh = anaWksh
        Set Create = .Self
    End With
End Function

Public Property Get Wksh() As Worksheet
    Set Wksh = this.anash
End Property

Public Property Set Wksh(ByVal anash As Worksheet)
    Set this.anash = anash
End Property

Public Property Get Self() As ILLAnalysis
    Set Self = Me
End Property

Private Sub CheckTablesRequirements()
    Dim loTable As BetterArray
    Dim counter As Long
    Dim Lo As listObject
    Dim loName As String

    Set loTable = New BetterArray

    loTable.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, _
                 GRAPHTIMESERIES, TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPORAL, _ 
                 SPATIOTEMPSPECS

    For counter = loTable.LowerBound To loTable.UpperBound
        loName = loTable.Items(counter)
        On Error Resume Next
        Set Lo = Wksh.ListObjects(loName)
        On Error GoTo 0
        If Lo Is Nothing Then ThrowError " Listobject " & loName & " does not exist in analysis worksheet "
    Next
End Sub


Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Private Sub Build(ByVal ll As Linelist)
    Dim anaout As IAnalysisOutput
    Dim sh As Worksheet
    Set sh = Wksh()
    Set anaout = AnalysisOutput.Create(sh, ll)
    anaout.WriteAnalysis
End Sub

'Import listobjects in analysis
Private Sub Import(ByVal impWksh As Worksheet)
    Dim anash As Worksheet
    Dim impRng As Range
    Dim actRng As Range
    Dim impAddress As String
    Dim impLo As listObject

    Set anash = Wksh()

    'Test if the sheet is empty
    If impWksh Is Nothing Then Exit Sub          'Add trace for import

    anash.Cells.Clear

    On Error GoTo LoFail
    Set impLo = impWksh.ListObjects(1)
    On Error GoTo 0


    For Each impLo In impWksh.ListObjects

        Set impRng = impLo.Range
        impAddress = impRng.Address
        Set actRng = anash.Range(impAddress)
        actRng.Value = impRng.Value

        actRng.Cells(-1, 1).Value = impRng.Cells(-1, 1).Value
        actRng.Cells(-1, 1).Font.Size = impRng.Cells(-1, 1).Font.Size
        actRng.Cells(-1, 1).Font.color = impRng.Cells(-1, 1).Font.color
        actRng.Cells(-1, 1).Font.Bold = True

        'Define the listobject
        With anash
            .ListObjects.Add(xlSrcRange, actRng, , xlYes).Name = impLo.Name
            .ListObjects(impLo.Name).TableStyle = impLo.TableStyle
        End With
    Next

    'Be sure all the tables exist after import
    CheckTablesRequirements

    anash.Cells(1, 1).Value = "analysis"
    anash.Cells(1, 1).Font.color = vbBlue

    'Sort the analysis worksheet on each one of the tables
    SortAnalysis

    Exit Sub
LoFail:
    'add trace for no list objects
End Sub


Private Sub SortAnalysis()
    Dim sh As Worksheet
    Dim Lo As listObject
    Dim csTab As ICustomTable
    Dim colName As String
    Dim tabl As BetterArray 'Will contain the list of listObjects to sort
    Dim counter As Long
    Dim loName As String
    Dim rowCount As Long

    Set tabl = New BetterArray
    tabl.LowerBound = 1

    tabl.Push UNIVARIATE, BIVARIATE, TIMESERIES, SPATIAL, GRAPHTIMESERIES, _ 
               TITLEGRAPHTIMESERIES, SPATIOTEMPORAL, SPATIOTEMPSPECS

    Set sh = Wksh()

    For counter = 1 To tabl.Length

        loName = tabl.Items(counter)
        Set Lo = sh.ListObjects(loName)
        Set csTab = CustomTable.Create(Lo)

        'Row count is the number of elements in the row that triggers deletion of the
        'entire row of the listObject
        rowCount = Switch(loName = TIMESERIES, 2, loName = GRAPHTIMESERIES, 4, True, 0)
        colName = Switch( _ 
            loName = GRAPHTIMESERIES Or loName = TITLEGRAPHTIMESERIES, "Graph ID", _ 
            loName = SPATIOTEMPORAL, "Section (select)", _ 
            True, "Section" _ 
        )

        csTab.Sort colName, onFirst:=True
        If loName = TIMESERIES Then csTab.Sort "Table order", onFirst:=False

        'Resize the listObject
        csTab.RemoveRows totalCount:=rowCount
    Next
End Sub

'Translate
Private Sub Translate(ByVal TransObject As ITranslation)
    Dim transCols As BetterArray
    Dim anash As Worksheet
    Dim counter As Long
    Dim Lo As listObject
    Dim csTab As ICustomTable
    Dim colRng As Range

    'TransObject is the translation object (from Translation.cls). It can
    'translate ranges, and formula ranges.

    Set transCols = New BetterArray
    Set anash = Wksh()
    transCols.Push "Section", "Table title", "Graph title", "Summary label", _
                   "Choices", "Label", "Section (select)"
 
    'choices is for the time series graphs translation

    For Each Lo In anash.ListObjects
        For counter = transCols.LowerBound To transCols.UpperBound
            Set csTab = CustomTable.Create(Lo)
            Set colRng = csTab.DataRange(transCols.Item(counter), strictSearch:=True) 
            If Not colRng Is Nothing Then TransObject.TranslateRange colRng
            Set colRng = Nothing
        Next

        'Formulas
        Set colRng = csTab.DataRange("Summary function", strictSearch:=True)
        If (Not colRng Is Nothing) Then TransObject.TranslateRange colRng, containsFormula:=True
    Next
End Sub

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, errorMessage
End Sub

'Interface
Public Sub ILLAnalysis_Import(ByVal impWksh As Worksheet)
    Import impWksh
End Sub

Public Sub ILLAnalysis_Translate(ByVal TransObject As ITranslation)
    Translate TransObject
End Sub

Public Property Get ILLAnalysis_Wksh() As Worksheet
    Set ILLAnalysis_Wksh = Wksh()
End Property

Public Sub ILLAnalysis_Build(ll As ILinelist)
    Build ll
End Sub
