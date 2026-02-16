

Attribute VB_Name = "AnalysisTestFixture"
Attribute VB_Description = "Shared helpers for Analysis tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Shared helpers for Analysis tests")

'@section Constants
'===============================================================================

Private Const ANALYSISTESTSHEET As String = "AnalysisFixture"
Private Const ANALYSISTRANSLATIONSHEET As String = "AnalysisTranslation"
Private Const ANALYSISTRANSLATIONTABLE As String = "tblTranslation"

Private Const TAB_GLOBAL_SUMMARY As String = "Tab_global_summary"
Private Const TAB_UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const TAB_BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TAB_TIME_SERIES As String = "Tab_TimeSeries_Analysis"
Private Const TAB_GRAPH_TIME_SERIES As String = "Tab_Graph_TimeSeries"
Private Const TAB_GRAPH_TITLE As String = "Tab_Label_TSGraph"
Private Const TAB_SPATIAL As String = "Tab_Spatial_Analysis"
Private Const TAB_SPATIO_TEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const TAB_SPATIO_TEMPORAL_SPECS As String = "Tab_SpatioTemporal_Specs"

'@section Fixture Data
'===============================================================================

Public Function AnalysisHeaders() As Variant
    AnalysisHeaders = Array("Section", "Table Title", "Summary function")
End Function

Public Function AnalysisRows(ByVal sectionValue As String) As Variant
    AnalysisRows = Array(Array(sectionValue, "Goodbye", "Summary"))
End Function

Private Function TranslationHeaders() As Variant
    TranslationHeaders = Array("tag", "English", "French")
End Function

Private Function TranslationRows() As Variant
    TranslationRows = Array( _
        Array("greeting", "Hello", "Bonjour"), _
        Array("farewell", "Goodbye", "Au revoir"))
End Function


'@section Worksheet Builders
'===============================================================================

Public Sub ClearTestAnalysisSheets()
    DeleteWorksheet ANALYSISTRANSLATIONSHEET
    DeleteWorksheet ANALYSISTESTSHEET  
End Sub

Public Function BuildAnalysisTable(ByVal hostSheet As Worksheet, ByVal sectionValue As String) As ListObject
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim tableRange As Range
    Dim analysisTable As ListObject

    headerMatrix = RowsToMatrix(Array(AnalysisHeaders()))
    dataMatrix = RowsToMatrix(AnalysisRows(sectionValue))

    WriteMatrix hostSheet.Cells(3, 1), headerMatrix
    WriteMatrix hostSheet.Cells(4, 1), dataMatrix

    Set tableRange = hostSheet.Range("A3").Resize( _
                      UBound(dataMatrix, 1) + UBound(headerMatrix, 1), _
                      UBound(headerMatrix, 2))

    On Error Resume Next
        hostSheet.ListObjects(TAB_GLOBAL_SUMMARY).Delete
    On Error GoTo 0

    Set analysisTable = hostSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                  Source:=tableRange, _
                                                  XlListObjectHasHeaders:=xlYes)
    analysisTable.Name = TAB_GLOBAL_SUMMARY

    Set BuildAnalysisTable = analysisTable
End Function

Public Function PrepareAnalysisSheet(Optional ByVal sectionValue As String = "Initial Section") As Worksheet
    Dim hostSheet As Worksheet
    Set hostSheet = EnsureWorksheet(ANALYSISTESTSHEET, clearSheet:=True, visibility:=xlSheetHidden)

    hostSheet.Cells(1, 1).Value = "Add or remove rows of Global Summary"
    BuildAnalysisTable hostSheet, sectionValue

    Set PrepareAnalysisSheet = hostSheet
End Function

Public Function AnalysisTable(ByVal tag As String,  _
                              Optional ByVal impSheet As Worksheet, _
                              Optional ByVal headerInstruction As String = "Add or remove rows of Global Summary")  _ 
                              As ListObject

    Dim hostSheet As Worksheet
    Dim loName As String

    If impSheet Is Nothing Then
        Set hostSheet = PrepareFullAnalysisWorksheet(headerInstruction)
    Else
        Set hostSheet = impSheet
    End If

    Select Case tag
        Case "global summary": loName = TAB_GLOBAL_SUMMARY
        Case "univariate analysis": loName = TAB_UNIVARIATE
        Case "bivariate analysis": loName = TAB_BIVARIATE
        Case "time series analysis": loName = TAB_TIME_SERIES
        Case "labels for time series graphs": loName = TAB_GRAPH_TITLE
        Case "graph on time series": loName = TAB_GRAPH_TIME_SERIES
        Case "spatial analysis": loName = TAB_SPATIAL
        Case "spatio-temporal specifications": loName = TAB_SPATIO_TEMPORAL_SPECS
        Case "spatio-temporal analysis": loName = TAB_SPATIO_TEMPORAL
        Case Else
            loName = TAB_GLOBAL_SUMMARY
    End Select

    Set AnalysisTable = hostSheet.ListObjects(loName)
End Function

Public Function PrepareFullAnalysisWorksheet(Optional ByVal headerInstruction As String = "Add or remove rows of Global Summary") As Worksheet

    Dim hostSheet As Worksheet
    Dim nextRow As Long
    Dim spatioTemporalRows As Variant

    Set hostSheet = EnsureWorksheet(ANALYSISTESTSHEET)
    ClearWorksheet hostSheet

    hostSheet.Cells(1, 1).Value = headerInstruction

    nextRow = 3
    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_GLOBAL_SUMMARY, _
                               AnalysisHeaders(), _
                               Array(Array("Initial Section", "Goodbye", "Summary"), _
                                     Array("Initial Section", "Hello", "Count"), _
                                     Array("Second Section", "World", "Percentage")))
    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_UNIVARIATE, _
                               AnalysisHeaders(), Array(Array("Univariate Section", "Univariate Title", "Summary Uni")))
    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_BIVARIATE, _
                               AnalysisHeaders(), Array(Array("Bivariate Section", "Bivariate Title", "Summary Bi")))
    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_TIME_SERIES, _
                               Array("Series ID", "Table order", "Label"), _
                               Array(Array("Series 1", 2, "Alpha")))
    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_GRAPH_TIME_SERIES, _
                               Array("Graph ID", "Section", "Table Title", "Summary label", "Choices"), _
                               Array(Array("Graph 5", "Section B", "Title B", "Summary B", "Choice B"), _
                                     Array("Graph 2", "Section A", "Title A", "Summary A", "Choice A"), _ 
                                     Array("Graph 3", "Section B", "Title C", "Summary C", "Choice C")))
    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_GRAPH_TITLE, _
                               Array("Graph ID", "Graph Title"), _
                               Array(Array("Graph 5", "Graph Title B")))
    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_SPATIAL, _
                               Array("Section", "Label", "Summary label", "Choices"), _
                               Array(Array("Spatial Section", "Spatial Label", "Spatial Summary", "Spatial Choice")))

    spatioTemporalRows = Array( _
        Array("Region A", "Label A", "Choice A", "Graph Title A"), _
        Array("Region B", "Label B", "Choice B", "Graph Title B"), _
        Array("Region C", "Label C", "Choice C", "Graph Title C"), _
        Array("Region A", "Label D", "Choice D", "Graph title D"), _
        Array(Empty, Empty, Empty, Empty))

    nextRow = AddAnalysisTable(hostSheet, nextRow, TAB_SPATIO_TEMPORAL, _
                               Array("Section", "Label", "Choices", "Graph Title"), _
                               spatioTemporalRows)
    Call AddAnalysisTable(hostSheet, nextRow, TAB_SPATIO_TEMPORAL_SPECS, _
                          Array("Section", "Label", "Summary label"), _
                          Array(Array("Specs Section", "Specs Label", "Specs Summary")))

    Set PreparefullAnalysisWorksheet = hostSheet
End Function

Public Function CreateAnalysisTranslator(Optional ByVal language As String = "French") As ITranslationObject
    Dim translationSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim translationTable As ListObject

    Set translationSheet = EnsureWorksheet(ANALYSISTRANSLATIONSHEET)
    ClearWorksheet translationSheet

    headerMatrix = RowsToMatrix(Array(TranslationHeaders()))
    dataMatrix = RowsToMatrix(TranslationRows())

    WriteMatrix translationSheet.Cells(1, 1), headerMatrix
    WriteMatrix translationSheet.Cells(2, 1), dataMatrix

    Set translationTable = translationSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                            Source:=translationSheet.Range("A1").CurrentRegion, _
                                                            XlListObjectHasHeaders:=xlYes)
    translationTable.Name = ANALYSISTRANSLATIONTABLE

    Set CreateAnalysisTranslator = TranslationObject.Create(translationTable, language)
End Function

'@section Internal Helpers
'===============================================================================

Private Function AddAnalysisTable(ByVal hostSheet As Worksheet, _
                                  ByVal startRow As Long, _
                                  ByVal tableName As String, _
                                   headers As Variant, _
                                  Optional  dataRows As Variant) As Long

    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim columnCount As Long
    Dim bottomRow As Long
    Dim tableRange As Range
    Dim listObject As ListObject
    Dim hasData As Boolean

    headerMatrix = RowsToMatrix(Array(headers))
    WriteMatrix hostSheet.Cells(startRow, 1), headerMatrix

    columnCount = UBound(headerMatrix, 2)
    bottomRow = startRow

    hasData = IsArray(dataRows)
    If hasData Then
        On Error Resume Next
            hasData = (UBound(dataRows) >= LBound(dataRows))
        On Error GoTo 0
    End If

    If hasData Then
        dataMatrix = RowsToMatrix(dataRows)
        WriteMatrix hostSheet.Cells(startRow + 1, 1), dataMatrix
        bottomRow = startRow + UBound(dataMatrix, 1)
    End If

    Set tableRange = hostSheet.Range(hostSheet.Cells(startRow, 1), _
                                     hostSheet.Cells(bottomRow, columnCount))

    On Error Resume Next
        hostSheet.ListObjects(tableName).Delete
    On Error GoTo 0

    Set listObject = hostSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                               Source:=tableRange, _
                                               XlListObjectHasHeaders:=xlYes)
    listObject.Name = tableName

    AddAnalysisTable = bottomRow + 8
End Function

