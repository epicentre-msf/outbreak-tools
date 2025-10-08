Attribute VB_Name = "AnalysisTestFixture"
Attribute VB_Description = "Shared helpers for Analysis tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Shared helpers for Analysis tests")

'@section Constants
'===============================================================================

Public Const ANALYSISTESTSHEET As String = "AnalysisFixture"
Public Const ANALYSISTRANSLATIONSHEET As String = "AnalysisTranslation"
Public Const ANALYSISTRANSLATIONTABLE As String = "tblTranslation"

Public Const TAB_GLOBAL_SUMMARY As String = "Tab_global_summary"
Public Const TAB_UNIVARIATE As String = "Tab_Univariate_Analysis"
Public Const TAB_BIVARIATE As String = "Tab_Bivariate_Analysis"
Public Const TAB_TIME_SERIES As String = "Tab_TimeSeries_Analysis"
Public Const TAB_GRAPH_TIME_SERIES As String = "Tab_Graph_TimeSeries"
Public Const TAB_GRAPH_TITLE As String = "Tab_Label_TSGraph"
Public Const TAB_SPATIAL As String = "Tab_Spatial_Analysis"
Public Const TAB_SPATIO_TEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Public Const TAB_SPATIO_TEMPORAL_SPECS As String = "Tab_SpatioTemporal_Specs"

'@section Fixture Data
'===============================================================================

Public Function AnalysisHeaders() As Variant
    AnalysisHeaders = Array("Section", "Table Title", "Summary function")
End Function

Public Function AnalysisRows(ByVal sectionValue As String) As Variant
    AnalysisRows = Array(Array(sectionValue, "Goodbye", "=""Summary"""))
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

Public Sub PrepareAnalysisSheet(Optional ByVal sectionValue As String = "Initial Section")
    Dim analysisSheet As Worksheet

    Set analysisSheet = EnsureWorksheet(ANALYSISTESTSHEET)
    ClearWorksheet analysisSheet

    analysisSheet.Cells(1, 1).Value = "Add or remove rows of Global Summary"
    BuildAnalysisTable analysisSheet, sectionValue
End Sub

Public Sub PrepareFullAnalysisWorksheet(Optional ByVal headerInstruction As String = "Add or remove rows of Global Summary")

    Dim analysisSheet As Worksheet
    Dim nextRow As Long
    Dim spatioTemporalRows As Variant

    Set analysisSheet = EnsureWorksheet(ANALYSISSHEET)
    ClearWorksheet analysisSheet

    analysisSheet.Cells(1, 1).Value = headerInstruction

    nextRow = 3
    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_GLOBAL_SUMMARY, _
                               AnalysisHeaders(), AnalysisRows("Initial Section"))
    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_UNIVARIATE, _
                               AnalysisHeaders(), Array(Array("Univariate Section", "Univariate Title", "Summary Uni")))
    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_BIVARIATE, _
                               AnalysisHeaders(), Array(Array("Bivariate Section", "Bivariate Title", "Summary Bi")))
    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_TIME_SERIES, _
                               Array("Series ID", "Table order", "Label"), _
                               Array(Array("Series 1", 2, "Alpha")))
    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_GRAPH_TIME_SERIES, _
                               Array("Graph ID", "Section", "Table Title", "Summary label", "Choices"), _
                               Array(Array("Graph 5", "Section B", "Title B", "Summary B", "Choice B"), _
                                     Array("Graph 2", "Section A", "Title A", "Summary A", "Choice A")))
    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_GRAPH_TITLE, _
                               Array("Graph ID", "Graph Title"), _
                               Array(Array("Graph 5", "Graph Title B")))
    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_SPATIAL, _
                               Array("Section", "Label", "Summary label", "Choices"), _
                               Array(Array("Spatial Section", "Spatial Label", "Spatial Summary", "Spatial Choice")))

    spatioTemporalRows = Array( _
        Array("Region A", "Label A", "Choice A", "Graph Title A"), _
        Array("Region B", "Label B", "Choice B", "Graph Title B"), _
        Array("Region C", "Label C", "Choice C", "Graph Title C"), _
        Array(Empty, Empty, Empty, Empty), _
        Array(Empty, Empty, Empty, Empty))

    nextRow = AddAnalysisTable(analysisSheet, nextRow, TAB_SPATIO_TEMPORAL, _
                               Array("Section (select)", "Label", "Choices", "Graph Title"), _
                               spatioTemporalRows)
    Call AddAnalysisTable(analysisSheet, nextRow, TAB_SPATIO_TEMPORAL_SPECS, _
                          Array("Section", "Label", "Summary label"), _
                          Array(Array("Specs Section", "Specs Label", "Specs Summary")))
End Sub

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

