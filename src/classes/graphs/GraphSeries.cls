VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Chart series for one built cross-table"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Chart series for one built cross-table")
'@IgnoreModule IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class GraphSeries
'@description
'The chart series of one built cross-table: univariate, bivariate, spatial or
'spatio-temporal. It reads the table scope and the graph settings off TableSpecs
'and answers one SeriesBuffer. AnalysisOutput walks that buffer and hands each
'entry to Graphs.
'
'THE SERIES ARE BUILT ON THE FIRST READ
'-------------------------------------------------------------------------------
'Series() builds once and keeps the answer, so a caller has one call to make and
'no order to remember. Read HasCheckings after Series(), because the reasons a
'series was left out are filed while the buffer is being built.
'
'EVERY SKIP IS REPORTED
'-------------------------------------------------------------------------------
'This half decides to build fewer series than the setup asked for in four
'places: a scope it has no builder for, a graph setting it does not recognise, a
'percentage chart on a table with no percentage column, and a table that ends
'with no series at all. Each one files an entry through LogInfo, and
'AnalysisOutput reads them into the generation report.
'@depends CrossTable, TableSpecs, SeriesBuffer, Checking
'@version 1.0 (2026-07-31)



'@section State
'===============================================================================

Private Type TGraphSeries
    tabl As CrossTable
    buffer As SeriesBuffer
    built As Boolean
    sealed As Boolean
    checkCounter As Long
    internalChecks As Checking
End Type

Private this As TGraphSeries

Private Const CLASS_NAME As String = "GraphSeries"

' The graph setting values this builder knows. TableSpecs.GraphMode trims the
' cell and folds its case before it answers, so a cell holding " Yes " matches.
Private Const MODE_VALUES As String = "values"
Private Const MODE_YES As String = "yes"
Private Const MODE_PERCENTAGES As String = "percentages"
Private Const MODE_BOTH As String = "both"


'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create the series builder of one cross-table
'@details
'Time series tables are drawn by TimeSeriesGraphs, which reads the three graph
'setup tables. Global summary tables carry no chart at all.
'@param tabl CrossTable. The cross-table whose named ranges hold the data to
'plot. The named ranges have to be on the worksheet before the buffer is read.
'@return GraphSeries. A builder ready for Series().
'@throws ProjectError.InvalidArgument When tabl is Nothing, a time series table
'or a global summary table.
'@export
Public Function Create(ByVal tabl As CrossTable) As GraphSeries
    Dim tableScope As Byte
    Dim newSeries As GraphSeries

    If tabl Is Nothing Then ThrowError InvalidArgument, "Cross-table is not defined"

    ' One read of the specification chain. It used to be walked twice.
    tableScope = tabl.Specifications.TableScope

    If tableScope = ScopeTimeSeries Then _
        ThrowError InvalidArgument, _
                   "A time series table is drawn by TimeSeriesGraphs, from the three graph setup tables."

    If tableScope = ScopeGlobalSummary Then _
        ThrowError InvalidArgument, "Global summary tables do not support graphs"

    Set newSeries = New GraphSeries
    Set newSeries.Table = tabl
    newSeries.Seal
    Set Create = newSeries
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes
'@details
'Called by the factory after the last setup write. Once sealed, the
'creation-only setters raise through ThrowError.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-creation writes
'@param propName String. The property name reported in the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:table
'@prop-title The cross-table this builder reads
'@return CrossTable. The stored cross-table.
Public Property Get Table() As CrossTable
    Set Table = this.tabl
End Property

'@label:table-set
'@prop-title Assign the cross-table
'@param tabl CrossTable. The cross-table to store.
Public Property Set Table(ByVal tabl As CrossTable)
    GuardNotSealed "Table"
    Set this.tabl = tabl
End Property


'@section Reading
'===============================================================================

'@label:series
'@prop-title The series of this table
'@details
'Builds on the first read and keeps the answer. A buffer whose Count is 0 says
'the table asks for a chart that carries nothing, and the reason is in the
'report.
'@return SeriesBuffer. One entry per series, in plot order.
'@export
Public Property Get Series() As SeriesBuffer
    If Not this.built Then BuildSeries
    Set Series = this.buffer
End Property

'@label:wksh
'@prop-title The worksheet the named ranges of this table live on
'@return Worksheet. The worksheet of the cross-table.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.tabl.Wksh
End Property


'@section Building
'===============================================================================

'@label:build-series
'@sub-title Build the buffer from the table scope and its graph settings
'@details
'FlipCoordinates controls the plot direction: horizontal bars when it answers
'True, vertical bars otherwise. For a spatio-temporal table the column count is
'the geographic unit count and the row categories are the ones of the section,
'because every table of a section shares them.
'@remarks TableSpecs.GeoCount carries the default of 5 and clamps the value to
'1 to 20, so an empty or unreadable cell needs no guard here.
Private Sub BuildSeries()
    Dim specs As TableSpecs
    Dim tabId As String
    Dim tabType As Byte
    Dim nbCols As Long
    Dim plotType As String
    Dim tabSecId As String

    this.built = True
    Set this.buffer = SeriesBuffer.Create()

    Set specs = this.tabl.Specifications
    tabId = specs.TableId
    tabType = specs.TableScope
    nbCols = this.tabl.NumberOfColumns

    If tabType = ScopeSpatioTemporal Then
        nbCols = specs.GeoCount()
        tabSecId = specs.TableSectionId
    End If

    ' The flag is read through TableSpecs, which trims it and folds its case.
    ' Comparing the raw cell here made a cell holding " Yes " open the gate in
    ' AnalysisOutput and match nothing in this builder.
    plotType = IIf(specs.FlipCoordinates(), "hbar", "bar")

    Select Case tabType

    Case ScopeUnivariate
        BuildUnivariateSeries tabId, plotType, specs.HasPercentage

    Case ScopeBivariate, ScopeSpatial, ScopeSpatioTemporal
        BuildMultiColumnSeries specs, tabId, tabSecId, tabType, nbCols, plotType

    Case Else
        LogInfo "Table " & tabId & " has the scope " & tabType & ", which has no " & _
                "series builder. No chart series were built for it.", checkingError

    End Select

    If this.buffer.Count = 0 Then
        LogInfo "Table " & tabId & " asks for a chart and no series could be built " & _
                "for it, so the chart is empty.", checkingWarning
    End If
End Sub

'@label:build-univariate-series
'@sub-title Build the series of a single-variable table
'@details
'One bar series over the value column, labelled by the row categories and named
'by the column label. A table carrying percentages gets a second series of
'points on the right axis, so the percentages sit over the bars. Horizontal bars
'are left without that overlay, because a point series on a horizontal chart
'reads poorly.
'@param tabId String. The table identifier every named range carries.
'@param plotType String. "bar" for vertical bars, "hbar" for horizontal ones.
'@param hasPerc Boolean. Whether the table carries a percentage column.
Private Sub BuildUnivariateSeries(ByVal tabId As String, _
                                  ByVal plotType As String, _
                                  ByVal hasPerc As Boolean)

    this.buffer.Add "VALUES_COL_1_" & tabId, plotType, "left", _
                    "ROW_CATEGORIES_" & tabId, "LABEL_COL_1_" & tabId

    If hasPerc And (plotType = "bar") Then
        this.buffer.Add "PERC_COL_1_" & tabId, "point", "right", _
                        "ROW_CATEGORIES_" & tabId, "PERC_LABEL_COL_" & tabId
    End If
End Sub

'@label:build-multi-column-series
'@sub-title Build the series of a table with several data columns
'@details
'The graph mode comes from TableSpecs.GraphMode. "values" and "yes" give one bar
'series per column over the value ranges. "percentages" gives one bar series per
'column over the percentage ranges, and reports when the table carries none.
'"both" gives value bars on the left axis and percentage points on the right,
'per column. A spatio-temporal table takes its row categories from its section.
'@param specs TableSpecs. The specifications of the table.
'@param tabId String. The table identifier every named range carries.
'@param tabSecId String. The section identifier, for a spatio-temporal table.
'@param tabType Byte. The table scope.
'@param nbCols Long. How many data columns to walk.
'@param plotType String. "bar" or "hbar".
'@remarks A graph mode outside the three the setup dropdowns offer files an
'error entry. It used to fall through the Select Case with no arm, which left
'the chart with no series and raised 1004 in Graphs.Format two frames later.
Private Sub BuildMultiColumnSeries(ByVal specs As TableSpecs, _
                                   ByVal tabId As String, _
                                   ByVal tabSecId As String, _
                                   ByVal tabType As Byte, _
                                   ByVal nbCols As Long, _
                                   ByVal plotType As String)
    Dim graphValue As String
    Dim counter As Long
    Dim rowCatName As String

    graphValue = specs.GraphMode()

    If tabType = ScopeSpatioTemporal Then
        rowCatName = "ROW_CATEGORIES_" & tabSecId
    Else
        rowCatName = "ROW_CATEGORIES_" & tabId
    End If

    Select Case graphValue

    Case MODE_VALUES, MODE_YES
        For counter = 1 To nbCols
            this.buffer.Add "VALUES_COL_" & counter & "_" & tabId, plotType, "left", _
                            rowCatName, "LABEL_COL_" & counter & "_" & tabId
        Next

    Case MODE_PERCENTAGES
        If specs.HasPercentage Then
            For counter = 1 To nbCols
                this.buffer.Add "PERC_COL_" & counter & "_" & tabId, plotType, "left", _
                                rowCatName, "LABEL_COL_" & counter & "_" & tabId
            Next
        Else
            LogInfo "Table " & tabId & " asks for a chart of percentages and carries " & _
                    "no percentage column, so no series were built for it.", checkingWarning
        End If

    Case MODE_BOTH
        For counter = 1 To nbCols
            this.buffer.Add "VALUES_COL_" & counter & "_" & tabId, plotType, "left", _
                            rowCatName, "LABEL_COL_" & counter & "_" & tabId
            If specs.HasPercentage Then
                this.buffer.Add "PERC_COL_" & counter & "_" & tabId, "point", "right", _
                                rowCatName, "LABEL_COL_" & counter & "_" & tabId
            End If
        Next

    Case Else
        LogInfo "Table " & tabId & " asks for a chart and its graph setting """ & _
                graphValue & """ is none of values, percentages or both. " & _
                "No series were built for it.", checkingError

    End Select
End Sub


'@section Checking
'===============================================================================
'@description The report of everything this builder decided to leave out.

'@label:specs-id
'@sub-title A name for this instance that no other instance shares
'@details
'AnalysisOutput merges the entries of every collaborator into one report, and
'Checking.Add raises on a duplicate key, so a key made of a bare counter
'collides as soon as two instances file an entry. One instance serves one table,
'so the table identifier names it.
'@return String. The identifier every key of this instance carries.
Private Function SpecsId() As String
    If Not this.tabl Is Nothing Then SpecsId = this.tabl.Specifications.TableId
    If LenB(SpecsId) = 0 Then SpecsId = "unplaced"
End Function

'@label:log-info
'@sub-title Record a diagnostic entry in the internal checking store
'@details
'Creates the Checking store on first use and appends the message under a key
'naming this class, this table and a counter.
'@param label String. The message to record.
'@param scope Optional Byte. Severity from CheckingLogType. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)
    Dim specsKey As String

    specsKey = SpecsId()

    If this.internalChecks Is Nothing Then _
        Set this.internalChecks = Checking.Create("Graph series of " & specsKey)

    this.internalChecks.Add CLASS_NAME & "-" & specsKey & "-" & CStr(this.checkCounter), _
                            label, scope
    this.checkCounter = this.checkCounter + 1
End Sub

'@label:has-checkings
'@prop-title Whether anything was reported while the series were built
'@return Boolean. True when at least one entry was filed.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = Not (this.internalChecks Is Nothing)
End Property

'@label:checking-values
'@prop-title The report of this build
'@return Checking. The store, or Nothing when nothing was filed.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@section ErrorHandling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
