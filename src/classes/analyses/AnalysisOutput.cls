VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Build and format all analysis worksheets"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Build and format all analysis worksheets")
'@IgnoreModule UnrecognizedAnnotation, IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IAnalysisOutput

'@class AnalysisOutput
'@description
'Orchestrates the complete analysis output pipeline for the linelist builder.
'Given a setup worksheet containing analysis definition ListObjects and an
'output linelist facade, this class iterates through each analysis scope
'(normal, time series, spatial, spatio-temporal), builds cross-tables from
'table specifications, writes formulas, registers metadata in AnaTabIds, and
'renders charts. Consumers interact through the IAnalysisOutput interface.
'@depends ILinelist, ILinelistSpecs, ILLFormat, ITranslationObject, ITableSpecs, ICrossTable, ICrossTableFormula, IGraphSpecs, IGraphs, IAnaTabIds, IButtons, IPasswords, IDropdownLists, ILLdictionary, ILLVariables, IFormulaData, BetterArray
'@version 1.0 (2026-02-10)


Private Const UPDATEFILTERCOMMAND As String = "UpdateFilterTables"
Private Const SHPFILTERNAME As String = "SHP_Filter"
Private Const GOTOSECCODE As String = "go_to_section"
Private Const UPDATESPCOMMAND As String = "UpdateSpTables"
Private Const CLASS_NAME As String = "AnalysisOutput"

' ListObject names in the setup worksheet
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"

Private Type TAnaOutput
    anaLL As ILinelist
    sh As Worksheet
    namObj As IAnaTabIds                             'Cached analysis tracking object
    ' Bug #2 fix: GoTo arrays are now local to each scope, not shared instance state.
    ' They are passed as parameters to WriteTable and managed per-scope in WriteAnalysis.
End Type

Private this As TAnaOutput

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create an AnalysisOutput orchestrator bound to a setup worksheet and output linelist.
'@details
'Factory method that creates a new AnalysisOutput instance tied to the given
'setup worksheet and output linelist. The setup worksheet contains the
'ListObjects that define which tables (global summary, univariate, bivariate,
'time series, spatial, spatio-temporal) should be generated. The linelist is
'the output workbook facade where analysis sheets will be created and
'populated. Both parameters are validated before construction.
'@param specSh Worksheet. The setup worksheet containing analysis ListObjects
'(Tab_global_summary, Tab_Univariate_Analysis, Tab_Bivariate_Analysis,
'Tab_TimeSeries_Analysis, Tab_Spatial_Analysis, Tab_SpatioTemporal_Analysis).
'@param ll ILinelist. The output linelist workbook facade where analysis sheets
'will be created.
'@return IAnalysisOutput. A fully initialised orchestrator instance ready for use.
'@throws InvalidArgument When the specifications worksheet is Nothing.
'@throws InvalidArgument When the linelist is Nothing.
'@export
Public Function Create(ByVal specSh As Worksheet, _
                       ByVal ll As ILinelist) As IAnalysisOutput

    If specSh Is Nothing Then ThrowError InvalidArgument, "Specifications worksheet not defined"
    If ll Is Nothing Then ThrowError InvalidArgument, "Linelist not defined"

    With New AnalysisOutput
        Set .OutputLL = ll
        Set .Wksh = specSh
        Set Create = .Self
    End With
End Function

'@section Internal Properties
'===============================================================================
'@description Properties used during factory construction and internal wiring.

'@label:self
'@prop-title Current object instance.
'@details
'Convenience accessor that returns the current instance cast to the
'IAnalysisOutput interface, enabling the predeclared factory pattern.
'@return IAnalysisOutput. The current instance cast to the interface.
Public Property Get Self() As IAnalysisOutput
    Set Self = Me
End Property

'@label:output-ll-get
'@prop-title Output linelist facade.
'@details
'Returns the ILinelist reference representing the output workbook where
'analysis worksheets are created and populated.
'@return ILinelist. The output linelist workbook facade.
Public Property Get OutputLL() As ILinelist
    Set OutputLL = this.anaLL
End Property

'@label:output-ll-set
'@prop-title Assign the output linelist facade.
'@param ll ILinelist. The output linelist workbook facade to store.
Public Property Set OutputLL(ByVal ll As ILinelist)
    Set this.anaLL = ll
End Property

'@label:wksh-get
'@prop-title Setup worksheet reference.
'@details
'Returns the setup worksheet containing analysis definition ListObjects.
'@return Worksheet. The setup worksheet containing analysis ListObjects.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the setup worksheet reference.
'@param specSh Worksheet. The setup worksheet to store.
Public Property Set Wksh(ByVal specSh As Worksheet)
    Set this.sh = specSh
End Property

'@section Dependency Properties
'===============================================================================
'@description Convenience accessors for frequently used collaborators retrieved
'through the output linelist and its metadata facade.

'@label:linelist-data
'@prop-title Linelist metadata facade.
'@details
'Convenience accessor that retrieves the ILinelistSpecs object from the
'stored output linelist. ILinelistSpecs is the central metadata facade
'providing access to the dictionary, translation objects, formula data,
'design format, passwords, and sheet configuration. This property avoids
'repeating OutputLL().LinelistData() throughout the class.
'@return ILinelistSpecs. Metadata facade for the output linelist.
Private Property Get LinelistData() As ILinelistSpecs
    Set LinelistData = OutputLL().LinelistData()
End Property

'@label:design-format
'@prop-title Design format object.
'@details
'Convenience accessor for the ILLFormat design format object used to
'apply visual formatting (fonts, colours, borders, cell styles) to analysis
'worksheets, cells, and chart elements. Delegates through LinelistData to
'retrieve the shared format singleton so all analysis output uses consistent
'styling from the setup workbook.
'@return ILLFormat. Design format object for the output linelist.
Private Property Get DesignFormat() As ILLFormat
    Set DesignFormat = LinelistData().DesignFormat()
End Property

'@label:trans-object
'@prop-title Translation object for a given scope.
'@details
'Retrieves an ITranslationObject for the requested translation scope.
'Translation objects resolve human-readable labels, button captions, sheet
'names, and other locale-dependent strings from the setup workbook. The
'scope parameter controls which translation table is used: messages (default),
'shape names, or other categories defined by the TranslationScope enum.
'@param scope Optional Byte. A constant from the TranslationScope enum.
'Defaults to TranslationOfMessages. Use TranslationOfShapes for button/shape labels.
'@return ITranslationObject. Bound to the requested translation scope.
Public Property Get TransObject(Optional ByVal scope As Byte = TranslationOfMessages) As ITranslationObject
    Set TransObject = LinelistData().TransObject(scope)
End Property

'@label:output-sh
'@prop-title Output worksheet for a given analysis scope.
'@details
'Resolves the output worksheet in the linelist workbook for the given
'analysis scope. Each scope maps to a translated sheet name stored in
'the messages translation table (e.g. "LLSHEET_Analysis" for normal,
'"LLSHEET_TemporalAnalysis" for time series). The translation ensures
'the output sheet names match the user's configured language.
'@param scope Optional Byte. A constant from the AnalysisScope enum indicating
'which analysis worksheet to return. Defaults to AnalysisScopeNormal.
'@return Worksheet. The output analysis sheet in the linelist workbook.
Private Property Get OutputSh(Optional ByVal scope As Byte = AnalysisScopeNormal) As Worksheet
    Dim trads As ITranslationObject
    Dim ll As ILinelist

    Set ll = OutputLL()
    Set trads = TransObject(TranslationOfMessages)

    Select Case scope
    Case AnalysisScopeTimeSeries
        Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_TemporalAnalysis"))
    Case AnalysisScopeSpatial
        Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_SpatialAnalysis"))
    Case AnalysisScopeNormal
        Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_Analysis"))
    Case AnalysisScopeSpatioTemporal
        Set OutputSh = ll.Wksh(trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"))
    End Select
End Property

'@label:names-object
'@prop-title Lazily-initialised analysis tracking object.
'@details
'Lazily-initialised accessor for the IAnaTabIds tracking object. AnaTabIds
'maintains a registry of all cross-table named ranges and graph metadata
'created during the analysis build. This registry is later used by the
'update/filter commands to locate tables and graphs at runtime. The object
'is created once on first access using the temporary analysis worksheet
'from the linelist, and all subsequent calls return the cached instance.
'@return IAnaTabIds. Tracking object that records table and graph metadata.
'@remarks The backing worksheet comes from lData.TemporarySheetName(TemporarySheetAnalysis),
'which must exist in the output linelist before this property is first accessed.
Private Property Get NamesObject() As IAnaTabIds
    If this.namObj Is Nothing Then
        Dim lData As ILinelistSpecs
        Dim ll As ILinelist
        Dim sh As Worksheet

        Set ll = OutputLL()
        Set lData = ll.LinelistData()
        Set sh = ll.Wksh(lData.TemporarySheetName(TemporarySheetAnalysis))
        Set this.namObj = AnaTabIds.Create(sh, False)
    End If
    Set NamesObject = this.namObj
End Property

'@section Application Helpers
'===============================================================================

'@label:busy-app
'@sub-title Suppress Excel visual feedback and recalculation for performance.
'@details
'Disables ScreenUpdating, DisplayAlerts, Calculation (switches to manual),
'and EnableAnimations on the Application object to maximise performance during
'batch write operations. These settings are restored externally after the
'full analysis build completes.
'@remarks This sub is called multiple times (e.g. inside SplitOutputSheet)
'because activating a worksheet can re-enable ScreenUpdating in some Excel
'versions.
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'@label:split-output-sheet
'@sub-title Apply freeze-pane split to an analysis output worksheet.
'@details
'Applies a freeze-pane split to the output analysis worksheet so that
'header rows and the left navigation columns remain visible while the
'user scrolls through tables. The row split is 2 for normal/spatial
'scopes and 3 for time-series and spatio-temporal scopes (which have
'an extra GoTo-header dropdown row). The column split is always 3.
'Zoom is set to 80% for a wider view of cross-tables.
'@param scope Optional Byte. A constant from the AnalysisScope enum.
'Defaults to AnalysisScopeNormal. Time-series and spatio-temporal scopes
'get an extra frozen row.
'@remarks The sub activates the worksheet (required by the FreezePanes API),
'wraps everything in On Error Resume Next to tolerate protected-sheet or
'hidden-window conditions, and calls BusyApp twice because Activate can
're-enable ScreenUpdating.
Private Sub SplitOutputSheet(Optional ByVal scope As Byte = AnalysisScopeNormal)
    Dim outsh As Worksheet
    Dim rowSplit As Long

    Set outsh = OutputSh(scope)
    rowSplit = IIf(scope = AnalysisScopeTimeSeries Or scope = AnalysisScopeSpatioTemporal, 3, 2)

    On Error Resume Next
    BusyApp
    outsh.Activate
    BusyApp
    With Application.ActiveWindow
        .SplitRow = rowSplit
        .SplitColumn = 3
        .Zoom = 80
        .FreezePanes = True
    End With
    On Error GoTo 0
End Sub

'@section Sheet Preparation
'===============================================================================

'@label:add-commands
'@sub-title Add a filter/update command button to the analysis output worksheet.
'@details
'Places a command button on the output worksheet that triggers either the
'standard filter-update macro (UPDATEFILTERCOMMAND) or the spatial-specific
'update macro (UPDATESPCOMMAND) depending on the scope. The button is placed
'at row 2 col 2, styled using the design format, and labelled with the
'translated SHP_Filter shape name. Exits early when the linelist is built
'from a template (mainobj.HasTemplate), because the template already
'contains the buttons.
'@param scope Optional Byte. A constant from the AnalysisScope enum. Defaults
'to AnalysisScopeNormal. When AnalysisScopeSpatial, the spatial-specific
'update command is assigned.
Private Sub AddCommands(Optional ByVal scope As Byte = AnalysisScopeNormal)
    Dim cellRng As Range
    Dim commandName As String
    Dim shapeLabel As String
    Dim design As ILLFormat
    Dim but As IButtons
    Dim outsh As Worksheet
    Dim trads As ITranslationObject
    Dim lData As ILinelistSpecs

    Set lData = LinelistData()
    If lData.HasTemplate() Then Exit Sub

    Set design = DesignFormat()
    Set outsh = OutputSh(scope)
    Set trads = TransObject(TranslationOfShapes)
    Set cellRng = outsh.Cells(2, 2)

    commandName = IIf(scope = AnalysisScopeSpatial, UPDATESPCOMMAND, UPDATEFILTERCOMMAND)
    shapeLabel = trads.TranslatedValue(SHPFILTERNAME)
    Set but = Buttons.Create(outrng:=cellRng, codeName:=SHPFILTERNAME, scope:=ButtonScopeSmall)
    but.Add actionCommand:=commandName, shapeLabel:=shapeLabel
    but.Format design
End Sub

'@label:prepare-sheet
'@sub-title Perform five-step initialisation of an analysis output worksheet.
'@details
'Initialises an analysis output worksheet before any tables are written.
'The five steps are: (1) write a scope-specific flag named range at row 1
'col 10 (e.g. RNG_HasTable_UNI), formatted as a hidden cell; (2) apply
'the AllAnalysisSheet base format to the entire worksheet; (3) write a
'sheet-info identifier at row 1 col 3 as a hidden cell; (4) place the
'ComputedOnFiltered() warning formula at row 2 col 5 with LinelistWarning
'formatting; (5) delegate to AddCommands and SplitOutputSheet for button
'and freeze-pane setup.
'@param scope Optional Byte. A constant from the AnalysisScope enum. Defaults
'to AnalysisScopeNormal. Determines the flag range name, sheet-info string,
'and freeze-pane row count.
Private Sub PrepareSheet(Optional ByVal scope As Byte = AnalysisScopeNormal)
    Dim outsh As Worksheet
    Dim cellRng As Range
    Dim sheetInfo As String
    Dim rngName As String
    Dim design As ILLFormat
    Dim shHn As IHiddenNames

    Select Case scope
    Case AnalysisScopeTimeSeries
        sheetInfo = "TS-Analysis"
        rngName = "RNG_HasTable_TS"
    Case AnalysisScopeSpatial
        sheetInfo = "SP-Analysis"
        rngName = "RNG_HasTable_SP"
    Case AnalysisScopeNormal
        sheetInfo = "Uni-Bi-Analysis"
        rngName = "RNG_HasTable_UNI"
    Case AnalysisScopeSpatioTemporal
        sheetInfo = "SPT-Analysis"
        rngName = "RNG_HasTable_SPT"
    End Select

    Set outsh = OutputSh(scope)
    Set design = DesignFormat()

    Set cellRng = outsh.Cells(1, 10)
    cellRng.Name = rngName
    design.ApplyFormat cellRng, scope:=LinelistHiddenCell

    design.ApplyFormat outsh, scope:=AllAnalysisSheet

    'Store the sheet type as a worksheet-level HiddenName
    Set shHn = HiddenNames.Create(outsh)
    shHn.EnsureName "sheet_type", sheetInfo, HiddenNameTypeString

    Set cellRng = outsh.Cells(2, 5)
    cellRng.Formula = "= ComputedOnFiltered()"
    design.ApplyFormat cellRng, scope:=LinelistWarning

    AddCommands scope:=scope
    SplitOutputSheet scope:=scope
End Sub

'@section GoTo Dropdown Helpers
'===============================================================================

'@label:add-goto
'@sub-title Create a data-validation dropdown for section/graph navigation.
'@details
'Creates a data-validation dropdown in a single cell that lets the user
'navigate to a named section, header, or graph on the analysis sheet. The
'entries in dropArray become the validation list; selecting one triggers a
'GoTo macro that scrolls to the corresponding named range. The cell is
'formatted as a LinelistSelectionDropdown, its initial value is set to
'goToLabel, and it receives a workbook-level named range. If dropArray
'is empty the method exits silently so callers do not need to guard.
'@param cellRng Range. The single-cell Range where the dropdown will be placed.
'@param dropArray BetterArray. Display strings for the dropdown entries.
'Must have Length > 0 or the method exits early.
'@param goToLabel String. The default display value written into the cell.
'@param goToPrefix Optional String. A short prefix for the named range,
'scoped per analysis type. Defaults to "ua_".
'@param goToSuffix Optional String. An optional suffix appended to the named
'range, used to distinguish section-level vs graph-level dropdowns.
'Defaults to vbNullString.
Private Sub AddGoTo(ByVal cellRng As Range, _
                    ByVal dropArray As BetterArray, _
                    ByVal goToLabel As String, _
                    Optional ByVal goToPrefix As String = "ua_", _
                    Optional ByVal goToSuffix As String = vbNullString)

    Dim drop As IDropdownLists
    Dim ll As ILinelist
    Dim dropName As String
    Dim design As ILLFormat

    If dropArray.Length = 0 Then Exit Sub

    Set ll = OutputLL()
    Set drop = ll.Dropdown()
    Set design = DesignFormat()

    dropName = goToPrefix & "gotosection" & goToSuffix

    drop.Add dropArray, dropName
    drop.SetValidation cellRng, dropName

    design.ApplyFormat cellRng, scope:=LinelistSelectionDropdown

    cellRng.Value = goToLabel
    cellRng.Name = goToPrefix & GOTOSECCODE & goToSuffix
End Sub

'@section Temporal Section Helpers
'===============================================================================
'@description Min/Max formula finalisation and GoTo header dropdown for temporal sections.

'@label:min-max-formula
'@sub-title Finalise temporal section date-period boundary formulas and GoTo header dropdown.
'@details
'Finalises a temporal (time-series or spatio-temporal) section by writing
'its date-period boundary formulas and adding a GoTo Header dropdown. Each
'temporal section has two named ranges: MIN_MIN_DATE_{secId} and
'MAX_MAX_DATE_{secId}, which hold the overall minimum and maximum dates
'across all tables in that section. The minFormula and maxFormula strings
'are comma-separated MIN/MAX fragments accumulated by WriteTable; this
'method wraps them in a final worksheet formula. After writing, a GoTo
'Header dropdown is placed two rows below the SECTION_{secId} named range.
'@param outsh Worksheet. The output analysis worksheet where formulas and
'dropdown are written.
'@param secId String. The section identifier string (e.g. "SEC1") used to
'locate the named ranges.
'@param minFormula String. Comma-separated MIN expressions to wrap in the
'final = MIN(...) formula.
'@param maxFormula String. Comma-separated MAX expressions to wrap in the
'final = MAX(...) formula.
'@param goToPrefix String. Prefix for the GoTo named range ("ts_" or "spt_").
'@param headersList BetterArray. Header display strings for the dropdown entries.
'@param goToHeaderLabel String. Translated label for the GoTo Header dropdown
'default value.
Private Sub MinMaxFormula(ByVal outsh As Worksheet, ByVal secId As String, _
                          ByVal minFormula As String, ByVal maxFormula As String, _
                          ByVal goToPrefix As String, _
                          ByVal headersList As BetterArray, _
                          ByVal goToHeaderLabel As String)
    Dim cellRng As Range

    outsh.Range("MAX_MAX_DATE_" & secId).Formula = "= MAX(" & maxFormula & ")"
    outsh.Range("MIN_MIN_DATE_" & secId).Formula = "= MIN(" & minFormula & ")"

    ' Add goTo Header dropdown to this section
    Set cellRng = outsh.Range("SECTION_" & secId).Offset(2)
    AddGoTo cellRng:=cellRng, dropArray:=headersList, goToLabel:=goToHeaderLabel, _
            goToPrefix:=goToPrefix, goToSuffix:=secId
End Sub

'@section Table Creation
'===============================================================================
'@description Core table-creation loop and supporting write logic.

'@label:write-table
'@sub-title Iterate setup ListObject rows and build cross-tables on the output sheet.
'@details
'Core table-creation loop. Iterates over every data row of the setup
'ListObject identified by loName, and for each valid row: creates a
'TableSpecs, builds a CrossTable on the output sheet, applies
'CrossTableFormula formulas, formats the table with the design, registers
'the table's named ranges in the AnaTabIds tracking object, and optionally
'creates a chart (GraphSpecs + Graphs). For time-series and spatio-temporal
'scopes, additional logic accumulates MIN/MAX date formulas per section
'and writes them via MinMaxFormula when a section boundary is detected.
'GoTo section and header entries are pushed into the caller-provided
'sectionsList and headersList arrays. Each table is wrapped in a per-table
'error handler (Bug #3 fix) so that a single malformed row does not abort
'the entire analysis; errors are logged to the Immediate window.
'@param loName String. Name of the ListObject on the setup worksheet whose
'rows define the tables to build.
'@param sectionsList BetterArray. Accumulates GoTo section entries (passed
'ByVal but mutated via Push). The caller uses this after the call to
'create the section-level dropdown.
'@param headersList BetterArray. Accumulates GoTo header entries for
'time-series and spatio-temporal scopes.
'@param scope Optional Byte. A constant from the AnalysisScope enum. Defaults
'to AnalysisScopeNormal. Controls which output sheet receives the tables.
'@remarks The loop uses Do...Loop While with Offset(1) advancement. If the
'ListObject has no DataBodyRange, the method exits early.
Private Sub WriteTable(ByVal loName As String, _
                       ByVal sectionsList As BetterArray, _
                       ByVal headersList As BetterArray, _
                       Optional ByVal scope As Byte = AnalysisScopeNormal)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim tabRng As Range
    Dim hRng As Range
    Dim graphCounter As Long
    Dim tabl As ICrossTable
    Dim tabform As ICrossTableFormula
    Dim gr As IGraphs
    Dim lo As ListObject
    Dim specs As ITableSpecs
    Dim lData As ILinelistSpecs
    Dim trads As ITranslationObject
    Dim formData As IFormulaData
    Dim cellRng As Range
    Dim grSpecs As IGraphSpecs
    Dim tabId As String
    Dim tabType As Byte
    Dim design As ILLFormat
    Dim anchorRng As Range
    Dim containsTimeSeries As Boolean
    Dim spatRng As Range
    Dim sheetInfo As String
    Dim namObj As IAnaTabIds
    Dim tabRangesNames As BetterArray
    Dim goToSectionLabel As String
    Dim goToHeaderLabel As String

    ' Variables for time series
    Dim tableName As String
    Dim rowVar As String
    Dim minFormula As String
    Dim maxFormula As String
    Dim prevSpecs As ITableSpecs
    Dim nextSpecs As ITableSpecs
    Dim prevSecId As String
    Dim secId As String
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim grHeight As Long
    Dim isNewSection As Boolean
    Dim addPrevSecMinMaxFormula As Boolean
    Dim addActSecMinMaxFormula As Boolean
    Dim addSpatioTemporalName As Boolean
    Dim upRngName As String
    Dim goToPrefix As String

    Set sh = Wksh()
    Set outsh = OutputSh(scope)
    Set lo = sh.ListObjects(loName)
    Set lData = LinelistData()
    Set trads = TransObject()
    Set formData = lData.FormulaDataObject()
    Set dict = lData.Dictionary()
    Set vars = LLVariables.Create(dict)
    Set design = DesignFormat()
    Set namObj = NamesObject()

    Dim outshHn As IHiddenNames
    Set outshHn = HiddenNames.Create(outsh)
    sheetInfo = outshHn.ValueAsString("sheet_type")
    upRngName = "RNG_HasTable_" & UCase(Split(sheetInfo, "-")(0))
    goToSectionLabel = trads.TranslatedValue("MSG_GoToSec")
    goToHeaderLabel = trads.TranslatedValue("MSG_GoToHead")

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set hRng = lo.HeaderRowRange()
    Set anchorRng = hRng.Offset(lo.ListRows.Count)
    Set tabRng = lo.ListRows(1).Range

    Do
        Set specs = TableSpecs.Create(hRng, tabRng, lData.Dictionary())
        Set nextSpecs = Nothing
        addPrevSecMinMaxFormula = False
        addActSecMinMaxFormula = False
        addSpatioTemporalName = False

        If Not specs.ValidTable() Then GoTo ContinueLoop

        tabType = specs.TableScope()
        isNewSection = specs.IsNewSection()
        containsTimeSeries = (tabType = ScopeTimeSeries Or tabType = ScopeSpatioTemporal)
        tabId = specs.TableId()

        ' Bug #2 fix: push to LOCAL sectionsList, not shared instance array
        If isNewSection Then sectionsList.Push goToSectionLabel & ": " & specs.Value("section")

        ' Bug #3 fix: per-table error handling — wrap table creation in error handler
        On Error GoTo TableError

        Set tabl = CrossTable.Create(specs, outsh, lData)
        tabl.Build
        Set tabform = CrossTableFormula.Create(tabl, formData)
        tabform.AddFormulas
        tabl.Format design

        On Error GoTo 0

        Set tabRangesNames = tabl.NamedRangesList()
        namObj.AddTableInfos scope, tabId, tabRangesNames

        If IsEmpty(outsh.Range(upRngName)) Then outsh.Range(upRngName).Value = "yes"

        ' Non-temporal graphs
        If specs.HasGraph() And (Not containsTimeSeries) Then
            grHeight = 1
            On Error Resume Next
            grHeight = CLng(specs.Value("n geo"))
            On Error GoTo 0

            Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
            If tabType = ScopeSpatial Then Set cellRng = cellRng.Offset(, 2)

            Set grSpecs = GraphSpecs.Create(tabl)
            grSpecs.CreateSeries
            Set gr = Graphs.Create(outsh, cellRng)
            gr.Add

            For graphCounter = 1 To grSpecs.NumberOfSeries()
                gr.AddSeries grSpecs.SeriesName(graphCounter), _
                             grSpecs.SeriesType(graphCounter), _
                             grSpecs.SeriesPos(graphCounter)
                gr.AddLabels grSpecs.SeriesLabel(graphCounter), _
                             grSpecs.SeriesColumnLabel(graphCounter), _
                             hardCodeLabels:=False

                namObj.AddGraphInfo scope:=scope, tabId:=tabId, graphId:=tabId, _
                                    seriesName:=grSpecs.SeriesName(graphCounter), _
                                    seriesType:=grSpecs.SeriesType(graphCounter), _
                                    seriesPos:=grSpecs.SeriesPos(graphCounter), _
                                    seriesLabel:=grSpecs.SeriesLabel(graphCounter), _
                                    seriesColumnLabel:=grSpecs.SeriesColumnLabel(graphCounter), _
                                    hardCodeLabels:=False, _
                                    outRangeAddress:=cellRng.Address
            Next

            gr.Format catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                      valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                      scope:=scope, heightFactor:=grHeight

            namObj.AddGraphFormat scope:=scope, tabId:=tabId, graphId:=tabId, _
                                  catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                                  valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                                  heightFactor:=grHeight
        End If

        ' Time series / spatio-temporal min/max and GoTo logic
        If Not containsTimeSeries Then GoTo ContinueLoop

        ' Navigate to previous/next specs for section tracking.
        ' Previous() throws when IsNewSection is True, so guard explicitly
        ' instead of swallowing all errors with On Error Resume Next.
        If isNewSection Then
            Set prevSpecs = Nothing
        Else
            Set prevSpecs = specs.Previous()
        End If
        Set nextSpecs = specs.NextSpecs(anchorRng)
        If Not prevSpecs Is Nothing Then prevSecId = prevSpecs.TableSectionId()
        secId = specs.TableSectionId()

        rowVar = specs.Value("row")
        tableName = vars.Value(colName:="table name", varName:=rowVar)
        goToPrefix = IIf(tabType = ScopeTimeSeries, "ts_", "spt_")

        If prevSpecs Is Nothing Then
            minFormula = "MIN(" & tableName & "[" & rowVar & "])"
            maxFormula = "MAX(" & tableName & "[" & rowVar & "])"
            headersList.Push goToHeaderLabel & ": " & tabl.TimeSeriesHeader()
            addPrevSecMinMaxFormula = False
        End If

        If isNewSection And (Not prevSpecs Is Nothing) Then
            addPrevSecMinMaxFormula = True
            addSpatioTemporalName = (tabType = ScopeSpatioTemporal)
        End If

        If Not nextSpecs Is Nothing Then
            If nextSpecs.IsNewSection() And isNewSection Then
                addActSecMinMaxFormula = True
                addSpatioTemporalName = (tabType = ScopeSpatioTemporal)
            End If
        Else
            addActSecMinMaxFormula = True
            addSpatioTemporalName = (tabType = ScopeSpatioTemporal)
            Set cellRng = outsh.Range("ENDTABLE_" & secId).Offset(5, 2)
            If tabType = ScopeTimeSeries Then cellRng.Name = "START_GRAPHS_TS"
            If tabType = ScopeSpatioTemporal Then cellRng.Name = "START_GRAPHS_SPT"
            headersList.Push goToHeaderLabel & ": " & tabl.TimeSeriesHeader()
        End If

        If (Not isNewSection) And (Not nextSpecs Is Nothing) Then
            addPrevSecMinMaxFormula = False
            addActSecMinMaxFormula = False
            minFormula = minFormula & ", MIN(" & tableName & "[" & rowVar & "])"
            maxFormula = maxFormula & ", MAX(" & tableName & "[" & rowVar & "])"
            headersList.Push goToHeaderLabel & ": " & tabl.TimeSeriesHeader()
        End If

        If addPrevSecMinMaxFormula Then
            MinMaxFormula outsh:=outsh, secId:=prevSecId, _
                          minFormula:=minFormula, maxFormula:=maxFormula, _
                          goToPrefix:=goToPrefix, headersList:=headersList, _
                          goToHeaderLabel:=goToHeaderLabel
            ' Bug #2 fix: clear headers for the new section after writing previous section's dropdown
            headersList.Clear
        End If

        If addActSecMinMaxFormula Then
            MinMaxFormula outsh:=outsh, secId:=secId, _
                          minFormula:=minFormula, maxFormula:=maxFormula, _
                          goToPrefix:=goToPrefix, headersList:=headersList, _
                          goToHeaderLabel:=goToHeaderLabel
            headersList.Clear
        End If

        If isNewSection Then
            minFormula = "MIN(" & tableName & "[" & rowVar & "])"
            maxFormula = "MAX(" & tableName & "[" & rowVar & "])"
        End If

        If addSpatioTemporalName Then
            With outsh
                Set cellRng = .Range("COLUMN_CATEGORIES_" & tabId)
                Set spatRng = .Range( _
                    .Range("COLUMN_CATEGORIES_" & secId).Cells(1, 1), _
                    cellRng.Cells(1, cellRng.Columns.Count))
            End With
            spatRng.Name = "SPT_FORMULA_COLUMN_" & secId
        End If

        GoTo ContinueLoop

TableError:
        Debug.Print "Error on table " & tabId & ": " & Err.Number & " " & Err.Description
        On Error GoTo 0
        Resume ContinueLoop

ContinueLoop:
        Set tabRng = tabRng.Offset(1)
    Loop While tabRng.Row <= anchorRng.Row
End Sub

'@section Graph Creation
'===============================================================================
'@description Methods that build time-series and spatio-temporal charts after
'tables have been written.

'@label:write-time-series-graphs
'@sub-title Build all time-series charts from the graph definition ListObjects.
'@details
'Reads three setup ListObjects: the graph definition table (graphLoName),
'the graph title/label table (graphTitleLoName), and the original time-series
'table (tabLoName). These are bundled into a BetterArray and passed to
'GraphSpecs.CreateRangeSpecs, which operates in complex mode -- resolving
'multi-series graphs with mixed chart types and separate axis positions.
'The method then loops over each graph, creates it via Graphs.Create,
'iterates its series to add data ranges and labels, formats the chart,
'records metadata in AnaTabIds, and advances the anchor cell by 50 rows.
'Finally, a GoTo graph dropdown is placed at row 3 col 3.
'@param graphLoName String. Name of the ListObject defining graph series.
'@param graphTitleLoName String. Name of the ListObject defining graph
'titles and labels.
'@param tabLoName String. Name of the time-series table ListObject, used by
'GraphSpecs to resolve named ranges for series data.
'@param graphsList BetterArray. Accumulates GoTo graph display strings for
'the navigation dropdown.
'@remarks Exits early if the RNG_HasTable_TS flag is empty (no tables were
'written) or if either the graph or title ListObject has no data rows.
Private Sub WriteTimeSeriesGraphs(ByVal graphLoName As String, _
                                   ByVal graphTitleLoName As String, _
                                   ByVal tabLoName As String, _
                                   ByVal graphsList As BetterArray)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim graphCounter As Long
    Dim cellRng As Range
    Dim lData As ILinelistSpecs
    Dim seriesCounter As Long

    Dim grLo As ListObject
    Dim tabLo As ListObject
    Dim grTitleLo As ListObject
    Dim grSpecs As IGraphSpecs
    Dim gr As IGraphs
    Dim trads As ITranslationObject
    Dim design As ILLFormat
    Dim namObj As IAnaTabIds
    Dim goToGraphLabel As String

    Dim grBAinput As BetterArray
    Dim grBAoutput As BetterArray
    Dim tabGraphInfos As BetterArray
    Dim tabSeries As BetterArray
    Dim tabTypes As BetterArray
    Dim tabPos As BetterArray
    Dim tabLabel As BetterArray
    Dim tabColLabel As BetterArray
    Dim tabPrefLabel As BetterArray
    Dim plotTitle As String

    Set sh = Wksh()
    Set outsh = OutputSh(AnalysisScopeTimeSeries)

    If IsEmpty(outsh.Range("RNG_HasTable_TS")) Then Exit Sub

    Set grBAinput = New BetterArray
    Set namObj = NamesObject()
    Set lData = LinelistData()
    Set cellRng = outsh.Range("START_GRAPHS_TS")
    Set grLo = sh.ListObjects(graphLoName)
    Set tabLo = sh.ListObjects(tabLoName)
    Set grTitleLo = sh.ListObjects(graphTitleLoName)
    Set trads = TransObject()
    Set design = DesignFormat()
    goToGraphLabel = trads.TranslatedValue("MSG_GoToGraph")

    If (grLo.DataBodyRange Is Nothing) Or (grTitleLo.DataBodyRange Is Nothing) Then Exit Sub

    grBAinput.Push grLo, tabLo, grTitleLo
    Set grSpecs = GraphSpecs.CreateRangeSpecs(grBAinput, outsh, lData)

    If grSpecs.Valid() Then
        Set grBAoutput = grSpecs.SpecsLists()
        For graphCounter = 1 To grSpecs.NumberOfGraphs()
            Set tabGraphInfos = grBAoutput.Item(graphCounter)
            Set tabSeries = tabGraphInfos.Item(1)
            Set tabTypes = tabGraphInfos.Item(2)
            Set tabPos = tabGraphInfos.Item(3)
            Set tabLabel = tabGraphInfos.Item(4)
            Set tabColLabel = tabGraphInfos.Item(5)
            Set tabPrefLabel = tabGraphInfos.Item(6)

            Set gr = Graphs.Create(outsh, cellRng)
            gr.Add

            For seriesCounter = 1 To tabSeries.Length
                gr.AddSeries rngName:=tabSeries.Item(seriesCounter), _
                             chrtType:=tabTypes.Item(seriesCounter), _
                             axisPos:=tabPos.Item(seriesCounter)
                gr.AddLabels catName:=tabLabel.Item(seriesCounter), _
                             lblName:=tabColLabel.Item(seriesCounter), _
                             prefix:=tabPrefLabel.Item(seriesCounter), prefixOnly:=True

                namObj.AddGraphInfo scope:=AnalysisScopeTimeSeries, _
                                    tabId:=graphCounter, graphId:=graphCounter, _
                                    seriesName:=tabSeries.Item(seriesCounter), _
                                    seriesType:=tabTypes.Item(seriesCounter), _
                                    seriesPos:=tabPos.Item(seriesCounter), _
                                    seriesLabel:=tabLabel.Item(seriesCounter), _
                                    seriesColumnLabel:=vbNullString, _
                                    outRangeAddress:=cellRng.Address, _
                                    prefix:=tabPrefLabel.Item(seriesCounter), _
                                    prefixOnly:=True
            Next

            plotTitle = tabGraphInfos.Item(7)
            gr.Format catTitle:=trads.TranslatedValue("MSG_Period"), _
                      valuesTitle:=trads.TranslatedValue("MSG_Values"), _
                      plotTitle:=plotTitle, scope:=GraphScopeTimeSeries

            cellRng.Value = plotTitle

            namObj.AddGraphFormat scope:=AnalysisScopeTimeSeries, _
                                  tabId:=graphCounter, graphId:=graphCounter, _
                                  catTitle:=trads.TranslatedValue("MSG_Period"), _
                                  valuesTitle:=trads.TranslatedValue("MSG_Values"), _
                                  plotTitle:=plotTitle

            design.ApplyFormat cellRng, scope:=LinelistHiddenCell
            graphsList.Push goToGraphLabel & ": " & plotTitle
            Set cellRng = cellRng.Offset(50)
        Next
    End If

    ' Add GoTo graph dropdown
    AddGoTo outsh.Cells(3, 3), dropArray:=graphsList, _
            goToLabel:=goToGraphLabel, goToPrefix:="ts_", goToSuffix:="_graph"
End Sub

'@label:write-spatio-temporal-graph
'@sub-title Build spatio-temporal charts from the table definition ListObject.
'@details
'Builds spatio-temporal charts by re-reading the same setup ListObject
'that was used to create the spatio-temporal tables (loName). Unlike
'WriteTimeSeriesGraphs which uses a separate graph-definition ListObject,
'this method iterates over the original table rows, checks HasGraph for
'each, and creates one chart per qualifying table using the simpler
'GraphSpecs.Create (non-complex) mode. Each graph's plot title is composed
'from the translated "MSG_Graph" label plus the STARTCOL named-range value.
'Charts are spaced 50 rows apart. After all charts are created, a GoTo
'graph dropdown is placed at row 3 col 3.
'@param loName String. Name of the spatio-temporal ListObject whose rows
'define both the tables and the graphs.
'@param graphsList BetterArray. Accumulates GoTo graph display strings for
'the navigation dropdown.
'@remarks Skips rows where ValidTable is False or HasGraph is False. Exits
'early if the ListObject has no DataBodyRange.
Private Sub WriteSpatioTemporalGraph(ByVal loName As String, _
                                      ByVal graphsList As BetterArray)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim cellRng As Range
    Dim specs As ITableSpecs
    Dim lData As ILinelistSpecs
    Dim design As ILLFormat
    Dim hRng As Range
    Dim tabRng As Range
    Dim anchorRng As Range
    Dim gr As IGraphs
    Dim grSpecs As IGraphSpecs
    Dim plotTitle As String
    Dim tabId As String
    Dim graphCounter As Long
    Dim trads As ITranslationObject
    Dim lo As ListObject
    Dim tabl As ICrossTable
    Dim namObj As IAnaTabIds
    Dim goToGraphLabel As String

    Set sh = Wksh()
    Set outsh = OutputSh(AnalysisScopeSpatioTemporal)
    Set lo = sh.ListObjects(loName)
    Set lData = LinelistData()
    Set trads = TransObject()
    Set design = DesignFormat()
    Set namObj = NamesObject()
    goToGraphLabel = trads.TranslatedValue("MSG_GoToGraph")

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set cellRng = outsh.Range("START_GRAPHS_SPT")
    Set hRng = lo.HeaderRowRange()
    Set anchorRng = hRng.Offset(lo.ListRows.Count)
    Set tabRng = lo.ListRows(1).Range

    Do
        Set specs = TableSpecs.Create(hRng, tabRng, lData.Dictionary())

        If Not specs.ValidTable() Then GoTo ContinueLoop
        If Not specs.HasGraph() Then GoTo ContinueLoop

        tabId = specs.TableId()
        Set tabl = CrossTable.Create(specs, outsh, lData)

        Set grSpecs = GraphSpecs.Create(tabl)
        plotTitle = trads.TranslatedValue("MSG_Graph") & _
                    outsh.Range("STARTCOL_" & tabId).Value

        grSpecs.CreateSeries
        Set gr = Graphs.Create(outsh, cellRng)
        gr.Add

        For graphCounter = 1 To grSpecs.NumberOfSeries()
            gr.AddSeries grSpecs.SeriesName(graphCounter), _
                         grSpecs.SeriesType(graphCounter), _
                         grSpecs.SeriesPos(graphCounter)
            gr.AddLabels grSpecs.SeriesLabel(graphCounter), _
                         grSpecs.SeriesColumnLabel(graphCounter), _
                         hardCodeLabels:=False

            namObj.AddGraphInfo scope:=AnalysisScopeSpatioTemporal, tabId:=tabId, _
                                graphId:=tabId, _
                                seriesName:=grSpecs.SeriesName(graphCounter), _
                                seriesType:=grSpecs.SeriesType(graphCounter), _
                                seriesPos:=grSpecs.SeriesPos(graphCounter), _
                                seriesLabel:=grSpecs.SeriesLabel(graphCounter), _
                                seriesColumnLabel:=grSpecs.SeriesColumnLabel(graphCounter), _
                                hardCodeLabels:=False, _
                                outRangeAddress:=cellRng.Address
        Next

        gr.Format catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                  valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                  scope:=GraphScopeTimeSeries, plotTitle:=plotTitle

        namObj.AddGraphFormat scope:=AnalysisScopeSpatioTemporal, tabId:=tabId, _
                              graphId:=tabId, _
                              catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                              valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                              heightFactor:=1, plotTitle:=plotTitle

        cellRng.Value = plotTitle
        design.ApplyFormat cellRng, scope:=LinelistHiddenCell
        graphsList.Push goToGraphLabel & ": " & plotTitle
        Set cellRng = cellRng.Offset(50)

ContinueLoop:
        Set tabRng = tabRng.Offset(1)
    Loop While tabRng.Row <= anchorRng.Row

    AddGoTo outsh.Cells(3, 3), dropArray:=graphsList, _
            goToLabel:=goToGraphLabel, goToPrefix:="spt_", goToSuffix:="_graph"
End Sub

'@section Orchestration
'===============================================================================
'@description Top-level entry point that builds all four analysis output worksheets.

'@label:write-analysis
'@sub-title Write all analysis tables, formulas, and graphs to the output.
'@details
'Top-level orchestrator that builds all four analysis output worksheets in
'sequence: (1) Normal (univariate + bivariate + global summary), (2) Time
'Series, (3) Spatial, and (4) Spatio-Temporal. For each scope the method
'prepares the sheet (PrepareSheet), allocates fresh local GoTo arrays
'(Bug #2 fix), calls WriteTable for each relevant setup ListObject,
'optionally calls graph-creation subs, builds the section-level GoTo
'dropdown (AddGoTo), re-adds command buttons (AddCommands), password-
'protects the sheet, and hides it if no tables were written (the
'RNG_HasTable_* flag is still empty). Application performance mode is
'enabled once at the top via BusyApp.
'@export
'@remarks The normal scope writes three ListObjects (GLOBALSUMMARY,
'UNIVARIATE, BIVARIATE) in sequence. Time-series and spatio-temporal
'scopes also create graphs and a graph-level GoTo dropdown. The caller
'(IAnalysisOutput_WriteAnalysis) is responsible for restoring Application
'settings after this method returns.
Private Sub WriteAnalysis()
    Dim outsh As Worksheet
    Dim cellRng As Range
    Dim scope As Byte
    Dim pass As IPasswords
    Dim lData As ILinelistSpecs
    Dim trads As ITranslationObject
    Dim goToSectionLabel As String

    ' Bug #2 fix: GoTo arrays are LOCAL to each scope
    Dim sectionsList As BetterArray
    Dim headersList As BetterArray
    Dim graphsList As BetterArray

    Set lData = LinelistData()
    Set pass = lData.Password()
    Set trads = TransObject()
    goToSectionLabel = trads.TranslatedValue("MSG_GoToSec")

    BusyApp

    ' Univariate and Bivariate Analysis ========================================
    scope = AnalysisScopeNormal
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    ' Bug #2 fix: fresh arrays per scope
    Set sectionsList = New BetterArray
    Set headersList = New BetterArray

    WriteTable loName:=GLOBALSUMMARY, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope
    sectionsList.Push goToSectionLabel & ""
    WriteTable loName:=UNIVARIATE, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope
    WriteTable loName:=BIVARIATE, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="ua_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_UNI")) Then outsh.Visible = xlSheetHidden

    ' Time Series Analysis =====================================================
    scope = AnalysisScopeTimeSeries
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    Set sectionsList = New BetterArray
    Set headersList = New BetterArray
    Set graphsList = New BetterArray

    WriteTable loName:=TIMESERIES, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    WriteTimeSeriesGraphs graphLoName:=GRAPHTIMESERIES, _
                          graphTitleLoName:=TITLEGRAPHTIMESERIES, _
                          tabLoName:=TIMESERIES, graphsList:=graphsList

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="ts_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_TS")) Then outsh.Visible = xlSheetHidden

    ' Spatial Analysis =========================================================
    scope = AnalysisScopeSpatial
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    Set sectionsList = New BetterArray
    Set headersList = New BetterArray

    WriteTable loName:=SPATIAL, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="sp_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_SP")) Then outsh.Visible = xlSheetHidden

    ' Spatio-Temporal Analysis =================================================
    scope = AnalysisScopeSpatioTemporal
    Set outsh = OutputSh(scope)
    PrepareSheet scope:=scope

    Set sectionsList = New BetterArray
    Set headersList = New BetterArray
    Set graphsList = New BetterArray

    WriteTable loName:=SPATIOTEMPORAL, sectionsList:=sectionsList, _
               headersList:=headersList, scope:=scope

    WriteSpatioTemporalGraph loName:=SPATIOTEMPORAL, graphsList:=graphsList

    Set cellRng = outsh.Cells(2, 3)
    AddGoTo cellRng, dropArray:=sectionsList, goToLabel:=goToSectionLabel, goToPrefix:="spt_"
    AddCommands scope:=scope

    pass.Protect outsh, True
    If IsEmpty(outsh.Range("RNG_HasTable_SPT")) Then outsh.Visible = xlSheetHidden
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a project error with standardised source.
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the IAnalysisOutput contract.
'See the corresponding Public/Private members above for full documentation.

Private Sub IAnalysisOutput_WriteAnalysis()
    WriteAnalysis
End Sub
