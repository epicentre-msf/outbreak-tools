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

'@section Enumerations
'===============================================================================
'@description
'How far a build goes. The analysis scopes themselves live in AnaTabIds: that
'class routes on the value, the generated linelist carries it, and one
'definition serves this class, the export and the linelist.
'
'The stage is a different thing from the scope, and it used to be a fifth
'member of the scope enum. Four Select Case blocks that route on a scope had no
'branch for it, so a stage value reaching any of them gave back Nothing or an
'empty string several frames from where it would fail.

Public Enum AnalysisBuildStage
    AnalysisBuildStageAll = 0
    AnalysisBuildStageNormal = 1
    AnalysisBuildStageTimeSeriesTables = 2
    AnalysisBuildStageTimeSeries = 3
    AnalysisBuildStageSpatial = 4
End Enum

'@class AnalysisOutput
'@description
'Writes every analysis sheet of a generated linelist. Create takes the setup
'worksheet that holds the analysis definition ListObjects and the linelist
'facade the sheets are written into. WriteAnalysis walks each analysis scope --
'normal, time series, spatial and spatio-temporal -- builds the cross-tables
'from their table specifications, writes the formulas, records the chart
'metadata in AnaTabIds and draws the charts. TransObject answers the
'translation provider, and entries filed along the way leave through
'HasCheckings and CheckingValues.
'
'THE CLASS OWNS THE STATE ITS OWN WORK NEEDS
'-------------------------------------------------------------------------------
'WriteAnalysis opens an ApplicationState of its own, turns screen updating,
'alerts, automatic calculation, animations and events off for the length of the
'build, and restores on the way out -- on the error path too. The sheets are
'therefore written under the state the analysis needs whoever calls it, and a
'caller that keeps its own scope, which is what clickGenerate does, is not
'disturbed: the snapshot taken here is the caller's busy state, so restoring
'puts that back rather than the pre-generation values.
'
'A SECTION IS TOLD TO THE LOOP
'-------------------------------------------------------------------------------
'The date bounds and the navigation dropdown of a temporal section are held by
'a TemporalSection. The loop closes the section it holds when it meets a table
'that starts a new one, and closes the last one after the loop. Nothing reads
'the row below to decide where a section ends.
'@depends Linelist, LinelistSpecs, LLFormat, TranslationObject, TableSpecs, CrossTable, CrossTableFormula, Graphs, GraphSeries, TimeSeriesGraphs, SeriesBuffer, AnaTabIds, TemporalSection, ApplicationState, Buttons, Passwords, DropdownLists, HiddenNames, LLdictionary, LLVariables, FormulaData, Checking, BetterArray
'@version 2026-08-12


Private Const UPDATEFILTERCOMMAND As String = "UpdateFilterTables"
Private Const SHPFILTERNAME As String = "SHP_Filter"
Private Const GOTOSECCODE As String = "go_to_section"
Private Const UPDATESPCOMMAND As String = "UpdateSpTables"
Private Const CLASS_NAME As String = "AnalysisOutput"
Private Const GRAPHROWSTEP As Long = 50

' ListObject names in the setup worksheet
Private Const GLOBALSUMMARY As String = "Tab_Global_Summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"

Private Type TAnaOutput
    anaLL As Linelist
    sh As Worksheet
    namObj As AnaTabIds                             'Cached analysis tracking object
    outSheets(1 To 4) As Worksheet                  'One output worksheet per scope
    outNames(1 To 4) As HiddenNames                 'One metadata reader per scope
    appScope As ApplicationState                    'Busy state held for the length of one build
    hasChecks As Boolean
    tablesWritten As Long                           'Analysis tables put on a sheet
    graphsWritten As Long                           'Charts put on a sheet
    sealed As Boolean
End Type

Private this As TAnaOutput

'Checking support
Private checkCounter As Long
Private internalChecks As Checking

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create an AnalysisOutput orchestrator bound to a setup worksheet and output linelist.
'@details
'Factory method that creates a new AnalysisOutput instance tied to the given
'setup worksheet and output linelist. The setup worksheet holds the ListObjects
'that define which tables (global summary, univariate, bivariate, time series,
'spatial, spatio-temporal) should be generated. The linelist is the output
'workbook facade where analysis sheets are created and populated. Both
'parameters are validated before construction.
'@param specSh Worksheet. The setup worksheet holding the analysis ListObjects
'(Tab_Global_Summary, Tab_Univariate_Analysis, Tab_Bivariate_Analysis,
'Tab_TimeSeries_Analysis, Tab_Spatial_Analysis, Tab_SpatioTemporal_Analysis).
'@param ll Linelist. The output linelist workbook facade where analysis sheets
'will be created.
'@return AnalysisOutput. A fully initialised orchestrator instance ready for use.
'@throws InvalidArgument When the specifications worksheet is Nothing.
'@throws InvalidArgument When the linelist is Nothing.
'@export
Public Function Create(ByVal specSh As Worksheet, _
                       ByVal ll As Linelist) As AnalysisOutput

    If specSh Is Nothing Then ThrowError InvalidArgument, "Specifications worksheet not defined"
    If ll Is Nothing Then ThrowError InvalidArgument, "Linelist not defined"

    Dim instance As AnalysisOutput
    Set instance = New AnalysisOutput
    Set instance.OutputLL = ll
    Set instance.Wksh = specSh
    instance.Seal
    Set Create = instance
End Function

'@section Internal Properties
'===============================================================================
'@description Properties used during factory construction and internal wiring.

'@label:seal
'@sub-title Seal the instance against further setup writes.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against post-creation writes.
'@param propName String. The property name for the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:output-ll-get
'@prop-title Output linelist facade.
'@details
'Returns the Linelist reference representing the output workbook where
'analysis worksheets are created and populated.
'@return Linelist. The output linelist workbook facade.
Public Property Get OutputLL() As Linelist
    Set OutputLL = this.anaLL
End Property

'@label:output-ll-set
'@prop-title Assign the output linelist facade.
'@param ll Linelist. The output linelist workbook facade to store.
Public Property Set OutputLL(ByVal ll As Linelist)
    GuardNotSealed "OutputLL"
    Set this.anaLL = ll
End Property

'@label:wksh-get
'@prop-title Setup worksheet reference.
'@details
'Returns the setup worksheet holding the analysis definition ListObjects.
'@return Worksheet. The setup worksheet holding the analysis ListObjects.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the setup worksheet reference.
'@param specSh Worksheet. The setup worksheet to store.
Public Property Set Wksh(ByVal specSh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sh = specSh
End Property

'@section Dependency Properties
'===============================================================================
'@description Convenience accessors for frequently used collaborators retrieved
'through the output linelist and its metadata facade.

'@label:linelist-data
'@prop-title Linelist metadata facade.
'@details
'Convenience accessor that retrieves the LinelistSpecs object from the stored
'output linelist. LinelistSpecs is the central metadata facade providing access
'to the dictionary, translation objects, formula data, design format, passwords
'and sheet configuration.
'@return LinelistSpecs. Metadata facade for the output linelist.
Private Property Get LinelistData() As LinelistSpecs
    Set LinelistData = OutputLL().LinelistData()
End Property

'@label:design-format
'@prop-title Design format object.
'@details
'Convenience accessor for the LLFormat design format object used to apply
'visual formatting (fonts, colours, borders, cell styles) to analysis
'worksheets, cells and chart elements. Delegates through LinelistData to the
'shared format singleton so all analysis output uses consistent styling from
'the setup workbook.
'@return LLFormat. Design format object for the output linelist.
Private Property Get DesignFormat() As LLFormat
    Set DesignFormat = LinelistData().DesignFormat()
End Property

'@label:trans-object
'@prop-title Translation object for a given scope.
'@details
'Retrieves a TranslationObject for the requested translation scope. Translation
'objects resolve human-readable labels, button captions, sheet names and other
'locale-dependent strings from the setup workbook.
'@param scope Optional Byte. A constant from the TranslationScope enum.
'Defaults to TranslationOfMessages. Use TranslationOfShapes for button labels.
'@return TranslationObject. Bound to the requested translation scope.
Public Property Get TransObject(Optional ByVal scope As Byte = TranslationOfMessages) As TranslationObject
    Set TransObject = LinelistData().TransObject(scope)
End Property

'@label:output-sh
'@prop-title Output worksheet for a given analysis scope.
'@details
'Resolves the output worksheet in the linelist workbook for the given analysis
'scope. Each scope maps to a translated sheet name held in the messages
'translation table (for example "LLSHEET_Analysis" for the normal scope), so
'the output sheet names match the configured language.
'
'Each sheet is resolved once and kept. Resolving one costs a translated value
'and a worksheet lookup, and this property is read from the sheet preparation,
'the button placement, the freeze panes, every table and both chart passes.
'@param scope Optional Byte. A constant from the AnalysisScope enum saying
'which analysis worksheet to answer. Defaults to AnalysisScopeNormal.
'@return Worksheet. The output analysis sheet in the linelist workbook.
'@throws ProjectError.InvalidArgument When the scope is outside the four.
Private Property Get OutputSh(Optional ByVal scope As Byte = AnalysisScopeNormal) As Worksheet
    Dim trads As TranslationObject
    Dim ll As Linelist
    Dim sheetTag As String

    If Not IsKnownScope(scope) Then _
        ThrowError ProjectError.InvalidArgument, "Unknown analysis scope: " & scope

    If Not this.outSheets(scope) Is Nothing Then
        Set OutputSh = this.outSheets(scope)
        Exit Property
    End If

    Set ll = OutputLL()
    Set trads = TransObject(TranslationOfMessages)

    Select Case scope
    Case AnalysisScopeTimeSeries
        sheetTag = "LLSHEET_TemporalAnalysis"
    Case AnalysisScopeSpatial
        sheetTag = "LLSHEET_SpatialAnalysis"
    Case AnalysisScopeNormal
        sheetTag = "LLSHEET_Analysis"
    Case AnalysisScopeSpatioTemporal
        sheetTag = "LLSHEET_SpatioTemporalAnalysis"
    End Select

    Set this.outSheets(scope) = ll.Wksh(trads.TranslatedValue(sheetTag))
    Set OutputSh = this.outSheets(scope)
End Property

'@label:output-names
'@prop-title Metadata reader for the output worksheet of a scope.
'@details
'One HiddenNames per output sheet, resolved once. Each Create walks the
'worksheet names and reads a comment on every tracked one, and this class used
'to build about ten of them per run.
'@param scope Byte. A constant from the AnalysisScope enum.
'@return HiddenNames. Reader bound to that scope's output sheet.
Private Property Get OutputNames(ByVal scope As Byte) As HiddenNames
    If this.outNames(scope) Is Nothing Then
        Set this.outNames(scope) = HiddenNames.Create(OutputSh(scope))
    End If
    Set OutputNames = this.outNames(scope)
End Property

'@label:names-object
'@prop-title Lazily-initialised analysis tracking object.
'@details
'Lazily-initialised accessor for the AnaTabIds tracking object. AnaTabIds keeps
'the registry of chart metadata created during the analysis build, and the
'export reads it back to rebuild the charts in the exported workbook. The
'object is created once on first access, over the temporary analysis worksheet
'of the linelist.
'@return AnaTabIds. Tracking object that records the chart metadata.
'@remarks The backing worksheet comes from lData.TemporarySheetName(TempSheetAnalysis),
'which must exist in the output linelist before this property is first read.
Private Property Get NamesObject() As AnaTabIds
    If this.namObj Is Nothing Then
        Dim lData As LinelistSpecs
        Dim ll As Linelist
        Dim sh As Worksheet

        Set ll = OutputLL()
        Set lData = ll.LinelistData()
        Set sh = ll.Wksh(lData.TemporarySheetName(TempSheetAnalysis))
        Set this.namObj = AnaTabIds.Create(sh, False)
    End If
    Set NamesObject = this.namObj
End Property

'@section Scope Mapping
'===============================================================================
'@description The analysis scope, the chart scope, the sheet tag and the
'navigation prefix are four names for one thing. Each translation is stated
'here once, so no call site has to assume that two enums line up by number.

'@label:is-known-scope
'@sub-title Whether a value is one of the four analysis scopes
'@param scope Byte. The value to test.
'@return Boolean. True for the four scopes.
Private Function IsKnownScope(ByVal scope As Byte) As Boolean
    Select Case scope
    Case AnalysisScopeTimeSeries, AnalysisScopeSpatial, _
         AnalysisScopeNormal, AnalysisScopeSpatioTemporal
        IsKnownScope = True
    End Select
End Function

'@label:graph-scope-of
'@sub-title The chart scope that goes with an analysis scope
'@details
'GraphScope has three members and AnalysisScope has four. The two agree by
'number on the first three, and this class used to pass an analysis scope
'straight into a chart scope parameter. Spatio-temporal has no chart scope of
'its own, and its charts are laid out like the time series ones.
'@param scope Byte. A constant from the AnalysisScope enum.
'@return Byte. A constant from the GraphScope enum.
Private Function GraphScopeOf(ByVal scope As Byte) As Byte
    Select Case scope
    Case AnalysisScopeTimeSeries, AnalysisScopeSpatioTemporal
        GraphScopeOf = GraphScopeTimeSeries
    Case AnalysisScopeSpatial
        GraphScopeOf = GraphScopeSpatial
    Case Else
        GraphScopeOf = GraphScopeNormal
    End Select
End Function

'@label:sheet-tag-of
'@sub-title The sheet_type value that goes with an analysis scope
'@details
'EventLinelist, LLExporter, LLImporter and the epidemiological week helper all
'branch on this string, so an unknown scope raises here instead of stamping an
'empty tag.
'@param scope Byte. A constant from the AnalysisScope enum.
'@return String. The sheet_type value.
'@throws ProjectError.InvalidArgument When the scope is outside the four.
Private Function SheetTagOf(ByVal scope As Byte) As String
    Select Case scope
    Case AnalysisScopeTimeSeries
        SheetTagOf = "TS-Analysis"
    Case AnalysisScopeSpatial
        SheetTagOf = "SP-Analysis"
    Case AnalysisScopeNormal
        SheetTagOf = "Uni-Bi-Analysis"
    Case AnalysisScopeSpatioTemporal
        SheetTagOf = "SPT-Analysis"
    Case Else
        ThrowError ProjectError.InvalidArgument, "Unknown analysis scope: " & scope
    End Select
End Function

'@label:goto-prefix-of
'@sub-title The navigation prefix that goes with an analysis scope
'@param scope Byte. A constant from the AnalysisScope enum.
'@return String. The prefix every navigation name of that sheet carries.
Private Function GoToPrefixOf(ByVal scope As Byte) As String
    Select Case scope
    Case AnalysisScopeTimeSeries
        GoToPrefixOf = "ts_"
    Case AnalysisScopeSpatial
        GoToPrefixOf = "sp_"
    Case AnalysisScopeSpatioTemporal
        GoToPrefixOf = "spt_"
    Case Else
        GoToPrefixOf = "ua_"
    End Select
End Function

'@section Application Helpers
'===============================================================================

'@label:app-scope
'@prop-title The application state this class builds under.
'@details
'Built on first read and kept for the life of the instance. The snapshot is
'taken when the object is built, which is inside whatever state the caller was
'already in, so restoring hands that state back rather than guessing at one.
'@return ApplicationState. Bound to the running Excel Application.
'@remarks WriteAnalysis is what opens and restores it: see the note at the top
'of this class.
Private Property Get AppScope() As ApplicationState
    If this.appScope Is Nothing Then
        Set this.appScope = ApplicationState.Create(Application)
    End If

    Set AppScope = this.appScope
End Property

'@label:split-output-sheet
'@sub-title Apply the freeze-pane split to an analysis output worksheet.
'@details
'Freezes the header rows and the left navigation columns so they stay visible
'while the user scrolls through the tables. The row split is 2 for the normal
'and spatial scopes and 3 for the time series and spatio-temporal ones, which
'carry an extra row for the header dropdown. The column split is always 3 and
'the zoom is 80%.
'@param scope Byte. A constant from the AnalysisScope enum.
'@remarks Activating the sheet is what the freeze-pane API needs. A protected
'sheet or a hidden window makes it fail, and the failure is filed as a warning
'rather than passed over in silence.
'@remarks The busy state is written again after the activation, because
'activating a worksheet turns screen updating back on in some Excel versions.
'The scope is already busy at that point, which is why the call forces.
Private Sub SplitOutputSheet(ByVal scope As Byte)
    Dim outsh As Worksheet
    Dim rowSplit As Long

    Set outsh = OutputSh(scope)
    rowSplit = IIf(scope = AnalysisScopeTimeSeries Or scope = AnalysisScopeSpatioTemporal, 3, 2)

    On Error GoTo SplitFailed
    outsh.Activate
    AppScope.ApplyBusyState suppressEvents:=True, force:=True
    With Application.ActiveWindow
        .SplitRow = rowSplit
        .SplitColumn = 3
        .Zoom = 80
        .FreezePanes = True
    End With
    On Error GoTo 0
    Exit Sub

SplitFailed:
    LogInfo "Sheet " & outsh.Name & " has no frozen panes: " & Err.Description, _
            checkingWarning
End Sub

'@section Sheet Preparation
'===============================================================================

'@label:add-commands
'@sub-title Add the filter/update command button to an analysis output worksheet.
'@details
'Places the command button that runs either the standard filter-update macro or
'the spatial one, depending on the scope. The button sits at row 2 column 2, is
'styled with the design format and is labelled with the translated SHP_Filter
'shape name. Exits early when the linelist is built from a template, because
'the template already carries the buttons.
'
'This runs once per scope, from PrepareSheet, after the sheet format has been
'applied. It used to run a second time at the end of each scope; Buttons.Add
'short-circuits on an existing shape, so the second call stacked nothing and
'cost a Buttons instance, a shape scan, a translation and a format per scope.
'What the second call did lose was its own diagnostics, because the instance
'was created inside it and dropped.
'@param scope Byte. A constant from the AnalysisScope enum. The spatial scope
'gets the spatial update command.
Private Sub AddCommands(ByVal scope As Byte)
    Dim cellRng As Range
    Dim commandName As String
    Dim shapeLabel As String
    Dim design As LLFormat
    Dim but As Buttons
    Dim outsh As Worksheet
    Dim trads As TranslationObject
    Dim lData As LinelistSpecs

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

    If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues
End Sub

'@label:stamp-sheet-type
'@sub-title Write the sheet_type tag of an analysis output worksheet
'@details
'The tag is what every runtime reader of the linelist uses to tell one analysis
'sheet from another, so a sheet that ends up empty gets it too.
'@param scope Byte. A constant from the AnalysisScope enum.
Private Sub StampSheetType(ByVal scope As Byte)
    Dim shHn As HiddenNames

    Set shHn = OutputNames(scope)
    shHn.EnsureName "sheet_type", SheetTagOf(scope), HiddenNameTypeString
End Sub

'@label:prepare-sheet
'@sub-title Initialise an analysis output worksheet before any table is written.
'@details
'Four steps: apply the AllAnalysisSheet base format to the whole worksheet,
'stamp the sheet_type tag, place the ComputedOnFiltered() warning formula at
'row 2 column 5 with the warning format, and add the command button.
'
'The freeze-pane split is not part of this. Activating a worksheet forces a
'repaint that screen updating does not suppress, so the four splits are applied
'in one pass at the end of the build.
'@param scope Byte. A constant from the AnalysisScope enum. Decides the
'sheet_type tag and which update command the button carries.
Private Sub PrepareSheet(ByVal scope As Byte)
    Dim outsh As Worksheet
    Dim cellRng As Range
    Dim design As LLFormat

    Set outsh = OutputSh(scope)
    Set design = DesignFormat()

    design.ApplyFormat outsh, scope:=AllAnalysisSheet

    StampSheetType scope

    Set cellRng = outsh.Cells(2, 5)
    cellRng.Formula = "= ComputedOnFiltered()"
    design.ApplyFormat cellRng, scope:=LinelistWarning

    AddCommands scope:=scope
End Sub

'@section GoTo Dropdown Helpers
'===============================================================================

'@label:add-goto
'@sub-title Create a data-validation dropdown for section, header or graph navigation.
'@details
'Creates a data-validation dropdown in one cell that navigates to a named
'section, header or graph on the analysis sheet. Selecting an entry runs the
'GoTo macro that scrolls to the matching named range. The cell is formatted as
'a selection dropdown, its initial value is the label, and it is given a
'workbook-level name.
'
'An empty list leaves the cell without a dropdown and still names it.
'EventLinelist reads those names without guarding, and the one entry that used
'to keep the normal sheet's list non-empty was a dropdown entry that navigated
'nowhere.
'@param cellRng Range. The single cell the dropdown goes in.
'@param dropArray BetterArray. Display strings for the entries.
'@param goToLabel String. The default display value written into the cell.
'@param goToPrefix Optional String. Short prefix for the name, one per analysis
'sheet. Defaults to "ua_".
'@param goToSuffix Optional String. Suffix telling a section-level dropdown
'from a graph-level one. Defaults to vbNullString.
Private Sub AddGoTo(ByVal cellRng As Range, _
                    ByVal dropArray As BetterArray, _
                    ByVal goToLabel As String, _
                    Optional ByVal goToPrefix As String = "ua_", _
                    Optional ByVal goToSuffix As String = vbNullString)

    Dim drop As DropdownLists
    Dim ll As Linelist
    Dim dropName As String
    Dim design As LLFormat

    ' Five statements below can raise: the list add, the validation, the format,
    ' the value write and the name. One of the two call sites is inside
    ' BuildScope, which has nothing above it -- a raise there ends WriteAnalysis,
    ' and clickGenerate's cleanup restores the application state and skips the
    ' checkings flush, so the report the user needs to understand the failure is
    ' the thing the failure destroys. A missing dropdown is worth a warning and
    ' the rest of the sheet.
    On Error GoTo DropdownFailed

    Set ll = OutputLL()
    Set design = DesignFormat()

    If dropArray.Length > 0 Then
        Set drop = ll.Dropdown()
        dropName = goToPrefix & "gotosection" & goToSuffix
        drop.Add dropArray, dropName
        drop.SetValidation cellRng, dropName
        design.ApplyFormat cellRng, scope:=LinelistSelectionDropdown
        cellRng.Value = goToLabel
    End If

    cellRng.Name = goToPrefix & GOTOSECCODE & goToSuffix
    On Error GoTo 0
    Exit Sub

DropdownFailed:
    LogInfo "The navigation dropdown " & goToPrefix & GOTOSECCODE & goToSuffix & _
            " could not be built: " & Err.Description, checkingWarning
End Sub

'@section Temporal Section Helpers
'===============================================================================
'@description Writing out a temporal section once the loop has finished with it.

'@label:write-section-bounds
'@sub-title Write the two date bounds of a temporal section
'@details
'MIN_MIN_DATE_<secId> and MAX_MAX_DATE_<secId> are created by CrossTable in the
'new-section branch of its temporal header block, so a section whose anchor
'table failed leaves them absent and the write raises 1004. The absence is
'filed as a warning and the rest of the section is kept.
'@param outsh Worksheet. The analysis sheet the names live on.
'@param secId String. Identifier of the section.
'@param minFormula String. The complete formula for the earliest date.
'@param maxFormula String. The complete formula for the latest date.
Private Sub WriteSectionBounds(ByVal outsh As Worksheet, ByVal secId As String, _
                               ByVal minFormula As String, ByVal maxFormula As String)
    If LenB(minFormula) = 0 Then Exit Sub

    On Error GoTo BoundsMissing
    outsh.Range("MAX_MAX_DATE_" & secId).Formula = maxFormula
    outsh.Range("MIN_MIN_DATE_" & secId).Formula = minFormula
    On Error GoTo 0
    Exit Sub

BoundsMissing:
    LogInfo "Section " & secId & " has no date bounds to fill: " & Err.Description, _
            checkingWarning
End Sub

'@label:name-spatio-temporal-column
'@sub-title Name the formula column that spans a spatio-temporal section
'@details
'GeoModule reads SPT_FORMULA_COLUMN_<secId> at runtime. It spans from the
'column categories of the first table of the section to those of the last one,
'so it can only be named once the section is finished.
'@param outsh Worksheet. The analysis sheet.
'@param secId String. Identifier of the section.
'@param lastTabId String. Identifier of the last table of the section.
Private Sub NameSpatioTemporalColumn(ByVal outsh As Worksheet, ByVal secId As String, _
                                     ByVal lastTabId As String)
    Dim cellRng As Range
    Dim spatRng As Range

    If LenB(lastTabId) = 0 Then Exit Sub

    On Error GoTo ColumnMissing
    With outsh
        Set cellRng = .Range("COLUMN_CATEGORIES_" & lastTabId)
        Set spatRng = .Range( _
            .Range("COLUMN_CATEGORIES_" & secId).Cells(1, 1), _
            cellRng.Cells(1, cellRng.Columns.Count))
    End With
    spatRng.Name = "SPT_FORMULA_COLUMN_" & secId
    On Error GoTo 0
    Exit Sub

ColumnMissing:
    LogInfo "Section " & secId & " has no spatio-temporal formula column: " & Err.Description, _
            checkingWarning
End Sub

'@label:close-temporal-section
'@sub-title Finish a temporal section and write everything it held
'@details
'Writes the two date bounds, adds the header dropdown two rows under the
'section anchor, and names the spatio-temporal formula column. Does nothing
'when no section is open, so the caller can call it at every section boundary
'and once more after the loop.
'@param outsh Worksheet. The analysis sheet.
'@param section TemporalSection. The section holder.
'@param scope Byte. A constant from the AnalysisScope enum.
'@param goToHeaderLabel String. Translated label of the header dropdown.
Private Sub CloseTemporalSection(ByVal outsh As Worksheet, _
                                 ByVal section As TemporalSection, _
                                 ByVal scope As Byte, _
                                 ByVal goToHeaderLabel As String)
    Dim secId As String
    Dim lastTabId As String

    If Not section.IsOpen() Then Exit Sub

    secId = section.SectionId()
    lastTabId = section.LastTableId()
    section.EndSection

    WriteSectionBounds outsh, secId, section.MinFormula(), section.MaxFormula()
    AddSectionGoTo outsh, secId, section.Headers(), scope, goToHeaderLabel

    If scope = AnalysisScopeSpatioTemporal Then _
        NameSpatioTemporalColumn outsh, secId, lastTabId
End Sub

'@label:add-section-go-to
'@sub-title Hang the header dropdown of one section under its anchor
'@details
'The anchor is SECTION_<secId>, which CrossTable writes only in its new-section
'branch, so a section whose anchor table failed has none. That failure used to
'end CloseTemporalSection, which took NameSpatioTemporalColumn with it and left
'SPT_FORMULA_COLUMN_<secId> uncreated -- a name GeoModule reads at runtime in the
'delivered linelist. The two are guarded apart now, the way WriteSectionBounds
'already was.
'@param outsh Worksheet. The analysis sheet.
'@param secId String. Identifier of the section.
'@param headers BetterArray. The entries of the dropdown.
'@param scope Byte. A constant from the AnalysisScope enum.
'@param goToHeaderLabel String. Translated label of the dropdown.
Private Sub AddSectionGoTo(ByVal outsh As Worksheet, ByVal secId As String, _
                           ByVal headers As BetterArray, ByVal scope As Byte, _
                           ByVal goToHeaderLabel As String)
    Dim cellRng As Range

    On Error GoTo AnchorMissing
    Set cellRng = outsh.Range("SECTION_" & secId).Offset(2)
    On Error GoTo 0

    AddGoTo cellRng:=cellRng, dropArray:=headers, _
            goToLabel:=goToHeaderLabel, goToPrefix:=GoToPrefixOf(scope), _
            goToSuffix:=secId
    Exit Sub

AnchorMissing:
    LogInfo "Section " & secId & " has no anchor to hang its header dropdown on: " & _
            Err.Description, checkingWarning
End Sub

'@label:name-graph-start
'@sub-title Name the cell the charts of a temporal sheet start at
'@details
'The charts go five rows under the last table of the sheet. Both chart passes
'read this name, and it is created once the tables are done.
'@param outsh Worksheet. The analysis sheet.
'@param secId String. Identifier of the last section of the sheet.
'@param scope Byte. A constant from the AnalysisScope enum.
Private Sub NameGraphStart(ByVal outsh As Worksheet, ByVal secId As String, _
                           ByVal scope As Byte)
    Dim cellRng As Range

    If LenB(secId) = 0 Then Exit Sub

    On Error GoTo EndTableMissing
    Set cellRng = outsh.Range("ENDTABLE_" & secId).Offset(5, 2)
    On Error GoTo 0

    If scope = AnalysisScopeTimeSeries Then cellRng.Name = "START_GRAPHS_TS"
    If scope = AnalysisScopeSpatioTemporal Then cellRng.Name = "START_GRAPHS_SPT"
    Exit Sub

EndTableMissing:
    LogInfo "Sheet " & outsh.Name & " has no end-of-table anchor for its charts: " & _
            Err.Description, checkingWarning
End Sub

'@section Table Creation
'===============================================================================
'@description Core table-creation loop and supporting write logic.

'@label:write-table
'@sub-title Iterate setup ListObject rows and build cross-tables on the output sheet.
'@details
'Core table-creation loop. Walks every data row of the setup ListObject named
'by loName and, for each valid row, creates a TableSpecs, builds a CrossTable
'on the output sheet, applies the formulas, formats the table and, for a row
'that asks for one, creates a chart. For the time series and spatio-temporal
'scopes it also carries the open TemporalSection: a row that starts a new
'section closes the one before it, and the last section is closed after the
'loop.
'
'Each row runs inside a per-table error handler, so one malformed row is
'skipped and named in the generation report while the rest of the scope is
'built.
'@param loName String. Name of the setup ListObject whose rows define the tables.
'@param sectionsList BetterArray. Collects the section entries of the
'sheet-level navigation dropdown. Passed ByVal and grown through Push.
'@param scope Byte. A constant from the AnalysisScope enum. Decides which
'output sheet receives the tables.
'@remarks Exits early when the ListObject has no data row.
Private Sub WriteTable(ByVal loName As String, _
                       ByVal sectionsList As BetterArray, _
                       ByVal scope As Byte)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim tabRng As Range
    Dim hRng As Range
    Dim seriesCounter As Long
    Dim tabl As CrossTable
    Dim tabform As CrossTableFormula
    Dim gr As Graphs
    Dim lo As ListObject
    Dim specs As TableSpecs
    Dim lData As LinelistSpecs
    Dim trads As TranslationObject
    Dim formData As FormulaData
    Dim cellRng As Range
    Dim grSeries As GraphSeries
    Dim buffer As SeriesBuffer
    Dim tabId As String
    Dim tabType As Byte
    Dim design As LLFormat
    Dim anchorRng As Range
    Dim containsTimeSeries As Boolean
    Dim namObj As AnaTabIds
    Dim goToSectionLabel As String
    Dim goToHeaderLabel As String
    Dim outshHn As HiddenNames
    Dim hasTable As Boolean

    ' The last table this loop built. CrossTable.Previous reads the table above
    ' it to place a continuation table, and building that ancestor from scratch
    ' costs a cross-table and a specification walk per table. The loop already
    ' holds it, so it hands it over. CrossTable compares the identifiers before
    ' it trusts the gift: a row that passes validation and then fails inside
    ' Build is still the nearest valid row above, so the two can disagree.
    Dim prevTable As CrossTable

    ' Temporal state
    Dim section As TemporalSection
    Dim lastSecId As String
    Dim sectionLabel As String
    Dim openSectionLabel As String
    Dim tableName As String
    Dim rowVar As String
    Dim dict As LLdictionary
    Dim vars As LLVariables
    Dim grHeight As Long
    Dim isNewSection As Boolean

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
    Set outshHn = OutputNames(scope)
    Set section = TemporalSection.Create()

    goToSectionLabel = trads.TranslatedValue("MSG_GoToSec")
    goToHeaderLabel = trads.TranslatedValue("MSG_GoToHead")
    hasTable = (LenB(outshHn.ValueAsString("has_table")) > 0)

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set hRng = lo.HeaderRowRange()
    Set anchorRng = hRng.Offset(lo.ListRows.Count)
    Set tabRng = lo.ListRows(1).Range

    Do
        ' The handler is armed before the specification is read, so a malformed
        ' setup row is skipped like any other bad table. Reading it outside the
        ' guarded window took the whole analysis down on one unreadable row.
        ' Until TableId is known the row number stands in for it, because the
        ' handler names the table it lost and an empty name would point at the
        ' table before this one.
        tabId = "row " & tabRng.Row
        On Error GoTo TableError

        Set specs = TableSpecs.Create(hRng, tabRng, dict)

        If Not specs.ValidTable() Then GoTo ContinueLoop

        tabType = specs.TableScope()
        isNewSection = specs.IsNewSection()
        containsTimeSeries = (tabType = ScopeTimeSeries Or tabType = ScopeSpatioTemporal)
        tabId = specs.TableId()

        ' The list belongs to the scope being built, so two scopes can never
        ' mix their sections into one dropdown.
        If isNewSection Then sectionsList.Push goToSectionLabel & ": " & specs.Value("section")

        Set tabl = CrossTable.Create(specs, outsh, lData, previousTable:=prevTable, _
                                     outputNames:=outshHn)
        tabl.Build
        Set prevTable = tabl
        Set tabform = CrossTableFormula.Create(tabl, formData)
        tabform.AddFormulas
        tabl.Format design

        ' The handler stays armed over the rest of the table. Disarming it here
        ' left the checkings harvest, the charts and the section logic
        ' unguarded, so a failure in any of them escaped the per-table skip and
        ' ended the run.

        ' Harvest cross-table checkings
        If tabl.HasCheckings Then HarvestCollaboratorCheckings tabl.CheckingValues

        ' And the formula writer's. Without this line the report showed the
        ' analysis phase clean while a table carried a cell Excel had refused.
        If tabform.HasCheckings Then HarvestCollaboratorCheckings tabform.CheckingValues

        ' The table and its formulas are on the sheet, so the milestone is
        ' true. A step after this one that fails is named by TableError.
        LogMilestone "table " & tabId & " written on " & outsh.Name
        this.tablesWritten = this.tablesWritten + 1

        ' Mark the sheet as having at least one table. The name is written once
        ' per sheet; asking the worksheet per row cost a name scan and a
        ' RefersTo read on every table.
        If Not hasTable Then
            outshHn.EnsureName "has_table", "yes", HiddenNameTypeString
            hasTable = True
        End If

        ' Non-temporal charts
        If specs.HasGraph() And (Not containsTimeSeries) Then
            ' GeoCount carries the default of 5 and clamps to 1..20. Reading
            ' the cell raw raised type mismatch 13 on an empty one.
            grHeight = specs.GeoCount()

            Set cellRng = outsh.Cells(tabl.StartRow - 2, tabl.EndColumn + 2)
            If tabType = ScopeSpatial Then Set cellRng = cellRng.Offset(, 2)

            Set grSeries = GraphSeries.Create(tabl)
            Set buffer = grSeries.Series()

            ' And the series builder's. It decides to build fewer series than
            ' the setup asked for in four places, and every one of them was
            ' silent until this line.
            If grSeries.HasCheckings Then HarvestCollaboratorCheckings grSeries.CheckingValues

            Set gr = Graphs.Create(outsh, cellRng)
            gr.Add

            For seriesCounter = 1 To buffer.Count
                AddGraphSeries gr:=gr, namObj:=namObj, scope:=scope, graphId:=tabId, _
                               seriesName:=buffer.RangeName(seriesCounter), _
                               seriesType:=buffer.ChartType(seriesCounter), _
                               seriesPos:=buffer.AxisSide(seriesCounter), _
                               seriesLabel:=buffer.CategoryRange(seriesCounter), _
                               seriesColumnLabel:=buffer.LegendRange(seriesCounter), _
                               hardCodeLabels:=False, _
                               outRangeAddress:=cellRng.Address
            Next

            FinishGraph gr:=gr, namObj:=namObj, scope:=scope, graphId:=tabId, _
                        catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                        valuesTitle:=specs.Value("label"), _
                        hardCodeLabels:=False, heightFactor:=grHeight, _
                        plotTitle:=vbNullString

            LogMilestone "chart of table " & tabId & " written on " & outsh.Name
            this.graphsWritten = this.graphsWritten + 1
        End If

        ' Time series and spatio-temporal section tracking
        If Not containsTimeSeries Then GoTo ContinueLoop

        ' A row that starts a section closes the one being held. A section that
        ' is not open yet is started whatever the row says, because every table
        ' of a temporal sheet belongs to some section.
        '
        ' The third test is issue #337. IsNewSection compares against the
        ' nearest row above that the build will DRAW, and a row can pass
        ' validation and still fail inside Build -- which lands on TableError and
        ' skips StartSection. The next row of that section then answers
        ' IsNewSection = False against an anchor that was never opened, finds a
        ' section still open, and appends its header to the section BEFORE it.
        ' Comparing the section label as well closes the section that is open
        ' whenever the label moves, so a failed anchor costs its own section a
        ' warning rather than costing two sections the wrong dropdown.
        sectionLabel = specs.Value("section")

        If isNewSection Or (Not section.IsOpen()) Or _
           (StrComp(sectionLabel, openSectionLabel, vbTextCompare) <> 0) Then
            CloseTemporalSection outsh, section, scope, goToHeaderLabel
            section.StartSection tabId
            openSectionLabel = sectionLabel
        End If

        rowVar = specs.Value("row")
        tableName = vars.Value(colName:="table name", varName:=rowVar)
        section.AddTable tabId:=tabId, tableName:=tableName, rowVar:=rowVar, _
                         header:=goToHeaderLabel & ": " & tabl.TimeSeriesHeader()

        GoTo ContinueLoop

TableError:
        LogInfo "Table " & tabId & " failed: " & Err.Description, checkingError
        ' Resume clears the error and hands control back to the loop, which
        ' arms the handler again for the next table. Disabling the handler
        ' here first left Resume with no error to clear, so a single bad
        ' table took the whole run down instead of being skipped.
        Resume ContinueLoop

ContinueLoop:
        Set tabRng = tabRng.Offset(1)
    Loop While tabRng.Row <= anchorRng.Row

    ' The last section of the sheet, and the anchor its charts hang from.
    If section.IsOpen() Then
        lastSecId = section.SectionId()
        CloseTemporalSection outsh, section, scope, goToHeaderLabel
        NameGraphStart outsh, lastSecId, scope
    End If
End Sub

'@section Graph Creation
'===============================================================================
'@description The two chart passes that run after the tables of a sheet are
'written, and the two steps they share.

'@label:add-graph-series
'@sub-title Add one series to a chart and register it for the export
'@details
'The chart and the registry have to be told the same thing about a series, and
'three places told them separately. AnaTabIds.AddGraphInfo takes nine required
'arguments; one of those three left the ninth out, which is a compile fault in
'an early-bound call, so the whole file stopped compiling. One call site cannot
'make that mistake twice.
'@param gr Graphs. The chart being built.
'@param namObj AnaTabIds. The registry the export reads back.
'@param scope Byte. A constant from the AnalysisScope enum.
'@param graphId String. Identifier of the chart.
'@param seriesName String. Named range holding the series data.
'@param seriesType String. Chart type of the series.
'@param seriesPos String. Axis the series is plotted against.
'@param seriesLabel String. Named range holding the category labels.
'@param seriesColumnLabel String. Named range holding the column header labels.
'@param hardCodeLabels Boolean. True writes the labels as literals.
'@param outRangeAddress String. Address the chart is placed at.
'@param prefix Optional String. Label prefix. Defaults to vbNullString.
'@param prefixOnly Optional Boolean. True uses the prefix as the whole label.
Private Sub AddGraphSeries(ByVal gr As Graphs, ByVal namObj As AnaTabIds, _
                           ByVal scope As Byte, ByVal graphId As String, _
                           ByVal seriesName As String, ByVal seriesType As String, _
                           ByVal seriesPos As String, ByVal seriesLabel As String, _
                           ByVal seriesColumnLabel As String, _
                           ByVal hardCodeLabels As Boolean, _
                           ByVal outRangeAddress As String, _
                           Optional ByVal prefix As String = vbNullString, _
                           Optional ByVal prefixOnly As Boolean = False)

    gr.AddSeries rngName:=seriesName, chrtType:=seriesType, axisPos:=seriesPos

    If prefixOnly Then
        gr.AddLabels catName:=seriesLabel, lblName:=seriesColumnLabel, _
                     hardCodeLabels:=hardCodeLabels, prefix:=prefix, prefixOnly:=True
    Else
        gr.AddLabels catName:=seriesLabel, lblName:=seriesColumnLabel, _
                     hardCodeLabels:=hardCodeLabels
    End If

    namObj.AddGraphInfo scope:=scope, graphId:=graphId, _
                        seriesName:=seriesName, seriesType:=seriesType, _
                        seriesPos:=seriesPos, seriesLabel:=seriesLabel, _
                        seriesColumnLabel:=seriesColumnLabel, _
                        hardCodeLabels:=hardCodeLabels, _
                        outRangeAddress:=outRangeAddress, _
                        prefix:=prefix, prefixOnly:=prefixOnly
End Sub

'@label:finish-graph
'@sub-title Format a chart, register its look and harvest what it reported
'@details
'The second half every chart shares. The chart scope comes from the analysis
'scope through GraphScopeOf, so a chart is never formatted against a number
'that happens to line up.
'@param gr Graphs. The chart being finished.
'@param namObj AnaTabIds. The registry the export reads back.
'@param scope Byte. A constant from the AnalysisScope enum.
'@param graphId String. Identifier of the chart.
'@param catTitle String. Category axis title.
'@param valuesTitle String. Value axis title.
'@param hardCodeLabels Boolean. True writes the labels as literals.
'@param heightFactor Long. Height multiplier of the chart.
'@param plotTitle String. Chart title.
Private Sub FinishGraph(ByVal gr As Graphs, ByVal namObj As AnaTabIds, _
                        ByVal scope As Byte, ByVal graphId As String, _
                        ByVal catTitle As String, ByVal valuesTitle As String, _
                        ByVal hardCodeLabels As Boolean, _
                        ByVal heightFactor As Long, ByVal plotTitle As String)

    gr.Format catTitle:=catTitle, valuesTitle:=valuesTitle, _
              hardCodeLabels:=hardCodeLabels, scope:=GraphScopeOf(scope), _
              heightFactor:=heightFactor, plotTitle:=plotTitle

    namObj.AddGraphFormat scope:=scope, graphId:=graphId, _
                          catTitle:=catTitle, valuesTitle:=valuesTitle, _
                          hardCodeLabels:=hardCodeLabels, _
                          heightFactor:=heightFactor, plotTitle:=plotTitle

    If gr.HasCheckings Then HarvestCollaboratorCheckings gr.CheckingValues
End Sub

'@label:graph-anchor
'@sub-title The cell the charts of a sheet start at
'@details
'The name is created after the last table of the sheet is written, so a scope
'whose tables all failed does not carry it. Answering Nothing is what lets the
'caller give up on the charts rather than end the whole build.
'@param outsh Worksheet. The analysis sheet.
'@param anchorName String. Name of the anchor cell.
'@return Range. The anchor cell, or Nothing when the name is absent.
Private Function GraphAnchor(ByVal outsh As Worksheet, ByVal anchorName As String) As Range
    On Error GoTo AnchorMissing
    Set GraphAnchor = outsh.Range(anchorName)
    On Error GoTo 0
    Exit Function

AnchorMissing:
    LogInfo "Sheet " & outsh.Name & " carries no chart anchor " & anchorName & _
            ", so its charts are skipped", checkingWarning
End Function

'@label:write-time-series-graphs
'@sub-title Build all time series charts from the graph definition ListObjects.
'@details
'Reads three setup ListObjects: the graph definition table, the graph title
'table and the time series table. TimeSeriesGraphs resolves them into
'multi-series graphs with mixed chart types and separate axis positions.
'Each graph is then created, its series added and registered, formatted and
'given its title cell, and the anchor moves down fifty rows. A graph dropdown
'is placed at row 3 column 3 at the end.
'@param graphLoName String. Name of the ListObject defining the graph series.
'@param graphTitleLoName String. Name of the ListObject defining the titles.
'@param tabLoName String. Name of the time series table ListObject.
'@param graphsList BetterArray. Collects the graph entries of the dropdown.
'@remarks Exits early when the sheet holds no table, when either definition
'ListObject has no data row, or when the charts have no anchor to start from.
Private Sub WriteTimeSeriesGraphs(ByVal graphLoName As String, _
                                  ByVal graphTitleLoName As String, _
                                  ByVal tabLoName As String, _
                                  ByVal graphsList As BetterArray)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim graphCounter As Long
    Dim cellRng As Range
    Dim lData As LinelistSpecs
    Dim seriesCounter As Long

    Dim grLo As ListObject
    Dim tabLo As ListObject
    Dim grTitleLo As ListObject
    Dim tsGraphs As TimeSeriesGraphs
    Dim buffer As SeriesBuffer
    Dim gr As Graphs
    Dim trads As TranslationObject
    Dim design As LLFormat
    Dim namObj As AnaTabIds
    Dim goToGraphLabel As String
    Dim periodTitle As String
    Dim valuesTitle As String

    Dim grBAinput As BetterArray
    Dim plotTitle As String

    Dim outshHn As HiddenNames

    Set sh = Wksh()
    Set outsh = OutputSh(AnalysisScopeTimeSeries)
    Set outshHn = OutputNames(AnalysisScopeTimeSeries)

    If LenB(outshHn.ValueAsString("has_table")) = 0 Then Exit Sub

    Set grLo = sh.ListObjects(graphLoName)
    Set tabLo = sh.ListObjects(tabLoName)
    Set grTitleLo = sh.ListObjects(graphTitleLoName)

    ' Asked before the anchor is resolved. A sheet that holds tables but no
    ' graph definition has nothing to draw, and reading the anchor first turned
    ' that into a 1004.
    If (grLo.DataBodyRange Is Nothing) Or (grTitleLo.DataBodyRange Is Nothing) Then Exit Sub

    Set cellRng = GraphAnchor(outsh, "START_GRAPHS_TS")
    If cellRng Is Nothing Then Exit Sub

    Set grBAinput = New BetterArray
    Set namObj = NamesObject()
    Set lData = LinelistData()
    Set trads = TransObject()
    Set design = DesignFormat()

    ' Both axis titles are the same for every chart of the sheet. They were
    ' read four times per chart, and every read used to walk the whole
    ' translation table.
    goToGraphLabel = trads.TranslatedValue("MSG_GoToGraph")
    periodTitle = trads.TranslatedValue("MSG_Period")
    valuesTitle = trads.TranslatedValue("MSG_Values")

    grBAinput.Push grLo, tabLo, grTitleLo
    Set tsGraphs = TimeSeriesGraphs.Create(grBAinput, outsh, lData)

    ' A time series series is labelled by its prefix alone, so the chart and the
    ' registry both read the prefix and leave the column label empty.
    For graphCounter = 1 To tsGraphs.Count
        On Error GoTo GraphError

        Set buffer = tsGraphs.Series(graphCounter)
        plotTitle = tsGraphs.Title(graphCounter)

        Set gr = Graphs.Create(outsh, cellRng)
        gr.Add

        For seriesCounter = 1 To buffer.Count
            AddGraphSeries gr:=gr, namObj:=namObj, _
                           scope:=AnalysisScopeTimeSeries, _
                           graphId:=CStr(graphCounter), _
                           seriesName:=buffer.RangeName(seriesCounter), _
                           seriesType:=buffer.ChartType(seriesCounter), _
                           seriesPos:=buffer.AxisSide(seriesCounter), _
                           seriesLabel:=buffer.CategoryRange(seriesCounter), _
                           seriesColumnLabel:=vbNullString, _
                           hardCodeLabels:=True, _
                           outRangeAddress:=cellRng.Address, _
                           prefix:=buffer.LabelPrefix(seriesCounter), _
                           prefixOnly:=True
        Next

        FinishGraph gr:=gr, namObj:=namObj, scope:=AnalysisScopeTimeSeries, _
                    graphId:=CStr(graphCounter), catTitle:=periodTitle, _
                    valuesTitle:=valuesTitle, hardCodeLabels:=True, _
                    heightFactor:=1, plotTitle:=plotTitle

        cellRng.Value = plotTitle
        design.ApplyFormat cellRng, scope:=LinelistHiddenCell
        graphsList.Push goToGraphLabel & ": " & plotTitle

        LogMilestone "time series chart " & CStr(graphCounter) & _
                     " written on " & outsh.Name
        this.graphsWritten = this.graphsWritten + 1

        On Error GoTo 0
        GoTo NextGraph

GraphError:
        LogInfo "Time series chart " & graphCounter & " failed: " & Err.Description, _
                checkingError
        Resume NextGraph

NextGraph:
        Set cellRng = cellRng.Offset(GRAPHROWSTEP)
    Next

    ' Read after the loop, so a run that built no graph at all still reports
    ' why. Count is what triggers the build, so the entries exist by here.
    If tsGraphs.HasCheckings Then HarvestCollaboratorCheckings tsGraphs.CheckingValues

    ' Add GoTo graph dropdown
    AddGoTo outsh.Cells(3, 3), dropArray:=graphsList, _
            goToLabel:=goToGraphLabel, goToPrefix:="ts_", goToSuffix:="_graph"
End Sub

'@label:write-spatio-temporal-graph
'@sub-title Build the spatio-temporal charts from the table definition ListObject.
'@details
'Walks the same setup ListObject the spatio-temporal tables were built from and
'creates one chart per row that asks for one. The plot title is the translated
'graph label and the value of the table's start column. Charts are spaced fifty
'rows apart, and a graph dropdown is placed at row 3 column 3 at the end.
'@param loName String. Name of the spatio-temporal ListObject.
'@param graphsList BetterArray. Collects the graph entries of the dropdown.
'@remarks Skips a row that is not a valid table or asks for no chart. Exits
'early when the ListObject has no data row or the charts have no anchor.
Private Sub WriteSpatioTemporalGraph(ByVal loName As String, _
                                     ByVal graphsList As BetterArray)
    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim cellRng As Range
    Dim specs As TableSpecs
    Dim lData As LinelistSpecs
    Dim design As LLFormat
    Dim hRng As Range
    Dim tabRng As Range
    Dim anchorRng As Range
    Dim gr As Graphs
    Dim grSeries As GraphSeries
    Dim buffer As SeriesBuffer
    Dim plotTitle As String
    Dim tabId As String
    Dim seriesCounter As Long
    Dim trads As TranslationObject
    Dim lo As ListObject
    Dim tabl As CrossTable
    Dim namObj As AnaTabIds
    Dim goToGraphLabel As String
    Dim graphLabel As String
    Dim dict As LLdictionary

    Set sh = Wksh()
    Set outsh = OutputSh(AnalysisScopeSpatioTemporal)
    Set lo = sh.ListObjects(loName)
    Set lData = LinelistData()
    Set dict = lData.Dictionary()
    Set trads = TransObject()
    Set design = DesignFormat()
    Set namObj = NamesObject()
    goToGraphLabel = trads.TranslatedValue("MSG_GoToGraph")
    graphLabel = trads.TranslatedValue("MSG_Graph")

    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set cellRng = GraphAnchor(outsh, "START_GRAPHS_SPT")
    If cellRng Is Nothing Then Exit Sub

    Set hRng = lo.HeaderRowRange()
    Set anchorRng = hRng.Offset(lo.ListRows.Count)
    Set tabRng = lo.ListRows(1).Range

    Do
        tabId = "row " & tabRng.Row
        On Error GoTo GraphError

        Set specs = TableSpecs.Create(hRng, tabRng, dict)

        If Not specs.ValidTable() Then GoTo ContinueLoop
        If Not specs.HasGraph() Then GoTo ContinueLoop

        tabId = specs.TableId()

        ' The cross-table is created for its specification alone. The series
        ' builder takes the column count from the geographic unit count for
        ' this scope, so it never reads the geometry a Build would have set.
        Set tabl = CrossTable.Create(specs, outsh, lData)
        Set grSeries = GraphSeries.Create(tabl)

        plotTitle = graphLabel & ": " & outsh.Range("STARTCOL_" & tabId).Value

        Set buffer = grSeries.Series()

        ' The series builder reports every series it left out.
        If grSeries.HasCheckings Then HarvestCollaboratorCheckings grSeries.CheckingValues

        Set gr = Graphs.Create(outsh, cellRng)
        gr.Add

        For seriesCounter = 1 To buffer.Count
            AddGraphSeries gr:=gr, namObj:=namObj, _
                           scope:=AnalysisScopeSpatioTemporal, graphId:=tabId, _
                           seriesName:=buffer.RangeName(seriesCounter), _
                           seriesType:=buffer.ChartType(seriesCounter), _
                           seriesPos:=buffer.AxisSide(seriesCounter), _
                           seriesLabel:=buffer.CategoryRange(seriesCounter), _
                           seriesColumnLabel:=buffer.LegendRange(seriesCounter), _
                           hardCodeLabels:=False, _
                           outRangeAddress:=cellRng.Address
        Next

        FinishGraph gr:=gr, namObj:=namObj, scope:=AnalysisScopeSpatioTemporal, _
                    graphId:=tabId, catTitle:="LABEL_ROW_CATEGORIES_" & tabId, _
                    valuesTitle:=specs.Value("label"), hardCodeLabels:=False, _
                    heightFactor:=1, plotTitle:=plotTitle

        cellRng.Value = plotTitle
        design.ApplyFormat cellRng, scope:=LinelistHiddenCell
        graphsList.Push goToGraphLabel & ": " & plotTitle

        LogMilestone "spatio-temporal chart of table " & tabId & _
                     " written on " & outsh.Name
        this.graphsWritten = this.graphsWritten + 1

        Set cellRng = cellRng.Offset(GRAPHROWSTEP)

        GoTo ContinueLoop

GraphError:
        LogInfo "Spatio-temporal chart " & tabId & " failed: " & Err.Description, _
                checkingError
        Resume ContinueLoop

ContinueLoop:
        Set tabRng = tabRng.Offset(1)
    Loop While tabRng.Row <= anchorRng.Row

    AddGoTo outsh.Cells(3, 3), dropArray:=graphsList, _
            goToLabel:=goToGraphLabel, goToPrefix:="spt_", goToSuffix:="_graph"
End Sub

'@section Orchestration
'===============================================================================
'@description Top-level entry point that builds the four analysis output
'worksheets.

'@label:scope-has-rows
'@sub-title Whether any setup ListObject of a scope carries a row
'@details
'Formatting an analysis sheet writes a font size and a column width over all
'sixteen thousand columns, and a scope with no row is hidden at the end of its
'own build anyway.
'@param loNames BetterArray. Names of the setup ListObjects of the scope.
'@return Boolean. True when at least one of them has a data row.
Private Function ScopeHasRows(ByVal loNames As BetterArray) As Boolean
    Dim sh As Worksheet
    Dim counter As Long
    Dim lo As ListObject

    Set sh = Wksh()

    For counter = loNames.LowerBound To loNames.UpperBound
        Set lo = Nothing
        ' A setup workbook that predates one of these tables has no such
        ' ListObject, and that is a reason to skip the name rather than to end
        ' the build.
        On Error Resume Next
        Set lo = sh.ListObjects(CStr(loNames.Item(counter)))
        On Error GoTo 0

        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                ScopeHasRows = True
                Exit Function
            End If
        End If
    Next
End Function

'@label:build-scope
'@sub-title Build one analysis sheet from end to end
'@details
'The four scopes were four copies of the same twenty lines. Each one prepares
'its sheet, writes its tables, adds its charts, hangs the section dropdown at
'row 2 column 3, protects the sheet and hides it when nothing was written.
'@param scope Byte. A constant from the AnalysisScope enum.
'@param loNames BetterArray. Names of the setup ListObjects to walk, in order.
'@param buildGraphs Boolean. True runs the chart pass of the scope.
Private Sub BuildScope(ByVal scope As Byte, ByVal loNames As BetterArray, _
                       ByVal buildGraphs As Boolean)
    Dim outsh As Worksheet
    Dim outshHn As HiddenNames
    Dim sectionsList As BetterArray
    Dim graphsList As BetterArray
    Dim pass As Passwords
    Dim counter As Long
    Dim goToSectionLabel As String

    Set outsh = OutputSh(scope)
    Set outshHn = OutputNames(scope)
    Set pass = LinelistData().Password()
    goToSectionLabel = TransObject().TranslatedValue("MSG_GoToSec")

    ' A scope with no specification row gets its tag and nothing else.
    If Not ScopeHasRows(loNames) Then
        StampSheetType scope
        outsh.Visible = xlSheetHidden
        Exit Sub
    End If

    PrepareSheet scope:=scope

    Set sectionsList = New BetterArray
    Set graphsList = New BetterArray

    For counter = loNames.LowerBound To loNames.UpperBound
        WriteTable loName:=CStr(loNames.Item(counter)), _
                   sectionsList:=sectionsList, scope:=scope
    Next

    If buildGraphs Then
        Select Case scope
        Case AnalysisScopeTimeSeries
            WriteTimeSeriesGraphs graphLoName:=GRAPHTIMESERIES, _
                                  graphTitleLoName:=TITLEGRAPHTIMESERIES, _
                                  tabLoName:=TIMESERIES, graphsList:=graphsList
        Case AnalysisScopeSpatioTemporal
            WriteSpatioTemporalGraph loName:=SPATIOTEMPORAL, graphsList:=graphsList
        End Select
    End If

    AddGoTo outsh.Cells(2, 3), dropArray:=sectionsList, _
            goToLabel:=goToSectionLabel, goToPrefix:=GoToPrefixOf(scope)

    pass.Protect outsh, True
    If LenB(outshHn.ValueAsString("has_table")) = 0 Then outsh.Visible = xlSheetHidden
End Sub

'@label:scope-tables
'@sub-title The setup ListObjects one scope is built from
'@param scope Byte. A constant from the AnalysisScope enum.
'@return BetterArray. The ListObject names, in the order they are written.
Private Function ScopeTables(ByVal scope As Byte) As BetterArray
    Dim tableNames As BetterArray

    Set tableNames = New BetterArray
    tableNames.LowerBound = 1

    Select Case scope
    Case AnalysisScopeNormal
        tableNames.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE
    Case AnalysisScopeTimeSeries
        tableNames.Push TIMESERIES
    Case AnalysisScopeSpatial
        tableNames.Push SPATIAL
    Case AnalysisScopeSpatioTemporal
        tableNames.Push SPATIOTEMPORAL
    End Select

    Set ScopeTables = tableNames
End Function

'@label:write-analysis
'@sub-title Write all analysis tables, formulas and charts to the output.
'@details
'Builds the four analysis sheets in order: normal (global summary, univariate
'and bivariate), time series, spatial and spatio-temporal. The stage says how
'far to go, and every sheet up to that point is finished, protected and hidden
'when it stayed empty.
'
'The freeze-pane split of all four sheets runs at the end, in one pass.
'Activating a worksheet forces a repaint and a scroll sync that screen updating
'does not suppress, and it used to happen in the middle of every scope.
'@param part Optional Byte. A constant from the AnalysisBuildStage enum.
'Defaults to AnalysisBuildStageAll, which builds the four sheets. Any other
'value stops after the stage it names.
'@export
'@remarks This is what opens and restores the application state the build runs
'under: see the note at the top of this class. A raise inside the build restores
'it before the error travels on, so no failure leaves Excel with a frozen
'screen.
'@throws Whatever the build raises, after the application state is restored.
Public Sub WriteAnalysis(Optional ByVal part As Byte = AnalysisBuildStageAll)
    On Error GoTo Cleanup
    AppScope.ApplyBusyState suppressEvents:=True

    ' Univariate and bivariate analysis =======================================
    BuildScope AnalysisScopeNormal, ScopeTables(AnalysisScopeNormal), False
    LogInfo "Univariate and bivariate analysis completed"
    If part = AnalysisBuildStageNormal Then GoTo SplitSheets

    ' Time series analysis ====================================================
    BuildScope AnalysisScopeTimeSeries, ScopeTables(AnalysisScopeTimeSeries), _
               (part <> AnalysisBuildStageTimeSeriesTables)

    If part = AnalysisBuildStageTimeSeriesTables Then
        LogInfo "Time series tables completed"
        GoTo SplitSheets
    End If

    LogInfo "Time series analysis completed (tables and charts)"
    If part = AnalysisBuildStageTimeSeries Then GoTo SplitSheets

    ' Spatial analysis ========================================================
    BuildScope AnalysisScopeSpatial, ScopeTables(AnalysisScopeSpatial), False
    LogInfo "Spatial analysis completed"
    If part = AnalysisBuildStageSpatial Then GoTo SplitSheets

    ' Spatio-temporal analysis ================================================
    BuildScope AnalysisScopeSpatioTemporal, ScopeTables(AnalysisScopeSpatioTemporal), True
    LogInfo "Spatio-temporal analysis completed"

SplitSheets:
    SplitOutputSheet AnalysisScopeNormal
    SplitOutputSheet AnalysisScopeTimeSeries
    SplitOutputSheet AnalysisScopeSpatial
    SplitOutputSheet AnalysisScopeSpatioTemporal

    'The handler stays armed over the restore: a restore that fails is caught
    'below rather than escaping with the screen still off.
    AppScope.Restore
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String

    errNumber = Err.Number
    errDesc = Err.Description

    'Restore silently: the build's own failure is the one worth reporting, and a
    'restore that raises on top of it would hide it.
    On Error Resume Next
    AppScope.Restore silent:=True
    On Error GoTo 0

    ThrowError errNumber, errDesc
End Sub

'@section Checkings
'===============================================================================

'@label:log-info
'@sub-title Record one entry in this class's own checking store
'@details
'The key names this class. Every collaborator of a build files its entries into
'one report, Checking.Add raises on a duplicate key, and a bare counter
'starting at zero is what each of them used to write.
'@param label String. The message to record.
'@param scope Optional Byte. Severity from CheckingLogType. Defaults to checkingSuccess.
Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingSuccess)
    If internalChecks Is Nothing Then
        Set internalChecks = Checking.Create("Analysis output")
    End If
    internalChecks.Add CLASS_NAME & "-" & CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    this.hasChecks = True
End Sub

'@label:harvest-collaborator-checkings
'@sub-title Fold a collaborator's entries into the report of this build
'@param collaboratorChecks Checking. The collaborator's store.
Private Sub HarvestCollaboratorCheckings(ByVal collaboratorChecks As Checking)
    If collaboratorChecks Is Nothing Then Exit Sub
    If internalChecks Is Nothing Then
        Set internalChecks = Checking.Create("Analysis output")
    End If
    internalChecks.Append collaboratorChecks
    this.hasChecks = True
End Sub

'@label:log-milestone
'@sub-title File one entry for a table or a chart put on a sheet
'@details
'A linelist carries tens of analysis tables and charts, so these entries
'go to the __check worksheet with the problems. They are what shows the
'analysis phase did work, where a clean phase used to look the same
'whether it built forty tables or none.
'@param label String. The message for the report.
Private Sub LogMilestone(ByVal label As String)
    LogInfo label, checkingInfo
End Sub

'@label:tables-written
'@prop-title How many analysis tables this build put on a sheet
'@return Long. The number of tables written.
'@export
Public Property Get TablesWritten() As Long
    TablesWritten = this.tablesWritten
End Property

'@label:graphs-written
'@prop-title How many charts this build put on a sheet
'@return Long. The number of charts written.
'@export
Public Property Get GraphsWritten() As Long
    GraphsWritten = this.graphsWritten
End Property

'@label:has-checkings
'@prop-title Whether anything was reported during the build
'@return Boolean. True when at least one entry was filed.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title The report of this build
'@return Checking. The store, or Nothing when nothing was filed.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property


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
