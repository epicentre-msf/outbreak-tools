VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds the named-range strings shared by the analysis writers"

'@PredeclaredId
'@Folder("Analyses")
'@ModuleDescription("Builds the named-range strings shared by the analysis writers")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends Checking

Option Explicit

'@class AnalysisRanges
'@description
'One instance holds one range identifier and answers the named-range strings
'built from it. An analysis output sheet carries about sixty named ranges per
'table, and this class is where each of those names is spelled.
'
'Create takes the identifier, BindId changes it and Id answers it. Each member
'gives one name: the value and label columns, the percentage columns, the row
'and column category ranges, the total and missing rows and columns and the
'cells where they cross, the section and table bounds, the time period names,
'and the spatial input cells.
'
'WHICH NAMES LIVE HERE
'-------------------------------------------------------------------------------
'The members are the name families built by MORE THAN ONE file, counted across
'the whole source tree. The second builder is often the graph writer, the
'linelist runtime or a standard module. A name with a single builder stays
'where it is.
'
'A member whose other builder sits in a class the test harness cannot compile
'is defined here first and adopted when that class is registered.
'
'ONE IDENTIFIER PER INSTANCE
'-------------------------------------------------------------------------------
'The analysis writers work with two identifiers: the id of a table, and the id
'of the first table of its section. Several names are built from the table id
'in one class and from the section id in another. Binding one identifier per
'instance puts that choice in the constructor call, so tabNames and secNames
'read differently at the point of use.
'
'THE SPATIAL TAG IS A PARAMETER
'-------------------------------------------------------------------------------
'SpatialInput takes the spatial type from its caller. Telling a geo variable
'from a health facility one needs the dictionary, and this class holds no
'dictionary, no worksheet and no specification. It spells the string and
'TableSpecs decides the type. The class names no collaborator, so the builder
'and the generated linelist both carry it and read one definition of every
'string.
'@depends Checking
'@version 1.0 (2026-07-30)


'@section State and constants
'===============================================================================
Private Type TAnalysisRanges
    rangeId As String                            'The table or section identifier this instance answers for
End Type

Private this As TAnalysisRanges

Private Const CLASS_NAME As String = "AnalysisRanges"

' The two spatial input tags. These are the strings C2 was about: the cells are
' named with one of them, and the label formulas reference the same choice. Both
' live here so the writer and the reader cannot pick differently.
Private Const TAG_SPATIAL_GEO As String = "INPUTSPTGEO_"
Private Const TAG_SPATIAL_HF As String = "INPUTSPTHF_"

' Spatial type values, as SpatialTableScopes answers them.
Private Const SPATIAL_TYPE_HF As String = "hf"
Private Const SPATIAL_TYPE_GEO As String = "geo"


'@section Factory
'===============================================================================
'Initialisation and self-reference.

'@label:create
'@sub-title Factory returning an instance bound to one range identifier
'@details
'Binds the identifier every id-suffixed member below appends. Pass a table id
'for the names that belong to one table and a section id for the names that
'belong to the section, so the call site states which it means.
'@param rangeId String. A table id such as "BA_tab5" or a section id.
'@return AnalysisRanges. Ready to use.
'@throws InvalidArgument When rangeId is empty.
'@export
Public Function Create(ByVal rangeId As String) As AnalysisRanges
    Dim rangeNames As AnalysisRanges

    ' An empty id builds "VALUES_COL_1_", which is a legal name that every
    ' table would claim in turn. Refusing it here is the difference between a
    ' raise and a sheet whose names silently point at the wrong table.
    If LenB(rangeId) = 0 Then _
        ThrowError ProjectError.InvalidArgument, _
                   "A range identifier is required"

    Set rangeNames = New AnalysisRanges
    rangeNames.BindId rangeId

    Set Create = rangeNames
End Function

'@label:bind-id
'@sub-title Store the range identifier
'@details
'Called by Create only. Public because the factory builds a second instance and
'VBA gives a class no privileged access to another instance of itself.
'@param rangeId String. The identifier to hold.
'@export
Public Sub BindId(ByVal rangeId As String)
    this.rangeId = rangeId
End Sub

'@label:id
'@prop-title The identifier this instance answers for
'@return String. The table or section id passed to Create.
'@export
Public Property Get Id() As String
    Id = this.rangeId
End Property


'@section Column families
'===============================================================================
'@description The three per-column families, each indexed by column number.
'One loop in the table writer names all three, and the formula writer reads
'them back by the same index.

'@label:values-col
'@prop-title Name of the nth value column
'@param n Long. The one-based column number.
'@return String. For example "VALUES_COL_3_BA_tab5".
'@export
Public Property Get ValuesCol(ByVal n As Long) As String
    ValuesCol = "VALUES_COL_" & n & "_" & this.rangeId
End Property

'@label:label-col
'@prop-title Name of the nth label column
'@param n Long. The one-based column number.
'@return String. For example "LABEL_COL_3_BA_tab5".
'@export
Public Property Get LabelCol(ByVal n As Long) As String
    LabelCol = "LABEL_COL_" & n & "_" & this.rangeId
End Property

'@label:perc-col
'@prop-title Name of the nth percentage column
'@param n Long. The one-based column number.
'@return String. For example "PERC_COL_3_BA_tab5".
'@export
Public Property Get PercCol(ByVal n As Long) As String
    PercCol = "PERC_COL_" & n & "_" & this.rangeId
End Property

'@label:perc-label-col
'@prop-title Name of the percentage column header
'@details
'One header serves the percentage columns of the table, so this one carries no
'column number while the three families above do.
'@return String. For example "PERC_LABEL_COL_BA_tab5".
'@export
Public Property Get PercLabelCol() As String
    PercLabelCol = "PERC_LABEL_COL_" & this.rangeId
End Property

'@label:spatial-input
'@prop-title Name of the nth spatial input cell
'@details
'The tag depends on whether the table is grouped by an administrative level or
'by a health facility, and the caller supplies that answer because working it
'out needs a dictionary. An unrecognised type raises rather than falling back
'to the geographic tag: falling back is what named the cells one way while the
'formulas referenced the other, and a spatio-temporal table that passes
'validation always has one of the two prefixes.
'@param n Long. The one-based input row number.
'@param spatialType String. Either "hf" or "geo", as SpatialTableScopes answers.
'@return String. For example "INPUTSPTHF_2_SPT_tab1".
'@throws InvalidArgument When spatialType is neither "hf" nor "geo".
'@export
Public Property Get SpatialInput(ByVal n As Long, ByVal spatialType As String) As String
    SpatialInput = SpatialInputTag(spatialType) & n & "_" & this.rangeId
End Property

'@label:spatial-input-tag
'@prop-title The spatial input tag for a spatial type
'@details
'Exposed on its own because a caller that names a whole block of input cells
'wants the tag once rather than per cell, and because the reverse direction
'below compares against the same two strings.
'@param spatialType String. Either "hf" or "geo".
'@return String. Either "INPUTSPTHF_" or "INPUTSPTGEO_".
'@throws InvalidArgument When spatialType is neither "hf" nor "geo".
'@export
Public Property Get SpatialInputTag(ByVal spatialType As String) As String
    Select Case LCase$(spatialType)
    Case SPATIAL_TYPE_HF
        SpatialInputTag = TAG_SPATIAL_HF
    Case SPATIAL_TYPE_GEO
        SpatialInputTag = TAG_SPATIAL_GEO
    Case Else
        ThrowError ProjectError.InvalidArgument, _
                   "Unknown spatial type " & Chr$(34) & spatialType & Chr$(34) & _
                   ", expected " & Chr$(34) & SPATIAL_TYPE_HF & Chr$(34) & _
                   " or " & Chr$(34) & SPATIAL_TYPE_GEO & Chr$(34)
    End Select
End Property

'@label:id-of-spatial-input
'@sub-title The range id a spatial input cell name was built from
'@details
'SpatialInput builds "INPUTSPTGEO_<n>_<rangeId>", and this is the way back:
'the tag and the row number are cut off the front and the rest is the id.
'GeoModule reads the name off the active cell at run time, and an unnamed
'cell hands an empty string over, so the answer is empty for a name this
'class did not build: no tag at the front, no row number behind it, or
'nothing after the row number. Instance state is never read, so the member
'is callable on the predeclared instance.
'@param cellName String. A name such as "INPUTSPTGEO_2_SPT_tab1".
'@return String. The range id, or an empty string.
'@export
Public Function IdOfSpatialInput(ByVal cellName As String) As String
    Dim rest As String
    Dim sepAt As Long

    If Left$(cellName, Len(TAG_SPATIAL_GEO)) = TAG_SPATIAL_GEO Then
        rest = Mid$(cellName, Len(TAG_SPATIAL_GEO) + 1)
    ElseIf Left$(cellName, Len(TAG_SPATIAL_HF)) = TAG_SPATIAL_HF Then
        rest = Mid$(cellName, Len(TAG_SPATIAL_HF) + 1)
    Else
        Exit Function
    End If

    sepAt = InStr(1, rest, "_")
    If sepAt <= 1 Then Exit Function
    If Not IsNumeric(Left$(rest, sepAt - 1)) Then Exit Function

    IdOfSpatialInput = Mid$(rest, sepAt + 1)
End Function


'@section Category and row families
'===============================================================================
'@description Names of the category bands, the value blocks and the markers
'that locate the table on the sheet.

'@label:row-categories
'@prop-title Name of the row categories band
'@return String. For example "ROW_CATEGORIES_BA_tab5".
'@export
Public Property Get RowCategories() As String
    RowCategories = "ROW_CATEGORIES_" & this.rangeId
End Property

'@label:label-row-categories
'@prop-title Name of the header cell above the row categories
'@return String. For example "LABEL_ROW_CATEGORIES_BA_tab5".
'@export
Public Property Get LabelRowCategories() As String
    LabelRowCategories = "LABEL_ROW_CATEGORIES_" & this.rangeId
End Property

'@label:column-categories
'@prop-title Name of the column categories band
'@return String. For example "COLUMN_CATEGORIES_BA_tab5".
'@export
Public Property Get ColumnCategories() As String
    ColumnCategories = "COLUMN_CATEGORIES_" & this.rangeId
End Property

'@label:interior-values
'@prop-title Name of the block of data cells
'@return String. For example "INTERIOR_VALUES_BA_tab5".
'@export
Public Property Get InteriorValues() As String
    InteriorValues = "INTERIOR_VALUES_" & this.rangeId
End Property

'@label:outer-values
'@prop-title Name of the block of data cells including the totals
'@return String. For example "OUTER_VALUES_BA_tab5".
'@export
Public Property Get OuterValues() As String
    OuterValues = "OUTER_VALUES_" & this.rangeId
End Property

'@label:section
'@prop-title Name of the section heading cell
'@return String. For example "SECTION_BA_tab5".
'@export
Public Property Get Section() As String
    Section = "SECTION_" & this.rangeId
End Property

'@label:end-table
'@prop-title Name of the cell marking the end of the table
'@return String. For example "ENDTABLE_BA_tab5".
'@export
Public Property Get EndTable() As String
    EndTable = "ENDTABLE_" & this.rangeId
End Property

'@label:start-row
'@prop-title Name of the start-row marker
'@return String. For example "STARTROW_BA_tab5".
'@export
Public Property Get StartRow() As String
    StartRow = "STARTROW_" & this.rangeId
End Property

'@label:start-col
'@prop-title Name of the start-column marker
'@return String. For example "STARTCOL_BA_tab5".
'@export
Public Property Get StartCol() As String
    StartCol = "STARTCOL_" & this.rangeId
End Property


'@section Total and missing intersections
'===============================================================================
'@description The total and missing bands, and the four cells where they cross.
'The writer names them and the formula writer fills them, which is why every
'one of these has two builders today.

'@label:total-row
'@prop-title Name of the total row band
'@return String. For example "TOTAL_ROW_BA_tab5".
'@export
Public Property Get TotalRow() As String
    TotalRow = "TOTAL_ROW_" & this.rangeId
End Property

'@label:total-col
'@prop-title Name of the total column band
'@return String. For example "TOTAL_COL_BA_tab5".
'@export
Public Property Get TotalCol() As String
    TotalCol = "TOTAL_COL_" & this.rangeId
End Property

'@label:missing-row
'@prop-title Name of the missing row band
'@return String. For example "MISSING_ROW_BA_tab5".
'@export
Public Property Get MissingRow() As String
    MissingRow = "MISSING_ROW_" & this.rangeId
End Property

'@label:missing-col
'@prop-title Name of the missing column band
'@return String. For example "MISSING_COL_BA_tab5".
'@export
Public Property Get MissingCol() As String
    MissingCol = "MISSING_COL_" & this.rangeId
End Property

'@label:total-row-values
'@prop-title Name of the values inside the total row
'@return String. For example "TOTAL_ROW_VALUES_BA_tab5".
'@export
Public Property Get TotalRowValues() As String
    TotalRowValues = "TOTAL_ROW_VALUES_" & this.rangeId
End Property

'@label:total-col-values
'@prop-title Name of the values inside the total column
'@return String. For example "TOTAL_COL_VALUES_BA_tab5".
'@export
Public Property Get TotalColValues() As String
    TotalColValues = "TOTAL_COL_VALUES_" & this.rangeId
End Property

'@label:missing-row-values
'@prop-title Name of the values inside the missing row
'@return String. For example "MISSING_ROW_VALUES_BA_tab5".
'@export
Public Property Get MissingRowValues() As String
    MissingRowValues = "MISSING_ROW_VALUES_" & this.rangeId
End Property

'@label:missing-col-values
'@prop-title Name of the values inside the missing column
'@return String. For example "MISSING_COL_VALUES_BA_tab5".
'@export
Public Property Get MissingColValues() As String
    MissingColValues = "MISSING_COL_VALUES_" & this.rangeId
End Property

'@label:total-total
'@prop-title Name of the total row and total column intersection
'@return String. For example "TOTAL_TOTAL_BA_tab5".
'@export
Public Property Get TotalTotal() As String
    TotalTotal = "TOTAL_TOTAL_" & this.rangeId
End Property

'@label:total-missing
'@prop-title Name of the total row and missing column intersection
'@return String. For example "TOTAL_MISSING_BA_tab5".
'@export
Public Property Get TotalMissing() As String
    TotalMissing = "TOTAL_MISSING_" & this.rangeId
End Property

'@label:missing-total
'@prop-title Name of the missing row and total column intersection
'@return String. For example "MISSING_TOTAL_BA_tab5".
'@export
Public Property Get MissingTotal() As String
    MissingTotal = "MISSING_TOTAL_" & this.rangeId
End Property

'@label:missing-missing
'@prop-title Name of the missing row and missing column intersection
'@return String. For example "MISSING_MISSING_BA_tab5".
'@export
Public Property Get MissingMissing() As String
    MissingMissing = "MISSING_MISSING_" & this.rangeId
End Property


'@section Temporal families
'===============================================================================
'@description The date inputs, the period markers and the time unit control.

'@label:time-unit
'@prop-title Name of the time unit control cell
'@return String. For example "TIME_UNIT_TS_tab2".
'@export
Public Property Get TimeUnit() As String
    TimeUnit = "TIME_UNIT_" & this.rangeId
End Property

'@label:start-date
'@prop-title Name of the start date cell
'@return String. For example "START_DATE_TS_tab2".
'@export
Public Property Get StartDate() As String
    StartDate = "START_DATE_" & this.rangeId
End Property

'@label:end-date
'@prop-title Name of the end date cell
'@return String. For example "END_DATE_TS_tab2".
'@export
Public Property Get EndDate() As String
    EndDate = "END_DATE_" & this.rangeId
End Property

'@label:start-time-period
'@prop-title Name of the period start marker
'@return String. For example "START_TIME_PERIOD_TS_tab2".
'@export
Public Property Get StartTimePeriod() As String
    StartTimePeriod = "START_TIME_PERIOD_" & this.rangeId
End Property

'@label:end-time-period
'@prop-title Name of the period end marker
'@return String. For example "END_TIME_PERIOD_TS_tab2".
'@export
Public Property Get EndTimePeriod() As String
    EndTimePeriod = "END_TIME_PERIOD_" & this.rangeId
End Property

'@label:first-value-start-time
'@prop-title Name of the first value of the period
'@details
'One of the names built from the table id in the table writer and from the
'section id in the formula writer. Bind the instance to whichever the call
'means and the difference stops being invisible.
'@return String. For example "FIRST_VALUE_START_TIME_TS_tab2".
'@export
Public Property Get FirstValueStartTime() As String
    FirstValueStartTime = "FIRST_VALUE_START_TIME_" & this.rangeId
End Property

'@label:info-start-date
'@prop-title Name of the start date information cell
'@return String. For example "INFO_START_DATE_TS_tab2".
'@export
Public Property Get InfoStartDate() As String
    InfoStartDate = "INFO_START_DATE_" & this.rangeId
End Property

'@label:info-end-date
'@prop-title Name of the end date information cell
'@return String. For example "INFO_END_DATE_TS_tab2".
'@export
Public Property Get InfoEndDate() As String
    InfoEndDate = "INFO_END_DATE_" & this.rangeId
End Property

'@label:info-ana-period
'@prop-title Name of the analysis period information cell
'@return String. For example "INFO_ANA_PERIOD_TS_tab2".
'@export
Public Property Get InfoAnaPeriod() As String
    InfoAnaPeriod = "INFO_ANA_PERIOD_" & this.rangeId
End Property

'@label:validation-min-date
'@prop-title Name of the minimum date validation cell
'@return String. For example "VALIDATION_MIN_DATE_TS_tab2".
'@export
Public Property Get ValidationMinDate() As String
    ValidationMinDate = "VALIDATION_MIN_DATE_" & this.rangeId
End Property

'@label:validation-max-date
'@prop-title Name of the maximum date validation cell
'@return String. For example "VALIDATION_MAX_DATE_TS_tab2".
'@export
Public Property Get ValidationMaxDate() As String
    ValidationMaxDate = "VALIDATION_MAX_DATE_" & this.rangeId
End Property

'@label:min-min-date
'@prop-title Name of the earliest date across the section
'@return String. For example "MIN_MIN_DATE_TS_tab2".
'@export
Public Property Get MinMinDate() As String
    MinMinDate = "MIN_MIN_DATE_" & this.rangeId
End Property

'@label:max-max-date
'@prop-title Name of the latest date across the section
'@return String. For example "MAX_MAX_DATE_TS_tab2".
'@export
Public Property Get MaxMaxDate() As String
    MaxMaxDate = "MAX_MAX_DATE_" & this.rangeId
End Property

'@label:user-start-date
'@prop-title Name of the cell the reader types the first date into
'@details
'The table writer names the cell and six formulas of the formula writer read it,
'so it belongs here beside the two dates derived from it.
'@return String. For example "USER_START_DATE_TS_tab2".
'@export
Public Property Get UserStartDate() As String
    UserStartDate = "USER_START_DATE_" & this.rangeId
End Property

'@label:user-end-date
'@prop-title Name of the cell the reader types the last date into
'@return String. For example "USER_END_DATE_TS_tab2".
'@export
Public Property Get UserEndDate() As String
    UserEndDate = "USER_END_DATE_" & this.rangeId
End Property


'@section Spatial control families
'===============================================================================
'@description The administrative level dropdown and the population divisor.

'@label:adm-dropdown
'@prop-title Name of the administrative level dropdown cell
'@return String. For example "ADM_DROPDOWN_SPT_tab1".
'@export
Public Property Get AdmDropdown() As String
    AdmDropdown = "ADM_DROPDOWN_" & this.rangeId
End Property

'@label:devide-pop
'@prop-title Name of the population divisor control cell
'@return String. For example "DEVIDEPOP_SPT_tab1".
'@export
Public Property Get DevidePop() As String
    DevidePop = "DEVIDEPOP_" & this.rangeId
End Property

'@label:pop-fact
'@prop-title Name of the population factor cell
'@return String. For example "POPFACT_SPT_tab1".
'@export
Public Property Get PopFact() As String
    PopFact = "POPFACT_" & this.rangeId
End Property

'@label:pop-fact-label
'@prop-title Name of the population factor header cell
'@return String. For example "POPFACTLABEL_SPT_tab1".
'@export
Public Property Get PopFactLabel() As String
    PopFactLabel = "POPFACTLABEL_" & this.rangeId
End Property

'@label:pop-prev-fact
'@prop-title Name of the previous population factor cell
'@return String. For example "POPPREVFACT_SPT_tab1".
'@export
Public Property Get PopPrevFact() As String
    PopPrevFact = "POPPREVFACT_" & this.rangeId
End Property

'@label:previous-adm
'@prop-title Name of the previous administrative level cell
'@return String. For example "PREVIOUS_ADM_SPT_tab1".
'@export
Public Property Get PreviousAdm() As String
    PreviousAdm = "PREVIOUS_ADM_" & this.rangeId
End Property


'@section Workbook-wide lists
'===============================================================================
'@description Names carrying no identifier, because one copy of each serves
'every table on the sheet. Read them off the predeclared instance without
'calling Create.

'@label:time-unit-list
'@prop-title Name of the time unit choice list
'@details
'ONE LIST SERVES THE WHOLE WORKBOOK
'-------------------------------------------------------------------------------
'The time units are a workbook-level dropdown on the list sheet, added by
'Linelist.Prepare through the standard DropdownLists manager. Every temporal
'table on every sheet binds its dropdown to this one name, and GetAgg in
'CustomLinelistFunctions reads the same name to map the chosen label back to a
'time unit.
'
'The sheet used to carry its own copy: a time series sheet built
'TIME_UNIT_LIST and a spatio-temporal sheet built SPTIME_UNIT_LIST, because a
'range name resolves to one cell block and the two sheets each wanted their
'own. The reader in CustomLinelistFunctions asked for the plain name whatever
'sheet it ran on, so a workbook whose only temporal analyses were
'spatio-temporal had no such name and every period cell answered #VALUE!.
'
'The spelling is what DropdownLists derives from the list name TIME_UNIT_NAME
'below and the "dropdown_" prefix of the standard manager. TestAnalysisRanges
'pins the two together.
'@return String. "dropdown_time_unit".
'@export
Public Property Get TimeUnitList() As String
    TimeUnitList = "dropdown_time_unit"
End Property

'@label:time-unit-list-name
'@prop-title The dropdown list name the time units are registered under
'@details
'What Linelist.Prepare passes to DropdownLists.Add. The workbook-level name
'above is what every reader uses; this is the registry key that produces it.
'@return String. "time_unit".
'@export
Public Property Get TimeUnitListName() As String
    TimeUnitListName = "time_unit"
End Property

'@label:adm-unit-list
'@prop-title Name of the administrative unit choice list
'@return String. "ADM_UNIT_LIST".
'@export
Public Property Get AdmUnitList() As String
    AdmUnitList = "ADM_UNIT_LIST"
End Property

'@label:population-factor-list
'@prop-title Name of the population factor choice list
'@return String. "POPULATION_FACTOR_LIST".
'@export
Public Property Get PopulationFactorList() As String
    PopulationFactorList = "POPULATION_FACTOR_LIST"
End Property


'@section Error management
'===============================================================================

'@label:throw-error
'@sub-title Raise a project error naming this class
'@param errNumber Long. A ProjectError member.
'@param message String. The message to carry.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
