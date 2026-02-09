VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Analysis table specifications"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Analysis table specifications")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ITableSpecs

Private Type TTableSpecs
    hRng As Range                                'Header range of the setup listobject
    tRng As Range                                'Row range for this table spec
    lData As ILinelistSpecs                      'Linelist specifications
End Type

Private this As TTableSpecs
Private Const CLASS_NAME As String = "TableSpecs"

'@section Factory
'===============================================================================

' @description Factory method that creates a new TableSpecs instance representing one row
'              from an analysis setup ListObject. Each row in the setup table defines a
'              single analysis table (e.g., a univariate frequency table or a time series).
'              The TableSpecs object wraps that row and exposes properties to query its
'              type, validity, categories, navigation to neighboring rows, and display flags
'              (total, missing, percentage, graph). Validation is performed immediately via
'              CheckRequirements before the instance is populated.
' @param headerRng The header row Range of the analysis setup ListObject. Column names in
'                  this range are used by Value() and ColumnIndex() to locate fields such as
'                  "row", "column", "total", "section", etc.
' @param tableSpecsRng A single data-row Range from the same ListObject. Must be below
'                      headerRng and have the same number of columns. This is the row whose
'                      cell values define the table specification.
' @param lData The ILinelistSpecs object that provides access to the linelist dictionary,
'              variable metadata, and category lists needed for validation and category
'              retrieval.
' @return A new ITableSpecs interface pointer to the fully initialized TableSpecs instance.
' @remark The method uses the PredeclaredId factory pattern: it creates a fresh instance via
'         "With New TableSpecs", sets internal properties, then returns .Self. Callers should
'         always receive the result as ITableSpecs, not as the concrete TableSpecs class.
Public Function Create(ByVal headerRng As Range, _
                       ByVal tableSpecsRng As Range, _
                       ByVal lData As ILinelistSpecs) As ITableSpecs

    CheckRequirements headerRng, tableSpecsRng, lData

    With New TableSpecs
        Set .HeaderRange = headerRng
        Set .TableRange = tableSpecsRng
        Set .LinelistSpecifications = lData
        Set Create = .Self
    End With
End Function

'@section Internal properties (used by factory)
'===============================================================================

Public Property Get Self() As ITableSpecs
    Set Self = Me
End Property

Public Property Get HeaderRange() As Range
    Set HeaderRange = this.hRng
End Property

Public Property Set HeaderRange(ByVal hRng As Range)
    Set this.hRng = hRng
End Property

Public Property Get TableRange() As Range
    Set TableRange = this.tRng
End Property

Public Property Set TableRange(ByVal tRng As Range)
    Set this.tRng = tRng
End Property

Public Property Get LinelistSpecifications() As ILinelistSpecs
    Set LinelistSpecifications = this.lData
End Property

Public Property Set LinelistSpecifications(ByVal lData As ILinelistSpecs)
    Set this.lData = lData
End Property

'@section Core properties
'===============================================================================

' @description Determines the AnalysisTablesType enum value for this table specification by
'              reading the type label that sits two rows above the header range (cell offset
'              -1 from the header row). In the setup workbook layout, each analysis block has
'              a merged title cell above its header; this property reads that cell, trims and
'              lowercases it, then maps it to the corresponding enum constant:
'              TypeGlobalSummary (0), TypeUnivariate (1), TypeBivariate (2), TypeTimeSeries (3),
'              TypeSpatial (4), or TypeSpatioTemporal (5). If the label does not match any
'              known type string, a runtime error is raised.
' @return An AnalysisTablesType enum value indicating whether this table is a global summary,
'         univariate analysis, bivariate analysis, time series, spatial, or spatio-temporal.
' @remark The cell offset hRng.Cells(-1, 1) means row index -1 relative to the header, which
'         in Excel's Range.Cells addressing is 2 rows above the header row (since Cells(1,1)
'         is the header itself and Cells(0,1) is one row above). If the setup layout changes,
'         this offset must be updated accordingly.
Public Property Get TableType() As AnalysisTablesType
    Dim hRng As Range
    Dim anaType As String

    Set hRng = HeaderRange()
    ' The type label is two rows above the header (row offset -1 from header row)
    anaType = LCase(Application.WorksheetFunction.Trim(hRng.Cells(-1, 1).Value))

    Select Case anaType
    Case "global summary"
        TableType = TypeGlobalSummary
    Case "univariate analysis"
        TableType = TypeUnivariate
    Case "bivariate analysis"
        TableType = TypeBivariate
    Case "time series analysis"
        TableType = TypeTimeSeries
    Case "spatial analysis"
        TableType = TypeSpatial
    Case "spatio-temporal analysis"
        TableType = TypeSpatioTemporal
    Case Else
        ThrowError ErrorUnexpectedState, _
                   "Could not find the table type. Actual type is: " & anaType
    End Select
End Property

' @description Generates a unique string identifier for this table specification by combining
'              a short type-based prefix with the row offset from the header. The prefix is
'              derived from the TableType: "GS" for GlobalSummary, "UA" for Univariate, "BA"
'              for Bivariate, "TS" for TimeSeries, "SA" for Spatial, and "SPT" for
'              SpatioTemporal. The row offset (lineNum) is computed as TableRange.Row minus
'              HeaderRange.Row, producing a stable integer that uniquely locates this row
'              within its analysis block. The resulting format is "prefix_tabN", for example
'              "UA_tab3" means the univariate table at row offset 3 from the header.
' @return A string identifier such as "GS_tab1", "BA_tab5", or "SPT_tab12". This id is used
'         throughout the analysis engine to reference specific table specs and to build
'         section identifiers via TableSectionId.
' @remark The id is positional, not semantic. If rows are inserted or deleted in the setup
'         ListObject, the ids for downstream rows will change. The id is deterministic for a
'         given row position.
Public Property Get TableId() As String
    Dim prefix As String
    Dim lineNum As Long

    lineNum = TableRange.Row - HeaderRange.Row

    Select Case TableType
    Case TypeGlobalSummary
        prefix = "GS"
    Case TypeUnivariate
        prefix = "UA"
    Case TypeBivariate
        prefix = "BA"
    Case TypeTimeSeries
        prefix = "TS"
    Case TypeSpatial
        prefix = "SA"
    Case TypeSpatioTemporal
        prefix = "SPT"
    End Select

    TableId = prefix & "_tab" & lineNum
End Property

' @description Returns the TableId of the first table in the current section group. This
'              property walks backward through the Previous chain to find the originating
'              table of the section. If this table is itself the start of a new section
'              (IsNewSection = True) or is a GlobalSummary table, it returns its own TableId
'              immediately without recursion. Otherwise it delegates to Previous.TableSectionId,
'              which continues walking backward until a section boundary is reached. This
'              allows all tables within the same section to share a common section identifier,
'              which is used to group related tables together in output formatting.
' @return The TableId string of the first table in this section, e.g. "UA_tab2". All tables
'         in the same section will return the same value from this property.
' @remark This is a recursive property. For very long sections with many rows, it will
'         traverse one row at a time via Previous. GlobalSummary tables short-circuit
'         immediately because they are standalone (each is its own section).
Public Property Get TableSectionId() As String
    If IsNewSection() Or (TableType = TypeGlobalSummary) Then
        TableSectionId = TableId()
    Else
        TableSectionId = Previous.TableSectionId
    End If
End Property

' @description Convenience shortcut that returns the ILLdictionary from the linelist
'              specifications object. The dictionary contains all variable definitions,
'              control types, category lists, and sheet metadata. Multiple properties in
'              this class (ValidTable, SpatialTableTypes, CategoriesData) need frequent
'              access to the dictionary, so this property avoids repeating the chained
'              call LinelistSpecifications().Dictionary() throughout the class.
' @return The ILLdictionary instance held by the linelist specifications.
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = LinelistSpecifications().Dictionary()
End Property

' @description Determines the spatial sub-type of a Spatial or SpatioTemporal table by
'              checking whether the relevant variable has a health-facility or administrative
'              geography prefix in the dictionary. For Spatial tables, the variable comes from
'              the "row" field; for SpatioTemporal tables, the spatial variable is in the
'              "column" field instead (because the row variable is a date in spatio-temporal
'              analyses). The method prepends "hf_" or "adm1_" to the variable name and checks
'              whether that prefixed variable exists in the dictionary. If the "hf_" prefixed
'              variable exists, the sub-type is "hf" (health facility); if "adm1_" exists, it
'              is "geo" (administrative geography). If neither prefix matches, returns
'              vbNullString.
' @return "geo" if the variable has an adm1_ prefix in the dictionary, "hf" if it has an
'         hf_ prefix, or vbNullString if neither prefix is found.
' @remark This property is meaningful only for Spatial and SpatioTemporal table types. Calling
'         it on other table types will still execute but will likely return vbNullString since
'         those variables typically lack geo/hf prefixes. The Switch function evaluates
'         conditions in order, so "hf" takes priority over "geo" if both somehow exist.
Public Property Get SpatialTableTypes() As String
    Dim rowVar As String
    Dim dict As ILLdictionary
    Dim spatType As String

    Set dict = Dictionary()

    ' For spatio-temporal, the spatial variable is in the "column" field
    If TableType = TypeSpatioTemporal Then
        rowVar = Value("column")
    Else
        rowVar = Value("row")
    End If

    spatType = Switch(dict.VariableExists("hf_" & rowVar), "hf", _
                      dict.VariableExists("adm1_" & rowVar), "geo", _
                      True, vbNullString)
    SpatialTableTypes = spatType
End Property

'@section Validation properties
'===============================================================================

' @description Validates whether this table specification row contains all required fields
'              and references valid variables for its analysis type. Each table type has
'              different validation rules:
'              - GlobalSummary: requires non-empty "label" and "function" fields.
'              - Univariate: requires the "row" variable to exist in the dictionary and its
'                control type to be "choice_manual" or "choice_formula" (categorical variable).
'              - Bivariate: requires both "row" and "column" variables to exist and both must
'                be choice variables (choice_manual or choice_formula).
'              - TimeSeries: requires the "row" variable to be of type "date". The "column"
'                variable is optional, but if present must be a choice variable.
'              - Spatial: requires the "row" variable to have an "adm1_" or "hf_" prefixed
'                counterpart in the dictionary (indicating a geographic variable). The "column"
'                variable is optional but must be a choice variable if present.
'              - SpatioTemporal: requires the "column" variable to have an "adm1_" or "hf_"
'                prefix in the dictionary and the "row" variable to be of type "date".
'              Invalid table specs are skipped during navigation (Previous, NextSpecs).
' @return True if the table specification meets all requirements for its type, False otherwise.
' @remark This property creates a temporary LLVariables instance each time it is called.
'         For performance-sensitive loops, consider caching the result. The validation does
'         not check for the existence of categories or other downstream concerns, only
'         structural correctness of the specification row.
Public Property Get ValidTable() As Boolean
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim rowVar As String
    Dim colVar As String
    Dim colControl As String
    Dim rowControl As String
    Dim rowType As String
    Dim isValid As Boolean

    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)

    rowVar = Value("row")
    rowControl = vars.Value(colName:="control", varName:=rowVar)
    colVar = Value("column")
    colControl = vars.Value(colName:="control", varName:=colVar)

    Select Case TableType

    Case TypeGlobalSummary
        isValid = (Value("label") <> vbNullString) And _
                  (Value("function") <> vbNullString)

    Case TypeUnivariate
        isValid = vars.Contains(rowVar) And _
                  (rowControl = "choice_manual" Or rowControl = "choice_formula")

    Case TypeBivariate
        isValid = vars.Contains(rowVar) And vars.Contains(colVar) And _
                  (colControl = "choice_manual" Or colControl = "choice_formula") And _
                  (rowControl = "choice_manual" Or rowControl = "choice_formula")

    Case TypeTimeSeries
        rowType = vars.Value(colName:="variable type", varName:=rowVar)
        isValid = (colVar = vbNullString Or _
                   colControl = "choice_manual" Or _
                   colControl = "choice_formula") And _
                  (rowType = "date")

    Case TypeSpatial
        isValid = (dict.VariableExists("adm1_" & rowVar) Or _
                   dict.VariableExists("hf_" & rowVar)) And _
                  (colVar = vbNullString Or _
                   colControl = "choice_manual" Or _
                   colControl = "choice_formula")

    Case TypeSpatioTemporal
        rowType = vars.Value(colName:="variable type", varName:=rowVar)
        isValid = (dict.VariableExists("adm1_" & colVar) Or _
                   dict.VariableExists("hf_" & colVar)) And _
                  (rowType = "date")

    End Select

    ValidTable = isValid
End Property

' @description Determines whether this table specification row begins a new section within
'              the analysis setup block. Sections are used to group related tables (e.g.,
'              multiple univariate tables for the same thematic area). The method checks the
'              "section" column: if the current row's section value differs from the previous
'              row's section value, or if this is the first data row (where the previous row
'              is the header), the table is considered a new section. GlobalSummary tables are
'              never considered new sections regardless of their section value, because each
'              GlobalSummary table is standalone and does not participate in section grouping.
'              If no "section" column exists in the header, every non-GlobalSummary table is
'              treated as a new section.
' @return True if this table starts a new section, False if it continues the same section as
'         the row above it.
' @remark The comparison uses the cell at row offset 0 (prevCell = tRng.Cells(0, sectIndex)),
'         which in Excel Range addressing corresponds to the row immediately above the current
'         row. The guard "prevCell.Row <> hRng.Row" ensures that the first data row (whose
'         previous row is the header) is always treated as a new section. Section values are
'         compared with exact string equality (case-sensitive).
Public Property Get IsNewSection() As Boolean
    Dim sectIndex As Long
    Dim actualSection As String
    Dim previousSection As String
    Dim isSameSection As Boolean
    Dim tRng As Range
    Dim hRng As Range
    Dim prevCell As Range

    Set tRng = TableRange()
    Set hRng = HeaderRange()
    isSameSection = False

    If ColumnExists("section") Then
        sectIndex = ColumnIndex("section")
        actualSection = Value("section")
        Set prevCell = tRng.Cells(0, sectIndex)  '0 offset = previous row
        previousSection = prevCell.Value
        isSameSection = (actualSection = previousSection) And (prevCell.Row <> hRng.Row)
    End If

    IsNewSection = (Not isSameSection) And (TableType <> TypeGlobalSummary)
End Property

'@section Flag properties (HasTotal, HasMissing, HasPercentage, HasGraph, TotalRequested)
'===============================================================================

' @description Determines whether a total row/column is needed for this table. The logic
'              varies by table type:
'              - GlobalSummary and SpatioTemporal: always False (totals are not applicable).
'              - Univariate: always True (frequency tables always include a total row).
'              - Bivariate: always True (cross-tabulations always include totals).
'              - TimeSeries: True when the user explicitly requested total="yes" OR when
'                percentage is set to "row" or "column", because percentage formulas need
'                the total column as a denominator for their calculations. Also requires a
'                column variable to be present (column <> vbNullString). This is the key
'                distinction from TotalRequested (Bug #1 context): HasTotal can be True even
'                when the user did not ask for totals, solely because percentage needs them.
'              - Spatial: True only when a column variable is specified.
' @return True if a total row/column should be created in the output table, False otherwise.
' @remark This property drives whether the total row/column is physically created in the
'         CrossTable output. It works in tandem with TotalRequested: when HasTotal is True
'         but TotalRequested is False, the total column is created (for percentage computation)
'         but hidden from the user's view. See TotalRequested for the companion logic.
Public Property Get HasTotal() As Boolean
    Select Case TableType

    Case TypeGlobalSummary, TypeSpatioTemporal
        HasTotal = False

    Case TypeUnivariate
        HasTotal = True

    Case TypeBivariate
        HasTotal = True

    Case TypeTimeSeries
        HasTotal = ((Value("total") = "yes") Or _
                    (Value("percentage") = "row") Or _
                    (Value("percentage") = "column")) And _
                   (Value("column") <> vbNullString)

    Case TypeSpatial
        HasTotal = (Value("column") <> vbNullString)

    End Select
End Property

' @description Returns True only when the user explicitly wrote "yes" in the "total" column
'              of the setup row. This is the Bug #1 fix companion to HasTotal. The distinction
'              matters for TimeSeries tables: HasTotal can be True because percentage needs a
'              total column for its formulas, but if the user never asked for totals, the total
'              column should be hidden. CrossTable.Format uses this property to decide
'              visibility: if HasTotal = True AND TotalRequested = False, the total column is
'              created but its column is hidden so the user does not see it. If TotalRequested
'              is True, the total column is both created and visible.
' @return True if the "total" field in this row equals "yes", False for any other value or
'         if the "total" column is empty or missing.
' @remark For Univariate and Bivariate tables, TotalRequested is not consulted because
'         HasTotal is unconditionally True and totals are always shown. This property is
'         primarily relevant for TimeSeries and Spatial types.
Public Property Get TotalRequested() As Boolean
    TotalRequested = (Value("total") = "yes")
End Property

' @description Determines whether a "missing" row or column should be included in the output
'              table. The missing row/column shows counts of records where the variable value
'              is blank or undefined. The logic varies by table type:
'              - GlobalSummary and SpatioTemporal: always False (missing not applicable).
'              - Univariate: True when missing="yes" in the setup row.
'              - Bivariate: True when missing="row", "column", or "all", allowing the user to
'                choose which dimension(s) include a missing category.
'              - TimeSeries: True when missing="yes" AND a column variable is present. Without
'                a column variable, there is no categorical axis to show missing values on.
'              - Spatial: True when missing="yes" AND a column variable is present, for the
'                same reason as TimeSeries.
' @return True if missing data should be shown in the output table, False otherwise.
' @remark For TimeSeries and Spatial types, the requirement for a column variable means that
'         simple (single-variable) time series or spatial tables cannot display missing data.
'         Only cross-tabulated versions (with a column variable) support it.
Public Property Get HasMissing() As Boolean
    Select Case TableType

    Case TypeGlobalSummary, TypeSpatioTemporal
        HasMissing = False

    Case TypeUnivariate
        HasMissing = (Value("missing") = "yes")

    Case TypeBivariate
        HasMissing = (Value("missing") = "row") Or _
                     (Value("missing") = "column") Or _
                     (Value("missing") = "all")

    Case TypeTimeSeries
        HasMissing = (Value("missing") = "yes") And _
                     (Value("column") <> vbNullString)

    Case TypeSpatial
        HasMissing = (Value("missing") = "yes") And _
                     (Value("column") <> vbNullString)

    End Select
End Property

' @description Determines whether percentage values should be computed and displayed for this
'              table. Percentage calculations require a total to serve as the denominator, so
'              for types where totals are conditional, HasPercentage depends on HasTotal. The
'              logic by table type:
'              - GlobalSummary and SpatioTemporal: always False (percentage not applicable).
'              - Univariate: True when percentage="yes" in the setup row.
'              - Bivariate: True when percentage="row", "column", or "total", allowing the
'                user to choose the percentage direction (row-wise, column-wise, or of grand
'                total).
'              - TimeSeries: True when percentage="row" or "column" AND HasTotal is True.
'                HasTotal must be True because the percentage formula references the total
'                column values. Note that setting percentage here will cause HasTotal to also
'                return True (they are mutually reinforcing for TimeSeries).
'              - Spatial: True when percentage="yes" AND HasTotal is True.
' @return True if percentage values should be included in the output table, False otherwise.
' @remark The interplay between HasPercentage and HasTotal for TimeSeries is intentional:
'         requesting percentage implicitly forces HasTotal=True (since percentage needs the
'         total), and HasPercentage checks HasTotal as a guard. This circular dependency works
'         because HasTotal evaluates the percentage field directly rather than calling
'         HasPercentage.
Public Property Get HasPercentage() As Boolean
    Select Case TableType

    Case TypeGlobalSummary, TypeSpatioTemporal
        HasPercentage = False

    Case TypeUnivariate
        HasPercentage = (Value("percentage") = "yes")

    Case TypeBivariate
        HasPercentage = (Value("percentage") = "row") Or _
                        (Value("percentage") = "column") Or _
                        (Value("percentage") = "total")

    Case TypeTimeSeries
        HasPercentage = ((Value("percentage") = "row") Or _
                         (Value("percentage") = "column")) And HasTotal()

    Case TypeSpatial
        HasPercentage = (Value("percentage") = "yes") And HasTotal()

    End Select
End Property

' @description Determines whether a chart/graph should be generated alongside the output
'              table. The "graph" column in the setup row controls this behavior. The accepted
'              values and logic vary by table type:
'              - GlobalSummary: always False (no graphs for summary statistics).
'              - Univariate: True when graph="yes". Produces a bar chart of frequencies.
'              - Bivariate: True when graph="percentage", "values", or "both". These three
'                options let the user choose whether the chart shows raw counts, percentage
'                values, or both side by side.
'              - TimeSeries: True when graph="yes". Produces a line/bar chart over time.
'              - Spatial: True when graph="yes". Produces a map-style or bar chart of spatial
'                data.
'              - SpatioTemporal: True when graph="yes". Produces a combined spatial-temporal
'                visualization.
' @return True if a graph should be created for this table, False otherwise.
' @remark Bivariate is the only type that supports multiple graph modes ("percentage",
'         "values", "both"). All other types use a simple "yes"/"no" toggle. The actual graph
'         rendering is handled downstream by the chart builder; this property only signals
'         whether one should be created.
Public Property Get HasGraph() As Boolean
    Select Case TableType

    Case TypeGlobalSummary
        HasGraph = False

    Case TypeUnivariate
        HasGraph = (Value("graph") = "yes")

    Case TypeBivariate
        HasGraph = (Value("graph") = "percentage" Or _
                    Value("graph") = "values" Or _
                    Value("graph") = "both")

    Case TypeTimeSeries
        HasGraph = (Value("graph") = "yes")

    Case TypeSpatial
        HasGraph = (Value("graph") = "yes")

    Case TypeSpatioTemporal
        HasGraph = (Value("graph") = "yes")

    End Select
End Property

'@section Navigation properties
'===============================================================================

' @description Returns the nearest valid table specification that precedes this one within
'              the same section. Starting from the row immediately above the current table
'              range, it creates a TableSpecs for that row and checks its validity. If the
'              row is invalid, it continues walking backward (one row at a time) until it
'              finds a valid table or encounters a section boundary (IsNewSection on the
'              candidate). The walk stops at the header row to prevent reading outside the
'              data area. This property is used by TableSectionId to recursively find the
'              first table in a section, and by output builders that need to chain tables
'              together within a section.
' @return An ITableSpecs for the nearest valid preceding table in the same section.
' @remark Calling this property on a table where IsNewSection is True will raise an error,
'         because there is no "previous table in the same section" for a section-starting
'         table. Callers must check IsNewSection before accessing Previous. If all preceding
'         rows in the section are invalid, the returned ITableSpecs will be the last candidate
'         checked (which may itself be invalid or a new-section boundary). The walk creates a
'         new TableSpecs instance for each row examined, so performance degrades with many
'         consecutive invalid rows.
Public Property Get Previous() As ITableSpecs
    Dim hRng As Range
    Dim tRng As Range
    Dim prevRng As Range
    Dim prevSpec As ITableSpecs

    If IsNewSection() Then
        ThrowError InvalidArgument, _
                   "Table is on a new section, cannot return previous table"
    End If

    Set hRng = HeaderRange()
    Set tRng = TableRange()
    Set prevRng = tRng.Offset(-1)

    ' Create spec for the row just above
    Set prevSpec = Create(hRng, prevRng, LinelistSpecifications())
    Set prevRng = prevRng.Offset(-1)

    ' Walk backward until we find a valid table or reach section boundary
    Do While (Not prevSpec.ValidTable()) And (Not (prevRng.Row <= hRng.Row))
        Set prevSpec = Create(hRng, prevRng, LinelistSpecifications())
        If prevSpec.IsNewSection() Then Exit Do
        Set prevRng = prevRng.Offset(-1)
    Loop

    Set Previous = prevSpec
End Property

' @description Returns the next valid table specification after this row by walking forward
'              through the setup ListObject. Starting from the row immediately below the
'              current table range, it creates a TableSpecs for that row. If the row is
'              invalid, it continues advancing one row at a time until it finds a valid table
'              or passes the anchor range (the last data row of the ListObject). Unlike
'              Previous, NextSpecs does NOT stop at section boundaries -- it will cross into
'              new sections to find the next valid table. If no valid table is found within
'              the data area, the property returns Nothing.
' @param anchorRng A Range representing the last row of the setup ListObject. The forward
'                  walk will not advance past this row. Typically this is the last DataBodyRange
'                  row of the ListObject.
' @return An ITableSpecs for the next valid table after this row, or Nothing if no valid
'         table exists between this row and the anchor range.
' @remark This property is the primary iteration mechanism for processing all tables in an
'         analysis block. Callers typically start with the first table and call NextSpecs
'         repeatedly until Nothing is returned. Each call creates new TableSpecs instances
'         for each row examined, similar to Previous.
Public Property Get NextSpecs(ByVal anchorRng As Range) As ITableSpecs
    Dim tRng As Range
    Dim hRng As Range
    Dim nextSpec As ITableSpecs

    Set tRng = TableRange()
    Set hRng = HeaderRange()
    Set tRng = tRng.Offset(1)

    If tRng.Row > anchorRng.Row Then Exit Property

    Set nextSpec = Create(hRng, tRng, LinelistSpecifications())
    Set tRng = tRng.Offset(1)

    Do While (Not nextSpec.ValidTable()) And (tRng.Row <= anchorRng.Row)
        Set nextSpec = Create(hRng, tRng, LinelistSpecifications())
        Set tRng = tRng.Offset(1)
    Loop

    If Not nextSpec.ValidTable() Then Exit Property

    Set NextSpecs = nextSpec
End Property

'@section Data access
'===============================================================================

' @description Retrieves the cell value from this table specification row for a given column
'              name. The column is located by searching the header range for a cell whose
'              value contains the requested name (case-insensitive, partial match via xlPart).
'              Once the column index is found, the corresponding cell in the table range row
'              is read. This is the fundamental data-access method used by nearly every other
'              property in this class (HasTotal reads Value("total"), ValidTable reads
'              Value("row"), Value("column"), etc.). If the column name does not exist in the
'              header, the method returns vbNullString rather than raising an error, making it
'              safe to call on optional columns.
' @param colName The column name to search for in the header range. The search uses xlPart
'                matching, so "row" will match a header cell containing "row". Case-insensitive.
' @return The string value of the cell at the intersection of this table's row and the found
'         column. Returns vbNullString if the column name is not found in the header.
' @remark Because the search uses xlPart (partial match), a colName of "row" could
'         theoretically match a header like "row_extra". In practice the setup headers are
'         designed with distinct names to avoid ambiguity. The ColumnIndex helper performs the
'         actual Range.Find call.
Public Function Value(ByVal colName As String) As String
    Dim valIndex As Long

    Value = vbNullString
    If ColumnExists(colName) Then
        valIndex = ColumnIndex(colName)
        Value = TableRange.Cells(1, valIndex).Value
    End If
End Function

' @description Returns a BetterArray of category values for the row variable of this table.
'              The categories are retrieved from the linelist specifications dictionary based
'              on the variable name in the "row" field. For example, if the row variable is
'              "sex", this returns the category list for "sex" (e.g., "Male", "Female"). These
'              categories are used to build the row headers of the output analysis table.
'              Internally delegates to CategoriesData("row").
' @return A BetterArray (1-based) containing the category strings for the row variable. If
'         the row variable is not found in the dictionary, returns an empty BetterArray.
Public Property Get RowCategories() As BetterArray
    Set RowCategories = CategoriesData("row")
End Property

' @description Returns a BetterArray of category values for the column variable of this table.
'              For most table types, the categories come from the linelist dictionary based on
'              the variable name in the "column" field. For SpatioTemporal tables, however,
'              the column categories are empty placeholder strings: the number of placeholders
'              is determined by the "n geo" field in the setup row (defaulting to 5 if not
'              specified or not a valid number). These placeholders are filled in later at
'              runtime with actual geographic location names. Internally delegates to
'              CategoriesData("column").
' @return A BetterArray (1-based) containing the category strings for the column variable.
'         For SpatioTemporal, contains N empty strings where N is the "n geo" value. For other
'         types, contains the dictionary categories. Returns an empty BetterArray if the column
'         variable is not found or not specified.
Public Property Get ColumnCategories() As BetterArray
    Set ColumnCategories = CategoriesData("column")
End Property

'@section Private helpers
'===============================================================================

' @description Checks whether a column with the given name exists in the header range. Uses
'              Range.Find with xlPart matching and case-insensitive comparison, so a search
'              for "row" will match a header cell containing "row" regardless of case. This is
'              a guard function called by Value and ColumnIndex before attempting to read or
'              locate a column, preventing errors when optional columns are absent from the
'              setup table.
' @param colName The column name to search for in the header range (partial, case-insensitive).
' @return True if at least one cell in the header range contains the specified name, False
'         otherwise.
Private Function ColumnExists(ByVal colName As String) As Boolean
    ColumnExists = (Not HeaderRange.Find(What:=colName, _
                                         lookAt:=xlPart, _
                                         MatchCase:=False) Is Nothing)
End Function

' @description Returns the 1-based column index of a named column relative to the start of
'              the header range. The index is computed as the absolute column number of the
'              found cell minus the absolute column number of the header's first cell, plus 1.
'              For example, if the header starts at column C and "total" is found at column E,
'              the returned index is 3 (E - C + 1). This index can be used directly with
'              TableRange.Cells(1, index) to read the corresponding cell value. Returns -1 if
'              the column name is not found in the header.
' @param colName The column name to search for (partial match, case-insensitive, via xlPart).
' @return A Long representing the 1-based position of the column within the header range, or
'         -1 if the column does not exist.
Private Function ColumnIndex(ByVal colName As String) As Long
    If ColumnExists(colName) Then
        ColumnIndex = HeaderRange.Find(What:=colName, _
                                       lookAt:=xlPart, _
                                       MatchCase:=False).Column - _
                      HeaderRange.Column + 1
    Else
        ColumnIndex = -1
    End If
End Function

' @description Shared implementation for RowCategories and ColumnCategories that builds a
'              BetterArray of category values from the linelist dictionary for either the "row"
'              or "column" variable. The method first reads the variable name from the
'              specified field using Value(rowOrCol). It then handles a special case for
'              SpatioTemporal tables: when rowOrCol is "column", the categories are not looked
'              up from the dictionary but instead populated with N empty strings (where N
'              comes from the "n geo" setup field, defaulting to 5). These empty placeholders
'              represent geographic locations that will be filled in at runtime when actual
'              data is available. For all other cases, if the variable name exists in the
'              dictionary, the method retrieves the categories via
'              LinelistSpecifications().Categories(nameVar). The resulting BetterArray is
'              cloned before returning to prevent external mutations from affecting internal
'              state.
' @param rowOrCol Either "row" or "column", indicating which variable field to read from the
'                 setup row. This string is passed directly to Value() to retrieve the variable
'                 name.
' @return A cloned BetterArray (1-based LowerBound) containing the category strings. For
'         SpatioTemporal column categories, contains N vbNullString values. For other cases,
'         contains the dictionary-defined categories. Empty if the variable is not found.
' @remark The "n geo" value is read with On Error Resume Next to gracefully handle non-numeric
'         or missing values, falling back to 5 placeholders. The LowerBound is explicitly set
'         to 1 for consistency with other BetterArray usage throughout the codebase.
Private Function CategoriesData(ByVal rowOrCol As String) As BetterArray
    Dim varsCat As BetterArray
    Dim nameVar As String
    Dim vars As ILLVariables
    Dim nbGeo As Long
    Dim counter As Long

    Set vars = LLVariables.Create(Dictionary())
    Set varsCat = New BetterArray
    varsCat.LowerBound = 1

    nameVar = Value(rowOrCol)

    ' For SpatioTemporal, column categories are empty geo placeholders
    If (TableType = TypeSpatioTemporal) And (rowOrCol = "column") Then
        nbGeo = 5
        On Error Resume Next
        nbGeo = CLng(Value("n geo"))
        On Error GoTo 0
        For counter = 1 To nbGeo
            varsCat.Push vbNullString
        Next
    ElseIf vars.Contains(nameVar) Then
        Set varsCat = LinelistSpecifications().Categories(nameVar)
    End If

    Set CategoriesData = varsCat.Clone()
End Function

'@section Error handling
'===============================================================================

' @description Raise a ProjectError-based exception
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

' @description Validates all preconditions required by the Create factory before populating a
'              new TableSpecs instance. Performs the following checks in order:
'              1. headerRng must not be Nothing.
'              2. tableSpecsRng must not be Nothing.
'              3. lData must not be Nothing.
'              4. headerRng and tableSpecsRng must have the same number of columns (ensuring
'                 they come from the same ListObject and column lookups will be valid).
'              5. headerRng.Row must be strictly less than tableSpecsRng.Row (the header must
'                 be above the data row).
'              6. headerRng.Row must be >= 1 (basic sanity check for valid row positioning).
'              If any check fails, a descriptive error is raised via ThrowError with the
'              InvalidArgument error code.
' @param headerRng The header range to validate.
' @param tableSpecsRng The table specification row range to validate.
' @param lData The linelist specifications object to validate.
' @remark This method is called at the very beginning of Create, before any property
'         assignment. It acts as a fail-fast guard to produce clear error messages rather
'         than cryptic runtime errors later during property access.
Private Sub CheckRequirements(ByVal headerRng As Range, _
                              ByVal tableSpecsRng As Range, _
                              ByVal lData As ILinelistSpecs)

    If headerRng Is Nothing Then
        ThrowError InvalidArgument, "Header range not defined"
    End If

    If tableSpecsRng Is Nothing Then
        ThrowError InvalidArgument, "Table specification range not defined"
    End If

    If lData Is Nothing Then
        ThrowError InvalidArgument, "Linelist specifications not defined"
    End If

    If headerRng.Columns.Count <> tableSpecsRng.Columns.Count Then
        ThrowError InvalidArgument, _
                   "Header range and specification range must have same column count"
    End If

    If headerRng.Row >= tableSpecsRng.Row Then
        ThrowError InvalidArgument, _
                   "Header range must be above the specification range"
    End If

    If headerRng.Row < 1 Then
        ThrowError InvalidArgument, "Header range must be at row >= 1"
    End If
End Sub

'@section Interface implementation
'===============================================================================

Private Property Get ITableSpecs_TableType() As AnalysisTablesType
    ITableSpecs_TableType = TableType()
End Property

Private Property Get ITableSpecs_Dictionary() As ILLdictionary
    Set ITableSpecs_Dictionary = Dictionary()
End Property

Private Property Get ITableSpecs_RowCategories() As BetterArray
    Set ITableSpecs_RowCategories = RowCategories()
End Property

Private Property Get ITableSpecs_ColumnCategories() As BetterArray
    Set ITableSpecs_ColumnCategories = ColumnCategories()
End Property

Private Property Get ITableSpecs_IsNewSection() As Boolean
    ITableSpecs_IsNewSection = IsNewSection()
End Property

Private Property Get ITableSpecs_TableId() As String
    ITableSpecs_TableId = TableId()
End Property

Private Property Get ITableSpecs_ValidTable() As Boolean
    ITableSpecs_ValidTable = ValidTable()
End Property

Private Property Get ITableSpecs_HasMissing() As Boolean
    ITableSpecs_HasMissing = HasMissing()
End Property

Private Property Get ITableSpecs_HasTotal() As Boolean
    ITableSpecs_HasTotal = HasTotal()
End Property

Private Property Get ITableSpecs_TotalRequested() As Boolean
    ITableSpecs_TotalRequested = TotalRequested()
End Property

Private Property Get ITableSpecs_HasGraph() As Boolean
    ITableSpecs_HasGraph = HasGraph()
End Property

Private Property Get ITableSpecs_HasPercentage() As Boolean
    ITableSpecs_HasPercentage = HasPercentage()
End Property

Private Property Get ITableSpecs_SpatialTableTypes() As String
    ITableSpecs_SpatialTableTypes = SpatialTableTypes()
End Property

Private Property Get ITableSpecs_Previous() As ITableSpecs
    Set ITableSpecs_Previous = Previous()
End Property

Private Property Get ITableSpecs_NextSpecs(ByVal anchorRng As Range) As ITableSpecs
    Set ITableSpecs_NextSpecs = NextSpecs(anchorRng)
End Property

Private Property Get ITableSpecs_TableSectionId() As String
    ITableSpecs_TableSectionId = TableSectionId()
End Property

Private Function ITableSpecs_Value(ByVal colName As String) As String
    ITableSpecs_Value = Value(colName)
End Function
