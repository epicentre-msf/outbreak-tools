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

' @description Create a new TableSpecs from a setup listobject row
' @param headerRng Header row of the analysis setup listobject
' @param tableSpecsRng Data row for this specific table
' @param lData Linelist specifications object
' @return ITableSpecs interface to the new instance
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

' @description Determine the analysis table type from the setup header
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

' @description Unique identifier for this table: prefix + row offset from header
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

' @description Id of the first table in the current section.
'              For GlobalSummary or new sections, it is this table's own id.
Public Property Get TableSectionId() As String
    If IsNewSection() Or (TableType = TypeGlobalSummary) Then
        TableSectionId = TableId()
    Else
        TableSectionId = Previous.TableSectionId
    End If
End Property

' @description Dictionary from linelist specifications
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = LinelistSpecifications().Dictionary()
End Property

' @description Spatial sub-type based on variable prefix in dictionary.
'              Returns "hf" if health-facility prefix exists, "geo" if admin prefix exists.
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

' @description Whether this table specification is valid for its type.
'              Checks variable existence, control types, and required fields.
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

' @description Whether this table starts a new section.
'              Compares "section" column value with the previous row.
'              GlobalSummary tables are never considered new sections.
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

' @description Whether total is needed for computation.
'              For time series: True when total="yes" OR percentage="row"/"column"
'              (because percentage formulas reference the total column data).
'              For univariate/bivariate: always True.
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

' @description Whether the user explicitly requested total (total="yes").
'              Used to control visibility of the total column in temporal tables:
'              when HasTotal is True only because of percentage, the total column
'              exists for computation but should be hidden from view.
Public Property Get TotalRequested() As Boolean
    TotalRequested = (Value("total") = "yes")
End Property

' @description Whether missing data should be included.
'              Time series/spatial require a column variable to have missing.
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

' @description Whether percentage is enabled for this table type.
'              SpatioTemporal and GlobalSummary never have percentage.
'              Time series/spatial require HasTotal for percentage.
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

' @description Whether a graph should be created for this table.
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

' @description Previous valid table specification in the same section.
'              Walks backward from this row, skipping invalid tables.
'              Throws if called on a new-section table.
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

' @description Next valid table specification after this row.
' @param anchorRng Last row of the setup listobject
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

' @description Retrieve a column value from the table specification row.
'              Case-insensitive column name matching.
' @param colName Column name to look up in the header
' @return String value at the intersection; vbNullString if column not found
Public Function Value(ByVal colName As String) As String
    Dim valIndex As Long

    Value = vbNullString
    If ColumnExists(colName) Then
        valIndex = ColumnIndex(colName)
        Value = TableRange.Cells(1, valIndex).Value
    End If
End Function

' @description Row categories from the dictionary for the row variable
Public Property Get RowCategories() As BetterArray
    Set RowCategories = CategoriesData("row")
End Property

' @description Column categories from the dictionary for the column variable.
'              For SpatioTemporal, returns empty placeholders based on "n geo" setting.
Public Property Get ColumnCategories() As BetterArray
    Set ColumnCategories = CategoriesData("column")
End Property

'@section Private helpers
'===============================================================================

' @description Whether a column name exists in the header range
Private Function ColumnExists(ByVal colName As String) As Boolean
    ColumnExists = (Not HeaderRange.Find(What:=colName, _
                                         lookAt:=xlPart, _
                                         MatchCase:=False) Is Nothing)
End Function

' @description Get 1-based column index relative to header start
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

' @description Build category data for row or column variable
' @param rowOrCol "row" or "column" to select the variable
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

' @description Validate factory requirements
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
