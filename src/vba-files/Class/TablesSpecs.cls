VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Analysis")
'@ModuleDescription("analysis tables specifications")
'@IgnoreModule

Option Explicit
Implements ITablesSpecs

Private Type TTablesSpecs
    hRng As Range
    tRng As Range
    lData As ILinelistSpecs
    tType As AnalysisTablesType
End Type

Private this As TTablesSpecs
Private Const CLASSNAME As String = "TableSpecs"

'Instanciation

Public Function Create(headerRng As Range, tableSpecsRng As Range, linelistData As ILinelistSpecs) As ITablesSpecs
    If headerRng Is Nothing Then ThrowErrorUnableInstanciate "Header range not defined"
    If tableSpecsRng Is Nothing Then ThrowErrorUnableInstanciate "Table range not defined"
    If linelistData Is Nothing Then ThrowErrorUnableInstanciate "Linelist specifications not defined"
    If headerRng.Columns.Count <> tableSpecsRng.Columns.Count Then ThrowErrorUnableInstanciate "Header range and sepicification range don't have same length"
    If headerRng.Row > tableSpecsRng.Row Then ThrowErrorUnableInstanciate "Seems like header comes range after specification range. Maybe you reversed the order? "
    If headerRng.Row < 1 Then ThrowErrorUnableInstanciate "Header range should be at row > 1"

    With New TablesSpecs
        Set .HeaderRange = headerRng
        Set .TableRange = tableSpecsRng
        Set .LinelistSpecifications = linelistData
        Set Create = .Self
    End With
End Function

'Properties
Public Property Get HeaderRange() As Range
    Set HeaderRange = this.hRng
End Property

Public Property Get TableRange() As Range
    Set TableRange = this.tRng
End Property

Public Property Get LinelistSpecifications() As ILinelistSpecs
    Set LinelistSpecifications = this.lData
End Property

Public Property Set HeaderRange(ByVal hRng As Range)
    Set this.hRng = hRng
End Property

Public Property Set TableRange(ByVal tRng As Range)
    Set this.tRng = tRng
End Property

Public Property Set LinelistSpecifications(ByVal lData As ILinelistSpecs)
    Set this.lData = lData
End Property

Public Property Get TableType() As AnalysisTablesType
    Dim hRng As Range
    Dim tType As AnalysisTablesType
    Dim anaType As String
    
    Set hRng = HeaderRange()
    anaType = LCase(Application.WorksheetFunction.Trim(hRng.Cells(-1, 1).Value))
    Select Case anaType
    Case "global summary"
        tType = TypeGlobalSummary
    Case "univariate analysis"
        tType = TypeUnivariate
    Case "bivariate analysis"
        tType = TypeBivariate
    Case "time series analysis"
        tType = TypeTimeSeries
    Case "spatial analysis"
        tType = TypeSpatial
    Case Else
        ThrowError "Could not find the table type [uni-bivariate, spatial or timeseries? :(]" & ".Actual type is : " & anaType
    End Select
  
    TableType = tType
End Property

Public Property Get Self() As ITablesSpecs
    Set Self = Me
End Property

Private Property Get TableId() As String
    Dim prefix As String
    Dim tRng As Range
    Dim hRng As Range
    Dim lineNum As Long

    Set tRng = TableRange
    Set hRng = HeaderRange

    lineNum = tRng.Row - hRng.Row

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
    End Select

    TableId = prefix & "_tab" & lineNum
End Property

Private Property Get TableSectionId() As String
    If isNewSection() Or (TableType = TypeGlobalSummary) Then
        TableSectionId = TableId()
    Else
        TableSectionId = Previous.TableSectionId
    End If
End Property

Private Property Get HasPercent() As Boolean
    Select Case TableType

    Case TypeGlobalSummary
        HasPercent = False
    Case TypeUnivariate
        HasPercent = (Value("percentage") = "yes")
    Case TypeBivariate
        HasPercent = (Value("percentage")) = "row" Or (Value("percentage") = "column") Or (Value("percentage") = "total")
    Case TypeTimeSeries
        HasPercent = (Value("percentage")) = "row" Or (Value("percentage") = "column")
    Case TypeSpatial
        HasPercent = (Value("percentage") = "column" And Value("column") <> vbNullString) Or _
                                                                                          (Value("percentage") = "row") Or (Value("percentage") = "total")
    End Select
End Property

Private Property Get HasTotal() As Boolean
    Select Case TableType

    Case TypeGlobalSummary
        HasTotal = False
    Case TypeUnivariate
        HasTotal = True
    Case TypeBivariate
        HasTotal = True
    Case TypeTimeSeries
        HasTotal = (Value("total") = "yes")
    Case TypeSpatial
        HasTotal = (Value("total") <> "no")
    End Select
End Property

Public Property Get HasGraph() As Boolean
    Select Case TableType
    Case TypeGlobalSummary
        HasGraph = False
    Case TypeUnivariate
        HasGraph = (Value("graph") = "yes")
    Case TypeBivariate
        HasGraph = (Value("graph") = "percentage" Or Value("graph") = "values" Or Value("graph") = "both")
    Case TypeSpatial
        HasGraph = (Value("graph") = "yes")
    End Select
End Property

Private Property Get HasMissing() As Boolean
    Select Case TableType

    Case TypeGlobalSummary
        HasMissing = False
    Case TypeUnivariate
        HasMissing = Value("missing") = "yes"
    Case TypeBivariate
        HasMissing = (Value("missing")) = "row" Or _
                                          (Value("missing") = "column") Or _
                                          (Value("missing") = "all")
    Case TypeTimeSeries
        HasMissing = (Value("missing") = "yes")
    Case TypeSpatial
        HasMissing = (Value("missing")) = "row" Or _
                                          (Value("missing") = "column") Or _
                                          (Value("missing") = "all")
    End Select

End Property

Public Property Get ValidTable() As Boolean
    
    Dim dict As ILLdictionary
    Dim lData As ILinelistSpecs
    Dim valid As Boolean
    Dim rowVar As String
    Dim colVar As String
    Dim colControl As String
    Dim rowControl As String
    Dim vars As ILLVariables

    Set vars = LLVariables.Create(dict)

    Set lData = LinelistSpecifications
    Set dict = lData.Dictionary

    rowVar = Value("row")
    rowControl = vars.Value(colName:="choices", varName:=rowVar)

    colVar = Value("column")
    colControl = vars.Value(colName:="choices", varName:=colVar)

    Select Case TableType
    Case TypeGlobalSummary
        valid = True
    Case TypeUnivariate
        valid = vars.Contains(rowVar) And (rowControl = "choice")
    Case TypeBivariate
        valid = vars.Contains(rowVar) And vars.Contains(colVar) And (colControl = "choice")
    Case TypeTimeSeries
        rowVar = Value("row")
        colVar = Value("column")
        valid = (dict.VariableExists(rowVar) And (colVar = vbNullString Or dict.VariableExists(colVar)))
    Case TypeSpatial
        rowVar = Value("row")
        colVar = Value("column")
        valid = (dict.VariableExists("adm1_" & rowVar) And (colVar = vbNullString Or dict.VariableExists(colVar)))
    End Select

    ValidTable = valid
End Property

Private Property Get Dictionary() As ILLdictionary
    Dim lData As ILinelistSpecs
    Set lData = LinelistSpecifications()
    Set Dictionary = lData.Dictionary()
End Property

Private Function ColumnExists(ByVal colName As String) As Boolean
    Dim hRng As Range
    Set hRng = HeaderRange
    ColumnExists = (Not hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=False) Is Nothing)
End Function

Private Function ColumnIndex(ByVal colName As String) As Long
    Dim hRng As Range
    Set hRng = HeaderRange
    If ColumnExists(colName) Then
        ColumnIndex = hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=False).Column - hRng.Column + 1
    Else
        ColumnIndex = -1
    End If
End Function

Private Function Value(ByVal colName As String) As String
    Dim retVal As String
    Dim valIndex As Long
    Dim tRng As Range

    Set tRng = TableRange
    retVal = vbNullString

    If ColumnExists(colName) Then
        valIndex = ColumnIndex(colName)
        retVal = tRng.Cells(1, valIndex).Value
    End If

    Value = retVal
End Function

Private Property Get Data(ByVal rowOrCol As String) As BetterArray
    Dim varsCat As BetterArray
    Dim nameVar As String
    Dim vars As ILLVariables
    Dim lData As ILinelistSpecs

    Set lData = LinelistSpecifications()
    Set vars = LLVariables.Create(lData.Dictionary)
    Set varsCat = New BetterArray
    nameVar = Value(rowOrCol)
    If vars.Contains(nameVar) Then
        Set varsCat = lData.Categories(nameVar)
    End If
    Set Data = varsCat.Clone()
End Property

Private Property Get Previous() As ITablesSpecs
    Dim lData As ILinelistSpecs
    Dim hRng As Range
    Dim tRng As Range
    Dim actRng As Range                          'Actual range
    Dim taType As AnalysisTablesType
    Dim endCol As Long
    
    
    If isNewSection() Then ThrowErrorUnableInstanciate "Your table is on a new section, I can't return the previous table"
    
    taType = TableType()
    Set hRng = HeaderRange()
    Set actRng = TableRange()
    Set lData = LinelistSpecifications()
    Set tRng = actRng.Offset(-1)
    
    Set Previous = Create(hRng, tRng, lData)

End Property

Private Property Get isNewSection() As Boolean
    Dim sectIndex As Long
    Dim ActualSection As String
    Dim previousSection As String
    Dim oldSec As Boolean
    Dim tRng As Range
    Dim hRng As Range
    Dim prevCell As Range

    Set tRng = TableRange()
    Set hRng = HeaderRange()
    oldSec = False

    If ColumnExists("section") Then
        sectIndex = ColumnIndex("section")
        ActualSection = Value("section")
        Set prevCell = tRng.Cells(0, sectIndex)  '0 for previous line (we can also offset)
        previousSection = prevCell.Value
        oldSec = (ActualSection = previousSection) And (prevCell.Row <> hRng.Row)
    End If

    isNewSection = (Not oldSec) And (TableType <> TypeGlobalSummary)
End Property

'Errors
'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error: " & errorMessage
End Sub

'Interface
Public Property Get ITablesSpecs_TableType() As AnalysisTablesType
    ITablesSpecs_TableType = TableType()
End Property

Public Property Get ITablesSpecs_RowCategories() As BetterArray
    Set ITablesSpecs_RowCategories = Data("row")
End Property

Public Property Get ITablesSpecs_ColumnCategories() As BetterArray
    Set ITablesSpecs_ColumnCategories = Data("column")
End Property

Public Property Get ITablesSpecs_Dictionary() As ILLdictionary
    Set ITablesSpecs_Dictionary = Dictionary()
End Property

Public Property Get ITablesSpecs_TableId() As String
    ITablesSpecs_TableId = TableId()
End Property

Public Property Get ITablesSpecs_isNewSection() As Boolean
    ITablesSpecs_isNewSection = isNewSection
End Property

Public Function ITablesSpecs_Value(ByVal colName As String) As String
    ITablesSpecs_Value = Value(colName)
End Function

Public Property Get ITablesSpecs_HasTotal() As Boolean
    ITablesSpecs_HasTotal = HasTotal
End Property

Public Property Get ITablesSpecs_HasPercentage() As Boolean
    ITablesSpecs_HasPercentage = HasPercent
End Property

Public Property Get ITablesSpecs_HasMissing() As Boolean
    ITablesSpecs_HasMissing = HasMissing
End Property

Public Property Get ITablesSpecs_HasGraph() As Boolean
    ITablesSpecs_HasGraph = HasGraph()
End Property

Public Property Get ITablesSpecs_Previous() As ITablesSpecs
    Set ITablesSpecs_Previous = Previous
End Property

Public Property Get ITablesSpecs_TableSectionId() As String
    ITablesSpecs_TableSectionId = TableSectionId
End Property

Public Property Get ITablesSpecs_ValidTable() As Boolean
    ITablesSpecs_ValidTable = ValidTable()
End Property

