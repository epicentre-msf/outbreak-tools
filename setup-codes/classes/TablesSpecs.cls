VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "analysis tables specifications (univariate, bivariate, time series and spatial)"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("analysis tables specifications (univariate, bivariate, time series and spatial)")

Option Explicit
Implements ITablesSpecs

Private Type TTablesSpecs
    hRng As Range
    tRng As Range
    dict As ILLdictionary
    choi As ILLchoice
    reason As String
End Type

Private this As TTablesSpecs
Private Const CLASSNAME As String = "TableSpecs"
Private Const CONTROLCOLUMN As String = "Control"
Private Const CONTROLDETAILSCOLUMN As String = "Control Details"
Private Const TYPECOLUMN As String = "Variable Type"

'Instanciation
'headerRange is the headerRange of the listObject
'tableRange is the range of one line of the table, corresponding to the
'specifications of one table in the analysis. (rowvar, colvar, etc.)
Public Function Create(ByVal headerRng As Range, ByVal tableSpecsRng As Range, _
                       ByVal dict As ILLdictionary, ByVal choi As ILLchoice) As ITablesSpecs

    CheckRequirements headerRng, tableSpecsRng, dict, choi

    With New TablesSpecs
        Set .HeaderRange = headerRng
        Set .TableRange = tableSpecsRng
        Set .Dictionary = dict
        Set .Choice = choi
        Set Create = .Self
    End With
End Function

'Properties
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

Public Property Get Choice() As ILLchoice
    Set Choice = this.choi
End Property

Public Property Set Choice(ByVal choi As ILLchoice)
    Set this.choi = choi
End Property

Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
End Property

Public Property Get Self() As ITablesSpecs
    Set Self = Me
End Property

'------------------------ private properties and methods
'validity reason
Private Property Get ValidityReason() As String
    ValidityReason = this.reason
End Property

Private Property Let ValidityReason(ByVal reason As String)
    this.reason = reason
End Property

'Get the table type of the one table
Private Property Get TableType() As AnalysisTablesType
    Dim hRng As Range
    Dim scope As AnalysisTablesType
    Dim anaType As String

    Set hRng = HeaderRange()
    anaType = LCase(Application.WorksheetFunction.Trim(hRng.Cells(-1, 1).Value))
    Select Case anaType
    Case "global summary"
        scope = TypeGlobalSummary
    Case "univariate analysis"
        scope = TypeUnivariate
    Case "bivariate analysis"
        scope = TypeBivariate
    Case "time series analysis"
        scope = TypeTimeSeries
    Case "spatial analysis"
        scope = TypeSpatial
    Case "spatio-temporal analysis"
        scope = TypeSpatioTemporal
    Case "graph on time series"
        scope = TypeTimeSeriesGraph
    Case Else
        ThrowError "Could not find the table type :(" & ".Actual type is : " & anaType
    End Select

    TableType = scope
End Property

'For spatial analyses, the table can be either on geo or hf (health facility)
Private Property Get SpatialTableTypes() As String
    Dim dict As ILLdictionary
    Dim vars As ILLVariables

    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)

    'Find the spatial table type.
    'Use the spatial table type only if the table is valid (you need to check it before)
    SpatialTableTypes = vars.Value(colName:="Control", varName:=Value("row"))
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
    Case TypeSpatioTemporal
        prefix = "SPT"
    Case TypeTimeSeriesGraph
        prefix = "TSGR"
    End Select

    TableId = prefix & "_tab" & lineNum
End Property

Private Property Get TableSectionId() As String
    If isNewSection() Or (TableType = TypeGlobalSummary) Or (TableType = TypeTimeSeriesGraph) Then
        TableSectionId = TableId()
    Else
        TableSectionId = Previous.TableSectionId
    End If
End Property

Private Property Get HasPercent() As Boolean
    Select Case TableType
    Case TypeGlobalSummary, TypeSpatioTemporal
        HasPercent = False
    Case TypeUnivariate, TypeSpatial
        HasPercent = (Value("percentage") = "yes")
    Case TypeBivariate
        HasPercent = (Value("percentage")) = "row" Or (Value("percentage") = "column") Or (Value("percentage") = "total")
    Case TypeTimeSeries
        HasPercent = ((Value("percentage")) = "row" Or (Value("percentage") = "column")) And HasTotal()
    Case Else
        HasPercent = False
    End Select
End Property

Private Property Get HasTotal() As Boolean
    Select Case TableType

    Case TypeGlobalSummary, TypeSpatioTemporal
        HasTotal = False
    Case TypeUnivariate
        HasTotal = True
    Case TypeBivariate
        HasTotal = True
    Case TypeTimeSeries
        HasTotal = ((Value("total") = "yes") Or (Value("percentage") = "row") Or _
                    (Value("percentage") = "column")) And (Value("column") <> vbNullString)
    Case TypeSpatial
        HasTotal = (Value("column") <> vbNullString)
    Case Else
        HasTotal = False
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
        HasGraph = (Value("graph") = "yes") And (Value("row") <> vbNullString)
    Case TypeSpatioTemporal
        HasGraph = (Value("graph") = "yes")
    Case Else
        HasGraph = False
    End Select
End Property

Private Property Get HasMissing() As Boolean
    Select Case TableType

    Case TypeGlobalSummary, TypeSpatioTemporal
        HasMissing = False
    Case TypeUnivariate
        HasMissing = Value("missing") = "yes"
    Case TypeBivariate
        HasMissing = (Value("missing") = "row") Or _
                                                (Value("missing") = "column") Or _
                                                (Value("missing") = "all")
    Case TypeTimeSeries
        HasMissing = (Value("missing") = "yes") And (Value("column") <> vbNullString)
    Case TypeSpatial
        HasMissing = (Value("missing") = "yes") And (Value("column") <> vbNullString)
    'For spatio - temporal, there is no missing on column
    Case Else
        HasMissing = False
    End Select
End Property

Public Property Get ValidTable() As Boolean

    Dim dict As ILLdictionary
    Dim Valid As Boolean
    Dim rowVar As String
    Dim colVar As String
    Dim colControl As String
    Dim rowControl As String
    Dim rowType As String
    Dim vars As ILLVariables
    Dim graphId As String
    Dim seriesId As String
    Dim graphOrder As String
    Dim choicesVal As String

    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)

    rowVar = Value("row")
    rowControl = vars.Value(colName:=CONTROLCOLUMN, varName:=rowVar)

    colVar = Value("column")
    colControl = vars.Value(colName:=CONTROLCOLUMN, varName:=colVar)
    ValidityReason = "The table seems valid"

    Select Case TableType

    Case TypeGlobalSummary
        Valid = (Value("label") <> vbNullString) And (Value("function") <> vbNullString)

        If Not Valid Then
        'add reasons of non validitydisn
            If (Value("label") = vbNullString) And (Value("function") <> vbNullString) Then
                ValidityReason = "Summary label is missing"
            ElseIf (Value("label") <> vbNullString) And (Value("function") = vbNullString) Then
                ValidityReason = "Summary function is missing"
            ElseIf (Value("label") = vbNullString) And (Value("function") = vbNullString) Then
                ValidityReason = "Summary label and summary function are missing"
            End If
        End If

    Case TypeUnivariate
        Valid = vars.Contains(rowVar) And (rowControl = "choice_manual" Or rowControl = "choice_formula")

        If Not Valid Then
            If (rowVar = vbNullString) Then
                ValidityReason = "The row variable is not specified"
            ElseIf (Not vars.Contains(rowVar)) Then
                ValidityReason = "The variable " & rowVar & " does not exists in the dictionary"
            Else
                ValidityReason = "The control of variable " & rowVar & " is " & rowControl & _
                              ", which is not choice_manual or choice_formula"
            End If
        End If
    Case TypeBivariate
        Valid = vars.Contains(rowVar) And vars.Contains(colVar) And _
                                                                (colControl = "choice_manual" Or colControl = "choice_formula") And _
                                                                (rowControl = "choice_manual" Or rowControl = "choice_formula")
        If Not Valid Then
            If (rowVar = vbNullString) Then
                ValidityReason = "The row variable is not specified"
            ElseIf (colVar = vbNullString) Then
                ValidityReason = "The column variable is not specified"
            ElseIf (Not vars.Contains(rowVar)) Then
                ValidityReason = "The row variable " & rowVar & " does not exists in the dictionary"
            ElseIf (Not vars.Contains(colVar)) Then
                ValidityReason = "The column variable " & colVar & " does not exists in the dictionary"
            ElseIf (colControl <> "choice_manual" And colControl <> "choice_formula") Then
                ValidityReason = "The control of variable " & colVar & " is " & colControl & _
                              "which is not choice_manual or choice_formula"
            ElseIf (rowControl <> "choice_manual" And rowControl <> "choice_formula") Then
                ValidityReason = "The control of variable " & rowVar & " is " & rowControl & _
                              "which is not choice_manual or choice_formula"
            End If
        End If

    Case TypeTimeSeries
        rowType = vars.Value(colName:=TYPECOLUMN, varName:=rowVar)
        Valid = (colVar = vbNullString Or _
                 colControl = "choice_manual" Or _
                 colControl = "choice_formula") And _
                 (rowType = "date")
        If Not Valid Then
            If (rowVar = vbNullString) Then
                ValidityReason = "The row variable is not specified"
            ElseIf (Not vars.Contains(rowVar)) Then
                ValidityReason = "The row variable " & rowVar & " does not exists in the dictionary"
            ElseIf (colVar <> vbNullString And (Not vars.Contains(colVar))) Then
                ValidityReason = "The column variable " & colVar & " does not exists in the dictionary"
            ElseIf (colVar <> vbNullString And colControl <> "choice_manual" And colControl <> "choice_formula") Then
                ValidityReason = "The control of variable " & _
                              colVar & " is " & colControl & _
                              "which is not choice_manual or choice_formula"
            ElseIf (rowType <> "date") Then
                ValidityReason = "The variable " & rowVar & _
                              " is not of type date. Time series should be on date variables."
            End If
        End If
    Case TypeSpatial
        Valid = (rowControl = "geo" Or rowControl = "hf") And _
                (colVar = vbNullString Or colControl = "choice_manual" Or colControl = "choice_formula")

        If Not Valid Then
            If (rowVar = vbNullString) Then
                ValidityReason = "The row variable is not specified"
            ElseIf (colVar <> vbNullString And colControl <> "choice_manual" And colControl <> "choice_formula") Then
                ValidityReason = "The control of variable " & colVar & " is " & rowControl & _
                "which is not choice_manual or choice_formula"
            ElseIf (rowControl <> "geo" And rowControl <> "hf") Then
                ValidityReason = "The control of variable " & rowVar & " is " & rowControl & _
                " which is not geo or hf. Spatial tables should be on geo or hf variables."
            End If
        End If

    Case TypeSpatioTemporal
        rowType = vars.Value(colName:=TYPECOLUMN, varName:=rowVar)
        Valid = (rowType = "date") And (colControl = "geo" Or colControl = "hf")

        'Add non validity reasons
        If Not Valid Then
            If (rowVar = vbNullString) Then
                ValidityReason = "The row variable is not specified"
            ElseIf (colVar = vbNullString) Then
                ValidityReason = "The column variable is not specified"
            ElseIf (rowType <> "date") Then
                ValidityReason = "The row variable " & rowVar & _ 
                 "is not of type date. Spatio temporal tables should have dates as row variables." 
            ElseIf (colControl <> "geo" And colControl <> "hf") Then
                ValidityReason = "The column variable " & colVar & " is " & _ 
                "neither geo nor hf. Columns of spatio-temporal tables shoud be geo or hf variables."
            End If
        End If

    Case TypeTimeSeriesGraph
        
        graphId = Value("graph id")
        seriesId = Value("series id")
        graphOrder = Value("graph order")
        
        Valid = (rowVar <> vbNullString) And (graphId <> vbNullString) And _ 
                (seriesId <> vbNullString) And (graphOrder <> vbNullString)

        If Not Valid Then
            If (rowVar = vbNullString) Then
                ValidityReason = "Unable to find the time variable. Series title may be missing."
            ElseIf (seriesId = vbNullString) Then
                ValidityReason = "Unable to find the series ID. Series title may be missing."
            ElseIf (graphOrder = vbNullString) Or (graphId = vbNullString) Then
                ValidityReason = "Unable to find the 'graph ID'/'graph order'. Graph title may be missing."
            End If
        End If
    End Select
    ValidTable = Valid
End Property


Private Function ColumnExists(ByVal colName As String) As Boolean
    Dim hRng As Range
    Set hRng = HeaderRange
    ColumnExists = (Not hRng.Find(What:=colName, lookAt:=xlPart, MatchCase:=False) Is Nothing)
End Function

Private Function ColumnIndex(ByVal colName As String) As Long
    Dim hRng As Range
    Set hRng = HeaderRange
    If ColumnExists(colName) Then
        '@Ignore ExcelMemberMayReturnNothing
        ColumnIndex = hRng.Find(What:=colName, lookAt:=xlPart, MatchCase:=False).Column - hRng.Column + 1
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

'Get the categories of one variable on row or columns
Private Property Get Data(ByVal rowOrCol As String) As BetterArray
    Dim varsCat As BetterArray
    Dim dict As ILLdictionary
    Dim nameVar As String
    Dim varControl As String 'control and control details of the variable
    Dim varControlDet As String
    Dim vars As ILLVariables
    Dim choi As ILLchoice
    Dim choiform As ChoiceFormula

    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)
    Set varsCat = New BetterArray
    'name var is the name of the variable from which we want to
    'retrieve the categories

    nameVar = Value(rowOrCol)
    If vars.Contains(nameVar) Then
        varControl = vars.Value(varName:=nameVar, colName:=CONTROLCOLUMN)
        varControlDet = vars.Value(varName:=nameVar, colName:=CONTROLDETAILSCOLUMN)

        'For choices, categories are drown from the choices
        'For choice formula, use the choice formula object
        If (varControl = "choice_manual") Then
            Set choi = Choice()
            Set varsCat = choi.Categories(varControlDet)
        ElseIf (varControl = "choice_formula") Then
            Set choiform = ChoiceFormula.Create(varControlDet)
            Set varsCat = choiform.Categories()
        End If  
    End If
    Set Data = varsCat.Clone()
End Property

Private Property Get Previous() As ITablesSpecs
    Dim hRng As Range
    Dim tRng As Range
    Dim actRng As Range                          'Actual range
    Dim precSec As ITablesSpecs
    Dim dict As ILLdictionary
    Dim choi As ILLchoice


    If isNewSection() Then ThrowErrorUnableInstanciate "Your table is on a new section, I can't return the previous table"

    Set hRng = HeaderRange()
    Set actRng = TableRange()
    Set tRng = actRng.Offset(-1)
    Set dict = Dictionary()
    Set choi = Choice()

    'Loop until valid table
    Set precSec = Create(hRng, tRng, dict, choi)
    Set tRng = tRng.Offset(-1)

    Do While (Not precSec.ValidTable()) And (Not (tRng.Row <= hRng.Row))
        Set precSec = Create(hRng, tRng, dict, choi)
        If precSec.isNewSection() Then Exit Do
        Set tRng = tRng.Offset(-1)
    Loop

    Set Previous = precSec
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

    isNewSection = (Not oldSec) And (TableType <> TypeGlobalSummary) And _ 
                  (TableType <> TypeTimeSeriesGraph)
End Property

'Errors
'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error: " & errorMessage
End Sub


Private Sub CheckRequirements(ByVal headerRng As Range, ByVal tableSpecsRng As Range, _
                            ByVal dict As ILLdictionary, ByVal choi As ILLchoice)
    If headerRng Is Nothing Then ThrowErrorUnableInstanciate "Header range not defined"
    If tableSpecsRng Is Nothing Then ThrowErrorUnableInstanciate "Table range not defined"
    If dict Is Nothing Then ThrowErrorUnableInstanciate "dictionary not defined"
    If choi Is Nothing Then ThrowErrorUnableInstanciate "dictionary not defined"
    If headerRng.Columns.Count <> tableSpecsRng.Columns.Count Then ThrowErrorUnableInstanciate "Header range and sepicification range don't have same length"
    If headerRng.Row > tableSpecsRng.Row Then ThrowErrorUnableInstanciate "Seems like header comes range after specification range. Maybe you reversed the order? "
    If headerRng.Row < 1 Then ThrowErrorUnableInstanciate "Header range should be at row > 1"
End Sub

'Interface
Private Property Get ITablesSpecs_TableType() As AnalysisTablesType
    ITablesSpecs_TableType = TableType()
End Property

Private Property Get ITablesSpecs_RowCategories() As BetterArray
    Set ITablesSpecs_RowCategories = Data("row")
End Property

Private Property Get ITablesSpecs_ColumnCategories() As BetterArray
    Set ITablesSpecs_ColumnCategories = Data("column")
End Property

Private Property Get ITablesSpecs_TableId() As String
    ITablesSpecs_TableId = TableId()
End Property

Private Property Get ITablesSpecs_isNewSection() As Boolean
    ITablesSpecs_isNewSection = isNewSection
End Property

Private Function ITablesSpecs_Value(ByVal colName As String) As String
    ITablesSpecs_Value = Value(colName)
End Function

Private Property Get ITablesSpecs_HasTotal() As Boolean
    ITablesSpecs_HasTotal = HasTotal
End Property

Private Property Get ITablesSpecs_HasPercentage() As Boolean
    ITablesSpecs_HasPercentage = HasPercent
End Property

Private Property Get ITablesSpecs_HasMissing() As Boolean
    ITablesSpecs_HasMissing = HasMissing
End Property

Private Property Get ITablesSpecs_HasGraph() As Boolean
    ITablesSpecs_HasGraph = HasGraph()
End Property

Private Property Get ITablesSpecs_Previous() As ITablesSpecs
    Set ITablesSpecs_Previous = Previous
End Property

Private Property Get ITablesSpecs_TableSectionId() As String
    ITablesSpecs_TableSectionId = TableSectionId
End Property

Private Property Get ITablesSpecs_ValidTable() As Boolean
    ITablesSpecs_ValidTable = ValidTable()
End Property

Private Property Get ITablesSpecs_SpatialTableTypes() As String
    ITablesSpecs_SpatialTableTypes = SpatialTableTypes()
End Property

Private Property Get ITablesSpecs_ValidityReason() As String
    ITablesSpecs_ValidityReason = ValidityReason()
End Property


