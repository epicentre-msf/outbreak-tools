VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpatialTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Spatial table creation at build time"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Spatial table creation at build time")
'@IgnoreModule UnrecognizedAnnotation, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ISpatialTables

'@section State
'===============================================================================

Private Type TSpatialTables
    sh As Worksheet
    tab As ICrossTable
    condVars As BetterArray
    condTests As BetterArray
End Type

Private this As TSpatialTables
Private Const CLASS_NAME As String = "SpatialTables"
Private Const SPATIAL_SHEET As String = "spatial_tables__"
Private Const ADMIN As String = "adm"
Private Const HF As String = "hf"
Private Const CONCAT As String = "concat_"

'@section Factory
'===============================================================================

' @description Create a new SpatialTables instance from a cross-table.
'              Validates the cross-table and that the spatial sheet exists.
' @param tabl The cross-table whose row variable defines the spatial tables
' @return ISpatialTables interface pointer
Public Function Create(ByVal tabl As ICrossTable) As ISpatialTables
    CheckRequirements tabl

    With New SpatialTables
        Set .Table = tabl
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ISpatialTables
    Set Self = Me
End Property

Public Property Get Table() As ICrossTable
    Set Table = this.tab
End Property

Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tab = tabl
End Property

'@section Worksheet
'===============================================================================

' @description Lazy-loaded spatial worksheet reference.
Private Property Get Wksh() As Worksheet
    If this.sh Is Nothing Then
        Dim wb As Workbook
        Set wb = Table.Wksh().Parent
        Set this.sh = wb.Worksheets(SPATIAL_SHEET)
    End If

    Set Wksh = this.sh
End Property

'@section Naming Conventions
'===============================================================================

' @description Build the header named range name for a spatial table.
Private Property Get CurrentHeaderName(ByVal varName As String, _
                                       ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

' @description Build the ListObject name for a spatial table.
Private Property Get CurrentLoName(ByVal varName As String, _
                                   ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'@section Exists
'===============================================================================

' @description Check if a variable already has all 4 admin-level ListObjects.
Private Function Exists(ByVal varName As String) As Boolean
    Dim counter As Long
    Dim testExistence As Boolean

    For counter = 1 To 4
        testExistence = LoExists(CurrentLoName(varName, ADMIN & counter))
        If Not testExistence Then Exit For
    Next

    Exists = testExistence
End Function

' @description Check if a ListObject exists on the spatial sheet.
Private Function LoExists(ByVal loName As String) As Boolean
    Dim Lo As ListObject

    On Error Resume Next
    Set Lo = Wksh().ListObjects(loName)
    On Error GoTo 0

    LoExists = Not (Lo Is Nothing)
End Function

'@section Column Helpers
'===============================================================================

' @description Find the next available column on the spatial sheet.
Private Property Get LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long

    Set sh = Wksh()
    colIndex = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column
    colIndex = IIf(colIndex = 3, 7, colIndex + 2)

    LastColumn = colIndex
End Property

'@section Formula Helpers
'===============================================================================

' @description Build a geo population INDEX/MATCH formula for an admin level.
Private Function GeoPopulationFunction(ByVal adminName As String, _
                                       ByVal rangeAddress As String) As String
    Dim tabName As String
    Dim concatName As String
    Dim centralPart As String
    Dim popColIndex As Long

    tabName = "T_" & UCase(adminName)
    concatName = LCase(adminName) & "_concat"

    On Error Resume Next
    popColIndex = CLng(Replace(adminName, ADMIN, vbNullString)) + 1
    On Error GoTo 0

    centralPart = "INDEX(" & tabName & ", MATCH(" & rangeAddress & _
                  ", " & concatName & ", 0), " & popColIndex & ")"
    GeoPopulationFunction = "IF(ISERROR(" & centralPart & "), 0," & centralPart & ")"
End Function

' @description Test whether a formula expression is valid by evaluating it on a cell.
Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim prevAlert As Long
    Dim sh As Worksheet

    prevAlert = Application.DisplayAlerts
    Set sh = Wksh()

    With sh
        Application.DisplayAlerts = False
        .Cells(1, 1).Clear

        On Error GoTo Fail
        .Cells(1, 1).FormulaArray = "= " & frmName
        .Cells(1, 1).Clear
        Application.DisplayAlerts = prevAlert
        FailedFormula = False
        Exit Function
    End With

Fail:
    FailedFormula = True
    Application.DisplayAlerts = prevAlert
End Function

' @description Empty string constant for formula conditions.
Private Function EmpStr() As String
    EmpStr = Chr(34) & Chr(34)
End Function

'@section Condition Management
'===============================================================================

' @description Add a variable/condition pair for formula generation.
Private Sub AddTableConditions(ByVal varName As String, ByVal conditionName As String)
    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
    End If

    If varName <> vbNullString Then
        this.condVars.Push varName
        this.condTests.Push conditionName
    End If
End Sub

' @description Clear condition arrays, optionally removing only the last N entries.
Private Sub ClearTableConditions(Optional ByVal removeNumber As Long = 0)
    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
        removeNumber = 0
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
        removeNumber = 0
    End If

    If removeNumber <= 0 Then
        this.condVars.Clear
        this.condTests.Clear
    Else
        Dim counter As Long
        For counter = 1 To removeNumber
            this.condVars.Pop
            this.condTests.Pop
        Next
    End If
End Sub

' @description Build an Excel formula string using the current conditions.
Private Property Get CurrentExcelFormula(ByVal tabFormObject As IFormulas, _
                                         Optional ByVal scope As Byte = FormulaScopeAnalysis, _
                                         Optional ByVal prefix As String = "f") As String
    Dim cond As IFormulaCondition
    Dim eform As String

    If Not tabFormObject.Valid() Then Exit Property

    If scope = FormulaScopeLinelist Then
        eform = tabFormObject.ParsedLinelistFormula(useTableName:=True, tablePrefix:=prefix)
    Else
        If this.condVars Is Nothing Or this.condTests Is Nothing Then Exit Property
        Set cond = FormulaCondition.Create(this.condVars, this.condTests)
        eform = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:=prefix)
    End If

    If FailedFormula(eform) Then eform = vbNullString

    CurrentExcelFormula = eform
End Property

'@section Prepare
'===============================================================================

' @description Initialize the spatial sheet with the "listofgeovars" ListObject
'              and named ranges for scratch/test columns.
Private Sub Prepare()
    Dim rng As Range
    Dim sh As Worksheet

    Set sh = Wksh()

    With sh
        Set rng = .Range(.Cells(1, 3), .Cells(2, 3))
        rng.Cells(1, 1).Value = "listofvars"
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "listofgeovars"

        .Cells(1, 5).Name = "RNG_PastingCol"
        .Cells(1, 1).Name = "RNG_TestingFormula"
    End With
End Sub

' @description Register a variable name in the listofgeovars ListObject.
Private Sub AddVarNameToList(ByVal varName As String)
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim nbRows As Long

    Set sh = Wksh()
    Set Lo = sh.ListObjects("listofgeovars")

    nbRows = Application.WorksheetFunction.CountA(Lo.Range)
    sh.Cells(nbRows + 1, Lo.Range.Column).Value = varName
End Sub

'@section Add
'===============================================================================

' @description Create spatial ListObjects for the current table's row variable.
'              Builds 4 admin-level tables (or 1 HF table) with formula,
'              population, and attack rate columns.
Private Sub Add(ByVal formObject As IFormulas)
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim varName As String
    Dim headerName As String
    Dim loName As String
    Dim rng As Range
    Dim counter As Long
    Dim specs As ITableSpecs
    Dim tabl As ICrossTable
    Dim cellRng As Range
    Dim finalFormula As String
    Dim adminName As String
    Dim rowVarName As String
    Dim inputVarName As String
    Dim endCounter As Long
    Dim spatType As String

    Set tabl = Table
    Set specs = tabl.Specifications
    spatType = specs.SpatialTableTypes()
    inputVarName = specs.Value("row")
    varName = inputVarName & "_" & specs.TableId()

    If Exists(varName) Then Exit Sub
    If Not LoExists("listofgeovars") Then Prepare

    AddVarNameToList varName

    Set sh = Wksh()

    With sh
        Select Case spatType
        Case "geo"
            endCounter = 4
        Case "hf"
            endCounter = 1
        Case Else
            endCounter = 4
        End Select

        For counter = 1 To endCounter
            ClearTableConditions

            If spatType = "hf" Then
                loName = CurrentLoName(varName, HF)
                adminName = HF
                rowVarName = HF & "_" & inputVarName
            Else
                loName = CurrentLoName(varName, ADMIN & counter)
                adminName = ADMIN & counter
                rowVarName = CONCAT & adminName & "_" & inputVarName
            End If

            colIndex = LastColumn()
            headerName = CurrentHeaderName(varName, adminName)

            Set cellRng = .Cells(1, colIndex)
            cellRng.Value = headerName

            'Add condition on the row variable
            AddTableConditions rowVarName, _
                "= " & cellRng.Cells(2, 1).Address(RowAbsolute:=False)

            'Column headers
            cellRng.Cells(1, 2).Value = "formula_" & adminName
            cellRng.Cells(1, 3).Value = "population_" & adminName
            cellRng.Cells(1, 4).Value = "attack_rate_" & adminName

            'Formula column value
            finalFormula = CurrentExcelFormula(formObject)
            If finalFormula <> vbNullString Then
                cellRng.Cells(2, 2).FormulaArray = "= " & finalFormula
            End If

            'Population and attack rate (geo only)
            If spatType = "geo" Then
                finalFormula = GeoPopulationFunction(adminName, _
                    cellRng.Cells(2, 1).Address(RowAbsolute:=False))
                cellRng.Cells(2, 3).Formula = "= " & finalFormula

                finalFormula = "= " & _
                    cellRng.Cells(2, 2).Address(RowAbsolute:=False) & "/" & _
                    cellRng.Cells(2, 3).Address(RowAbsolute:=False)
                cellRng.Cells(2, 4).Formula = finalFormula
            End If

            'Create the ListObject
            Set rng = .Range(cellRng.Cells(1, 1), cellRng.Cells(2, 4))
            .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loName
        Next
    End With
End Sub

'@section Validation
'===============================================================================

' @description Validate that the cross-table is not Nothing and the spatial sheet exists.
Private Sub CheckRequirements(ByVal tabl As ICrossTable)
    Dim wb As Workbook
    Dim sh As Worksheet

    If tabl Is Nothing Then
        ThrowError ObjectNotInitialized, _
            "No table provided for spatial tables creation"
    End If

    Set sh = tabl.Wksh()
    Set wb = sh.Parent

    On Error Resume Next
    Set sh = wb.Worksheets(SPATIAL_SHEET)
    On Error GoTo 0

    If sh Is Nothing Then
        ThrowError ElementNotFound, _
            "Worksheet " & SPATIAL_SHEET & " not found in workbook"
    End If
End Sub

'@section Error Handling
'===============================================================================

' @description Raise a ProjectError-based exception.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================

Private Property Get ISpatialTables_Exists(ByVal varName As String) As Boolean
    ISpatialTables_Exists = Exists(varName)
End Property

Private Sub ISpatialTables_Add(ByVal formObject As IFormulas)
    Add formObject
End Sub
