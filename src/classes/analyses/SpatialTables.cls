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

'@class SpatialTables
'@description
'Creates spatial analysis ListObjects on the "spatial_tables__" worksheet at
'linelist build time. For each cross-table row variable, builds either four
'admin-level tables (adm1-adm4 for geo type) or one health-facility table
'(for hf type), each with formula, population, and attack rate columns.
'Manages formula conditions and validates that the spatial sheet exists
'before writing.
'@depends ICrossTable, ITableSpecs, IFormulas, IFormulaCondition, BetterArray
'@version 2025-01-15
'@author EpiData


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

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a SpatialTables instance from a cross-table
'@details
'Factory method on the predeclared instance. Validates that the cross-table
'is not Nothing and that the "spatial_tables__" worksheet exists in the same
'workbook. Binds the cross-table to the new instance so subsequent Add calls
'can read the row variable and table specifications.
'@param tabl ICrossTable. The cross-table whose row variable defines the
'spatial tables to create.
'@return ISpatialTables. A fully initialised spatial tables builder.
'@throws ProjectError.ObjectNotInitialized When tabl is Nothing.
'@throws ProjectError.ElementNotFound When the spatial worksheet is missing.
'@depends ICrossTable, ITableSpecs
'@export
Public Function Create(ByVal tabl As ICrossTable) As ISpatialTables
    CheckRequirements tabl

    With New SpatialTables
        Set .Table = tabl
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor for fluent retrieval of the interface reference from
'the predeclared Create method.
'@return ISpatialTables. The current instance cast to the interface.
Public Property Get Self() As ISpatialTables
    Set Self = Me
End Property

'@label:table
'@prop-title Cross-table bound to this builder
'@details
'Returns the ICrossTable instance whose row variable and specifications
'drive the spatial table creation in Add.
'@return ICrossTable. The bound cross-table.
Public Property Get Table() As ICrossTable
    Set Table = this.tab
End Property

'@label:table-set
'@prop-title Assign the cross-table reference
'@param tabl ICrossTable. The cross-table to bind.
Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tab = tabl
End Property

'@section Worksheet
'===============================================================================

'@label:wksh
'@prop-title Lazy-loaded spatial worksheet reference
'@details
'Returns the "spatial_tables__" worksheet from the cross-table workbook.
'The reference is cached after the first access to avoid repeated lookups.
'@return Worksheet. The spatial analysis worksheet.
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
'@description Helpers for constructing spatial table and range names.

'@label:current-header-name
'@prop-title Build the header named range name for a spatial table
'@details
'Concatenates "tabl_", the admin name, and the variable name to form a
'unique header range name (e.g. "tabl_adm1_cases_sp1").
'@param varName String. The spatial variable name.
'@param adminName String. The admin level identifier (e.g. "adm1", "hf").
'@return String. The header named range name.
Private Property Get CurrentHeaderName(ByVal varName As String, _
                                       ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

'@label:current-lo-name
'@prop-title Build the ListObject name for a spatial table
'@details
'Concatenates "spatial_", the admin name, and the variable name to form a
'unique ListObject name (e.g. "spatial_adm1_cases_sp1").
'@param varName String. The spatial variable name.
'@param adminName String. The admin level identifier.
'@return String. The ListObject name.
Private Property Get CurrentLoName(ByVal varName As String, _
                                   ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'@section Existence Checks
'===============================================================================

'@label:exists
'@sub-title Check if spatial ListObjects already exist for a variable
'@details
'Iterates through admin levels 1-4 and tests whether each corresponding
'ListObject exists on the spatial worksheet. Returns True only when all
'four exist. Used by Add to skip duplicate creation.
'@param varName String. The spatial variable name to check.
'@return Boolean. True when all four admin-level ListObjects exist.
Private Function Exists(ByVal varName As String) As Boolean
    Dim counter As Long
    Dim testExistence As Boolean

    For counter = 1 To 4
        testExistence = LoExists(CurrentLoName(varName, ADMIN & counter))
        If Not testExistence Then Exit For
    Next

    Exists = testExistence
End Function

'@label:lo-exists
'@sub-title Check if a ListObject exists on the spatial sheet
'@details
'Uses the standard VBA error-trapping pattern to test for ListObject
'existence on the spatial worksheet.
'@param loName String. The ListObject name to look for.
'@return Boolean. True when the ListObject exists.
Private Function LoExists(ByVal loName As String) As Boolean
    Dim Lo As ListObject

    On Error Resume Next
    Set Lo = Wksh().ListObjects(loName)
    On Error GoTo 0

    LoExists = Not (Lo Is Nothing)
End Function

'@section Column Helpers
'===============================================================================

'@label:last-column
'@prop-title Next available column on the spatial sheet
'@details
'Finds the last used column on row 1 and returns the next available
'position with a 2-column gap. Returns column 7 when the sheet only has
'the initial 3-column setup (listofgeovars occupies column 3).
'@return Long. The 1-based column index for the next spatial table.
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
'@description Helpers for building and testing Excel formula expressions.

'@label:geo-population-function
'@sub-title Build a geo population INDEX/MATCH formula for an admin level
'@details
'Constructs an INDEX/MATCH formula that looks up the population value for
'an admin level from the corresponding geo reference table (T_ADM1 through
'T_ADM4). The formula is wrapped in ISERROR to return 0 when the lookup
'fails. The population column index is derived from the admin level number.
'@param adminName String. The admin level identifier (e.g. "adm1").
'@param rangeAddress String. The cell address used as the MATCH lookup value.
'@return String. The complete INDEX/MATCH formula string.
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

'@label:failed-formula
'@sub-title Test whether a formula expression is valid
'@details
'Writes the formula to cell A1 of the spatial sheet inside an error
'handler. If the assignment succeeds, the formula is valid and the cell
'is cleared. If it fails, the method returns True. DisplayAlerts is
'temporarily disabled to suppress Excel prompts.
'@param frmName String. The formula expression to test (without leading "=").
'@return Boolean. True when the formula is invalid.
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

'@label:emp-str
'@sub-title Empty string constant for formula conditions
'@details
'Returns a pair of double-quote characters for use in Excel formula strings.
'@return String. Two double-quote characters (empty string literal in Excel).
Private Function EmpStr() As String
    EmpStr = Chr(34) & Chr(34)
End Function

'@section Condition Management
'===============================================================================
'@description Manages variable/condition pairs for formula generation.

'@label:add-table-conditions
'@sub-title Add a variable/condition pair for formula generation
'@details
'Pushes the variable name and condition test onto the parallel condVars
'and condTests BetterArrays. Lazily initialises both arrays with
'LowerBound 1 on first use. Ignores empty variable names.
'@param varName String. The variable name for the condition.
'@param conditionName String. The condition expression.
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

'@label:clear-table-conditions
'@sub-title Clear condition arrays
'@details
'When removeNumber is zero or negative, clears both arrays entirely.
'When positive, removes only the last N entries from both arrays via Pop.
'Lazily initialises the arrays if they are Nothing.
'@param removeNumber Optional Long. Number of trailing entries to remove.
'Defaults to 0 (clear all).
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

'@label:current-excel-formula
'@prop-title Build an Excel formula string using the current conditions
'@details
'Constructs a parsed formula from the IFormulas object, applying either
'linelist scope (with table name prefix) or analysis scope (with condition
'pairs from condVars/condTests). Validates the result via FailedFormula
'and returns an empty string when the formula is invalid or the object
'is not valid.
'@param tabFormObject IFormulas. The formula definition to parse.
'@param scope Optional Byte. FormulaScopeLinelist or FormulaScopeAnalysis.
'Defaults to FormulaScopeAnalysis.
'@param prefix Optional String. Table name prefix for linelist scope.
'Defaults to "f".
'@return String. The parsed formula string, or empty when invalid.
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

'@section Preparation
'===============================================================================

'@label:prepare
'@sub-title Initialise the spatial sheet with the listofgeovars ListObject
'@details
'Creates the "listofgeovars" ListObject in column 3 and defines the
'RNG_PastingCol and RNG_TestingFormula named ranges. Called by Add when
'the listofgeovars table does not yet exist.
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

'@label:add-var-name-to-list
'@sub-title Register a variable name in the listofgeovars ListObject
'@details
'Appends the variable name to the next empty row in the listofgeovars
'ListObject. Uses CountA to determine the current row count.
'@param varName String. The spatial variable name to register.
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

'@label:add
'@sub-title Create spatial ListObjects for the cross-table row variable
'@details
'Builds the spatial table set for the current cross-table. For geo-type
'tables, creates four admin-level tables (adm1-adm4) each with formula,
'population, and attack rate columns. For hf-type tables, creates a single
'health-facility table with a formula column only. Each table is created
'as a 2-row ListObject (header + one data row) positioned at the next
'available column on the spatial sheet. The formula column is populated
'from the parsed analysis formula via CurrentExcelFormula. Skips creation
'when the tables already exist for this variable.
'@param formObject IFormulas. The formula object whose parsed expression
'is written into each spatial table formula column.
'@export
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
    spatType = specs.SpatialTableScopes()
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

'@label:check-requirements
'@sub-title Validate cross-table and spatial sheet existence
'@details
'Guards the Create factory by ensuring the cross-table is not Nothing and
'that a worksheet named "spatial_tables__" exists in the same workbook.
'@param tabl ICrossTable. The cross-table to validate.
'@throws ProjectError.ObjectNotInitialized When tabl is Nothing.
'@throws ProjectError.ElementNotFound When the spatial worksheet is missing.
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

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME.
'@param errNumber Long. The ProjectError code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError. Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ISpatialTables contract.
'See the corresponding Private members above for full documentation.

Private Property Get ISpatialTables_Exists(ByVal varName As String) As Boolean
    ISpatialTables_Exists = Exists(varName)
End Property

Private Sub ISpatialTables_Add(ByVal formObject As IFormulas)
    Add formObject
End Sub
