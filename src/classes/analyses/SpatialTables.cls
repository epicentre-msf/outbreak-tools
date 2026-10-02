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

'@class SpatialTables
'@description
'Creates the spatial analysis ListObjects on the "__spatial_tables" worksheet
'at build time. Add takes one cross-table and builds, for each of its row
'variables, either four administrative tables (adm1 to adm4) or one health
'facility table. Entries filed along the way leave through HasCheckings and
'CheckingValues.
'
'THE CONTRACT WITH LLSPATIAL
'-------------------------------------------------------------------------------
'This class writes what LLSpatial reads in the generated linelist, and the two
'never reference each other: the contract is the set of names below, agreed by
'string convention across a workbook boundary. Change a name here and change it
'in LLSpatial in the same commit.
'
'    listofgeovars              one row per spatial variable
'    spatial_<level>_<var>      one ListObject per level
'      column 1  tabl_<level>_<var>    the lookup key
'      column 2  formula_<level>       the value LLSpatial sorts on
'      column 3  population_<level>    administrative levels only
'      column 4  attack_rate_<level>   administrative levels only
'
'Every row of listofgeovars resolves to a real ListObject, because LLSpatial
'walks that list and asks for the table by name. So a variable is registered
'after its tables are built.
'
'WHAT A FACILITY TABLE HOLDS
'-------------------------------------------------------------------------------
'A facility table is two columns: the lookup key and the value. The population
'and the attack rate belong to an administrative table alone. The geobase gives
'a population per administrative level, adm1_pop to adm4_pop, and gives none
'for a facility.
'
'LLSpatial.Update resizes a spatial table as it refreshes it, and reads the
'width off the table itself.
'@depends CrossTable, TableSpecs, Formulas, FormulaBuilder, Checking
'@version 2026-08-12



'@section State
'===============================================================================

Private Type TSpatialTables
    sh As Worksheet
    tab As CrossTable
    build As FormulaBuilder
    checks As Checking
    sealed As Boolean
End Type

Private this As TSpatialTables
Private checkCounter As Long

Private Const CLASS_NAME As String = "SpatialTables"
Private Const SPATIAL_SHEET As String = "__spatial_tables"
Private Const ADMIN As String = "adm"
Private Const HF As String = "hf"
Private Const GEO As String = "geo"
Private Const CONCAT As String = "concat_"

' The registry of spatial variables, and the column Prepare puts it in.
Private Const REGISTRY_TABLE As String = "listofgeovars"
Private Const REGISTRY_HEADER As String = "listofvars"
Private Const REGISTRY_COL As Long = 3

' Where the first spatial table starts, how wide one is, and how far past the
' last used column the next one starts. Prepare and LastColumn share these; the
' two used to hold the same numbers as literals in different methods.
Private Const FIRST_TABLE_COL As Long = 7
Private Const ADMIN_TABLE_WIDTH As Long = 4
Private Const HF_TABLE_WIDTH As Long = 2
Private Const TABLE_GAP As Long = 2

' A geographic variable is spread over these four administrative levels.
Private Const ADMIN_LEVELS As Long = 4

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a SpatialTables instance from a cross-table
'@details
'Factory method on the predeclared instance. Validates that the cross-table is
'given and that the "__spatial_tables" worksheet exists in the same workbook,
'then hands the resolved worksheet to the instance so it is looked up once.
'@param tabl CrossTable. The cross-table whose row variable defines the spatial
'tables to create.
'@return SpatialTables. A fully initialised spatial tables builder.
'@throws ProjectError.ObjectNotInitialized When tabl is Nothing.
'@throws ProjectError.ElementNotFound When the spatial worksheet is missing.
'@depends CrossTable, TableSpecs
'@export
Public Function Create(ByVal tabl As CrossTable) As SpatialTables
    Dim spatialSh As Worksheet

    Set spatialSh = CheckRequirements(tabl)

    Dim instance As SpatialTables
    Set instance = New SpatialTables
    Set instance.Table = tabl
    Set instance.SpatialSheet = spatialSh
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes.
'@export
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

'@label:table
'@prop-title Cross-table bound to this builder
'@details
'Returns the CrossTable instance whose row variable and specifications drive the
'spatial table creation in Add.
'@return CrossTable. The bound cross-table.
'@export
Public Property Get Table() As CrossTable
    Set Table = this.tab
End Property

'@label:table-set
'@prop-title Assign the cross-table reference
'@param tabl CrossTable. The cross-table to bind.
'@export
Public Property Set Table(ByVal tabl As CrossTable)
    GuardNotSealed "Table"
    Set this.tab = tabl
End Property

'@label:spatial-sheet-set
'@prop-title Assign the spatial worksheet the factory resolved
'@param sh Worksheet. The "__spatial_tables" worksheet.
'@export
Public Property Set SpatialSheet(ByVal sh As Worksheet)
    GuardNotSealed "SpatialSheet"
    Set this.sh = sh
End Property

'@section Worksheet
'===============================================================================

'@label:wksh
'@prop-title The spatial worksheet
'@details
'The factory resolves the worksheet and hands it over, so this is a read of what
'the instance already holds.
'@return Worksheet. The spatial analysis worksheet.
Private Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:id
'@prop-title The identifier of the table these spatial tables belong to
'@return String. The table identifier.
Private Property Get ID() As String
    ID = Table.Specifications.TableId()
End Property

'@section Naming Conventions
'===============================================================================
'@description Helpers for building the spatial table and header names. LLSpatial
'builds the same two strings from the same pieces.

'@label:current-header-name
'@prop-title Build the header name of a spatial table
'@details
'The shape is "tabl_<level>_<variable>", for instance "tabl_adm1_cases_sp1". It
'is written into the first cell of the table and is the lookup key LLSpatial
'reads the unit names under.
'@param varName String. The spatial variable name.
'@param adminName String. The level identifier, such as "adm1" or "hf".
'@return String. The header name.
Private Property Get CurrentHeaderName(ByVal varName As String, _
                                       ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

'@label:current-lo-name
'@prop-title Build the ListObject name of a spatial table
'@details
'The shape is "spatial_<level>_<variable>", for instance "spatial_adm1_cases_sp1".
'@param varName String. The spatial variable name.
'@param adminName String. The level identifier.
'@return String. The ListObject name.
Private Property Get CurrentLoName(ByVal varName As String, _
                                   ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'@label:level-count
'@prop-title How many tables one spatial variable takes
'@details
'A facility variable takes one table and a geographic one takes four. An empty
'spatial type answers four, which is what a table whose spatial type cell was
'never filled produces.
'@param spatType String. The spatial type, "geo" or "hf".
'@return Long. The number of tables to build.
Private Property Get LevelCount(ByVal spatType As String) As Long
    LevelCount = IIf(spatType = HF, 1, ADMIN_LEVELS)
End Property

'@label:table-width
'@prop-title How wide one table of a spatial variable is
'@details
'An administrative table carries the value, the population and the attack rate.
'A facility table carries the value alone, because a health facility has no
'population to divide by.
'@param spatType String. The spatial type, "geo" or "hf".
'@return Long. The number of columns.
Private Property Get TableWidth(ByVal spatType As String) As Long
    TableWidth = IIf(spatType = HF, HF_TABLE_WIDTH, ADMIN_TABLE_WIDTH)
End Property

'@label:level-name
'@prop-title The level identifier of one table of a spatial variable
'@param spatType String. The spatial type, "geo" or "hf".
'@param counter Long. The 1-based table number.
'@return String. "hf", or "adm" followed by the level number.
Private Property Get LevelName(ByVal spatType As String, ByVal counter As Long) As String
    LevelName = IIf(spatType = HF, HF, ADMIN & counter)
End Property

'@section Existence Checks
'===============================================================================

'@label:exists
'@sub-title Whether the spatial tables of a variable are already built
'@details
'Asks for the tables the spatial type actually produces. It used to probe the
'four administrative levels whatever the type, so a facility variable answered
'False every time: the guard never fired, the variable was registered twice and
'ListObjects.Add raised 1004 on the second pass over the same workbook.
'@param varName String. The spatial variable name.
'@param spatType String. The spatial type, "geo" or "hf".
'@return Boolean. True when every table the type needs is present.
Private Function Exists(ByVal varName As String, ByVal spatType As String) As Boolean
    Dim counter As Long

    If spatType = HF Then
        Exists = LoExists(CurrentLoName(varName, HF))
        Exit Function
    End If

    For counter = 1 To ADMIN_LEVELS
        If Not LoExists(CurrentLoName(varName, ADMIN & counter)) Then Exit Function
    Next

    Exists = True
End Function

'@label:lo-exists
'@sub-title Whether a ListObject is on the spatial sheet
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
'@prop-title The column the next spatial table starts in
'@details
'Reads the last used column of row 1 and leaves a gap after it. A sheet carrying
'only the registry answers with the first table column, so the columns between
'the registry and the first table stay empty.
'@return Long. The 1-based column index.
Private Property Get LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long

    Set sh = Wksh()
    colIndex = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column
    LastColumn = IIf(colIndex = REGISTRY_COL, FIRST_TABLE_COL, colIndex + TABLE_GAP)
End Property

'@section Formula Helpers
'===============================================================================

'@label:geo-population-function
'@sub-title Build the population lookup of one administrative level
'@details
'An INDEX and MATCH over the geobase table of the level, wrapped in ISERROR so a
'unit the geobase does not carry reads 0. The population column index comes from
'the level number: T_ADM1 holds it in column 2, T_ADM2 in column 3, and so on.
'@param adminName String. The level identifier, such as "adm1".
'@param rangeAddress String. The cell address MATCH looks the unit up by.
'@return String. The formula text with no leading "=", or an empty string when
'the name carries no level number.
'@remarks A NAME THAT IS NOT adm<n> USED TO GIVE COLUMN ZERO. The level number
'was read with CLng under On Error Resume Next, so anything else left the index
'at 0 -- and INDEX(table, row, 0) is a whole-row reference rather than an error,
'so ISERROR answered False and the cell took the first column of the row. The
'level number is checked before it is used now, and a name without one files a
'warning and leaves the population cell empty.
Private Function GeoPopulationFunction(ByVal adminName As String, _
                                       ByVal rangeAddress As String) As String
    Dim tabName As String
    Dim concatName As String
    Dim centralPart As String
    Dim levelText As String
    Dim popColIndex As Long

    levelText = Replace(LCase(adminName), ADMIN, vbNullString)

    If Not IsNumeric(levelText) Then
        LogInfo "The population column of " & adminName & _
                " could not be worked out: the name carries no level number", _
                checkingWarning
        Exit Function
    End If

    popColIndex = CLng(levelText) + 1
    tabName = "T_" & UCase(adminName)
    concatName = LCase(adminName) & "_concat"

    centralPart = "INDEX(" & tabName & ", MATCH(" & rangeAddress & _
                  ", " & concatName & ", 0), " & popColIndex & ")"
    GeoPopulationFunction = "IF(ISERROR(" & centralPart & "), 0," & centralPart & ")"
End Function

'@label:builder
'@prop-title The formula builder this class writes through
'@details
'Holds the criteria of the formula being built, asks the parser for the text and
'writes it into the cell. CrossTableFormula writes through one of these too, so
'both agree on how a formula is made and on what a refusal looks like.
'@return FormulaBuilder. The builder, created on first use.
Private Property Get Builder() As FormulaBuilder
    Dim entries As Checking

    If this.build Is Nothing Then
        Set entries = Checks
        Set this.build = FormulaBuilder.Create(CLASS_NAME, ID(), entries)
    End If

    Set Builder = this.build
End Property

'@section Preparation
'===============================================================================

'@label:prepare
'@sub-title Put the registry on the spatial sheet
'@details
'Creates the listofgeovars ListObject LLSpatial walks. Add calls this the first
'time it finds no registry.
Private Sub Prepare()
    Dim rng As Range
    Dim sh As Worksheet

    Set sh = Wksh()

    With sh
        Set rng = .Range(.Cells(1, REGISTRY_COL), .Cells(2, REGISTRY_COL))
        rng.Cells(1, 1).Value = REGISTRY_HEADER
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = REGISTRY_TABLE
    End With
End Sub

'@label:add-var-name-to-list
'@sub-title Register a variable in the registry
'@details
'Appends one row through the ListObject itself, which puts the value in the row
'the table ends at and resizes the table with it. The row used to be computed
'from CountA over the whole table range. CountA answers how many cells hold
'something, and the write needed a row number, so one blank row anywhere in the
'registry made the next variable overwrite an existing entry. LLSpatial walks that list to decide
'which variables get refreshed, so a lost entry is a spatial table that never
'updates in the delivered linelist.
'@param varName String. The spatial variable name to register.
Private Sub AddVarNameToList(ByVal varName As String)
    Dim Lo As ListObject
    Dim newRow As ListRow

    Set Lo = Wksh().ListObjects(REGISTRY_TABLE)

    If Lo.ListRows.Count = 0 Then
        Set newRow = Lo.ListRows.Add
    ElseIf Lo.ListRows(Lo.ListRows.Count).Range.Cells(1, 1).Value = vbNullString Then
        Set newRow = Lo.ListRows(Lo.ListRows.Count)
    Else
        Set newRow = Lo.ListRows.Add
    End If

    newRow.Range.Cells(1, 1).Value = varName
End Sub

'@section Add
'===============================================================================

'@label:add
'@sub-title Build the spatial tables of the cross-table row variable
'@details
'A geographic variable gets four tables, one per administrative level, each with
'the value formula, the population lookup and the attack rate. A facility
'variable gets one table with the value formula. Every table is two rows and four
'columns, and they are laid out left to right with a gap between them.
'
'The variable is registered in listofgeovars after its tables are built. It used
'to be registered first, so a failure inside the loop left a name in the registry
'with no table behind it and LLSpatial walked to a table that was not there.
'@param formObject Formulas. The summary function the value column is built from.
'@export
Public Sub Add(ByVal formObject As Formulas)
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim varName As String
    Dim headerName As String
    Dim loName As String
    Dim rng As Range
    Dim counter As Long
    Dim specs As TableSpecs
    Dim tabl As CrossTable
    Dim cellRng As Range
    Dim finalFormula As String
    Dim adminName As String
    Dim rowVarName As String
    Dim inputVarName As String
    Dim endCounter As Long
    Dim spatType As String
    Dim formulaIsValid As Boolean

    Set tabl = Table
    Set specs = tabl.Specifications
    spatType = specs.SpatialTableScopes()
    inputVarName = specs.Value("row")
    varName = inputVarName & "_" & specs.TableId()

    If Exists(varName, spatType) Then Exit Sub
    If Not LoExists(REGISTRY_TABLE) Then Prepare

    If Not formObject Is Nothing Then formulaIsValid = formObject.Valid()

    If Not formulaIsValid Then _
        LogInfo "The summary function of table " & ID() & _
                " could not be read, so its spatial tables carry no value formula", _
                checkingError

    Set sh = Wksh()
    endCounter = LevelCount(spatType)

    ' The answer is the same scan for every level of one variable, so it is read
    ' once and advanced by the width of a table plus the gap.
    colIndex = LastColumn()

    With sh
        For counter = 1 To endCounter
            Builder.ClearConditions

            adminName = LevelName(spatType, counter)
            loName = CurrentLoName(varName, adminName)

            If spatType = HF Then
                rowVarName = HF & "_" & inputVarName
            Else
                rowVarName = CONCAT & adminName & "_" & inputVarName
            End If

            headerName = CurrentHeaderName(varName, adminName)

            Set cellRng = .Cells(1, colIndex)
            cellRng.Value = headerName

            'The value of the table is read for the unit named one row below
            Builder.AddCondition rowVarName, _
                "= " & cellRng.Cells(2, 1).Address(RowAbsolute:=False)

            'Column headers. The population and the attack rate belong to an
            'administrative table alone.
            cellRng.Cells(1, 2).Value = "formula_" & adminName

            If spatType <> HF Then
                cellRng.Cells(1, 3).Value = "population_" & adminName
                cellRng.Cells(1, 4).Value = "attack_rate_" & adminName
            End If

            'Formula column value
            finalFormula = vbNullString
            If formulaIsValid Then finalFormula = Builder.ExcelFormula(formObject)
            Builder.WriteFormula cellRng.Cells(2, 2), finalFormula

            'Population and attack rate, which the geobase carries per
            'administrative level. A spatial type the setup workbook never filled
            'reaches here too, and it gets the four administrative tables with no
            'population, the way it always has.
            If spatType = GEO Then
                Builder.WriteFormula cellRng.Cells(2, 3), _
                    GeoPopulationFunction(adminName, _
                        cellRng.Cells(2, 1).Address(RowAbsolute:=False)), False

                'The population lookup answers 0 for a unit the geobase does
                'not carry, and a bare division by it left #DIV/0! in the
                'attack rate. An error value outranks every number when the
                'table is ordered on this column, so such a unit took rank 1.
                'The 0 puts it last instead.
                Builder.WriteFormula cellRng.Cells(2, 4), _
                    "IFERROR(" & _
                    cellRng.Cells(2, 2).Address(RowAbsolute:=False) & "/" & _
                    cellRng.Cells(2, 3).Address(RowAbsolute:=False) & ", 0)", False
            End If

            'Create the ListObject
            Set rng = .Range(cellRng.Cells(1, 1), _
                             cellRng.Cells(2, TableWidth(spatType)))
            .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loName

            colIndex = colIndex + TableWidth(spatType) - 1 + TABLE_GAP
        Next
    End With

    AddVarNameToList varName
End Sub

'@section Checkings
'===============================================================================
'@description What the class reports back to the generation report.
'CrossTableFormula harvests these after Add, and AnalysisOutput harvests them
'from there. A blank formula column used to reach the delivered table with
'nothing said about it, and the attack rate two cells over divided by it.

'@label:checks
'@prop-title The report entries of this instance
'@return Checking. The entries, created on first use.
Private Property Get Checks() As Checking
    If this.checks Is Nothing Then Set this.checks = Checking.Create(CLASS_NAME)
    Set Checks = this.checks
End Property

'@label:log-info
'@sub-title File one entry
'@details
'Checking.Add raises on a duplicate key, and one report carries the entries of
'several classes over several tables, so the key names the class, the table and a
'counter.
'@param message String. The message text.
'@param scope Optional Byte. The checking scope. Defaults to checkingSuccess.
Private Sub LogInfo(ByVal message As String, _
                    Optional ByVal scope As Byte = checkingSuccess)
    Dim entries As Checking

    Set entries = Checks
    checkCounter = checkCounter + 1
    entries.Add CLASS_NAME & "-" & ID() & "-" & CStr(checkCounter), message, scope
End Sub

'@label:has-checkings
'@prop-title Whether this instance has anything to report
'@return Boolean. True once one entry has been filed.
'@export
Public Property Get HasCheckings() As Boolean
    If this.checks Is Nothing Then Exit Property
    HasCheckings = (this.checks.Length() > 0)
End Property

'@label:checking-values
'@prop-title The entries filed while the spatial tables were built
'@return Checking. The entries, or Nothing when none were filed.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = this.checks
End Property

'@section Validation
'===============================================================================

'@label:check-requirements
'@sub-title Validate the cross-table and find the spatial worksheet
'@details
'The lookup used to overwrite the variable already holding the analysis
'worksheet, and a failed Set leaves the target at what it held, so the guard
'below could never fire and a missing worksheet surfaced later as a bare
'subscript error with nothing pointing at the sheet.
'@param tabl CrossTable. The cross-table to validate.
'@return Worksheet. The spatial worksheet of the same workbook.
'@throws ProjectError.ObjectNotInitialized When tabl is Nothing.
'@throws ProjectError.ElementNotFound When the spatial worksheet is missing.
Private Function CheckRequirements(ByVal tabl As CrossTable) As Worksheet
    Dim wb As Workbook
    Dim spatialSh As Worksheet

    If tabl Is Nothing Then
        ThrowError ObjectNotInitialized, _
            "No table provided for spatial tables creation"
    End If

    Set wb = tabl.Wksh().Parent

    On Error Resume Next
    Set spatialSh = wb.Worksheets(SPATIAL_SHEET)
    On Error GoTo 0

    If spatialSh Is Nothing Then
        ThrowError ElementNotFound, _
            "Worksheet " & SPATIAL_SHEET & " not found in workbook"
    End If

    Set CheckRequirements = spatialSh
End Function

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
