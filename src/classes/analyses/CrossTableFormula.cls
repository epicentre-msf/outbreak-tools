VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Formulas for all the cross tables in the analysis"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Formulas for all the cross tables in the analysis")
'@IgnoreModule UnrecognizedAnnotation, IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ICrossTableFormula

Private Type TCrossTableFormula
    tabl As ICrossTable
    fData As IFormulaData
    forml As IFormulas
    condVars As BetterArray
    condTests As BetterArray
End Type

Private this As TCrossTableFormula
Private Const CLASS_NAME As String = "CrossTableFormula"

'@section Factory
'===============================================================================

' @description Creates and returns a new CrossTableFormula instance bound to a specific cross-table
'              and formula-data object. CrossTableFormula is responsible for writing Excel formulas
'              into the cells of a pre-built cross-table on an analysis worksheet. It translates
'              the table's specifications (row/column variables, summary function, table type) into
'              concrete COUNTIFS/SUMPRODUCT-style array formulas that Excel can evaluate at runtime.
' @param tabl The ICrossTable object representing the already-constructed cross-table layout
'             (named ranges, headers, worksheet reference) that formulas will be written into.
' @param fData The IFormulaData object that provides dictionary-based lookups for variable names,
'              ranges, and formula templates used during formula generation.
' @return An ICrossTableFormula interface reference to the newly created instance.
' @remark Both parameters are required and must not be Nothing; an InvalidArgument error is
'         raised otherwise. Uses the PredeclaredId factory pattern (With New CrossTableFormula).
Public Function Create(ByVal tabl As ICrossTable, _
                       ByVal fData As IFormulaData) As ICrossTableFormula

    If tabl Is Nothing Then ThrowError InvalidArgument, "Cross-table not defined"
    If fData Is Nothing Then ThrowError InvalidArgument, "Formula data not defined"

    With New CrossTableFormula
        Set .Table = tabl
        Set .formData = fData
        Set Create = .Self
    End With
End Function

'@section Internal properties (factory use)
'===============================================================================

Public Property Get Self() As ICrossTableFormula
    Set Self = Me
End Property

Public Property Get Table() As ICrossTable
    Set Table = this.tabl
End Property

Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tabl = tabl
End Property

Public Property Get formData() As IFormulaData
    Set formData = this.fData
End Property

Public Property Set formData(ByVal fData As IFormulaData)
    Set this.fData = fData
End Property

'@section Shortcut properties
'===============================================================================

' @description Provides a shortcut to the ITableSpecs object associated with this cross-table.
'              ITableSpecs encapsulates all configuration metadata for a table: its row/column
'              variable names, summary function, table type, missing/total/percentage flags,
'              and the section it belongs to. This property avoids repeatedly chaining
'              Table.Specifications throughout the class methods.
' @return The ITableSpecs interface for the underlying cross-table's specification dictionary.
Private Property Get Specifications() As ITableSpecs
    Set Specifications = Table.Specifications
End Property

' @description Returns the summary function name string from the table's specifications.
'              The summary function determines the type of aggregation Excel will perform
'              in the generated formulas. Common values include "N" (count non-empty),
'              "N()" (count all), "SUM", "MEAN", "MEDIAN", and "MIN"/"MAX". This value
'              drives formula generation logic throughout the class, especially in
'              AddNonEmptyTableConditions where "N"/"N()" use a different operator pattern.
' @return A String such as "N", "N()", "SUM", "MEAN", etc., as defined in the setup dictionary.
Private Property Get SummaryFunction() As String
    SummaryFunction = Specifications.Value("summary function")
End Property

' @description Returns the unique identifier string for this cross-table. The table ID is
'              used extensively as a suffix for named ranges on the analysis worksheet
'              (e.g., "ROW_CATEGORIES_<id>", "VALUES_COL_1_<id>", "TOTAL_ROW_VALUES_<id>").
'              Every named range that belongs to this table is suffixed with this ID, making
'              it the primary key for locating table regions on the worksheet.
' @return A String uniquely identifying this table within the analysis workbook.
Private Property Get ID() As String
    ID = Specifications.TableId()
End Property

' @description Returns the section identifier for the group of tables this cross-table belongs to.
'              A section is a logical grouping of related tables on the analysis worksheet (e.g.,
'              all time-series tables sharing the same date controls). The section ID is the table
'              ID of the first (anchor) table in the section. It is used to reference shared named
'              ranges like "FIRST_VALUE_START_TIME_<sectionId>" and "TOTAL_ROW_<sectionId>" that
'              are defined once per section rather than per table.
' @return A String representing the anchor table's ID for this table's section.
Private Property Get SectionId() As String
    SectionId = Specifications.TableSectionId
End Property

' @description Returns the AnalysisTablesType byte enum value indicating what kind of cross-table
'              this is. The table type determines which branch of the AddFormulas Select Case is
'              executed: TypeGlobalSummary, TypeUnivariate, TypeBivariate, TypeTimeSeries,
'              TypeSpatioTemporal, or TypeSpatial. Each type has fundamentally different formula
'              structures, named range layouts, and condition-building logic.
' @return A Byte value corresponding to one of the AnalysisTablesType enum constants.
' @remark Returns Byte rather than the enum type for macOS VBA compatibility.
Private Property Get TableType() As Byte
    TableType = Specifications.TableType()
End Property

' @description Returns the IFormulas object for this table, creating it on first access via
'              lazy initialization. The IFormulas object (created by Formulas.Create) is the
'              core formula-parsing engine: it knows how to translate the table's summary function
'              and variable dictionary into concrete Excel formula strings for both linelist-scope
'              and analysis-scope contexts. Once created, the instance is cached in this.forml
'              so subsequent calls reuse the same object without re-parsing.
' @return An IFormulas interface that provides ParsedLinelistFormula and ParsedAnalysisFormula
'         methods for generating Excel formula text.
' @remark The IFormulas object is initialized with the table's specification dictionary,
'         the IFormulaData reference, and the summary function string. This is intentionally
'         lazy because not all code paths need the formula engine (e.g., if Valid is False).
Private Property Get TableFormula() As IFormulas
    If this.forml Is Nothing Then
        Dim specs As ITableSpecs
        Set specs = Specifications
        Set this.forml = Formulas.Create(specs.Dictionary, formData(), SummaryFunction)
    End If
    Set TableFormula = this.forml
End Property

' @description Checks whether the formula engine considers this table's formula valid for
'              analysis use. Validity is determined by the IFormulas.Valid method with the
'              "analysis" formula type. A formula may be invalid if the summary function is
'              unsupported, required variables are missing from the dictionary, or the formula
'              template cannot be parsed. When Valid returns False, the AddFormulas method
'              exits immediately without writing any formulas to the worksheet.
' @return True if the formula can be used for analysis-scope formula generation, False otherwise.
Private Property Get Valid() As Boolean
    Valid = TableFormula.Valid(formulaType:="analysis")
End Property

'@section Range and formula helpers
'===============================================================================

' @description Checks whether a named range exists on the cross-table's worksheet by attempting
'              to resolve it. This is used before writing formulas to optional table regions
'              (e.g., "MISSING_TOTAL_<id>", "MISSING_MISSING_<id>", "TOTAL_TOTAL_<id>") that
'              may or may not be present depending on the table's missing/total configuration.
'              The method uses On Error Resume Next to silently handle the case where the range
'              does not exist, then checks whether the range variable was successfully set.
' @param rngName The name of the named range to look up on the cross-table worksheet
'                (e.g., "MISSING_TOTAL_T001").
' @return True if the named range resolves to a valid Range object, False if it does not exist.
Private Function RangeExists(ByVal rngName As String) As Boolean
    Dim sh As Worksheet
    Dim rng As Range
    Set sh = Table.Wksh()
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

' @description Tests whether a generated Excel formula string is valid by attempting to write it
'              as an array formula into cell(1,1) of the cross-table's worksheet, then immediately
'              clearing that cell. If the assignment succeeds without error, the formula is
'              considered valid and FailedFormula returns False. If an error occurs (e.g., invalid
'              function name, bad range reference, or formula syntax error), execution jumps to the
'              Fail label via On Error GoTo Fail and returns True. This acts as a safety net to
'              prevent writing broken formulas into the actual table cells.
' @param frmName The Excel formula string to test (without the leading "= " prefix, which is
'                prepended internally). For example: "COUNTIFS(range1,cond1,range2,cond2)".
' @return True if the formula assignment raised an error (formula is broken), False if it succeeded.
' @remark Application.DisplayAlerts is temporarily disabled and restored afterward to suppress
'         any Excel warning dialogs during the test. Cell(1,1) is always cleared regardless of
'         success or failure to avoid leaving test artifacts on the worksheet.
Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim prevAlert As Long
    Dim sh As Worksheet

    prevAlert = Application.DisplayAlerts
    Set sh = Table.Wksh

    With sh
        Application.DisplayAlerts = False
        On Error GoTo Fail
        .Cells(1, 1).Clear
        .Cells(1, 1).FormulaArray = "= " & frmName
        .Cells(1, 1).Clear
        FailedFormula = False
        Application.DisplayAlerts = prevAlert
        Exit Function
    End With

Fail:
    FailedFormula = True
    Application.DisplayAlerts = prevAlert
End Function

' @description Returns a two-character string consisting of two double-quote characters (Chr(34) &
'              Chr(34)), which represents the Excel formula literal for an empty string ("").
'              This is used extensively when building condition expressions that test whether a
'              cell is empty, e.g., '= ""' in COUNTIFS criteria. Using Chr(34) avoids VBA
'              double-quote escaping issues and keeps the formula-building code readable.
' @return The string consisting of two Chr(34) characters, i.e., the literal text: ""
Private Function EmpStr() As String
    EmpStr = Chr(34) & Chr(34)
End Function

' @description Returns a four-character string representing the Excel formula literal "<>"
'              (Chr(34) & "<>" & Chr(34)). This is the COUNTIFS criteria for "not empty" when
'              used with the "N" or "N()" summary functions, which need a comparison operator
'              inside the criteria string. For example, a COUNTIFS condition using this would
'              look like: COUNTIFS(range, "<>") — matching any cell that is not blank.
' @return The string literal "<>" wrapped in double-quote characters for Excel formula use.
Private Function NonEmpStr() As String
    NonEmpStr = Chr(34) & "<>" & Chr(34)
End Function

' @description Builds an Excel formula string that computes a percentage (numerator / denominator)
'              with an ISERR guard to handle division-by-zero or other errors gracefully. The
'              generated formula follows the pattern: IF(ISERR(num/denom), "", num/denom). This
'              ensures that when the denominator is zero or the ranges are empty, the cell displays
'              blank instead of an #DIV/0! error. The address references use mixed absolute/relative
'              addressing so the formula can be auto-filled down rows while keeping the denominator
'              column fixed.
' @param denomRng The Range whose address will be used as the denominator in the division.
'                 Its column is always absolute; its row absoluteness is controlled by rowabs.
' @param numRng The Range whose address will be used as the numerator. Its row reference is
'               always relative (RowAbsolute:=False) so the formula adjusts when auto-filled.
' @param rowabs Optional Boolean (default True). When True, the denominator's row reference is
'               absolute (e.g., $A$5), locking it during auto-fill. When False, the row is
'               relative (e.g., $A5), allowing it to shift — used for row-based percentages
'               where each row has its own total.
' @return A String containing the complete IF(ISERR(...)) Excel formula text (without "= " prefix).
Private Function Percentage(ByVal denomRng As Range, ByVal numRng As Range, _
                            Optional ByVal rowabs As Boolean = True) As String
    Dim parsedForm As String
    parsedForm = numRng.Address(RowAbsolute:=False) & "/" & _
                 denomRng.Address(RowAbsolute:=rowabs, ColumnAbsolute:=True)
    Percentage = "IF(ISERR(" & parsedForm & ")," & EmpStr & "," & parsedForm & ")"
End Function

' @description Wraps a formula value inside an IF guard that returns an empty string when a
'              specified condition cell is blank. The generated pattern is:
'              IF(conditionCell = "", "", formulaValue). This is primarily used in time-series
'              and spatio-temporal tables where formula cells should display blank when the
'              corresponding time period label cell is empty (i.e., the row falls outside the
'              user's selected date range). Without this guard, formulas would show misleading
'              zero/error values for unused time period rows.
' @param rngCondition The Range (typically a single cell) that is tested for blankness. If this
'                     cell equals "", the formula returns "" instead of computing formulaValue.
' @param formulaValue The Excel formula string to evaluate when rngCondition is not blank.
'                     This should be a complete formula expression (without "= " prefix).
' @param rowabs Optional Boolean (default False). Controls whether the rngCondition address
'               uses an absolute row reference. False (relative) allows auto-fill to shift the
'               condition cell row-by-row.
' @return A String containing the complete IF(...="","",formulaValue) Excel formula text.
Private Function Condition(ByVal rngCondition As Range, ByVal formulaValue As String, _
                           Optional ByVal rowabs As Boolean = False) As String
    Condition = "IF(" & rngCondition.Address(RowAbsolute:=rowabs) & "=" & _
                EmpStr & "," & EmpStr & "," & formulaValue & ")"
End Function

'@section Condition management
'===============================================================================

' @description Appends a variable name and its corresponding condition test string to the
'              parallel accumulator arrays (condVars and condTests). These arrays are consumed
'              by CurrentExcelFormula to build multi-criteria COUNTIFS/SUMPRODUCT formulas.
'              Each call adds one criteria pair; for example, AddTableConditions "age_group",
'              "= $A3" would produce a COUNTIFS criterion matching the age_group column against
'              cell $A3. The BetterArray objects are lazily initialized on first use with a
'              1-based lower bound. If varName is vbNullString, nothing is pushed (silent no-op).
' @param varName The linelist variable name (e.g., "age_group", "sex", "adm1") that identifies
'                which column of the linelist data the condition applies to.
' @param conditionName The condition/criteria string for the formula (e.g., "= $A3", "= """,
'                      "<>""", ">= $B2"). This becomes the criteria argument in COUNTIFS or
'                      the comparison in SUMPRODUCT boolean arrays.
' @remark condVars and condTests always have the same length and are consumed in parallel.
'         Call ClearTableConditions to reset them before building a new formula.
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

' @description Clears conditions from the condVars/condTests accumulator arrays. When called
'              with removeNumber = 0 (the default), all accumulated conditions are cleared,
'              resetting the formula builder to a blank state. When called with a positive
'              removeNumber, only the last N condition pairs are popped off the arrays, leaving
'              earlier conditions intact. This selective removal is useful when a base set of
'              conditions is shared across multiple formulas and only the last few criteria
'              need to change between iterations (e.g., switching column labels in a bivariate loop).
' @param removeNumber Optional Long (default 0). The number of condition pairs to remove from
'                     the end of the arrays. 0 means clear everything. A positive value pops
'                     that many entries from both condVars and condTests.
' @remark If the BetterArray objects have not been initialized yet (Nothing), they are created
'         fresh and removeNumber is forced to 0 to avoid popping from an empty array.
Private Sub ClearTableConditions(Optional ByVal removeNumber As Long = 0)
    Dim counter As Long
    Dim remNum As Long

    remNum = removeNumber

    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
        remNum = 0
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
        remNum = 0
    End If

    If remNum <= 0 Then
        this.condVars.Clear
        this.condTests.Clear
    Else
        For counter = 1 To remNum
            this.condVars.Pop
            this.condTests.Pop
        Next
    End If
End Sub

' @description Adds a "non-empty" filter condition for a variable, automatically adapting the
'              condition operator to the current summary function. For count-based functions
'              ("N" and "N()"), the condition uses the equality operator with the NonEmpStr
'              literal ("= <>"), because COUNTIFS interprets "<>" as "not blank" inside a
'              criteria string. For all other summary functions (SUM, MEAN, etc.), the condition
'              uses the inequality operator with EmpStr ("<>"""), which is the standard
'              SUMPRODUCT-style not-equal-to-empty test. This distinction is critical because
'              COUNTIFS and SUMPRODUCT handle criteria strings differently.
' @param varName The linelist variable name to add the non-empty filter for (e.g., "age_group").
'                The condition will exclude rows where this variable's cell is blank.
Private Sub AddNonEmptyTableConditions(ByVal varName As String)
    If SummaryFunction = "N" Or SummaryFunction = "N()" Then
        AddTableConditions varName, "= " & NonEmpStr()
    Else
        AddTableConditions varName, "<>" & EmpStr()
    End If
End Sub

'@section Excel formula builder
'===============================================================================

' @description Builds and returns the complete Excel formula string from the currently accumulated
'              condition pairs (condVars/condTests). This is the central formula-generation entry
'              point that all table-type branches in AddFormulas call after setting up their
'              conditions. It operates in two modes depending on the scope parameter:
'              (1) Linelist scope (FormulaScopeLinelist): calls ParsedLinelistFormula on the
'              IFormulas object, which generates a formula referencing the linelist table directly
'              without any condition filtering — used for global summaries and total rows.
'              (2) Analysis scope (default, FormulaScopeAnalysis): wraps the accumulated condVars
'              and condTests into a FormulaCondition object, then calls ParsedAnalysisFormula to
'              generate a COUNTIFS/SUMPRODUCT formula with all the criteria baked in.
'              After generating the formula string, it is validated via FailedFormula. If the
'              formula fails validation (e.g., references a missing range), a fallback error
'              string "##== formula parsing failed :( ==##" is returned instead.
' @param scope Optional Byte (default 2 = FormulaScopeAnalysis). Use FormulaScopeLinelist (1)
'              for formulas that operate on the entire linelist without conditions, or
'              FormulaScopeAnalysis (2) for formulas with accumulated COUNTIFS/SUMPRODUCT criteria.
' @param prefix Optional String (default "f"). The table name prefix passed to the formula parser,
'               used to construct the linelist table reference (e.g., "f" produces "fTableName").
'               Pass vbNullString for an unprefixed reference (used in GlobalSummary's "all data").
' @return A String containing the Excel formula text (without "= " prefix) ready to be assigned
'         to a cell's FormulaArray property, or an error placeholder string if parsing failed.
' @remark If condVars or condTests is Nothing when scope is analysis, execution jumps to Fail
'         immediately. Always call AddTableConditions before requesting an analysis-scope formula.
Private Property Get CurrentExcelFormula(Optional ByVal scope As Byte = 2, _
                                         Optional ByVal prefix As String = "f") As String
    Dim cond As IFormulaCondition
    Dim tabFormObject As IFormulas
    Dim eform As String

    Set tabFormObject = TableFormula

    If scope = FormulaScopeLinelist Then
        eform = tabFormObject.ParsedLinelistFormula(useTableName:=True, tablePrefix:=prefix)
    Else
        If this.condVars Is Nothing Or this.condTests Is Nothing Then GoTo Fail
        Set cond = FormulaCondition.Create(this.condVars, this.condTests)
        eform = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:=prefix)
    End If

    If FailedFormula(eform) Then GoTo Fail

    CurrentExcelFormula = eform
    Exit Property

Fail:
    eform = Chr(34) & "##== formula parsing failed :( ==##" & Chr(34)
    CurrentExcelFormula = eform
End Property

'@section Number format helper
'===============================================================================

' @description Applies an Excel NumberFormat string to a range based on a human-readable format
'              keyword from the table specifications. This translates configuration-friendly names
'              into proper Excel custom number format strings. Supported keywords and their formats:
'              - "integer": ###0;-###0;;@ (whole numbers, no thousands separator)
'              - "round0": #,##0;-#,##0;0;@ (whole numbers with thousands separator)
'              - "percentage0": 0% (percentage with no decimals)
'              - "roundN" (e.g., "round2"): #,###0.00;-#,###0.00;0.00;@ (N decimal places)
'              - "percentageN" (e.g., "percentage1"): 0.0% (percentage with N decimals)
'              - "euros": #,##0.00 $;-#,##0.00 $;0.00 $;@ (Euro currency format)
'              - "dollars": [$$-en-US] #,##0.00 (US Dollar currency format)
'              - "text": @ (text format)
'              - Any other string: treated as a custom Excel NumberFormat applied directly.
'              If formatValue is vbNullString, the sub exits without modifying the range.
' @param formatValue The format keyword string from the table specifications (e.g., "integer",
'                    "round2", "percentage1", "euros", "dollars", "text"), or a custom Excel
'                    NumberFormat string if none of the keywords match.
' @param rng The Range to apply the number format to. All cells in this range will receive
'            the same NumberFormat.
' @remark For "roundN" and "percentageN" keywords, the numeric suffix is extracted via Replace
'         and CLng conversion. If the suffix is not a valid number, nbDecimals defaults to 1.
'         Custom format strings are applied inside On Error Resume Next to silently ignore
'         invalid format strings.
Private Sub AddRngFormat(ByVal formatValue As String, ByVal rng As Range)
    Dim formatText As String
    Dim nbDecimals As Long
    Dim zerosDecimals As String

    nbDecimals = 1

    If formatValue = vbNullString Then
        Exit Sub
    ElseIf formatValue = "integer" Then
        formatText = "###0;-###0;;@"
    ElseIf formatValue = "round0" Then
        formatText = "#,##0;-#,##0;0;@"
    ElseIf formatValue = "percentage0" Then
        formatText = "0%"
    ElseIf InStr(1, formatValue, "round") > 0 Then
        On Error Resume Next
        nbDecimals = CLng(Replace(formatValue, "round", vbNullString))
        On Error GoTo 0
        zerosDecimals = String(nbDecimals, "0")
        formatText = "#,###0." & zerosDecimals & ";-#,###0." & zerosDecimals & ";0." & zerosDecimals & ";@"
    ElseIf InStr(1, formatValue, "percentage") > 0 Then
        On Error Resume Next
        nbDecimals = CLng(Replace(formatValue, "percentage", vbNullString))
        On Error GoTo 0
        zerosDecimals = String(nbDecimals, "0")
        formatText = "0." & zerosDecimals & "%"
    ElseIf formatValue = "euros" Then
        formatText = "#,##0.00 $;-#,##0.00 $;0.00 $;@"
    ElseIf formatValue = "dollars" Then
        formatText = "[$$-en-US] #,##0.00;[$$-en-US] -#,##0.00;[$$-en-US] 0.00;@"
    ElseIf formatValue = "text" Then
        formatText = "@"
    Else
        formatText = "custom"
    End If

    If formatText <> "custom" Then
        rng.NumberFormat = formatText
    Else
        On Error Resume Next
        rng.NumberFormat = formatValue
        On Error GoTo 0
    End If
End Sub

'@section AddFormulas — main formula writing logic
'===============================================================================

' @description The main entry point that writes all Excel formulas into the cross-table's worksheet
'              cells. This is the largest and most important method in the class (~670 lines). It
'              reads the table's type from Specifications and dispatches to one of five branches,
'              each handling a fundamentally different table layout:
'
'              **TypeGlobalSummary**: The simplest case. Writes exactly two formulas into a single
'              row: column 2 gets an unfiltered linelist formula (all data, prefix=vbNullString),
'              and column 3 gets a filtered linelist formula (prefix="f"). Both use linelist scope
'              (no conditions). Applies the table's number format to both cells.
'
'              **TypeUnivariate**: Writes formulas for a one-variable frequency table. Builds a
'              condition matching the row variable against each row category label, then auto-fills
'              down the VALUES_COL_1 range. Handles missing rows (empty row variable), total rows
'              (sum of non-empty for N/N(), or full linelist formula for other functions), and
'              optional percentage columns (value / total with ISERR guard).
'
'              **TypeBivariate**: Writes formulas for a two-variable cross-tabulation grid. Loops
'              across columns from VALUES_COL_1 to the table's EndColumn, building row x column
'              condition pairs. For each column: writes the value formula + auto-fills rows, handles
'              missing-row cells, total-row cells, and percentage cells (supporting row/column/total
'              percentage modes). After the column loop, handles missing-column ranges, total-column
'              ranges, and all four corner cells (missing-total, missing-missing, total-missing,
'              total-total). Total-total for N/N() uses SUM of the total column; other functions
'              use conditional formulas adapted to the missing configuration.
'
'              **TypeTimeSeries / TypeSpatioTemporal**: The most complex branch. For new sections,
'              sets up all the time infrastructure: date label formulas (FormatDateFromLastDay),
'              start/end period columns, total period row, validation min/max dates, effective
'              start/end dates, info cells, analysis period display, and time unit validation
'              dropdown. Then loops across columns writing time-bounded conditions (>=startTime,
'              <=endTime) wrapped in Condition() guards (blank if period label is empty). Handles
'              missing rows, percentages, and spatio-temporal label formulas (INPUTSPTGEO/INPUTSPTHF
'              lookups). For SpatioTemporal tables, the column variable is prefixed with "concat_adm1_"
'              or "hf_" depending on the spatial table type.
'
'              **TypeSpatial**: Handles geographic and health-facility tables. Sets up validation
'              dropdowns (ADM_DROPDOWN, DEVIDEPOP for geo type), order numbering (1, 2, 3...),
'              row category formulas (FindTopAdmin/FindTopHF with population lookups for geo),
'              then loops across columns similar to bivariate but with an extra non-empty guard on
'              the row variable. Handles missing/total columns and corner cells. Finally creates a
'              SpatialTables management object to register the table's formula engine for spatial
'              table refresh operations.
'
'              The method exits immediately if Valid() returns False (formula cannot be parsed).
' @remark This method modifies the worksheet extensively — it writes FormulaArray values, applies
'         number formats, sets up data validation, and auto-fills ranges. It assumes all named
'         ranges have already been created by the cross-table builder before this method is called.
Private Sub AddFormulas()
    Dim tabId As String
    Dim tabType As Byte
    Dim sh As Worksheet
    Dim rng As Range
    Dim totRng As Range
    Dim rngLab As Range
    Dim misRng As Range
    Dim cellRng As Range
    Dim finalFormula As String
    Dim tabEndColumn As Long
    Dim secId As String
    Dim perRng As Range
    Dim rowVar As String
    Dim colVar As String
    Dim tabFormat As String
    Dim rowCatAddress As String
    Dim startTimeAddress As String
    Dim endTimeAddress As String
    Dim hasMissing As Boolean
    Dim hasTotal As Boolean
    Dim hasPercentage As Boolean
    Dim tabl As ICrossTable
    Dim spTab As ISpatialTables
    Dim tabFormObject As IFormulas
    Dim spTabType As String
    Dim counter As Long
    Dim specs As ITableSpecs

    If Not Valid() Then Exit Sub

    Set specs = Specifications
    Set tabl = Table()
    Set sh = tabl.Wksh()

    tabId = ID()
    secId = SectionId()
    tabType = TableType
    hasMissing = specs.HasMissing
    hasTotal = specs.HasTotal
    hasPercentage = specs.HasPercentage
    tabFormat = specs.Value("format")

    With sh

        Select Case tabType

        Case TypeGlobalSummary
            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:=vbNullString)
            Set rng = .Range("STARTROW_" & tabId)
            rng.Cells(1, 2).FormulaArray = "= " & finalFormula
            AddRngFormat tabFormat, rng.Cells(1, 2)

            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:="f")
            rng.Cells(1, 3).FormulaArray = "= " & finalFormula
            AddRngFormat tabFormat, rng.Cells(1, 3)

        Case TypeUnivariate

            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            rowVar = specs.Value("row")
            ClearTableConditions
            AddTableConditions rowVar, "=" & rng.Cells(1, 1).Address(RowAbsolute:=False)
            finalFormula = CurrentExcelFormula()

            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set cellRng = rng.Cells(1, 1)
            cellRng.FormulaArray = "= " & finalFormula
            cellRng.AutoFill Destination:=rng, Type:=xlFillValues
            AddRngFormat tabFormat, rng

            ' Missing and total
            If hasMissing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                Set rng = .Range("MISSING_ROW_VALUES_" & tabId)
                rng.Cells(1, 1).FormulaArray = "=" & finalFormula

                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & NonEmpStr()
                    finalFormula = finalFormula & " + " & CurrentExcelFormula()
                    rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                Else
                    ClearTableConditions
                    finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist)
                    rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                End If
            Else
                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)
                ClearTableConditions
                AddNonEmptyTableConditions rowVar
                finalFormula = CurrentExcelFormula()
                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
            End If

            ' Percentage
            If hasPercentage Then
                Set rng = .Range("VALUES_COL_1_" & tabId)
                Set cellRng = .Range("TOTAL_ROW_VALUES_" & tabId).Cells(1, 1)
                rng.Offset(, 1).Cells(1, 1).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), True)
                rng.Offset(, 1).Cells(1, 1).AutoFill Destination:=rng.Offset(, 1), Type:=xlFillValues
                cellRng.Cells(1, 2).Formula = "= " & Percentage(cellRng.Cells(1, 1), cellRng.Cells(1, 1))
            End If

        Case TypeBivariate

            rowVar = specs.Value("row")
            colVar = specs.Value("column")
            rowCatAddress = .Range("ROW_CATEGORIES_" & tabId).Cells(1, 1).Address(RowAbsolute:=False)
            tabEndColumn = tabl.EndColumn

            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)

            Do While rng.Column < tabEndColumn

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "=" & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                If specs.Value("missing") = "row" Or specs.Value("missing") = "all" Then
                    Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & EmpStr()
                    AddTableConditions colVar, "= " & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(misRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula

                    ClearTableConditions
                    AddTableConditions colVar, "= " & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula
                Else
                    ClearTableConditions
                    AddNonEmptyTableConditions rowVar
                    AddTableConditions colVar, "=" & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula
                End If

                ' Percentage
                If hasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill _
                        Destination:=.Range(.Cells(rng.Row, rng.Column), .Cells(totRng.Row, rng.Column)), _
                        Type:=xlFillValues
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
            Loop

            ' Missing column and total columns
            If specs.Value("missing") = "column" Or specs.Value("missing") = "all" Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                misRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                finalFormula = CurrentExcelFormula()
            Else
                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddNonEmptyTableConditions colVar
                finalFormula = CurrentExcelFormula()
            End If

            Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
            totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
            totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
            AddRngFormat tabFormat, totRng

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Total row, missing column
            If RangeExists("TOTAL_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("TOTAL_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("TOTAL_MISSING_" & tabId)
            End If

            ' Total row, total column
            If RangeExists("TOTAL_TOTAL_" & tabId) Then
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    .Range("TOTAL_TOTAL_" & tabId).Formula = _
                        "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
                Else
                    Select Case specs.Value("missing")
                    Case "all"
                        finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist)
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                    Case "row"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Case "column"
                        ClearTableConditions
                        AddNonEmptyTableConditions rowVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Case "no"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        AddNonEmptyTableConditions rowVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    End Select
                End If
                AddRngFormat tabFormat, .Range("TOTAL_TOTAL_" & tabId)
            End If

        Case TypeTimeSeries, TypeSpatioTemporal

            rowVar = specs.Value("time")
            colVar = specs.Value("column")
            counter = 1

            ' For spatio-temporal, column variable depends on geo/hf type
            If tabType = TypeSpatioTemporal Then
                spTabType = specs.SpatialTableTypes()
                colVar = Switch(spTabType = "geo", "concat_adm1_" & colVar, _
                                spTabType = "hf", "hf_" & colVar, _
                                True, vbNullString)
            End If

            ' Time period formulas (only for new sections)
            If specs.IsNewSection() Then
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FormatDateFromLastDay(TIME_UNIT_" & tabId & "," & _
                                  cellRng.Cells(1, -1).Address(RowAbsolute:=False) & "," & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & ", " & _
                                  "END_DATE_" & tabId & ")"
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' Start period column
                Set cellRng = .Range("FIRST_VALUE_START_TIME_" & tabId)
                cellRng.Formula = "= START_DATE_" & tabId
                cellRng.Cells(2, 1).Formula = "= " & _
                    cellRng.Cells(1, 2).Address(RowAbsolute:=False) & " + 1"
                Set rng = .Range("START_TIME_PERIOD_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' End period column
                Set rng = .Range("END_TIME_PERIOD_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindLastDay(TIME_UNIT_" & tabId & "," & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & ")"
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' Total period
                Set rng = .Range("TOTAL_ROW_" & tabId)
                rng.Cells(1, -1).Formula = "= START_DATE_" & tabId
                rng.Cells(1, 0).Formula = "= END_DATE_" & tabId

                ' Validation max/min
                Set rng = .Range("VALIDATION_MAX_DATE_" & tabId)
                rng.Formula = "=ValidMax(USER_START_DATE_" & tabId & _
                              ", USER_END_DATE_" & tabId & ", MIN_MIN_DATE_" & tabId & _
                              ", MAX_MAX_DATE_" & tabId & ", TIME_UNIT_" & tabId & ")"

                Set rng = .Range("VALIDATION_MIN_DATE_" & tabId)
                rng.Formula = "=ValidMin(USER_START_DATE_" & tabId & _
                              ", USER_END_DATE_" & tabId & ", MIN_MIN_DATE_" & tabId & _
                              ", MAX_MAX_DATE_" & tabId & ", TIME_UNIT_" & tabId & ")"

                ' Start/End dates
                Set rng = .Range("START_DATE_" & tabId)
                rng.Formula = "= MAX(USER_START_DATE_" & tabId & _
                              ", VALIDATION_MIN_DATE_" & tabId & ")"

                Set rng = .Range("END_DATE_" & tabId)
                rng.Formula = "= MIN(USER_END_DATE_" & tabId & _
                              ", VALIDATION_MAX_DATE_" & tabId & ")"

                ' Info cells
                Set rng = .Range("INFO_START_DATE_" & tabId)
                rng.Formula = "= InfoUser(USER_START_DATE_" & tabId & _
                              ", START_DATE_" & tabId & ")"

                Set rng = .Range("INFO_END_DATE_" & tabId)
                rng.Formula = "= InfoUser(USER_END_DATE_" & tabId & _
                              ", END_DATE_" & tabId & ", 2)"

                Set rng = .Range("INFO_ANA_PERIOD_" & tabId)
                rng.Formula = "= RNG_AnaPeriod & " & _
                              " PLAGE_VALUE(START_DATE_" & tabId & "," & _
                              "END_DATE_" & tabId & ")"

                ' Time unit validation
                On Error Resume Next
                With .Range("TIME_UNIT_" & tabId).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:="=TIME_UNIT_LIST"
                End With
                On Error GoTo 0
            End If

            ' Column formulas
            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1))
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)
            Set cellRng = .Range("FIRST_VALUE_START_TIME_" & secId)
            Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
            Set perRng = cellRng.Cells(1, 3)
            startTimeAddress = cellRng.Address(RowAbsolute:=False)
            endTimeAddress = cellRng.Cells(1, 2).Address(RowAbsolute:=False)
            tabEndColumn = tabl.EndColumn

            Do While rng.Column <= tabEndColumn

                ClearTableConditions
                AddTableConditions rowVar, ">=" & startTimeAddress
                AddTableConditions rowVar, "<=" & endTimeAddress
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()
                Set cellRng = .Cells(misRng.Row, rng.Column)
                cellRng.FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, cellRng

                ' Percentage
                If hasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                End If

                ' Spatio-temporal label formulas
                If tabType = TypeSpatioTemporal Then
                    If spTabType = "geo" Then
                        rngLab.Formula = "= IF(ISBLANK(INPUTSPTGEO_" & counter & "_" & secId & ")" & _
                                         ", " & Chr(34) & Chr(34) & ", " & _
                                         "INPUTSPTGEO_" & counter & "_" & secId & ")"
                    Else
                        rngLab.Formula = "= IF(ISBLANK(INPUTSPTHF_" & counter & "_" & secId & ")" & _
                                         ", " & Chr(34) & Chr(34) & ", " & _
                                         "INPUTSPTHF_" & counter & "_" & secId & ")"
                    End If
                    counter = counter + 1
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
            Loop

            ' Missing column and total columns
            If hasMissing Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
                Set misRng = .Range(misRng.Cells(1, 1), misRng.Cells(misRng.Rows.Count + 1, 1))

                ClearTableConditions
                AddTableConditions rowVar, ">=" & startTimeAddress
                AddTableConditions rowVar, "<=" & endTimeAddress
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()

                misRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            Else
                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    If colVar <> vbNullString Then AddNonEmptyTableConditions colVar
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            End If

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Total row, total column
            If RangeExists("TOTAL_TOTAL_" & tabId) Then
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    .Range("TOTAL_TOTAL_" & tabId).Formula = _
                        "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
                Else
                    ClearTableConditions
                    Set cellRng = .Range("TOTAL_ROW_" & secId).Cells(1, 0)
                    AddTableConditions rowVar, "<=" & cellRng.Address
                    AddTableConditions rowVar, ">=" & cellRng.Cells(1, 0).Address

                    If hasMissing Then
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Else
                        AddNonEmptyTableConditions colVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    End If
                End If
                AddRngFormat tabFormat, .Range("TOTAL_TOTAL_" & tabId)
            End If

        Case TypeSpatial

            spTabType = specs.SpatialTableTypes()
            rowVar = specs.Value("row")
            rowVar = Switch(spTabType = "geo", "concat_adm1_" & rowVar, _
                            spTabType = "hf", "hf_" & rowVar, _
                            True, vbNullString)
            colVar = specs.Value("column")

            ' Validation dropdowns
            On Error Resume Next
            If spTabType = "geo" Then
                With .Range("ADM_DROPDOWN_" & tabId).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:="=ADM_UNIT_LIST"
                End With

                With .Range("DEVIDEPOP_" & tabId).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:="=POPULATION_FACTOR_LIST"
                End With
            End If
            On Error GoTo 0

            ' Order formula
            Set rng = .Range("ROW_CATEGORIES_" & tabId).Offset(, -2)
            Set cellRng = rng.Cells(1, 1)
            cellRng.Value = 1
            Set cellRng = rng.Cells(2, 1)
            cellRng.Formula = "= " & cellRng.Offset(-1).Address(RowAbsolute:=False) & " + 1"
            Set rng = .Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
            cellRng.AutoFill Destination:=rng, Type:=xlFillValues

            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            rowCatAddress = rng.Cells(1, 1).Address(RowAbsolute:=False)

            If spTabType = "geo" Then
                ' Population formula
                Set rng = .Range("ROW_CATEGORIES_" & tabId).Offset(, -1)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindTopPop(ADM_DROPDOWN_" & tabId & "," & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & "," & _
                                  Chr(34) & rowVar & Chr(34) & "," & Chr(34) & tabId & Chr(34) & ")"
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' Top Admin formula
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindTopAdmin(ADM_DROPDOWN_" & tabId & "," & _
                                  cellRng.Cells(1, -1).Address(RowAbsolute:=False) & "," & _
                                  Chr(34) & rowVar & Chr(34) & "," & Chr(34) & tabId & Chr(34) & ")"

            ElseIf spTabType = "hf" Then
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindTopHF(" & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & "," & _
                                  Chr(34) & rowVar & Chr(34) & "," & Chr(34) & tabId & Chr(34) & ")"
            End If

            cellRng.AutoFill Destination:=rng, Type:=xlFillValues

            ' Column value formulas (similar to bivariate)
            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
            tabEndColumn = tabl.EndColumn

            Do While rng.Column <= tabEndColumn

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddNonEmptyTableConditions rowVar
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()
                Set cellRng = .Cells(misRng.Row, rng.Column)
                cellRng.FormulaArray = "= " & finalFormula

                ' Percentage (spatial only supports row percentage)
                If hasPercentage Then
                    Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                    rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Set rng = rng.Offset(, 1)
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
            Loop

            ' Missing column and total columns
            If hasMissing Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()

                misRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & rowCatAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            Else
                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & rowCatAddress
                    AddNonEmptyTableConditions colVar
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            End If

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Create spatial table management
            Set spTab = SpatialTables.Create(tabl)
            Set tabFormObject = TableFormula
            spTab.Add tabFormObject

        End Select
    End With
End Sub

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface implementation
'===============================================================================

Private Property Get ICrossTableFormula_Valid() As Boolean
    ICrossTableFormula_Valid = Valid
End Property

Private Sub ICrossTableFormula_AddFormulas()
    AddFormulas
End Sub
