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

'@class CrossTableFormula
'@description
'Translates cross-table specifications into concrete Excel formulas and writes
'them into the cells of analysis worksheets. Each instance is bound to one
'ICrossTable and its companion IFormulaData, and the central AddFormulas method
'dispatches to type-specific branches (GlobalSummary, Univariate, Bivariate,
'TimeSeries, SpatioTemporal, Spatial) to produce COUNTIFS / SUMPRODUCT-style
'array formulas, percentage columns, number formatting, time-period scaffolding,
'spatial lookups, and data-validation dropdowns. Consumers obtain an instance
'through the predeclared Create factory and interact via the ICrossTableFormula
'interface.
'@depends ICrossTable, IFormulaData, IFormulas, Formulas, IFormulaCondition, FormulaCondition, ITableSpecs, ISpatialTables, SpatialTables, BetterArray
'@version 1.0 (2026-02-10)


Private Type TCrossTableFormula
    tabl As ICrossTable
    fData As IFormulaData
    forml As IFormulas
    condVars As BetterArray
    condTests As BetterArray
End Type

Private this As TCrossTableFormula
Private Const CLASS_NAME As String = "CrossTableFormula"

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a CrossTableFormula instance bound to a cross-table and formula data source.
'@details
'Factory method that returns a new CrossTableFormula through the predeclared-id
'pattern. Both parameters are validated for Nothing before the instance is
'constructed. The returned object exposes the ICrossTableFormula interface so
'callers program against the contract rather than the concrete class.
'@param tabl ICrossTable. The already-constructed cross-table layout providing named ranges, headers, and worksheet reference.
'@param fData IFormulaData. The dictionary-based lookup object supplying variable names, ranges, and formula templates.
'@return ICrossTableFormula. A fully initialised instance ready to write formulas.
'@throws InvalidArgument When tabl is Nothing.
'@throws InvalidArgument When fData is Nothing.
'@export
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

'@section Internal Properties
'===============================================================================
'@description Properties used during factory construction and internal wiring.

'@label:self
'@prop-title Current object instance cast to the interface.
'@details
'Convenience accessor so the Create factory can return the newly constructed
'object through the ICrossTableFormula interface.
'@return ICrossTableFormula. The current instance.
Public Property Get Self() As ICrossTableFormula
    Set Self = Me
End Property

'@label:table-get
'@prop-title ICrossTable backing store.
'@details
'Returns the cross-table object that this formula writer targets. The
'cross-table provides the worksheet reference, named ranges, and layout
'metadata needed to place formulas.
'@return ICrossTable. The bound cross-table instance.
Public Property Get Table() As ICrossTable
    Set Table = this.tabl
End Property

'@label:table-set
'@prop-title Assign the ICrossTable backing store.
'@param tabl ICrossTable. The cross-table to bind to this formula writer.
Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tabl = tabl
End Property

'@label:form-data-get
'@prop-title IFormulaData backing store.
'@details
'Returns the formula-data object that supplies dictionary-based lookups for
'variable names, linelist ranges, and formula templates during generation.
'@return IFormulaData. The bound formula-data instance.
Public Property Get formData() As IFormulaData
    Set formData = this.fData
End Property

'@label:form-data-set
'@prop-title Assign the IFormulaData backing store.
'@param fData IFormulaData. The formula-data source to bind.
Public Property Set formData(ByVal fData As IFormulaData)
    Set this.fData = fData
End Property

'@section Shortcut Properties
'===============================================================================
'@description Convenience accessors that delegate to the cross-table or its
'specifications, reducing repetitive chaining throughout the class.

'@label:specifications
'@prop-title ITableSpecs shortcut for the cross-table specifications.
'@details
'Provides a shortcut to the ITableSpecs object associated with this cross-table.
'ITableSpecs encapsulates all configuration metadata for a table: its row/column
'variable names, summary function, table scope, missing/total/percentage flags,
'and the section it belongs to. This property avoids repeatedly chaining
'Table.Specifications throughout the class methods.
'@return ITableSpecs. The specification dictionary for the underlying cross-table.
Private Property Get Specifications() As ITableSpecs
    Set Specifications = Table.Specifications
End Property

'@label:summary-function
'@prop-title Summary function name from the table specifications.
'@details
'Returns the summary function name string from the table specifications.
'The summary function determines the type of aggregation Excel will perform
'in the generated formulas. Common values include "N" (count non-empty),
'"N()" (count all), "SUM", "MEAN", "MEDIAN", and "MIN"/"MAX". This value
'drives formula generation logic throughout the class, especially in
'AddNonEmptyTableConditions where "N"/"N()" use a different operator pattern.
'@return String. A value such as "N", "N()", "SUM", "MEAN", etc.
Private Property Get SummaryFunction() As String
    SummaryFunction = Specifications.Value("summary function")
End Property

'@label:id
'@prop-title Unique identifier string for this cross-table.
'@details
'Returns the unique identifier string for this cross-table. The table ID is
'used extensively as a suffix for named ranges on the analysis worksheet
'(e.g., "ROW_CATEGORIES_<id>", "VALUES_COL_1_<id>", "TOTAL_ROW_VALUES_<id>").
'Every named range that belongs to this table is suffixed with this ID, making
'it the primary key for locating table regions on the worksheet.
'@return String. The table identifier unique within the analysis workbook.
Private Property Get ID() As String
    ID = Specifications.TableId()
End Property

'@label:section-id
'@prop-title Section identifier for the group this table belongs to.
'@details
'Returns the section identifier for the group of tables this cross-table belongs
'to. A section is a logical grouping of related tables on the analysis worksheet
'(e.g., all time-series tables sharing the same date controls). The section ID is
'the table ID of the first (anchor) table in the section. It is used to reference
'shared named ranges like "FIRST_VALUE_START_TIME_<sectionId>" and
'"TOTAL_ROW_<sectionId>" that are defined once per section rather than per table.
'@return String. The anchor table ID for this table section.
Private Property Get SectionId() As String
    SectionId = Specifications.TableSectionId
End Property

'@label:table-type
'@prop-title AnalysisTableScope byte value indicating the cross-table kind.
'@details
'Returns the AnalysisTableScope byte enum value indicating what kind of
'cross-table this is. The table scope determines which branch of the AddFormulas
'Select Case is executed: ScopeGlobalSummary, ScopeUnivariate, ScopeBivariate,
'ScopeTimeSeries, ScopeSpatioTemporal, or ScopeSpatial. Each type has fundamentally
'different formula structures, named range layouts, and condition-building logic.
'@return Byte. One of the AnalysisTableScope enum constants.
'@remarks Returns Byte rather than the enum type for macOS VBA compatibility.
Private Property Get TableScope() As Byte
    TableScope = Specifications.TableScope()
End Property

'@label:table-formula
'@prop-title Lazy-loaded IFormulas parsing engine for this table.
'@details
'Returns the IFormulas object for this table, creating it on first access via
'lazy initialisation. The IFormulas object (created by Formulas.Create) is the
'core formula-parsing engine: it knows how to translate the table summary function
'and variable dictionary into concrete Excel formula strings for both linelist-scope
'and analysis-scope contexts. Once created, the instance is cached in this.forml
'so subsequent calls reuse the same object without re-parsing.
'@return IFormulas. The formula parsing engine providing ParsedLinelistFormula and ParsedAnalysisFormula methods.
'@remarks The IFormulas object is initialised with the table specification dictionary,
'the IFormulaData reference, and the summary function string. This is intentionally
'lazy because not all code paths need the formula engine (e.g., if Valid is False).
Private Property Get TableFormula() As IFormulas
    If this.forml Is Nothing Then
        Dim specs As ITableSpecs
        Set specs = Specifications
        Set this.forml = Formulas.Create(specs.Dictionary, formData(), SummaryFunction)
    End If
    Set TableFormula = this.forml
End Property

'@label:valid
'@prop-title Whether the formula object is valid for the analysis context.
'@details
'Checks whether the formula engine considers this table formula valid for
'analysis use. Validity is determined by the IFormulas.Valid method with the
'"analysis" formula type. A formula may be invalid if the summary function is
'unsupported, required variables are missing from the dictionary, or the formula
'template cannot be parsed. When Valid returns False, the AddFormulas method
'exits immediately without writing any formulas to the worksheet.
'@return Boolean. True if the formula can be used for analysis-scope generation.
'@export
Private Property Get Valid() As Boolean
    Valid = TableFormula.Valid(formulaType:="analysis")
End Property

'@section Range and Formula Helpers
'===============================================================================
'@description Utility functions for range existence checks, formula validation,
'string literals, percentage formulas, and conditional wrapping.

'@label:range-exists
'@sub-title Check whether a named range exists on the cross-table worksheet.
'@details
'Attempts to resolve a named range on the cross-table worksheet using On Error
'Resume Next to silently handle missing ranges. This is used before writing
'formulas to optional table regions (e.g., "MISSING_TOTAL_<id>",
'"MISSING_MISSING_<id>", "TOTAL_TOTAL_<id>") that may or may not be present
'depending on the table missing/total configuration.
'@param rngName String. The named range to look up (e.g., "MISSING_TOTAL_T001").
'@return Boolean. True if the named range resolves to a valid Range object.
Private Function RangeExists(ByVal rngName As String) As Boolean
    Dim sh As Worksheet
    Dim rng As Range
    Set sh = Table.Wksh()
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

'@label:failed-formula
'@sub-title Test whether a generated Excel formula string is syntactically valid.
'@details
'Evaluates the formula string in memory using Application.Evaluate without
'writing to any worksheet cell. If Evaluate raises an error or returns an
'error variant (e.g., invalid function name, bad range reference, or formula
'syntax error), the formula is considered broken and FailedFormula returns
'True. This acts as a safety net to prevent writing broken formulas into
'the actual table cells.
'@param frmName String. The Excel formula string to test (without the leading "=" prefix).
'@return Boolean. True if the formula raised an error (is broken), False if it succeeded.
Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim result As Variant

    On Error Resume Next
    result = Application.Evaluate(frmName)
    FailedFormula = (Err.Number <> 0) Or IsError(result)
    On Error GoTo 0
End Function

'@label:emp-str
'@sub-title Return the Excel formula literal for an empty string.
'@details
'Returns a two-character string consisting of two double-quote characters
'(Chr(34) & Chr(34)), representing the Excel formula literal for an empty
'string (""). This is used extensively when building condition expressions
'that test whether a cell is empty. Using Chr(34) avoids VBA double-quote
'escaping issues and keeps the formula-building code readable.
'@return String. The literal text consisting of two Chr(34) characters.
Private Function EmpStr() As String
    EmpStr = Chr(34) & Chr(34)
End Function

'@label:non-emp-str
'@sub-title Return the Excel formula literal for "not empty" criteria.
'@details
'Returns a four-character string representing the Excel formula literal "<>"
'wrapped in double-quote characters (Chr(34) & "<>" & Chr(34)). This is the
'COUNTIFS criteria for "not empty" when used with the "N" or "N()" summary
'functions, which need a comparison operator inside the criteria string.
'@return String. The literal "<>" wrapped in double-quote characters for Excel formula use.
Private Function NonEmpStr() As String
    NonEmpStr = Chr(34) & "<>" & Chr(34)
End Function

'@label:percentage
'@sub-title Build an IF(ISERR(...)) guarded percentage formula string.
'@details
'Builds an Excel formula string that computes a percentage (numerator /
'denominator) with an ISERR guard to handle division-by-zero or other errors
'gracefully. The generated formula follows the pattern:
'IF(ISERR(num/denom), "", num/denom). This ensures that when the denominator
'is zero or the ranges are empty, the cell displays blank instead of a
'#DIV/0! error. Address references use mixed absolute/relative addressing so
'the formula can be auto-filled down rows while keeping the denominator column
'fixed.
'@param denomRng Range. The range whose address is used as the denominator in the division.
'@param numRng Range. The range whose address is used as the numerator.
'@param rowabs Optional Boolean. When True, the denominator row reference is absolute. Defaults to True.
'@return String. The complete IF(ISERR(...)) Excel formula text without "=" prefix.
Private Function Percentage(ByVal denomRng As Range, ByVal numRng As Range, _
                            Optional ByVal rowabs As Boolean = True) As String
    Dim parsedForm As String
    parsedForm = numRng.Address(RowAbsolute:=False) & "/" & _
                 denomRng.Address(RowAbsolute:=rowabs, ColumnAbsolute:=True)
    Percentage = "IF(ISERR(" & parsedForm & ")," & EmpStr & "," & parsedForm & ")"
End Function

'@label:condition
'@sub-title Wrap a formula in an IF guard that blanks when a condition cell is empty.
'@details
'Wraps a formula value inside an IF guard that returns an empty string when a
'specified condition cell is blank. The generated pattern is:
'IF(conditionCell = "", "", formulaValue). This is primarily used in time-series
'and spatio-temporal tables where formula cells should display blank when the
'corresponding time period label cell is empty (i.e., the row falls outside the
'user selected date range). Without this guard, formulas would show misleading
'zero/error values for unused time period rows.
'@param rngCondition Range. The cell tested for blankness.
'@param formulaValue String. The Excel formula string to evaluate when rngCondition is not blank.
'@param rowabs Optional Boolean. Controls whether rngCondition uses an absolute row reference. Defaults to False.
'@return String. The complete IF(...="","",formulaValue) Excel formula text.
Private Function Condition(ByVal rngCondition As Range, ByVal formulaValue As String, _
                           Optional ByVal rowabs As Boolean = False) As String
    Condition = "IF(" & rngCondition.Address(RowAbsolute:=rowabs) & "=" & _
                EmpStr & "," & EmpStr & "," & formulaValue & ")"
End Function

'@section Condition Management
'===============================================================================
'@description Methods that maintain the parallel condVars/condTests accumulator
'arrays used to build multi-criteria COUNTIFS/SUMPRODUCT formulas.

'@label:add-table-conditions
'@sub-title Append a variable-name / condition-test pair to the accumulators.
'@details
'Appends a variable name and its corresponding condition test string to the
'parallel accumulator arrays (condVars and condTests). These arrays are consumed
'by CurrentExcelFormula to build multi-criteria COUNTIFS/SUMPRODUCT formulas.
'Each call adds one criteria pair. The BetterArray objects are lazily initialised
'on first use with a 1-based lower bound. If varName is vbNullString, nothing
'is pushed (silent no-op).
'@param varName String. The linelist variable name identifying the column the condition applies to.
'@param conditionName String. The criteria string for the formula (e.g., "= $A3", "=""", "<>""").
'@remarks condVars and condTests always have the same length and are consumed in parallel.
'Call ClearTableConditions to reset them before building a new formula.
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
'@sub-title Clear or trim conditions from the accumulator arrays.
'@details
'Clears conditions from the condVars/condTests accumulator arrays. When called
'with removeNumber = 0 (the default), all accumulated conditions are cleared,
'resetting the formula builder to a blank state. When called with a positive
'removeNumber, only the last N condition pairs are popped off the arrays, leaving
'earlier conditions intact. This selective removal is useful when a base set of
'conditions is shared across multiple formulas and only the last few criteria
'need to change between iterations.
'@param removeNumber Optional Long. Number of condition pairs to remove from the end. Defaults to 0 (clear all).
'@remarks If the BetterArray objects have not been initialised yet (Nothing), they
'are created fresh and removeNumber is forced to 0 to avoid popping from an empty array.
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

'@label:add-non-empty-table-conditions
'@sub-title Add a "non-empty" filter condition adapted to the summary function.
'@details
'Adds a "non-empty" filter condition for a variable, automatically adapting the
'condition operator to the current summary function. For count-based functions
'("N" and "N()"), the condition uses the equality operator with the NonEmpStr
'literal ("= <>"), because COUNTIFS interprets "<>" as "not blank" inside a
'criteria string. For all other summary functions (SUM, MEAN, etc.), the condition
'uses the inequality operator with EmpStr ("<>"""), which is the standard
'SUMPRODUCT-style not-equal-to-empty test.
'@param varName String. The linelist variable name to add the non-empty filter for.
Private Sub AddNonEmptyTableConditions(ByVal varName As String)
    If SummaryFunction = "N" Or SummaryFunction = "N()" Then
        AddTableConditions varName, "=" & NonEmpStr()
    Else
        AddTableConditions varName, "<>" & EmpStr()
    End If
End Sub

'@section Excel Formula Builder
'===============================================================================
'@description Central entry point that assembles the final Excel formula string
'from the accumulated condition pairs.

'@label:current-excel-formula
'@prop-title Assembled Excel formula from the current condition accumulators.
'@details
'Builds and returns the complete Excel formula string from the currently
'accumulated condition pairs (condVars/condTests). This is the central
'formula-generation entry point that all table-type branches call after setting
'up their conditions. It operates in two modes: (1) Linelist scope calls
'ParsedLinelistFormula for unfiltered formulas; (2) Analysis scope wraps the
'accumulated condVars and condTests into a FormulaCondition object and calls
'ParsedAnalysisFormula for COUNTIFS/SUMPRODUCT formulas with all criteria. After
'generation, the formula is validated via FailedFormula, returning a fallback
'error placeholder if parsing failed.
'@param scope Optional Byte. FormulaScopeLinelist (1) for unfiltered formulas or FormulaScopeAnalysis (2) for conditional formulas. Defaults to 2.
'@param prefix Optional String. Table name prefix for the linelist reference. Defaults to "f".
'@return String. The Excel formula text without "=" prefix, or an error placeholder if parsing failed.
'@remarks If condVars or condTests is Nothing when scope is analysis, execution
'jumps to Fail immediately. Always call AddTableConditions before requesting
'an analysis-scope formula.
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

'@section Number Format Helper
'===============================================================================
'@description Translates human-readable format keywords into Excel NumberFormat strings.

'@label:add-rng-format
'@sub-title Apply an Excel NumberFormat to a range based on a format keyword.
'@details
'Translates configuration-friendly format keywords into proper Excel custom
'number format strings and applies them to the target range. Supported keywords
'include "integer", "round0", "roundN", "percentage0", "percentageN", "euros",
'"dollars", and "text". Any unrecognised string is treated as a custom Excel
'NumberFormat and applied directly inside On Error Resume Next. If formatValue
'is vbNullString, the sub exits without modifying the range.
'@param formatValue String. The format keyword or custom Excel NumberFormat string.
'@param rng Range. The range to apply the number format to.
'@remarks For "roundN" and "percentageN" keywords, the numeric suffix is extracted
'via Replace and CLng conversion. If the suffix is not a valid number, nbDecimals
'defaults to 1.
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

'@section Core Operations
'===============================================================================
'@description The main formula-writing method that dispatches by table scope.

'@label:add-formulas
'@sub-title Write all formulas to the cross-table on the worksheet.
'@details
'The main entry point that writes all Excel formulas into the cross-table
'worksheet cells. It reads the table scope from Specifications and dispatches to
'one of five branches, each handling a fundamentally different table layout:
'ScopeGlobalSummary writes two unfiltered linelist formulas into a single row;
'ScopeUnivariate writes a one-variable frequency table with missing/total/percentage
'support; ScopeBivariate writes a two-variable cross-tabulation grid with column
'loops and corner cells; ScopeTimeSeries and ScopeSpatioTemporal set up time-period
'scaffolding (date labels, validation, dropdowns) then write time-bounded conditions
'wrapped in blank-when-empty guards; ScopeSpatial sets up geographic or health-facility
'lookups, validation dropdowns, ordering, and registers the table with SpatialTables
'for refresh operations. The method exits immediately if Valid returns False.
'@export
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
    tabType = TableScope
    hasMissing = specs.HasMissing
    hasTotal = specs.HasTotal
    hasPercentage = specs.HasPercentage
    tabFormat = specs.Value("format")

    With sh

        Select Case tabType

        Case ScopeGlobalSummary
            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:=vbNullString)
            Set rng = .Range("STARTROW_" & tabId)
            rng.Cells(1, 2).FormulaArray = "=" & finalFormula
            AddRngFormat tabFormat, rng.Cells(1, 2)

            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:="f")
            rng.Cells(1, 3).FormulaArray = "=" & finalFormula
            AddRngFormat tabFormat, rng.Cells(1, 3)

        Case ScopeUnivariate

            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            rowVar = specs.Value("row")
            ClearTableConditions
            AddTableConditions rowVar, "=" & rng.Cells(1, 1).Address(RowAbsolute:=False)
            finalFormula = CurrentExcelFormula()

            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set cellRng = rng.Cells(1, 1)
            cellRng.FormulaArray = "=" & finalFormula
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
                    AddTableConditions rowVar, "=" & NonEmpStr()
                    finalFormula = finalFormula & " + " & CurrentExcelFormula()
                    rng.Cells(1, 1).FormulaArray = "=" & finalFormula
                Else
                    ClearTableConditions
                    finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist)
                    rng.Cells(1, 1).FormulaArray = "=" & finalFormula
                End If
            Else
                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)
                ClearTableConditions
                AddNonEmptyTableConditions rowVar
                finalFormula = CurrentExcelFormula()
                rng.Cells(1, 1).FormulaArray = "=" & finalFormula
            End If

            ' Percentage
            If hasPercentage Then
                Set rng = .Range("VALUES_COL_1_" & tabId)
                Set cellRng = .Range("TOTAL_ROW_VALUES_" & tabId).Cells(1, 1)
                rng.Offset(, 1).Cells(1, 1).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1), True)
                rng.Offset(, 1).Cells(1, 1).AutoFill Destination:=rng.Offset(, 1), Type:=xlFillValues
                cellRng.Cells(1, 2).Formula = "=" & Percentage(cellRng.Cells(1, 1), cellRng.Cells(1, 1))
            End If

        Case ScopeBivariate

            rowVar = specs.Value("row")
            colVar = specs.Value("column")
            rowCatAddress = .Range("ROW_CATEGORIES_" & tabId).Cells(1, 1).Address(RowAbsolute:=False)
            tabEndColumn = tabl.EndColumn

            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)

            Do While rng.Column < tabEndColumn

                ClearTableConditions
                AddTableConditions rowVar, "=" & rowCatAddress
                AddTableConditions colVar, "=" & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "=" & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                If specs.Value("missing") = "row" Or specs.Value("missing") = "all" Then
                    Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & EmpStr()
                    AddTableConditions colVar, "=" & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(misRng.Row, rng.Column)
                    cellRng.FormulaArray = "=" & finalFormula

                    ClearTableConditions
                    AddTableConditions colVar, "=" & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "=" & finalFormula
                Else
                    ClearTableConditions
                    AddNonEmptyTableConditions rowVar
                    AddTableConditions colVar, "=" & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "=" & finalFormula
                End If

                ' Percentage
                If hasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1))
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
                AddTableConditions rowVar, "=" & rowCatAddress
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                misRng.Cells(1, 1).FormulaArray = "=" & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                ClearTableConditions
                AddTableConditions rowVar, "=" & rowCatAddress
                finalFormula = CurrentExcelFormula()
            Else
                ClearTableConditions
                AddTableConditions rowVar, "=" & rowCatAddress
                AddNonEmptyTableConditions colVar
                finalFormula = CurrentExcelFormula()
            End If

            Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
            totRng.Cells(1, 1).FormulaArray = "=" & finalFormula
            totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
            AddRngFormat tabFormat, totRng

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Total row, missing column
            If RangeExists("TOTAL_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("TOTAL_MISSING_" & tabId).FormulaArray = "=" & finalFormula
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
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
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

        Case ScopeTimeSeries, ScopeSpatioTemporal

            rowVar = specs.Value("time")
            colVar = specs.Value("column")
            counter = 1

            ' For spatio-temporal, column variable depends on geo/hf type
            If tabType = ScopeSpatioTemporal Then
                spTabType = specs.SpatialTableScopes()
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
                cellRng.Cells(2, 1).Formula = "=" & _
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
                AddTableConditions colVar, "=" & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "=" & Condition(perRng, finalFormula)
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & rngLab.Address
                finalFormula = CurrentExcelFormula()
                Set cellRng = .Cells(misRng.Row, rng.Column)
                cellRng.FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, cellRng

                ' Percentage
                If hasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                End If

                ' Spatio-temporal label formulas
                If tabType = ScopeSpatioTemporal Then
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
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()

                misRng.Cells(1, 1).FormulaArray = "=" & Condition(perRng, finalFormula)
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "=" & Condition(perRng, finalFormula)
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
                    totRng.Cells(1, 1).FormulaArray = "=" & Condition(perRng, finalFormula)
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            End If

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Total row, missing column
            If RangeExists("TOTAL_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("TOTAL_MISSING_" & tabId).FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, .Range("TOTAL_MISSING_" & tabId)
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

        Case ScopeSpatial

            spTabType = specs.SpatialTableScopes()
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
            cellRng.Formula = "=" & cellRng.Offset(-1).Address(RowAbsolute:=False) & " + 1"
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
                AddTableConditions rowVar, "=" & rowCatAddress
                AddNonEmptyTableConditions rowVar
                AddTableConditions colVar, "=" & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "=" & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & rngLab.Address
                finalFormula = CurrentExcelFormula()
                Set cellRng = .Cells(misRng.Row, rng.Column)
                cellRng.FormulaArray = "=" & finalFormula

                ' Percentage (spatial only supports row percentage)
                If hasPercentage Then
                    Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                    rng.Cells(1, 2).Formula = "=" & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
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
                AddTableConditions rowVar, "=" & rowCatAddress
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()

                misRng.Cells(1, 1).FormulaArray = "=" & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & rowCatAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "=" & finalFormula
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            Else
                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & rowCatAddress
                    AddNonEmptyTableConditions colVar
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "=" & finalFormula
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            End If

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "=" & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Total row, missing column
            If RangeExists("TOTAL_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions colVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("TOTAL_MISSING_" & tabId).FormulaArray = "=" & finalFormula
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
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
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

            ' Create spatial table management
            Set spTab = SpatialTables.Create(tabl)
            Set tabFormObject = TableFormula
            spTab.Add tabFormObject

        End Select
    End With
End Sub

'@section Error Handling
'===============================================================================
'@description Standardised error-raising helper for all methods in this class.

'@label:throw-error
'@sub-title Raise a project error with a standardised source name.
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber Long. The error code to raise (typically a ProjectError constant).
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ICrossTableFormula contract.
'See the corresponding Public/Private members above for full documentation.

Private Property Get ICrossTableFormula_Valid() As Boolean
    ICrossTableFormula_Valid = Valid
End Property

Private Sub ICrossTableFormula_AddFormulas()
    AddFormulas
End Sub
