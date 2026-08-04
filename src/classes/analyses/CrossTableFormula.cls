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

'@class CrossTableFormula
'@description
'Writes the Excel formulas of one cross-table into the cells of an analysis
'worksheet. Each instance is bound to one CrossTable and its FormulaData.
'AddFormulas dispatches on the table scope -- global summary, univariate,
'bivariate, time series, spatio-temporal and spatial -- and each branch writes
'the COUNTIFS and SUMPRODUCT array formulas, the percentage columns, the number
'formats, the time period scaffolding, the spatial lookups and the data
'validation dropdowns that scope needs. Valid says whether the table can be
'written, and entries filed along the way leave through HasCheckings and
'CheckingValues.
'
'EXCEL IS THE JUDGE OF A FORMULA
'-------------------------------------------------------------------------------
'Every formula reaches its cell through WriteFormula, which assigns it and
'traps what Excel raises. A formula Excel refuses leaves the cell empty and
'files a check, so the generation report names the table and the cell. Read
'AddFormulas for what is written and Checkings for what is reported.
'
'ONE BUILDER, SHARED WITH SPATIALTABLES
'-------------------------------------------------------------------------------
'The criteria of a formula, the text the parser gives back, the write into the
'cell and the 255 character limit of an array formula all live in
'FormulaBuilder, and SpatialTables writes through one of those too.
'
'WHICH COLUMNS THE VALUE LOOP OWNS
'-------------------------------------------------------------------------------
'The three column loops walk the data columns, one turn per column of
'CrossTable.NumberOfColumns, and advance by two when the table carries
'percentages. The Missing column, the Total column and the four cells where
'they cross are written by their own blocks below each loop.
'CrossTable.EndColumn counts the missing and the total columns, so a loop that
'walked to it would write value formulas into them with the criteria of a data
'column.
'@depends CrossTable, TableSpecs, FormulaData, Formulas, FormulaBuilder, SpatialTables, AnalysisRanges, Checking
'@version 1.1 (2026-07-30)


Private Type TCrossTableFormula
    tabl As CrossTable
    fData As FormulaData
    forml As Formulas
    build As FormulaBuilder
    sealed As Boolean
End Type

Private this As TCrossTableFormula
Private internalChecks As Checking
Private checkCounter As Long

Private Const CLASS_NAME As String = "CrossTableFormula"

' The formula context the parser is asked for. Formulas.Valid and
' Formulas.ParsedAnalysisFormula both key off this string.
Private Const CONTEXT_ANALYSIS As String = "analysis"

' The two count summary functions, which take a different formula shape from
' every other summary function.
Private Const COUNT_FUNCTION As String = "N"
Private Const COUNT_FUNCTION_CALL As String = "N()"

'@section Instantiation
'===============================================================================

'@label:create

'@sub-title Create a CrossTableFormula instance bound to a cross-table and formula data source.
'@details
'Factory method that returns a new CrossTableFormula through the predeclared-id
'pattern. Both parameters are validated for Nothing before the instance is
'constructed.
'@param tabl CrossTable. The already-constructed cross-table layout providing named ranges, headers, and worksheet reference.
'@param fData FormulaData. The dictionary-based lookup object supplying variable names, ranges, and formula templates.
'@return CrossTableFormula. A fully initialised instance ready to write formulas.
'@throws InvalidArgument When tabl is Nothing.
'@throws InvalidArgument When fData is Nothing.
'@export
Public Function Create(ByVal tabl As CrossTable, _
                       ByVal fData As FormulaData) As CrossTableFormula

    If tabl Is Nothing Then ThrowError InvalidArgument, "Cross-table not defined"
    If fData Is Nothing Then ThrowError InvalidArgument, "Formula data not defined"

    Dim instance As CrossTableFormula
    Set instance = New CrossTableFormula
    Set instance.Table = tabl
    Set instance.formData = fData
    instance.Seal
    Set Create = instance
End Function

'@section Internal Properties
'===============================================================================
'@description Properties used during factory construction and internal wiring.

'@label:seal
'@sub-title Seal the instance against further setup writes.
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

'@label:table-get
'@prop-title CrossTable backing store.
'@details
'Returns the cross-table object that this formula writer targets. The
'cross-table provides the worksheet reference, named ranges, and layout
'metadata needed to place formulas.
'@return CrossTable. The bound cross-table instance.
Public Property Get Table() As CrossTable
    Set Table = this.tabl
End Property

'@label:table-set
'@prop-title Assign the CrossTable backing store.
'@param tabl CrossTable. The cross-table to bind to this formula writer.
Public Property Set Table(ByVal tabl As CrossTable)
    GuardNotSealed "Table"
    Set this.tabl = tabl
End Property

'@label:form-data-get
'@prop-title FormulaData backing store.
'@details
'Returns the formula-data object that supplies dictionary-based lookups for
'variable names, linelist ranges, and formula templates during generation.
'@return FormulaData. The bound formula-data instance.
Public Property Get formData() As FormulaData
    Set formData = this.fData
End Property

'@label:form-data-set
'@prop-title Assign the FormulaData backing store.
'@param fData FormulaData. The formula-data source to bind.
Public Property Set formData(ByVal fData As FormulaData)
    GuardNotSealed "formData"
    Set this.fData = fData
End Property

'@section Shortcut Properties
'===============================================================================
'@description Convenience accessors that delegate to the cross-table or its
'specifications, reducing repetitive chaining throughout the class.

'@label:specifications
'@prop-title TableSpecs shortcut for the cross-table specifications.
'@details
'Provides a shortcut to the TableSpecs object associated with this cross-table.
'TableSpecs encapsulates all configuration metadata for a table: its row/column
'variable names, summary function, table scope, missing/total/percentage flags,
'and the section it belongs to. This property avoids repeatedly chaining
'Table.Specifications throughout the class methods.
'@return TableSpecs. The specification dictionary for the underlying cross-table.
Private Property Get Specifications() As TableSpecs
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

'@label:is-count-function
'@prop-title Whether the summary function counts rows.
'@details
'The two count spellings take a different formula shape from every other
'summary function: a total is a sum of counts and a non-empty test is written as
'a COUNTIFS criteria. Six places asked this question with the same pair of
'string tests.
'@return Boolean. True for "N" and "N()".
Private Property Get IsCountFunction() As Boolean
    Dim summaryName As String

    summaryName = SummaryFunction
    IsCountFunction = (summaryName = COUNT_FUNCTION) Or (summaryName = COUNT_FUNCTION_CALL)
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
'@prop-title Lazy-loaded Formulas parsing engine for this table.
'@details
'Returns the Formulas object for this table, creating it on first access via
'lazy initialisation. The Formulas object (created by Formulas.Create) is the
'core formula-parsing engine: it knows how to translate the table summary function
'and variable dictionary into concrete Excel formula strings for both linelist-scope
'and analysis-scope contexts. Once created, the instance is cached in this.forml
'so subsequent calls reuse the same object without re-parsing.
'@return Formulas. The formula parsing engine providing ParsedLinelistFormula and ParsedAnalysisFormula methods.
'@remarks The Formulas object is initialised with the table specification dictionary,
'the FormulaData reference, and the summary function string. This is intentionally
'lazy because not all code paths need the formula engine (e.g., if Valid is False).
Private Property Get TableFormula() As Formulas
    If this.forml Is Nothing Then
        Dim specs As TableSpecs
        Set specs = Specifications
        Set this.forml = Formulas.Create(specs.Dictionary, formData(), SummaryFunction)
    End If
    Set TableFormula = this.forml
End Property

'@label:valid
'@prop-title Whether the formula object is valid for the analysis context.
'@details
'Checks whether the formula engine considers this table formula valid for
'analysis use. Validity is determined by the Formulas.Valid method with the
'analysis formula context. A formula may be invalid if the summary function is
'unsupported, required variables are missing from the dictionary, or the formula
'template cannot be parsed. When Valid returns False, the AddFormulas method
'exits immediately without writing any formulas to the worksheet.
'@return Boolean. True if the formula can be used for analysis-scope generation.
'@export
Public Property Get Valid() As Boolean
    Valid = TableFormula.Valid(formulaType:=CONTEXT_ANALYSIS)
End Property

'@section Formula Builder
'===============================================================================
'@description The one place a formula is assembled and written. SpatialTables
'writes through one of these too, so the two classes agree on how a formula is
'made, on the 255 character limit of an array formula, and on what a refusal
'looks like in the report.

'@label:builder
'@prop-title The builder this class writes its formulas through.
'@details
'Created on first use, because it carries the table identifier and the entries
'both travel into every message it files.
'@return FormulaBuilder. The builder of this instance.
Private Property Get Builder() As FormulaBuilder
    Dim entries As Checking

    If this.build Is Nothing Then
        Set entries = Checks
        Set this.build = FormulaBuilder.Create(CLASS_NAME, ID(), entries)
    End If

    Set Builder = this.build
End Property

'@section Range and Formula Helpers
'===============================================================================
'@description Utility functions for range existence checks, string literals,
'percentage formulas, and conditional wrapping.

'@label:range-exists
'@sub-title Check whether a named range exists on the cross-table worksheet.
'@details
'Attempts to resolve a named range on the cross-table worksheet, answering False
'when the name is absent. This is used before writing formulas to optional table
'regions (e.g., "MISSING_TOTAL_<id>", "TOTAL_TOTAL_<id>") that may or may not be
'present depending on the table missing/total configuration.
'@param rngName String. The named range to look up (e.g., "MISSING_TOTAL_T001").
'@return Boolean. True if the named range resolves to a valid Range object.
Private Function RangeExists(ByVal rngName As String) As Boolean
    RangeExists = Not (ResolvedRange(rngName) Is Nothing)
End Function

'@label:resolved-range
'@sub-title Resolve a named range on the cross-table worksheet, or answer Nothing.
'@details
'One resolution serves both the question and the answer. Twelve places asked
'RangeExists and then resolved the same name a second time, so each optional
'region cost two lookups.
'@param rngName String. The named range to look up.
'@return Range. The range, or Nothing when the name is absent.
Private Function ResolvedRange(ByVal rngName As String) As Range
    Dim sh As Worksheet
    Dim rng As Range

    Set sh = Table.Wksh()
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    Set ResolvedRange = rng
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
    EmpStr = Builder.EmptyText
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
    NonEmpStr = Builder.NotEmptyText
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
    Percentage = Builder.Percentage(denomRng, numRng, rowabs)
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
    Condition = Builder.Condition(rngCondition, formulaValue, rowabs)
End Function

'@section Writing Formulas
'===============================================================================
'@description The one way a formula reaches a cell. Excel accepts it or raises,
'and a raise is reported through Checkings.

'@label:write-formula
'@sub-title Write one formula into one cell and report what Excel says.
'@details
'Assigns the formula and traps the assignment. Excel rejects a malformed
'formula with error 1004 at assignment time, which makes it the only judge the
'class needs. A rejected formula leaves the cell empty and files a check
'carrying the address and the message Excel gave.
'
'A formula over 255 characters cannot go through Range.FormulaArray at all, so
'it is entered through a stub and Range.Replace. See the class description.
'@param target Range. The cell the formula belongs in.
'@param frm String. The formula text with no leading "=".
'@param asArray Optional Boolean. True to enter the formula as an array formula. Defaults to True.
'@return Boolean. True when the formula reached the cell.
Private Function WriteFormula(ByVal target As Range, ByVal frm As String, _
                              Optional ByVal asArray As Boolean = True) As Boolean
    WriteFormula = Builder.WriteFormula(target, frm, asArray)
End Function

'@label:fill-down
'@sub-title Copy the formula of one cell over a range.
'@details
'AutoFill needs a destination taller than one cell and raises 1004 otherwise.
'A value column is one cell tall whenever the table has a single category row,
'which a univariate table with one category and a spatial table with a geo count
'of one both produce, so the guard is on the ordinary path.
'@param sourceCell Range. The cell holding the formula to copy.
'@param destination Range. The range to fill, including the source cell.
Private Sub FillDown(ByVal sourceCell As Range, ByVal destination As Range)
    Builder.FillDown sourceCell, destination
End Sub

'@label:write-percentage-twin
'@sub-title Write the percentage column that sits beside one value column.
'@details
'A table with percentages carries one percentage column immediately right of
'every value column, including the ones heading the Missing and the Total
'columns. The formula is written into the first cell and copied down to the
'last cell the caller names, because a bivariate table fills its percentage
'column down to the total row while a temporal one stops at the value block.
'@param valueRng Range. The value column the percentage is computed from.
'@param denomRng Range. The cell holding the denominator.
'@param lastCell Range. The last cell of the percentage column to fill.
'@param rowabs Boolean. True to hold the denominator row fixed.
Private Sub WritePercentageTwin(ByVal valueRng As Range, ByVal denomRng As Range, _
                                ByVal lastCell As Range, ByVal rowabs As Boolean)
    Dim firstCell As Range
    Dim sh As Worksheet

    If denomRng Is Nothing Then Exit Sub

    Set sh = Table.Wksh()
    Set firstCell = valueRng.Cells(1, 2)

    If WriteFormula(firstCell, Percentage(denomRng, valueRng.Cells(1, 1), rowabs), False) Then
        FillDown firstCell, sh.Range(firstCell, lastCell)
    End If
End Sub

'@label:percentage-denominator
'@sub-title Resolve the denominator one percentage column divides by.
'@details
'The setup sheet offers three percentages and each names a different
'denominator: a row percentage divides by the total column of the same row, a
'column percentage by the total row of the same column, and a total percentage
'by the cell where the two cross. Three arms of AddFormulas held the same
'Select Case.
'@param tabNames AnalysisRanges. The names of this table.
'@param percentageValue String. The percentage setting of the table.
'@param totalRow Long. The row holding the totals.
'@param valueColumn Long. The column the percentage belongs to.
'@return Range. The denominator cell, or Nothing when the setting names none.
Private Function PercentageDenominator(ByVal tabNames As AnalysisRanges, _
                                       ByVal percentageValue As String, _
                                       ByVal totalRow As Long, _
                                       ByVal valueColumn As Long) As Range
    Dim sh As Worksheet
    Dim rng As Range

    Set sh = Table.Wksh()

    Select Case percentageValue
    Case "row"
        Set rng = ResolvedRange(tabNames.TotalColValues)
        If Not rng Is Nothing Then Set rng = rng.Cells(1, 1)
    Case "column"
        Set rng = sh.Cells(totalRow, valueColumn)
    Case "total"
        Set rng = ResolvedRange(tabNames.TotalTotal)
    Case Else
        ' A setting the setup sheet never offered. The callers already test for
        ' Nothing, so the cell is left alone; saying so is what stops an empty
        ' percentage column reading as a formula the parser refused.
        LogInfo "Table " & ID() & " asks for the percentage setting '" & _
                percentageValue & "', which names no denominator", checkingWarning
    End Select

    Set PercentageDenominator = rng
End Function

'@label:percentage-row-absolute
'@sub-title Whether the percentage denominator holds its row fixed.
'@param percentageValue String. The percentage setting of the table.
'@return Boolean. False for a row percentage, True for the other two.
Private Function PercentageRowAbsolute(ByVal percentageValue As String) As Boolean
    PercentageRowAbsolute = (percentageValue <> "row")
End Function

'@section Condition Management
'===============================================================================
'@description The criteria of the formula being built. The builder holds them,
'and it starts a new set on the first criterion added after a formula was built,
'so an arm that forgets to clear cannot carry the criteria of the previous cell.

'@label:add-table-conditions
'@sub-title Add one criterion to the formula being built.
'@details
'One criterion is a variable name and the test applied to it, and the two travel
'as a pair into the COUNTIFS or SUMPRODUCT formula CurrentExcelFormula asks for.
'An empty variable name is ignored, which is how a caller passes an absent
'column.
'@param varName String. The linelist variable name identifying the column the condition applies to.
'@param conditionName String. The criteria string for the formula.
Private Sub AddTableConditions(ByVal varName As String, ByVal conditionName As String)
    Builder.AddCondition varName, conditionName
End Sub

'@label:clear-table-conditions
'@sub-title Empty the criteria set before a new formula is built.
'@details
'Every arm opens with this and says what it is about to build. The builder also
'clears for itself, so a forgotten call costs nothing. The routine used to take a
'count of pairs to pop off the end, and all sixty-odd call sites took the
'default, so that branch was unreachable.
Private Sub ClearTableConditions()
    Builder.ClearConditions
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
    If IsCountFunction() Then
        AddTableConditions varName, "=" & NonEmpStr()
    Else
        AddTableConditions varName, "<>" & EmpStr()
    End If
End Sub

'@section Excel Formula Text
'===============================================================================
'@description Central entry point that asks the builder for the formula text the
'current criteria describe.

'@label:current-excel-formula
'@prop-title Assembled Excel formula from the current condition accumulators.
'@details
'Builds and returns the complete Excel formula string from the currently
'criteria the arm has added. This is the central formula-generation entry point
'that all table-type branches call after setting up their conditions. It operates
'in two modes: (1) Linelist scope reads the whole column and ignores the
'criteria; (2) Analysis scope builds a COUNTIFS or SUMPRODUCT formula carrying
'every criterion.
'
'The text is handed back as the parser built it. Excel judges it when
'WriteFormula assigns it.
'@param scope Optional Byte. FormulaScopeLinelist for unfiltered formulas or FormulaScopeAnalysis for conditional formulas. Defaults to the analysis scope.
'@param prefix Optional String. Table name prefix for the linelist reference. Defaults to "f".
'@return String. The Excel formula text without "=" prefix, or an empty string when the parser produced none.
Private Property Get CurrentExcelFormula(Optional ByVal scope As Byte = FormulaScopeAnalysis, _
                                         Optional ByVal prefix As String = "f") As String
    Dim tabFormObject As Formulas

    Set tabFormObject = TableFormula
    CurrentExcelFormula = Builder.ExcelFormula(tabFormObject, scope, prefix)
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
    Builder.ApplyFormat formatValue, rng
End Sub

'@section Validation Lists
'===============================================================================
'@description The three dropdowns a table offers its reader.

'@label:add-list-validation
'@sub-title Point a cell at a workbook-wide choice list.
'@details
'The time unit control, the administrative unit control and the population
'divisor each take their values from one list shared by every table on the
'sheet. The three writes used to sit inside blanket error handlers, so a table
'whose dropdown was never created looked exactly like one whose dropdown works.
'@param targetName String. The name of the cell carrying the dropdown.
'@param listName String. The name of the choice list the dropdown reads.
Private Sub AddListValidation(ByVal targetName As String, ByVal listName As String)
    Dim target As Range
    Dim failureReason As String

    Set target = ResolvedRange(targetName)

    If target Is Nothing Then
        LogInfo "Cell " & targetName & " is missing, so its dropdown was not created", _
                checkingError
        Exit Sub
    End If

    On Error GoTo ValidationFailed
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & listName
    End With
    On Error GoTo 0
    Exit Sub

ValidationCleanup:
    On Error Resume Next
    LogInfo "The dropdown of " & targetName & " reading " & listName & _
            " was not created: " & failureReason, checkingError
    On Error GoTo 0
    Exit Sub

ValidationFailed:
    failureReason = Err.Description
    Resume ValidationCleanup
End Sub

'@section Core Operations
'===============================================================================
'@description The main formula-writing method that dispatches by table scope.

'@label:add-formulas
'@sub-title Write all formulas to the cross-table on the worksheet.
'@details
'The main entry point that writes all Excel formulas into the cross-table
'worksheet cells. It reads the table scope from Specifications and dispatches to
'one of six branches, each handling a fundamentally different table layout:
'ScopeGlobalSummary writes two unfiltered linelist formulas into a single row;
'ScopeUnivariate writes a one-variable frequency table with missing/total/percentage
'support; ScopeBivariate writes a two-variable cross-tabulation grid with column
'loops and corner cells; ScopeTimeSeries and ScopeSpatioTemporal set up time-period
'scaffolding (date labels, validation, dropdowns) then write time-bounded conditions
'wrapped in blank-when-empty guards; ScopeSpatial sets up geographic or health-facility
'lookups, validation dropdowns, ordering, and registers the table with SpatialTables
'for refresh operations. The method exits immediately if Valid returns False.
'@export
Public Sub AddFormulas()
    Dim tabId As String
    Dim tabType As Byte
    Dim sh As Worksheet
    Dim rng As Range
    Dim totRng As Range
    Dim rngLab As Range
    Dim misRng As Range
    Dim cellRng As Range
    Dim percRng As Range
    Dim popRng As Range
    Dim anchorRng As Range
    Dim denomRng As Range
    Dim finalFormula As String
    Dim countFormula As String
    Dim secId As String
    Dim perRng As Range
    Dim rowVar As String
    Dim colVar As String
    Dim tabFormat As String
    Dim rowCatAddress As String
    Dim labAddress As String
    Dim startTimeAddress As String
    Dim endTimeAddress As String
    Dim hasMissing As Boolean
    Dim hasTotal As Boolean
    Dim hasPercentage As Boolean
    Dim missingValue As String
    Dim percentageValue As String
    Dim nbDataColumns As Long
    Dim columnStride As Long
    Dim colCounter As Long
    Dim totalRowNumber As Long
    Dim tabl As CrossTable
    Dim spTab As SpatialTables
    Dim spatialChecks As Checking
    Dim tableChecks As Checking
    Dim spTabType As String
    Dim specs As TableSpecs
    Dim tabNames As AnalysisRanges
    Dim secNames As AnalysisRanges

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

    ' The two setup answers used to be re-read inside every column loop, and
    ' TableSpecs.Value is a header search followed by a cell read.
    missingValue = specs.Value("missing")
    percentageValue = specs.Value("percentage")

    ' A percentage table carries one percentage column right of each value
    ' column, so the walk steps two columns at a time.
    columnStride = IIf(hasPercentage, 2, 1)
    nbDataColumns = tabl.NumberOfColumns

    ' The two name builders. TOTAL_ROW_ is built from the table id in one place
    ' of this class and from the section id in another, so each call site says
    ' which of the two it means.
    Set tabNames = AnalysisRanges.Create(tabId)
    Set secNames = AnalysisRanges.Create(secId)

    With sh

        Select Case tabType

        Case ScopeGlobalSummary
            Set rng = .Range(tabNames.StartRow)

            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:=vbNullString)
            If WriteFormula(rng.Cells(1, 2), finalFormula) Then _
                AddRngFormat tabFormat, rng.Cells(1, 2)

            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:="f")
            If WriteFormula(rng.Cells(1, 3), finalFormula) Then _
                AddRngFormat tabFormat, rng.Cells(1, 3)

        Case ScopeUnivariate

            Set rng = .Range(tabNames.RowCategories)
            rowVar = specs.Value("row")
            ClearTableConditions
            AddTableConditions rowVar, "=" & rng.Cells(1, 1).Address(RowAbsolute:=False)
            finalFormula = CurrentExcelFormula()

            Set rng = .Range(tabNames.ValuesCol(1))
            Set cellRng = rng.Cells(1, 1)
            If WriteFormula(cellRng, finalFormula) Then FillDown cellRng, rng
            AddRngFormat tabFormat, rng

            ' Missing and total
            If hasMissing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                Set rng = .Range(tabNames.MissingRowValues)
                WriteFormula rng.Cells(1, 1), finalFormula

                Set rng = .Range(tabNames.TotalRowValues)
                If IsCountFunction() Then
                    ' A count total is the missing count plus the non-empty
                    ' count, so both halves have to be there.
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & NonEmpStr()
                    countFormula = CurrentExcelFormula()
                    If LenB(finalFormula) > 0 And LenB(countFormula) > 0 Then
                        WriteFormula rng.Cells(1, 1), finalFormula & " + " & countFormula
                    Else
                        WriteFormula rng.Cells(1, 1), vbNullString
                    End If
                Else
                    ClearTableConditions
                    WriteFormula rng.Cells(1, 1), CurrentExcelFormula(scope:=FormulaScopeLinelist)
                End If
            Else
                Set rng = .Range(tabNames.TotalRowValues)
                ClearTableConditions
                AddNonEmptyTableConditions rowVar
                WriteFormula rng.Cells(1, 1), CurrentExcelFormula()
            End If

            ' Percentage
            If hasPercentage Then
                Set rng = .Range(tabNames.ValuesCol(1))
                Set cellRng = .Range(tabNames.TotalRowValues).Cells(1, 1)
                Set percRng = rng.Offset(, 1)
                If WriteFormula(percRng.Cells(1, 1), _
                                Percentage(cellRng, rng.Cells(1, 1), True), False) Then
                    FillDown percRng.Cells(1, 1), percRng
                End If
                WriteFormula cellRng.Cells(1, 2), _
                             Percentage(cellRng.Cells(1, 1), cellRng.Cells(1, 1)), False
            End If

        Case ScopeBivariate

            rowVar = specs.Value("row")
            colVar = specs.Value("column")
            rowCatAddress = .Range(tabNames.RowCategories).Cells(1, 1).Address(RowAbsolute:=False)

            Set rng = .Range(tabNames.ValuesCol(1))
            Set rngLab = .Range(tabNames.LabelCol(1))
            Set totRng = .Range(tabNames.TotalRowValues)
            totalRowNumber = totRng.Row
            Set misRng = ResolvedRange(tabNames.MissingRowValues)

            If (missingValue = "row" Or missingValue = "all") And misRng Is Nothing Then
                LogInfo "Table " & tabId & " asks for a missing row and has none", checkingError
            End If

            For colCounter = 1 To nbDataColumns

                labAddress = rngLab.Address

                ClearTableConditions
                AddTableConditions rowVar, "=" & rowCatAddress
                AddTableConditions colVar, "=" & labAddress
                finalFormula = CurrentExcelFormula()

                If WriteFormula(rng.Cells(1, 1), finalFormula) Then FillDown rng.Cells(1, 1), rng
                AddRngFormat tabFormat, rng

                ' Missing row
                If (missingValue = "row" Or missingValue = "all") And Not misRng Is Nothing Then
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & EmpStr()
                    AddTableConditions colVar, "=" & labAddress
                    WriteFormula .Cells(misRng.Row, rng.Column), CurrentExcelFormula()

                    ClearTableConditions
                    AddTableConditions colVar, "=" & labAddress
                    WriteFormula .Cells(totalRowNumber, rng.Column), CurrentExcelFormula()
                Else
                    ClearTableConditions
                    AddNonEmptyTableConditions rowVar
                    AddTableConditions colVar, "=" & labAddress
                    WriteFormula .Cells(totalRowNumber, rng.Column), CurrentExcelFormula()
                End If

                ' Percentage
                If hasPercentage Then
                    Set denomRng = PercentageDenominator(tabNames, percentageValue, _
                                                         totalRowNumber, rng.Column)
                    WritePercentageTwin rng, denomRng, _
                                        .Cells(totalRowNumber, rng.Column + 1), _
                                        PercentageRowAbsolute(percentageValue)
                End If

                Set rng = rng.Offset(, columnStride)
                Set rngLab = rngLab.Offset(, columnStride)
            Next

            ' Missing column
            If missingValue = "column" Or missingValue = "all" Then
                Set misRng = ResolvedRange(tabNames.MissingColValues)

                If misRng Is Nothing Then
                    LogInfo "Table " & tabId & " asks for a missing column and has none", _
                            checkingError
                Else
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & rowCatAddress
                    AddTableConditions colVar, "=" & EmpStr()
                    finalFormula = CurrentExcelFormula()

                    If WriteFormula(misRng.Cells(1, 1), finalFormula) Then _
                        FillDown misRng.Cells(1, 1), misRng
                    AddRngFormat tabFormat, misRng

                    If hasPercentage Then
                        Set denomRng = PercentageDenominator(tabNames, percentageValue, _
                                                             totalRowNumber, misRng.Column)
                        WritePercentageTwin misRng, denomRng, _
                                            .Cells(totalRowNumber, misRng.Column + 1), _
                                            PercentageRowAbsolute(percentageValue)
                    End If
                End If

                ClearTableConditions
                AddTableConditions rowVar, "=" & rowCatAddress
                finalFormula = CurrentExcelFormula()
            Else
                ClearTableConditions
                AddTableConditions rowVar, "=" & rowCatAddress
                AddNonEmptyTableConditions colVar
                finalFormula = CurrentExcelFormula()
            End If

            ' Total column
            Set totRng = .Range(tabNames.TotalColValues)
            If WriteFormula(totRng.Cells(1, 1), finalFormula) Then _
                FillDown totRng.Cells(1, 1), totRng
            AddRngFormat tabFormat, totRng

            If hasPercentage Then
                Set denomRng = PercentageDenominator(tabNames, percentageValue, _
                                                     totalRowNumber, totRng.Column)
                WritePercentageTwin totRng, denomRng, _
                                    .Cells(totalRowNumber, totRng.Column + 1), _
                                    PercentageRowAbsolute(percentageValue)
            End If

            ' Missing row, total column
            Set cellRng = ResolvedRange(tabNames.MissingTotal)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Missing row, missing column
            Set cellRng = ResolvedRange(tabNames.MissingMissing)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Total row, missing column
            Set cellRng = ResolvedRange(tabNames.TotalMissing)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions colVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Total row, total column
            Set cellRng = ResolvedRange(tabNames.TotalTotal)
            If Not cellRng Is Nothing Then
                If IsCountFunction() Then
                    WriteFormula cellRng, _
                                 "SUM(" & .Range(tabNames.TotalColValues).Address & ")", False
                Else
                    Select Case missingValue
                    Case "all"
                        WriteFormula cellRng, CurrentExcelFormula(scope:=FormulaScopeLinelist)
                    Case "row"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        WriteFormula cellRng, CurrentExcelFormula()
                    Case "column"
                        ClearTableConditions
                        AddNonEmptyTableConditions rowVar
                        WriteFormula cellRng, CurrentExcelFormula()
                    Case "no"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        AddNonEmptyTableConditions rowVar
                        WriteFormula cellRng, CurrentExcelFormula()
                    Case Else
                        ' A missing setting outside the four leaves the cell
                        ' where the total row meets the total column with no
                        ' formula, and the table still reads as complete.
                        LogInfo "Table " & ID() & " asks for the missing setting '" & _
                                missingValue & "', which this class does not write", _
                                checkingWarning
                    End Select
                End If
                AddRngFormat tabFormat, cellRng
            End If

        Case ScopeTimeSeries, ScopeSpatioTemporal

            rowVar = specs.Value("time")
            colVar = specs.Value("column")

            ' For spatio-temporal, the column variable is the prefixed one the
            ' geo lines of the dictionary carry. TableSpecs owns the mapping
            ' because an administrative table reads the concatenated column and
            ' a facility table reads the plain one.
            If tabType = ScopeSpatioTemporal Then
                spTabType = specs.SpatialTableScopes()
                If LenB(spTabType) = 0 Then
                    LogInfo "Table " & tabId & " groups by a variable that is neither " & _
                            "administrative nor a health facility, so no formula was written", _
                            checkingError
                    Exit Sub
                End If
                colVar = specs.SpatialConcatVariableName(colVar)
            End If

            ' Time period formulas (only for new sections)
            If specs.IsNewSection() Then
                Set rng = .Range(tabNames.RowCategories)
                Set cellRng = rng.Cells(1, 1)
                If WriteFormula(cellRng, _
                                "FormatDateFromLastDay(" & tabNames.TimeUnit & "," & _
                                cellRng.Cells(1, -1).Address(RowAbsolute:=False) & "," & _
                                cellRng.Cells(1, 0).Address(RowAbsolute:=False) & ", " & _
                                tabNames.EndDate & ")", False) Then
                    FillDown cellRng, rng
                End If

                ' Start period column
                Set cellRng = .Range(tabNames.FirstValueStartTime)
                WriteFormula cellRng, tabNames.StartDate, False
                WriteFormula cellRng.Cells(2, 1), _
                             cellRng.Cells(1, 2).Address(RowAbsolute:=False) & " + 1", False
                Set rng = .Range(tabNames.StartTimePeriod)
                FillDown rng.Cells(1, 1), rng

                ' End period column
                Set rng = .Range(tabNames.EndTimePeriod)
                Set cellRng = rng.Cells(1, 1)
                If WriteFormula(cellRng, _
                                "FindLastDay(" & tabNames.TimeUnit & "," & _
                                cellRng.Cells(1, 0).Address(RowAbsolute:=False) & ")", _
                                False) Then
                    FillDown cellRng, rng
                End If

                ' Total period
                Set rng = .Range(tabNames.TotalRow)
                WriteFormula rng.Cells(1, -1), tabNames.StartDate, False
                WriteFormula rng.Cells(1, 0), tabNames.EndDate, False

                ' Validation max/min
                WriteFormula .Range(tabNames.ValidationMaxDate), _
                             "ValidMax(" & tabNames.UserStartDate & ", " & _
                             tabNames.UserEndDate & ", " & tabNames.MinMinDate & ", " & _
                             tabNames.MaxMaxDate & ", " & tabNames.TimeUnit & ")", False

                WriteFormula .Range(tabNames.ValidationMinDate), _
                             "ValidMin(" & tabNames.UserStartDate & ", " & _
                             tabNames.UserEndDate & ", " & tabNames.MinMinDate & ", " & _
                             tabNames.MaxMaxDate & ", " & tabNames.TimeUnit & ")", False

                ' Start/End dates
                WriteFormula .Range(tabNames.StartDate), _
                             "MAX(" & tabNames.UserStartDate & ", " & _
                             tabNames.ValidationMinDate & ")", False

                WriteFormula .Range(tabNames.EndDate), _
                             "MIN(" & tabNames.UserEndDate & ", " & _
                             tabNames.ValidationMaxDate & ")", False

                ' Info cells
                WriteFormula .Range(tabNames.InfoStartDate), _
                             "InfoUser(" & tabNames.UserStartDate & ", " & _
                             tabNames.StartDate & ")", False

                WriteFormula .Range(tabNames.InfoEndDate), _
                             "InfoUser(" & tabNames.UserEndDate & ", " & _
                             tabNames.EndDate & ", 2)", False

                WriteFormula .Range(tabNames.InfoAnaPeriod), _
                             "RNG_AnaPeriod & PLAGE_VALUE(" & tabNames.StartDate & "," & _
                             tabNames.EndDate & ")", False

                ' Time unit validation. The spatio-temporal tables are written to
                ' a worksheet of their own and carry a time unit list of their
                ' own, so the scope decides which list the dropdown reads.
                AddListValidation tabNames.TimeUnit, _
                                  AnalysisRanges.TimeUnitListOfScope( _
                                      tabType = ScopeSpatioTemporal)
            End If

            ' The section anchor carries the period cells every table of the
            ' section reads. A section whose anchor row failed validation was
            ' skipped by the builder, so the names were never created and
            ' reading one raises several frames away from the cause.
            If Not RangeExists(secNames.FirstValueStartTime) Then
                LogInfo "The section anchor " & secId & " has no start period cell, so " & _
                        "the formulas of table " & tabId & " were skipped", checkingError
                Exit Sub
            End If

            ' Column formulas
            Set rng = .Range(tabNames.ValuesCol(1))
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1))
            Set rngLab = .Range(tabNames.LabelCol(1))
            Set totRng = .Range(tabNames.TotalRowValues)
            totalRowNumber = totRng.Row
            Set cellRng = .Range(secNames.FirstValueStartTime)
            Set misRng = ResolvedRange(tabNames.MissingRowValues)
            Set perRng = cellRng.Cells(1, 3)
            startTimeAddress = cellRng.Address(RowAbsolute:=False)
            endTimeAddress = cellRng.Cells(1, 2).Address(RowAbsolute:=False)

            If hasMissing And misRng Is Nothing Then
                LogInfo "Table " & tabId & " asks for a missing row and has none", checkingError
            End If

            For colCounter = 1 To nbDataColumns

                labAddress = rngLab.Address

                ClearTableConditions
                AddTableConditions rowVar, ">=" & startTimeAddress
                AddTableConditions rowVar, "<=" & endTimeAddress
                AddTableConditions colVar, "=" & labAddress
                finalFormula = CurrentExcelFormula()

                If WriteFormula(rng.Cells(1, 1), Condition(perRng, finalFormula)) Then _
                    FillDown rng.Cells(1, 1), rng
                AddRngFormat tabFormat, rng

                ' Missing row. The time variable is empty on those records, so
                ' the two date bounds are left off on purpose: a record with no
                ' date sits in no period.
                If Not misRng Is Nothing Then
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & EmpStr()
                    AddTableConditions colVar, "=" & labAddress
                    Set cellRng = .Cells(misRng.Row, rng.Column)
                    If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                        AddRngFormat tabFormat, cellRng
                End If

                ' Percentage
                If hasPercentage Then
                    Set denomRng = PercentageDenominator(tabNames, percentageValue, _
                                                         totalRowNumber, rng.Column)
                    WritePercentageTwin rng, denomRng, _
                                        rng.Cells(rng.Rows.Count, 2), _
                                        PercentageRowAbsolute(percentageValue)
                End If

                ' Spatio-temporal label formulas
                If tabType = ScopeSpatioTemporal Then
                    If RangeExists(secNames.SpatialInput(colCounter, spTabType)) Then
                        WriteFormula rngLab, _
                                     "IF(ISBLANK(" & _
                                     secNames.SpatialInput(colCounter, spTabType) & ")," & _
                                     EmpStr() & "," & _
                                     secNames.SpatialInput(colCounter, spTabType) & ")", _
                                     False
                    Else
                        LogInfo "Input cell " & _
                                secNames.SpatialInput(colCounter, spTabType) & _
                                " is missing, so column " & colCounter & " of table " & _
                                tabId & " has no header formula", checkingError
                    End If
                End If

                Set rng = rng.Offset(, columnStride)
                Set rngLab = rngLab.Offset(, columnStride)
            Next

            ' Missing column and total columns
            If hasMissing Then
                Set misRng = ResolvedRange(tabNames.MissingColValues)

                If misRng Is Nothing Then
                    LogInfo "Table " & tabId & " asks for a missing column and has none", _
                            checkingError
                Else
                    Set misRng = .Range(misRng.Cells(1, 1), misRng.Cells(misRng.Rows.Count + 1, 1))

                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    AddTableConditions colVar, "=" & EmpStr()
                    finalFormula = CurrentExcelFormula()

                    If WriteFormula(misRng.Cells(1, 1), Condition(perRng, finalFormula)) Then _
                        FillDown misRng.Cells(1, 1), misRng
                    AddRngFormat tabFormat, misRng

                    If hasPercentage Then
                        Set denomRng = PercentageDenominator(tabNames, percentageValue, _
                                                             totalRowNumber, misRng.Column)
                        WritePercentageTwin misRng, denomRng, _
                                            misRng.Cells(misRng.Rows.Count, 2), _
                                            PercentageRowAbsolute(percentageValue)
                    End If
                End If

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range(tabNames.TotalColValues)
                    If WriteFormula(totRng.Cells(1, 1), Condition(perRng, finalFormula)) Then _
                        FillDown totRng.Cells(1, 1), totRng
                    AddRngFormat tabFormat, totRng
                End If
            Else
                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    If colVar <> vbNullString Then AddNonEmptyTableConditions colVar
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range(tabNames.TotalColValues)
                    If WriteFormula(totRng.Cells(1, 1), Condition(perRng, finalFormula)) Then _
                        FillDown totRng.Cells(1, 1), totRng
                    AddRngFormat tabFormat, totRng
                End If
            End If

            If hasTotal And hasPercentage Then
                Set totRng = ResolvedRange(tabNames.TotalColValues)
                If Not totRng Is Nothing Then
                    Set denomRng = PercentageDenominator(tabNames, percentageValue, _
                                                         totalRowNumber, totRng.Column)
                    WritePercentageTwin totRng, denomRng, _
                                        totRng.Cells(totRng.Rows.Count, 2), _
                                        PercentageRowAbsolute(percentageValue)
                End If
            End If

            ' Missing row, total column
            Set cellRng = ResolvedRange(tabNames.MissingTotal)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Missing row, missing column
            Set cellRng = ResolvedRange(tabNames.MissingMissing)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Total row, missing column
            Set cellRng = ResolvedRange(tabNames.TotalMissing)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions colVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Total row, total column
            Set cellRng = ResolvedRange(tabNames.TotalTotal)
            If Not cellRng Is Nothing Then
                If IsCountFunction() Then
                    WriteFormula cellRng, _
                                 "SUM(" & .Range(tabNames.TotalColValues).Address & ")", False
                Else
                    Set anchorRng = ResolvedRange(secNames.TotalRow)
                    If anchorRng Is Nothing Then
                        LogInfo "The section anchor " & secId & " has no total row, so the " & _
                                "grand total of table " & tabId & " was not written", _
                                checkingError
                    Else
                        Set anchorRng = anchorRng.Cells(1, 0)
                        ClearTableConditions
                        AddTableConditions rowVar, "<=" & anchorRng.Address
                        AddTableConditions rowVar, ">=" & anchorRng.Cells(1, 0).Address
                        If Not hasMissing Then AddNonEmptyTableConditions colVar
                        WriteFormula cellRng, CurrentExcelFormula()
                    End If
                End If
                AddRngFormat tabFormat, cellRng
            End If

        Case ScopeSpatial

            spTabType = specs.SpatialTableScopes()

            ' Neither prefix answered, so there is no variable to summarise and
            ' no dropdown to build. The order column used to be filled across
            ' whatever range the geographic branch would have used.
            If LenB(spTabType) = 0 Then
                LogInfo "Table " & tabId & " groups by a variable that is neither " & _
                        "administrative nor a health facility, so no formula was written", _
                        checkingError
                Exit Sub
            End If

            rowVar = specs.SpatialConcatVariableName(specs.Value("row"))
            colVar = specs.Value("column")

            ' Validation dropdowns
            If spTabType = "geo" Then
                AddListValidation tabNames.AdmDropdown, AnalysisRanges.AdmUnitList
                AddListValidation tabNames.DevidePop, AnalysisRanges.PopulationFactorList
            End If

            ' Order formula
            Set rng = .Range(tabNames.RowCategories).Offset(, -2)
            rng.Cells(1, 1).Value = 1
            If rng.Rows.Count > 1 Then
                Set cellRng = rng.Cells(2, 1)
                If WriteFormula(cellRng, _
                                cellRng.Offset(-1).Address(RowAbsolute:=False) & " + 1", _
                                False) Then
                    Set rng = .Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
                    FillDown cellRng, rng
                End If
            End If

            Set rng = .Range(tabNames.RowCategories)
            rowCatAddress = rng.Cells(1, 1).Address(RowAbsolute:=False)

            If spTabType = "geo" Then
                ' Population formula
                Set popRng = rng.Offset(, -1)
                Set cellRng = popRng.Cells(1, 1)
                If WriteFormula(cellRng, _
                                "FindTopPop(" & tabNames.AdmDropdown & "," & _
                                cellRng.Cells(1, 0).Address(RowAbsolute:=False) & "," & _
                                Chr(34) & rowVar & Chr(34) & "," & _
                                Chr(34) & tabId & Chr(34) & ")", False) Then
                    FillDown cellRng, popRng
                End If

                ' Top Admin formula
                Set cellRng = rng.Cells(1, 1)
                If WriteFormula(cellRng, _
                                "FindTopAdmin(" & tabNames.AdmDropdown & "," & _
                                cellRng.Cells(1, -1).Address(RowAbsolute:=False) & "," & _
                                Chr(34) & rowVar & Chr(34) & "," & _
                                Chr(34) & tabId & Chr(34) & ")", False) Then
                    FillDown cellRng, rng
                End If
            Else
                Set cellRng = rng.Cells(1, 1)
                If WriteFormula(cellRng, _
                                "FindTopHF(" & _
                                cellRng.Cells(1, 0).Address(RowAbsolute:=False) & "," & _
                                Chr(34) & rowVar & Chr(34) & "," & _
                                Chr(34) & tabId & Chr(34) & ")", False) Then
                    FillDown cellRng, rng
                End If
            End If

            ' Column value formulas (similar to bivariate)
            Set rng = .Range(tabNames.ValuesCol(1))
            Set rngLab = .Range(tabNames.LabelCol(1))
            Set misRng = ResolvedRange(tabNames.MissingRowValues)

            For colCounter = 1 To nbDataColumns

                labAddress = rngLab.Address

                ClearTableConditions
                AddTableConditions rowVar, "=" & rowCatAddress
                AddNonEmptyTableConditions rowVar
                AddTableConditions colVar, "=" & labAddress
                finalFormula = CurrentExcelFormula()

                If WriteFormula(rng.Cells(1, 1), finalFormula) Then FillDown rng.Cells(1, 1), rng
                AddRngFormat tabFormat, rng

                ' Missing row
                If Not misRng Is Nothing Then
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & EmpStr()
                    AddTableConditions colVar, "=" & labAddress
                    WriteFormula .Cells(misRng.Row, rng.Column), CurrentExcelFormula()
                End If

                ' Percentage. A spatial table offers the row percentage alone.
                If hasPercentage Then
                    Set denomRng = PercentageDenominator(tabNames, "row", 0, rng.Column)
                    WritePercentageTwin rng, denomRng, rng.Cells(rng.Rows.Count, 2), False
                End If

                Set rng = rng.Offset(, columnStride)
                Set rngLab = rngLab.Offset(, columnStride)
            Next

            ' Missing column and total columns
            If hasMissing Then
                Set misRng = ResolvedRange(tabNames.MissingColValues)

                If misRng Is Nothing Then
                    LogInfo "Table " & tabId & " asks for a missing column and has none", _
                            checkingError
                Else
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & rowCatAddress
                    AddTableConditions colVar, "=" & EmpStr()
                    finalFormula = CurrentExcelFormula()

                    If WriteFormula(misRng.Cells(1, 1), finalFormula) Then _
                        FillDown misRng.Cells(1, 1), misRng
                    AddRngFormat tabFormat, misRng

                    If hasPercentage Then
                        Set denomRng = PercentageDenominator(tabNames, "row", 0, misRng.Column)
                        WritePercentageTwin misRng, denomRng, _
                                            misRng.Cells(misRng.Rows.Count, 2), False
                    End If
                End If

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & rowCatAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range(tabNames.TotalColValues)
                    If WriteFormula(totRng.Cells(1, 1), finalFormula) Then _
                        FillDown totRng.Cells(1, 1), totRng
                    AddRngFormat tabFormat, totRng
                End If
            Else
                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & rowCatAddress
                    AddNonEmptyTableConditions colVar
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range(tabNames.TotalColValues)
                    If WriteFormula(totRng.Cells(1, 1), finalFormula) Then _
                        FillDown totRng.Cells(1, 1), totRng
                    AddRngFormat tabFormat, totRng
                End If
            End If

            ' Missing row, total column
            Set cellRng = ResolvedRange(tabNames.MissingTotal)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Missing row, missing column
            Set cellRng = ResolvedRange(tabNames.MissingMissing)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                AddTableConditions colVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Total row, missing column
            Set cellRng = ResolvedRange(tabNames.TotalMissing)
            If Not cellRng Is Nothing Then
                ClearTableConditions
                AddTableConditions colVar, "=" & EmpStr()
                If WriteFormula(cellRng, CurrentExcelFormula()) Then _
                    AddRngFormat tabFormat, cellRng
            End If

            ' Total row, total column
            Set cellRng = ResolvedRange(tabNames.TotalTotal)
            If Not cellRng Is Nothing Then
                If IsCountFunction() Then
                    WriteFormula cellRng, _
                                 "SUM(" & .Range(tabNames.TotalColValues).Address & ")", False
                Else
                    Select Case missingValue
                    Case "all"
                        WriteFormula cellRng, CurrentExcelFormula(scope:=FormulaScopeLinelist)
                    Case "row"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        WriteFormula cellRng, CurrentExcelFormula()
                    Case "column"
                        ClearTableConditions
                        AddNonEmptyTableConditions rowVar
                        WriteFormula cellRng, CurrentExcelFormula()
                    Case "no"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        AddNonEmptyTableConditions rowVar
                        WriteFormula cellRng, CurrentExcelFormula()
                    Case Else
                        ' A missing setting outside the four leaves the cell
                        ' where the total row meets the total column with no
                        ' formula, and the table still reads as complete.
                        LogInfo "Table " & ID() & " asks for the missing setting '" & _
                                missingValue & "', which this class does not write", _
                                checkingWarning
                    End Select
                End If
                AddRngFormat tabFormat, cellRng
            End If

            ' Write the spatial tables the generated linelist refreshes, and
            ' carry what that reported into this report. Nothing harvested from
            ' it before, so a spatial table with a blank value column reached
            ' the delivered file with the analysis phase reading clean.
            Set spTab = SpatialTables.Create(tabl)
            spTab.Add TableFormula

            If spTab.HasCheckings Then
                Set spatialChecks = spTab.CheckingValues
                Set tableChecks = Checks
                tableChecks.Append spatialChecks
            End If

        Case Else
            ' A scope outside the five wrote no formula and said nothing, so the
            ' table reached the delivered file empty and the report read clean.
            LogInfo "Table " & tabId & " has a scope this class writes no formulas for", _
                    checkingWarning

        End Select
    End With
End Sub

'@section Checkings
'===============================================================================
'@description What the class reports back to the generation report. The report
'used to show the analysis phase clean while a table carried a cell Excel had
'refused. The formula builder and the spatial table builder file into the same
'entries, so one harvest carries everything written for the table.

'@label:checks
'@prop-title The report entries of this instance.
'@return Checking. The entries, created on first use.
Private Property Get Checks() As Checking
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(CLASS_NAME)
    Set Checks = internalChecks
End Property

'@label:log-info
'@sub-title File one diagnostic entry.
'@details
'Checking.Add raises on a duplicate key, and AnalysisOutput pours the entries of
'several classes over several tables into one report, so the key names the class,
'the table and a counter. A bare counter collided with the cross-table first
'entry on the same table.
'@param message String. The diagnostic message text.
'@param scope Optional Byte. The checking scope (default checkingSuccess).
Private Sub LogInfo(ByVal message As String, _
                    Optional ByVal scope As Byte = checkingSuccess)
    Dim entries As Checking

    Set entries = Checks
    checkCounter = checkCounter + 1
    entries.Add CLASS_NAME & "-" & ID() & "-" & CStr(checkCounter), message, scope
End Sub

'@label:has-checkings
'@prop-title Whether this instance has anything to report.
'@return Boolean. True once one entry has been filed.
'@export
Public Property Get HasCheckings() As Boolean
    If internalChecks Is Nothing Then Exit Property
    HasCheckings = (internalChecks.Length() > 0)
End Property

'@label:checking-values
'@prop-title The entries filed while the formulas were written.
'@details
'AnalysisOutput harvests these after AddFormulas, the same way it harvests from
'the cross-table.
'@return Checking. The entries, or Nothing when none were filed.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property

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
