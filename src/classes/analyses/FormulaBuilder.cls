VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds analysis formulas and writes them into cells"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Builds analysis formulas and writes them into cells")
'@IgnoreModule UnrecognizedAnnotation, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class FormulaBuilder
'@description
'Holds the criteria of one analysis formula, asks the parser for the formula
'text and writes it into a cell. AddCondition adds one criterion,
'ConditionCount answers how many are held and ClearConditions drops them.
'ExcelFormula gives the text, WriteFormula puts it in a cell, FillDown copies
'it down a range and ApplyFormat sets the number format. Percentage and
'Condition build the two smaller formulas, and EmptyText and NotEmptyText give
'the strings an empty cell shows. CrossTableFormula and SpatialTables both
'build their formulas through one of these.
'
'THE CRITERIA
'-------------------------------------------------------------------------------
'A formula is built from the criteria added since the last one was built. The
'first AddCondition after a build starts a new set. ClearConditions empties the
'set at any point.
'
'EXCEL IS THE JUDGE OF A FORMULA
'-------------------------------------------------------------------------------
'WriteFormula assigns the formula and traps what Excel raises. A refused
'formula leaves the cell empty and files a check naming the cell.
'
'THE 255 CHARACTER LIMIT ON FORMULAARRAY
'-------------------------------------------------------------------------------
'Range.FormulaArray refuses a formula longer than 255 characters, and an
'analysis formula over three criteria passes that on ordinary variable names.
'Those go in through a short stub formula and Range.Replace, which carries no
'such limit and keeps the cell an array formula. When the host refuses that
'too, the formula is entered as an ordinary formula and a warning is filed.
'
'WHERE THE CHECKS GO
'-------------------------------------------------------------------------------
'The owning class hands its own Checking to Create, so everything filed here
'reaches the generation report through that owner. A key carries the owner
'name, the table and an F. Checking.Add raises on a duplicate key.
'@depends Formulas, FormulaCondition, Checking, BetterArray
'@version 2026-07-30

'@section Enumerations
'===============================================================================
'@description
'Target context for a generated formula. An analysis formula carries the criteria
'of its cell; a linelist formula reads the whole column. The enum lived in
'CrossTableFormula. That class and SpatialTables both read it from here now.

Public Enum ExcelFormulaScope
    FormulaScopeLinelist = 1
    FormulaScopeAnalysis = 2
End Enum

'@section State
'===============================================================================

Private Type TFormulaBuilder
    ownerName As String
    contextId As String
    condVars As BetterArray
    condTests As BetterArray
    checks As Checking
    formulaBuilt As Boolean
    sealed As Boolean
End Type

Private this As TFormulaBuilder
Private checkCounter As Long

Private Const CLASS_NAME As String = "FormulaBuilder"

' Range.FormulaArray refuses a longer formula. See the class description.
Private Const FORMULA_ARRAY_LIMIT As Long = 255

' The stub a long array formula is entered through. It is a name no workbook
' defines, so the cell holds a name error for as long as the swap takes, and it
' cannot appear inside a generated formula.
Private Const LONG_FORMULA_STUB As String = "OBTLONGFORMULASTUB"

'@section Instantiation
'===============================================================================

'@label:create
'@sub-title Create a builder for one owner and one table.
'@details
'Factory method on the predeclared instance. The owner name and the table
'identifier travel into every message and every checking key, so a reader of the
'generation report knows which class wrote which cell.
'@param ownerName String. The class the builder works for.
'@param contextId String. The table identifier the formulas belong to.
'@param checks Checking. The report entries are filed into this.
'@return FormulaBuilder. A builder with an empty criteria set.
'@throws ProjectError.ObjectNotInitialized When checks is Nothing.
'@export
Public Function Create(ByVal ownerName As String, ByVal contextId As String, _
                       ByVal checks As Checking) As FormulaBuilder
    If checks Is Nothing Then _
        ThrowError ProjectError.ObjectNotInitialized, _
                   "No checking provided for " & ownerName & " formulas"

    Dim instance As FormulaBuilder
    Set instance = New FormulaBuilder
    instance.OwnerName = ownerName
    instance.ContextId = contextId
    Set instance.Checkings = checks
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

'@label:owner-name
'@prop-title The class this builder works for.
'@return String. The owner class name.
'@export
Public Property Get OwnerName() As String
    OwnerName = this.ownerName
End Property

'@label:owner-name-set
'@prop-title Assign the owner class name.
'@param ownerValue String. The owner class name.
'@export
Public Property Let OwnerName(ByVal ownerValue As String)
    GuardNotSealed "OwnerName"
    this.ownerName = ownerValue
End Property

'@label:context-id
'@prop-title The table identifier the formulas belong to.
'@return String. The table identifier.
'@export
Public Property Get ContextId() As String
    ContextId = this.contextId
End Property

'@label:context-id-set
'@prop-title Assign the table identifier.
'@param idValue String. The table identifier.
'@export
Public Property Let ContextId(ByVal idValue As String)
    GuardNotSealed "ContextId"
    this.contextId = idValue
End Property

'@label:checkings
'@prop-title The report entries are filed into this.
'@return Checking. The owner report.
'@export
Public Property Get Checkings() As Checking
    Set Checkings = this.checks
End Property

'@label:checkings-set
'@prop-title Assign the report the entries are filed into.
'@param checks Checking. The owner report.
'@export
Public Property Set Checkings(ByVal checks As Checking)
    GuardNotSealed "Checkings"
    Set this.checks = checks
End Property

'@section Criteria
'===============================================================================
'@description The criteria of the formula being built. condVars and condTests
'are two arrays of the same length, read in step by FormulaCondition.

'@label:add-condition
'@sub-title Add one criterion to the formula being built.
'@details
'The first criterion added after a formula was built starts a new set. An empty
'variable name is ignored, which is how a caller passes an absent column.
'@param varName String. The linelist variable the criterion reads.
'@param conditionName String. The criteria text, such as "= $A3" or "<>""".
'@export
Public Sub AddCondition(ByVal varName As String, ByVal conditionName As String)
    If this.formulaBuilt Then ClearConditions

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

'@label:clear-conditions
'@sub-title Empty the criteria set.
'@export
Public Sub ClearConditions()
    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
    End If

    this.condVars.Clear
    this.condTests.Clear
    this.formulaBuilt = False
End Sub

'@label:condition-count
'@prop-title How many criteria the current formula carries.
'@return Long. The number of criteria added since the last build.
'@export
Public Property Get ConditionCount() As Long
    If this.condVars Is Nothing Then Exit Property
    ConditionCount = this.condVars.Length
End Property

'@section Formula Text
'===============================================================================

'@label:empty-text
'@prop-title The Excel literal for an empty string.
'@details
'Two double quote characters. Written with Chr(34) so the VBA source carries no
'doubled quotes of its own.
'@return String. The two characters.
'@export
Public Property Get EmptyText() As String
    EmptyText = Chr(34) & Chr(34)
End Property

'@label:not-empty-text
'@prop-title The COUNTIFS criteria for a cell that holds something.
'@details
'The two characters "<>" inside a pair of double quotes. COUNTIFS reads that as
'"not blank", which is what a count summary function needs.
'@return String. The quoted operator.
'@export
Public Property Get NotEmptyText() As String
    NotEmptyText = Chr(34) & "<>" & Chr(34)
End Property

'@label:percentage
'@sub-title Build a percentage guarded against a zero denominator.
'@details
'The shape is IF(ISERR(num/denom), "", num/denom), so an empty denominator shows
'a blank cell. The denominator column is held absolute, so the formula survives
'a copy down the rows.
'@param denomRng Range. The cell the division divides by.
'@param numRng Range. The cell the division divides.
'@param rowabs Optional Boolean. True to hold the denominator row fixed. Defaults to True.
'@return String. The formula text with no leading "=".
'@export
Public Function Percentage(ByVal denomRng As Range, ByVal numRng As Range, _
                           Optional ByVal rowabs As Boolean = True) As String
    Dim parsedForm As String

    parsedForm = numRng.Address(RowAbsolute:=False) & "/" & _
                 denomRng.Address(RowAbsolute:=rowabs, ColumnAbsolute:=True)
    Percentage = "IF(ISERR(" & parsedForm & ")," & EmptyText & "," & parsedForm & ")"
End Function

'@label:condition
'@sub-title Blank a formula when its guard cell is empty.
'@details
'The shape is IF(guard = "", "", formula). A time period row the user date range
'leaves unused has an empty label cell, and every value beside it reads blank.
'@param rngCondition Range. The guard cell.
'@param formulaValue String. The formula text to keep when the guard holds something.
'@param rowabs Optional Boolean. True to hold the guard row fixed. Defaults to False.
'@return String. The guarded formula text, empty when formulaValue is empty.
'@export
Public Function Condition(ByVal rngCondition As Range, ByVal formulaValue As String, _
                          Optional ByVal rowabs As Boolean = False) As String
    If LenB(formulaValue) = 0 Then Exit Function
    Condition = "IF(" & rngCondition.Address(RowAbsolute:=rowabs) & "=" & _
                EmptyText & "," & EmptyText & "," & formulaValue & ")"
End Function

'@label:excel-formula
'@sub-title Ask the parser for the formula the current criteria describe.
'@details
'The analysis scope wraps the criteria into a FormulaCondition and asks for a
'COUNTIFS or SUMPRODUCT formula. The linelist scope reads the whole column and
'ignores the criteria. The text comes back as the parser built it, and Excel
'judges it when WriteFormula assigns it.
'@param formObject Formulas. The parsed summary function.
'@param scope Optional Byte. FormulaScopeAnalysis or FormulaScopeLinelist. Defaults to the analysis scope.
'@param prefix Optional String. The table name prefix of the linelist reference. Defaults to "f".
'@return String. The formula text with no leading "=".
'@export
Public Property Get ExcelFormula(ByVal formObject As Formulas, _
                                 Optional ByVal scope As Byte = FormulaScopeAnalysis, _
                                 Optional ByVal prefix As String = "f") As String
    Dim cond As FormulaCondition

    If formObject Is Nothing Then Exit Property

    If scope = FormulaScopeLinelist Then
        ExcelFormula = formObject.ParsedLinelistFormula(useTableName:=True, tablePrefix:=prefix)
        this.formulaBuilt = True
        Exit Property
    End If

    If this.condVars Is Nothing Or this.condTests Is Nothing Then Exit Property

    Set cond = FormulaCondition.Create(this.condVars, this.condTests)
    ExcelFormula = formObject.ParsedAnalysisFormula(cond, tablePrefix:=prefix)

    'A condition that refuses a variable, a table or an empty list says so
    'in a store of its own. Nobody read it, so a table whose condition was
    'dropped reached the delivered file with the report clean.
    HarvestConditionCheckings cond

    this.formulaBuilt = True
End Property

'@label:harvest-condition-checkings
'@sub-title Take what the condition reported into the owner report
'@details
'The keys are rewritten under this builder's own shape, because
'`FormulaCondition` counts from 0 and one owner report holds the entries
'of many conditions.
'@param cond FormulaCondition. The condition just built.
Private Sub HarvestConditionCheckings(ByVal cond As FormulaCondition)
    Dim sourceKeys As BetterArray
    Dim counter As Long
    Dim sourceKey As String

    If cond Is Nothing Then Exit Sub
    If Not cond.HasCheckings Then Exit Sub
    If this.checks Is Nothing Then Exit Sub

    Set sourceKeys = cond.CheckingValues.ListOfKeys()
    If sourceKeys Is Nothing Then Exit Sub

    For counter = sourceKeys.LowerBound To sourceKeys.UpperBound
        sourceKey = CStr(sourceKeys.Item(counter))
        LogInfo cond.CheckingValues.ValueOf(sourceKey, checkingLabel), checkingError
    Next counter
End Sub

'@section Writing Formulas
'===============================================================================
'@description The one way a formula reaches a cell. Excel accepts it or raises,
'and a raise is reported through the owner checkings.

'@label:write-formula
'@sub-title Write one formula into one cell and report what Excel says.
'@details
'Assigns the formula and traps the assignment. Excel rejects a malformed formula
'with error 1004 at assignment time, which makes it the only judge needed. A
'refused formula leaves the cell empty and files a check carrying the address and
'the message Excel gave.
'@param target Range. The cell the formula belongs in.
'@param frm String. The formula text with no leading "=".
'@param asArray Optional Boolean. True to enter the formula as an array formula. Defaults to True.
'@return Boolean. True when the formula reached the cell.
'@export
Public Function WriteFormula(ByVal target As Range, ByVal frm As String, _
                             Optional ByVal asArray As Boolean = True) As Boolean
    Dim failureReason As String

    'The address is read only when a message carries it. Reading it up front
    'cost one host call per written formula on the path that never fails.
    If LenB(frm) = 0 Then
        LogInfo "No formula could be built for cell " & _
                target.Address(RowAbsolute:=False, ColumnAbsolute:=False) & _
                " of table " & this.contextId, checkingError
        Exit Function
    End If

    On Error GoTo WriteFailed

    If asArray Then
        If Len(frm) > FORMULA_ARRAY_LIMIT Then
            WriteLongArrayFormula target, frm
        Else
            target.FormulaArray = "=" & frm
        End If
    Else
        target.Formula = "=" & frm
    End If

    On Error GoTo 0
    WriteFormula = True
    Exit Function

WriteCleanup:
    On Error Resume Next
    target.ClearContents
    LogInfo "Excel refused the formula of cell " & _
            target.Address(RowAbsolute:=False, ColumnAbsolute:=False) & " of table " & _
            this.contextId & ": " & failureReason, checkingError
    On Error GoTo 0
    Exit Function

WriteFailed:
    failureReason = Err.Description
    Resume WriteCleanup
End Function

'@label:write-long-array-formula
'@sub-title Enter an array formula longer than 255 characters.
'@details
'Range.FormulaArray refuses the assignment, so a short stub goes in first and
'Range.Replace swaps the real text in. Replace edits the cell the way a user
'would and carries no length limit, and the cell stays an array formula.
'Range.Replace is called as a statement and the cell is read back to see whether
'the swap happened. Its return value is a Variant that this host does not always
'fill, and testing it with Not raised a type mismatch that read as Excel refusing
'the formula.
'
'When the host refuses the swap as well, the formula goes in as an ordinary
'formula and a warning is filed carrying the step that refused it. Excel with
'dynamic arrays gives the same answer either way; an older Excel returns the
'first value of the range where the aggregate belongs, so the warning is the
'reader only sign of it.
'@param target Range. The cell the formula belongs in.
'@param frm String. The formula text with no leading "=".
Private Sub WriteLongArrayFormula(ByVal target As Range, ByVal frm As String)
    Dim stage As String
    Dim refusal As String

    On Error GoTo ArrayEntryFailed

    stage = "stub"
    target.FormulaArray = "=" & LONG_FORMULA_STUB

    stage = "swap"
    target.Replace What:=LONG_FORMULA_STUB, Replacement:=frm, _
                   LookAt:=xlPart, MatchCase:=True

    stage = "read back"
    If InStr(1, CStr(target.Formula), LONG_FORMULA_STUB, vbBinaryCompare) > 0 Then
        ThrowError ProjectError.SomethingWentWrong, "the stub survived the swap"
    End If

    On Error GoTo 0
    Exit Sub

ArrayEntryPlainFormula:
    On Error GoTo 0
    LogInfo "The formula of cell " & _
            target.Address(RowAbsolute:=False, ColumnAbsolute:=False) & " of table " & _
            this.contextId & " is " & Len(frm) & " characters, and the host refused an " & _
            "array entry at [" & stage & "] with " & refusal & _
            ". It was entered as an ordinary formula.", checkingWarning
    target.Formula = "=" & frm
    Exit Sub

ArrayEntryFailed:
    refusal = "(" & Err.Number & ") " & Err.Description
    Resume ArrayEntryPlainFormula
End Sub

'@label:fill-down
'@sub-title Copy the formula of one cell over a range.
'@details
'AutoFill needs a destination taller than one cell and raises 1004 otherwise. A
'value column is one cell tall whenever the table has a single category row,
'which a univariate table with one category and a spatial table with a geo count
'of one both produce, so the guard is on the ordinary path.
'@param sourceCell Range. The cell holding the formula to copy.
'@param destination Range. The range to fill, including the source cell.
'@export
Public Sub FillDown(ByVal sourceCell As Range, ByVal destination As Range)
    If destination.Cells.Count <= 1 Then Exit Sub

    On Error GoTo FillFailed
    sourceCell.AutoFill Destination:=destination, Type:=xlFillValues
    On Error GoTo 0
    Exit Sub

FillCleanup:
    On Error Resume Next
    LogInfo "The formula of cell " & _
            sourceCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & _
            " could not be copied over " & _
            destination.Address(RowAbsolute:=False, ColumnAbsolute:=False) & _
            " of table " & this.contextId, checkingError
    On Error GoTo 0
    Exit Sub

FillFailed:
    Resume FillCleanup
End Sub

'@section Number Format
'===============================================================================

'@label:apply-format
'@sub-title Apply a number format to a range from a setup keyword.
'@details
'Translates the keywords the setup workbook offers into Excel number formats:
'"integer", "round0", "roundN", "percentage0", "percentageN", "euros", "dollars"
'and "text". Anything else is handed to Excel as a number format of its own,
'because the setup sheet allows free text there.
'@param formatValue String. The keyword or an Excel number format.
'@param rng Range. The range to format.
'@export
Public Sub ApplyFormat(ByVal formatValue As String, ByVal rng As Range)
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

'@section Checkings
'===============================================================================

'@label:log-info
'@sub-title File one entry into the owner report.
'@details
'Checking.Add raises on a duplicate key, so the key carries the owner name, the
'table and a counter. The F tells a reader the entry came from the builder rather
'than from the owner itself, and it is what lets two owners file into one report.
'@param message String. The message text.
'@param scope Optional Byte. The checking scope. Defaults to checkingSuccess.
Private Sub LogInfo(ByVal message As String, _
                    Optional ByVal scope As Byte = checkingSuccess)
    checkCounter = checkCounter + 1
    this.checks.Add this.ownerName & "-" & this.contextId & "-F" & CStr(checkCounter), _
                    message, scope
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a project error with a standardised source name.
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError. Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
