VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Variables of the dictionary"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Variables of the dictionary")
'@depends BetterArray, Checking, DataSheet, LLdictionary
'@IgnoreModule PropertyNotUsed, ReadOnlyPropertyAssignment, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'This class answers questions about ONE variable of the dictionary: the row it
'sits on, the value of any column on that row, and the small set of named
'fields the rest of the tree reads most (sheet name, control type, table name).
'
'Speed rule for this file: no worksheet search. Thirty production sites build
'an LLVariables, and `Formulas.ParsedLinelistFormula` asks `Contains` then
'`Value` for every token of every formula in the workbook. Each of those used
'to be a `Range.Find` plus a column lookup. The variable-name column and the
'header row are each read ONCE as a block, and every lookup after that is a
'string compare in memory.
'
'What the two snapshots promise:
'
'- The variable names are held until `InvalidateCaches`. That is the same
'  contract the captured search range already had: `LLdictionary.Prepare`
'  inserts geo rows, so an object built before preparation cannot see them.
'  `EventSetup` is the one production caller that drops them.
'- The header row is held, and a column that the snapshot cannot answer makes
'  the header row be read again before the answer is "missing". A column added
'  after this object was built therefore resolves. `LinelistSpecs.AddListAuto`
'  needs exactly that: it adds "list auto" to the dictionary and then writes to
'  it through a variables object it built earlier.
'- A column the snapshot CAN answer is checked with one cell read before the
'  index is handed out, so a header renamed on the worksheet can never make
'  this class read the wrong column.

Private Type TVariable
    dict As LLdictionary
    hostWksh As Worksheet
    varRng As Range
    variableNames() As String
    variableCount As Long
    firstVariableRow As Long
    namesLoaded As Boolean
    headers() As String
    headerCount As Long
    firstHeaderColumn As Long
    headersLoaded As Boolean
    sealed As Boolean
End Type

Private this As TVariable
Private Const VARIABLENAMECOLUMN As String = "variable name"
Private Const INDEXHEADER As String = "column index"
Private Const CLASSNAME As String = "LLVariables"
Private Const COL_SHEET_NAME As String = "sheet name"
Private Const COL_CONTROL As String = "control"
Private Const COL_TABLE_NAME As String = "table name"

'@section Creation
'===============================================================================

'@label:create
'@sub-title Build a variables reader over one dictionary
'@details
'Validates the dictionary before anything else: a missing object, a dictionary
'with no `DataSheet`, and a dictionary with no "variable name" column each
'raise their own error. `TestLLVariables` pins all three numbers.
'@param dict LLdictionary. The dictionary to read.
'@return LLVariables. A sealed instance bound to that dictionary.
'@throws ProjectError.InvalidArgument When dict is Nothing.
'@throws ProjectError.ObjectNotInitialized When the dictionary has no DataSheet.
'@throws ProjectError.ElementNotFound When the variable-name column is missing.
'@export
Public Function Create(ByVal dict As LLdictionary) As LLVariables

    If dict Is Nothing Then
        ThrowError ProjectError.InvalidArgument, _
                   "Unable to create LLVariables without a dictionary reference"
    End If

    If dict.Data Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary instance does not expose a DataSheet implementation"
    End If

    If Not dict.ColumnExists(VARIABLENAMECOLUMN) Then
        ThrowError ProjectError.ElementNotFound, _
                   "Unable to create LLVariables, required column '" & VARIABLENAMECOLUMN & "' is missing"
    End If

    Dim newVariables As LLVariables
    Set newVariables = New LLVariables
    Set newVariables.Dictionary = dict
    newVariables.Seal
    Set Create = newVariables

End Function

'@label:dictionary
'@prop-title Dictionary this object reads
'@return LLdictionary
'@export
Public Property Get Dictionary() As LLdictionary
    Set Dictionary = this.dict
End Property

'@label:dictionary-set
'@prop-title Bind the dictionary, at creation only
'@param dict LLdictionary. The dictionary to read.
Public Property Set Dictionary(ByVal dict As LLdictionary)
    GuardNotSealed "Dictionary"
    Set this.dict = dict
    InvalidateCaches
    RefreshVariableRange
End Property

'@label:seal
'@sub-title Prevent further changes to creation-only setters
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Refuse a write to a creation-only setter
'@param propName String. Name used in the message.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@section Cache management
'===============================================================================
'@description
'The two snapshots and the host worksheet live here. Read the class header for
'what each one promises.

'@label:InvalidateCaches
'@sub-title Drop everything this object holds about the worksheet
'@details
'The way out for a caller that changed the dictionary behind this object:
'`EventSetup` calls it after the setup sheets are rebuilt. The search range
'goes too, so the next read rebuilds it from the dictionary.
'@export
Public Sub InvalidateCaches()
    DropSnapshots
    Set this.varRng = Nothing
End Sub

'@label:drop-snapshots
'@sub-title Forget the two blocks and the host worksheet
Private Sub DropSnapshots()
    Erase this.variableNames
    Erase this.headers
    this.namesLoaded = False
    this.variableCount = 0
    this.firstVariableRow = 0
    this.headersLoaded = False
    this.headerCount = 0
    this.firstHeaderColumn = 0
    Set this.hostWksh = Nothing
End Sub

'@label:EnsureDictionaryReady
'@sub-title Refuse to read without a dictionary
'@throws ProjectError.ObjectNotInitialized When no dictionary is bound.
Private Sub EnsureDictionaryReady()
    If this.dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "LLVariables dictionary reference is missing"
    End If
End Sub

'@label:host-sheet
'@prop-title Worksheet holding the dictionary
'@details
'Held on the instance. Reaching it through `dict.Data.Wksh` cost two crossings
'before any work, on every cell this class resolved.
'@return Worksheet
Private Property Get HostSheet() As Worksheet
    If this.hostWksh Is Nothing Then
        EnsureDictionaryReady
        Set this.hostWksh = this.dict.Data.Wksh
    End If
    Set HostSheet = this.hostWksh
End Property

'@label:RefreshVariableRange
'@sub-title Rebuild the range holding the variable names
'@details
'One path. `DataSheet.DataRange` resolves the column with `shouldExist:=True`,
'so a dictionary without a "variable name" column raises here, and it always
'hands back at least one data row. The second hand-built path this routine
'used to carry could never run.
'@throws ProjectError.ElementNotFound When the variable-name column is missing.
Private Sub RefreshVariableRange()
    EnsureDictionaryReady
    Set this.varRng = this.dict.DataRange(VARIABLENAMECOLUMN)
End Sub

'@label:EnsureVariableRange
'@sub-title Build the search range when it is missing
Private Sub EnsureVariableRange()
    If this.varRng Is Nothing Then RefreshVariableRange
End Sub

'@label:ensure-names-loaded
'@sub-title Read the variable-name column once
'@details
'One crossing for the whole column. Every `Contains`, `RowIndex`,
'`VariableNames` and cell resolve is answered from this array afterwards.
Private Sub EnsureNamesLoaded()
    Dim rawValues As Variant
    Dim counter As Long
    Dim nbRows As Long

    If this.namesLoaded Then Exit Sub

    EnsureVariableRange

    nbRows = this.varRng.Rows.Count
    this.firstVariableRow = this.varRng.Row
    ReDim this.variableNames(1 To nbRows)

    If nbRows = 1 Then
        this.variableNames(1) = CellText(this.varRng.Cells(1, 1).Value)
    Else
        rawValues = this.varRng.Columns(1).Value
        For counter = 1 To nbRows
            this.variableNames(counter) = CellText(rawValues(counter, 1))
        Next
    End If

    this.variableCount = nbRows
    this.namesLoaded = True
End Sub

'@label:reload-headers
'@sub-title Read the header row once
Private Sub ReloadHeaders()
    Dim dictData As DataSheet
    Dim headerRng As Range
    Dim rawValues As Variant
    Dim counter As Long
    Dim nbColumns As Long

    EnsureDictionaryReady
    Set dictData = this.dict.Data
    Set headerRng = dictData.HeaderRange

    nbColumns = headerRng.Columns.Count
    this.firstHeaderColumn = headerRng.Column
    ReDim this.headers(1 To nbColumns)

    If nbColumns = 1 Then
        this.headers(1) = CellText(headerRng.Cells(1, 1).Value)
    Else
        rawValues = headerRng.Rows(1).Value
        For counter = 1 To nbColumns
            this.headers(counter) = CellText(rawValues(1, counter))
        Next
    End If

    this.headerCount = nbColumns
    this.headersLoaded = True
End Sub

'@label:ensure-headers-loaded
'@sub-title Read the header row when it has not been read yet
'@details
'Says whether it read the sheet. A snapshot built in this very call is already
'fresh, so the caller can skip both the cell check and the second read.
'@return Boolean. True when the header row was read in this call.
Private Function EnsureHeadersLoaded() As Boolean
    If this.headersLoaded Then Exit Function
    ReloadHeaders
    EnsureHeadersLoaded = True
End Function

'@label:header-column
'@sub-title Find a column in the header snapshot
'@details
'Whole string, case-insensitive, leftmost wins. That is the same rule
'`DataSheet.FindColumn` applies for the arguments this class used to pass it
'(`strictSearch:=True`, `matchCase:=False`).
'@param colName String. Column header to look for.
'@return Long. The worksheet column, or 0 when the snapshot does not hold it.
Private Function HeaderColumn(ByVal colName As String) As Long
    Dim counter As Long

    For counter = 1 To this.headerCount
        If StrComp(this.headers(counter), colName, vbTextCompare) = 0 Then
            HeaderColumn = this.firstHeaderColumn + counter - 1
            Exit Function
        End If
    Next
End Function

'@label:header-still-reads
'@sub-title Check one header cell against the name it is expected to hold
'@details
'One crossing, and it is what stops a stale index. A header renamed on the
'worksheet leaves the snapshot pointing at a column that now holds something
'else, and handing that index out would read the wrong column's value.
'@param columnNumber Long. The worksheet column to test.
'@param colName String. The header it should hold.
'@return Boolean. True when the cell still reads as colName.
Private Function HeaderStillReads(ByVal columnNumber As Long, ByVal colName As String) As Boolean
    Dim headerText As String

    headerText = CellText(HostSheet().Cells(this.dict.Data.DataStartRow, columnNumber).Value)
    HeaderStillReads = (StrComp(headerText, colName, vbTextCompare) = 0)
End Function

'@label:ResolveColumnIndex
'@sub-title Resolve a dictionary column to a worksheet column
'@details
'Answers from the header snapshot, checks the answer with one cell read, and
'reads the header row again when either step comes up short. The second read
'is what lets a column added after this object was built resolve.
'@param colName String. Column header to resolve.
'@param required Optional Boolean. When True, a missing column raises. Defaults to False.
'@return Long. The worksheet column, or 0 when it is missing and required is False.
'@throws ProjectError.InvalidArgument When colName is empty and required is True.
'@throws ProjectError.ElementNotFound When the column is missing and required is True.
Private Function ResolveColumnIndex(ByVal colName As String, Optional ByVal required As Boolean = False) As Long
    Dim columnNumber As Long
    Dim headersAreFresh As Boolean

    If LenB(colName) = 0 Then
        If required Then
            ThrowError ProjectError.InvalidArgument, "Column name cannot be empty"
        End If
        Exit Function
    End If

    EnsureDictionaryReady
    headersAreFresh = EnsureHeadersLoaded()

    columnNumber = HeaderColumn(colName)

    If Not headersAreFresh Then
        If columnNumber > 0 Then
            If HeaderStillReads(columnNumber, colName) Then
                ResolveColumnIndex = columnNumber
                Exit Function
            End If
        End If

        'The snapshot is behind the worksheet. Read the header row again and
        'ask once more; the answer that comes back now is fresh.
        ReloadHeaders
        columnNumber = HeaderColumn(colName)
    End If

    If columnNumber > 0 Then
        ResolveColumnIndex = columnNumber
        Exit Function
    End If

    If required Then
        ThrowError ProjectError.ElementNotFound, _
                   "Dictionary column '" & colName & "' not found"
    End If
End Function

'@label:cell-text
'@sub-title Read one cell value as text
'@details
'A dictionary cell can hold an error value (`#N/A`, `#REF!`) or a Null, and
'`CStr` raises on both. Every read in this class goes through here.
'The parameter is called `cellValue` because this module also holds a public
'`Value` function. VBA renormalises every use of an identifier to whichever
'spelling it saw last, across the whole project.
'@param cellValue Variant. The raw cell value.
'@return String. The text, or an empty string when there is none.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@section Reading
'===============================================================================

'@label:contains
'@sub-title Test whether a variable is in the dictionary
'@details
'Case-sensitive by default. `LLdictionary.UniqueVarNames` normalises variable
'names during preparation, so the stored case is whatever the user typed, and
'`LLSheets.VariableAddress` and `LLExport.ResolveVariableChunk` both rely on
'the case-sensitive answer.
'The match is a string compare in memory, so a name holding `*`, `?`, `[`, `]`
'or `#` is matched as written. The `Range.Find` this replaced had to escape
'those characters first.
'@param varName String. Variable to look for.
'@param matchCase Optional Boolean. When True, the match is case-sensitive. Defaults to True.
'@return Boolean. True when the variable is present.
'@export
Public Function Contains(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Boolean
    Contains = (RowIndex(varName, matchCase) > 0)
End Function

'@label:row-index
'@sub-title Find the worksheet row of one variable
'@details
'Answers 0 when the variable is absent, and every caller tests for 0.
'A duplicated variable name answers with the FIRST row that holds it.
'`Range.Find` started after the first cell and wrapped, so it used to answer
'with the second. `LLdictionary.UniqueVarNames` makes the names unique during
'preparation, so two equal names are a fault in the sheet either way.
'@param varName String. Variable to look for.
'@param matchCase Optional Boolean. When True, the match is case-sensitive. Defaults to True.
'@return Long. The worksheet row, or 0 when the variable is absent.
Private Function RowIndex(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Long
    Dim counter As Long
    Dim compareMethod As Long

    If LenB(varName) = 0 Then Exit Function

    EnsureNamesLoaded
    compareMethod = IIf(matchCase, vbBinaryCompare, vbTextCompare)

    For counter = 1 To this.variableCount
        If StrComp(this.variableNames(counter), varName, compareMethod) = 0 Then
            RowIndex = this.firstVariableRow + counter - 1
            Exit Function
        End If
    Next
End Function

'@label:cell-range-internal
'@sub-title Resolve the cell holding one column of one variable
'@param colName String. Column header.
'@param varName String. Variable name.
'@param matchCase Boolean. When True, the variable match is case-sensitive.
'@return Range. The cell, or Nothing when the row or the column is missing.
Private Function CellRangeInternal(ByVal colName As String, _
                                   ByVal varName As String, _
                                   ByVal matchCase As Boolean) As Range
    Dim rowNum As Long
    Dim colNum As Long

    EnsureDictionaryReady

    rowNum = RowIndex(varName, matchCase)
    If rowNum = 0 Then Exit Function

    colNum = ResolveColumnIndex(colName)
    If colNum = 0 Then Exit Function

    Set CellRangeInternal = HostSheet().Cells(rowNum, colNum)
End Function

'@label:read-column-value
'@sub-title Read one column of one variable as text
'@details
'The single reader. `Value`, `SheetName`, `ControlType` and `TableName` are all
'one line over this; the two argument orders those members expose are the only
'thing that used to differ between them.
'@param varName String. Variable name.
'@param colName String. Column header.
'@param matchCase Boolean. When True, the variable match is case-sensitive.
'@return String. The cell text, or an empty string.
Private Function ReadColumnValue(ByVal varName As String, _
                                 ByVal colName As String, _
                                 ByVal matchCase As Boolean) As String
    Dim targetCell As Range

    Set targetCell = CellRangeInternal(colName, varName, matchCase)
    If targetCell Is Nothing Then Exit Function

    ReadColumnValue = CellText(targetCell.Value)
End Function

'@label:value
'@sub-title Value of one column for one variable
'@details
'The argument order is column first. Twenty production sites call it with
'named arguments, so both names are part of the contract.
'@param colName String. Column header.
'@param varName String. Variable name.
'@return String. The cell text, or an empty string when either is missing.
'@export
Public Function Value(ByVal colName As String, ByVal varName As String) As String
    Value = ReadColumnValue(varName, colName, True)
End Function

'@label:SheetName
'@sub-title Sheet a variable belongs to
'@param varName String. Variable name.
'@param matchCase Optional Boolean. When True, the variable match is case-sensitive. Defaults to True.
'@return String. The sheet name, or an empty string.
'@export
Public Function SheetName(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    SheetName = ReadColumnValue(varName, COL_SHEET_NAME, matchCase)
End Function

'@label:ControlType
'@sub-title Control type of a variable
'@param varName String. Variable name.
'@param matchCase Optional Boolean. When True, the variable match is case-sensitive. Defaults to True.
'@return String. The control value, or an empty string.
'@export
Public Function ControlType(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    ControlType = ReadColumnValue(varName, COL_CONTROL, matchCase)
End Function

'@label:TableName
'@sub-title Table a variable belongs to
'@details
'The table name column is written by `LLdictionary.Prepare`, so this answers
'an empty string on a dictionary that has not been prepared.
'@param varName String. Variable name.
'@return String. The table name, or an empty string.
'@export
Public Function TableName(ByVal varName As String) As String
    TableName = ReadColumnValue(varName, COL_TABLE_NAME, True)
End Function

'@label:cell-range
'@prop-title Cell holding one column of one variable
'@param colName String. Column header.
'@param varName String. Variable name.
'@return Range. The cell, or Nothing when the row or the column is missing.
'@export
Public Property Get CellRange(ByVal colName As String, ByVal varName As String) As Range
    Set CellRange = CellRangeInternal(colName, varName, True)
End Property

'@label:VariableNames
'@sub-title List of every variable name in the dictionary
'@details
'Built from the snapshot, so it costs no worksheet read once the column has
'been loaded. Empty cells are left out.
'@return BetterArray. The names, lower bound 1.
'@export
Public Function VariableNames() As BetterArray
    Dim names As BetterArray
    Dim counter As Long

    Set names = New BetterArray
    names.LowerBound = 1

    EnsureNamesLoaded

    For counter = 1 To this.variableCount
        If LenB(this.variableNames(counter)) > 0 Then
            names.Push this.variableNames(counter)
        End If
    Next

    Set VariableNames = names
End Function

'@section Writing
'===============================================================================

'@label:SetValue
'@sub-title Write one column of one variable
'@details
'`onEmpty` tests `IsEmpty` on the target cell. That is False for a cell holding
'an empty string and for a formula that returns "", which is the behaviour
'`LinelistSpecs` relies on for the "list auto" column.
'The return value says whether the write happened. A skipped write used to be
'recorded in a checking object that nothing ever read.
'@param varName String. Variable to write to.
'@param colName String. Column header to write to.
'@param newValue String. The value to write.
'@param onEmpty Optional Boolean. When True, write only into an empty cell. Defaults to False.
'@return Boolean. True when the value was written.
'@throws ProjectError.ElementNotFound When the variable or the column is missing.
'@export
Public Function SetValue(ByVal varName As String, _
                         ByVal colName As String, _
                         ByVal newValue As String, _
                         Optional ByVal onEmpty As Boolean = False) As Boolean

    Dim rowNum As Long
    Dim colNum As Long
    Dim targetCell As Range

    EnsureDictionaryReady

    rowNum = RowIndex(varName)
    If rowNum = 0 Then
        ThrowError ProjectError.ElementNotFound, "Variable '" & varName & "' not found"
    End If

    colNum = ResolveColumnIndex(colName, required:=True)
    Set targetCell = HostSheet().Cells(rowNum, colNum)

    If onEmpty Then
        If Not IsEmpty(targetCell.Value) Then Exit Function
    End If

    targetCell.Value = newValue
    SetValue = True
End Function

'@label:index
'@sub-title Column index of a variable on its own sheet
'@details
'Reads the "column index" column, which `LLdictionary.Prepare` creates. Before
'preparation the cell is empty and this raises `ElementNotFound`; that is the
'contract `LLSheets.VariableAddress` turns into its own message.
'@param varName String. Variable name.
'@return Long. The index stored for that variable.
'@throws ProjectError.ElementNotFound When the column index is missing.
'@throws ProjectError.InvalidArgument When the stored index is not a number.
'@export
Public Function Index(ByVal varName As String) As Long
    Dim sheetIndex As String

    sheetIndex = Value(colName:=INDEXHEADER, varName:=varName)

    If LenB(sheetIndex) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "Column index is missing for variable '" & varName & "'"
    End If

    'Text in the column index cell used to come back as a bare type mismatch
    'from a class that raises ProjectError everywhere else, and the message
    'named neither the variable nor the value.
    If Not IsNumeric(sheetIndex) Then
        ThrowError ProjectError.InvalidArgument, _
                   "Column index for variable '" & varName & "' is not a number. Found: " & sheetIndex
    End If

    Index = CLng(sheetIndex)
End Function

'@section Errors
'===============================================================================

'@label:throw-error
'@sub-title Raise an error from this class
'@param errNumb Long. The ProjectError value.
'@param errorMessage String. The message.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub
