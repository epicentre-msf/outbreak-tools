VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Spatial analysis table management"

'@PredeclaredId
'@Folder("Geo")
'@ModuleDescription("Spatial analysis table management")
'@IgnoreModule UnrecognizedAnnotation, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class LLSpatial
'@description
'Fills the spatial analysis tables of a generated linelist, orders them, and
'reads ranked values back out for the TopAdmin, TopPop and TopHF worksheet
'functions.
'
'THE CONTRACT WITH SPATIALTABLES
'-------------------------------------------------------------------------------
'SpatialTables builds these tables at build time and this class fills them at
'runtime. The two never reference each other: the contract is a set of names
'agreed by string convention across a workbook boundary, and both files hold
'their own copy of the four constants and the two name builders. That is the
'house rule — a class keeps every constant it needs inside itself. Change a name
'here and change it in SpatialTables in the same commit.
'
'    listofgeovars              one row per spatial variable, "<var>_<tabId>"
'    spatial_<level>_<var>      one ListObject per level
'      column 1  tabl_<level>_<var>    the lookup key this class writes
'      column 2  formula_<level>       the value this class sorts on
'      column 3  population_<level>    administrative levels only
'      column 4  attack_rate_<level>   administrative levels only
'
'The width of a table and the meaning of each column are SpatialTables' to
'decide. Every column is resolved by name here, so a table that grows a column
'keeps working.
'
'WHERE THE VALUES COME FROM
'-------------------------------------------------------------------------------
'Update walks the HList sheets of the workbook. A column controlled by geo1 or
'by hf names a variable, and the filtered companion sheet of that HList carries
'the concatenated geo values of that variable, one column per administrative
'level. The distinct values of that column are the rows of the spatial table.
'The reading and the de-duplication happen in memory, so nothing is pasted
'through the RNG_PastingCol scratch column that SpatialTables still creates.
'
'WHAT A MISSING REGISTRY MEANS
'-------------------------------------------------------------------------------
'Linelist creates the spatial worksheet on every linelist and SpatialTables
'creates the listofgeovars table only when a spatial cross-table is defined. So
'a linelist with geo variables and no spatial analysis has the sheet with no
'table in it, and Update on such a sheet is a legitimate no-op: Exists, the
'variable lookup and Sort all answer empty when the table is absent.
'
'RNG_TESTINGFORMULA BELONGS TO SPATIALTABLES
'-------------------------------------------------------------------------------
'SpatialTables.Prepare creates it and SpatialTables.FailedFormula reads it. This
'class never touches it.
'@depends BetterArray, HiddenNames, Checking
'@version 2026-07-31


'@section State
'===============================================================================

Private Type TLLSpatial
    sh As Worksheet
    sealed As Boolean
End Type

Private this As TLLSpatial
Private Const CLASS_NAME As String = "LLSpatial"
Private Const SPATIAL_SHEET As String = "spatial_tables__"
Private Const ADMIN As String = "adm"
Private Const CONCAT As String = "concat_"
Private Const HF As String = "hf"

' The registry of spatial variables. SpatialTables.Prepare builds it.
Private Const REGISTRY_TABLE As String = "listofgeovars"

' How many tables a geographic variable carries.
Private Const ADMIN_LEVELS As Long = 4

' The hidden names an HList sheet carries, and the suffix VarWriter appends to a
' variable name for its control value.
Private Const SHEET_TAG As String = "sheet_type"
Private Const FILTERED_SHEET As String = "filtered_sheet"
Private Const TABLE_NAME As String = "table_name"
Private Const CONTROL_SUFFIX As String = " -- control"

' How wide one dictionary control is on an HList sheet. LLdictionary expands a
' geo control into 12 contiguous columns — 4 levels, 4 pcodes and 4 concats —
' and an hf control into 2, the facility and its pcode.
Private Const GEO_COLUMN_SPAN As Long = 12
Private Const HF_COLUMN_SPAN As Long = 2

' The separator of the de-duplication index. It sits on both sides of every
' entry, so a short entry cannot match inside a longer one. A Const takes a
' constant expression, which rules out a Chr$ or ChrW$ call.
Private Const KEY_SEPARATOR As String = vbNullChar

' What a rank property answers when it is called without a table identifier. It
' goes into the cell of the worksheet function, so the reader sees why the cell
' is not a value.
Private Const NO_TABLE_ID As String = "#missing table id"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new LLSpatial instance from the spatial tables worksheet
'@details
'Factory method that checks the supplied worksheet is given and is named
'"spatial_tables__", then returns a new instance. Called on the predeclared
'instance.
'
'The listofgeovars table is left out of the check on purpose: the worksheet and
'the table have different lifecycles, and a linelist that carries the sheet with
'no table still needs an instance for the two callers that build one on every
'run. See the class description.
'@param sh Worksheet. The spatial tables worksheet.
'@return LLSpatial. A fully initialised instance ready for use.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.InvalidArgument When the worksheet name is not "spatial_tables__".
'@export
Public Function Create(ByVal sh As Worksheet) As LLSpatial
    CheckRequirements sh

    Dim instance As LLSpatial
    Set instance = New LLSpatial
    Set instance.Wksh = sh
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Seal the instance against further setup writes.
'@details
'Marks the instance as sealed so setup-only setters raise if called again.
'Called by the factory after the last setup write.
Public Sub Seal()
    this.sealed = True
End Sub

'@sub-title Guard a setup-only setter against post-seal writes.
'@param propName String. The property name used in the error message.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:wksh
'@prop-title Spatial tables worksheet
'@return Worksheet. The backing spatial tables worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the spatial tables worksheet
'@param sh Worksheet. The worksheet to store.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sh = sh
End Property

'@section Naming Conventions
'===============================================================================

'@label:current-lo-name
'@prop-title ListObject name for a spatial table
'@details
'Builds the conventional ListObject name for a spatial table. The pattern is
'"spatial_{adminName}_{varName}", for instance "spatial_adm1_cases_sp1".
'SpatialTables builds the same string from the same pieces.
'@param varName String. The variable name component, with its table identifier.
'@param adminName String. The admin level component (e.g. "adm1", "hf").
'@return String. The constructed ListObject name.
Private Property Get CurrentLoName(ByVal varName As String, _
                                   ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'@section Worksheet Lookups
'===============================================================================
'@description The guarded resolvers every other section reads through. Each one
'declares its own local, so a failed lookup answers Nothing rather than the
'answer of the call before it.

'@label:lo-or-nothing
'@sub-title The named ListObject of the spatial sheet, or Nothing
'@param loName String. The ListObject name to resolve.
'@return ListObject. The table, or Nothing when the sheet carries no such name.
Private Function LoOrNothing(ByVal loName As String) As ListObject
    Dim Lo As ListObject

    On Error Resume Next
    Set Lo = Wksh().ListObjects(loName)
    On Error GoTo 0

    Set LoOrNothing = Lo
End Function

'@label:column-body-or-nothing
'@sub-title The data body of one named column of a table, or Nothing
'@param Lo ListObject. The table to read.
'@param colName String. The column name.
'@return Range. The data body of that column, or Nothing.
Private Function ColumnBodyOrNothing(ByVal Lo As ListObject, _
                                     ByVal colName As String) As Range
    Dim col As ListColumn

    On Error Resume Next
    Set col = Lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then Exit Function
    Set ColumnBodyOrNothing = col.DataBodyRange
End Function

'@label:column-range-or-nothing
'@sub-title One named column of a table with its header, or Nothing
'@param Lo ListObject. The table to read.
'@param colName String. The column name.
'@return Range. The whole column including its header, or Nothing.
Private Function ColumnRangeOrNothing(ByVal Lo As ListObject, _
                                      ByVal colName As String) As Range
    Dim col As ListColumn

    On Error Resume Next
    Set col = Lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then Exit Function
    Set ColumnRangeOrNothing = col.Range
End Function

'@label:range-values
'@sub-title Read a range into a two dimensional array in one crossing
'@details
'Range.Value hands back a scalar for a single cell and an array for anything
'wider, so a one-cell range is wrapped to give every caller the same shape.
'@param rng Range. The range to read.
'@return Variant. A 1-based two dimensional array of the values.
Private Function RangeValues(ByVal rng As Range) As Variant
    Dim oneCell(1 To 1, 1 To 1) As Variant

    If rng.Cells.Count = 1 Then
        oneCell(1, 1) = rng.Value
        RangeValues = oneCell
    Else
        RangeValues = rng.Value
    End If
End Function

'@section The Registry Of Spatial Variables
'===============================================================================

'@label:registered-names
'@sub-title Every name the listofgeovars table holds
'@details
'One crossing reads the whole column into memory, and every match in this class
'is then a string test. Range.Find inherits LookIn, SearchOrder and
'SearchDirection from the last search of the Excel session, including one the
'user ran from the Find dialog, so a lookup built on it can miss for a reason
'that has nothing to do with the workbook.
'
'The header cell is left out because the data body is read on its own. A search
'over the whole table range answered True for "list" on an empty table, and a
'header hit could become a variable name.
'@return BetterArray. The registered names, empty when the table is absent.
Private Function RegisteredNames() As BetterArray
    Dim Lo As ListObject
    Dim body As Range
    Dim readValues As Variant
    Dim counter As Long
    Dim entry As String
    Dim held As BetterArray

    Set held = New BetterArray

    Set Lo = LoOrNothing(REGISTRY_TABLE)

    If Not Lo Is Nothing Then
        Set body = Lo.ListColumns(1).DataBodyRange

        If Not body Is Nothing Then
            readValues = RangeValues(body)

            For counter = LBound(readValues, 1) To UBound(readValues, 1)
                If Not IsError(readValues(counter, 1)) Then
                    entry = CStr(readValues(counter, 1))
                    If LenB(entry) > 0 Then held.Push entry
                End If
            Next
        End If
    End If

    Set RegisteredNames = held
End Function

'@label:list-of-existing-var-names
'@sub-title The spatial names one variable carries
'@details
'A registered name is always "<variable>_<tableId>", and one variable can carry
'several spatial tables, so the test is on the prefix and the underscore that
'follows it. A bare substring test also caught every other variable whose name
'contains this one: with "cases" and "new_cases" both registered, updating
'"cases" rewrote the four "new_cases_sp1" tables from the "concat_adm1_cases"
'column, and the values landed in the wrong tables looking plausible.
'@param varName String. The bare variable name.
'@return BetterArray. The matching registered names.
Private Function ListOfExistingVarNames(ByVal varName As String) As BetterArray
    Dim names As BetterArray
    Dim held As BetterArray
    Dim counter As Long
    Dim entry As String
    Dim prefix As String

    Set names = RegisteredNames()
    Set held = New BetterArray
    prefix = varName & "_"

    For counter = names.LowerBound To names.UpperBound
        entry = names.Item(counter)
        If Left$(entry, Len(prefix)) = prefix Then held.Push entry
    Next

    Set ListOfExistingVarNames = held
End Function

'@label:registered-name-for-table
'@sub-title The variable a table identifier belongs to
'@details
'The test is on the suffix. A substring test matched "sp10" and "sp11" when it
'was asked for "sp1", so it broke on the day a linelist carried ten spatial
'cross-tables.
'@param tabId String. The table identifier.
'@return String. The registered name, empty when nothing carries that identifier.
Private Function RegisteredNameForTable(ByVal tabId As String) As String
    Dim names As BetterArray
    Dim counter As Long
    Dim entry As String
    Dim suffix As String

    Set names = RegisteredNames()
    suffix = "_" & tabId

    For counter = names.LowerBound To names.UpperBound
        entry = names.Item(counter)
        If Right$(entry, Len(suffix)) = suffix Then
            RegisteredNameForTable = entry
            Exit Function
        End If
    Next
End Function

'@label:exists
'@sub-title Check if a variable has spatial tables
'@details
'Answers True when the listofgeovars table carries at least one name built on
'this variable. A spatial sheet with no listofgeovars table answers False.
'@param varName String. The bare variable name to look up.
'@return Boolean. True when the variable is registered.
'@export
Public Function Exists(ByVal varName As String) As Boolean
    Exists = (ListOfExistingVarNames(varName).Length > 0)
End Function

'@section Update
'===============================================================================

'@label:update
'@sub-title Update all spatial tables from HList filtered data
'@details
'Walks the worksheets of the workbook and works on the HList ones. For each of
'them it resolves the filtered companion sheet, then walks the variable columns
'of the table looking for a control value of geo1 or of hf. A geo1 column takes
'the concat column of the companion and fills the four administrative tables of
'the variable; an hf column fills the single facility table. The spatial sheet is
'recalculated at the end so the formula columns answer for the rows just written.
'
'The control value of a variable is a hidden name of the sheet, written by
'VarWriter as "<variable> -- control". This routine used to read it from a hidden
'worksheet row that stopped being written on 2026-02-06, so it matched nothing
'and every spatial table stayed at the one blank row the build gave it.
'@throws Any error raised while walking a sheet, after the spatial sheet has been
'recalculated. The caller is what puts Excel back the way it found it.
'@export
Public Sub Update()
    Dim actsh As Worksheet
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim shHn As HiddenNames
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    Set actsh = Wksh()
    Set wb = actsh.Parent

    On Error GoTo UpdateFailed

    For Each sh In wb.Worksheets
        Set shHn = HiddenNames.Create(sh)
        If shHn.ValueAsString(SHEET_TAG) = "HList" Then UpdateOneHList sh, wb, shHn
    Next

    actsh.UsedRange.Calculate
    Exit Sub

UpdateFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description

    On Error Resume Next
    actsh.UsedRange.Calculate
    On Error GoTo 0

    Err.Raise errNumber, errSource, errMessage
End Sub

'@label:update-one-hlist
'@sub-title Fill the spatial tables of the variables of one HList sheet
'@details
'The columns of one variable sit side by side, so a geo column is followed by
'the eleven columns of its expansion and a facility column by its pcode. A sheet
'with no filtered companion, no table name or no first table is left alone.
'@param sh Worksheet. The HList worksheet.
'@param wb Workbook. The workbook holding the sheet.
'@param shHn HiddenNames. The hidden name store of that sheet.
Private Sub UpdateOneHList(ByVal sh As Worksheet, ByVal wb As Workbook, _
                           ByVal shHn As HiddenNames)
    Dim filtsh As Worksheet
    Dim Lo As ListObject
    Dim cellRng As Range
    Dim rng As Range
    Dim spatialNames As BetterArray
    Dim tableName As String
    Dim varName As String
    Dim controlValue As String

    On Error Resume Next
    Set filtsh = wb.Worksheets(shHn.ValueAsString(FILTERED_SHEET))
    On Error GoTo 0
    If filtsh Is Nothing Then Exit Sub

    On Error Resume Next
    Set Lo = filtsh.ListObjects(1)
    On Error GoTo 0
    If Lo Is Nothing Then Exit Sub
    If Lo.DataBodyRange Is Nothing Then Exit Sub

    tableName = shHn.ValueAsString(TABLE_NAME)
    If LenB(tableName) = 0 Then Exit Sub

    On Error Resume Next
    Set cellRng = sh.Range(tableName & "_START")
    On Error GoTo 0
    If cellRng Is Nothing Then Exit Sub

    Do While Not IsEmpty(cellRng)
        controlValue = shHn.ValueAsString(cellRng.Value & CONTROL_SUFFIX)

        If controlValue = "geo1" Then
            varName = NameWithoutPrefix(cellRng.Value, ADMIN & "1_")
            Set spatialNames = ListOfExistingVarNames(varName)

            If spatialNames.Length > 0 Then
                Set rng = ColumnBodyOrNothing(Lo, CONCAT & ADMIN & "1_" & varName)
                If Not rng Is Nothing Then UpdateSpatialValue rng, spatialNames
            End If

            Set cellRng = cellRng.Offset(, GEO_COLUMN_SPAN - 1)

        ElseIf controlValue = HF Then
            varName = NameWithoutPrefix(cellRng.Value, HF & "_")
            Set spatialNames = ListOfExistingVarNames(varName)

            If spatialNames.Length > 0 Then
                Set rng = ColumnBodyOrNothing(Lo, HF & "_" & varName)
                If Not rng Is Nothing Then UpdateSpatialValue rng, spatialNames, isHF:=True
            End If

            Set cellRng = cellRng.Offset(, HF_COLUMN_SPAN - 1)
        End If

        Set cellRng = cellRng.Offset(, 1)
    Loop
End Sub

'@label:name-without-prefix
'@sub-title Cut a known prefix off the front of a name
'@details
'The header cell of a geo column always carries its prefix at the front. Replace
'took the prefix out wherever it sat, so a variable named "adm1_site_adm1_ref"
'came back as "sitref".
'@param fullName String. The name as the header cell holds it.
'@param prefix String. The prefix to cut.
'@return String. The name with the leading prefix removed.
Private Function NameWithoutPrefix(ByVal fullName As String, _
                                   ByVal prefix As String) As String
    If Left$(fullName, Len(prefix)) = prefix Then
        NameWithoutPrefix = Mid$(fullName, Len(prefix) + 1)
    Else
        NameWithoutPrefix = fullName
    End If
End Function

'@label:update-spatial-value
'@sub-title Fill the spatial tables of one variable from a filtered column
'@details
'A geographic variable has one table per administrative level, and the columns
'of the companion sheet hold concat_adm1 through concat_adm4 side by side, so the
'input range walks one column right per level. A facility variable has one table
'and one column.
'@param inputRng Range. The filtered column the first level reads.
'@param spatialNames BetterArray. The registered names of this variable.
'@param isHF Optional Boolean. True for the facility table. Defaults to False.
Private Sub UpdateSpatialValue(ByVal inputRng As Range, _
                               ByVal spatialNames As BetterArray, _
                               Optional ByVal isHF As Boolean = False)
    Dim rng As Range
    Dim Lo As ListObject
    Dim counter As Long
    Dim adminCounter As Long
    Dim endCounter As Long
    Dim varName As String
    Dim adminName As String

    endCounter = IIf(isHF, 1, ADMIN_LEVELS)

    For counter = spatialNames.LowerBound To spatialNames.UpperBound
        varName = spatialNames.Item(counter)
        Set rng = inputRng

        For adminCounter = 1 To endCounter
            adminName = IIf(isHF, HF, ADMIN & adminCounter)
            Set Lo = LoOrNothing(CurrentLoName(varName, adminName))

            If Not Lo Is Nothing Then FillOneTable Lo, adminName, DistinctValues(rng)

            Set rng = rng.Offset(, 1)
        Next
    Next
End Sub

'@label:distinct-values
'@sub-title The distinct values of a range, read in one crossing
'@details
'The values used to go out to a scratch column, through RemoveDuplicates, and
'back again — five worksheet round trips per level per variable per sheet, on a
'range that can run to tens of thousands of rows. The index is a delimited
'string, and the separator sits on both sides of every entry so a short entry
'cannot match inside a longer one.
'
'Blanks are dropped here. EventLinelist.GeoConcat answers an empty string for a
'record whose geo levels are not all filled, RemoveDuplicates collapsed those to
'one blank, and that blank became a data row that could take rank 1 once the
'tables were ordered.
'
'Two values that differ only in case are one value, which is what
'RemoveDuplicates answered and what COUNTIFS answers over the same data.
'@param rng Range. The column to read.
'@return Variant. A 1-based two dimensional array of the distinct values, or
'Empty when the range holds none.
Private Function DistinctValues(ByVal rng As Range) As Variant
    Dim readValues As Variant
    Dim held() As String
    Dim result() As Variant
    Dim seen As String
    Dim entry As String
    Dim counter As Long
    Dim total As Long

    readValues = RangeValues(rng)
    ReDim held(1 To UBound(readValues, 1))
    seen = KEY_SEPARATOR

    For counter = LBound(readValues, 1) To UBound(readValues, 1)
        If Not IsError(readValues(counter, 1)) Then
            entry = CStr(readValues(counter, 1))

            If LenB(entry) > 0 Then
                If InStr(1, seen, KEY_SEPARATOR & LCase$(entry) & KEY_SEPARATOR, _
                         vbBinaryCompare) = 0 Then
                    seen = seen & LCase$(entry) & KEY_SEPARATOR
                    total = total + 1
                    held(total) = entry
                End If
            End If
        End If
    Next

    If total = 0 Then Exit Function

    ReDim result(1 To total, 1 To 1)
    For counter = 1 To total
        result(counter, 1) = held(counter)
    Next

    DistinctValues = result
End Function

'@label:fill-one-table
'@sub-title Write the values of one spatial table and order it
'@details
'The rows below the first are cleared before the resize, so a table that shrinks
'leaves nothing of its old self standing under it. The first data row keeps the
'formulas SpatialTables wrote, and they are copied down over the new rows.
'
'The table is calculated before it is ordered. Update recalculates the whole
'sheet, and that happens once every sheet has been walked, so the sort would read
'the values of the run before.
'@param Lo ListObject. The spatial table to fill.
'@param adminName String. The level identifier the sort column is named for.
'@param values Variant. The values for the first column, or Empty for none.
Private Sub FillOneTable(ByVal Lo As ListObject, ByVal adminName As String, _
                         ByRef values As Variant)
    Dim headerRng As Range
    Dim nbRows As Long

    ClearDataRowsBelowFirst Lo

    If IsEmpty(values) Then
        If Not Lo.DataBodyRange Is Nothing Then _
            Lo.ListColumns(1).DataBodyRange.ClearContents
        Exit Sub
    End If

    nbRows = UBound(values, 1)
    Set headerRng = Lo.HeaderRowRange

    Lo.Resize Wksh().Range(headerRng.Cells(1, 1), _
                           headerRng.Cells(nbRows + 1, Lo.ListColumns.Count))

    Lo.ListColumns(1).DataBodyRange.Value = values
    FillFormulaColumns Lo
    Lo.DataBodyRange.Calculate
    SortOneTable Lo, adminName, onAR:=False
End Sub

'@label:clear-data-rows-below-first
'@sub-title Empty every data row of a table but the first
'@details
'ClearContents leaves the number formats, the borders and the validation of the
'column standing. Clear took those with the values, so the key column lost its
'formatting on every update.
'@param Lo ListObject. The table to clear.
Private Sub ClearDataRowsBelowFirst(ByVal Lo As ListObject)
    Dim body As Range

    Set body = Lo.DataBodyRange
    If body Is Nothing Then Exit Sub
    If body.Rows.Count <= 1 Then Exit Sub

    Wksh().Range(body.Cells(2, 1), _
                 body.Cells(body.Rows.Count, body.Columns.Count)).ClearContents
End Sub

'@label:fill-formula-columns
'@sub-title Copy the formulas of the first data row over every row
'@details
'The rows used to be left to the calculated-column autofill of the host, which
'is an option a user can switch off, and Resize is a weaker trigger for it than
'a row insert. SpatialTables writes the value column as an array formula, and an
'array formula is the least likely of all to be filled that way, so only the
'first row of a table held a number and the sort ranked it against blanks.
'
'AutoFill needs a destination taller than one cell and raises 1004 otherwise,
'which a table of one row produces on the ordinary path.
'@param Lo ListObject. The table whose formula columns are filled.
Private Sub FillFormulaColumns(ByVal Lo As ListObject)
    Dim colRng As Range
    Dim counter As Long

    If Lo.DataBodyRange Is Nothing Then Exit Sub
    If Lo.ListRows.Count <= 1 Then Exit Sub

    For counter = 2 To Lo.ListColumns.Count
        Set colRng = Lo.ListColumns(counter).DataBodyRange

        On Error Resume Next
        colRng.Cells(1, 1).AutoFill Destination:=colRng, Type:=xlFillValues
        On Error GoTo 0
    Next
End Sub

'@section Sort
'===============================================================================

'@label:sort
'@sub-title Order the spatial tables of one analysis table
'@details
'Finds the variable the table identifier belongs to, then orders its tables on
'the value column or on the attack rate column, highest first. A facility
'variable carries one table and a geographic one carries four.
'@param tabId String. The identifier of the analysis table.
'@param onAR Optional Boolean. True to order on the attack rate. Defaults to False.
'@export
Public Sub Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False)
    Dim Lo As ListObject
    Dim varName As String
    Dim adminCounter As Long

    varName = RegisteredNameForTable(tabId)
    If LenB(varName) = 0 Then Exit Sub

    'A facility variable has a single table and no administrative one. The loop
    'below built administrative names alone, so the attack rate toggle of a
    'facility spatial sheet reordered nothing.
    Set Lo = LoOrNothing(CurrentLoName(varName, HF))

    If Not Lo Is Nothing Then
        SortOneTable Lo, HF, onAR
        Exit Sub
    End If

    For adminCounter = 1 To ADMIN_LEVELS
        Set Lo = LoOrNothing(CurrentLoName(varName, ADMIN & adminCounter))
        If Not Lo Is Nothing Then SortOneTable Lo, ADMIN & adminCounter, onAR
    Next
End Sub

'@label:sort-one-table
'@sub-title Order one spatial table on one of its value columns
'@details
'Excel wants the sort key inside the range it sorts and raises 1004 when it sits
'outside. Both sort sites of this class ordered the first column alone with a key
'in the second or the fourth, under On Error Resume Next, so no spatial table has
'ever been ordered and every "top N admin units" figure was in insertion order.
'
'Sorting the whole table is safe: SpatialTables writes the population and the
'attack rate as relative references to the cells of their own row, so Excel moves
'them with the row they belong to.
'@param Lo ListObject. The table to order.
'@param adminName String. The level identifier the value columns are named for.
'@param onAR Boolean. True to order on the attack rate.
Private Sub SortOneTable(ByVal Lo As ListObject, ByVal adminName As String, _
                         ByVal onAR As Boolean)
    Dim sortRng As Range
    Dim sortColName As String

    If Lo.ListRows.Count = 0 Then Exit Sub

    sortColName = IIf(onAR, "attack_rate_", "formula_") & adminName
    Set sortRng = ColumnRangeOrNothing(Lo, sortColName)
    If sortRng Is Nothing Then Exit Sub

    Lo.Range.Sort key1:=sortRng, Order1:=xlDescending, Header:=xlYes
End Sub

'@section TopGeoValue / TopHFValue
'===============================================================================

'@label:top-geo-value
'@prop-title Value at a given rank from an ordered spatial table
'@details
'Reads the table of one variable at one administrative level and answers the
'value at the requested rank, either the admin unit itself or its population.
'Answers an empty string when the table is absent, holds no data, or is shorter
'than the rank asked for.
'
'A rank below 1 used to read the header cell or the cells above the table, and
'the rank comes straight from the argument of a worksheet function, so
'=TopAdmin("adm1", 0, ...) in a cell answered with the internal table name.
'@param adminName String. Level identifier ("adm1" to "adm4", or "hf").
'@param orderValue Long. Rank position in the ordered table, 1 for the top one.
'@param inputVarName String. The bare variable name.
'@param tabId Optional String. The identifier of the analysis table. Defaults to vbNullString.
'@param columnChoice Optional Byte. 1 for the admin unit, anything else for its population. Defaults to 1.
'@return String. The value at that rank, or an empty string.
'@export
Public Property Get TopGeoValue(ByVal adminName As String, _
                                 ByVal orderValue As Long, _
                                 ByVal inputVarName As String, _
                                 Optional ByVal tabId As String = vbNullString, _
                                 Optional ByVal columnChoice As Byte = 1) As String
    Dim Lo As ListObject
    Dim rng As Range

    'A spatial table name always carries the identifier of its analysis table, so
    'an empty one builds a name nothing can match. It used to answer an empty
    'string, which reads on the sheet like a table with no data in it.
    If LenB(tabId) = 0 Then
        TopGeoValue = NO_TABLE_ID
        Exit Property
    End If

    Set Lo = LoOrNothing(CurrentLoName(inputVarName & "_" & tabId, adminName))
    If Lo Is Nothing Then Exit Property
    If Lo.DataBodyRange Is Nothing Then Exit Property

    If columnChoice = 1 Then
        Set rng = Lo.ListColumns(1).DataBodyRange
    Else
        Set rng = ColumnBodyOrNothing(Lo, "population_" & adminName)
    End If

    If rng Is Nothing Then Exit Property
    If orderValue < 1 Or orderValue > rng.Rows.Count Then Exit Property

    TopGeoValue = CStr(rng.Cells(orderValue, 1).Value)
End Property

'@label:top-hf-value
'@prop-title Value at a given rank from an ordered facility table
'@details
'The facility table is the same shape as an administrative one under another
'level name, so this is that read with the level fixed. The two used to be two
'copies of one function, and every fix to one of them had to be made twice.
'@param orderValue Long. Rank position in the ordered table, 1 for the top one.
'@param inputVarName String. The bare variable name.
'@param tabId Optional String. The identifier of the analysis table. Defaults to vbNullString.
'@return String. The facility at that rank, or an empty string.
'@export
Public Property Get TopHFValue(ByVal orderValue As Long, _
                                ByVal inputVarName As String, _
                                Optional ByVal tabId As String = vbNullString) As String
    TopHFValue = TopGeoValue(HF, orderValue, inputVarName, tabId)
End Property

'@section Validation
'===============================================================================

'@label:check-requirements
'@sub-title Validate that the worksheet is the spatial tables sheet
'@details
'Guards the factory by checking the worksheet is given and its name matches the
'expected "spatial_tables__" constant.
'@param sh Worksheet. The worksheet to validate.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.InvalidArgument When the worksheet name does not match "spatial_tables__".
Private Sub CheckRequirements(ByVal sh As Worksheet)
    If sh Is Nothing Then
        ThrowError ObjectNotInitialized, "Worksheet not provided"
    End If

    If sh.Name <> SPATIAL_SHEET Then
        ThrowError InvalidArgument, _
            "Expected worksheet named " & SPATIAL_SHEET & " but got " & sh.Name
    End If
End Sub

'@section Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber Long. The ProjectError error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
