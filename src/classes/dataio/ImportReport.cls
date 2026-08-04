VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "The four tables the last import report is kept in"

'@PredeclaredId
'@Folder("DataIO")
'@ModuleDescription("The four tables the last import report is kept in")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ProcedureNotUsed

Option Explicit

'@class ImportReport
'@description
'The store an import report is kept in. It owns four ListObjects on the
'import_rep__ worksheet. ProvisionTables builds the tables, SaveSheets and
'SaveVariables write the entries, Clear empties them, HasEntries says whether
'anything was recorded, and SheetNames and VariableEntries read the entries
'back. The class knows nothing about forms.
'
'THE STORE MAKES ITS OWN TABLES
'-------------------------------------------------------------------------------
'Create builds any of the four that is absent. A linelist built by the current
'generation carries all four, and one built by an older generation may carry
'none.
'
'THE REPORT LIVES ON A WORKSHEET
'-------------------------------------------------------------------------------
'The report survives the workbook closing, so a user can open the last one from
'the advanced form days later.
'
'THE FOUR TABLES
'-------------------------------------------------------------------------------
'  reptab_sheetsNotImp    a worksheet of the file this linelist has none of
'  reptab_sheetsNotTouch  a worksheet of this linelist the file has none of
'  reptab_varsNotImp      a variable of the file this linelist has none of
'  reptab_varsNotTouch    a variable of this linelist the file has none of
'
'The two variable tables carry the sheet a variable belongs to in a second
'column. The values stored here are names of worksheets and variables, and they
'are stored untranslated. The form translates its own labels.
'@depends BetterArray, LLImporter
'@version 1.0 (2026-08-01)


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "ImportReport"

'The worksheet a linelist keeps its last import report on
Private Const STORE_SHEET As String = "import_rep__"

'The four tables and the columns each one starts at
Private Const TAB_SHEETS_NOT_IMP As String = "reptab_sheetsNotImp"
Private Const TAB_SHEETS_NOT_TOUCH As String = "reptab_sheetsNotTouch"
Private Const TAB_VARS_NOT_IMP As String = "reptab_varsNotImp"
Private Const TAB_VARS_NOT_TOUCH As String = "reptab_varsNotTouch"

Private Const COL_SHEETS_NOT_IMP As Long = 1
Private Const COL_SHEETS_NOT_TOUCH As Long = 3
Private Const COL_VARS_NOT_IMP As Long = 5
Private Const COL_VARS_NOT_TOUCH As Long = 8


'@section State
'===============================================================================

Private Type TImportReport
    sh As Worksheet
    sealed As Boolean
End Type

Private this As TImportReport


'@section Factory
'===============================================================================

'@label:create
'@sub-title Bind to the import report worksheet of a linelist
'@details
'The worksheet is made when the workbook has none, and any of the four tables
'that is absent is built. The sheet is left very hidden, which is how the
'generation makes it: the form reads the tables and nobody opens the sheet.
'@param wkb Workbook. The linelist workbook.
'@return ImportReport. A ready store.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Function Create(ByVal wkb As Workbook) As ImportReport

    Dim newReport As ImportReport
    Dim sh As Worksheet

    If wkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "A workbook is needed to find the import report store"
    End If

    Set sh = Nothing
    On Error Resume Next
    Set sh = wkb.Worksheets(STORE_SHEET)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wkb.Worksheets.Add(After:=wkb.Worksheets(wkb.Worksheets.Count))
        sh.Name = STORE_SHEET
        sh.Visible = xlSheetVeryHidden
    End If

    Set newReport = New ImportReport
    Set newReport.Wksh = sh
    newReport.Seal
    newReport.ProvisionTables
    Set Create = newReport
End Function

'@label:wksh
'@prop-title The worksheet the four tables sit on
'@return Worksheet. The store worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the worksheet the four tables sit on
'@param sh Worksheet. The store worksheet.
Public Property Set Wksh(ByVal sh As Worksheet)
    GuardNotSealed "Wksh"
    Set this.sh = sh
End Property

'@label:seal
'@prop-title Prevent further changes to setup-only properties
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Raise when a setup-only property is written after sealing
'@param propName String. The name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.ElementShouldNotExists, _
                   propName & " is set at creation only"
End Sub

'@label:provision-tables
'@sub-title Build any of the four tables the worksheet is missing
'@details
'Public because Create calls it on the instance it just made, after sealing.
'Never activates or selects the sheet: import_rep__ is very hidden and
'activating a very hidden sheet raises.
'@export
Public Sub ProvisionTables()
    EnsureTable TAB_SHEETS_NOT_IMP, COL_SHEETS_NOT_IMP, _
                "sheetNotImported", vbNullString
    EnsureTable TAB_SHEETS_NOT_TOUCH, COL_SHEETS_NOT_TOUCH, _
                "sheetNotTouched", vbNullString
    EnsureTable TAB_VARS_NOT_IMP, COL_VARS_NOT_IMP, _
                "variableNotImported", "correspondingSheet"
    EnsureTable TAB_VARS_NOT_TOUCH, COL_VARS_NOT_TOUCH, _
                "variableNotFound", "correspondingSheet"
End Sub


'@section Writing the report
'===============================================================================

'@label:clear
'@sub-title Empty the four tables
'@details
'Called at the start of every import, so the store always describes the last
'one rather than every one that ever ran.
'@export
Public Sub Clear()
    ClearOneTable TAB_SHEETS_NOT_IMP
    ClearOneTable TAB_SHEETS_NOT_TOUCH
    ClearOneTable TAB_VARS_NOT_IMP
    ClearOneTable TAB_VARS_NOT_TOUCH
End Sub

'@label:save-sheets
'@sub-title Write one list of worksheet names into its table
'@param scope Byte. An ImportReportScope value.
'@param sheetsList BetterArray. The worksheet names.
'@export
Public Sub SaveSheets(ByVal scope As Byte, ByVal sheetsList As BetterArray)

    Dim tableName As String

    tableName = SheetTableOf(scope)
    If LenB(tableName) = 0 Then Exit Sub

    WriteOneColumn tableName, sheetsList
End Sub

'@label:save-variables
'@sub-title Write one list of variables and their sheets into its table
'@param scope Byte. An ImportReportScope value.
'@param varsList BetterArray. Entries of Array(variableName, sheetName).
'@export
Public Sub SaveVariables(ByVal scope As Byte, ByVal varsList As BetterArray)

    Dim tableName As String

    tableName = VariableTableOf(scope)
    If LenB(tableName) = 0 Then Exit Sub

    WriteTwoColumns tableName, varsList
End Sub


'@section Reading the report back
'===============================================================================

'@label:has-entries
'@prop-title Whether the store holds anything at all
'@return Boolean. True when at least one of the four tables holds a row.
'@export
Public Property Get HasEntries() As Boolean

    HasEntries = True

    If RowCountOf(TAB_SHEETS_NOT_IMP) > 0 Then Exit Property
    If RowCountOf(TAB_SHEETS_NOT_TOUCH) > 0 Then Exit Property
    If RowCountOf(TAB_VARS_NOT_IMP) > 0 Then Exit Property
    If RowCountOf(TAB_VARS_NOT_TOUCH) > 0 Then Exit Property

    HasEntries = False
End Property

'@label:sheet-names
'@prop-title The worksheet names one table holds
'@details
'Named SheetNames rather than Sheets, because Sheets is the workbook collection
'and a member of that name reads as one everywhere it is used.
'@param scope Byte. An ImportReportScope value.
'@return BetterArray. The names. Never Nothing.
'@export
Public Property Get SheetNames(ByVal scope As Byte) As BetterArray
    Set SheetNames = ReadOneColumn(SheetTableOf(scope))
End Property

'@label:variable-entries
'@prop-title The variables and their sheets one table holds
'@param scope Byte. An ImportReportScope value.
'@return BetterArray. Entries of Array(variableName, sheetName). Never Nothing.
'@export
Public Property Get VariableEntries(ByVal scope As Byte) As BetterArray
    Set VariableEntries = ReadTwoColumns(VariableTableOf(scope))
End Property


'@section Private - The four tables
'===============================================================================

'@label:sheet-table-of
'@sub-title The worksheet table one scope writes into
'@param scope Byte. An ImportReportScope value.
'@return String. The table name, empty for an unknown scope.
Private Function SheetTableOf(ByVal scope As Byte) As String
    Select Case scope
    Case ImportReportNotImported
        SheetTableOf = TAB_SHEETS_NOT_IMP
    Case ImportReportNotTouched
        SheetTableOf = TAB_SHEETS_NOT_TOUCH
    End Select
End Function

'@label:variable-table-of
'@sub-title The variable table one scope writes into
'@param scope Byte. An ImportReportScope value.
'@return String. The table name, empty for an unknown scope.
Private Function VariableTableOf(ByVal scope As Byte) As String
    Select Case scope
    Case ImportReportNotImported
        VariableTableOf = TAB_VARS_NOT_IMP
    Case ImportReportNotTouched
        VariableTableOf = TAB_VARS_NOT_TOUCH
    End Select
End Function

'@label:named-table
'@sub-title The ListObject of one name, or Nothing
'@param tableName String. The table to find.
'@return ListObject. The table, or Nothing when the sheet carries none.
Private Function NamedTable(ByVal tableName As String) As ListObject
    On Error Resume Next
    Set NamedTable = this.sh.ListObjects(tableName)
    On Error GoTo 0
End Function

'@label:ensure-table
'@sub-title Build one table when the worksheet has none of that name
'@param tableName String. The table name.
'@param startColumn Long. The column its first header sits in.
'@param firstHeader String. The first header text.
'@param secondHeader String. The second header text, empty for a one-column table.
Private Sub EnsureTable(ByVal tableName As String, ByVal startColumn As Long, _
                        ByVal firstHeader As String, ByVal secondHeader As String)

    Dim lo As ListObject
    Dim headerRange As Range
    Dim lastColumn As Long

    Set lo = NamedTable(tableName)
    If Not lo Is Nothing Then Exit Sub

    lastColumn = startColumn
    this.sh.Cells(1, startColumn).Value = firstHeader

    If LenB(secondHeader) > 0 Then
        lastColumn = startColumn + 1
        this.sh.Cells(1, lastColumn).Value = secondHeader
    End If

    Set headerRange = this.sh.Range(this.sh.Cells(1, startColumn), _
                                    this.sh.Cells(2, lastColumn))
    this.sh.ListObjects.Add(xlSrcRange, headerRange, , xlYes).Name = tableName
End Sub

'@label:clear-one-table
'@sub-title Empty one table and take it back to a single blank row
'@param tableName String. The table to empty.
Private Sub ClearOneTable(ByVal tableName As String)

    Dim lo As ListObject
    Dim loRng As Range

    Set lo = NamedTable(tableName)
    If lo Is Nothing Then Exit Sub

    On Error Resume Next
    lo.DataBodyRange.ClearContents
    On Error GoTo 0

    Set loRng = lo.Range
    lo.Resize this.sh.Range(loRng.Cells(1, 1), _
                            loRng.Cells(2, loRng.Columns.Count))
End Sub

'@label:row-count-of
'@sub-title How many filled rows one table holds
'@details
'A table always keeps one data row, which is blank when nothing was written, so
'the count is taken off the first column rather than off the table height.
'@param tableName String. The table to measure.
'@return Long. The number of filled rows.
Private Function RowCountOf(ByVal tableName As String) As Long

    Dim lo As ListObject
    Dim counter As Long

    Set lo = NamedTable(tableName)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For counter = 1 To lo.DataBodyRange.Rows.Count
        If LenB(CStr(lo.DataBodyRange.Cells(counter, 1).Value)) = 0 Then Exit For
        RowCountOf = counter
    Next counter
End Function

'@label:write-one-column
'@sub-title Write a list of names into a one-column table
'@param tableName String. The table to write into.
'@param valuesList BetterArray. The names.
Private Sub WriteOneColumn(ByVal tableName As String, ByVal valuesList As BetterArray)

    Dim lo As ListObject
    Dim loRng As Range

    Set lo = NamedTable(tableName)
    If lo Is Nothing Then Exit Sub
    If valuesList Is Nothing Then Exit Sub

    ClearOneTable tableName
    If valuesList.Length = 0 Then Exit Sub

    Set loRng = lo.Range
    valuesList.ToExcelRange loRng.Cells(2, 1)

    lo.Resize this.sh.Range(loRng.Cells(1, 1), _
                            loRng.Cells(valuesList.Length + 1, loRng.Columns.Count))
End Sub

'@label:write-two-columns
'@sub-title Write a list of variable and sheet pairs into a two-column table
'@param tableName String. The table to write into.
'@param valuesList BetterArray. Entries of Array(variableName, sheetName).
Private Sub WriteTwoColumns(ByVal tableName As String, ByVal valuesList As BetterArray)

    Dim lo As ListObject
    Dim loRng As Range
    Dim entry As Variant
    Dim counter As Long

    Set lo = NamedTable(tableName)
    If lo Is Nothing Then Exit Sub
    If valuesList Is Nothing Then Exit Sub

    ClearOneTable tableName
    If valuesList.Length = 0 Then Exit Sub

    Set loRng = lo.Range

    'The pair is read through LBound rather than from 0. BetterArray.Push
    'REBASES an inner array to the outer LowerBound, so Array(name, sheet)
    'pushed into a 1-based list comes back 1-based and every reader written
    'against element 0 raises subscript out of range.
    For counter = valuesList.LowerBound To valuesList.UpperBound
        entry = valuesList.Item(counter)
        loRng.Cells(counter - valuesList.LowerBound + 2, 1).Value = _
            CStr(entry(LBound(entry)))
        loRng.Cells(counter - valuesList.LowerBound + 2, 2).Value = _
            CStr(entry(LBound(entry) + 1))
    Next counter

    lo.Resize this.sh.Range(loRng.Cells(1, 1), _
                            loRng.Cells(valuesList.Length + 1, loRng.Columns.Count))
End Sub

'@label:read-one-column
'@sub-title Read the filled rows of a one-column table
'@param tableName String. The table to read.
'@return BetterArray. The names. Never Nothing.
Private Function ReadOneColumn(ByVal tableName As String) As BetterArray

    Dim lo As ListObject
    Dim answer As BetterArray
    Dim rowCount As Long
    Dim counter As Long

    Set answer = New BetterArray
    answer.LowerBound = 1
    Set ReadOneColumn = answer

    If LenB(tableName) = 0 Then Exit Function

    Set lo = NamedTable(tableName)
    If lo Is Nothing Then Exit Function

    rowCount = RowCountOf(tableName)
    For counter = 1 To rowCount
        answer.Push CStr(lo.DataBodyRange.Cells(counter, 1).Value)
    Next counter
End Function

'@label:read-two-columns
'@sub-title Read the filled rows of a two-column table
'@param tableName String. The table to read.
'@return BetterArray. Entries of Array(variableName, sheetName). Never Nothing.
Private Function ReadTwoColumns(ByVal tableName As String) As BetterArray

    Dim lo As ListObject
    Dim answer As BetterArray
    Dim rowCount As Long
    Dim counter As Long

    Set answer = New BetterArray
    answer.LowerBound = 1
    Set ReadTwoColumns = answer

    If LenB(tableName) = 0 Then Exit Function

    Set lo = NamedTable(tableName)
    If lo Is Nothing Then Exit Function

    rowCount = RowCountOf(tableName)
    For counter = 1 To rowCount
        answer.Push Array(CStr(lo.DataBodyRange.Cells(counter, 1).Value), _
                          CStr(lo.DataBodyRange.Cells(counter, 2).Value))
    Next counter
End Function


'@section Private - Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a class level error
'@param errNumber ProjectError. The error to raise.
'@param message String. What went wrong, for a person reading it.
'@throws ProjectError Always raises.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
