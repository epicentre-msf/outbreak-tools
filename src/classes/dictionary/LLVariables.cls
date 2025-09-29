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


'@Folder("Dictionary")
'@ModuleDescription("Variables of the dictionary")
'@depends Checking
'@IgnoreModule PropertyNotUsed, ReadOnlyPropertyAssignment, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

Implements ILLVariables

Private Type TVariable
    dict As ILLdictionary
    varRng As Range
    columnCache As BetterArray
    variablesCache As BetterArray
    checkCounter As Long
    internalChecks As IChecking
End Type

Private this As TVariable
Private Const VARIABLENAME As String = "variable name"
Private Const INDEXHEADER  As String = "column index"
Private Const CLASSNAME As String = "LLVariables"
Private Const COL_SHEET_NAME As String = "sheet name"
Private Const COL_CONTROL As String = "control"
Private Const COL_TABLE_NAME As String = "table name"

Public Function Create(ByVal dict As ILLdictionary) As ILLVariables

    If dict Is Nothing Then
        Err.Raise CLng(ProjectError.InvalidArgument), CLASSNAME, _
                  "Unable to create LLVariables without a dictionary reference"
    End If

    If dict.Data Is Nothing Then
        Err.Raise CLng(ProjectError.ObjectNotInitialized), CLASSNAME, _
                  "Dictionary instance does not expose an IDataSheet implementation"
    End If

    If Not dict.ColumnExists(VARIABLENAME, matchCase:=False) Then
        ThrowError ProjectError.ElementNotFound, _
            "Unable to create LLVariables, required column '" & VARIABLENAME & "' is missing"
    End If

    With New LLVariables
        Set .Dictionary = dict
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ILLVariables
    Set Self = Me
End Property

Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    Set this.dict = dict
    ResetCaches
    RefreshVariableRange
End Property

Friend Property Get VarRange() As Range
    EnsureVariableRange
    Set VarRange = this.varRng
End Property

'@section CacheManagement
'===============================================================================
'@label:ResetCaches
Private Sub ResetCaches()
    Set this.columnCache = New BetterArray
    this.columnCache.LowerBound = 1
    Set this.variablesCache = Nothing
    this.checkCounter = 0
    Set this.internalChecks = Nothing
End Sub

'@label:InvalidateCaches
Private Sub InvalidateCaches()
    ResetCaches
    Set this.varRng = Nothing
End Sub

'@label:EnsureDictionaryReady
Private Sub EnsureDictionaryReady()
    If this.dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "LLVariables dictionary reference is missing"
    End If
End Sub

'@label:RefreshVariableRange
Private Sub RefreshVariableRange()
    EnsureDictionaryReady

    Set this.varRng = Nothing

    On Error Resume Next
        Set this.varRng = this.dict.DataRange(VARIABLENAME)
    On Error GoTo 0

    If this.varRng Is Nothing Then
        Dim headerRange As Range
        On Error Resume Next
            Set headerRange = this.dict.DataRange(VARIABLENAME, includeHeaders:=True)
        On Error GoTo 0

        If Not headerRange Is Nothing Then
            If headerRange.Rows.Count > 1 Then
                Set this.varRng = headerRange.Offset(1, 0).Resize(headerRange.Rows.Count - 1, 1)
            Else
                Set this.varRng = headerRange.Offset(1, 0).Resize(1, 1)
            End If
        Else
            ThrowError ProjectError.ElementNotFound, _
                       "Dictionary does not expose a usable '" & VARIABLENAME & "' column"
        End If
    End If
End Sub

'@label:EnsureVariableRange
Private Sub EnsureVariableRange()
    If this.varRng Is Nothing Then RefreshVariableRange
End Sub

'@label:CacheKeyForColumn
Private Function CacheKeyForColumn(ByVal colName As String) As String
    CacheKeyForColumn = LCase$(Trim$(colName))
End Function

'@label:TryGetCachedColumnIndex
Private Function TryGetCachedColumnIndex(ByVal cacheKey As String, ByRef cachedIndex As Long) As Boolean
    Dim idx As Long
    Dim entry As Variant

    cachedIndex = 0
    TryGetCachedColumnIndex = False

    If this.columnCache Is Nothing Then Exit Function
    If this.columnCache.Length = 0 Then Exit Function

    For idx = this.columnCache.LowerBound To this.columnCache.UpperBound
        entry = this.columnCache.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), cacheKey, vbBinaryCompare) = 0 Then
                cachedIndex = CLng(entry(1))
                TryGetCachedColumnIndex = True
                Exit Function
            End If
        End If
    Next idx
End Function

'@label:StoreColumnIndex
Private Sub StoreColumnIndex(ByVal cacheKey As String, ByVal columnIndexValue As Long)
    Dim entry As Variant

    If columnIndexValue <= 0 Then Exit Sub
    If this.columnCache Is Nothing Then ResetCaches

    RemoveCachedColumnIndex cacheKey
    entry = Array(cacheKey, columnIndexValue)
    this.columnCache.Push entry
End Sub

Private Sub RemoveCachedColumnIndex(ByVal cacheKey As String)
    Dim idx As Long

    If this.columnCache Is Nothing Then Exit Sub
    If this.columnCache.Length = 0 Then Exit Sub

    For idx = this.columnCache.UpperBound To this.columnCache.LowerBound Step -1
        If IsArray(this.columnCache.Item(idx)) Then
            If StrComp(CStr(this.columnCache.Item(idx)(0)), cacheKey, vbBinaryCompare) = 0 Then
                this.columnCache.Remove idx
            End If
        End If
    Next idx
End Sub

Private Function ColumnIndexStillValid(ByVal columnIndexValue As Long, ByVal expectedColumnName As String) As Boolean
    Dim headerCell As Range
    Dim headerText As String
    Dim dataSheet As IDataSheet
    Dim relativeIndex As Long

    ColumnIndexStillValid = False

    EnsureDictionaryReady
    Set dataSheet = this.dict.Data

    relativeIndex = columnIndexValue - dataSheet.DataStartColumn + 1
    If relativeIndex < 1 Then Exit Function
    If relativeIndex > dataSheet.HeaderRange.Columns.Count Then Exit Function

    On Error Resume Next
        Set headerCell = dataSheet.HeaderRange.Cells(1, relativeIndex)
    On Error GoTo 0

    If headerCell Is Nothing Then Exit Function

    headerText = LCase$(Trim$(CStr(headerCell.Value)))
    ColumnIndexStillValid = (headerText = LCase$(Trim$(expectedColumnName)))
End Function

'@label:SanitizeSearchTerm
Private Function SanitizeSearchTerm(ByVal value As String) As String
    Dim sanitized As String

    sanitized = Replace(value, "~", "~~")
    sanitized = Replace(sanitized, "*", "~*")
    sanitized = Replace(sanitized, "?", "~?")
    sanitized = Replace(sanitized, "[", "~[")
    sanitized = Replace(sanitized, "]", "~]")
    sanitized = Replace(sanitized, "#", "~#")

    SanitizeSearchTerm = sanitized
End Function

'@label:FindVariableCell
Private Function FindVariableCell(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Range
    Dim searchValue As String

    EnsureVariableRange

    If LenB(varName) = 0 Then Exit Function

    searchValue = SanitizeSearchTerm(varName)

    On Error Resume Next
        Set FindVariableCell = this.varRng.Find(What:=searchValue, _
                                               LookAt:=xlWhole, _
                                               MatchCase:=matchCase, _
                                               LookIn:=xlValues, _
                                               SearchOrder:=xlByRows)
    On Error GoTo 0
End Function

'@label:ResolveColumnIndex
Private Function ResolveColumnIndex(ByVal colName As String, Optional ByVal required As Boolean = False) As Long
    Dim cacheKey As String
    Dim cachedIndex As Long
    Dim resolvedIndex As Long

    If LenB(colName) = 0 Then
        If required Then
            ThrowError ProjectError.InvalidArgument, "Column name cannot be empty"
        End If
        Exit Function
    End If

    EnsureDictionaryReady

    cacheKey = CacheKeyForColumn(colName)
    If TryGetCachedColumnIndex(cacheKey, cachedIndex) Then
        If ColumnIndexStillValid(cachedIndex, colName) Then
            ResolveColumnIndex = cachedIndex
            Exit Function
        Else
            RemoveCachedColumnIndex cacheKey
        End If
    End If

    If Not this.dict.ColumnExists(colName) Then
        If required Then
            ThrowError ProjectError.ElementNotFound, _
                       "Dictionary column '" & colName & "' not found"
        End If
        Exit Function
    End If

    On Error Resume Next
        resolvedIndex = this.dict.Data.ColumnIndex(colName, shouldExist:=True)
    On Error GoTo 0

    If resolvedIndex <= 0 Then
        If required Then
            ThrowError ProjectError.ElementNotFound, _
                       "Unable to resolve dictionary column '" & colName & "'"
        End If
        Exit Function
    End If

    StoreColumnIndex cacheKey, resolvedIndex
    ResolveColumnIndex = resolvedIndex
End Function

'@label:VariableNamesList
Private Function VariableNamesList() As BetterArray
    Dim cell As Range

    EnsureVariableRange

    If this.variablesCache Is Nothing Then
        Set this.variablesCache = New BetterArray
        this.variablesCache.LowerBound = 1

        For Each cell In this.varRng.Cells
            If LenB(CStr(cell.Value)) > 0 Then
                this.variablesCache.Push CStr(cell.Value)
            End If
        Next cell
    End If

    If this.variablesCache Is Nothing Then
        Set VariableNamesList = New BetterArray
        VariableNamesList.LowerBound = 1
    Else
        Set VariableNamesList = this.variablesCache.Clone()
    End If
End Function

'@section Logging
'===============================================================================
'@label:LogInfo
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal varName As String = vbNullString, _
                    Optional ByVal colName As String = vbNullString)

    Dim checkName As String
    Dim message As String

    checkName = "LLVariables trace"
    message = label
    If LenB(varName) > 0 Then message = message & " | Var=" & varName
    If LenB(colName) > 0 Then message = message & " | Col=" & colName

    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create(checkName)
    this.internalChecks.Add CStr(this.checkCounter), message
    this.checkCounter = this.checkCounter + 1
End Sub

Private Property Get HasInternalChecks() As Boolean
    HasInternalChecks = Not (this.internalChecks Is Nothing)
End Property

'Functions --------------------------------------------------------------------------

'Test if a variable is present
Private Function Contains(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Boolean
    Contains = Not (FindVariableCell(varName, matchCase) Is Nothing)
End Function

'Find the row index of one variable
Private Function RowIndex(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Long
    Dim foundCell As Range

    Set foundCell = FindVariableCell(varName, matchCase)
    If Not foundCell Is Nothing Then RowIndex = foundCell.Row Else RowIndex = -1
End Function

'Value of one column for one variable
Private Function Value(ByVal colName As String, ByVal varName As String) As String
    'Row and columns to return
    Dim retValue As String
    Dim retRng As Range
    'return value
    retValue = vbNullString
    Set retRng = CellRangeInternal(colName, varName, True)
    If Not (retRng Is Nothing) Then retValue = retRng.Value
    Value = retValue
End Function

'Metadata helpers -------------------------------------------------------------

'@label:SheetName
Private Function SheetName(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    SheetName = ColumnValue(varName, COL_SHEET_NAME, matchCase)
End Function

'@label:ControlType
Private Function ControlType(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    ControlType = ColumnValue(varName, COL_CONTROL, matchCase)
End Function

'@label:TableName
Private Function TableName(ByVal varName As String) As String
    TableName = ColumnValue(varName, COL_TABLE_NAME, True)
End Function

Private Property Get CellRange(ByVal colName As String, ByVal varName As String) As Range

'Row and columns to return
    Set CellRange = CellRangeInternal(colName, varName, True)
End Property

'@label:CellRangeInternal
Private Function CellRangeInternal(ByVal colName As String, ByVal varName As String, ByVal matchCase As Boolean) As Range

    Dim rowNum As Long
    Dim colNum As Long
    Dim dict As ILLdictionary
    Dim Wksh As Worksheet
    EnsureDictionaryReady
    EnsureVariableRange
    Set dict = Dictionary()
    Set Wksh = dict.Data.Wksh
    rowNum = RowIndex(varName, matchCase)
    If rowNum <= 0 Then Exit Property

    colNum = ResolveColumnIndex(colName, required:=True)
    If colNum <= 0 Then Exit Property

    Set CellRangeInternal = Wksh.Cells(rowNum, colNum)
End Function

'@label:ColumnValue
Private Function ColumnValue(ByVal varName As String, ByVal columnName As String, ByVal matchCase As Boolean) As String
    Dim targetCell As Range

    Set targetCell = CellRangeInternal(columnName, varName, matchCase)
    If Not targetCell Is Nothing Then ColumnValue = CStr(targetCell.Value)
End Function

Private Sub SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)

    'Row and columns to return
    Dim rowNum As Long
    Dim colNum As Long
    Dim Update As Boolean

    Dim dict As ILLdictionary
    Dim Wksh As Worksheet

    EnsureDictionaryReady
    EnsureVariableRange
    Set dict = Dictionary()
    Set Wksh = dict.Data.Wksh

    rowNum = RowIndex(varName)
    If rowNum <= 0 Then
        ThrowError ProjectError.ElementNotFound, "Variable '" & varName & "' not found"
    End If

    colNum = ResolveColumnIndex(colName, required:=True)

    Update = True
    If onEmpty Then Update = IsEmpty(Wksh.Cells(rowNum, colNum))

    If Update Then
        Wksh.Cells(rowNum, colNum).Value = newValue
    Else
        LogInfo "Skipped SetValue because target cell already contains data", varName, colName
    End If
End Sub

'Index of a column or a Row of a variable in a worksheet
Private Function Index(ByVal varName As String) As Long
    Dim sheetIndex As String
    sheetIndex = Value(colName:=INDEXHEADER, varName:=varName)
    If sheetIndex = vbNullString Then
        ThrowError ProjectError.ElementNotFound, _
                   "Column index is missing for variable '" & varName & "'"
    End If
    Index = CLng(sheetIndex)
End Function

'Errors ------------------------------------------------------------------------------------------------------------------------------------
Private Sub ThrowError(ByVal errNumb As ProjectError, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub

'Interface ---------------------------------------------------------------------------------------------------------------------------------

'Properties
Public Property Get ILLVariables_Dictionary() As ILLdictionary
    Set ILLVariables_Dictionary = Dictionary()
End Property

'Exposed methods
Public Function ILLVariables_Value(ByVal colName As String, ByVal varName As String) As String
    ILLVariables_Value = Value(colName, varName)
End Function

'column or row index in the sheet of one variable
Public Function ILLVariables_Index(ByVal varName As String) As Long
    ILLVariables_Index = Index(varName)
End Function

Public Function ILLVariables_Contains(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As Boolean
    ILLVariables_Contains = Contains(varName, matchCase)
End Function

Public Sub ILLVariables_SetValue(ByVal varName As String, ByVal colName As String, ByVal newValue As String, Optional ByVal onEmpty As Boolean = False)
    SetValue varName, colName, newValue, onEmpty
End Sub

Public Property Get ILLVariables_CellRange(ByVal colName As String, ByVal varName As String) As Range
    Set ILLVariables_CellRange = CellRange(colName, varName)
End Property

Public Function ILLVariables_VariableNames() As BetterArray
    Set ILLVariables_VariableNames = VariableNamesList()
End Function

Public Function ILLVariables_SheetName(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    ILLVariables_SheetName = SheetName(varName, matchCase)
End Function

Public Function ILLVariables_ControlType(ByVal varName As String, Optional ByVal matchCase As Boolean = True) As String
    ILLVariables_ControlType = ControlType(varName, matchCase)
End Function

Public Function ILLVariables_TableName(ByVal varName As String) As String
    ILLVariables_TableName = TableName(varName)
End Function

Public Sub ILLVariables_InvalidateCaches()
    InvalidateCaches
End Sub

Public Function ILLVariables_HasCheckings() As Boolean
    ILLVariables_HasCheckings = HasInternalChecks
End Function

Public Property Get ILLVariables_CheckingValues() As IChecking
    If HasInternalChecks Then Set ILLVariables_CheckingValues = this.internalChecks
End Property
