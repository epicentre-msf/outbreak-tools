VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSectionContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Shared dictionary context for section and subsection writers")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSectionContext

Private Type TLLSectionContext
    ll As ILinelist
    startRow As Long
    dict As ILLdictionary
    specs As ILinelistSpecs
    columnCache As Collection
End Type

Private this As TLLSectionContext
Private Const CLASSNAME As String = "LLSectionContext"

'@section Lifecycle helpers
'===============================================================================
Private Sub ResetColumnCache()
    Set this.columnCache = New Collection
End Sub

Private Sub EnsureInitialised()
    If this.ll Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Section context has not been initialised with a linelist"
    End If
    If this.startRow <= 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Section context start row has not been configured"
    End If
End Sub

Private Sub EnsureDictionary()
    EnsureInitialised
    If this.dict Is Nothing Then
        Set this.dict = this.ll.Dictionary
    End If
    If this.dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Section context could not resolve the dictionary"
    End If
End Sub

Private Sub EnsureSpecs()
    EnsureInitialised
    If this.specs Is Nothing Then
        Set this.specs = this.ll.LinelistData
    End If
    If this.specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Section context could not resolve linelist specifications"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Function ColumnCacheKey(ByVal columnName As String) As String
    ColumnCacheKey = LCase$(Trim$(columnName))
End Function

Private Function FetchColumnRange(ByVal columnName As String) As Range
    Dim key As String
    Dim cached As Variant

    EnsureDictionary
    key = ColumnCacheKey(columnName)

    On Error Resume Next
        Set cached = this.columnCache.Item(key)
    On Error GoTo 0

    If TypeName(cached) = "Range" Then
        Set FetchColumnRange = cached
        Exit Function
    End If

    On Error Resume Next
        Set cached = this.dict.DataRange(colName:=columnName)
    On Error GoTo 0

    If cached Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Column '" & columnName & "' was not found in the dictionary"
    End If

    On Error Resume Next
        this.columnCache.Add cached, key
    On Error GoTo 0

    Set FetchColumnRange = cached
End Function

Private Function NormaliseColumns(ByVal columnNames As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    Set result = New BetterArray
    result.LowerBound = 1

    If IsArray(columnNames) Then
        For idx = LBound(columnNames) To UBound(columnNames)
            result.Push Trim$(CStr(columnNames(idx)))
        Next idx
    Else
        result.Push Trim$(CStr(columnNames))
    End If

    If result.Length = 0 Then
        ThrowError ProjectError.InvalidArgument, "At least one column name is required"
    End If

    Set NormaliseColumns = result
End Function

Private Function ResolveValue(ByVal columnName As String, ByVal rowIndex As Long) As String
    Dim targetRange As Range
    Dim relativeRow As Long

    Set targetRange = FetchColumnRange(columnName)

    relativeRow = rowIndex - targetRange.Row + 1
    If relativeRow < 1 Or relativeRow > targetRange.Rows.Count Then Exit Function

    On Error Resume Next
        ResolveValue = CStr(targetRange.Cells(relativeRow, 1).Value)
    On Error GoTo 0
End Function

Private Function ComputeEndRow(ByVal startRow As Long, ByVal columns As BetterArray) As Long
    Dim currentRow As Long
    Dim lastRow As Long
    Dim idx As Long
    Dim baseValue As String
    Dim compareValue As String

    EnsureDictionary
    lastRow = DataEndRowInternal()
    currentRow = startRow

    Do While currentRow <= lastRow
        currentRow = currentRow + 1

        If currentRow > lastRow Then Exit Do

        For idx = columns.LowerBound To columns.UpperBound
            baseValue = ResolveValue(columns.Item(idx), startRow)
            compareValue = ResolveValue(columns.Item(idx), currentRow)
            If StrComp(baseValue, compareValue, vbBinaryCompare) <> 0 Then
                GoTo Finish
            End If
        Next idx
    Loop

Finish:
    ComputeEndRow = currentRow - 1
End Function

Private Function DataEndRowInternal() As Long
    Dim dataSheet As IDataSheet
    EnsureDictionary
    Set dataSheet = this.dict.Data
    If dataSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary does not expose a backing data sheet"
    End If
    DataEndRowInternal = dataSheet.DataEndRow
End Function

Private Function StartRowSafe() As Long
    EnsureInitialised
    StartRowSafe = this.startRow
End Function

Private Function VariableColumnName() As String
    VariableColumnName = "variable name"
End Function

'@section Interface implementation
'===============================================================================
Private Sub ILLSectionContext_Initialise(ByVal ll As ILinelist, ByVal startRow As Long)

    If ll Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist reference is required"
    End If

    If startRow <= 0 Then
        ThrowError ProjectError.InvalidArgument, "Start row must be greater than zero"
    End If

    Set this.ll = ll
    this.startRow = startRow
    Set this.dict = Nothing
    Set this.specs = Nothing
    ResetColumnCache
End Sub

Private Sub ILLSectionContext_ResetCaches()
    Set this.dict = Nothing
    Set this.specs = Nothing
    ResetColumnCache
End Sub

Private Property Get ILLSectionContext_Linelist() As ILinelist
    EnsureInitialised
    Set ILLSectionContext_Linelist = this.ll
End Property

Private Property Get ILLSectionContext_StartRow() As Long
    ILLSectionContext_StartRow = StartRowSafe()
End Property

Private Property Let ILLSectionContext_StartRow(ByVal newValue As Long)
    If newValue <= 0 Then
        ThrowError ProjectError.InvalidArgument, "Start row must be greater than zero"
    End If
    this.startRow = newValue
End Property

Private Function ILLSectionContext_Dictionary() As ILLdictionary
    EnsureDictionary
    Set ILLSectionContext_Dictionary = this.dict
End Function

Private Function ILLSectionContext_Specs() As ILinelistSpecs
    EnsureSpecs
    Set ILLSectionContext_Specs = this.specs
End Function

Private Function ILLSectionContext_Design() As ILLFormat
    EnsureSpecs
    Set ILLSectionContext_Design = this.specs.DesignFormat
End Function

Private Function ILLSectionContext_Translation() As ITranslationObject
    EnsureSpecs
    Set ILLSectionContext_Translation = this.specs.TransObject
End Function

Private Function ILLSectionContext_DataEndRow() As Long
    ILLSectionContext_DataEndRow = DataEndRowInternal()
End Function

Private Function ILLSectionContext_Value(ByVal columnName As String, _
                                         Optional ByVal rowOffset As Long = 0) As String
    Dim targetRow As Long

    targetRow = StartRowSafe() + rowOffset
    ILLSectionContext_Value = ResolveValue(columnName, targetRow)
End Function

Private Function ILLSectionContext_EndRow(ByVal columnNames As Variant) As Long
    Dim columns As BetterArray
    Dim endRowValue As Long

    Set columns = NormaliseColumns(columnNames)
    If columns.Length = 0 Then
        ThrowError ProjectError.InvalidArgument, "At least one column is required to compute the end row"
    End If
    endRowValue = ComputeEndRow(StartRowSafe(), columns)
    ILLSectionContext_EndRow = endRowValue
End Function

Private Function ILLSectionContext_HasNext(ByVal columnNames As Variant) As Boolean
    Dim endRowValue As Long
    Dim columns As BetterArray

    endRowValue = ILLSectionContext_EndRow(columnNames)
    Set columns = NormaliseColumns(columnNames)

    If endRowValue >= DataEndRowInternal() Then
        ILLSectionContext_HasNext = False
        Exit Function
    End If

    ILLSectionContext_HasNext = True

    Dim idx As Long
    Dim baseValue As String
    Dim compareValue As String

    For idx = columns.LowerBound To columns.UpperBound
        baseValue = ResolveValue(columns.Item(idx), StartRowSafe())
        compareValue = ResolveValue(columns.Item(idx), endRowValue + 1)
        If StrComp(baseValue, compareValue, vbBinaryCompare) <> 0 Then
            ILLSectionContext_HasNext = False
            Exit Function
        End If
    Next idx
End Function

Private Function ILLSectionContext_NextStartRow(ByVal columnNames As Variant) As Long
    Dim endRowValue As Long

    If Not ILLSectionContext_HasNext(columnNames) Then
        ThrowError ProjectError.ElementNotFound, "No further section matches the provided boundaries"
    End If

    endRowValue = ILLSectionContext_EndRow(columnNames)
    ILLSectionContext_NextStartRow = endRowValue + 1
End Function

Private Function ILLSectionContext_VariableNames(ByVal firstRow As Long, _
                                                 ByVal lastRow As Long) As BetterArray
    Dim names As BetterArray
    Dim variableRange As Range
    Dim rowIndex As Long

    If lastRow < firstRow Then
        ThrowError ProjectError.InvalidArgument, _
                   "Variable range start row must be less than or equal to end row"
    End If

    Set names = New BetterArray
    names.LowerBound = 1

    Set variableRange = FetchColumnRange(VariableColumnName)

    For rowIndex = firstRow To lastRow
        names.Push variableRange.Cells(rowIndex, 1).Value
    Next rowIndex

    Set ILLSectionContext_VariableNames = names
End Function
