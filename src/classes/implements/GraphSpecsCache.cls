VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSpecsCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Indexes graph list-object rows for fast lookups by graph identifier")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IGraphSpecsCache

Private Type TCacheState
    graphList As ListObject
    graphBuckets As Collection
    graphIds As BetterArray
    columnKeys As BetterArray
    columnIndexes As BetterArray
    isPrimed As Boolean
End Type

Private this As TCacheState
Private Const CLASSNAME As String = "GraphSpecsCache"

'@section Factory
'===============================================================================
Public Function Create(ByVal graphList As ListObject) As IGraphSpecsCache
    ValidateListObject graphList, "graphList"

    With New GraphSpecsCache
        Set .GraphList = graphList
        .Refresh
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IGraphSpecsCache
    Set Self = Me
End Property

'@section Accessors
'===============================================================================
Public Property Get GraphList() As ListObject
    Set GraphList = this.graphList
End Property

Public Property Set GraphList(ByVal graphList As ListObject)
    ValidateListObject graphList, "graphList"
    Set this.graphList = graphList
    InvalidateCache
End Property

Public Function GraphIds() As BetterArray
    EnsurePrimed
    Set GraphIds = this.graphIds.Clone()
End Function

Public Function ColumnValues(ByVal graphId As String, ByVal columnName As String) As BetterArray
    Dim bucket As Collection
    Dim values As BetterArray
    Dim key As String

    EnsurePrimed

    Set ColumnValues = New BetterArray
    ColumnValues.LowerBound = 1

    key = NormalizeKey(graphId)
    If key = vbNullString Then Exit Function

    Set bucket = LookupGraphBucket(key)
    If bucket Is Nothing Then Exit Function

    On Error Resume Next
        Set values = bucket.Item(NormalizeKey(columnName))
    On Error GoTo 0

    If values Is Nothing Then Exit Function

    Set ColumnValues = values.Clone()
End Function

'@section Cache management
'===============================================================================
Public Sub Invalidate()
    InvalidateCache
End Sub

Public Sub Refresh()
    InvalidateCache
    PrimeCache
End Sub

'@section Internal helpers
'===============================================================================
Private Sub EnsurePrimed()
    If Not this.isPrimed Then PrimeCache
End Sub

Private Sub InvalidateCache()
    Set this.graphBuckets = Nothing
    Set this.graphIds = Nothing
    Set this.columnKeys = Nothing
    Set this.columnIndexes = Nothing
    this.isPrimed = False
End Sub

Private Sub PrimeCache()
    Dim dataRange As Range
    Dim headerRow As Range
    Dim column As ListColumn
    Dim rowIndex As Long
    Dim graphId As String
    Dim graphKey As String
    Dim columnIdx As Long
    Dim columnKey As String
    Dim bucket As Collection
    Dim values As BetterArray
    Dim rawValue As Variant

    ValidateListObject this.graphList, "graphList"

    Set this.graphBuckets = New Collection
    Set this.graphIds = New BetterArray
    this.graphIds.LowerBound = 1
    Set this.columnKeys = New BetterArray
    this.columnKeys.LowerBound = 1
    Set this.columnIndexes = New BetterArray
    this.columnIndexes.LowerBound = 1

    Set headerRow = this.graphList.HeaderRowRange
    For Each column In this.graphList.ListColumns
        columnKey = NormalizeKey(CStr(column.Name))
        If columnKey <> vbNullString Then
            this.columnKeys.Push columnKey
            this.columnIndexes.Push CLng(column.Index)
        End If
    Next column

    On Error Resume Next
        Set dataRange = this.graphList.DataBodyRange
    On Error GoTo 0

    If dataRange Is Nothing Then
        this.isPrimed = True
        Exit Sub
    End If

    Dim graphColumnIndex As Long
    graphColumnIndex = ColumnIndexFor("graph id")
    If graphColumnIndex <= 0 Then
        ThrowError CLng(ProjectError.ElementNotFound), "graph id column missing in graph list"
    End If

    For rowIndex = 1 To dataRange.Rows.Count
        graphId = CStr(dataRange.Cells(rowIndex, graphColumnIndex).Value)
        graphKey = NormalizeKey(graphId)
        If graphKey <> vbNullString Then
            Set bucket = EnsureGraphBucket(graphKey)
            If Not GraphIdRecorded(graphKey) Then this.graphIds.Push graphId

            For columnIdx = this.columnKeys.LowerBound To this.columnKeys.UpperBound
                columnKey = CStr(this.columnKeys.Item(columnIdx))
                rawValue = dataRange.Cells(rowIndex, CLng(this.columnIndexes.Item(columnIdx))).Value
                Set values = EnsureColumnBucket(bucket, columnKey)
                values.Push rawValue
            Next columnIdx
        End If
    Next rowIndex

    this.isPrimed = True
End Sub

Private Function EnsureGraphBucket(ByVal graphKey As String) As Collection
    Dim bucket As Collection

    On Error Resume Next
        Set bucket = this.graphBuckets.Item(graphKey)
    On Error GoTo 0

    If bucket Is Nothing Then
        Set bucket = New Collection
        On Error Resume Next
            this.graphBuckets.Add bucket, graphKey
        On Error GoTo 0
    End If

    Set EnsureGraphBucket = bucket
End Function

Private Function LookupGraphBucket(ByVal graphKey As String) As Collection
    On Error Resume Next
        Set LookupGraphBucket = this.graphBuckets.Item(graphKey)
    On Error GoTo 0
End Function

Private Function EnsureColumnBucket(ByVal bucket As Collection, ByVal columnKey As String) As BetterArray
    Dim values As BetterArray

    On Error Resume Next
        Set values = bucket.Item(columnKey)
    On Error GoTo 0

    If values Is Nothing Then
        Set values = New BetterArray
        values.LowerBound = 1
        On Error Resume Next
            bucket.Add values, columnKey
        On Error GoTo 0
    End If

    Set EnsureColumnBucket = values
End Function

Private Function GraphIdRecorded(ByVal graphKey As String) As Boolean
    Dim idx As Long

    GraphIdRecorded = False

    If this.graphIds Is Nothing Then Exit Function
    If this.graphIds.Length = 0 Then Exit Function

    For idx = this.graphIds.LowerBound To this.graphIds.UpperBound
        If NormalizeKey(CStr(this.graphIds.Item(idx))) = graphKey Then
            GraphIdRecorded = True
            Exit Function
        End If
    Next idx
End Function

Private Function ColumnIndexFor(ByVal columnName As String) As Long
    Dim idx As Long
    Dim key As String

    ColumnIndexFor = -1

    If this.columnKeys Is Nothing Then Exit Function
    If this.columnKeys.Length = 0 Then Exit Function

    For idx = this.columnKeys.LowerBound To this.columnKeys.UpperBound
        key = CStr(this.columnKeys.Item(idx))
        If key = NormalizeKey(columnName) Then
            ColumnIndexFor = CLng(this.columnIndexes.Item(idx))
            Exit Function
        End If
    Next idx
End Function

Private Function NormalizeKey(ByVal valueText As String) As String
    NormalizeKey = LCase$(Trim$(valueText))
End Function

Private Sub ValidateListObject(ByVal candidate As ListObject, ByVal argumentName As String)
    If candidate Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
        Exit Sub
    End If

    If StrComp(TypeName(candidate), "ListObject", vbBinaryCompare) <> 0 Then
        ThrowError CLng(ProjectError.InvalidArgument), argumentName & " must be a ListObject"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function IGraphSpecsCache_Create(ByVal graphList As ListObject) As IGraphSpecsCache
    Set IGraphSpecsCache_Create = Create(graphList)
End Function

Private Property Get IGraphSpecsCache_Self() As IGraphSpecsCache
    Set IGraphSpecsCache_Self = Self
End Property

Private Property Get IGraphSpecsCache_GraphList() As ListObject
    Set IGraphSpecsCache_GraphList = GraphList
End Property

Private Property Set IGraphSpecsCache_GraphList(ByVal graphList As ListObject)
    Set Me.GraphList = graphList
End Property

Private Function IGraphSpecsCache_GraphIds() As BetterArray
    Set IGraphSpecsCache_GraphIds = GraphIds()
End Function

Private Function IGraphSpecsCache_ColumnValues(ByVal graphId As String, ByVal columnName As String) As BetterArray
    Set IGraphSpecsCache_ColumnValues = ColumnValues(graphId, columnName)
End Function

Private Sub IGraphSpecsCache_Invalidate()
    Invalidate
End Sub

Private Sub IGraphSpecsCache_Refresh()
    Refresh
End Sub

