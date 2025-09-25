Attribute VB_Name = "GraphListObjectUtilities"
Option Explicit

'@Folder("Analysis.Graphs")
'@ModuleDescription("Shared list-object utilities optimised for graph/table specification workflows")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'
'The helper routines centralise common lookups such as column indexing,
'case-insensitive value detection, and unique value extraction so that
'graph-related builders can operate without duplicating the fragile
'worksheet navigation logic used inside legacy classes.

Private Const MODULE_NAME As String = "GraphListObjectUtilities"

'@section Normalisation helpers
'===============================================================================

'Normalise a text key for case-insensitive dictionary usage.
Public Function NormalizeGraphKey(ByVal valueText As String) As String
    NormalizeGraphKey = LCase$(Trim$(valueText))
End Function

'Determine whether a worksheet range contains a text value.
Public Function RangeContainsValue(ByVal targetRange As Range, _
                                   ByVal valueText As String, _
                                   Optional ByVal strictMatch As Boolean = False) As Boolean
    Dim normalizedValue As String
    Dim searchRange As Range

    RangeContainsValue = False

    If targetRange Is Nothing Then Exit Function

    normalizedValue = Trim$(valueText)
    If normalizedValue = vbNullString Then Exit Function

    Set searchRange = targetRange
    If searchRange.Rows.Count = 0 Or searchRange.Columns.Count = 0 Then Exit Function

    If strictMatch Then
        RangeContainsValue = Not (searchRange.Find(What:=normalizedValue, _
                                                   LookAt:=xlWhole, _
                                                   MatchCase:=True) Is Nothing)
    Else
        RangeContainsValue = Not (searchRange.Find(What:=normalizedValue, _
                                                   LookAt:=xlPart, _
                                                   MatchCase:=False) Is Nothing)
    End If
End Function

'@section ListObject helpers
'===============================================================================

'Returns the column index for the requested header within a ListObject.
Public Function ListObjectColumnIndex(ByVal source As ListObject, _
                                      ByVal columnName As String, _
                                      Optional ByVal relativeIndex As Boolean = True) As Long
    Dim headerRange As Range
    Dim hit As Range

    ValidateListObject source, "source"

    ListObjectColumnIndex = -1

    Set headerRange = SafeHeaderRange(source)
    If headerRange Is Nothing Then Exit Function

    Set hit = headerRange.Find(What:=Trim$(columnName), _
                               LookAt:=xlPart, _
                               MatchCase:=False)

    If Not hit Is Nothing Then
        If relativeIndex Then
            ListObjectColumnIndex = hit.Column - headerRange.Column + 1
        Else
            ListObjectColumnIndex = hit.Column
        End If
    End If
End Function

'Extracts unique values from a column and returns them in a BetterArray.
Public Function ListObjectUniqueColumnValues(ByVal source As ListObject, _
                                             ByVal columnName As String) As BetterArray
    Dim index As Long
    Dim dataRange As Range
    Dim values As BetterArray
    Dim seen As Collection
    Dim rowIndex As Long
    Dim cellValue As Variant
    Dim key As String

    ValidateListObject source, "source"

    Set ListObjectUniqueColumnValues = New BetterArray
    ListObjectUniqueColumnValues.LowerBound = 1

    index = ListObjectColumnIndex(source, columnName)
    If index < 1 Then Exit Function

    On Error Resume Next
        Set dataRange = source.ListColumns(index).DataBodyRange
    On Error GoTo 0

    If dataRange Is Nothing Then Exit Function

    Set values = New BetterArray
    values.LowerBound = 1
    Set seen = New Collection

    For rowIndex = 1 To dataRange.Rows.Count
        cellValue = dataRange.Cells(rowIndex, 1).Value
        key = NormalizeGraphKey(CStr(cellValue))

        If (key <> vbNullString) Then
            On Error Resume Next
                seen.Add True, key
            If Err.Number = 0 Then values.Push cellValue
            Err.Clear
            On Error GoTo 0
        End If
    Next rowIndex

    Set ListObjectUniqueColumnValues = values
End Function

'Safely retrieve the header row range, accounting for listobjects without headers.
Private Function SafeHeaderRange(ByVal source As ListObject) As Range
    On Error Resume Next
        Set SafeHeaderRange = source.HeaderRowRange
    On Error GoTo 0
End Function

'Validate that a candidate ListObject is usable.
Private Sub ValidateListObject(ByVal candidate As ListObject, ByVal argumentName As String)
    If candidate Is Nothing Then
        RaiseProjectError CLng(ProjectError.InvalidArgument), _
                          argumentName & " cannot be Nothing"
        Exit Sub
    End If

    If StrComp(TypeName(candidate), "ListObject", vbBinaryCompare) <> 0 Then
        RaiseProjectError CLng(ProjectError.InvalidArgument), _
                          argumentName & " must be a ListObject"
    End If
End Sub

'Centralise error raising to keep message formatting consistent.
Private Sub RaiseProjectError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, MODULE_NAME, messageText
End Sub

