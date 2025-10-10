New value_of implementation

```vb
Public Function VALUE_OF(rng As Range, lookupSheetName As String, colLookupIndex As Long, colValueIndex As Long) As Variant

    Application.Volatile

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lookupValue As Variant
    Dim lookupColumn As Range
    Dim valueColumn As Range
    Dim matchPos As Variant

    lookupValue = rng.Value

    If lookupValue = vbNullString Then
        VALUE_OF = vbNullString
        Exit Function
    End If
    
    If Trim$(lookupSheetName) = vbNullString Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(lookupSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        VALUE_OF = vbNullString
        Exit Function
    End If 

    If ws.ListObjects.Count = 0 Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    Set lo = ws.ListObjects(1)

    If colLookupIndex < 1 Or colLookupIndex > lo.ListColumns.Count Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    If colValueIndex < 1 Or colValueIndex > lo.ListColumns.Count Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    If lo.ListColumns(colLookupIndex).DataBodyRange Is Nothing Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    If lo.ListColumns(colValueIndex).DataBodyRange Is Nothing Then
        VALUE_OF = vbNullString
        Exit Function
    End If

    Set lookupColumn = lo.ListColumns(colLookupIndex).DataBodyRange
    Set valueColumn = lo.ListColumns(colValueIndex).DataBodyRange

    matchPos = Application.Match(lookupValue, lookupColumn, 0)

    If IsError(matchPos) Then
        VALUE_OF = vbNullString
    Else
        VALUE_OF = valueColumn.Cells(matchPos, 1).Value
    End If
End Function
```

