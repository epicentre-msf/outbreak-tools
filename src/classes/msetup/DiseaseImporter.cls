VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Coordinates disease worksheet imports with merge and reporting support"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Coordinates disease worksheet imports with merge and reporting support")
'@depends DiseaseImportSummary, DiseaseLogger, DiseaseSheet, BetterArray
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'Merges a foreign disease table into an existing one, existing or foreign
'values winning by priority, and answers a DiseaseImportSummary of the merge.

'@section Enumerations
'===============================================================================

Public Enum DiseaseImportPriority
    DiseaseImportPriority_Existing = 0
    DiseaseImportPriority_Foreign = 1
End Enum

'The variable name keys every merge. DiseaseSheet builds it as column 2
'of the disease table (Order, Name, Section, ...).
Private Const VARIABLE_NAME_COLUMN As Long = 2

'@section Factory
'===============================================================================

'@sub-title Build a fresh importer.
'@return DiseaseImporter ready to merge disease tables.
Public Function Create() As DiseaseImporter
    Set Create = New DiseaseImporter
End Function

'@section Operations
'===============================================================================

'@sub-title Import disease data, merging when requested.
'@param priority Byte holding a DiseaseImportPriority value. Defaults to foreign.
Public Function MergeDisease(ByVal targetTable As ListObject, ByVal sourceTable As ListObject, _
                             Optional ByVal mergeValues As Boolean = True, _
                             Optional ByVal priority As Byte = DiseaseImportPriority_Foreign, _
                             Optional ByVal logger As DiseaseLogger = Nothing) As DiseaseImportSummary
    Dim summary As DiseaseImportSummary

    ValidateTable targetTable, "targetTable"
    ValidateTable sourceTable, "sourceTable"

    Set summary = DiseaseImportSummary.Create()

    If Not mergeValues Then
        ReplaceTable targetTable, sourceTable, summary, logger
    Else
        MergeTables targetTable, sourceTable, priority, summary, logger
    End If

    Set MergeDisease = summary
End Function

'@section Replacement
'===============================================================================

'@sub-title Replace the target table with the source content.
Private Sub ReplaceTable(ByVal targetTable As ListObject, ByVal sourceTable As ListObject, ByVal summary As DiseaseImportSummary, ByVal logger As DiseaseLogger)
    Dim sourceRows As Long
    Dim sourceData As Variant
    Dim columnCount As Long
    Dim rowIndex As Long
    Dim variableName As String

    columnCount = targetTable.ListColumns.Count
    If sourceTable.ListColumns.Count <> columnCount Then
        ThrowError ProjectError.InvalidArgument, "Source table structure does not match target table"
    End If

    ClearTable targetTable

    sourceRows = sourceTable.ListRows.Count
    If sourceRows = 0 Then Exit Sub

    sourceData = Ensure2D(sourceTable.DataBodyRange.Value, columnCount)
    AppendBlankRows targetTable, sourceRows
    targetTable.DataBodyRange.Value = sourceData

    For rowIndex = 1 To sourceRows
        variableName = CStr(sourceData(rowIndex, VARIABLE_NAME_COLUMN))
        If LenB(variableName) > 0 Then
            summary.RegisterAppended variableName
            Log logger, "ImportReplace", DiseaseLogInfo, "Appended variable during replace", BuildDetails("Variable", variableName)
        End If
    Next rowIndex
End Sub

'@section Merge
'===============================================================================

'@sub-title Merge the source data into the target according to the priority.
Private Sub MergeTables(ByVal targetTable As ListObject, ByVal sourceTable As ListObject, ByVal priority As Byte, ByVal summary As DiseaseImportSummary, ByVal logger As DiseaseLogger)
    Dim targetRows As Long
    Dim sourceRows As Long
    Dim columnCount As Long
    Dim targetData As Variant
    Dim sourceData As Variant
    Dim targetNames() As String
    Dim targetIndex As Collection
    Dim seenSource As Collection
    Dim appendedRows As BetterArray

    columnCount = targetTable.ListColumns.Count
    If sourceTable.ListColumns.Count <> columnCount Then
        ThrowError ProjectError.InvalidArgument, "Source table structure does not match target table"
    End If

    targetRows = targetTable.ListRows.Count
    sourceRows = sourceTable.ListRows.Count

    If targetRows > 0 Then targetData = Ensure2D(targetTable.DataBodyRange.Value, columnCount)
    If sourceRows > 0 Then sourceData = Ensure2D(sourceTable.DataBodyRange.Value, columnCount)

    targetNames = BuildTargetIndex(targetData, targetRows, targetIndex)
    Set seenSource = New Collection
    'The rows pushed here are 2D arrays based at 1, and Push re-bases any
    'array whose bound disagrees with the container's. The container has to
    'start at 1 too, or the readers index 1 into a row that starts at 0.
    Set appendedRows = New BetterArray
    appendedRows.LowerBound = 1

    ApplyMerge targetData, sourceData, targetRows, sourceRows, columnCount, priority, targetIndex, seenSource, appendedRows, summary, logger

    If targetRows > 0 And priority = DiseaseImportPriority_Foreign Then
        targetTable.DataBodyRange.Value = targetData
    End If

    AppendImportedRows targetTable, appendedRows, logger
    RegisterMissingVariables targetData, targetRows, targetNames, seenSource, summary, logger
End Sub

'@sub-title Build lookup structures for target variables.
Private Function BuildTargetIndex(ByVal targetData As Variant, ByVal targetRows As Long, ByRef targetIndex As Collection) As String()
    Dim names() As String
    Dim rowIndex As Long
    Dim key As String

    Set targetIndex = New Collection

    If targetRows > 0 Then
        ReDim names(1 To targetRows)

        For rowIndex = 1 To targetRows
            key = Normalize(CStr(targetData(rowIndex, VARIABLE_NAME_COLUMN)))
            names(rowIndex) = key
            If LenB(key) > 0 Then
                AddIfMissing targetIndex, key, rowIndex
            End If
        Next rowIndex
    End If

    BuildTargetIndex = names
End Function

'@sub-title Perform the merge between target and source matrices.
Private Sub ApplyMerge(ByRef targetData As Variant, ByRef sourceData As Variant, ByVal targetRows As Long, ByVal sourceRows As Long, _
                       ByVal columnCount As Long, ByVal priority As Byte, ByVal targetIndex As Collection, _
                       ByVal seenSource As Collection, ByVal appendedRows As BetterArray, ByVal summary As DiseaseImportSummary, ByVal logger As DiseaseLogger)

    Dim rowIndex As Long
    Dim colIndex As Long
    Dim variableName As String
    Dim normalizedName As String
    Dim targetRow As Long
    Dim appendedRow As Variant

    If sourceRows = 0 Then Exit Sub

    For rowIndex = 1 To sourceRows
        variableName = CStr(sourceData(rowIndex, VARIABLE_NAME_COLUMN))
        normalizedName = Normalize(variableName)
        If LenB(normalizedName) = 0 Then GoTo ContinueLoop

        targetRow = LookupRow(targetIndex, normalizedName)
        If targetRow > 0 Then
            AddIfMissingValue seenSource, normalizedName

            If priority = DiseaseImportPriority_Foreign And targetRows > 0 Then
                For colIndex = 1 To columnCount
                    targetData(targetRow, colIndex) = sourceData(rowIndex, colIndex)
                Next colIndex
                summary.RegisterUpdated variableName
                Log logger, "ImportMerge", DiseaseLogInfo, "Updated variable from import", BuildDetails("Variable", variableName)
            End If
        Else
            appendedRow = BuildRow(sourceData, rowIndex, columnCount)
            appendedRows.Push appendedRow
            summary.RegisterAppended variableName
            Log logger, "ImportMerge", DiseaseLogInfo, "Appended new variable", BuildDetails("Variable", variableName)
        End If

ContinueLoop:
    Next rowIndex
End Sub

'@sub-title Append merged rows to the target ListObject.
Private Sub AppendImportedRows(ByVal targetTable As ListObject, ByVal appendedRows As BetterArray, ByVal logger As DiseaseLogger)
    Dim appendCount As Long
    Dim columnCount As Long
    Dim appendedMatrix As Variant
    Dim startRow As Long
    Dim freeRows As Long
    Dim rowIndex As Long
    Dim rowValues As Variant

    If appendedRows Is Nothing Then Exit Sub
    appendCount = appendedRows.Length
    If appendCount = 0 Then Exit Sub

    columnCount = targetTable.ListColumns.Count

    'The rows land under the last named line. A built disease table carries
    'blank lines, so those are filled before any row is added.
    startRow = LastNamedRow(targetTable) + 1
    freeRows = targetTable.ListRows.Count - startRow + 1
    If freeRows < appendCount Then AppendBlankRows targetTable, appendCount - freeRows

    appendedMatrix = ConvertRowsToMatrix(appendedRows, appendCount, columnCount)

    targetTable.DataBodyRange.Rows(startRow).Resize(appendCount, columnCount).Value = appendedMatrix

    If Not logger Is Nothing Then
        For rowIndex = appendedRows.LowerBound To appendedRows.UpperBound
            rowValues = appendedRows.Item(rowIndex)
            Log logger, "ImportAppend", DiseaseLogInfo, "Appended new row from import", BuildDetails("Variable", CStr(rowValues(1, VARIABLE_NAME_COLUMN)))
        Next rowIndex
    End If
End Sub

'@sub-title Determine which existing variables were missing from the import.
Private Sub RegisterMissingVariables(ByVal targetData As Variant, ByVal targetRows As Long, ByRef targetNames() As String, ByVal seenSource As Collection, ByVal summary As DiseaseImportSummary, ByVal logger As DiseaseLogger)
    Dim rowIndex As Long
    Dim normalizedName As String

    If targetRows = 0 Then Exit Sub

    For rowIndex = 1 To targetRows
        normalizedName = vbNullString
        If Not IsEmpty(targetNames) Then
            normalizedName = targetNames(rowIndex)
        End If
        If LenB(normalizedName) = 0 Then GoTo NextRow

        If Not CollectionContains(seenSource, normalizedName) Then
            summary.RegisterMissing CStr(targetData(rowIndex, VARIABLE_NAME_COLUMN))
            Log logger, "ImportMerge", DiseaseLogWarning, "Variable missing from import", BuildDetails("Variable", CStr(targetData(rowIndex, VARIABLE_NAME_COLUMN)))
        End If

NextRow:
    Next rowIndex
End Sub

'@section Table Utilities
'===============================================================================

'@sub-title Remove all rows from the ListObject.
Private Sub ClearTable(ByVal targetTable As ListObject)
    Do While targetTable.ListRows.Count > 0
        targetTable.ListRows(targetTable.ListRows.Count).Delete
    Loop
End Sub

'@sub-title The index of the last line carrying a variable name, 0 when none does.
Private Function LastNamedRow(ByVal targetTable As ListObject) As Long
    Dim names As Variant
    Dim rowIndex As Long

    If targetTable.DataBodyRange Is Nothing Then Exit Function

    names = targetTable.DataBodyRange.Columns(VARIABLE_NAME_COLUMN).Value2
    If Not IsArray(names) Then
        If LenB(CleanName(names)) > 0 Then LastNamedRow = 1
        Exit Function
    End If

    For rowIndex = UBound(names, 1) To 1 Step -1
        If LenB(CleanName(names(rowIndex, 1))) > 0 Then
            LastNamedRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

'@sub-title A cell value as trimmed text; an error value reads as empty.
Private Function CleanName(ByVal candidate As Variant) As String
    If IsError(candidate) Then Exit Function
    CleanName = Trim$(CStr(candidate))
End Function

'@sub-title Append blank rows to the target table.
'@details A row added through ListRows.Add takes the vertical edges of the
'row above and no horizontal edge, so the disease frame is painted again
'over the grown table.
Private Sub AppendBlankRows(ByVal targetTable As ListObject, ByVal rowCount As Long)
    Dim counter As Long

    If rowCount <= 0 Then Exit Sub

    For counter = 1 To rowCount
        targetTable.ListRows.Add
    Next counter

    DiseaseSheet.FrameTable targetTable
End Sub

'@sub-title Create a row array from the source matrix.
Private Function BuildRow(ByRef sourceData As Variant, ByVal rowIndex As Long, ByVal columnCount As Long) As Variant
    Dim rowValues() As Variant
    Dim colIndex As Long

    ReDim rowValues(1 To 1, 1 To columnCount)
    For colIndex = 1 To columnCount
        rowValues(1, colIndex) = sourceData(rowIndex, colIndex)
    Next colIndex

    BuildRow = rowValues
End Function

'@sub-title Convert a BetterArray of rows into a contiguous matrix.
Private Function ConvertRowsToMatrix(ByVal rows As BetterArray, ByVal appendCount As Long, ByVal columnCount As Long) As Variant
    Dim matrix() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowValues As Variant

    ReDim matrix(1 To appendCount, 1 To columnCount)

    For rowIndex = 1 To appendCount
        rowValues = rows.Item(rowIndex - 1 + rows.LowerBound)
        For colIndex = 1 To columnCount
            matrix(rowIndex, colIndex) = rowValues(1, colIndex)
        Next colIndex
    Next rowIndex

    ConvertRowsToMatrix = matrix
End Function

'@sub-title Convert single values into 2D matrices when necessary.
Private Function Ensure2D(ByVal values As Variant, ByVal columnCount As Long) As Variant
    Dim matrix() As Variant

    If IsArray(values) Then
        Ensure2D = values
    Else
        ReDim matrix(1 To 1, 1 To columnCount)
        matrix(1, 1) = values
        Ensure2D = matrix
    End If
End Function

'@section Collections
'===============================================================================

'@sub-title Add row index to collection if the key is missing.
Private Sub AddIfMissing(ByVal index As Collection, ByVal key As String, ByVal rowIndex As Long)
    On Error Resume Next
        index.Add rowIndex, key
    On Error GoTo 0
End Sub

'@sub-title Register presence of a source variable.
Private Sub AddIfMissingValue(ByVal seenSource As Collection, ByVal key As String)
    On Error Resume Next
        seenSource.Add True, key
    On Error GoTo 0
End Sub

'@sub-title Read a row index from the collection.
Private Function LookupRow(ByVal index As Collection, ByVal key As String) As Long
    If index Is Nothing Then Exit Function
    On Error Resume Next
        LookupRow = CLng(index(key))
        If Err.Number <> 0 Then
            LookupRow = 0
            Err.Clear
        End If
    On Error GoTo 0
End Function

'@sub-title Test whether the collection contains the supplied key.
Private Function CollectionContains(ByVal values As Collection, ByVal key As String) As Boolean
    If values Is Nothing Then Exit Function
    On Error Resume Next
        values.Item key
        CollectionContains = (Err.Number = 0)
        Err.Clear
    On Error GoTo 0
End Function

'@section Validation
'===============================================================================

'@sub-title Guard against invalid ListObject references.
Private Sub ValidateTable(ByVal table As ListObject, ByVal argumentName As String)
    If table Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, argumentName & " cannot be Nothing"
    End If
    If table.ListColumns.Count = 0 Then
        ThrowError ProjectError.InvalidArgument, argumentName & " must contain at least one column"
    End If
End Sub

'@section Text Helpers
'===============================================================================

'@sub-title Normalise identifiers for comparison.
Private Function Normalize(ByVal value As String) As String
    Normalize = LCase$(Trim$(value))
End Function

Private Sub Log(ByVal logger As DiseaseLogger, ByVal operation As String, ByVal severity As Byte, ByVal message As String, Optional ByVal details As BetterArray)
    If logger Is Nothing Then Exit Sub
    logger.Record operation, severity, message, details
End Sub

Private Function BuildDetails(ByVal key As String, ByVal value As String) As BetterArray
    Dim details As BetterArray
    Set details = New BetterArray
    details.lowerBound = 0
    details.Push key
    details.Push value
    Set BuildDetails = details
End Function

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseImporter", errMessage
End Sub

