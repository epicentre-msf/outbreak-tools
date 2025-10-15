VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseReportManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Maintains disease import report tables without mutating legacy logic"

'@Folder("Msetup")
'@ModuleDescription("Maintains disease import report tables without mutating legacy logic")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseReportManager

'@section Queries
'===============================================================================

'@sub-title Check whether a disease row exists with the expected pending value.
'@param reportTable ListObject hosting the import report entries.
'@param diseaseName String disease identifier to search for.
'@param expectedValue Optional String pending flag, defaults to ""yes"".
'@return Boolean True when the disease requires a report.
Public Function HasPendingReport(ByVal reportTable As ListObject, ByVal diseaseName As String, Optional ByVal expectedValue As String = "yes") As Boolean
    Dim namesColumn As Range
    Dim statusColumn As Range

    ValidateReportTable reportTable
    ValidateDiseaseName diseaseName

    Set namesColumn = ResolveColumn(reportTable, 1)
    Set statusColumn = ResolveColumn(reportTable, 2)

    If namesColumn Is Nothing Or statusColumn Is Nothing Then Exit Function

    HasPendingReport = ColumnContains(namesColumn, statusColumn, diseaseName, expectedValue)
End Function

'@section Commands
'===============================================================================

'@sub-title Remove disease rows from all report tables on the worksheet.
'@param reportSheet Worksheet hosting the import report tables.
'@param diseaseName String disease identifier to purge.
Public Sub ClearReportStatus(ByVal reportSheet As Worksheet, ByVal diseaseName As String)
    Dim reportTable As ListObject
    Dim removalIndexes As BetterArray

    ValidateReportSheet reportSheet
    ValidateDiseaseName diseaseName

    For Each reportTable In reportSheet.ListObjects
        Set removalIndexes = LocateRows(reportTable, diseaseName)
        DeleteRows reportTable, removalIndexes
    Next reportTable
End Sub

'@section Internal Helpers
'===============================================================================

'@sub-title Ensure the report table reference is valid.
Private Sub ValidateReportTable(ByVal reportTable As ListObject)
    If reportTable Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Report table reference cannot be Nothing"
    End If
End Sub

'@sub-title Ensure the report sheet reference is valid.
Private Sub ValidateReportSheet(ByVal reportSheet As Worksheet)
    If reportSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Report sheet reference cannot be Nothing"
    End If
End Sub

'@sub-title Ensure the disease name is provided.
Private Sub ValidateDiseaseName(ByVal diseaseName As String)
    If LenB(diseaseName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name cannot be empty"
    End If
End Sub

'@sub-title Retrieve the Nth column from the table's data body.
Private Function ResolveColumn(ByVal reportTable As ListObject, ByVal columnIndex As Long) As Range
    Dim dataBody As Range

    Set dataBody = reportTable.DataBodyRange
    If dataBody Is Nothing Then Exit Function
    If columnIndex < 1 Or columnIndex > reportTable.ListColumns.Count Then Exit Function

    Set ResolveColumn = dataBody.Columns(columnIndex)
End Function

'@sub-title Determine whether a disease row matches the expected status.
Private Function ColumnContains(ByVal namesColumn As Range, ByVal statusColumn As Range, ByVal diseaseName As String, ByVal expectedValue As String) As Boolean
    Dim namesData As Variant
    Dim statusData As Variant
    Dim rowIndex As Long
    Dim totalRows As Long
    Dim targetName As String
    Dim targetStatus As String

    namesData = ReadColumn(namesColumn)
    statusData = ReadColumn(statusColumn)

    totalRows = UBound(namesData, 1)

    For rowIndex = 1 To totalRows
        targetName = CStr(namesData(rowIndex, 1))
        If StrComp(targetName, diseaseName, vbTextCompare) = 0 Then
            targetStatus = CStr(statusData(rowIndex, 1))
            ColumnContains = (StrComp(Normalize(targetStatus), Normalize(expectedValue), vbBinaryCompare) = 0)
            Exit Function
        End If
    Next rowIndex
End Function

'@sub-title Locate all row indexes belonging to the disease.
Private Function LocateRows(ByVal reportTable As ListObject, ByVal diseaseName As String) As BetterArray
    Dim rowsData As Variant
    Dim rowIndex As Long
    Dim matches As BetterArray
    Dim totalRows As Long

    Set matches = New BetterArray

    If reportTable.ListRows.Count = 0 Then
        Set LocateRows = matches
        Exit Function
    End If

    rowsData = ReadColumn(reportTable.DataBodyRange.Columns(1))
    totalRows = UBound(rowsData, 1)

    For rowIndex = totalRows To 1 Step -1
        If StrComp(CStr(rowsData(rowIndex, 1)), diseaseName, vbTextCompare) = 0 Then
            matches.Push rowIndex
        End If
    Next rowIndex

    Set LocateRows = matches
End Function

'@sub-title Delete the supplied rows from the table, starting from the bottom.
Private Sub DeleteRows(ByVal reportTable As ListObject, ByVal rowIndexes As BetterArray)
    Dim counter As Long
    Dim rowNumber As Long

    If rowIndexes Is Nothing Then Exit Sub
    If rowIndexes.Length = 0 Then Exit Sub

    For counter = rowIndexes.LowerBound To rowIndexes.UpperBound
        rowNumber = CLng(rowIndexes.Item(counter))
        If rowNumber >= 1 And rowNumber <= reportTable.ListRows.Count Then
            reportTable.ListRows(rowNumber).Delete
        End If
    Next counter
End Sub

'@sub-title Normalize text for case-insensitive comparisons.
Private Function Normalize(ByVal value As String) As String
    Normalize = LCase$(Trim$(value))
End Function

'@sub-title Read a single-column range into a 2D Variant matrix.
Private Function ReadColumn(ByVal columnRange As Range) As Variant
    Dim values As Variant

    If columnRange Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Column range cannot be Nothing"
    End If

    values = columnRange.Value
    If IsArray(values) Then
        If UBound(values, 2) = 1 Then
            ReadColumn = values
        Else
            ThrowError ProjectError.InvalidArgument, "Column range must be single-column"
        End If
    Else
        Dim singleValue(1 To 1, 1 To 1) As Variant
        singleValue(1, 1) = values
        ReadColumn = singleValue
    End If
End Function

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseReportManager", errMessage
End Sub

'@section Interface
'===============================================================================

Private Function IDiseaseReportManager_HasPendingReport(ByVal reportTable As ListObject, ByVal diseaseName As String, Optional ByVal expectedValue As String = "yes") As Boolean
    IDiseaseReportManager_HasPendingReport = HasPendingReport(reportTable, diseaseName, expectedValue)
End Function

Private Sub IDiseaseReportManager_ClearReportStatus(ByVal reportSheet As Worksheet, ByVal diseaseName As String)
    ClearReportStatus reportSheet, diseaseName
End Sub
