VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableLayoutPlanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Computes layout coordinates for cross tables without mutating worksheets"

'@Folder("Implement")
'@ModuleDescription("Compute deterministic start cells and footprints for cross tables")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTableLayoutPlanner
    specs As ITablesSpecs
    targetSheet As Worksheet
End Type

Private this As TCrossTableLayoutPlanner
Private Const CLASSNAME As String = "CrossTableLayoutPlanner"
Private Const STANDARD_START_COLUMN As Long = 3
Private Const MINIMUM_START_ROW As Long = 6
Private Const TIME_SERIES_ROW_COUNT As Long = 56

'@section Factory
'===============================================================================

'@details Factory method wiring specifications and worksheet into a planner instance.
Public Function Create(ByVal specs As ITablesSpecs, ByVal targetSheet As Worksheet) As CrossTableLayoutPlanner
    Dim planner As CrossTableLayoutPlanner

    If specs Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Table specifications are required"
    End If

    If targetSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Target worksheet is required"
    End If

    Set planner = New CrossTableLayoutPlanner
    Set planner.Specifications = specs
    Set planner.TargetSheet = targetSheet
    Set Create = planner
End Function

'@section Dependencies
'===============================================================================

'@details Associates the planner with a table specification source.
Public Property Set Specifications(ByVal specs As ITablesSpecs)
    Set this.specs = specs
End Property

Public Property Get Specifications() As ITablesSpecs
    Set Specifications = this.specs
End Property

'@details Associates the planner with the worksheet that will host the cross table.
Public Property Set TargetSheet(ByVal sh As Worksheet)
    Set this.targetSheet = sh
End Property

Public Property Get TargetSheet() As Worksheet
    Set TargetSheet = this.targetSheet
End Property

'@section Public API
'===============================================================================

'@details Computes the layout summary (start coordinates, spans) without altering the worksheet.
Public Function ComputeLayout() As CrossTableLayoutSummary
    Dim summary As CrossTableLayoutSummary
    Dim startRow As Long
    Dim startColumn As Long
    Dim headerRows As Long
    Dim dataRows As Long
    Dim dataColumns As Long

    startRow = ResolveStartRow
    startColumn = ResolveStartColumn(startRow)
    headerRows = EstimateHeaderRows
    dataRows = EstimateDataRows
    dataColumns = EstimateDataColumns

    Set summary = New CrossTableLayoutSummary
    summary.Initialise startRow, startColumn, headerRows, dataRows, dataColumns
    Set ComputeLayout = summary
End Function

'@section Internal helpers
'===============================================================================

Private Function ResolveStartRow() As Long
    Dim tabId As String
    Dim existingRow As Long
    Dim tableType As Byte
    Dim newSection As Boolean

    tabId = Trim$(Specifications.TableId)
    If TryGetNamedRow("STARTROW_" & tabId, existingRow) Then
        ResolveStartRow = existingRow
        Exit Function
    End If

    tableType = CByte(Specifications.TableType)
    newSection = Specifications.isNewSection

    If Not newSection And IsTimeSeriesType(tableType) Then
        Dim previousSpecs As ITablesSpecs
        Set previousSpecs = Specifications.Previous

        If Not previousSpecs Is Nothing Then
            Dim previousPlanner As CrossTableLayoutPlanner
            Dim previousSummary As CrossTableLayoutSummary

            Set previousPlanner = CrossTableLayoutPlanner.Create(previousSpecs, TargetSheet)
            Set previousSummary = previousPlanner.ComputeLayout()
            ResolveStartRow = previousSummary.StartRow
            Exit Function
        End If
    End If

    ResolveStartRow = CalculateStartRowForSection(tableType)
End Function

Private Function ResolveStartColumn(ByVal startRow As Long) As Long
    Dim tabId As String
    Dim existingColumn As Long
    Dim tableType As Byte

    tabId = Trim$(Specifications.TableId)
    If TryGetNamedColumn("STARTCOL_" & tabId, existingColumn) Then
        ResolveStartColumn = existingColumn
        Exit Function
    End If

    tableType = CByte(Specifications.TableType)

    'Existing time-series sections append horizontally until a new section starts.
    If Not Specifications.isNewSection And IsTimeSeriesType(tableType) Then
        ResolveStartColumn = LastUsedColumn(startRow + 1)
    Else
        ResolveStartColumn = STANDARD_START_COLUMN
    End If
End Function

Private Function EstimateHeaderRows() As Long
    Dim tableType As Byte

    tableType = CByte(Specifications.TableType)

    Select Case tableType
        Case TypeGlobalSummary
            EstimateHeaderRows = 1
        Case TypeTimeSeries, TypeSpatioTemporal
            EstimateHeaderRows = 3
        Case Else
            EstimateHeaderRows = 2
    End Select
End Function

Private Function EstimateDataRows() As Long
    Dim tableType As Byte
    Dim count As Long

    tableType = CByte(Specifications.TableType)

    If IsTimeSeriesType(tableType) Then
        EstimateDataRows = TIME_SERIES_ROW_COUNT
        Exit Function
    End If

    count = CountBetterArray(Specifications.RowCategories)
    If count = 0 Then count = 1

    If Specifications.HasTotal Then count = count + 1
    If Specifications.HasMissing Then count = count + 1

    EstimateDataRows = count
End Function

Private Function EstimateDataColumns() As Long
    Dim count As Long

    count = CountBetterArray(Specifications.ColumnCategories)
    If count = 0 Then count = 1

    If Specifications.HasTotal Then count = count + 1
    If Specifications.HasPercentage Then count = count + 1

    EstimateDataColumns = count
End Function

Private Function IsTimeSeriesType(ByVal tableType As Byte) As Boolean
    IsTimeSeriesType = (tableType = TypeTimeSeries) Or (tableType = TypeSpatioTemporal)
End Function

Private Function CalculateStartRowForSection(ByVal tableType As Byte) As Long
    Dim baseRow As Long

    'Start immediately after the last used row in the anchor column.
    baseRow = LastUsedRow(STANDARD_START_COLUMN) + 1
    If baseRow < MINIMUM_START_ROW Then baseRow = MINIMUM_START_ROW

    'Legacy layouts reserve additional spacing depending on table type.
    Select Case tableType
        Case TypeSpatial
            baseRow = baseRow + 6
        Case TypeGlobalSummary
            'No shift required for global summary tables.
        Case Else
            baseRow = baseRow + 4
    End Select

    CalculateStartRowForSection = baseRow
End Function

Private Function LastUsedRow(ByVal columnIndex As Long) As Long
    Dim sheet As Worksheet
    Dim lastRow As Long

    Set sheet = TargetSheet
    lastRow = sheet.Cells(sheet.Rows.Count, columnIndex).End(xlUp).Row
    LastUsedRow = lastRow
End Function

Private Function LastUsedColumn(ByVal rowIndex As Long) As Long
    Dim sheet As Worksheet
    Dim lastColumn As Long

    Set sheet = TargetSheet
    If rowIndex < 1 Then rowIndex = 1

    lastColumn = sheet.Cells(rowIndex, sheet.Columns.Count).End(xlToLeft).Column
    If lastColumn < STANDARD_START_COLUMN Then lastColumn = STANDARD_START_COLUMN

    LastUsedColumn = lastColumn
End Function

Private Function TryGetNamedRow(ByVal nameText As String, ByRef rowIndex As Long) As Boolean
    Dim namedRange As Range

    Set namedRange = ResolveNamedRange(nameText)
    If namedRange Is Nothing Then Exit Function

    rowIndex = namedRange.Row
    TryGetNamedRow = True
End Function

Private Function TryGetNamedColumn(ByVal nameText As String, ByRef startColumn As Long) As Boolean
    Dim namedRange As Range

    Set namedRange = ResolveNamedRange(nameText)
    If namedRange Is Nothing Then Exit Function

    startColumn = namedRange.Column - 1
    If startColumn < 1 Then startColumn = 1
    TryGetNamedColumn = True
End Function

Private Function ResolveNamedRange(ByVal nameText As String) As Range
    Dim sheet As Worksheet
    Dim wb As Workbook
    Dim rng As Range

    Set sheet = TargetSheet
    Set wb = sheet.Parent

    On Error Resume Next
        Set rng = sheet.Range(nameText)
    On Error GoTo 0

    If rng Is Nothing Then
        On Error Resume Next
            Set rng = wb.Names(nameText).RefersToRange
        On Error GoTo 0
    End If

    Set ResolveNamedRange = rng
End Function

Private Function CountBetterArray(ByVal items As BetterArray) As Long
    If items Is Nothing Then Exit Function

    On Error Resume Next
        CountBetterArray = items.Length
    On Error GoTo 0
End Function

End Class
