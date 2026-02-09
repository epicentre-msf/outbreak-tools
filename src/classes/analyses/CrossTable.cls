VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Create and format cross-tables for analysis"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Create and format cross-tables for analysis")
'@IgnoreModule UnrecognizedAnnotation, IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ICrossTable

Private Type TCrossTable
    specs As ITableSpecs
    outputSh As Worksheet
    trans As ITranslationObject
    endRw As Long
    endCol As Long
    nbCols As Long
    tsHeader As String                           'Header for time series (used in go-to section)
End Type

Private this As TCrossTable
Private Const STANDARD_START_COL As Long = 3
Private Const CLASS_NAME As String = "CrossTable"
Private Const NB_ROWS_TIME_SERIES As Long = 56
Private allRangeNames As BetterArray

'@section Factory
'===============================================================================

' @description Create a cross-table from table specifications
' @param specs Table specification row from setup
' @param outputWksh Worksheet where the table will be built
' @param trans Translation object for labels
Public Function Create(ByVal specs As ITableSpecs, _
                       ByVal outputWksh As Worksheet, _
                       ByVal trans As ITranslationObject) As ICrossTable

    If specs Is Nothing Then ThrowError InvalidArgument, "Table specifications not defined"
    If outputWksh Is Nothing Then ThrowError InvalidArgument, "Output worksheet not defined"
    If trans Is Nothing Then ThrowError InvalidArgument, "Translations not defined"

    With New CrossTable
        Set .Specifications = specs
        Set .Wksh = outputWksh
        Set .Translations = trans
        Set Create = .Self
    End With
End Function

'@section Internal properties (factory use)
'===============================================================================

Public Property Get Self() As ICrossTable
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.outputSh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.outputSh = sh
End Property

Public Property Get Specifications() As ITableSpecs
    Set Specifications = this.specs
End Property

Public Property Set Specifications(ByVal specs As ITableSpecs)
    Set this.specs = specs
End Property

Public Property Get Translations() As ITranslationObject
    Set Translations = this.trans
End Property

Public Property Set Translations(ByVal trans As ITranslationObject)
    Set this.trans = trans
End Property

Private Property Get NumberOfColumns() As Long
    NumberOfColumns = this.nbCols
End Property

Private Property Let NumberOfColumns(ByVal col As Long)
    this.nbCols = col
End Property

Private Property Get NamedRangesList() As BetterArray
    Set NamedRangesList = allRangeNames.Clone()
End Property

'@section Position properties
'===============================================================================

Private Function RangeExists(ByVal rngName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Wksh().Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

' @description Starting row for this table on the output worksheet.
'              For non-new-section temporal tables, uses previous table's start row.
'              Bug #3 fix: if Previous start row can't be resolved (section infrastructure
'              missing because previous table was invalid/not built), falls back to
'              computing a fresh start row as if this were a new section.
Private Property Get StartRow() As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim strtRow As Long
    Dim tabId As String

    Set specs = Specifications()
    tabId = specs.TableId
    Set sh = Wksh()

    With sh
        If Not RangeExists("STARTROW_" & tabId) Then
            If specs.IsNewSection() Or _
               (specs.TableType <> TypeTimeSeries And specs.TableType <> TypeSpatioTemporal) Then
                ' New section or non-temporal: compute from last used row
                strtRow = .Cells(.Rows.Count, STANDARD_START_COL).End(xlUp).Row + 1
                strtRow = IIf(strtRow = 2, 6, strtRow)

                strtRow = Switch(specs.TableType = TypeSpatial, strtRow + 6, _
                                 specs.TableType = TypeGlobalSummary, strtRow, _
                                 True, strtRow + 4)

                .Cells(strtRow, STANDARD_START_COL).Name = "STARTROW_" & tabId

            ElseIf Not specs.IsNewSection() Then
                ' Non-new-section temporal: use previous table's start row
                ' Bug #3 fix: check that previous table was actually built
                Dim prevTabId As String
                On Error Resume Next
                prevTabId = specs.Previous.TableId
                On Error GoTo 0

                If prevTabId <> vbNullString And RangeExists("ENDTABLE_" & prevTabId) Then
                    strtRow = Previous.StartRow
                Else
                    ' Previous table was never built (invalid); treat as de-facto new section
                    strtRow = .Cells(.Rows.Count, STANDARD_START_COL).End(xlUp).Row + 1
                    strtRow = IIf(strtRow = 2, 6, strtRow)
                    strtRow = strtRow + 4
                    .Cells(strtRow, STANDARD_START_COL).Name = "STARTROW_" & tabId
                End If
            End If
            StartRow = strtRow
        Else
            StartRow = .Range("STARTROW_" & tabId).Row
        End If
    End With
End Property

' @description Previous cross-table in the same section
Private Property Get Previous() As ICrossTable
    Set Previous = Create(Specifications.Previous, Wksh(), Translations())
End Property

' @description Starting column for this table.
'              For non-new-section temporal tables, starts after the previous table's last column.
Private Property Get StartColumn() As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim strtCol As Long
    Dim tabId As String
    Dim tablType As Byte

    Set specs = Specifications
    tabId = specs.TableId
    tablType = specs.TableType
    Set sh = Wksh

    With sh
        If Not RangeExists("STARTCOL_" & tabId) Then
            If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And _
                (Not specs.IsNewSection()) Then
                strtCol = .Cells(StartRow + 1, .Columns.Count).End(xlToLeft).Column
            Else
                strtCol = STANDARD_START_COL
            End If
            StartColumn = strtCol
            .Cells(StartRow - 1, strtCol + 1).Name = "STARTCOL_" & tabId
        Else
            StartColumn = .Range("STARTCOL_" & tabId).Column - 1
        End If
    End With
End Property

Private Property Let EndRow(ByVal rw As Long)
    this.endRw = rw
End Property

Private Property Let EndColumn(ByVal col As Long)
    this.endCol = col
End Property

Private Property Get EndRow() As Long
    EndRow = this.endRw
End Property

Private Property Get EndColumn() As Long
    EndColumn = this.endCol
End Property

'@section Range access properties
'===============================================================================

Public Property Get HeaderRange() As Range
    Dim sh As Worksheet
    Set sh = Wksh()
    With sh
        Set HeaderRange = .Range(.Cells(StartRow, StartColumn + 1), .Cells(StartRow, EndColumn))
    End With
End Property

Private Property Get TimeSeriesHeader() As String
    TimeSeriesHeader = this.tsHeader
End Property

Private Property Let TimeSeriesHeader(ByVal headerText As String)
    this.tsHeader = headerText
End Property

Public Property Get RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, _
                                         Optional ByVal onlyCategories As Boolean = False) As Range
    Dim startRw As Long
    Dim endRw As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim tablType As Byte

    Set specs = Specifications
    tablType = specs.TableType
    If tablType = TypeGlobalSummary Then ThrowError ErrorUnexpectedState, "Global summary has no row categories"

    Set sh = Wksh()

    Select Case tablType
    Case TypeBivariate, TypeTimeSeries, TypeSpatial, TypeSpatioTemporal
        startRw = StartRow + 1
    Case TypeUnivariate
        startRw = StartRow
    End Select

    If Not includeHeaders Then startRw = startRw + 1

    endRw = Switch( _
        (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And onlyCategories, Me.EndRow - 2, _
        (specs.HasTotal Or tablType = TypeSpatial) And onlyCategories, Me.EndRow - 1, _
        True, Me.EndRow)

    With sh
        Set RowsCategoriesRange = .Range(.Cells(startRw, STANDARD_START_COL), _
                                          .Cells(endRw, STANDARD_START_COL))
    End With
End Property

Private Function ValueExists(ByVal rng As Range, ByVal searchValue As String) As Boolean
    ValueExists = (Not rng.Find(What:=searchValue, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Function

Public Property Get RowRange(ByVal rowVal As String, _
                              Optional ByVal includeHeaders As Boolean = True) As Range
    Dim rng As Range
    Dim valueRow As Long
    Dim sh As Worksheet
    Dim startCol As Long

    Set rng = RowsCategoriesRange()
    Set sh = Wksh()

    If ValueExists(rng, rowVal) Then
        valueRow = rng.Find(What:=rowVal, LookAt:=xlWhole, MatchCase:=True).Row
        If includeHeaders Then
            startCol = StartColumn
        Else
            startCol = StartColumn + 1
        End If
        With sh
            Set RowRange = .Range(.Cells(valueRow, startCol), .Cells(valueRow, EndColumn))
        End With
    End If
End Property

Public Property Get ColumnRange(ByVal colVal As String, _
                                 Optional ByVal onlyCategories As Boolean = False, _
                                 Optional ByVal includeHeaders As Boolean = False) As Range
    Dim rng As Range
    Dim valueColumn As Long
    Dim startRw As Long
    Dim endRw As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim tablType As Byte

    Set rng = HeaderRange()
    Set sh = Wksh()
    Set specs = Specifications
    tablType = specs.TableType

    startRw = Switch(tablType = TypeBivariate, StartRow + 2, _
                     (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal), StartRow + 2, _
                     tablType = TypeSpatial, StartRow + 2, _
                     True, StartRow + 1)

    endRw = Switch( _
        (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And onlyCategories, Me.EndRow - 2, _
        (specs.HasTotal Or tablType = TypeSpatial) And onlyCategories, Me.EndRow - 1, _
        True, Me.EndRow)

    If includeHeaders Then startRw = startRw - 1

    If ValueExists(rng, colVal) Then
        valueColumn = rng.Find(What:=colVal, LookAt:=xlWhole, MatchCase:=True).Column
        With sh
            Set ColumnRange = .Range(.Cells(startRw, valueColumn), .Cells(endRw, valueColumn))
        End With
    End If
End Property

'@section AddRows — populate row categories
'===============================================================================

Private Sub AddRows()
    Dim startRw As Long
    Dim specs As ITableSpecs
    Dim sh As Worksheet
    Dim vars As ILLVariables
    Dim rowVar As String
    Dim rowsData As BetterArray
    Dim trans As ITranslationObject
    Dim transMissing As String
    Dim transTotal As String
    Dim counter As Long
    Dim spatType As String

    startRw = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Translations()
    Set vars = LLVariables.Create(specs.Dictionary)

    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    Set rowsData = specs.RowCategories()

    With sh
        Select Case specs.TableType()

        Case TypeGlobalSummary
            .Cells(startRw, StartColumn).Value = specs.Value("label")

        Case TypeUnivariate
            If specs.Value("missing") = "yes" Then rowsData.Push transMissing
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(startRw + 1, StartColumn)
            rowVar = specs.Value("row")
            .Cells(startRw, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)
            EndRow = startRw + rowsData.Length

        Case TypeBivariate
            If specs.Value("missing") = "row" Or specs.Value("missing") = "all" Then
                rowsData.Push transMissing
            End If
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(startRw + 2, StartColumn)
            rowVar = specs.Value("row")
            .Cells(startRw + 1, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)
            EndRow = startRw + 1 + rowsData.Length

        Case TypeTimeSeries, TypeSpatioTemporal
            ' Add rows only on new sections (or de-facto new sections for Bug #3)
            If specs.IsNewSection() Or IsDeFactoNewSection() Then
                rowsData.Clear
                rowsData.Push trans.TranslatedValue("MSG_Period"), vbNullString
                For counter = 1 To 13
                    rowsData.Push vbNullString, vbNullString, vbNullString, vbNullString
                Next
                ' Total and Missing rows are ALWAYS present in time series
                rowsData.Push transTotal, transMissing
                rowsData.ToExcelRange .Cells(startRw + 1, StartColumn)
            End If
            EndRow = startRw + NB_ROWS_TIME_SERIES

        Case TypeSpatial
            rowVar = specs.Value("row")
            rowsData.Clear
            spatType = specs.SpatialTableTypes()

            If spatType = "geo" Then
                rowsData.Push vars.Value(colName:="sub section", varName:="adm1_" & rowVar)
            ElseIf spatType = "hf" Then
                rowsData.Push vars.Value(colName:="main label", varName:="hf_" & rowVar)
            End If

            For counter = 1 To CLng(specs.Value("geo max"))
                rowsData.Push vbNullString
            Next

            rowsData.Push transMissing
            rowsData.ToExcelRange .Cells(startRw + 1, StartColumn)
            EndRow = startRw + rowsData.Length

        End Select
    End With
End Sub

'@section Percentage label helper
'===============================================================================

Private Function PercentLabel(ByVal percentVal As String, _
                               Optional ByVal percentType As String = "all") As String
    Select Case percentType
    Case "row"
        PercentLabel = percentVal & " " & ChrW(8596)
    Case "column"
        PercentLabel = percentVal & " " & ChrW(8597)
    Case Else
        PercentLabel = percentVal
    End Select
End Function

'@section BivariateColumns — shared column-building logic for bivariate/temporal/spatial
'===============================================================================

Private Sub BivariateColumns(ByVal startRw As Long, ByVal specs As ITableSpecs, _
                              ByVal trans As ITranslationObject, ByVal sh As Worksheet)
    Dim colLab As String
    Dim colData As BetterArray
    Dim newColData As BetterArray
    Dim labData As BetterArray
    Dim percentageValue As String
    Dim transMissing As String
    Dim transTotal As String
    Dim transPercent As String
    Dim counter As Long

    Set colData = New BetterArray
    Set newColData = New BetterArray
    Set labData = New BetterArray
    colData.LowerBound = 1
    labData.LowerBound = 1
    newColData.LowerBound = 1

    percentageValue = specs.Value("percentage")
    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    transPercent = trans.TranslatedValue("MSG_Percent")

    With sh
        colLab = specs.Value("label")
        If specs.HasMissing Then labData.Push transMissing
        If specs.HasTotal Then labData.Push transTotal
        Set colData = specs.ColumnCategories()

        If colData.Length > 0 Then
            NumberOfColumns = colData.Length

            colData.CONCAT labData.Items
            labData.Clear

            If specs.HasPercentage Then
                counter = 1
                Do While counter <= colData.Length
                    newColData.Push colData.Item(counter), vbNullString
                    labData.Push colLab, PercentLabel(transPercent, percentageValue)
                    counter = counter + 1
                Loop
            Else
                For counter = 1 To colData.Length
                    labData.Push colLab
                Next
                Set newColData = colData.Clone()
            End If

            newColData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            labData.ToExcelRange .Cells(startRw + 1, StartColumn + 1), TransposeValues:=True
        Else
            ' No column data — just a summary function
            labData.Push colLab
            NumberOfColumns = 1
            labData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            ' Placeholder for End(xlToLeft) detection in next table's StartColumn
            .Cells(startRw + 1, StartColumn + 1).Value = " "
        End If

        EndColumn = StartColumn + labData.Length
    End With
End Sub

'@section AddColumns — populate column headers
'===============================================================================

Private Sub AddColumns()
    Dim startRw As Long
    Dim specs As ITableSpecs
    Dim sh As Worksheet
    Dim vars As ILLVariables
    Dim trans As ITranslationObject
    Dim colData As BetterArray
    Dim colVarLab As String
    Dim sumLab As String
    Dim timeVarLab As String
    Dim tabId As String
    Dim tablType As Byte
    Dim spType As String
    Dim percentageValue As String
    Dim transPercent As String

    startRw = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Translations()
    Set vars = LLVariables.Create(specs.Dictionary)
    tabId = specs.TableId()
    tablType = specs.TableType()
    Set colData = New BetterArray
    colData.LowerBound = 1
    spType = specs.SpatialTableTypes()
    percentageValue = specs.Value("percentage")
    transPercent = trans.TranslatedValue("MSG_Percent")

    With sh
        Select Case tablType

        Case TypeGlobalSummary
            If Not RangeExists("COLGS_SET") Then
                .Cells(startRw - 1, StartColumn + 1).Value = trans.TranslatedValue("MSG_AllData")
                .Cells(startRw - 1, StartColumn + 2).Value = trans.TranslatedValue("MSG_FilteredData")
                .Cells(startRw - 1, StartColumn + 1).Name = "COLGS_SET"
            End If
            EndColumn = StartColumn + 2
            NumberOfColumns = 2

        Case TypeUnivariate
            Dim colLab As String
            colLab = specs.Value("label")
            colData.Push colLab
            If percentageValue = "yes" Then colData.Push transPercent
            colData.ToExcelRange .Cells(startRw, StartColumn + 1), TransposeValues:=True
            EndColumn = StartColumn + colData.Length
            NumberOfColumns = 1

        Case TypeBivariate
            BivariateColumns startRw, specs, trans, sh
            .Range("STARTCOL_" & tabId).Value = vars.Value( _
                colName:="main label", varName:=specs.Value("column"))

        Case TypeTimeSeries, TypeSpatioTemporal
            BivariateColumns startRw, specs, trans, sh

            If tablType = TypeTimeSeries Then
                colVarLab = vars.Value(colName:="main label", varName:=specs.Value("column"))
            ElseIf spType = "geo" Then
                colVarLab = vars.Value(colName:="sub section", varName:="adm1_" & specs.Value("column"))
            ElseIf spType = "hf" Then
                colVarLab = vars.Value(colName:="main label", varName:="hf_" & specs.Value("column"))
            End If

            timeVarLab = vars.Value(colName:="main label", varName:=specs.Value("time"))
            sumLab = specs.Value("label")

            If specs.Value("column") = vbNullString Then
                TimeSeriesHeader = sumLab & " " & ChrW(9472) & " " & timeVarLab
            Else
                TimeSeriesHeader = sumLab & " " & ChrW(9472) & " " & timeVarLab & _
                                   " " & ChrW(9472) & " " & colVarLab
            End If

            .Range("STARTCOL_" & tabId).Value = TimeSeriesHeader

        Case TypeSpatial
            BivariateColumns startRw, specs, trans, sh

        End Select
    End With
End Sub

'@section AddHeader — create section headers, titles, date controls
'===============================================================================

Private Sub AddHeader()
    Dim startRw As Long
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim counter As Long
    Dim trans As ITranslationObject
    Dim tabId As String
    Dim rng As Range
    Dim cellRng As Range
    Dim nGeo As Long
    Dim tablType As Byte
    Dim nbShift As Long
    Dim timeUnitName As String
    Dim sptGeoRngTag As String

    Set sh = Wksh()
    Set specs = Specifications
    Set trans = Translations()

    Set allRangeNames = New BetterArray

    startRw = StartRow
    tabId = specs.TableId
    tablType = specs.TableType
    sptGeoRngTag = "INPUTSPTGEO_"
    If specs.Value("spatial type") = "hf" Then sptGeoRngTag = "INPUTSPTHF_"

    With sh
        Select Case tablType

        Case TypeUnivariate, TypeBivariate
            .Cells(startRw - 2, STANDARD_START_COL).Value = specs.Value("title")
            .Cells(startRw - 2, STANDARD_START_COL).Name = "TITLE_" & tabId
            allRangeNames.Push "TITLE_" & tabId

            If specs.IsNewSection() And (Not RangeExists("SECTION_" & tabId)) Then
                For counter = 1 To 3
                    .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Range("TITLE_" & tabId).Cells(-1, 1) = specs.Value("section")
                .Range("TITLE_" & tabId).Cells(-1, 1).Name = "SECTION_" & tabId
                allRangeNames.Push "SECTION_" & tabId
            End If

        Case TypeGlobalSummary
            If Not RangeExists("ROWGS_SET") Then
                For counter = 1 To 3
                    .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Cells(startRw - 2, STANDARD_START_COL).Value = trans.TranslatedValue("MSG_GlobalSummary")
                .Cells(startRw - 2, STANDARD_START_COL).Name = "ROWGS_SET"
                allRangeNames.Push "ROWGS_SET"
            End If

        Case TypeTimeSeries, TypeSpatioTemporal
            ' Bug #3 fix: treat as new section if previous table was never built
            If (specs.IsNewSection() Or IsDeFactoNewSection()) And _
               (Not RangeExists("SECTION_" & tabId)) Then

                If tablType = TypeSpatioTemporal Then
                    nGeo = 5
                    On Error Resume Next
                    nGeo = CLng(specs.Value("n geo"))
                    On Error GoTo 0
                    nbShift = nGeo + 14
                Else
                    nbShift = 7
                End If

                For counter = 1 To nbShift
                    .Rows(startRw - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Cells(startRw - 2, StartColumn).Value = specs.Value("section")
                .Cells(startRw - 2, StartColumn).Name = "SECTION_" & tabId
                allRangeNames.Push "SECTION_" & tabId

                Set rng = .Range("SECTION_" & tabId)

                ' Spatio-temporal geo-level selection
                If tablType = TypeSpatioTemporal Then
                    Set rng = rng.Offset(4)
                    rng.Value = trans.TranslatedValue("MSG_InputGeoLevels")
                    rng.Name = "INPUTSPTGEO_LABEL_" & tabId
                    allRangeNames.Push "INPUTSPTGEO_LABEL_" & tabId
                    rng.Cells(1, 2).Value = nGeo
                    rng.Cells(1, 2).Name = "INPUTSPTGEO_NGEO_" & tabId

                    Set cellRng = rng.Offset(2)
                    For counter = 1 To nGeo
                        cellRng.Cells(counter, 1).Name = sptGeoRngTag & counter & "_" & tabId
                        cellRng.Cells(counter, 2).Value = 1
                    Next

                    Set rng = .Range(cellRng.Cells(1, 1), cellRng.Cells(nGeo, 1))
                    rng.Name = "INPUTSPTGEO_USER_" & tabId
                    allRangeNames.Push "INPUTSPTGEO_USER_" & tabId
                    rng.Locked = False
                    rng.Offset(, 1).Name = "INPUTPREVADM_USER_" & tabId
                    Set rng = cellRng.Offset(nGeo + 3)
                End If

                ' Date controls (start date, time unit, end date)
                rng.Cells(6, 1).Name = "INFO_ANA_PERIOD_" & tabId
                rng.Cells(3, 3).Value = trans.TranslatedValue("MSG_StartDate")
                rng.Cells(3, 3).Name = "START_DATE_LABEL_" & tabId
                rng.Cells(4, 3).Locked = False
                rng.Cells(4, 3).Name = "USER_START_DATE_" & tabId
                rng.Cells(5, 3).Name = "INFO_START_DATE_" & tabId
                rng.Cells(6, 3).Name = "START_DATE_" & tabId

                rng.Cells(3, 5).Value = trans.TranslatedValue("MSG_TimeUnit")
                rng.Cells(3, 5).Name = "TIME_UNIT_LABEL_" & tabId
                rng.Cells(4, 5).Name = "TIME_UNIT_" & tabId
                rng.Cells(4, 5).Locked = False

                rng.Cells(3, 7).Value = trans.TranslatedValue("MSG_EndDate")
                rng.Cells(3, 7).Name = "END_DATE_LABEL_" & tabId
                rng.Cells(4, 7).Name = "USER_END_DATE_" & tabId
                rng.Cells(4, 7).Locked = False
                rng.Cells(5, 7).Name = "INFO_END_DATE_" & tabId
                rng.Cells(6, 7).Name = "END_DATE_" & tabId

                rng.Cells(4, 9).Name = "VALIDATION_MIN_DATE_" & tabId
                rng.Cells(4, 11).Name = "VALIDATION_MAX_DATE_" & tabId
                rng.Cells(4, 13).Name = "MIN_MIN_DATE_" & tabId
                rng.Cells(4, 15).Name = "MAX_MAX_DATE_" & tabId

                allRangeNames.Push "INFO_ANA_PERIOD_" & tabId, _
                    "START_DATE_LABEL_" & tabId, "USER_START_DATE_" & tabId, _
                    "INFO_START_DATE_" & tabId, "START_DATE_" & tabId, _
                    "END_DATE_LABEL_" & tabId, "USER_END_DATE_" & tabId, _
                    "INFO_END_DATE_" & tabId, "END_DATE_" & tabId

                timeUnitName = "TIME_UNIT_LIST"
                If tablType = TypeSpatioTemporal Then timeUnitName = "SP" & timeUnitName

                Set cellRng = .Cells(1, 26)
                If Not RangeExists(timeUnitName) Then
                    cellRng.Cells(1, 1).Value = trans.TranslatedValue("MSG_Day")
                    cellRng.Cells(2, 1).Value = trans.TranslatedValue("MSG_Week")
                    cellRng.Cells(3, 1).Value = trans.TranslatedValue("MSG_Month")
                    cellRng.Cells(4, 1).Value = trans.TranslatedValue("MSG_Quarter")
                    cellRng.Cells(5, 1).Value = trans.TranslatedValue("MSG_Year")
                    .Range(.Cells(1, cellRng.Column), .Cells(5, cellRng.Column)).Name = timeUnitName
                    .Range(timeUnitName).Font.Color = vbWhite
                    .Range(timeUnitName).Interior.Color = vbWhite
                    .Range(timeUnitName).Locked = True
                End If

                .Range("TIME_UNIT_" & tabId).Value = cellRng.Cells(2, 1).Value
            End If

        Case TypeSpatial
            If Not RangeExists("ADMIN_UNIT_LIST") Then
                Set cellRng = .Cells(1, 26)
                cellRng.Cells(1, 1).Formula = "= RNG_ADM1NAME"
                cellRng.Cells(2, 1).Formula = "= RNG_ADM2NAME"
                cellRng.Cells(3, 1).Formula = "= RNG_ADM3NAME"
                cellRng.Cells(4, 1).Formula = "= RNG_ADM4NAME"
                .Range(.Cells(1, cellRng.Column), .Cells(4, cellRng.Column)).Name = "ADM_UNIT_LIST"
                .Range("ADM_UNIT_LIST").Font.Color = vbWhite
                .Range("ADM_UNIT_LIST").Interior.Color = vbWhite
                .Range("ADM_UNIT_LIST").Locked = True
            End If

            If Not RangeExists("POPULATION_FACTOR_LIST") Then
                Set cellRng = .Cells(1, 27)
                cellRng.Cells(1, 1).Value = trans.TranslatedValue("MSG_NoDevide")
                cellRng.Cells(2, 1).Value = trans.TranslatedValue("MSG_Devide")
                .Range(.Cells(1, cellRng.Column), .Cells(2, cellRng.Column)).Name = "POPULATION_FACTOR_LIST"
                .Range("POPULATION_FACTOR_LIST").Font.Color = vbWhite
                .Range("POPULATION_FACTOR_LIST").Interior.Color = vbWhite
                .Range("POPULATION_FACTOR_LIST").Locked = True
            End If

            If (Not RangeExists("ADM_DROPDOWN_" & tabId)) And _
               (Not RangeExists("HF_DROPDOWN_" & tabId)) Then

                If specs.SpatialTableTypes = "geo" Then
                    .Cells(startRw - 2, STANDARD_START_COL).Name = "ADM_DROPDOWN_" & tabId
                    .Cells(startRw - 2, STANDARD_START_COL).Value = trans.TranslatedValue("MSG_SelectAdmin")

                    Set rng = .Range("ADM_DROPDOWN_" & tabId)
                    rng.Cells(1, 2).Value = "adm1"
                    rng.Cells(1, 2).Name = "PREVIOUS_ADM_" & tabId
                    rng.Cells(1, 3).Value = trans.TranslatedValue("MSG_SelectPOPFACT")
                    rng.Cells(1, 3).Name = "DEVIDEPOP_" & tabId
                    rng.Cells(0, 4).Value = trans.TranslatedValue("MSG_MultiplyBy")
                    rng.Cells(0, 4).Name = "POPFACTLABEL_" & tabId
                    rng.Cells(1, 4).Value = 100
                    rng.Cells(1, 4).Name = "POPFACT_" & tabId
                    rng.Cells(1, 5).Name = "POPPREVFACT_" & tabId
                    rng.Cells(1, 5).Value = "0"

                    allRangeNames.Push "DEVIDEPOP_" & tabId, "POPFACT_" & tabId, _
                                        "ADM_DROPDOWN_" & tabId

                ElseIf specs.SpatialTableTypes = "hf" Then
                    .Cells(startRw - 2, STANDARD_START_COL).Name = "HF_DROPDOWN_" & tabId
                    Set rng = .Range("HF_DROPDOWN_" & tabId)
                    rng.Cells(1, 1).Value = trans.TranslatedValue("MSG_HF")
                    allRangeNames.Push "HF_DROPDOWN_" & tabId
                End If

                If specs.IsNewSection() Then
                    For counter = 1 To 3
                        .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Next
                    rng.Cells(-3, 1).Value = specs.Value("section")
                    rng.Cells(-3, 1).Name = "SECTION_" & tabId
                    allRangeNames.Push "SECTION_" & tabId
                End If

                .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                rng.Cells(-1, 1).Value = specs.Value("title")
                rng.Cells(-1, 1).Name = "TITLE_" & tabId
                allRangeNames.Push "TITLE_" & tabId
            End If

        End Select
    End With
End Sub

'@section NameRanges — create all named ranges for the table
'===============================================================================

Private Sub NameRanges()
    Dim specs As ITableSpecs
    Dim sh As Worksheet
    Dim nbCols As Long
    Dim catRng As Range
    Dim rng As Range
    Dim nbColCounter As Long
    Dim counter As Long
    Dim trans As ITranslationObject
    Dim labCol As Long
    Dim cData As BetterArray
    Dim tData As BetterArray
    Dim rw As Long
    Dim col As Long
    Dim addBegin As String
    Dim addEnd As String
    Dim tabId As String
    Dim cellRng As Range
    Dim tablType As Byte
    Dim isNewSect As Boolean

    Set sh = Wksh()
    Set trans = Translations()
    Set specs = Specifications
    tabId = specs.TableId
    tablType = specs.TableType
    ' Bug #3 fix: treat as new section if previous table was never built
    isNewSect = specs.IsNewSection() Or IsDeFactoNewSection()

    If tablType = TypeGlobalSummary Then Exit Sub

    With sh
        If Not RangeExists("ENDTABLE_" & tabId) Then
            .Cells(EndRow + 1, StartColumn).Name = "ENDTABLE_" & tabId
        End If

        nbCols = NumberOfColumns
        If nbCols = 0 Then ThrowError ErrorUnexpectedState, _
                                      "Table " & tabId & " has 0 columns"

        ' Row categories — only for new sections or non-temporal tables
        If isNewSect Or (tablType <> TypeTimeSeries And tablType <> TypeSpatioTemporal) Then
            Set catRng = RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False)
            catRng.Name = "ROW_CATEGORIES_" & tabId
            catRng.Cells(0, 1).Name = "LABEL_ROW_CATEGORIES_" & tabId
            allRangeNames.Push "LABEL_ROW_CATEGORIES_" & tabId
            allRangeNames.Push "ROW_CATEGORIES_" & tabId

            If tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal Then
                Set rng = catRng.Offset(, -1)
                rng.Name = "END_TIME_PERIOD_" & tabId
                Set rng = rng.Offset(, -1)
                rng.Cells(1, 1).Name = "FIRST_VALUE_START_TIME_" & tabId
                Set rng = .Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
                rng.Name = "START_TIME_PERIOD_" & tabId
            End If
        End If

        ' Non-new-section temporal: derive catRng from ENDTABLE position
        ' Bug #3 fix: only enter this block if section infrastructure actually exists
        If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And Not isNewSect Then
            Dim prevTabId As String
            prevTabId = specs.Previous.TableId
            If Not RangeExists("ENDTABLE_" & prevTabId) Then
                ThrowError ErrorUnexpectedState, _
                           "Cannot append time series: no previous table found for " & prevTabId
            End If
            Set cellRng = .Range("ENDTABLE_" & specs.TableId)
            Set catRng = .Range(cellRng.Offset(1 - NB_ROWS_TIME_SERIES), cellRng.Offset(-3))
        End If

        ' Name all value columns
        counter = StartColumn
        nbColCounter = 1
        Set rng = catRng.Offset(, 1)

        Do While (counter <= EndColumn) And (nbColCounter <= nbCols)
            rng.Name = "VALUES_COL_" & nbColCounter & "_" & tabId
            allRangeNames.Push "VALUES_COL_" & nbColCounter & "_" & tabId

            Select Case tablType
            Case TypeBivariate, TypeTimeSeries, TypeSpatial, TypeSpatioTemporal
                labCol = -1
            Case TypeUnivariate
                labCol = 0
            End Select

            rng.Cells(labCol, 1).Name = "LABEL_COL_" & nbColCounter & "_" & tabId
            allRangeNames.Push "LABEL_COL_" & nbColCounter & "_" & tabId

            Set rng = rng.Offset(, 1)

            If specs.HasPercentage Then
                rng.Name = "PERC_COL_" & nbColCounter & "_" & tabId
                allRangeNames.Push "PERC_COL_" & nbColCounter & "_" & tabId

                If tablType = TypeUnivariate Then
                    rng.Cells(labCol, 1).Name = "PERC_LABEL_COL_" & tabId
                    allRangeNames.Push "PERC_LABEL_COL_" & tabId
                End If

                Set rng = rng.Offset(, 1)
                counter = counter + 1
            End If

            counter = counter + 1
            nbColCounter = nbColCounter + 1
        Loop

        ' Name Missing/Total/Percent rows and columns
        Set cData = New BetterArray
        Set tData = New BetterArray
        cData.Push "MISSING", "TOTAL", "PERC"
        tData.Push "MSG_NA", "MSG_Total", "MSG_Percent"

        For counter = cData.LowerBound To cData.UpperBound
            ' Row ranges
            Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=True)
            If Not rng Is Nothing Then
                rng.Name = cData.Item(counter) & "_ROW_" & tabId
                allRangeNames.Push cData.Item(counter) & "_ROW_" & tabId
                Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=False)
                rng.Name = cData.Item(counter) & "_ROW_VALUES_" & tabId
                allRangeNames.Push cData.Item(counter) & "_ROW_VALUES_" & tabId
            End If

            ' Column ranges (not for univariate)
            If tablType <> TypeUnivariate Then
                Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), _
                                      includeHeaders:=True, onlyCategories:=True)
                If Not rng Is Nothing Then
                    rng.Name = cData.Item(counter) & "_COL_" & tabId
                    allRangeNames.Push cData.Item(counter) & "_COL_" & tabId
                    rng.Cells(0, 1).Name = cData.Item(counter) & "_LABEL_COL_" & tabId
                    Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), _
                                          includeHeaders:=False, onlyCategories:=True)
                    rng.Name = cData.Item(counter) & "_COL_VALUES_" & tabId
                    allRangeNames.Push cData.Item(counter) & "_COL_VALUES_" & tabId

                    If specs.HasPercentage Then
                        Set rng = rng.Offset(, 1)
                        rng.Name = cData.Item(counter) & "_PERC_VALUES_" & tabId
                        allRangeNames.Push cData.Item(counter) & "_PERC_VALUES_" & tabId
                        .Range(cData.Item(counter) & "_COL_" & tabId).Offset(, 1).Name = _
                            cData.Item(counter) & "_PERC_" & tabId
                    End If
                End If
            End If
        Next

        ' Intersection cells (total/missing combinations)
        If RangeExists("TOTAL_COL_" & tabId) Then
            If RangeExists("TOTAL_ROW_" & tabId) Then
                rw = .Range("TOTAL_ROW_" & tabId).Row
                col = .Range("TOTAL_COL_" & tabId).Column
                .Cells(rw, col).Name = "TOTAL_TOTAL_" & tabId
                allRangeNames.Push "TOTAL_TOTAL_" & tabId
            End If
            If RangeExists("MISSING_ROW_" & tabId) Then
                rw = .Range("MISSING_ROW_" & tabId).Row
                col = .Range("TOTAL_COL_" & tabId).Column
                .Cells(rw, col).Name = "MISSING_TOTAL_" & tabId
                allRangeNames.Push "MISSING_TOTAL_" & tabId
                If RangeExists("MISSING_COL_" & tabId) Then
                    col = .Range("MISSING_COL_" & tabId).Column
                    .Cells(rw, col).Name = "MISSING_MISSING_" & tabId
                    allRangeNames.Push "MISSING_MISSING_" & tabId
                End If
            End If
        End If

        If RangeExists("TOTAL_ROW_" & tabId) Then
            If RangeExists("MISSING_COL_" & tabId) Then
                col = .Range("MISSING_COL_" & tabId).Column
                rw = .Range("TOTAL_ROW_" & tabId).Row
                .Cells(rw, col).Name = "TOTAL_MISSING_" & tabId
                allRangeNames.Push "TOTAL_MISSING_" & tabId
            End If
        End If

        ' Interior values (all data cells)
        addBegin = Split(.Range("VALUES_COL_1_" & tabId).Address, ":")(0)
        addEnd = IIf(specs.HasPercentage, "PERC_COL_", "VALUES_COL_")
        addEnd = Split(.Range(addEnd & nbCols & "_" & tabId).Address, ":")(1)
        .Range(addBegin & ":" & addEnd).Name = "INTERIOR_VALUES_" & tabId
        allRangeNames.Push "INTERIOR_VALUES_" & tabId

        ' Spatial outer values for population computation
        If tablType = TypeSpatial Then
            Set rng = .Range("INTERIOR_VALUES_" & tabId)
            Set rng = .Range(rng.Cells(1, -1), rng.Cells(rng.Rows.Count + 1, rng.Columns.Count + 4))
            rng.Name = "OUTER_VALUES_" & tabId
            allRangeNames.Push "OUTER_VALUES_" & tabId
        End If

        ' Column categories header
        Set rng = HeaderRange()
        rng.Name = "COLUMN_CATEGORIES_" & tabId
        allRangeNames.Push "COLUMN_CATEGORIES_" & tabId
    End With
End Sub

'@section Bug #3 helper — detect de-facto new section
'===============================================================================

' @description Check if this is a de-facto new section: the specs say it's not a new
'              section (same section string as previous row), but the previous table
'              was never built (invalid or skipped), so section infrastructure is missing.
'              In that case, this table should act as the section anchor.
Private Function IsDeFactoNewSection() As Boolean
    Dim specs As ITableSpecs
    Set specs = Specifications

    ' Only relevant for non-new-section temporal tables
    If specs.IsNewSection() Then
        IsDeFactoNewSection = False
        Exit Function
    End If

    If specs.TableType <> TypeTimeSeries And specs.TableType <> TypeSpatioTemporal Then
        IsDeFactoNewSection = False
        Exit Function
    End If

    ' Check if the previous table's end marker exists
    Dim prevTabId As String
    On Error Resume Next
    prevTabId = specs.Previous.TableId
    On Error GoTo 0

    If prevTabId = vbNullString Then
        IsDeFactoNewSection = True
        Exit Function
    End If

    ' If the previous table's ENDTABLE range doesn't exist, it was never built
    IsDeFactoNewSection = Not RangeExists("ENDTABLE_" & prevTabId)
End Function

'@section Format — apply design formatting + Bug #1 fix
'===============================================================================

Private Sub Format(ByVal desFormat As ILLFormat)
    Dim sh As Worksheet
    Dim specs As ITableSpecs
    Dim rng As Range
    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String
    Dim tabId As String
    Dim designFormat As ILLFormat
    Dim wholeTableStart As Long
    Dim hasPercent As Boolean
    Dim newSection As Boolean
    Dim tablType As Byte

    Set rngTable = New BetterArray
    Set specs = Specifications
    Set sh = Wksh()
    Set designFormat = desFormat

    tabId = specs.TableId
    tablType = specs.TableType
    hasPercent = specs.HasPercentage
    newSection = specs.IsNewSection() Or IsDeFactoNewSection()

    With sh
        ' Sections ==============================================================
        If newSection Then
            Set rng = .Range("SECTION_" & tabId)
            designFormat.ApplyFormat rng, AnalysisSection
        ElseIf RangeExists("ROWGS_SET") And (Not RangeExists("ROWGS_FORMATSET")) Then
            Set rng = .Range("ROWGS_SET")
            designFormat.ApplyFormat rng, scope:=AnalysisSection
            rng.Cells(1, 8).Name = "ROWGS_FORMATSET"
        End If

        ' Title (GlobalSummary and TimeSeries don't have titles)
        If (tablType <> TypeGlobalSummary) And (tablType <> TypeTimeSeries) And _
           (tablType <> TypeSpatioTemporal) Then
            Set rng = .Range("TITLE_" & tabId)
            designFormat.ApplyFormat rng, AnalysisTableTitle
        End If

        ' Spatial dropdown formatting
        If tablType = TypeSpatial Then
            If specs.SpatialTableTypes = "geo" Then
                Set rng = .Range("ADM_DROPDOWN_" & tabId)
                designFormat.ApplyFormat rng, scope:=LinelistSelectionDropdown
                Set rng = .Range("DEVIDEPOP_" & tabId)
                designFormat.ApplyFormat rng, scope:=LinelistSelectionDropdown
            ElseIf specs.SpatialTableTypes = "hf" Then
                Set rng = .Range("HF_DROPDOWN_" & tabId)
                designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            End If
            Set rng = .Range("ROW_CATEGORIES_" & tabId).Offset(, -1)
            designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            Set rng = rng.Offset(, -1)
            designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
        End If

        ' Headers ===============================================================
        Select Case tablType
        Case TypeGlobalSummary
            If Not RangeExists("COLGS_FORMATSET") Then
                Set rng = .Range("COLGS_SET")
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 2))
                designFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
                rng.Cells(1, 3).Name = "COLGS_FORMATSET"
            End If
        Case TypeUnivariate
            Set rng = .Range("COLUMN_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
            Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
        Case TypeBivariate, TypeSpatial
            Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
            Set rng = .Range("STARTCOL_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelColumn
        Case TypeTimeSeries, TypeSpatioTemporal
            If newSection Then
                Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
                designFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
            End If
            Set rng = .Range("STARTCOL_" & tabId)
            Set rng = .Range(.Cells(rng.Row, rng.Column), .Cells(rng.Row, EndColumn()))
            designFormat.ApplyFormat rng, scope:=AnalysisTimeSeriesHeader
        End Select

        ' Categories rows and columns ===========================================
        Select Case tablType
        Case TypeGlobalSummary
            Set rng = .Range("STARTROW_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisGlobalSummaryRow
        Case TypeUnivariate
            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesUnivariateRow
        Case TypeBivariate, TypeTimeSeries, TypeSpatial, TypeSpatioTemporal
            If ((tablType <> TypeTimeSeries) And (tablType <> TypeSpatioTemporal)) Or newSection Then
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                designFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow
                If tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal Then
                    Set rng = rng.Offset(, -1)
                    Set rng = .Range(rng.Cells(1, 0), rng.Cells(rng.Rows.Count + 1, 1))
                    designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
                    rngTable.Clear
                    rngTable.Push "START_DATE_", "END_DATE_", "VALIDATION_MIN_DATE_", _
                                  "VALIDATION_MAX_DATE_", "MIN_MIN_DATE_", "MAX_MAX_DATE_"
                    For counter = rngTable.LowerBound To rngTable.UpperBound
                        rngName = rngTable.Item(counter) & tabId
                        Set rng = .Range(rngName)
                        designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
                        designFormat.ApplyFormat rng, scope:=AnalysisDateCell
                    Next
                End If
            End If

            If RangeExists("INPUTSPTGEO_USER_" & tabId) Then
                Set rng = .Range("INPUTSPTGEO_USER_" & tabId)
                designFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow
            End If

            ' Column label merging
            Set rng = .Range("LABEL_COL_1_" & tabId)
            If Not RangeExists("MERGED_" & tabId) Then
                Do While rng.Column <= EndColumn()
                    If hasPercent Then
                        Set rng = rng.Offset(, 2)
                        .Range(rng.Cells(1, -1), rng.Cells(1, 0)).Merge
                    Else
                        Set rng = rng.Offset(, 1)
                    End If
                Loop
                rng.Cells(1, 3).Name = "MERGED_" & tabId
            End If

            Set rng = .Range("COLUMN_CATEGORIES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateColumn
            Set rng = rng.Offset(1)
            designFormat.ApplyFormat rng, scope:=AnalysisFirstRow
        End Select

        ' Interior values =======================================================
        If tablType <> TypeGlobalSummary Then
            Set rng = .Range("INTERIOR_VALUES_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisInteriorTable
        End If

        ' Unique cells — Total/Missing rows and columns ==========================
        rngTable.Clear
        rngTable.Push "TOTAL_ROW_VALUES_", "TOTAL_COL_VALUES_", _
                      "MISSING_ROW_VALUES_", "MISSING_COL_VALUES_", _
                      "TOTAL_PERC_VALUES_", "MISSING_PERC_VALUES_"

        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If (InStr(1, rngName, "ROW") > 0) And RangeExists(rngName) Then
                Set rng = .Range(rngName)
                If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And _
                   (Not newSection) Then
                    designFormat.ApplyFormat rng, scope:=AnalysisMissingRowWithoutHeader
                Else
                    designFormat.ApplyFormat rng, scope:=AnalysisMissingRow
                End If
            ElseIf RangeExists(rngName) Then
                Set rng = .Range(rngName)
                designFormat.ApplyFormat rng, scope:=AnalysisMissingColumn
            End If
        Next

        ' Single cells
        rngTable.Clear
        rngTable.Push "TOTAL_MISSING_", "TOTAL_TOTAL_", _
                      "USER_START_DATE_", "USER_END_DATE_", "TIME_UNIT_", _
                      "MISSING_TOTAL_", "MISSING_MISSING_", "START_DATE_LABEL_", _
                      "TIME_UNIT_LABEL_", "END_DATE_LABEL_", "POPFACT_", "POPFACTLABEL_", _
                      "INPUTSPTGEO_LABEL_"

        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                If (InStr(1, rngName, "LABEL") > 0) Or (InStr(1, rngName, "POP") > 0) Then
                    designFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
                ElseIf (InStr(1, rngName, "USER") > 0) Or (rngName = "TIME_UNIT_" & tabId) Then
                    designFormat.ApplyFormat rng, scope:=AnalysisOneCellCenter
                Else
                    designFormat.ApplyFormat rng, scope:=AnalysisOneCell
                End If
            End If
        Next

        ' Info cells
        rngTable.Clear
        rngTable.Push "INFO_START_DATE_", "INFO_END_DATE_", "INFO_ANA_PERIOD_"
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                designFormat.ApplyFormat rng, scope:=AnalysisInfoCell
                If InStr(1, rngName, "INFO_ANA_") > 0 Then
                    designFormat.ApplyFormat rng, scope:=AnalysisInfoCellLeft
                End If
            End If
        Next

        ' Hidden cells
        rngTable.Clear
        rngTable.Push "PREVIOUS_ADM_", "POPPREVFACT_", "POPFACT_", "POPFACTLABEL_", _
                      "INPUTSPTGEO_NGEO_", "INPUTPREVADM_USER_"
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                designFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            End If
        Next

        ' Percentage columns ====================================================
        If hasPercent Then
            Set rng = .Range("PERC_COL_1_" & tabId)
            If (specs.TableType <> TypeSpatial) Or specs.HasTotal Then
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1))
            End If
            Do While rng.Column <= EndColumn()
                designFormat.ApplyFormat rng, scope:=AnalysisPercent
                Set rng = rng.Offset(, 2)
            Loop
            rngTable.Clear
            rngTable.Push "MISSING_MISSING_", "TOTAL_TOTAL_", "MISSING_TOTAL_"
            For counter = rngTable.LowerBound To rngTable.UpperBound
                rngName = rngTable.Items(counter) & tabId
                If RangeExists(rngName) Then
                    Set rng = .Range(rngName).Offset(, 1)
                    designFormat.ApplyFormat rng, scope:=AnalysisOneCellCenter
                End If
            Next
        End If

        ' Whole table ===========================================================
        wholeTableStart = StartColumn
        If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And _
           (Not newSection) Then
            wholeTableStart = StartColumn + 1
        End If

        If (tablType <> TypeUnivariate) And (tablType <> TypeGlobalSummary) Then
            Set rng = .Range(.Cells(StartRow, wholeTableStart), .Cells(EndRow, EndColumn))
            designFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        ElseIf tablType = TypeUnivariate Then
            Set rng = .Range(.Cells(StartRow + 1, wholeTableStart), .Cells(EndRow, EndColumn))
            designFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        End If

        ' End table gap for graphs
        If RangeExists("ENDTABLE_" & tabId) And _
           (tablType <> TypeTimeSeries) And _
           (tablType <> TypeSpatioTemporal) Then
            Set rng = .Range("ENDTABLE_" & tabId)
            designFormat.ApplyFormat rng, scope:=AnalysisEndTable
        End If

        ' Bug #1 fix: Hide Total COLUMN when not explicitly requested by user
        ' For time series/spatio-temporal tables, HasTotal can be True because
        ' percentage="row"/"column" needs the total column for computation.
        ' When TotalRequested=False, hide the column to keep it invisible.
        If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And _
           specs.HasTotal And (Not specs.TotalRequested) Then
            If RangeExists("TOTAL_COL_" & tabId) Then
                .Range("TOTAL_COL_" & tabId).EntireColumn.Hidden = True
                ' Also hide the adjacent percentage column for Total if it exists
                If hasPercent And RangeExists("TOTAL_PERC_VALUES_" & tabId) Then
                    .Range("TOTAL_PERC_VALUES_" & tabId).EntireColumn.Hidden = True
                End If
            End If
        End If

    End With
End Sub

'@section Build — orchestrate table creation
'===============================================================================

Private Sub Build()
    AddHeader
    AddRows
    AddColumns
    NameRanges
End Sub

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface implementation
'===============================================================================

Private Sub ICrossTable_AddRows()
    AddRows
End Sub

Private Sub ICrossTable_AddColumns()
    AddColumns
End Sub

Private Sub ICrossTable_AddHeader()
    AddHeader
End Sub

Private Sub ICrossTable_NameRanges()
    NameRanges
End Sub

Private Sub ICrossTable_Format(ByVal designFormat As ILLFormat)
    Format designFormat
End Sub

Private Sub ICrossTable_Build()
    Build
End Sub

Private Property Get ICrossTable_HeaderRange() As Range
    Set ICrossTable_HeaderRange = HeaderRange
End Property

Private Property Get ICrossTable_TimeSeriesHeader() As String
    ICrossTable_TimeSeriesHeader = TimeSeriesHeader()
End Property

Private Property Get ICrossTable_RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, _
                                                      Optional ByVal onlyCategories As Boolean = False) As Range
    Set ICrossTable_RowsCategoriesRange = RowsCategoriesRange(includeHeaders, onlyCategories)
End Property

Private Property Get ICrossTable_ColumnRange(ByVal colVal As String, _
                                              Optional ByVal onlyCategories As Boolean = False, _
                                              Optional ByVal includeHeaders As Boolean = False) As Range
    Set ICrossTable_ColumnRange = ColumnRange(colVal, onlyCategories, includeHeaders)
End Property

Private Property Get ICrossTable_RowRange(ByVal rowVal As String, _
                                           Optional ByVal includeHeaders As Boolean = True) As Range
    Set ICrossTable_RowRange = RowRange(rowVal, includeHeaders)
End Property

Private Property Get ICrossTable_StartRow() As Long
    ICrossTable_StartRow = StartRow
End Property

Private Property Get ICrossTable_Previous() As ICrossTable
    Set ICrossTable_Previous = Previous
End Property

Private Property Get ICrossTable_Specifications() As ITableSpecs
    Set ICrossTable_Specifications = Specifications
End Property

Private Property Get ICrossTable_Wksh() As Worksheet
    Set ICrossTable_Wksh = Wksh()
End Property

Private Property Get ICrossTable_NumberOfColumns() As Long
    ICrossTable_NumberOfColumns = NumberOfColumns()
End Property

Private Property Get ICrossTable_EndColumn() As Long
    ICrossTable_EndColumn = EndColumn()
End Property

Private Property Get ICrossTable_EndRow() As Long
    ICrossTable_EndRow = EndRow()
End Property

Private Property Get ICrossTable_NamedRangesList() As BetterArray
    Set ICrossTable_NamedRangesList = NamedRangesList()
End Property
