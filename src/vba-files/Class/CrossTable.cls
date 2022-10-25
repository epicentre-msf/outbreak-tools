VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("")
'@IgnoreModule

Option Explicit

'Exposed methods
Implements ICrossTable

Private Type TCrossTable
    specs As ITablesSpecs
    outputsh As Worksheet
    trans As ITranslation
    endRw As Long
    endCol As Long
    nbCols As Long
End Type

Private this As TCrossTable
Private Const STANDARDSTARTCOL As Long = 3
Private Const CLASSNAME As String = "CrossTable"

Public Function Create(ByVal specs As ITablesSpecs, ByVal outputWksh As Worksheet, ByVal trans As ITranslation) As ICrossTable
    'Check if values are not defined
    If specs Is Nothing Then ThrowErrorUnableInstanciate " Table specifications not defined"
    If outputWksh Is Nothing Then ThrowErrorUnableInstanciate " Output worksheet not defined"
    If trans Is Nothing Then ThrowErrorUnableInstanciate " Translations not defined"

    With New CrossTable
        Set .Specifications = specs
        Set .Wksh = outputWksh
        Set .Trads = trans
        Set Create = .Self
    End With

End Function

Public Property Get Self() As ICrossTable
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.outputsh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.outputsh = sh
End Property

Public Property Get Specifications() As ITablesSpecs
    Set Specifications = this.specs
End Property

Public Property Set Specifications(ByVal specs As ITablesSpecs)
    Set this.specs = specs
End Property

Public Property Get Trads() As ITranslation
    Set Trads = this.trans
End Property

Public Property Set Trads(ByVal trans As ITranslation)
    Set this.trans = trans
End Property

Private Property Get NumberOfColumns() As Long
    NumberOfColumns = this.nbCols
End Property

Private Property Let NumberOfColumns(ByVal col As Long)
    this.nbCols = col
End Property

Private Function RangeExists(ByVal rngName As String) As Boolean
    Dim sh As Worksheet
    Dim rng As Range

    Set sh = Wksh()
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

Private Property Get StartRow() As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim strtRow As Long
    Dim rowSet As Boolean                        'Test there is a row specification

    Set specs = Specifications()
    Set sh = Wksh()
    With sh
        'Test if there is already a startrow
        If Not RangeExists("STARTROW_" & specs.TableId) Then
            If specs.isNewSection() Or specs.TableType <> TypeTimeSeries Then
                strtRow = .Cells(.Rows.Count, STANDARDSTARTCOL).End(xlUp).row + 1
                strtRow = IIf(strtRow = 2, 10, strtRow) 'Be sure to Start after the 1st lines
                strtRow = IIf(specs.TableType <> TypeGlobalSummary, strtRow + 4, strtRow) 'Do not shift for global summary
                .Cells(strtRow, STANDARDSTARTCOL).Name = "STARTROW_" & specs.TableId
            ElseIf (Not specs.isNewSection) Then
                strtRow = Previous.StartRow      'For time series on same section, go back to precendent start row
            End If
            StartRow = strtRow
        Else
            StartRow = .Range("STARTROW_" & specs.TableId).row
        End If

    End With
End Property

Private Property Get Previous() As ICrossTable
    'Previous crosstable
    Dim prevSpecs As ITablesSpecs
    Dim specs As ITablesSpecs
    Dim sh As Worksheet
    Dim trans As ITranslation

    Set specs = Specifications
    Set prevSpecs = specs.Previous
    Set sh = Wksh()
    Set trans = Trads()

    Set Previous = Create(prevSpecs, sh, trans)
End Property

Private Property Get StartColumn() As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim strtCol As Long

    Set specs = Specifications
    Set sh = Wksh
    With sh
        If Not RangeExists("STARTCOL_" & specs.TableId()) Then
            'On time series, the startcolumn depends on new section

            strtCol = IIf(specs.TableType = TypeTimeSeries And (Not specs.isNewSection()), _
                          .Cells(StartRow + 1, .Columns.Count).End(xlToLeft).Column, _
                          STANDARDSTARTCOL)

            StartColumn = strtCol
            .Cells(StartRow - 1, strtCol + 1).Name = "STARTCOL_" & specs.TableId
        Else
            StartColumn = .Range("STARTCOL_" & specs.TableId).Column - 1
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

Public Property Get HeaderRange() As Range
    Dim start As Long
    Dim sh As Worksheet
    Set sh = Wksh()
    With sh
        Set HeaderRange = .Range(.Cells(StartRow, StartColumn + 1), .Cells(StartRow, EndColumn))
    End With

End Property

Public Property Get RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, Optional ByVal onlyCategories As Boolean = False) As Range
    Dim start As Long
    Dim rend As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs

    Set specs = Specifications
    If specs.TableType = TypeGlobalSummary Then ThrowExecutionError "Global summary does not have a row category"
    Set sh = Wksh()
    'Start Row
    Select Case specs.TableType()

    Case TypeBivariate, TypeTimeSeries, TypeSpatial
        start = StartRow + 1
    Case TypeUnivariate, TypeGlobalSummary
        start = StartRow
    End Select
    If Not includeHeaders Then start = start + 1

    'End Row
    rend = Switch(specs.TableType = TypeTimeSeries, EndRow - 2, specs.hasTotal And onlyCategories, EndRow - 1, True, EndRow)

    With sh
        Set RowsCategoriesRange = .Range(.Cells(start, STANDARDSTARTCOL), .Cells(rend, STANDARDSTARTCOL))
    End With

End Property

'Test if a value Exists in a Range
Private Function ValueExists(ByVal rng As Range, ByVal Value As String)
    ValueExists = (Not rng.Find(What:=Value, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Function

Public Property Get RowRange(ByVal rowVal As String, Optional ByVal includeHeaders As Boolean = True) As Range
    Dim rng As Range
    Dim valueRow As Long
    Dim sh As Worksheet
    Dim start As Long

    Set rng = RowsCategoriesRange()
    Set sh = Wksh()

    'Test if Missing exists
    If ValueExists(rng, rowVal) Then
        valueRow = rng.Find(What:=rowVal, LookAt:=xlWhole, MatchCase:=True).row
        start = IIf(includeHeaders, StartColumn, StartColumn + 1)
        With sh
            Set RowRange = .Range(.Cells(valueRow, start), .Cells(valueRow, EndColumn))
        End With
    End If
End Property

Public Property Get ColumnRange(ByVal colVal As String, Optional ByVal includeHeaders = False) As Range
    Dim rng As Range
    Dim valueColumn As Long
    Dim start As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs

    Set rng = HeaderRange()
    Set sh = Wksh()
    Set specs = Specifications

    start = IIf(specs.TableType = TypeBivariate, StartRow + 2, StartRow + 1)
    If includeHeaders Then start = start - 1

    'Test if the value exists
    If ValueExists(rng, colVal) Then
        valueColumn = rng.Find(What:=colVal, LookAt:=xlWhole, MatchCase:=True).Column
        With sh
            Set ColumnRange = .Range(.Cells(start, valueColumn), .Cells(EndRow, valueColumn))
        End With
    End If
End Property

Private Sub AddRows()
    Dim start As Long
    Dim specs As ITablesSpecs
    Dim sh As Worksheet
    Dim vars As ILLVariables
    Dim rowVar As String
    Dim rowsData As BetterArray
    Dim trans As ITranslation
    Dim missingValue As String
    Dim transMissing As String
    Dim transTotal As String
    Dim counter As Long

    start = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Trads()
    Set vars = LLVariables.Create(specs.Dictionary)

    missingValue = specs.Value("missing")
    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    Set rowsData = specs.RowCategories()

    With sh
        Select Case specs.TableType()
        Case TypeGlobalSummary

            .Cells(start, StartColumn).Value = specs.Value("label")

        Case TypeUnivariate

            If missingValue = "yes" Then rowsData.Push transMissing
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(start + 1, StartColumn)
            rowVar = specs.Value("row")
            .Cells(start, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)

            EndRow = start + rowsData.Length

        Case TypeBivariate

            If missingValue = "row" Or missingValue = "all" Then rowsData.Push transMissing
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(start + 2, StartColumn)
            rowVar = specs.Value("row")
            .Cells(start + 1, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)

            'Define the endrow
            EndRow = start + 1 + rowsData.Length

        Case TypeTimeSeries
            'Add rows for time series only on new sections
            If specs.isNewSection() Then

                rowsData.Clear
                rowsData.Push trans.TranslatedValue("MSG_Period"), ""
                For counter = 1 To 13
                    rowsData.Push "", "", "", ""
                Next
                rowsData.Push transTotal, transMissing
                rowsData.ToExcelRange .Cells(start + 1, StartColumn)

                EndRow = start + rowsData.Length
            End If

        Case TypeSpatial
            rowVar = specs.Value("row")
            .Cells(start + 1, StartColumn).Value = vars.Value(colName:="sub section", varName:="adm1" & "_" & rowVar)
            EndRow = start + 1 + CLng(specs.Value("geo max"))
            .Cells(EndRow + 1, StartColumn).Value = transTotal
        End Select
    End With
End Sub

Private Function Percent(ByVal percentVal As String, Optional ByVal percentType As String = "all") As String
    Select Case percentType
    Case "row"
        Percent = percentVal & " " & ChrW(8596)
    Case "col"
        Percent = percentVal & " " & ChrW(8597)
    Case Else
        Percent = percentVal
    End Select
End Function

Private Sub BivariateColumns(ByVal start As Long, ByVal specs As ITablesSpecs, ByVal trans As ITranslation, _
                             ByVal sh As Worksheet)
    Dim colVar As String
    Dim colLab As String
    Dim colData As BetterArray
    Dim labData As BetterArray
    Dim missingValue As String
    Dim percentageValue As String
    Dim transMissing As String
    Dim transTotal As String
    Dim transPercent As String
    Dim counter As Long
    Dim tobound As Long
    Dim totalValue As String

    Set colData = New BetterArray
    Set labData = New BetterArray
    colData.LowerBound = 1
    labData.LowerBound = 1
    missingValue = specs.Value("missing")
    percentageValue = specs.Value("percentage")
    totalValue = specs.Value("total")
    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    transPercent = trans.TranslatedValue("MSG_Percent")

    With sh
        colLab = specs.Value("label")
        If missingValue = "column" Or missingValue = "all" Or missingValue = "yes" Then labData.Push transMissing
        If specs.hasTotal Then labData.Push transTotal
        Set colData = specs.ColumnCategories()

        If colData.Length > 0 Then

            NumberOfColumns = colData.Length  'Set the number of columns of the categories

            colData.Concat labData.Items
            labData.Clear
            'Percentage
            If specs.HasPercentage Then
                counter = 1
                tobound = 2 * colData.UpperBound

                Do While counter <= tobound

                    colData.Splice counter + 1, 0, ""
                    counter = counter + 2
                    labData.Push colLab, Percent(transPercent, percentageValue)
                Loop
            Else
                For counter = 1 To colData.UpperBound
                    labData.Push colLab
                Next
            End If
            colData.ToExcelRange .Cells(start, StartColumn + 1), TransposeValues:=True
            labData.ToExcelRange .Cells(start + 1, StartColumn + 1), TransposeValues:=True
        Else                                     'No column Data, just a summary function
            labData.Splice 1, 0, colLab
            labData.ToExcelRange .Cells(start + 1, StartColumn + 1), TransposeValues:=True
        End If

        EndColumn = StartColumn + labData.Length() 'Set the endcolumn of the table
    End With
End Sub

'Add Columns
Private Sub AddColumns()
    Dim start As Long
    Dim specs As ITablesSpecs
    Dim sh As Worksheet
    Dim vars As ILLVariables
    Dim trans As ITranslation
    Dim missingValue As String
    Dim percentageValue As String
    Dim transMissing As String
    Dim transTotal As String
    Dim transPercent As String
    Dim customStartCol As Long
    Dim colLab As String
    Dim colData As BetterArray
    Dim colVarLab As String
    Dim timeVarLab As String


    start = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = Trads()
    Set vars = LLVariables.Create(specs.Dictionary)
    Set colData = New BetterArray
    colData.LowerBound = 1

    missingValue = specs.Value("missing")
    percentageValue = specs.Value("percentage")
    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    transPercent = trans.TranslatedValue("MSG_Percent")

    With sh
        Select Case specs.TableType()
        Case TypeGlobalSummary
            'Test if the global summary is set

            If Not RangeExists("COLGS_SET") Then
                .Cells(start - 1, StartColumn + 1).Value = trans.TranslatedValue("MSG_AllData")
                .Cells(start - 1, StartColumn + 2).Value = trans.TranslatedValue("MSG_FilteredData")
                .Cells(start - 1, StartColumn + 1).Name = "COLGS_SET"
            End If
            EndColumn = StartColumn + 2
            NumberOfColumns = 2

        Case TypeUnivariate

            colLab = specs.Value("label")
            colData.Push colLab
            If percentageValue = "yes" Then colData.Push transPercent
            colData.ToExcelRange .Cells(start, StartColumn + 1), TransposeValues:=True

            EndColumn = StartColumn + colData.Length
            NumberOfColumns = colData.Length

        Case TypeBivariate

            BivariateColumns start, specs, trans, sh
            .Range("STARTCOL_" & specs.TableId).Value = vars.Value(colName:="main label", varName:=specs.Value("column"))

        Case TypeTimeSeries

            BivariateColumns start, specs, trans, sh
            colVarLab = vars.Value(colName:="main label", varName:=specs.Value("column"))
            timeVarLab = vars.Value(colName:="main label", varName:=specs.Value("time"))

            'Add the time variable label
            .Range("STARTCOL_" & specs.TableId).Value = colVarLab & " " & trans.TranslatedValue("MSG_per") & " " & timeVarLab

        Case TypeSpatial

            BivariateColumns start, specs, trans, sh
        End Select
    End With
End Sub

'Create the headers
Private Sub AddHeader()

    Dim start As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim counter As Long
    Dim trans As ITranslation
    Dim rng As Range

    Set sh = Wksh()
    Set specs = Specifications
    Set trans = Trads()
    start = StartRow

    With sh
        Select Case specs.TableType

        Case TypeUnivariate, TypeBivariate

            .Cells(start - 2, STANDARDSTARTCOL).Value = specs.Value("title")
            .Cells(start - 2, STANDARDSTARTCOL).Name = "TITLE_" & specs.TableId()

            If specs.isNewSection() And (Not RangeExists("SECTION_" & specs.TableId())) Then
                For counter = 1 To 3
                    .Rows(start - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                'Jump two rows before the title for the section
                .Range("TITLE_" & specs.TableId()).Cells(-1, 1) = specs.Value("section")
                .Range("TITLE_" & specs.TableId()).Cells(-1, 1).Name = "SECTION_" & specs.TableId()
            End If

        Case TypeGlobalSummary

            If Not RangeExists("ROWGS_SET") Then
                .Cells(start - 2, STANDARDSTARTCOL).Value = trans.TranslatedValue("MSG_GlobalSummary")
                .Cells(start - 2, STANDARDSTARTCOL).Name = "ROWGS_SET"
            End If

        Case TypeTimeSeries

            If specs.isNewSection() And (Not RangeExists("SECTION_" & specs.TableId())) Then
                For counter = 1 To 4
                    .Rows(start - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Cells(start - 2, StartColumn).Value = specs.Value("section")
                .Cells(start - 2, StartColumn).Name = "SECTION_" & specs.TableId()

                'Ranges for start and end dates and for minimum and maximum dates (for validations)
                Set rng = .Range("SECTION_" & specs.TableId())

                'Start Date
                rng.Cells(3, 3).Value = trans.TranslatedValue("MSG_StartDate")
                rng.Cells(4, 3).Name = "START_DATE_" & specs.TableId()
                rng.Cells(5, 3).Name = "INFO_START_DATE_" & specs.TableId()

                'Time Unit
                rng.Cells(3, 5).Value = trans.TranslatedValue("MSG_TimeUnit")
                rng.Cells(4, 5).Name = "TIME_UNIT_" & specs.TableId()

                'End Date
                rng.Cells(3, 7).Value = trans.TranslatedValue("MSG_EndDate")
                rng.Cells(4, 7).Name = "END_DATE_" & specs.TableId()
                rng.Cells(5, 7).Name = "INFO_END_DATE_" & specs.TableId()


                'Minimum cell for validiation
                rng.Cells(4, 9).Name = "MIN_DATE_" & specs.TableId()
                rng.Cells(4, 11).Name = "MAX_DATE_" & specs.TableId()
                'Maximum cell for validation
            End If

        Case TypeSpatial

            If Not RangeExists("ADM_DROPDOWN_" & specs.TableId()) Then
                .Rows(start - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                'Table Title
                .Cells(start - 3, STANDARDSTARTCOL).Value = specs.Value("title")
                .Cells(start - 3, STANDARDSTARTCOL).Name = "TITLE_" & specs.TableId()

                'Range for dropdown for the admins
                .Cells(start - 2, STANDARDSTARTCOL).Name = "ADM_DROPDOWN_" & specs.TableId()
                .Cells(start - 2, STANDARDSTARTCOL).Value = trans.TranslatedValue("MSG_SelectAdmin")
            End If

        End Select
    End With
End Sub

'Name the ranges of the table for further
Private Sub NameRanges()
    Dim specs As ITablesSpecs
    Dim sh As Worksheet
    Dim nbCols As Long
    Dim catRng As Range
    Dim rng As Range
    Dim nbcolCounter As Long
    Dim counter As Long
    Dim trans As ITranslation
    Dim labCol As Long
    Dim cData As BetterArray
    Dim tData As BetterArray
    Dim rw As Long
    Dim col As Long
    Dim addBegin As String
    Dim addEnd As String


    Set sh = Wksh()
    Set trans = Trads()

    With sh
        Set specs = Specifications

        'Test if the table has been filled
        If Not RangeExists("ENDTABLE_" & specs.TableId) Then .Cells(EndRow + 1, StartColumn).Name = "ENDTABLE_" & specs.TableId

        nbCols = NumberOfColumns
        If nbCols = 0 Then ThrowExecutionError "Execution Error: Found a table with number of columns 0: " & specs.TableId

        'The label of rows categories of the table (with the missing, will serve for graphs)

        Set catRng = RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False)
        catRng.Name = "ROW_CATEGORIES_" & specs.TableId
        catRng.Cells(0, 1).Name = "LABEL_ROW_CATEGORIES_" & specs.TableId

        'All columns of the data
        counter = StartColumn
        nbcolCounter = 1
        Set rng = catRng.Offset(, 1)

        Do While (counter < EndColumn) And (nbcolCounter <= nbCols)
            rng.Name = "VALUES_COL_" & nbcolCounter & "_" & specs.TableId

            Select Case specs.TableType

                Case TypeBivariate, TypeTimeSeries, TypeSpatial
                    labCol = -1
                Case TypeUnivariate, TypeGlobalSummary
                    labCol = 0
            End Select

            'Name the header witch will serve as labels
            rng.Cells(labCol, 1).Name = "LABEL_COL_" & nbcolCounter & "_" & specs.TableId

            'Movte to next Range
            Set rng = rng.Offset(, 1)

            If specs.HasPercentage Then
                rng.Name = "PERC_COL_" & nbcolCounter & "_" & specs.TableId
                Set rng = rng.Offset(, 1)
                counter = counter + 1
            End If

            counter = counter + 1
            nbcolCounter = nbcolCounter + 1
        Loop

        Set cData = New BetterArray
        Set tData = New BetterArray

        cData.Push "MISSING", "TOTAL"
        tData.Push "MSG_NA", "MSG_Total"

        For counter = cData.LowerBound To cData.UpperBound
             'Missing/Total Row (with label)

            Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=True)

            If (Not rng Is Nothing) Then
                rng.Name = cData.Item(counter) & "_ROW_" & specs.TableId
                Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=False)
                rng.Name = cData.Item(counter) & "_ROW_VALUES_" & specs.TableId
            End If

            'Missing /Total Column
            Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=True)
            If (Not rng Is Nothing) Then
                rng.Name = cData.Item(counter) & "_COL_" & specs.TableId
                rng.Cells(0, 1).Name = cData.Item(counter) & "_LABEL_COL_" & specs.TableId
                Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=False)
                rng.Name = cData.Item(counter) & "_COL_VALUES_" & specs.TableId
                'Percentages on Missing/Total column
                If specs.HasPercentage Then
                 Set rng = rng.Offset(, 1)
                 rng.Name = cData.Item(counter) & "_PERC_VALUES_" & specs.TableId
                 'All the percentage columns (including label)
                 .Range(cData.Item(counter) & "_COL_" & specs.TableId).Offset(, 1).Name = cData.Item(counter) & "_PERC_" & specs.TableId
                End If
            End If
        Next

        'Intersection of total and missings
        If RangeExists("TOTAL_COL_" & specs.TableId) Then
            'Total and Total
            If RangeExists("TOTAL_ROW_" & specs.TableId) Then
                'Total row
                rw = .Range("TOTAL_ROW_" & specs.TableId).row
                col = .Range("TOTAL_COL_" & specs.TableId).Column
                .Cells(rw, col).Name = "TOTAL_TOTAL_" & specs.TableId
            End If
            If RangeExists("MISSING_ROW_" & specs.TableId) Then
                rw = .Range("MISSING_ROW_" & specs.TableId).row
                col = .Range("TOTAL_COL_" & specs.TableId).Column
                .Cells(rw, col).Name = "MISSING_TOTAL_" & specs.TableId
            End If
        End If

        If RangeExists("TOTAL_ROW_" & specs.TableId) Then
            'Total and Total
            If RangeExists("MISSING_COL_" & specs.TableId) Then
                col = .Range("MISSING_COL_" & specs.TableId).Column
                rw = .Range("TOTAL_ROW_" & specs.TableId).row
                .Cells(rw, col).Name = "TOTAL_MISSING_" & specs.TableId
            End If
        End If

        'All the interior of the table
        addBegin = Split(.Range("VALUES_COL_1_" & specs.TableId).Address, ":")(0)
        addEnd = IIf(specs.HasPercentage, "PERC_COL_", "VALUES_COL_")
        addEnd = Split(.Range(addEnd & nbCols & "_" & specs.TableId).Address, ":")(1)
        .Range(addBegin & ":" & addEnd).Name = "INTERIOR_VALUES_" & specs.TableId

        'Header Row with column categories
        Set rng = HeaderRange()
        rng.Name = "COLUMN_CATEGORIES_" & specs.TableId
    End With
End Sub

'Build the table
Private Sub Build()
  AddRows 'Add categories for rows and column
  AddColumns
  AddHeaders 'Add the headers of the table
  NameRanges 'Name required ranges of the table (for plotting purposes
End Sub

'Get the formula of the table

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowExecutionError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error : " & errorMessage
End Sub

Public Sub ICrossTable_AddRows()
    AddRows
End Sub

Public Sub ICrossTable_AddColumns()
    AddColumns
End Sub

Public Sub ICrossTable_AddHeader()
    AddHeader
End Sub

Public Sub ICrossTable_NameRanges()
    NameRanges
End Sub

Public Sub ICrossTable_Build()
  Build
End Sub

Public Property Get ICrossTable_HeaderRange() As Range
    Set ICrossTable_HeaderRange = HeaderRange
End Property

Public Property Get ICrossTable_RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, Optional ByVal onlyCategories As Boolean = False) As Range
    Set ICrossTable_RowsCategoriesRange = RowsCategoriesRange(includeHeaders, onlyCategories)
End Property

Public Property Get ICrossTable_ColumnRange(ByVal translationKey As String, _
                                            Optional ByVal includeHeaders = False) As Range
    Set ICrossTable_ColumnRange = ColumnRange(translationKey, includeHeaders)
End Property

Public Property Get ICrossTable_RowRange(ByVal translationKey As String) As Range
    Set ICrossTable_RowRange = RowRange(translationKey)
End Property

Public Property Get ICrossTable_StartRow() As Long
    ICrossTable_StartRow = StartRow
End Property


