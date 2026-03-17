VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Create and format crosstables for analysis"

'@IgnoreModule IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Create and format crosstables for analysis")

Option Explicit

'Exposed methods
Implements ICrossTable

Private Type TCrossTable
    specs As ITablesSpecs
    OutputSh As Worksheet
    trans As ITranslation
    endrw As Long
    endCol As Long
    nbCols As Long
    tsHeader As String 'Header for time series (to use in to go section)
End Type

Private this As TCrossTable
Private Const STANDARDSTARTCOL As Long = 3
Private Const CLASSNAME As String = "CrossTable"
Private Const NBROWSTIMESERIES As Long = 56
Private ALLRANGESNAMES As BetterArray

Public Function Create(ByVal specs As ITablesSpecs, ByVal outputWksh As Worksheet, _
                       ByVal trans As ITranslation) As ICrossTable
    'Check if values are not defined
    If specs Is Nothing Then ThrowErrorUnableInstanciate " Table specifications not defined"
    If outputWksh Is Nothing Then ThrowErrorUnableInstanciate " Output worksheet not defined"
    If trans Is Nothing Then ThrowErrorUnableInstanciate " Translations not defined"

    With New CrossTable
        Set .Specifications = specs
        Set .Wksh = outputWksh
        Set .trads = trans
        Set Create = .Self
    End With

End Function

Public Property Get Self() As ICrossTable
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.OutputSh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.OutputSh = sh
End Property

Public Property Get Specifications() As ITablesSpecs
    Set Specifications = this.specs
End Property

Public Property Set Specifications(ByVal specs As ITablesSpecs)
    Set this.specs = specs
End Property

Public Property Get trads() As ITranslation
    Set trads = this.trans
End Property

Public Property Set trads(ByVal trans As ITranslation)
    Set this.trans = trans
End Property

Private Property Get NumberOfColumns() As Long
    NumberOfColumns = this.nbCols
End Property

Private Property Let NumberOfColumns(ByVal Col As Long)
    this.nbCols = Col
End Property

Private Property Get NamedRangesList() As BetterArray
    Set NamedRangesList = ALLRANGESNAMES.Clone()
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
    Dim tabId As String                          'Test there is a row specification

    Set specs = Specifications()
    tabId = specs.TableId
    Set sh = Wksh()
    With sh
        'Test if there is already a startrow
        If Not RangeExists("STARTROW_" & tabId) Then
            If specs.isNewSection() Or (specs.TableType <> TypeTimeSeries And specs.TableType <> TypeSpatioTemporal) Then
                strtRow = .Cells(.Rows.Count, STANDARDSTARTCOL).End(xlUp).Row + 1
                strtRow = IIf(strtRow = 2, 6, strtRow) 'Be sure to Start after the 1st lines
                
                strtRow = Switch(specs.TableType = TypeSpatial, strtRow + 6, _
                                 specs.TableType = TypeGlobalSummary, _
                                 strtRow, True, strtRow + 4) 'More shifts for spatial, no shift for global summary
                
                .Cells(strtRow, STANDARDSTARTCOL).Name = "STARTROW_" & tabId
            ElseIf (Not specs.isNewSection) Then
                strtRow = Previous.StartRow      'For time series on same section, go back to precendent start row
            End If
            StartRow = strtRow
        Else
            StartRow = .Range("STARTROW_" & tabId).Row
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
    Set trans = trads()

    Set Previous = Create(prevSpecs, sh, trans)
End Property

Private Property Get StartColumn() As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim strtcol As Long
    Dim tabId As String
    Dim tablType As Byte

    Set specs = Specifications
    tabId = specs.TableId
    tablType = specs.TableType

    Set sh = Wksh
    With sh
        If Not RangeExists("STARTCOL_" & tabId) Then
            'On time series, the startcolumn depends on new section
            If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And _
                 (Not specs.isNewSection()) Then
                strtcol = .Cells(StartRow + 1, .Columns.Count).End(xlToLeft).Column
            Else
                strtcol = STANDARDSTARTCOL
            End If
            StartColumn = strtcol
            .Cells(StartRow - 1, strtcol + 1).Name = "STARTCOL_" & tabId
        Else
            StartColumn = .Range("STARTCOL_" & tabId).Column - 1
        End If
    End With
End Property

Private Property Let EndRow(ByVal rw As Long)
    this.endrw = rw
End Property

Private Property Let EndColumn(ByVal Col As Long)
    this.endCol = Col
End Property

Private Property Get EndRow() As Long
    EndRow = this.endrw
End Property

Private Property Get EndColumn() As Long
    EndColumn = this.endCol
End Property

Public Property Get HeaderRange() As Range
    Dim sh As Worksheet
    Set sh = Wksh()
    With sh
        Set HeaderRange = .Range(.Cells(StartRow, StartColumn + 1), .Cells(StartRow, EndColumn))
    End With

End Property

'Header for each part of the time series
Private Property Get TimeSeriesHeader() As String
    TimeSeriesHeader = this.tsHeader
End Property

Private Property Let TimeSeriesHeader(ByVal headerText As String)
    this.tsHeader = headerText
End Property

Public Property Get RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, Optional ByVal onlyCategories As Boolean = False) As Range
    Dim Start As Long
    Dim rend As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim tablType As Byte

    Set specs = Specifications
    tablType = specs.TableType

    If tablType = TypeGlobalSummary Then ThrowExecutionError "Global summary does not have a row category"
    Set sh = Wksh()
    'Start Row
    Select Case tablType

    Case TypeBivariate, TypeTimeSeries, TypeSpatial, TypeSpatioTemporal
        Start = StartRow + 1
    Case TypeUnivariate, TypeGlobalSummary
        Start = StartRow
    End Select
    If Not includeHeaders Then Start = Start + 1

    'End Row
    rend = Switch((tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And onlyCategories, EndRow - 2, _
                  (specs.HasTotal Or tablType = TypeSpatial) And onlyCategories, EndRow - 1, _
                  True, EndRow)

    With sh
        Set RowsCategoriesRange = .Range(.Cells(Start, STANDARDSTARTCOL), .Cells(rend, STANDARDSTARTCOL))
    End With

End Property

'Test if a value Exists in a Range
Private Function ValueExists(ByVal rng As Range, ByVal Value As String) As Boolean
    ValueExists = (Not rng.Find(What:=Value, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
End Function

Public Property Get RowRange(ByVal rowVal As String, Optional ByVal includeHeaders As Boolean = True) As Range
    Dim rng As Range
    Dim valueRow As Long
    Dim sh As Worksheet
    Dim Start As Long

    Set rng = RowsCategoriesRange()
    Set sh = Wksh()

    'Test if Missing exists
    If ValueExists(rng, rowVal) Then
        valueRow = rng.Find(What:=rowVal, lookAt:=xlWhole, MatchCase:=True).Row
        
        If includeHeaders Then
            Start = StartColumn
        Else 
            start = StartColumn + 1
        End If

        With sh
            Set RowRange = .Range(.Cells(valueRow, Start), .Cells(valueRow, EndColumn))
        End With
    End If
End Property

Public Property Get ColumnRange(ByVal colVal As String, Optional ByVal includeHeaders As Boolean = False, Optional ByVal onlyCategories As Boolean = False) As Range
    Dim rng As Range
    Dim valueColumn As Long
    Dim Start As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim rend As Long
    Dim tablType As Byte

    Set rng = HeaderRange()
    Set sh = Wksh()
    Set specs = Specifications
    tablType = specs.TableType

    'startRow
    Start = Switch(tablType = TypeBivariate, StartRow + 2, _
                   (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal), StartRow + 2, _
                   tablType = TypeSpatial, StartRow + 2, _
                   True, StartRow + 1)
    'End Row
    rend = Switch((tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And onlyCategories, EndRow - 2, _
                  (specs.HasTotal Or tablType = TypeSpatial) And onlyCategories, EndRow - 1, _
                  True, EndRow)

    If includeHeaders Then Start = Start - 1

    'Test if the value exists
    If ValueExists(rng, colVal) Then
        valueColumn = rng.Find(What:=colVal, lookAt:=xlWhole, MatchCase:=True).Column
        With sh
            Set ColumnRange = .Range(.Cells(Start, valueColumn), .Cells(rend, valueColumn))
        End With
    End If
End Property

Private Sub AddRows()
    Dim Start As Long
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
    Dim spatType As String

    Start = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = trads()
    Set vars = LLVariables.Create(specs.Dictionary)

    missingValue = specs.Value("missing")
    transMissing = trans.TranslatedValue("MSG_NA")
    transTotal = trans.TranslatedValue("MSG_Total")
    Set rowsData = specs.RowCategories()

    With sh
        Select Case specs.TableType()
        Case TypeGlobalSummary

            .Cells(Start, StartColumn).Value = specs.Value("label")

        Case TypeUnivariate

            If missingValue = "yes" Then rowsData.Push transMissing
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(Start + 1, StartColumn)
            rowVar = specs.Value("row")
            .Cells(Start, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)

            EndRow = Start + rowsData.Length


        Case TypeBivariate

            If missingValue = "row" Or missingValue = "all" Then rowsData.Push transMissing
            rowsData.Push transTotal
            rowsData.ToExcelRange .Cells(Start + 2, StartColumn)
            rowVar = specs.Value("row")
            .Cells(Start + 1, StartColumn).Value = vars.Value(colName:="main label", varName:=rowVar)

            'Define the endrow
            EndRow = Start + 1 + rowsData.Length

        Case TypeTimeSeries, TypeSpatioTemporal
            'Add rows for time series only on new sections
            If specs.isNewSection() Then

                rowsData.Clear
                rowsData.Push trans.TranslatedValue("MSG_Period"), vbNullString
                For counter = 1 To 13
                    rowsData.Push vbNullString, vbNullString, _
                                  vbNullString, vbNullString
                Next
                rowsData.Push transTotal, transMissing
                rowsData.ToExcelRange .Cells(Start + 1, StartColumn)

            End If

            EndRow = Start + NBROWSTIMESERIES    'This is the number of rows (FIXED for time series)

        Case TypeSpatial
            rowVar = specs.Value("row")
            rowsData.Clear
            spatType = specs.SpatialTableTypes()
            
            
            If (spatType = "geo") Then
                rowsData.Push vars.Value(colName:="sub section", varName:="adm1_" & rowVar)
            ElseIf (spatType = "hf") Then
                rowsData.Push vars.Value(colName:="main label", varName:="hf_" & rowVar)
            End If

            For counter = 1 To CLng(specs.Value("geo max"))
                rowsData.Push vbNullString
            Next

            rowsData.Push transMissing
            rowsData.ToExcelRange .Cells(Start + 1, StartColumn)
            EndRow = Start + rowsData.Length
        End Select
    End With
End Sub

Private Function Percent(ByVal percentVal As String, Optional ByVal percentType As String = "all") As String
    Select Case percentType
    Case "row"
        Percent = percentVal & " " & ChrW(8596)
    Case "column"
        Percent = percentVal & " " & ChrW(8597)
    Case Else
        Percent = percentVal
    End Select
End Function

Private Sub BivariateColumns(ByVal Start As Long, ByVal specs As ITablesSpecs, ByVal trans As ITranslation, _
                             ByVal sh As Worksheet)
    Dim colLab As String
    Dim colData As BetterArray
    Dim newColData As BetterArray 'New column data to use for sending empty values when add percentage = True
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

            NumberOfColumns = colData.Length     'Set the number of columns of the categories (Property)

            colData.CONCAT labData.Items
            labData.Clear
            'Percentage
            If specs.HasPercentage Then
                counter = 1

                Do While counter <= colData.Length

                    newColData.Push colData.Item(counter), vbNullString
                    labData.Push colLab, Percent(transPercent, percentageValue)
                    counter = counter + 1
                Loop
            Else
                For counter = 1 To colData.Length
                    labData.Push colLab
                Next
                Set newColData = colData.Clone()
            End If
            
            'new col data is the column data with the empty values between elements
            newColData.ToExcelRange .Cells(Start, StartColumn + 1), TransposeValues:=True
            labData.ToExcelRange .Cells(Start + 1, StartColumn + 1), TransposeValues:=True
        Else                                     'No column Data, just a summary function
            labData.Push colLab
            NumberOfColumns = 1
            labData.ToExcelRange .Cells(Start, StartColumn + 1), TransposeValues:=True
            
            'This is a quick fix to include calculated columns (not splitted by a group by) to time series.
            'I use the End(xltoLeft) to find the next startcolumn, so something needs to be puts in this column
            .Cells(Start + 1, StartColumn + 1).Value = " "
        End If

        EndColumn = StartColumn + labData.Length() 'Set the endcolumn of the table
    End With
End Sub

'Add Columns
Private Sub AddColumns()
    Dim Start As Long
    Dim specs As ITablesSpecs
    Dim sh As Worksheet
    Dim vars As ILLVariables
    Dim trans As ITranslation
    Dim percentageValue As String
    Dim transPercent As String
    Dim colLab As String
    Dim colData As BetterArray
    Dim colVarLab As String
    Dim sumLab As String
    Dim timeVarLab As String
    Dim tabId As String
    Dim tablType As Byte
    Dim spType As String

    Start = StartRow
    Set sh = Wksh()
    Set specs = Specifications()
    Set trans = trads()
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
            'Test if the global summary is set

            If Not RangeExists("COLGS_SET") Then
                .Cells(Start - 1, StartColumn + 1).Value = trans.TranslatedValue("MSG_AllData")
                .Cells(Start - 1, StartColumn + 2).Value = trans.TranslatedValue("MSG_FilteredData")
                .Cells(Start - 1, StartColumn + 1).Name = "COLGS_SET"
            End If
            EndColumn = StartColumn + 2
            NumberOfColumns = 2

        Case TypeUnivariate

            colLab = specs.Value("label")
            colData.Push colLab
            If percentageValue = "yes" Then colData.Push transPercent
            colData.ToExcelRange .Cells(Start, StartColumn + 1), TransposeValues:=True

            EndColumn = StartColumn + colData.Length
            NumberOfColumns = 1                  'Univariate analysis has one variable on columns (even when there are percentages)

        Case TypeBivariate

            BivariateColumns Start, specs, trans, sh
            .Range("STARTCOL_" & tabId).Value = vars.Value(colName:="main label", varName:=specs.Value("column"))

        Case TypeTimeSeries, TypeSpatioTemporal

            BivariateColumns Start, specs, trans, sh

            If (tablType = TypeTimeSeries) Then
                colVarLab = vars.Value(colName:="main label", varName:=specs.Value("column"))
            ElseIf spType = "geo" Then
                colVarLab = vars.Value(colName:="sub section", varName:="adm1_" & specs.Value("column"))
            ElseIf spType = "hf" Then
                colVarLab = vars.Value(colName:="main label", varName:="hf_" & specs.Value("column"))
            End If

            timeVarLab = vars.Value(colName:="main label", varName:=specs.Value("time"))
            sumLab = specs.Value("label")

            'Use the label for header if there is no column (only an aggregate function)
            If specs.Value("column") = vbNullString Then
                TimeSeriesHeader = sumLab & " " & ChrW(9472) & " " & timeVarLab
            Else
                TimeSeriesHeader = sumLab & " " & ChrW(9472) & " " & timeVarLab & " " & ChrW(9472) & " " & colVarLab
            End If

            'Add the time variable label
            .Range("STARTCOL_" & tabId).Value = TimeSeriesHeader

        Case TypeSpatial

            BivariateColumns Start, specs, trans, sh

        End Select
    End With
End Sub

'Create the headers
Private Sub AddHeader()

    Dim Start As Long
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim counter As Long
    Dim trans As ITranslation
    Dim tabId As String
    Dim rng As Range
    Dim cellRng As Range
    Dim nGeo As Long
    Dim tablType As Byte
    Dim nbShift As Long
    Dim timeUnitName As String
    Dim sptGeorngTag As String 'Spatio - temporal geo range Tag

    Set sh = Wksh()
    Set specs = Specifications
    Set trans = trads()

    Set ALLRANGESNAMES = New BetterArray

    Start = StartRow
    tabId = specs.TableId
    tablType = specs.TableType

    sptGeorngTag = "INPUTSPTGEO_"
    If specs.Value("spatial type") = "hf" Then  sptGeorngTag = "INPUTSPTHF_"

    With sh
        Select Case tablType

        Case TypeUnivariate, TypeBivariate
            'Add headers for univariable and bivariate tables ---

            .Cells(Start - 2, STANDARDSTARTCOL).Value = specs.Value("title")
            .Cells(Start - 2, STANDARDSTARTCOL).Name = "TITLE_" & tabId
            ALLRANGESNAMES.Push "TITLE_" & tabId

            If specs.isNewSection() And (Not RangeExists("SECTION_" & tabId)) Then
                For counter = 1 To 3
                    .Rows(Start - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                'Jump two rows before the title for the section
                .Range("TITLE_" & tabId).Cells(-1, 1) = specs.Value("section")
                .Range("TITLE_" & tabId).Cells(-1, 1).Name = "SECTION_" & tabId
                ALLRANGESNAMES.Push "SECTION_" & tabId
            End If

        Case TypeGlobalSummary
            'Global summary header ---

            If Not RangeExists("ROWGS_SET") Then
                For counter = 1 To 3
                    .Rows(Start - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next

                .Cells(Start - 2, STANDARDSTARTCOL).Value = trans.TranslatedValue("MSG_GlobalSummary")
                .Cells(Start - 2, STANDARDSTARTCOL).Name = "ROWGS_SET"
                ALLRANGESNAMES.Push "ROWGS_SET"
            End If
        

        Case TypeTimeSeries, TypeSpatioTemporal
            'Time series headers are almost similar to spatio-temporal headers ---

            If specs.isNewSection() And (Not RangeExists("SECTION_" & tabId)) Then

                'For spatio temporal, we need to take in account the number of geo
                If (tablType = TypeSpatioTemporal) Then
                    nGeo = 5
                    On Error Resume Next
                    nGeo = CLng(specs.Value("n geo"))
                    On Error GoTo 0
                    'This is a bit touchy, if you want to add some rows before
                    'the table, remember to change the number of rows to shift
                    nbShift = nGeo + 14
                Else
                    nbShift = 7
                End If

                For counter = 1 To nbShift
                    .Rows(Start - 2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next
                .Cells(Start - 2, StartColumn).Value = specs.Value("section")
                .Cells(Start - 2, StartColumn).Name = "SECTION_" & tabId                
                ALLRANGESNAMES.Push "SECTION_" & tabId

                Set rng = .Range("SECTION_" & tabId)

                'For spatio-temporal, add part for selection of geo-levels

                If (tablType = TypeSpatioTemporal) Then
                    'Anchor for the future table header elements
                    Set rng = rng.Offset(4)

                    rng.Value = trans.TranslatedValue("MSG_InputGeoLevels")
                    rng.Name = "INPUTSPTGEO_LABEL_" & tabId
                    ALLRANGESNAMES.Push "INPUTSPTGEO_LABEL_" & tabId
                    rng.Cells(1, 2).Value = nGeo
                    rng.Cells(1, 2).Name = "INPUTSPTGEO_NGEO_" & tabId

                    'Add the number of geolevels

                    Set cellRng = rng.Offset(2)

                    For counter = 1 To nGeo
                        'Add HF or Geo depending on the section type
                        cellRng.Cells(counter, 1).Name = sptGeorngTag & counter & "_" & tabId
                        cellRng.Cells(counter, 2).Value = 1
                    Next

                    Set rng = .Range(cellRng.Cells(1, 1), cellRng.Cells(nGeo, 1))
                    rng.Name = "INPUTSPTGEO_USER_" & tabId
                    ALLRANGESNAMES.Push "INPUTSPTGEO_USER_" & tabId
                    
                    rng.Locked = False

                    'Where to write previous admin from 
                    rng.Offset(, 1).Name = "INPUTPREVADM_USER_" & tabId

                    'anchor for time series elements (on TimeSeries, it is the section range)
                    Set rng = cellRng.Offset(nGeo + 3)
                End If


                'Ranges for start and end dates and for minimum and maximum dates (for validations)

                'Analysis Period Date
                rng.Cells(6, 1).Name = "INFO_ANA_PERIOD_" & tabId

                'Start Date
                rng.Cells(3, 3).Value = trans.TranslatedValue("MSG_StartDate")
                rng.Cells(3, 3).Name = "START_DATE_LABEL_" & tabId
                rng.Cells(4, 3).Locked = False
                rng.Cells(4, 3).Name = "USER_START_DATE_" & tabId
                rng.Cells(5, 3).Name = "INFO_START_DATE_" & tabId
                rng.Cells(6, 3).Name = "START_DATE_" & tabId

                'Time Unit
                rng.Cells(3, 5).Value = trans.TranslatedValue("MSG_TimeUnit")
                rng.Cells(3, 5).Name = "TIME_UNIT_LABEL_" & tabId
                rng.Cells(4, 5).Name = "TIME_UNIT_" & tabId
                rng.Cells(4, 5).Locked = False

                'End Date
                rng.Cells(3, 7).Value = trans.TranslatedValue("MSG_EndDate")
                rng.Cells(3, 7).Name = "END_DATE_LABEL_" & tabId
                rng.Cells(4, 7).Name = "USER_END_DATE_" & tabId
                rng.Cells(4, 7).Locked = False
                rng.Cells(5, 7).Name = "INFO_END_DATE_" & tabId
                rng.Cells(6, 7).Name = "END_DATE_" & tabId


                'Minimum / Maximum cell for validiation
                rng.Cells(4, 9).Name = "VALIDATION_MIN_DATE_" & tabId
                rng.Cells(4, 11).Name = "VALIDATION_MAX_DATE_" & tabId
                'Minimum /Maximum of all dates

                rng.Cells(4, 13).Name = "MIN_MIN_DATE_" & tabId
                rng.Cells(4, 15).Name = "MAX_MAX_DATE_" & tabId

                ALLRANGESNAMES.Push "INFO_ANA_PERIOD_" & tabId, _ 
                    "START_DATE_LABEL_" & tabId, "USER_START_DATE_" & tabId, _ 
                    "INFO_START_DATE_" & tabId, "START_DATE_" & tabId, _ 
                    "END_DATE_LABEL_" & tabId, "USER_END_DATE_" & tabId, _ 
                    "INFO_END_DATE_" & tabId, "END_DATE_" & tabId


                timeUnitName = "TIME_UNIT_LIST"
                If tablType = TypeSpatioTemporal Then timeUnitName = "SP" & timeUnitName

                Set cellRng = .Cells(1, 26)
                If (Not RangeExists(timeUnitName)) Then
                    cellRng.Cells(1, 1).Value = trans.TranslatedValue("MSG_Day")
                    cellRng.Cells(2, 1).Value = trans.TranslatedValue("MSG_Week")
                    cellRng.Cells(3, 1).Value = trans.TranslatedValue("MSG_Month")
                    cellRng.Cells(4, 1).Value = trans.TranslatedValue("MSG_Quarter")
                    cellRng.Cells(5, 1).Value = trans.TranslatedValue("MSG_Year")
                    .Range(.Cells(1, cellRng.Column), .Cells(5, cellRng.Column)).Name = timeUnitName
                    .Range(timeUnitName).Font.color = vbWhite
                    .Range(timeUnitName).Interior.color = vbWhite
                    .Range(timeUnitName).Locked = True
                End If
                    
                'Add week As default time
                .Range("TIME_UNIT_" & tabId).Value = cellRng.Cells(2, 1).Value
            End If

        Case TypeSpatial
        
             'Add Dropdown for selecting the admins
            If (Not RangeExists("ADMIN_UNIT_LIST")) Then

                Set cellRng = .Cells(1, 26)
                cellRng.Cells(1, 1).formula = "= RNG_ADM1NAME"
                cellRng.Cells(2, 1).formula = "= RNG_ADM2NAME"
                cellRng.Cells(3, 1).formula = "= RNG_ADM3NAME"
                cellRng.Cells(4, 1).formula = "= RNG_ADM4NAME"
                
                .Range(.Cells(1, cellRng.Column), .Cells(4, cellRng.Column)).Name = "ADM_UNIT_LIST"
                .Range("ADM_UNIT_LIST").Font.color = vbWhite
                .Range("ADM_UNIT_LIST").Interior.color = vbWhite
                .Range("ADM_UNIT_LIST").Locked = True
            End If

            'Range for selecting values for the population
            If (Not RangeExists("POPULATION_FACTOR_LIST")) Then
                Set cellRng = .Cells(1, 27)
                cellRng.Cells(1, 1).Value = trans.TranslatedValue("MSG_NoDevide")
                cellRng.Cells(2, 1).Value = trans.TranslatedValue("MSG_Devide")

                .Range(.Cells(1, cellRng.Column), .Cells(2, cellRng.Column)).Name = "POPULATION_FACTOR_LIST"
                .Range("POPULATION_FACTOR_LIST").Font.color = vbWhite
                .Range("POPULATION_FACTOR_LIST").Interior.color = vbWhite
                .Range("POPULATION_FACTOR_LIST").Locked = True
            End If

            If (Not RangeExists("ADM_DROPDOWN_" & tabId)) And (Not RangeExists("HF_DROPDOWN_" & tabId)) Then

                If (specs.SpatialTableTypes = "geo") Then
                
                    'Range for dropdown for the admins
                    .Cells(Start - 2, STANDARDSTARTCOL).Name = "ADM_DROPDOWN_" & tabId
                    'Add the dropdown only on spatial tables of type geo
                    .Cells(Start - 2, STANDARDSTARTCOL).Value = trans.TranslatedValue("MSG_SelectAdmin")

                    Set rng = .Range("ADM_DROPDOWN_" & tabId)
                    'Previous value selected for ADM_DROPDOWN_
                    rng.Cells(1, 2).Value = "adm1"
                    rng.Cells(1, 2).Name = "PREVIOUS_ADM_" & tabId

                    'population factor dropdown
                    rng.Cells(1, 3).Value = trans.TranslatedValue("MSG_SelectPOPFACT")
                    rng.Cells(1, 3).Name = "DEVIDEPOP_" & tabId

                    'Add population to range 4
                    rng.Cells(0, 4).Value = trans.TranslatedValue("MSG_MultiplyBy")
                    rng.Cells(0, 4).Name = "POPFACTLABEL_" & tabId
                    rng.Cells(1, 4).Value = 100
                    rng.Cells(1, 4).Name = "POPFACT_" & tabId
                

                    'Add name to button where you divide by the population
                    rng.Cells(1, 5).Name = "POPPREVFACT_" & tabId
                    rng.Cells(1, 5).Value = "0"

                    ALLRANGESNAMES.Push "DEVIDEPOP_" & tabId, "POPFACT_" & tabId, _ 
                                        "ADM_DROPDOWN_" & tabId
                
                ElseIf (specs.SpatialTableTypes = "hf") Then
                    'Range for information on Type of admin level
                    .Cells(Start - 2, STANDARDSTARTCOL).Name = "HF_DROPDOWN_" & tabId
                    Set rng = .Range("HF_DROPDOWN_" & tabId)
                    rng.Cells(1, 1).Value = trans.TranslatedValue("MSG_HF")

                    ALLRANGESNAMES.Push "HF_DROPDOWN_" & tabId
                End If

                'New Section
                If specs.isNewSection() Then
                    For counter = 1 To 3
                        .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    Next
                    rng.Cells(-3, 1).Value = specs.Value("section")
                    rng.Cells(-3, 1).Name = "SECTION_" & tabId

                    ALLRANGESNAMES.Push "SECTION_" & tabId
                End If
                
                'Table Title
                .Rows(rng.Row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                rng.Cells(-1, 1).Value = specs.Value("title")
                rng.Cells(-1, 1).Name = "TITLE_" & tabId

                ALLRANGESNAMES.Push "TITLE_" & tabId
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
    Dim Col As Long
    Dim addBegin As String
    Dim addEnd As String
    Dim tabId As String
    Dim cellRng As Range
    Dim tablType As AnalysisTablesType
    Dim isNewSection As Boolean


    Set sh = Wksh()
    Set trans = trads()
    Set specs = Specifications
    tabId = specs.TableId
    tablType = specs.TableType
    isNewSection = specs.isNewSection()

    If tablType = TypeGlobalSummary Then Exit Sub 'No name ranges for global summary

    With sh

        'Test if the table has been filled
        If Not RangeExists("ENDTABLE_" & tabId) Then .Cells(EndRow + 1, StartColumn).Name = "ENDTABLE_" & tabId

        nbCols = NumberOfColumns
        If nbCols = 0 Then ThrowExecutionError "Execution Error: Found a table with number of columns 0: " & tabId

        'The label of rows categories of the table (with the missing, will serve for graphs)
        If isNewSection Or (tablType <> TypeTimeSeries And tablType <> TypeSpatioTemporal) Then

            Set catRng = RowsCategoriesRange(onlyCategories:=True, includeHeaders:=False)
            catRng.Name = "ROW_CATEGORIES_" & tabId
            catRng.Cells(0, 1).Name = "LABEL_ROW_CATEGORIES_" & tabId
            
            'Stock the row and label row categories name somewhere
            ALLRANGESNAMES.Push "LABEL_ROW_CATEGORIES_" & tabId
            ALLRANGESNAMES.Push "ROW_CATEGORIES_" & tabId

            'Set the start time period and last time period dates
            If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) Then
                Set rng = catRng.Offset(, -1)
                rng.Name = "END_TIME_PERIOD_" & tabId
                Set rng = rng.Offset(, -1)
                rng.Cells(1, 1).Name = "FIRST_VALUE_START_TIME_" & tabId
                Set rng = .Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
                rng.Name = "START_TIME_PERIOD_" & tabId
            End If
        End If


        'On time series and not on new sections, start with the previous end row

        If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And Not isNewSection Then
            If Not RangeExists("ENDTABLE_" & specs.Previous.TableId) Then ThrowExecutionError "I can't append the new time series to a previous table, No previous table  were found :("
            Set cellRng = .Range("ENDTABLE_" & specs.TableId)
            Set catRng = .Range(cellRng.Offset(1 - NBROWSTIMESERIES), cellRng.Offset(-3)) 'should be negative the offset (going backward)
        End If

        'All columns of the data
        counter = StartColumn
        nbcolCounter = 1
        Set rng = catRng.Offset(, 1)

        Do While (counter <= EndColumn) And (nbcolCounter <= nbCols)
            
            rng.Name = "VALUES_COL_" & nbcolCounter & "_" & tabId
            ALLRANGESNAMES.Push "VALUES_COL_" & nbcolCounter & "_" & tabId

            Select Case tablType

            Case TypeBivariate, TypeTimeSeries, TypeSpatial, TypeSpatioTemporal
                labCol = -1
            Case TypeUnivariate
                labCol = 0
            End Select

            'Name the header witch will serve as labels
            rng.Cells(labCol, 1).Name = "LABEL_COL_" & nbcolCounter & "_" & tabId
            ALLRANGESNAMES.Push "LABEL_COL_" & nbcolCounter & "_" & tabId
            
            'Movte to next Range
            Set rng = rng.Offset(, 1)
            
            DoEvents
            
            If specs.HasPercentage Then

                rng.Name = "PERC_COL_" & nbcolCounter & "_" & tabId
                ALLRANGESNAMES.Push "PERC_COL_" & nbcolCounter & "_" & tabId

                If tablType = TypeUnivariate Then 
                    rng.Cells(labCol, 1).Name = "PERC_LABEL_COL_" & tabId
                    ALLRANGESNAMES.Push "PERC_LABEL_COL_" & tabId
                End If

                Set rng = rng.Offset(, 1)
                counter = counter + 1
            End If

            counter = counter + 1
            nbcolCounter = nbcolCounter + 1
        Loop
        
        DoEvents
        Set cData = New BetterArray
        Set tData = New BetterArray

        cData.Push "MISSING", "TOTAL", "PERC"
        tData.Push "MSG_NA", "MSG_Total", "MSG_Percent"

        For counter = cData.LowerBound To cData.UpperBound
            'Missing/Total Row

            Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=True)

            If (Not rng Is Nothing) Then
                
                rng.Name = cData.Item(counter) & "_ROW_" & tabId
                ALLRANGESNAMES.Push cData.Item(counter) & "_ROW_" & tabId

                Set rng = RowRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=False)
                
                rng.Name = cData.Item(counter) & "_ROW_VALUES_" & tabId
                ALLRANGESNAMES.Push cData.Item(counter) & "_ROW_VALUES_" & tabId
            End If

            'Missing /Total Column
            If tablType <> TypeUnivariate Then
                Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=True, onlyCategories:=True)
                If (Not rng Is Nothing) Then
                    rng.Name = cData.Item(counter) & "_COL_" & tabId
                    ALLRANGESNAMES.Push cData.Item(counter) & "_COL_" & tabId
                    rng.Cells(0, 1).Name = cData.Item(counter) & "_LABEL_COL_" & tabId
                    Set rng = ColumnRange(trans.TranslatedValue(tData.Item(counter)), includeHeaders:=False, onlyCategories:=True)
                    rng.Name = cData.Item(counter) & "_COL_VALUES_" & tabId
                    ALLRANGESNAMES.Push cData.Item(counter) & "_COL_VALUES_" & tabId
                    'Percentages on Missing/Total column
                    If specs.HasPercentage Then
                        Set rng = rng.Offset(, 1)
                        rng.Name = cData.Item(counter) & "_PERC_VALUES_" & tabId
                        ALLRANGESNAMES.Push cData.Item(counter) & "_PERC_VALUES_" & tabId
                        'All the percentage columns (including label)
                        .Range(cData.Item(counter) & "_COL_" & tabId).Offset(, 1).Name = cData.Item(counter) & "_PERC_" & tabId
                    End If
                End If
            End If
        Next


        'Intersection of total and missings
        If RangeExists("TOTAL_COL_" & tabId) Then
            'Total and Total
            If RangeExists("TOTAL_ROW_" & tabId) Then
                'Total row
                rw = .Range("TOTAL_ROW_" & tabId).Row
                Col = .Range("TOTAL_COL_" & tabId).Column
                .Cells(rw, Col).Name = "TOTAL_TOTAL_" & tabId
                ALLRANGESNAMES.Push "TOTAL_TOTAL_" & tabId
            End If

            If RangeExists("MISSING_ROW_" & tabId) Then
                rw = .Range("MISSING_ROW_" & tabId).Row
                Col = .Range("TOTAL_COL_" & tabId).Column
                .Cells(rw, Col).Name = "MISSING_TOTAL_" & tabId
                ALLRANGESNAMES.Push "MISSING_TOTAL_" & tabId
                If RangeExists("MISSING_COL_" & tabId) Then
                    Col = .Range("MISSING_COL_" & tabId).Column
                    .Cells(rw, Col).Name = "MISSING_MISSING_" & tabId
                    ALLRANGESNAMES.Push "MISSING_MISSING_" & tabId
                End If
            End If
        End If

        If RangeExists("TOTAL_ROW_" & tabId) Then
            'Total and Total
            If RangeExists("MISSING_COL_" & tabId) Then
                Col = .Range("MISSING_COL_" & tabId).Column
                rw = .Range("TOTAL_ROW_" & tabId).Row
                .Cells(rw, Col).Name = "TOTAL_MISSING_" & tabId
                ALLRANGESNAMES.Push "TOTAL_MISSING_" & tabId
            End If
        End If

        'All the interior of the table
        addBegin = Split(.Range("VALUES_COL_1_" & tabId).Address, ":")(0)
        addEnd = IIf(specs.HasPercentage, "PERC_COL_", "VALUES_COL_")
        addEnd = Split(.Range(addEnd & nbCols & "_" & tabId).Address, ":")(1)
        .Range(addBegin & ":" & addEnd).Name = "INTERIOR_VALUES_" & tabId
        ALLRANGESNAMES.Push  "INTERIOR_VALUES_" & tabId

        'For spatial table, add a new name (outer_values for future computations)
        If tablType = TypeSpatial Then
            Set rng = .Range("INTERIOR_VALUES_" & tabId)
            Set rng = .Range(rng.Cells(1, -1), rng.Cells(rng.Rows.Count + 1, rng.Columns.Count + 4))
            rng.Name = "OUTER_VALUES_" & tabId
            ALLRANGESNAMES.Push  "OUTER_VALUES_" & tabId
        End If

        'Header Row with column categories
        Set rng = HeaderRange()
        rng.Name = "COLUMN_CATEGORIES_" & tabId
        ALLRANGESNAMES.Push "COLUMN_CATEGORIES_" & tabId
    End With
End Sub

'Format

Private Sub Format(ByVal desFormat As ILLFormat)
    Dim sh As Worksheet
    Dim specs As ITablesSpecs
    Dim rng As Range
    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String
    Dim tabId As String
    Dim DesignFormat As ILLFormat
    Dim wholeTableStart As Long
    Dim HasPercent As Boolean
    Dim NewSection As Boolean
    Dim tablType As AnalysisTablesType

    Set rngTable = New BetterArray
    Set specs = Specifications
    Set sh = Wksh()
    Set DesignFormat = desFormat

    tabId = specs.TableId
    tablType = specs.TableType
    HasPercent = specs.HasPercentage
    NewSection = specs.isNewSection

    'The formatting is divided in parts, from the begining to the end of the
    'table

    '- Sections
    '- Titles
    '- Headers
    '- Categories columns and Rows
    '- Interior Values
    '- uniques cells (missing_missing, total_total, missing_total, info_user, ect)
    '- Percentage columns (If they exists)
    '- Then the whole table
    '- There is a special format for the END_TABLE to take in accound gap for graphs

    With sh
        'Sections =============================================================
        If NewSection Then
            Set rng = .Range("SECTION_" & tabId)
            DesignFormat.ApplyFormat rng, AnalysisSection

        ElseIf RangeExists("ROWGS_SET") And (Not RangeExists("ROWGS_FORMATSET")) Then
            Set rng = .Range("ROWGS_SET")
            DesignFormat.ApplyFormat rng, scope:=AnalysisSection
            'Add informations on the sheet to avoid redoing the same thing
            rng.Cells(1, 8).Name = "ROWGS_FORMATSET"
        End If

        'Tittle (global summary and time series don't have a title)
        If (tablType <> TypeGlobalSummary) And (tablType <> TypeTimeSeries) And (tablType <> TypeSpatioTemporal) Then
            Set rng = .Range("TITLE_" & tabId)
            DesignFormat.ApplyFormat rng, AnalysisTableTitle
        End If
        
        'Dropdown on Geo/spatial table, and also buttons for the
        If tablType = TypeSpatial Then

            'Apply the dropdown only on geo
            If (specs.SpatialTableTypes = "geo") Then
                Set rng = .Range("ADM_DROPDOWN_" & tabId)
                DesignFormat.ApplyFormat rng, scope:=LinelistSelectionDropdown
                Set rng = .Range("DEVIDEPOP_" & tabId)
                DesignFormat.ApplyFormat rng, scope:=LinelistSelectionDropdown

            ElseIf (specs.SpatialTableTypes = "hf") Then
                Set rng = .Range("HF_DROPDOWN_" & tabId)
                DesignFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            End If
            
            'Range for the order (used in the formulas)
            Set rng = .Range("ROW_CATEGORIES_" & tabId).Offset(, -1)
            DesignFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            Set rng = rng.Offset(, -1)
            DesignFormat.ApplyFormat rng, scope:=LinelistHiddenCell
        End If

        'Headers ==============================================================

        Select Case tablType
        
        Case TypeGlobalSummary
            If Not RangeExists("COLGS_FORMATSET") Then
                Set rng = .Range("COLGS_SET")
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(1, 2))
                DesignFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
                rng.Cells(1, 3).Name = "COLGS_FORMATSET"
            End If
            
        Case TypeUnivariate
            Set rng = .Range("COLUMN_CATEGORIES_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
            Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
        
        Case TypeBivariate, TypeSpatial
            
            Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
            
            Set rng = .Range("STARTCOL_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelColumn

        Case TypeTimeSeries, TypeSpatioTemporal

            'Row categories are formatted only on new sections
            If NewSection Then
                Set rng = .Range("LABEL_ROW_CATEGORIES_" & tabId)
                DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesLabelRow
            End If
            
            'Headers with the start columns
            Set rng = .Range("STARTCOL_" & tabId)
            Set rng = .Range(.Cells(rng.Row, rng.Column), .Cells(rng.Row, EndColumn()))
            DesignFormat.ApplyFormat rng, scope:=AnalysisTimeSeriesHeader

        End Select

        'Categories column and row =============================================

        Select Case tablType

        Case TypeGlobalSummary
            Set rng = .Range("STARTROW_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisGlobalSummaryRow
        
        Case TypeUnivariate
            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesUnivariateRow
        
        Case TypeBivariate, TypeTimeSeries, TypeSpatial, TypeSpatioTemporal
            'rows
            If ((tablType <> TypeTimeSeries) And (tablType <> TypeSpatioTemporal)) Or NewSection Then
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow
                'start and end date columns
                If (tablType = TypeTimeSeries) Or (tablType = TypeSpatioTemporal) Then
                    Set rng = rng.Offset(, -1)
                    Set rng = .Range(rng.Cells(1, 0), rng.Cells(rng.Rows.Count + 1, 1))
                    DesignFormat.ApplyFormat rng, scope:=LinelistHiddenCell
                    rngTable.Clear
                    rngTable.Push "START_DATE_", "END_DATE_", "VALIDATION_MIN_DATE_", _
                                  "VALIDATION_MAX_DATE_", "MIN_MIN_DATE_", "MAX_MAX_DATE_"
                    For counter = rngTable.LowerBound To rngTable.UpperBound
                        rngName = rngTable.Item(counter) & tabId
                        Set rng = .Range(rngName)
                        DesignFormat.ApplyFormat rng, scope:=LinelistHiddenCell
                        DesignFormat.ApplyFormat rng, scope:=AnalysisDateCell
                    Next
                End If
            End If

            'For spatio-temporal, there is another row column for selected values
            'by the user, they will be formated as Rows
            If RangeExists("INPUTSPTGEO_USER_" & tabId) Then
                Set rng = .Range("INPUTSPTGEO_USER_" & tabId)
                DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateRow
            End If

            'columns
            Set rng = .Range("LABEL_COL_1_" & tabId)
            'Merge the columns if there is percentage
            If Not RangeExists("MERGED_" & tabId) Then
                Do While rng.Column <= EndColumn()
                    If HasPercent Then
                        Set rng = rng.Offset(, 2)
                        .Range(rng.Cells(1, -1), rng.Cells(1, 0)).Merge
                    Else
                        Set rng = rng.Offset(, 1)
                    End If
                Loop
                rng.Cells(1, 3).Name = "MERGED_" & tabId
            End If

            Set rng = .Range("COLUMN_CATEGORIES_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisCategoriesBivariateColumn
            Set rng = rng.Offset(1)
            DesignFormat.ApplyFormat rng, scope:=AnalysisFirstRow


        End Select
        
        'Interior values =======================================================

        If (tablType <> TypeGlobalSummary) Then
            Set rng = .Range("INTERIOR_VALUES_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisInteriorTable
        End If

        'Unique Cells ==========================================================

        'Total And missing lines
        rngTable.Clear
        rngTable.Push "TOTAL_ROW_VALUES_", "TOTAL_COL_VALUES_", _
                      "MISSING_ROW_VALUES_", "MISSING_COL_VALUES_", _
                      "TOTAL_PERC_VALUES_", "MISSING_PERC_VALUES_"
        
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If (InStr(1, rngName, "ROW") > 0) And RangeExists(rngName) Then
                Set rng = .Range(rngName)
                'On time series will shift the first cell on non new sections
                If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And (Not NewSection) Then
                    DesignFormat.ApplyFormat rng, scope:=AnalysisMissingRowWithoutHeader
                Else
                    DesignFormat.ApplyFormat rng, scope:=AnalysisMissingRow
                End If
                
            ElseIf RangeExists(rngName) Then
                Set rng = .Range(rngName)
                DesignFormat.ApplyFormat rng, scope:=AnalysisMissingColumn
            End If
        Next

        'Single cells (total_missing, missing_total, ect.)
        rngTable.Clear
        rngTable.Push "TOTAL_MISSING_", "TOTAL_TOTAL_", _
                      "USER_START_DATE_", "USER_END_DATE_", "TIME_UNIT_", _
                      "MISSING_TOTAL_", "MISSING_MISSING_", "START_DATE_LABEL_", _
                      "TIME_UNIT_LABEL_", "END_DATE_LABEL_", "POPFACT_", "POPFACTLABEL_", _
                      "INPUTSPTGEO_LABEL_"
        
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId

            'labels are formatted differently (using analysis univariate header format)
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                If (InStr(1, rngName, "LABEL") > 0) Or (InStr(1, rngName, "POP") > 0) Then
                    DesignFormat.ApplyFormat rng, scope:=AnalysisUnivariateHeader
                ElseIf (InStr(1, rngName, "USER") > 0) Or (rngName = "TIME_UNIT_" & tabId) Then
                    DesignFormat.ApplyFormat rng, scope:=AnalysisOneCellCenter
                Else
                    DesignFormat.ApplyFormat rng, scope:=AnalysisOneCell 'Will be formatted with the missing color
                End If
            End If
        Next

        'Informations Cells (Info on start date, info on end date)
        rngTable.Clear
        rngTable.Push "INFO_START_DATE_", "INFO_END_DATE_", "INFO_ANA_PERIOD_"
        
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                DesignFormat.ApplyFormat rng, scope:=AnalysisInfoCell
                If InStr(1, rngName, "INFO_ANA_") > 0 Then DesignFormat.ApplyFormat rng, scope:=AnalysisInfoCellLeft
            End If
        Next

        'Hidden cells (no matter what)
        rngTable.Clear
        rngTable.Push "PREVIOUS_ADM_", "POPPREVFACT_", "POPFACT_", "POPFACTLABEL_", _ 
                      "INPUTSPTGEO_NGEO_", "INPUTPREVADM_USER_"
        
        For counter = rngTable.LowerBound To rngTable.UpperBound
            rngName = rngTable.Item(counter) & tabId
            If RangeExists(rngName) Then
                Set rng = .Range(rngName)
                DesignFormat.ApplyFormat rng, scope:=LinelistHiddenCell
            End If
        Next

        'Percentages columns ========================================================

        If HasPercent Then
            Set rng = .Range("PERC_COL_1_" & tabId)
            If (specs.TableType <> TypeSpatial) Or (specs.HasTotal) Then
                Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1))
            End If

            Do While rng.Column <= EndColumn()
                DesignFormat.ApplyFormat rng, scope:=AnalysisPercent
                Set rng = rng.Offset(, 2)
            Loop
            rngTable.Clear
            rngTable.Push "MISSING_MISSING_", "TOTAL_TOTAL_", "MISSING_TOTAL_"
            
            For counter = rngTable.LowerBound To rngTable.UpperBound
                rngName = rngTable.Items(counter) & tabId
                If RangeExists(rngName) Then
                    Set rng = .Range(rngName).Offset(, 1)
                    DesignFormat.ApplyFormat rng, scope:=AnalysisOneCellCenter 'Percentages are centered
                End If
            Next
        End If
        
        'The whole table =======================================================
        wholeTableStart = StartColumn
        If (tablType = TypeTimeSeries Or tablType = TypeSpatioTemporal) And (Not NewSection) Then
            wholeTableStart = StartColumn + 1
        End If

        If (tablType <> TypeUnivariate) And (tablType <> TypeGlobalSummary) Then
            Set rng = .Range(.Cells(StartRow, wholeTableStart), .Cells(EndRow, EndColumn))
            DesignFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        ElseIf (tablType = TypeUnivariate) Then
            Set rng = .Range(.Cells(StartRow + 1, wholeTableStart), .Cells(EndRow, EndColumn))
            DesignFormat.ApplyFormat rng, scope:=AnalysisWholeTable
        End If

        'The end table for gap for graphs ======================================

        If RangeExists("ENDTABLE_" & tabId) And _ 
          (tablType <> TypeTimeSeries) And _ 
          (tablType <> TypeSpatioTemporal) Then
            Set rng = .Range("ENDTABLE_" & tabId)
            DesignFormat.ApplyFormat rng, scope:=AnalysisEndTable
        End If
    End With
End Sub

'Build the table
Private Sub Build()
    'Add the headers of the table before everything because of shifts of rows
    AddHeader
    AddRows                                      'Add categories for rows and column
    AddColumns
    NameRanges                                   'Name required ranges of the table (for plotting purposes
End Sub

'Get the formula of the table

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowExecutionError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error : " & errorMessage
End Sub

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

Private Sub ICrossTable_Format(ByVal DesignFormat As ILLFormat)
    Format DesignFormat
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

Private Property Get ICrossTable_RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, Optional ByVal onlyCategories As Boolean = False) As Range
    Set ICrossTable_RowsCategoriesRange = RowsCategoriesRange(includeHeaders, onlyCategories)
End Property

Private Property Get ICrossTable_ColumnRange(ByVal colVal As String, _
                                            Optional ByVal onlyCategories As Boolean = False, _
                                            Optional ByVal includeHeaders As Boolean = False) As Range
    Set ICrossTable_ColumnRange = ColumnRange(colVal, includeHeaders, onlyCategories)
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

Private Property Get ICrossTable_Specifications() As ITablesSpecs
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