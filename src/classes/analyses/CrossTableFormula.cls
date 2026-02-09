VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Formulas for all the cross tables in the analysis"

'@PredeclaredId
'@Folder("Analysis")
'@ModuleDescription("Formulas for all the cross tables in the analysis")
'@IgnoreModule UnrecognizedAnnotation, IIfSideEffect, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ICrossTableFormula

Private Type TCrossTableFormula
    tabl As ICrossTable
    fData As IFormulaData
    forml As IFormulas
    condVars As BetterArray
    condTests As BetterArray
End Type

Private this As TCrossTableFormula
Private Const CLASS_NAME As String = "CrossTableFormula"

'@section Factory
'===============================================================================

' @description Create a cross-table formula writer
' @param tabl Cross-table to write formulas into
' @param fData Formula data for function lookups
Public Function Create(ByVal tabl As ICrossTable, _
                       ByVal fData As IFormulaData) As ICrossTableFormula

    If tabl Is Nothing Then ThrowError InvalidArgument, "Cross-table not defined"
    If fData Is Nothing Then ThrowError InvalidArgument, "Formula data not defined"

    With New CrossTableFormula
        Set .Table = tabl
        Set .formData = fData
        Set Create = .Self
    End With
End Function

'@section Internal properties (factory use)
'===============================================================================

Public Property Get Self() As ICrossTableFormula
    Set Self = Me
End Property

Public Property Get Table() As ICrossTable
    Set Table = this.tabl
End Property

Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tabl = tabl
End Property

Public Property Get formData() As IFormulaData
    Set formData = this.fData
End Property

Public Property Set formData(ByVal fData As IFormulaData)
    Set this.fData = fData
End Property

'@section Shortcut properties
'===============================================================================

Private Property Get Specifications() As ITableSpecs
    Set Specifications = Table.Specifications
End Property

Private Property Get SummaryFunction() As String
    SummaryFunction = Specifications.Value("summary function")
End Property

Private Property Get ID() As String
    ID = Specifications.TableId()
End Property

Private Property Get SectionId() As String
    SectionId = Specifications.TableSectionId
End Property

Private Property Get TableType() As Byte
    TableType = Specifications.TableType()
End Property

Private Property Get TableFormula() As IFormulas
    If this.forml Is Nothing Then
        Dim specs As ITableSpecs
        Set specs = Specifications
        Set this.forml = Formulas.Create(specs.Dictionary, formData(), SummaryFunction)
    End If
    Set TableFormula = this.forml
End Property

Private Property Get Valid() As Boolean
    Valid = TableFormula.Valid(formulaType:="analysis")
End Property

'@section Range and formula helpers
'===============================================================================

Private Function RangeExists(ByVal rngName As String) As Boolean
    Dim sh As Worksheet
    Dim rng As Range
    Set sh = Table.Wksh()
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim prevAlert As Long
    Dim sh As Worksheet

    prevAlert = Application.DisplayAlerts
    Set sh = Table.Wksh

    With sh
        Application.DisplayAlerts = False
        On Error GoTo Fail
        .Cells(1, 1).Clear
        .Cells(1, 1).FormulaArray = "= " & frmName
        .Cells(1, 1).Clear
        FailedFormula = False
        Application.DisplayAlerts = prevAlert
        Exit Function
    End With

Fail:
    FailedFormula = True
    Application.DisplayAlerts = prevAlert
End Function

Private Function EmpStr() As String
    EmpStr = Chr(34) & Chr(34)
End Function

Private Function NonEmpStr() As String
    NonEmpStr = Chr(34) & "<>" & Chr(34)
End Function

Private Function Percentage(ByVal denomRng As Range, ByVal numRng As Range, _
                            Optional ByVal rowabs As Boolean = True) As String
    Dim parsedForm As String
    parsedForm = numRng.Address(RowAbsolute:=False) & "/" & _
                 denomRng.Address(RowAbsolute:=rowabs, ColumnAbsolute:=True)
    Percentage = "IF(ISERR(" & parsedForm & ")," & EmpStr & "," & parsedForm & ")"
End Function

Private Function Condition(ByVal rngCondition As Range, ByVal formulaValue As String, _
                           Optional ByVal rowabs As Boolean = False) As String
    Condition = "IF(" & rngCondition.Address(RowAbsolute:=rowabs) & "=" & _
                EmpStr & "," & EmpStr & "," & formulaValue & ")"
End Function

'@section Condition management
'===============================================================================

Private Sub AddTableConditions(ByVal varName As String, ByVal conditionName As String)
    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
    End If

    If varName <> vbNullString Then
        this.condVars.Push varName
        this.condTests.Push conditionName
    End If
End Sub

Private Sub ClearTableConditions(Optional ByVal removeNumber As Long = 0)
    Dim counter As Long
    Dim remNum As Long

    remNum = removeNumber

    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
        remNum = 0
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
        remNum = 0
    End If

    If remNum <= 0 Then
        this.condVars.Clear
        this.condTests.Clear
    Else
        For counter = 1 To remNum
            this.condVars.Pop
            this.condTests.Pop
        Next
    End If
End Sub

Private Sub AddNonEmptyTableConditions(ByVal varName As String)
    If SummaryFunction = "N" Or SummaryFunction = "N()" Then
        AddTableConditions varName, "= " & NonEmpStr()
    Else
        AddTableConditions varName, "<>" & EmpStr()
    End If
End Sub

'@section Excel formula builder
'===============================================================================

Private Property Get CurrentExcelFormula(Optional ByVal scope As Byte = 2, _
                                         Optional ByVal prefix As String = "f") As String
    Dim cond As IFormulaCondition
    Dim tabFormObject As IFormulas
    Dim eform As String

    Set tabFormObject = TableFormula

    If scope = FormulaScopeLinelist Then
        eform = tabFormObject.ParsedLinelistFormula(useTableName:=True, tablePrefix:=prefix)
    Else
        If this.condVars Is Nothing Or this.condTests Is Nothing Then GoTo Fail
        Set cond = FormulaCondition.Create(this.condVars, this.condTests)
        eform = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:=prefix)
    End If

    If FailedFormula(eform) Then GoTo Fail

    CurrentExcelFormula = eform
    Exit Property

Fail:
    eform = Chr(34) & "##== formula parsing failed :( ==##" & Chr(34)
    CurrentExcelFormula = eform
End Property

'@section Number format helper
'===============================================================================

Private Sub AddRngFormat(ByVal formatValue As String, ByVal rng As Range)
    Dim formatText As String
    Dim nbDecimals As Long
    Dim zerosDecimals As String

    nbDecimals = 1

    If formatValue = vbNullString Then
        Exit Sub
    ElseIf formatValue = "integer" Then
        formatText = "###0;-###0;;@"
    ElseIf formatValue = "round0" Then
        formatText = "#,##0;-#,##0;0;@"
    ElseIf formatValue = "percentage0" Then
        formatText = "0%"
    ElseIf InStr(1, formatValue, "round") > 0 Then
        On Error Resume Next
        nbDecimals = CLng(Replace(formatValue, "round", vbNullString))
        On Error GoTo 0
        zerosDecimals = String(nbDecimals, "0")
        formatText = "#,###0." & zerosDecimals & ";-#,###0." & zerosDecimals & ";0." & zerosDecimals & ";@"
    ElseIf InStr(1, formatValue, "percentage") > 0 Then
        On Error Resume Next
        nbDecimals = CLng(Replace(formatValue, "percentage", vbNullString))
        On Error GoTo 0
        zerosDecimals = String(nbDecimals, "0")
        formatText = "0." & zerosDecimals & "%"
    ElseIf formatValue = "euros" Then
        formatText = "#,##0.00 $;-#,##0.00 $;0.00 $;@"
    ElseIf formatValue = "dollars" Then
        formatText = "[$$-en-US] #,##0.00;[$$-en-US] -#,##0.00;[$$-en-US] 0.00;@"
    ElseIf formatValue = "text" Then
        formatText = "@"
    Else
        formatText = "custom"
    End If

    If formatText <> "custom" Then
        rng.NumberFormat = formatText
    Else
        On Error Resume Next
        rng.NumberFormat = formatValue
        On Error GoTo 0
    End If
End Sub

'@section AddFormulas — main formula writing logic
'===============================================================================

Private Sub AddFormulas()
    Dim tabId As String
    Dim tabType As Byte
    Dim sh As Worksheet
    Dim rng As Range
    Dim totRng As Range
    Dim rngLab As Range
    Dim misRng As Range
    Dim cellRng As Range
    Dim finalFormula As String
    Dim tabEndColumn As Long
    Dim secId As String
    Dim perRng As Range
    Dim rowVar As String
    Dim colVar As String
    Dim tabFormat As String
    Dim rowCatAddress As String
    Dim startTimeAddress As String
    Dim endTimeAddress As String
    Dim hasMissing As Boolean
    Dim hasTotal As Boolean
    Dim hasPercentage As Boolean
    Dim tabl As ICrossTable
    Dim spTab As ISpatialTables
    Dim tabFormObject As IFormulas
    Dim spTabType As String
    Dim counter As Long
    Dim specs As ITableSpecs

    If Not Valid() Then Exit Sub

    Set specs = Specifications
    Set tabl = Table()
    Set sh = tabl.Wksh()

    tabId = ID()
    secId = SectionId()
    tabType = TableType
    hasMissing = specs.HasMissing
    hasTotal = specs.HasTotal
    hasPercentage = specs.HasPercentage
    tabFormat = specs.Value("format")

    With sh

        Select Case tabType

        Case TypeGlobalSummary
            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:=vbNullString)
            Set rng = .Range("STARTROW_" & tabId)
            rng.Cells(1, 2).FormulaArray = "= " & finalFormula
            AddRngFormat tabFormat, rng.Cells(1, 2)

            finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist, prefix:="f")
            rng.Cells(1, 3).FormulaArray = "= " & finalFormula
            AddRngFormat tabFormat, rng.Cells(1, 3)

        Case TypeUnivariate

            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            rowVar = specs.Value("row")
            ClearTableConditions
            AddTableConditions rowVar, "=" & rng.Cells(1, 1).Address(RowAbsolute:=False)
            finalFormula = CurrentExcelFormula()

            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set cellRng = rng.Cells(1, 1)
            cellRng.FormulaArray = "= " & finalFormula
            cellRng.AutoFill Destination:=rng, Type:=xlFillValues
            AddRngFormat tabFormat, rng

            ' Missing and total
            If hasMissing Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                Set rng = .Range("MISSING_ROW_VALUES_" & tabId)
                rng.Cells(1, 1).FormulaArray = "=" & finalFormula

                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & NonEmpStr()
                    finalFormula = finalFormula & " + " & CurrentExcelFormula()
                    rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                Else
                    ClearTableConditions
                    finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist)
                    rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                End If
            Else
                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)
                ClearTableConditions
                AddNonEmptyTableConditions rowVar
                finalFormula = CurrentExcelFormula()
                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
            End If

            ' Percentage
            If hasPercentage Then
                Set rng = .Range("VALUES_COL_1_" & tabId)
                Set cellRng = .Range("TOTAL_ROW_VALUES_" & tabId).Cells(1, 1)
                rng.Offset(, 1).Cells(1, 1).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), True)
                rng.Offset(, 1).Cells(1, 1).AutoFill Destination:=rng.Offset(, 1), Type:=xlFillValues
                cellRng.Cells(1, 2).Formula = "= " & Percentage(cellRng.Cells(1, 1), cellRng.Cells(1, 1))
            End If

        Case TypeBivariate

            rowVar = specs.Value("row")
            colVar = specs.Value("column")
            rowCatAddress = .Range("ROW_CATEGORIES_" & tabId).Cells(1, 1).Address(RowAbsolute:=False)
            tabEndColumn = tabl.EndColumn

            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)

            Do While rng.Column < tabEndColumn

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "=" & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                If specs.Value("missing") = "row" Or specs.Value("missing") = "all" Then
                    Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & EmpStr()
                    AddTableConditions colVar, "= " & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(misRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula

                    ClearTableConditions
                    AddTableConditions colVar, "= " & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula
                Else
                    ClearTableConditions
                    AddNonEmptyTableConditions rowVar
                    AddTableConditions colVar, "=" & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula
                End If

                ' Percentage
                If hasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill _
                        Destination:=.Range(.Cells(rng.Row, rng.Column), .Cells(totRng.Row, rng.Column)), _
                        Type:=xlFillValues
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
            Loop

            ' Missing column and total columns
            If specs.Value("missing") = "column" Or specs.Value("missing") = "all" Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                misRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                finalFormula = CurrentExcelFormula()
            Else
                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddNonEmptyTableConditions colVar
                finalFormula = CurrentExcelFormula()
            End If

            Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
            totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
            totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
            AddRngFormat tabFormat, totRng

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Total row, missing column
            If RangeExists("TOTAL_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("TOTAL_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("TOTAL_MISSING_" & tabId)
            End If

            ' Total row, total column
            If RangeExists("TOTAL_TOTAL_" & tabId) Then
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    .Range("TOTAL_TOTAL_" & tabId).Formula = _
                        "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
                Else
                    Select Case specs.Value("missing")
                    Case "all"
                        finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist)
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                    Case "row"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Case "column"
                        ClearTableConditions
                        AddNonEmptyTableConditions rowVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Case "no"
                        ClearTableConditions
                        AddNonEmptyTableConditions colVar
                        AddNonEmptyTableConditions rowVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    End Select
                End If
                AddRngFormat tabFormat, .Range("TOTAL_TOTAL_" & tabId)
            End If

        Case TypeTimeSeries, TypeSpatioTemporal

            rowVar = specs.Value("time")
            colVar = specs.Value("column")
            counter = 1

            ' For spatio-temporal, column variable depends on geo/hf type
            If tabType = TypeSpatioTemporal Then
                spTabType = specs.SpatialTableTypes()
                colVar = Switch(spTabType = "geo", "concat_adm1_" & colVar, _
                                spTabType = "hf", "hf_" & colVar, _
                                True, vbNullString)
            End If

            ' Time period formulas (only for new sections)
            If specs.IsNewSection() Then
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FormatDateFromLastDay(TIME_UNIT_" & tabId & "," & _
                                  cellRng.Cells(1, -1).Address(RowAbsolute:=False) & "," & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & ", " & _
                                  "END_DATE_" & tabId & ")"
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' Start period column
                Set cellRng = .Range("FIRST_VALUE_START_TIME_" & tabId)
                cellRng.Formula = "= START_DATE_" & tabId
                cellRng.Cells(2, 1).Formula = "= " & _
                    cellRng.Cells(1, 2).Address(RowAbsolute:=False) & " + 1"
                Set rng = .Range("START_TIME_PERIOD_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' End period column
                Set rng = .Range("END_TIME_PERIOD_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindLastDay(TIME_UNIT_" & tabId & "," & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & ")"
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' Total period
                Set rng = .Range("TOTAL_ROW_" & tabId)
                rng.Cells(1, -1).Formula = "= START_DATE_" & tabId
                rng.Cells(1, 0).Formula = "= END_DATE_" & tabId

                ' Validation max/min
                Set rng = .Range("VALIDATION_MAX_DATE_" & tabId)
                rng.Formula = "=ValidMax(USER_START_DATE_" & tabId & _
                              ", USER_END_DATE_" & tabId & ", MIN_MIN_DATE_" & tabId & _
                              ", MAX_MAX_DATE_" & tabId & ", TIME_UNIT_" & tabId & ")"

                Set rng = .Range("VALIDATION_MIN_DATE_" & tabId)
                rng.Formula = "=ValidMin(USER_START_DATE_" & tabId & _
                              ", USER_END_DATE_" & tabId & ", MIN_MIN_DATE_" & tabId & _
                              ", MAX_MAX_DATE_" & tabId & ", TIME_UNIT_" & tabId & ")"

                ' Start/End dates
                Set rng = .Range("START_DATE_" & tabId)
                rng.Formula = "= MAX(USER_START_DATE_" & tabId & _
                              ", VALIDATION_MIN_DATE_" & tabId & ")"

                Set rng = .Range("END_DATE_" & tabId)
                rng.Formula = "= MIN(USER_END_DATE_" & tabId & _
                              ", VALIDATION_MAX_DATE_" & tabId & ")"

                ' Info cells
                Set rng = .Range("INFO_START_DATE_" & tabId)
                rng.Formula = "= InfoUser(USER_START_DATE_" & tabId & _
                              ", START_DATE_" & tabId & ")"

                Set rng = .Range("INFO_END_DATE_" & tabId)
                rng.Formula = "= InfoUser(USER_END_DATE_" & tabId & _
                              ", END_DATE_" & tabId & ", 2)"

                Set rng = .Range("INFO_ANA_PERIOD_" & tabId)
                rng.Formula = "= RNG_AnaPeriod & " & _
                              " PLAGE_VALUE(START_DATE_" & tabId & "," & _
                              "END_DATE_" & tabId & ")"

                ' Time unit validation
                On Error Resume Next
                With .Range("TIME_UNIT_" & tabId).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:="=TIME_UNIT_LIST"
                End With
                On Error GoTo 0
            End If

            ' Column formulas
            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1))
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)
            Set cellRng = .Range("FIRST_VALUE_START_TIME_" & secId)
            Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
            Set perRng = cellRng.Cells(1, 3)
            startTimeAddress = cellRng.Address(RowAbsolute:=False)
            endTimeAddress = cellRng.Cells(1, 2).Address(RowAbsolute:=False)
            tabEndColumn = tabl.EndColumn

            Do While rng.Column <= tabEndColumn

                ClearTableConditions
                AddTableConditions rowVar, ">=" & startTimeAddress
                AddTableConditions rowVar, "<=" & endTimeAddress
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()
                Set cellRng = .Cells(misRng.Row, rng.Column)
                cellRng.FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, cellRng

                ' Percentage
                If hasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                End If

                ' Spatio-temporal label formulas
                If tabType = TypeSpatioTemporal Then
                    If spTabType = "geo" Then
                        rngLab.Formula = "= IF(ISBLANK(INPUTSPTGEO_" & counter & "_" & secId & ")" & _
                                         ", " & Chr(34) & Chr(34) & ", " & _
                                         "INPUTSPTGEO_" & counter & "_" & secId & ")"
                    Else
                        rngLab.Formula = "= IF(ISBLANK(INPUTSPTHF_" & counter & "_" & secId & ")" & _
                                         ", " & Chr(34) & Chr(34) & ", " & _
                                         "INPUTSPTHF_" & counter & "_" & secId & ")"
                    End If
                    counter = counter + 1
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
            Loop

            ' Missing column and total columns
            If hasMissing Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
                Set misRng = .Range(misRng.Cells(1, 1), misRng.Cells(misRng.Rows.Count + 1, 1))

                ClearTableConditions
                AddTableConditions rowVar, ">=" & startTimeAddress
                AddTableConditions rowVar, "<=" & endTimeAddress
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()

                misRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            Else
                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    If colVar <> vbNullString Then AddNonEmptyTableConditions colVar
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            End If

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Total row, total column
            If RangeExists("TOTAL_TOTAL_" & tabId) Then
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    .Range("TOTAL_TOTAL_" & tabId).Formula = _
                        "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
                Else
                    ClearTableConditions
                    Set cellRng = .Range("TOTAL_ROW_" & secId).Cells(1, 0)
                    AddTableConditions rowVar, "<=" & cellRng.Address
                    AddTableConditions rowVar, ">=" & cellRng.Cells(1, 0).Address

                    If hasMissing Then
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Else
                        AddNonEmptyTableConditions colVar
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    End If
                End If
                AddRngFormat tabFormat, .Range("TOTAL_TOTAL_" & tabId)
            End If

        Case TypeSpatial

            spTabType = specs.SpatialTableTypes()
            rowVar = specs.Value("row")
            rowVar = Switch(spTabType = "geo", "concat_adm1_" & rowVar, _
                            spTabType = "hf", "hf_" & rowVar, _
                            True, vbNullString)
            colVar = specs.Value("column")

            ' Validation dropdowns
            On Error Resume Next
            If spTabType = "geo" Then
                With .Range("ADM_DROPDOWN_" & tabId).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:="=ADM_UNIT_LIST"
                End With

                With .Range("DEVIDEPOP_" & tabId).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:="=POPULATION_FACTOR_LIST"
                End With
            End If
            On Error GoTo 0

            ' Order formula
            Set rng = .Range("ROW_CATEGORIES_" & tabId).Offset(, -2)
            Set cellRng = rng.Cells(1, 1)
            cellRng.Value = 1
            Set cellRng = rng.Cells(2, 1)
            cellRng.Formula = "= " & cellRng.Offset(-1).Address(RowAbsolute:=False) & " + 1"
            Set rng = .Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
            cellRng.AutoFill Destination:=rng, Type:=xlFillValues

            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            rowCatAddress = rng.Cells(1, 1).Address(RowAbsolute:=False)

            If spTabType = "geo" Then
                ' Population formula
                Set rng = .Range("ROW_CATEGORIES_" & tabId).Offset(, -1)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindTopPop(ADM_DROPDOWN_" & tabId & "," & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & "," & _
                                  Chr(34) & rowVar & Chr(34) & "," & Chr(34) & tabId & Chr(34) & ")"
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues

                ' Top Admin formula
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindTopAdmin(ADM_DROPDOWN_" & tabId & "," & _
                                  cellRng.Cells(1, -1).Address(RowAbsolute:=False) & "," & _
                                  Chr(34) & rowVar & Chr(34) & "," & Chr(34) & tabId & Chr(34) & ")"

            ElseIf spTabType = "hf" Then
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.Formula = "= FindTopHF(" & _
                                  cellRng.Cells(1, 0).Address(RowAbsolute:=False) & "," & _
                                  Chr(34) & rowVar & Chr(34) & "," & Chr(34) & tabId & Chr(34) & ")"
            End If

            cellRng.AutoFill Destination:=rng, Type:=xlFillValues

            ' Column value formulas (similar to bivariate)
            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
            tabEndColumn = tabl.EndColumn

            Do While rng.Column <= tabEndColumn

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddNonEmptyTableConditions rowVar
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()

                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                AddRngFormat tabFormat, rng

                ' Missing row
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()
                Set cellRng = .Cells(misRng.Row, rng.Column)
                cellRng.FormulaArray = "= " & finalFormula

                ' Percentage (spatial only supports row percentage)
                If hasPercentage Then
                    Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                    rng.Cells(1, 2).Formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Set rng = rng.Offset(, 1)
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
            Loop

            ' Missing column and total columns
            If hasMissing Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()

                misRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                AddRngFormat tabFormat, misRng

                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & rowCatAddress
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            Else
                If hasTotal Then
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & rowCatAddress
                    AddNonEmptyTableConditions colVar
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                    AddRngFormat tabFormat, totRng
                End If
            End If

            ' Missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_TOTAL_" & tabId)
            End If

            ' Missing row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
                AddRngFormat tabFormat, .Range("MISSING_MISSING_" & tabId)
            End If

            ' Create spatial table management
            Set spTab = SpatialTables.Create(tabl)
            Set tabFormObject = TableFormula
            spTab.Add tabFormObject

        End Select
    End With
End Sub

'@section Error handling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface implementation
'===============================================================================

Private Property Get ICrossTableFormula_Valid() As Boolean
    ICrossTableFormula_Valid = Valid
End Property

Private Sub ICrossTableFormula_AddFormulas()
    AddFormulas
End Sub
