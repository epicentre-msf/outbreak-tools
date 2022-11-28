VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Analysis")
'@ModuleDescription("Formulas for all the cross tables in the analysis")
'@IgnoreModule

Option Explicit
Implements ICrossTableFormula

Private Enum ExcelFormulaScope
    FormulaScopeLinelist = 1
    FormulaScopeAnalysis = 2
End Enum

Private Type TCrossTableFormula
    tab As ICrossTable
    fData As IFormulaData
    forml As IFormulas
    condVars As BetterArray
    condTests As BetterArray
End Type

Private this As TCrossTableFormula
Private Const CLASSNAME As String = "CrossTableFormula"

Public Function Create(ByVal tabl As ICrossTable, fData As IFormulaData) As ICrossTableFormula
    With New CrossTableFormula
        Set .Table = tabl
        Set .formData = fData
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ICrossTableFormula
    Set Self = Me
End Property

Public Property Get Table() As ICrossTable
    Set Table = this.tab
End Property

Public Property Set Table(ByVal tabl As ICrossTable)
    Set this.tab = tabl
End Property

Public Property Get formData() As IFormulaData
    Set formData = this.fData
End Property

Public Property Set formData(ByVal fData As IFormulaData)
    Set this.fData = fData
End Property

Private Property Get Specifications() As ITablesSpecs
    Set Specifications = Table.Specifications
End Property

Private Property Get SummaryFunction() As String
    Dim specs As ITablesSpecs

    Set specs = Specifications
    SummaryFunction = specs.Value("summary function")

End Property

Private Property Get Id() As String
    Dim specs As ITablesSpecs
    Set specs = Specifications
    Id = specs.TableId()
End Property

Private Property Get sectionId() As String
    sectionId = Specifications.TableSectionId
End Property

Private Property Get TableType() As AnalysisTablesType
    Dim specs As ITablesSpecs
    Set specs = Specifications
    TableType = specs.TableType()
End Property

Private Property Get TableFormula() As IFormulas
    Dim tabform As String
    Dim fData As IFormulaData
    Dim specs As ITablesSpecs
    Dim dict As ILLdictionary
    Dim tabFormObject As IFormulas

    If (this.forml Is Nothing) Then
        Set fData = formData()
        Set specs = Specifications
        Set dict = specs.Dictionary
        tabform = SummaryFunction
        Set tabFormObject = Formulas.Create(dict, fData, tabform)

        Set this.forml = tabFormObject
        Set TableFormula = tabFormObject
    Else
        Set TableFormula = this.forml
    End If
End Property

Private Property Get valid() As Boolean
    Dim tabFormObject As IFormulas
    Set tabFormObject = TableFormula
    valid = tabFormObject.valid(formulaType:="analysis")
End Property

Private Function RangeExists(ByVal rngName As String) As Boolean
    'Test if range exists in the sheet
    Dim sh As Worksheet
    Dim rng As Range
    Set sh = Table.Wksh()
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0
    RangeExists = (Not rng Is Nothing)
End Function

Private Function FailedFormula(ByVal frmName As String) As Boolean
    Dim sh As Worksheet
    Dim formTest As String

    Set sh = Table.Wksh
    With sh
        .Cells(1, 1).Clear
        On Error GoTo Fail
        
        'Test the formula on one cell
        .Cells(1, 1).FormulaArray = "= " & frmName
        
        .Cells(1, 1).Clear
        FailedFormula = False
        Exit Function
    End With
    
Fail:
    FailedFormula = True
End Function

Private Function EmpStr() As String
    EmpStr = Chr(34) & Chr(34)
End Function

Private Function NonEmpStr() As String
    NonEmpStr = Chr(34) & "<>" & Chr(34)
End Function

Private Function Percentage(ByVal denomRng As Range, ByVal numRng As Range, Optional ByVal rowabs As Boolean = True) As String
    Dim parsedForm As String
    parsedForm = numRng.Address(Rowabsolute:=False) & "/" & denomRng.Address(Rowabsolute:=rowabs, ColumnAbsolute:=True)
    Percentage = "IF(ISERR(" & parsedForm & ")," & EmpStr & "," & parsedForm & ")"
End Function

'Write a formula on a cell if only another cell is not empty
Private Function Condition(ByVal rngCondition As Range, ByVal formulaValue As String, Optional ByVal rowabs As Boolean = False) As String
    Condition = "IF(" & rngCondition.Address(Rowabsolute:=rowabs) & "=" & EmpStr & "," & EmpStr & "," & formulaValue & ")"
End Function

'Add conditions for the formulas
Private Sub AddTableConditions(varName As String, conditionName As String)

    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
    End If

    this.condVars.Push varName
    this.condTests.Push conditionName
End Sub

Private Sub ClearTableConditions(Optional ByVal removeNumber As Integer = 0)
    Dim counter As Integer

    If this.condVars Is Nothing Then
        Set this.condVars = New BetterArray
        this.condVars.LowerBound = 1
        removeNumber = 0
    End If

    If this.condTests Is Nothing Then
        Set this.condTests = New BetterArray
        this.condTests.LowerBound = 1
        removeNumber = 0
    End If

    If removeNumber <= 0 Then
        this.condVars.Clear
        this.condTests.Clear
    Else
        For counter = 1 To removeNumber
            this.condVars.Pop
            this.condTests.Pop
        Next
    End If
End Sub

Private Property Get CurrentExcelFormula(Optional ByVal scope As ExcelFormulaScope = FormulaScopeAnalysis, _
                                         Optional ByVal prefix As String = "f") As String
    'use the formula
    Dim cond As IFormulaCondition
    Dim tabFormObject As IFormulas
    Dim eform As String

    Set tabFormObject = TableFormula
       
    If scope = 1 Then
        eform = tabFormObject.ParsedLinelistFormula(useTableName:=True, tablePrefix:=prefix)
    Else
        If this.condVars Is Nothing Or this.condTests Is Nothing Then GoTo Fail
        Set cond = FormulaCondition.Create(this.condVars, this.condTests)
        eform = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:=prefix)
    End If
  
    'Test here the formula before assigining it
    If FailedFormula(eform) Then
        GoTo Fail
    End If

    CurrentExcelFormula = eform
    Exit Property

Fail:
    'Trace Log here on failed formula?
    eform = Chr(34) & "##== formula parsing failed :( ==##" & Chr(34)
    CurrentExcelFormula = eform
End Property

'Parsing Formulas -----------------------------------------------------------------------
Private Sub AddFormulas()
    Dim tabId As String
    Dim tabType As AnalysisTablesType
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
    Dim rowCatAddress As String
    Dim colCatAddress As String
    Dim startTimeAddress As String
    Dim endTimeAddress As String
    Dim vars As ILLVariables
    
    Dim specs As ITablesSpecs
    
    If Not valid() Then Exit Sub                 'Trace here for validity

    Set specs = Specifications
    Set vars = LLVariables.Create(specs.Dictionary)
    Set sh = Table.Wksh()

    tabId = Id()
    secId = sectionId()
    tabType = TableType
    
    With sh
        Select Case tabType

        Case TypeGlobalSummary                   'Global summary formula
            finalFormula = CurrentExcelFormula(scope:=1, prefix:="") 'Parse the formula object using table names on analysis
            Set rng = .Range("STARTROW_" & tabId)
            rng.Cells(1, 2).FormulaArray = "= " & finalFormula

            'Now on filtered data (should normaly work if the previous didn't failed)
            finalFormula = CurrentExcelFormula(scope:=1, prefix:="f")
            rng.Cells(1, 3).FormulaArray = "= " & finalFormula

        Case TypeUnivariate

            Set rng = .Range("ROW_CATEGORIES_" & tabId)
            rowVar = specs.Value("row")
            ClearTableConditions
            AddTableConditions rowVar, "=" & rng.Cells(1, 1).Address(Rowabsolute:=False)
            finalFormula = CurrentExcelFormula()

            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set cellRng = rng.Cells(1, 1)
            cellRng.FormulaArray = "= " & finalFormula
            cellRng.AutoFill Destination:=rng, Type:=xlFillValues

            'Missing and percentages
            If specs.HasMissing() Then

                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr() 'empty string for empstr()

                finalFormula = CurrentExcelFormula()
                Set rng = .Range("MISSING_ROW_VALUES_" & tabId)
                rng.Cells(1, 1).FormulaArray = "=" & finalFormula

                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)
                
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & NonEmpStr() 'Non empty string for nonempstr()
                    'Formula on total (no choice but to put two conditions non empty and empty)
                    finalFormula = finalFormula & " + " & CurrentExcelFormula()
                    rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                Else
                
                    ClearTableConditions
                    finalFormula = CurrentExcelFormula(scope:=FormulaScopeLinelist)
                    rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                End If

            Else
                'No Missing
                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)
                ClearTableConditions
                AddTableConditions rowVar, "= " & NonEmpStr() 'rowvar = "<>"
                finalFormula = CurrentExcelFormula()
                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
            End If

            'Now percentage
            If specs.HasPercentage() Then
                Set rng = .Range("VALUES_COL_1_" & tabId)
                Set cellRng = .Range("TOTAL_ROW_VALUES_" & tabId).Cells(1, 1)

                rng.Offset(, 1).Cells(1, 1).formula = "= " & Percentage(cellRng, rng.Cells(1, 1), True)
                rng.Offset(, 1).Cells(1, 1).AutoFill Destination:=rng.Offset(, 1), Type:=xlFillValues

                'Last range on percentage
                cellRng.Cells(1, 2).formula = "= " & Percentage(cellRng.Cells(1, 1), cellRng.Cells(1, 1))

            End If

        Case TypeBivariate

            rowVar = specs.Value("row")
            colVar = specs.Value("column")
            rowCatAddress = .Range("ROW_CATEGORIES_" & tabId).Cells(1, 1).Address(Rowabsolute:=False)
            tabEndColumn = Table.EndColumn
            'Row categorie
            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)

            Do While (rng.Column < tabEndColumn)

                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "=" & rngLab.Address

                finalFormula = CurrentExcelFormula()

                'Test here the finalFormula
                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues

                'Missing Row
                If specs.Value("missing") = "row" Or specs.Value("missing") = "all" Then
                    Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
                    ClearTableConditions
                    AddTableConditions rowVar, "= " & EmpStr() 'Empty string
                    AddTableConditions colVar, "= " & rngLab.Address

                    finalFormula = CurrentExcelFormula()
                    
                    Set cellRng = .Cells(misRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula

                    'Formula on total:
                    'Define new formula condition on only the column
                    ClearTableConditions
                    AddTableConditions colVar, "= " & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula
                Else
                    'Formula on total should remove missing from row variable
                    ClearTableConditions
                    AddTableConditions rowVar, "=" & NonEmpStr()
                    AddTableConditions colVar, "=" & rngLab.Address
                    finalFormula = CurrentExcelFormula()
                    Set cellRng = .Cells(totRng.Row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula

                End If
                
                'Percentages
                If specs.HasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)    'percentage range
                    Set rngLab = rngLab.Offset(, 1)
                    'Include the total on the percentages
                    rng.Cells(1, 1).AutoFill Destination:=.Range(.Cells(rng.Row, rng.Column), .Cells(totRng.Row, rng.Column)), Type:=xlFillValues
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
                'Formula array on total
            Loop

            'Missing column and total columns
            If specs.Value("missing") = "column" Or specs.Value("missing") = "all" Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                misRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues

                'Total when there are missing columns
                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                finalFormula = CurrentExcelFormula()
                'The range is filled bellow
            Else
                'Total column, without missing
                ClearTableConditions
                AddTableConditions rowVar, "= " & rowCatAddress
                AddTableConditions colVar, "=" & NonEmpStr()
                finalFormula = CurrentExcelFormula()
            End If

            'Fill the total range with the formula
            Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
            totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
            totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues

            'missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'missing Row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'total row, missing column
            If RangeExists("TOTAL_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions colVar, "= " & EmpStr()
                finalFormula = CurrentExcelFormula()
                .Range("TOTAL_MISSING_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'total row, total column, depend on formula, I will just sum the column
            If RangeExists("TOTAL_TOTAL_" & tabId) Then
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    .Range("TOTAL_TOTAL_" & tabId).formula = "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
                Else
                    Select Case specs.Value("missing")
                    Case "all"
                        finalFormula = CurrentExcelFormula(scope:=1)
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
                    Case "row"                   'If missing on rows only, compute the total by removing non empty on columns
                        ClearTableConditions
                        AddTableConditions colVar, "= " & NonEmpStr()
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Case "column"                'If missing on columns only, compute the total by removing non empty on rows
                        ClearTableConditions
                        AddTableConditions rowVar, "= " & NonEmpStr()
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    End Select
                End If
            End If

        Case TypeTimeSeries                      ' FORMULAS FOR TIME SERIES =====================

            rowVar = specs.Value("time")
            colVar = specs.Value("column")

            'Rows for last day, and the period
            If specs.isNewSection() Then         'New sectio on time series
                'Period column
                Set rng = .Range("ROW_CATEGORIES_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                'cellRng.formula = "= FormatDateFromLastDay(TIME_UNIT_" & tabId & "," & _
                cellRng.Cells(1, -1).Address(rowabsolute:=False) & "," & _
                cellRng.Cells(1, 0).Address(rowabsolute:=False) & ", " & _
                "END_DATE_" & tabId & ")"
                'cellRng.AutoFill Destination:=rng, Type:=xlFillValues
                'Start period column
                Set cellRng = .Range("FIRST_VALUE_START_TIME_" & tabId)
                cellRng.formula = "= " & "START_DATE_" & tabId
                cellRng.Cells(2, 1).formula = "= " & cellRng.Cells(1, 2).Address(Rowabsolute:=False) & " + " & " 1 "
                Set rng = .Range("START_TIME_PERIOD_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues
      
                'End period column
                Set rng = .Range("END_TIME_PERIOD_" & tabId)
                Set cellRng = rng.Cells(1, 1)
                cellRng.formula = "= FindLastDay(TIME_UNIT_" & tabId & "," & _
                                  cellRng.Cells(1, 0).Address(Rowabsolute:=False) & ")"
                cellRng.AutoFill Destination:=rng, Type:=xlFillValues
                'Total period
                Set rng = .Range("TOTAL_ROW_" & tabId)
                rng.Cells(1, -1).formula = "= START_DATE_" & tabId
                rng.Cells(1, 0).formula = "= END_DATE_" & tabId
                'Validation max and validation min
                Set rng = .Range("VALIDATION_MAX_DATE_" & tabId)
                rng.formula = "=ValidMax(USER_START_DATE_" & tabId & _
                              ", USER_END_DATE_" & tabId & ", MIN_MIN_DATE_" & tabId & _
                              ", MAX_MAX_DATE_" & tabId & ", TIME_UNIT_" & tabId & ")"

                Set rng = .Range("VALIDATION_MIN_DATE_" & tabId)
                rng.formula = "=ValidMin(USER_START_DATE_" & tabId & _
                              ", USER_END_DATE_" & tabId & ", MIN_MIN_DATE_" & tabId & _
                              ", MAX_MAX_DATE_" & tabId & ", TIME_UNIT_" & tabId & ")"
                
                'Start Date and end Date
                Set rng = .Range("START_DATE_" & tabId)
                rng.formula = "= MAX(USER_START_DATE_" & tabId & ", VALIDATION_MIN_DATE_" & tabId & ")"

                Set rng = .Range("END_DATE_" & tabId)
                rng.formula = "= MIN(USER_END_DATE_" & tabId & ", VALIDATION_MAX_DATE_" & tabId & ")"

                'Informatation on start and end date for the user:
                Set rng = .Range("INFO_START_DATE_" & tabId)
                rng.formula = "= InfoUser(USER_START_DATE_" & tabId & ", START_DATE_" & tabId & ")"

                Set rng = .Range("INFO_END_DATE_" & tabId)
                rng.formula = "= InfoUser(USER_END_DATE_" & tabId & ", END_DATE_" & tabId & ", 2)"

                On Error Resume Next

                With .Range("TIME_UNIT_" & tabId).validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, Formula1:="=TIME_UNIT_LIST"
                End With
                On Error GoTo 0
            End If

            'Fill the ranges with formulas (Approximatly the same thing as previously)
            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1)) 'Add total line
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)
            Set cellRng = .Range("FIRST_VALUE_START_TIME_" & secId) 'Section Id because time series share same time period
            Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
            Set perRng = cellRng.Cells(1, 3)     'Period Range
            startTimeAddress = cellRng.Address(Rowabsolute:=False)
            endTimeAddress = cellRng.Cells(1, 2).Address(Rowabsolute:=False)
            tabEndColumn = Table.EndColumn
            
            Do While (rng.Column <= tabEndColumn)
                ClearTableConditions
                AddTableConditions rowVar, ">=" & startTimeAddress
                AddTableConditions rowVar, "<=" & endTimeAddress
                AddTableConditions colVar, "= " & rngLab.Address
                finalFormula = CurrentExcelFormula()

                'Formula on one column for time series
                rng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues

                'Missing Row
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & rngLab.Address
                
                finalFormula = CurrentExcelFormula()

                Set cellRng = .Cells(misRng.Row, rng.Column)
                cellRng.FormulaArray = "= " & finalFormula

                'Percentage
                If specs.HasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.Row, rng.Column)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)    'percentage column range
                    Set rngLab = rngLab.Offset(, 1)
                    rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
                    
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
            Loop

            'Missing column and total columns
            If specs.Value("missing") = "yes" Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
                Set misRng = .Range(misRng.Cells(1, 1), misRng.Cells(misRng.Rows.Count + 1, 1)) 'Add total line
                
                'Missing column
                ClearTableConditions
                AddTableConditions rowVar, ">=" & startTimeAddress
                AddTableConditions rowVar, "<=" & endTimeAddress
                AddTableConditions colVar, "= " & EmpStr()

                finalFormula = CurrentExcelFormula()
                
                misRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
                
                'Total including Missing
                If specs.Value("total") = "yes" Then
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    
                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                End If

            Else
                'Total, without missing
                If specs.Value("total") = "yes" Then
                    'Total column, without missing
                    ClearTableConditions
                    AddTableConditions rowVar, ">=" & startTimeAddress
                    AddTableConditions rowVar, "<=" & endTimeAddress
                    AddTableConditions colVar, "= " & NonEmpStr()

                    finalFormula = CurrentExcelFormula()

                    Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                    totRng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
                    totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
                End If
            End If

            'Row total, missing row

            'missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "=" & EmpStr()
                finalFormula = CurrentExcelFormula()

                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'missing Row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                ClearTableConditions
                AddTableConditions rowVar, "= " & EmpStr()
                AddTableConditions colVar, "= " & EmpStr()

                finalFormula = CurrentExcelFormula()
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'total row, total column, depend on formula, I will just sum the column
            If RangeExists("TOTAL_TOTAL_" & tabId) Then
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    .Range("TOTAL_TOTAL_" & tabId).formula = "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
            
                Else
                    ClearTableConditions
                    Set cellRng = .Range("TOTAL_ROW_" & secId).Cells(1, 0) 'section because I append tables on each new section
                    AddTableConditions rowVar, "<=" & cellRng.Address
                    AddTableConditions rowVar, ">=" & cellRng.Cells(1, 0).Address
              
                    If specs.Value("missing") = "yes" Then
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    Else
                        AddTableConditions colVar, "= " & NonEmpStr()
                        finalFormula = CurrentExcelFormula()
                        .Range("TOTAL_TOTAL_" & tabId).FormulaArray = "=" & finalFormula
                    End If
                End If
            End If

        Case TypeSpatial
        
        
        



        End Select
    End With
End Sub

'Errors
Private Sub ThrowExecutionError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error : " & errorMessage
End Sub

Private Sub UnFoundRangeError(ByVal sectionName As String, ByVal rngName As String, ByVal shName As String)
    ThrowExecutionError "On " & sectionName & ", range " & rngName & "Does not Exists in worksheet " & shName
End Sub

Public Property Get ICrossTableFormula_Valid() As Boolean
    ICrossTableFormula_Valid = valid
End Property

Public Sub ICrossTableFormula_AddFormulas()
    AddFormulas
End Sub


