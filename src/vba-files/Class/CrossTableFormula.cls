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

Private Type TCrossTableFormula
  tab As ICrossTable
  fdata As IFormulaData
  forml As IFormulas
End Type

Private this As TCrossTableFormula
Private Const CLASSNAME As String = "CrossTableFormula"

Public Function Create(ByVal tabl As ICrossTable, fdata As IFormulaData) As ICrossTableFormula
  With New CrossTableFormula
    Set .Table = tabl
    Set .formData = fdata
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
  Set formData = this.fdata
End Property

Public Property Set formData(ByVal fdata As IFormulaData)
  Set this.fdata = fdata
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
  Dim tabForm As String
  Dim fdata As IFormulaData
  Dim specs As ITablesSpecs
  Dim dict As ILLdictionary
  Dim tabFormObject As IFormulas

  If (this.forml Is Nothing) Then
    Set fdata = formData()
    Set specs = Specifications
    Set dict = specs.Dictionary
    tabForm = SummaryFunction
    Set tabFormObject = Formulas.Create(dict, fdata, tabForm)

    Set this.forml = tabFormObject
    Set TableFormula = tabFormObject
  Else
    Set TableFormula = this.forml
  End If
End Property

Private Property Get Valid() As Boolean
  Dim tabFormObject As IFormulas
  Set tabFormObject = TableFormula
  Valid = tabFormObject.Valid(formulaType:="analysis")
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

Private Property Get FailedFormula(ByVal frmName As String) As Boolean
  Dim sh As Worksheet
  Dim formTest As String

  Set sh = Table.Wksh
  With sh
    .Cells(1, 1).Clear
    On Error Resume Next
    .Cells(1, 1).FormulaArray = "= " & frmName
    On Error GoTo 0
    formTest = .Cells(1, 1).FormulaArray
  End With
  FailedFormula = (formTest = vbNullString)
End Property

Private Function EmpStr() As String
  EmpStr = Chr(34) & Chr(34)
End Function

Private Function NonEmpStr() As String
  NonEmpStr = Chr(34) & "<>" & Chr(34)
End Function

Private Function Percentage(ByVal denomRng As Range, ByVal numRng As Range, Optional ByVal rowabs As Boolean = True) As String
  Dim parsedForm As String
  parsedForm = numRng.Address(rowabsolute:=False) & "/" & denomRng.Address(rowabsolute:=rowabs, ColumnAbsolute:=True)
  Percentage = "IF(ISERR(" & parsedForm & ")," & EmpStr & "," & parsedForm & ")"
End Function

Private Function Condition(ByVal rngCondition As Range, ByVal formulaValue As String, Optional ByVal rowabs As Boolean = False) As String
  Condition = "IF(" & rngCondition.Address(rowabsolute:=rowabs) & "=" & EmpStr & "," & EmpStr & "," & formulaValue & ")"
End Function

'Parsing Formulas -----------------------------------------------------------------------
Private Sub AddFormulas()
    Dim tabId As String
    Dim tabFormObject As IFormulas
    Dim tabType As AnalysisTablesType
    Dim sh As Worksheet
    Dim rng As Range
    Dim totRng As Range
    Dim rngLab As Range
    Dim misRng As Range
    Dim cellRng As Range
    Dim finalFormula As String
    Dim cond As IFormulaCondition
    Dim condVars As BetterArray
    Dim condTest As BetterArray
    Dim cond2 As IFormulaCondition
    Dim condVars2 As BetterArray
    Dim condTest2 As BetterArray
    Dim vars As ILLVariables
    Dim tabName As String                        'Table names of variables in the same crosstable
    Dim targetRangeName As String
    Dim specs As ITablesSpecs
    Dim tabEndColumn As Long
    Dim secId As String
    Dim perRng As Range


    If Not Valid() Then Exit Sub                 'Trace here for validity

    Set specs = Specifications
    Set tabFormObject = TableFormula
    Set sh = Table.Wksh()
    Set vars = LLVariables.Create(specs.Dictionary)
    Set condVars = New BetterArray
    Set condTest = New BetterArray
    Set condVars2 = New BetterArray
    Set condTest2 = New BetterArray

    condVars.LowerBound = 1
    condTest.LowerBound = 1

    condVars2.LowerBound = 1
    condTest2.LowerBound = 1

    tabId = Id()
    secId = sectionId()
    tabType = TableType
    
    With sh
        Select Case tabType

        Case TypeGlobalSummary                   'Global summary formula
            'Test the formula in one sheet
            finalFormula = tabFormObject.ParsedLinelistFormula(useTableName:=True)

            'If it is empty, error, trace here
            If FailedFormula(finalFormula) Then Exit Sub 'Trace error on formula

            'No Errors, we can paste the formula
            'If the range does not exists, trace error on formula
            targetRangeName = "STARTROW_" & tabId

            If Not RangeExists(targetRangeName) Then UnFoundRangeError "Global summary", targetRangeName, sh.Name
            Set rng = .Range(targetRangeName)
            rng.Cells(1, 2).FormulaArray = "= " & finalFormula

            'Now on filtered data (should normaly work if the previous didn't failed)
            finalFormula = tabFormObject.ParsedLinelistFormula(useTableName:=True, tablePrefix:="f")
            rng.Cells(1, 3).FormulaArray = "= " & finalFormula

        Case TypeUnivariate

            'Test if the range on categories exists
            targetRangeName = "ROW_CATEGORIES_" & tabId
            If Not RangeExists(targetRangeName) Then UnFoundRangeError "Univariate analysis", targetRangeName, sh.Name
            Set rng = .Range(targetRangeName)

            condVars.Push specs.Value("row")
            condTest.Push "=" & rng.Cells(1, 1).Address(rowabsolute:=False)

            'Test the formula
            Set cond = FormulaCondition.Create(condVars, condTest)
            finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
            Debug.Print finalFormula

            If FailedFormula(finalFormula) Then Exit Sub 'trace error on formula on univariate

            targetRangeName = "VALUES_COL_1_" & tabId
            If Not RangeExists(targetRangeName) Then UnFoundRangeError "Univariate analysis", targetRangeName, sh.Name
            Set rng = .Range(targetRangeName)
            Set cellRng = rng.Cells(1, 1)

            cellRng.FormulaArray = "= " & finalFormula
            cellRng.AutoFill Destination:=rng, Type:=xlFillValues

            'Missing and percentages
            If specs.HasMissing() Then

                condTest.Clear
                condTest.Push "=" & EmpStr()
                Set cond = FormulaCondition.Create(condVars, condTest)

                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                Set rng = .Range("MISSING_ROW_VALUES_" & tabId)
                Debug.Print finalFormula

                rng.Cells(1, 1).FormulaArray = "=" & finalFormula

                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)

                condTest.Clear
                condTest.Push "= " & NonEmpStr()
                Set cond = FormulaCondition.Create(condVars, condTest)

                'Formula on total (no choice but to put two conditions non empty and empty)
                finalFormula = finalFormula & " + " & tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")

                rng.Cells(1, 1).FormulaArray = "= " & finalFormula

            Else
                'No Missing
                Set rng = .Range("TOTAL_ROW_VALUES_" & tabId)

                condTest.Clear
                condTest.Push "= " & NonEmpStr()
                Set cond = FormulaCondition.Create(condVars, condTest)

                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")

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
            condVars.Clear
            condTest.Clear
            condVars.Push specs.Value("row"), specs.Value("column")
            tabEndColumn = Table.EndColumn
            'Row categorie
            Set rng = .Range("VALUES_COL_1_" & tabId)
            Set rngLab = .Range("LABEL_COL_1_" & tabId)
            Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)

            Do While (rng.Column < tabEndColumn)
                condTest2.Clear
                condVars2.Clear
                condTest.Clear

                condTest.Push "=" & .Range("ROW_CATEGORIES_" & tabId).Cells(1, 1).Address(rowabsolute:=False)
                condTest.Push "=" & rngLab.Address

                Set cond = FormulaCondition.Create(condVars, condTest)
                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                
                Debug.Print finalFormula
                'Test here the finalFormula
                rng.Cells(1, 1).FormulaArray = "= " & finalFormula
                rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues

                'Missing Row
                If specs.Value("missing") = "row" Or specs.Value("missing") = "all" Then
                    Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
                    condVars2.Push specs.Value("column")
                    condTest2.Push "=" & rngLab.Address 'Label column address
                    condVars2.Push specs.Value("row")  'column variable
                    condTest2.Push "=" & EmpStr()     'Empty string

                    Set cond2 = FormulaCondition.Create(condVars2, condTest2)
                    finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
                    
                    Debug.Print finalFormula
                    
                    Set cellRng = .Cells(misRng.row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula

                    'Formula on total is only on the column (we don't care about missing)
                    condVars2.Pop
                    condTest2.Pop
                    'Define new formula condition on only the column
                    Set cond2 = FormulaCondition.Create(condVars2, condTest2)
                    finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
                    
                    Debug.Print finalFormula
                    
                    Set cellRng = .Cells(totRng.row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula

                Else
                    'Formula on total should remove missing from row variable
                    condVars2.Clear
                    condTest2.Clear

                    condVars2.Push specs.Value("column")
                    condTest2.Push "=" & rngLab.Address
                    condVars2.Push specs.Value("row")
                    condTest2.Push "=" & NonEmpStr()

                    Set cond2 = FormulaCondition.Create(condVars2, condTest2)
                    finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
                    Debug.Print finalFormula
                    
                    Set cellRng = .Cells(totRng.row, rng.Column)
                    cellRng.FormulaArray = "= " & finalFormula
                End If

                If specs.HasPercentage Then
                    Select Case specs.Value("percentage")
                    Case "row"
                        Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
                    Case "column"
                        Set cellRng = .Cells(totRng.row, rng.Column)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    Case "total"
                        Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                        rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
                    End Select
                    Set rng = rng.Offset(, 1)    'percentage range
                    Set rngLab = rngLab.Offset(, 1)
                    'Include the total on the percentages
                    rng.Cells(1, 1).AutoFill Destination:=.Range(.Cells(rng.row, rng.Column), .Cells(totRng.row, rng.Column)), Type:=xlFillValues
                End If

                Set rng = rng.Offset(, 1)
                Set rngLab = rngLab.Offset(, 1)
                'Formula array on total
            Loop

            'Missing column and total columns
            If specs.Value("missing") = "column" Or specs.Value("missing") = "all" Then
                Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
                condVars.Clear
                condTest.Clear
                condVars.Push specs.Value("row")
                condTest.Push "=" & .Range("ROW_CATEGORIES_" & tabId).Cells(1, 1).Address(rowabsolute:=False)
                condVars.Push specs.Value("col")
                condTest.Push "=" & EmpStr()
                Set cond = FormulaCondition.Create(condVars, condTest)
                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                
                Debug.Print finalFormula
                
                misRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues

                'Now work on total column
                condVars.Pop
                condTest.Pop
                'Define new formula condition on only the column
                Set cond = FormulaCondition.Create(condVars, condTest)
                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues

            Else
                'Total column, without missing
                condVars.Clear
                condTest.Clear

                condVars.Push specs.Value("row")
                condTest.Push "=" & .Range("ROW_CATEGORIES_" & tabId).Cells(1, 1).Address(rowabsolute:=False)
                condVars.Push specs.Value("column")
                condTest.Push "=" & NonEmpStr()

                Set cond = FormulaCondition.Create(condVars, condTest)
                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                Debug.Print finalFormula
                
                Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
                totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
                totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
            End If


            'missing row, total column
            If RangeExists("MISSING_TOTAL_" & tabId) Then
                condVars.Clear
                condTest.Clear
         
                condVars.Push specs.Value("row")
                condTest.Push "=" & EmpStr()
                Set cond = FormulaCondition.Create(condVars, condTest)
                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                
                Debug.Print finalFormula
                
                .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'missing Row, missing column
            If RangeExists("MISSING_MISSING_" & tabId) Then
                condVars.Clear
                condTest.Clear
          
                condVars.Push specs.Value("row")
                condVars.Push specs.Value("column")
                condTest.Push "=" & EmpStr(), "=" & EmpStr()
                Set cond = FormulaCondition.Create(condVars, condTest)
                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                Debug.Print finalFormula
                
                .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'missing column, total row
            If RangeExists("TOTAL_MISSING_" & tabId) Then
                condVars.Clear
                condTest.Clear
          
                condVars.Push specs.Value("column")
                condTest.Push "=" & EmpStr()
                Set cond = FormulaCondition.Create(condVars, condTest)
                finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
                
                Debug.Print finalFormula
                
                .Range("TOTAL_MISSING_" & tabId).FormulaArray = "= " & finalFormula
            End If

            'total row, total column, depend on formula, I will just sum the column
            If RangeExists("TOTAL_TOTAL_" & tabId) Then
                If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                    .Range("TOTAL_TOTAL_" & tabId).formula = "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
                Else
                    .Range("TOTAL_TOTAL_" & tabId).formula = ""
                    .Range("TOTAL_TOTAL_" & tabId).FormulaArray = ""
                End If
            End If

        Case TypeTimeSeries ' FORMULAS FOR TIME SERIES =====================

          'Rows for last day, and the period
          If specs.isNewSection() Then
              'Period column
              Set rng = .Range("ROW_CATEGORIES_" & tabId)
              Set cellRng = rng.Cells(1, 1)
              cellRng.formula = "= FormatDateFromLastDay(" & .Range("TIME_UNIT_" & tabId).Address & "," & _
                                                                                                  cellRng.Cells(1, -1).Address(rowabsolute:=False) & "," & _
                                                                                                  cellRng.Cells(1, 0).Address(rowabsolute:=False) & ", " & _
                                                                                                  .Range("END_DATE_" & tabId).Address & ")"
              cellRng.AutoFill Destination:=rng, Type:=xlFillValues
              'Start period column
              Set cellRng = .Range("FIRST_VALUE_START_TIME_" & tabId)
              cellRng.formula = "=" & Range("START_DATE_" & tabId).Address
              cellRng.Cells(2, 1).formula = "= " & cellRng.Cells(1, 2).Address(rowabsolute:=False) & " + " & " 1 "
              Set rng = .Range("START_TIME_PERIOD_" & tabId)
              Set cellRng = rng.Cells(1, 1)
              cellRng.AutoFill Destination:=rng, Type:=xlFillValues
      
              'End period column
              Set rng = .Range("END_TIME_PERIOD_" & tabId)
              Set cellRng = rng.Cells(1, 1)
              cellRng.formula = "= FindLastDay(" & .Range("TIME_UNIT_" & tabId).Address & "," & _
                                                                                        cellRng.Cells(1, 0).Address(rowabsolute:=False) & ")"
              cellRng.AutoFill Destination:=rng, Type:=xlFillValues
              'Total period
              Set rng = .Range("TOTAL_ROW_" & tabId)
              rng.Cells(1, -1).formula = "= " & .Range("START_DATE_" & tabId).Address
              rng.Cells(1, 0).formula = "= " & .Range("END_DATE_" & tabId).Address
              On Error Resume Next
              With .Range("TIME_UNIT_" & tabId).Validation
                  .Delete
                  .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                   Operator:=xlBetween, Formula1:="=TIME_UNIT_LIST"
              End With
      
              With .Range("START_DATE_" & tabId).Validation
                  .Delete
                  .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                       Operator:=xlBetween, Formula1:="=VALIDATION_MIN_DATE_" & tabId, _
                       Formula2:="=VALIDATION_MAX_DATE_" & tabId
              End With
              On Error GoTo 0
          End If

          'Fill the ranges with formulas (Approximatly the same thing as previously)
          Set rng = .Range("VALUES_COL_1_" & tabId)
          Set rng = .Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count + 1, 1)) 'Add total line
          Set rngLab = .Range("LABEL_COL_1_" & tabId)
          Set totRng = .Range("TOTAL_ROW_VALUES_" & tabId)
          Set cellRng = .Range("FIRST_VALUE_START_TIME_" & secId) 'Section Id because time series share same time period
          Set perRng = cellRng.Cells(1, 3)
          condVars.Push specs.Value("time"), specs.Value("time"), specs.Value("column")
          condVars2.Push specs.Value("time"), specs.Value("column")
          condTest2.Push "=" & EmpStr()
          condTest.Push ">=" & cellRng.Address(rowabsolute:=False)
          condTest.Push "<=" & cellRng.Cells(1, 2).Address(rowabsolute:=False)


          tabEndColumn = Table.EndColumn

          Do While (rng.Column < tabEndColumn)

            condTest.Push "=" & rngLab.Address
            Set cond = FormulaCondition.Create(condVars, condTest)
            finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
            'Test here the finalFormula
            rng.Cells(1, 1).FormulaArray = "= " & Condition(perRng, finalFormula)
            rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
            'Missing Row
            Set misRng = .Range("MISSING_ROW_VALUES_" & tabId)
            condTest2.Push "=" & rngLab.Address  'Label column address
            Set cond2 = FormulaCondition.Create(condVars2, condTest2)
            finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
            Set cellRng = .Cells(misRng.row, rng.Column)
            cellRng.FormulaArray = "= " & finalFormula

            'Percentage
            If specs.HasPercentage Then
              Select Case specs.Value("percentage")
              Case "row"
                  Set cellRng = .Range("TOTAL_COL_VALUES_" & tabId).Cells(1, 1)
                  rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1), rowabs:=False)
              Case "column"
                  Set cellRng = .Cells(totRng.row, rng.Column)
                  rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
              Case "total"
                  Set cellRng = .Range("TOTAL_TOTAL_" & tabId)
                  rng.Cells(1, 2).formula = "= " & Percentage(cellRng, rng.Cells(1, 1))
              End Select
              Set rng = rng.Offset(, 1)    'percentage range
              Set rngLab = rngLab.Offset(, 1)
              rng.Cells(1, 1).AutoFill Destination:=rng, Type:=xlFillValues
            End If

            Set rng = rng.Offset(, 1)
            Set rngLab = rngLab.Offset(, 1)
            condTest.Pop
            condTest2.Pop
          Loop

          'Missing column and total columns
          If specs.Value("missing") = "yes" Then
            Set misRng = .Range("MISSING_COL_VALUES_" & tabId)
            Set misRng = .Range(misRng.Cells(1, 1), misRng.Cells(misRng.Rows.Count + 1, 1)) 'Add total line
            condTest.Push "=" & EmpStr()
            Set cond = FormulaCondition.Create(condVars, condTest)
            finalFormula = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:="f")
            misRng.Cells(1, 1).FormulaArray = "= " & finalFormula
            misRng.Cells(1, 1).AutoFill Destination:=misRng, Type:=xlFillValues
            condTest.Pop

            If specs.Value("total") = "yes" Then
              'Now work on total column
              Set condVars2 = condVars.Clone()
              condVars2.Pop
              Set condTest2 = condTest.Clone()
              'Define new formula condition on only the column
              Set cond2 = FormulaCondition.Create(condVars2, condTest2)
              finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
              Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
              totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
              totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
            End If

          Else
            If specs.Value("total") = "yes" Then
              'Total column, without missing
              Set condVars2 = condVars.Clone()
              Set condTest2 = condTest.Clone()
              condTest2.Push "=" & NonEmpStr()
              Set cond2 = FormulaCondition.Create(condVars2, condTest2)
              finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
              Set totRng = .Range("TOTAL_COL_VALUES_" & tabId)
              totRng.Cells(1, 1).FormulaArray = "= " & finalFormula
              totRng.Cells(1, 1).AutoFill Destination:=totRng, Type:=xlFillValues
            End If
          End If

          'Row total, missing row

          'missing row, total column
          If RangeExists("MISSING_TOTAL_" & tabId) Then
            condVars2.Clear
            condVars2.Push specs.Value("time")
            condTest2.Clear
            condTest2.Push "=" & EmpStr()
            Set cond2 = FormulaCondition.Create(condVars2, condTest2)
            finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
            .Range("MISSING_TOTAL_" & tabId).FormulaArray = "= " & finalFormula
        End If

        'missing Row, missing column
        If RangeExists("MISSING_MISSING_" & tabId) Then
            condVars2.Clear
            condTest2.Clear
            condVars2.Push specs.Value("time"), specs.Value("column")
            condTest2.Push "=" & EmpStr(), "=" & EmpStr()
            Set cond2 = FormulaCondition.Create(condVars2, condTest2)
            finalFormula = tabFormObject.ParsedAnalysisFormula(cond2, tablePrefix:="f")
            .Range("MISSING_MISSING_" & tabId).FormulaArray = "= " & finalFormula
        End If

        'total row, total column, depend on formula, I will just sum the column
        If RangeExists("TOTAL_TOTAL_" & tabId) Then
            If SummaryFunction = "N" Or SummaryFunction = "N()" Then
                .Range("TOTAL_TOTAL_" & tabId).formula = "= SUM(" & .Range("TOTAL_COL_VALUES_" & tabId).Address & ")"
            Else
                 .Range("TOTAL_TOTAL_" & tabId).formula = ""
                 .Range("TOTAL_TOTAL_" & tabId).FormulaArray = ""
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
  ICrossTableFormula_Valid = Valid
End Property

Public Sub ICrossTableFormula_AddFormulas()
  AddFormulas
End Sub
