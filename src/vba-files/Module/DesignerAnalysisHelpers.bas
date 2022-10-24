Attribute VB_Name = "DesignerAnalysisHelpers"
Option Explicit
Option Private Module

'Format each analysis worksheet (global values for the worksheet)
Sub FormatAnalysisWorksheet(Wkb As Workbook, sSheetName As String, _
                            Optional sCodeName As String = vbNullString, _
                            Optional iColWidth As Integer = C_iLLFirstColumnsWidth)

    With Wkb.Worksheets(sSheetName)
        .Cells.EntireColumn.ColumnWidth = iColWidth
        .Cells.WrapText = True
        .Cells.EntireRow.AutoFit
    End With

    If sCodeName <> vbNullString Then TransferCodeWksh Wkb:=Wkb, sSheetName:=sSheetName, sNameModule:=sCodeName
End Sub

'FUNCTIONS USED TO BUILD UNIVARIATE ANALYSIS ===================================================================================================================

'Create New section
Sub CreateNewSection(Wksh As Worksheet, iRow As Long, iCol As Long, sSection As String, _
                     Optional sColor As String = "DarkBlue")
    With Wksh
        'New range, format the range
        FormatARange .Cells(iRow, iCol), sValue:=sSection, FontSize:=C_iAnalysisFontSize + 2, _
        sFontColor:=sColor, Horiz:=xlHAlignLeft, Verti:=xlVAlignBottom

        Range(.Cells(iRow, iCol), .Cells(iRow, iCol + 4)).Merge
        'Draw a border arround the section
        DrawLines rng:=Range(.Cells(iRow, iCol), .Cells(iRow, iCol + 6)), iWeight:=xlMedium, sColor:=sColor, At:="Bottom"

        'Heights for eventual charts
        .Cells(iRow - 1, iCol).RowHeight = C_iLLChartPartRowHeight
        .Cells(iRow + 1, iCol).RowHeight = C_iLLChartPartRowHeight
    End With
End Sub

'Create Headers for univariate analysis
Sub CreateUAHeaders(Wksh As Worksheet, iRow As Long, iCol As Long, _
                    sMainLab As String, sSummaryLabel As String, _
                    sPercent As String, Optional sColor As String = "DarkBlue")
    With Wksh
        'Variable Label from the dictionary
        FormatARange rng:=.Cells(iRow, iCol), sValue:=sMainLab, sFontColor:=sColor, isBold:=True, Horiz:=xlHAlignLeft
        'First column on sumary label
        FormatARange rng:=.Cells(iRow, iCol + 1), sValue:=sSummaryLabel, sFontColor:=sColor, isBold:=True
        'Add Percentage header column if required
        If sPercent = C_sYes Then FormatARange rng:=.Cells(iRow, iCol + 2), sValue:=TranslateLLMsg("MSG_Percent"), sFontColor:=sColor, isBold:=True
    End With
End Sub

'Create Headers for bivariate Analysis

Sub CreateBATable(Wksh As Worksheet, ColumnsData As BetterArray, _
                  iRow As Long, iCol As Long, _
                  sMainLabCol As String, _
                  sSummaryLabel As String, _
                  sPercent As String, sMiss As String, _
                  Optional RowsData As BetterArray, _
                  Optional sMainLabRow As String, _
                  Optional sInteriorColor As String = "VeryLightBlue", _
                  Optional sTotalInteriorColor As String = "VeryLightGreyBlue", _
                  Optional sNAFontColor As String = "GreyBlue", _
                  Optional sColor As String = "DarkBlue", _
                  Optional isTimeSeries As Boolean = False)

    Dim iEndRow As Long
    Dim i As Long
    Dim iLastCol As Long
    Dim iTotalLastCol As Long
    Dim iTotalFirstCol As Long
    Dim sArrow As String
    Dim HasPercent As Boolean
    Dim AddTotal As Boolean


    'Add Total for column
    AddTotal = True

    With Wksh
        'Variable Label from the dictionary for Row
        If Not isTimeSeries Then
            FormatARange rng:=.Cells(iRow + 2, iCol), sValue:=sMainLabRow, sFontColor:=sColor, isBold:=True

            'Merge the first and second rows of first column of bivariate analysis
            .Range(.Cells(iRow + 1, iCol), .Cells(iRow + 2, iCol)).Merge
            .Cells(iRow + 1, iCol).MergeArea.HorizontalAlignment = xlHAlignLeft
            .Cells(iRow + 1, iCol).MergeArea.VerticalAlignment = xlVAlignCenter

            'Add the rows Data -------------------------------------------------------------------------------------------------------------------------------

            RowsData.ToExcelRange .Cells(iRow + 3, iCol)
            'EndRow of the table
            iEndRow = iRow + 2 + RowsData.Length

            FormatARange rng:=.Range(.Cells(iRow + 3, iCol), .Cells(iEndRow, iCol)), sFontColor:=sColor, _
        sInteriorColor:=sInteriorColor, Horiz:=xlHAlignLeft
        Else
            iEndRow = iRow + 3 + C_iNbTime
        End If

        If sMiss = C_sAnaRow Or sMiss = C_sAnaAll Or isTimeSeries Then
            'This is to avoid adding Missing on time series for following column
            If iCol <= C_eStartColumnAnalysis + 2 Then .Cells(iEndRow + 1, iCol).Value = TranslateLLMsg("MSG_NA")
            'Format the last row, just in case we need
            FormatARange rng:=.Cells(iEndRow + 1, iCol), sFontColor:=sNAFontColor, sInteriorColor:=sTotalInteriorColor, _
        Horiz:=xlHAlignLeft
            iEndRow = iEndRow + 1
        End If

        'Variable label from the dictionary for column
        FormatARange rng:=.Cells(iRow, iCol + 1), sValue:=sMainLabCol, sFontColor:=sColor, isBold:=True, Horiz:=xlHAlignLeft

        'Now Add Percentage And Values for the column -----------------------------------------------------------------------------------------------------

        'If you have to add percentage :
        sArrow = vbNullString
        Select Case sPercent
        Case C_sAnaCol
            HasPercent = True
            sArrow = ChrW(8597)                  'Arrow is vertical
        Case C_sAnaRow
            HasPercent = True
            sArrow = ChrW(8596)                  'Arrow is horizontal
        Case C_sAnaTot
            HasPercent = True
        Case Else
            HasPercent = False
        End Select

        If ColumnsData.Length > 0 Then
            'There are categories related to the group on column

            If HasPercent Then
                i = 0
                Do While (i < ColumnsData.Length)
                    'There is percentage, we have to add the percentage
                    .Cells(iRow + 1, iCol + 2 * i + 1).Value = ColumnsData.Items(i + 1)
                    .Cells(iRow + 2, iCol + 2 * i + 1).Value = sSummaryLabel
                    .Cells(iRow + 2, iCol + 2 * i + 2).Value = TranslateLLMsg("MSG_Percent") & " " & sArrow
                    .Range(.Cells(iRow + 1, iCol + 2 * i + 1), .Cells(iRow + 1, iCol + 2 * i + 2)).Merge
                    'Write borders arround the different part of the columns
                    DrawLines rng:=.Range(.Cells(iRow + 1, iCol + 2 * i + 1), .Cells(iEndRow + 1, iCol + 2 * i + 2)), sColor:=sColor
                    DrawLines rng:=.Range(.Cells(iRow + 1, iCol + 2 * i + 1), .Cells(iEndRow + 1, iCol + 2 * i + 1)), At:="Left", iWeight:=xlThin, sColor:=sColor
                    i = i + 1
                Loop
                iLastCol = 2 * i                 'This is the last column of the table when there is percentage
            Else
                'There is no percentage, only column values
                i = 1
                Do While (i <= ColumnsData.Length)
                    .Cells(iRow + 1, iCol + i).Value = ColumnsData.Items(i)
                    .Cells(iRow + 2, iCol + i).Value = sSummaryLabel
                    'Draw lines arround all borders
                    DrawLines rng:=.Range(.Cells(iRow + 1, iCol + i), .Cells(iEndRow + 1, iCol + i)), sColor:=sColor
                    DrawLines rng:=.Range(.Cells(iRow + 1, iCol + i), .Cells(iEndRow + 1, iCol + i)), At:="Left", sColor:=sColor, iWeight:=xlThin
                    i = i + 1
                Loop
                iLastCol = i - 1                 'Last column of the table without the percentage
            End If
            iLastCol = iCol + iLastCol
            iTotalFirstCol = iLastCol + 1
        Else
            'There are no categories, only a custom function created by the user

            .Cells(iRow + 1, iCol + 1).Value = ""
            .Cells(iRow + 2, iCol + 1).Value = sSummaryLabel
            DrawLines rng:=.Range(.Cells(iRow + 1, iCol + 1), .Cells(iEndRow + 1, iCol + 1)), sColor:=sColor
            DrawLines rng:=.Range(.Cells(iRow + 1, iCol + 1), .Cells(iEndRow + 1, iCol + 1)), At:="Left", sColor:=sColor, iWeight:=xlThin
            iLastCol = iCol + 1
            iTotalFirstCol = iLastCol
            AddTotal = False
        End If

        'Add Missing for column --------------------------------------------------------------------------------------------------------------------------------

        If sMiss = C_sAnaCol Or sMiss = C_sAnaAll Or (sMiss = C_sYes And isTimeSeries) Then

            'Missing at the end of the column
            .Cells(iRow + 1, iTotalFirstCol).Value = TranslateLLMsg("MSG_NA")
            .Cells(iRow + 2, iTotalFirstCol).Value = sSummaryLabel
            iTotalFirstCol = iTotalFirstCol + 1

            'Add percentage
            If HasPercent Then

                .Cells(iRow + 2, iTotalFirstCol).Value = TranslateLLMsg("MSG_Percent") & " " & sArrow
                .Range(.Cells(iRow + 1, iTotalFirstCol - 1), .Cells(iRow + 1, iTotalFirstCol)).Merge

                'Now update the first column for total
                iTotalFirstCol = iTotalFirstCol + 1

            End If

            'Format the missing for column
            DrawLines rng:=.Range(.Cells(iRow + 1, iLastCol + 1), .Cells(iEndRow + 1, iTotalFirstCol - 1)), sColor:=sColor
            DrawLines rng:=.Range(.Cells(iRow + 1, iLastCol + 1), .Cells(iEndRow + 1, iTotalFirstCol - 1)), sColor:=sColor, iWeight:=xlThin, At:="Left"
            FormatARange rng:=.Range(.Cells(iRow + 1, iLastCol + 1), .Cells(iEndRow + 1, iTotalFirstCol - 1)), sInteriorColor:=sTotalInteriorColor, sFontColor:=sNAFontColor
        End If

        'The last column is initiated outside the total column formating
        iTotalLastCol = iTotalFirstCol

        'Add Total ------------------------------------------------------------------------------------------------------------------------------------------------
        If AddTotal Then

            .Cells(iRow + 1, iTotalFirstCol).Value = TranslateLLMsg("MSG_Total")
            .Cells(iRow + 2, iTotalFirstCol).Value = sSummaryLabel

            'In case it is needed, add percentage for total also
            If HasPercent Then
                .Cells(iRow + 2, iTotalLastCol + 1).Value = TranslateLLMsg("MSG_Percent") & " " & sArrow
                .Range(.Cells(iRow + 1, iTotalLastCol), .Cells(iRow + 1, iTotalLastCol + 1)).Merge
                iTotalLastCol = iTotalLastCol + 1
            End If

            'Format total

            'Add hairlines between cells
            DrawLines rng:=.Range(.Cells(iRow + 1, iTotalFirstCol), .Cells(iEndRow + 1, iTotalLastCol)), sColor:=sColor

            'Add a left double line
            DrawLines rng:=.Range(.Cells(iRow + 1, iTotalFirstCol), .Cells(iEndRow + 1, iTotalFirstCol)), sColor:=sColor, iLine:=xlDouble, At:="Left"

            'Add rigth thick line on time series

            'Format all the total range
            FormatARange rng:=.Range(.Cells(iRow + 1, iTotalFirstCol), .Cells(iEndRow + 1, iTotalLastCol)), sInteriorColor:=sTotalInteriorColor, isBold:=True
        End If

        'Add Missing for Rows (After total because we need total end column) -------------------------------------------------------------------------------------------------
        If sMiss = C_sAnaRow Or sMiss = C_sAnaAll Or isTimeSeries Then
            FormatARange rng:=.Range(.Cells(iEndRow, iCol + 1), .Cells(iEndRow, iTotalLastCol)), sInteriorColor:=sTotalInteriorColor, sFontColor:=sNAFontColor
        End If

        'Total on the Last line
        'Add total for time series only on first columns
        If iCol <= C_eStartColumnAnalysis + 2 Then .Cells(iEndRow + 1, iCol).Value = TranslateLLMsg("MSG_Total")
        FormatARange rng:=.Cells(iEndRow + 1, iCol), sInteriorColor:=sTotalInteriorColor, isBold:=True, Horiz:=xlHAlignLeft
        FormatARange rng:=.Range(.Cells(iEndRow + 1, iCol + 1), .Cells(iEndRow + 1, iTotalLastCol)), sInteriorColor:=sTotalInteriorColor, isBold:=True

        'Format Table Headers -------------------------------------------------------------------------------------------------------------------------------------------
        'First row with column categories
        FormatARange rng:=.Range(.Cells(iRow + 1, iCol + 1), .Cells(iRow + 1, iLastCol)), sFontColor:=sColor, sInteriorColor:=sInteriorColor
        'Second row with summary label with/without percentage
        FormatARange rng:=.Range(.Cells(iRow + 2, iCol + 1), .Cells(iRow + 2, iLastCol)), sFontColor:=sColor, FontSize:=C_iAnalysisFontSize - 1
        'Draw lines arround the first column of table
        If Not isTimeSeries Then DrawLines rng:=.Range(.Cells(iRow + 1, iCol), .Cells(iEndRow + 1, iCol)), sColor:=sColor
        'Thick line at the header row
        DrawLines rng:=.Range(.Cells(iRow + 2, iCol), .Cells(iRow + 2, iTotalLastCol)), At:="Bottom", iLine:=xlDouble, sColor:=sColor
        'Draw lines for Total
        DrawLines rng:=.Range(.Cells(iEndRow + 1, iCol), .Cells(iEndRow + 1, iTotalLastCol)), sColor:=sColor, iLine:=xlDouble, At:="Top"
        'Drawlines arround all the table
        If Not isTimeSeries Then WriteBorderLines oRange:=.Range(.Cells(iRow + 1, iCol), .Cells(iEndRow + 1, iTotalLastCol)), sColor:=sColor, iWeight:=xlThin

        'Put every values to right
        .Range(.Cells(iRow + 3, iCol + 1), .Cells(iEndRow + 1, iTotalLastCol)).HorizontalAlignment = xlHAlignCenter

    End With
End Sub

'Add interior formulas for the bivariate analysis

Sub AddInnerFormula(Wkb As Workbook, DictHeaders As BetterArray, sForm As String, _
                    iStartRow As Long, iStartCol As Long, iEndRow As Long, iEndCol, _
                    sPercent As String, sMiss As String, sVarRow As String, sVarColumn As String)

    Dim Wksh As Worksheet
    Dim iInnerEndRow As Long
    Dim iInnerEndCol As Long
    Dim i As Long
    Dim j As Long
    Dim istep As Long
    Dim sFormula As String

    Set Wksh = Wkb.Worksheets(sParamSheetAnalysis)

    iInnerEndRow = iEndRow - 1
    'There is a missing line
    If sMiss = C_sAnaRow Or sMiss = C_sAnaAll Then iInnerEndRow = iEndRow - 2

    iInnerEndCol = iEndCol - 1

    'There is missing at the end column
    If sMiss = C_sAnaCol Or sMiss = C_sAnaAll Then iInnerEndCol = iInnerEndCol - 1

    'Step for columns
    istep = 1

    If sPercent <> C_sNo Then                    'There is precentage.
        iInnerEndCol = iInnerEndCol - 1
        istep = 2
    End If

    i = iStartRow + 2
    j = iStartCol + 1

    With Wksh
        'Add Now the formulas
        Do While (i <= iInnerEndRow)

            j = iStartCol + 1

            Do While (j <= iInnerEndCol)

                sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, _
                                            sVarRow:=sVarRow, sVarColumn:=sVarColumn, _
                                            sConditionRow:=.Cells(i, iStartCol).Address, _
                                            sConditionColumn:=.Cells(iStartRow, j).Address, _
                                            isFiltered:=True)

                On Error Resume Next

                'Add the
                If sFormula <> vbNullString Then .Cells(i, j).FormulaArray = sFormula

                'adding the percentage columns

                If sPercent <> C_sNo Then

                    sFormula = vbNullString

                    Select Case sPercent

                    Case C_sAnaTot
                        'Percentage on all
                        sFormula = .Cells(i, j).Address & " / " & .Cells(iEndRow, iEndCol - 1).Address
                    Case C_sAnaCol
                        'Percentage on column
                        sFormula = .Cells(i, j).Address & " / " & .Cells(iEndRow, j).Address
                    Case C_sAnaRow
                        'Percentage on Row
                        sFormula = .Cells(i, j).Address & " / " & .Cells(i, iEndCol - 1).Address
                    End Select

                    'Add the percentage format now
                    With .Cells(i, j + 1)
                        .Style = "Percent"
                        .NumberFormat = "0.00%"

                        If sFormula <> vbNullString Then .formula = AddPercentage(sFormula)
                    End With

                End If

                On Error GoTo 0

                j = j + istep
            Loop

            i = i + 1

        Loop

    End With
End Sub

'Add formulas at Borders

Sub AddBordersFormula(Wkb As Workbook, DictHeaders As BetterArray, sForm As String, iStartRow As Long, iStartCol As Long, _
                      iEndRow As Long, iEndCol As Long, sVarRow As String, sVarColumn As String, sMiss As String, sPercent As String)


    Dim Wksh As Worksheet
    Dim i As Long
    Dim istep As Long

    Dim sFormula As String                       'Formula string
    Dim sFormula2 As String                      'Second formula when needed
    Dim includeMissing As Boolean
    Dim iTotalColumn As Long                     'Column for total, depending on wheter there is percentage or not
    Dim iMissingColumn As Long
    Dim iMissingRow As Long

    istep = 1
    iTotalColumn = iEndCol

    If sPercent <> C_sNo Then
        istep = 2
        iTotalColumn = iEndCol - 1
    End If

    Set Wksh = Wkb.Worksheets(sParamSheetAnalysis)

    With Wksh

        'Add Total Row ---------------------------------------------------------------------------------------------------------------------------------------
        i = iStartCol + 1
        includeMissing = (sMiss = C_sAnaRow Or sMiss = C_sAnaAll)

        Do While (i <= iEndCol)

            'Formula for the last line (Here I invert the places of rows and columns)
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarColumn, _
                                        OnTotal:=True, sConditionRow:=.Cells(iStartRow, i).Address, _
                                        sVarColumn:=sVarRow, includeMissing:=includeMissing, isFiltered:=True)


            If sFormula <> vbNullString Then .Cells(iEndRow, i).FormulaArray = sFormula

            'Formula for the percentage if there is one
            If sPercent <> C_sNo Then

                sFormula = vbNullString

                Select Case sPercent

                Case C_sAnaCol

                    sFormula = .Cells(iEndRow, i).Address & " / " & .Cells(iEndRow, i).Address
                    sFormula2 = .Cells(iEndRow - 1, i).Address & " / " & .Cells(iEndRow, i).Address   'Formula for missing percentage

                Case C_sAnaRow, C_sAnaTot

                    sFormula = .Cells(iEndRow, i).Address & " / " & .Cells(iEndRow, iTotalColumn).Address
                    sFormula2 = .Cells(iEndRow - 1, i).Address & " / " & .Cells(iEndRow, iTotalColumn).Address

                End Select

                'Add the percentage format now
                With .Cells(iEndRow, i + 1)
                    .Style = "Percent"
                    .NumberFormat = "0.00%"
                    If sFormula <> vbNullString Then .formula = AddPercentage(sFormula)
                End With
            End If

            'Formula for the missing if there is missing
            If includeMissing Then

                sFormula = vbNullString
                sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarRow, _
                                            OnTotal:=False, sConditionRow:=Chr(34) & Chr(34), _
                                            sVarColumn:=sVarColumn, sConditionColumn:=.Cells(iStartRow, i).Address)

                If sFormula <> vbNullString Then .Cells(iEndRow - 1, i).FormulaArray = sFormula

                'Update the missing row
                iMissingRow = iEndRow - 1

                If sPercent <> C_sNo Then
                    ' Add percentage for the missing
                    With .Cells(iEndRow - 1, i + 1)
                        .Style = "Percent"
                        .NumberFormat = "0.00%"
                        If sFormula2 <> vbNullString Then .formula = AddPercentage(sFormula2)
                    End With
                End If

            End If

            i = i + istep
        Loop

        'Add Total column --------------------------------------------------------------------------------------------------------------------------------------
        i = iStartRow + 2
        includeMissing = (sMiss = C_sAnaCol Or sMiss = C_sAnaAll)


        Do While (i <= iEndRow)

            sFormula = vbNullString
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarRow, _
                                        OnTotal:=True, sConditionRow:=.Cells(i, iStartCol).Address, _
                                        sVarColumn:=sVarColumn, includeMissing:=includeMissing, isFiltered:=True)

            'We need to add a percentage column If there is missing and if there is not
            If sPercent <> C_sNo Then
                'Percentages are a the end of the table, so we need to shift to -1
                If sFormula <> vbNullString Then .Cells(i, iEndCol - 1).FormulaArray = sFormula

                'Add the percentage on total
                Select Case sPercent

                    Case C_sAnaRow

                        sFormula = .Cells(i, iEndCol - 1).Address & " / " & .Cells(i, iEndCol - 1).Address
                        sFormula2 = .Cells(i, iEndCol - 3).Address & " / " & .Cells(i, iEndCol - 1).Address

                    Case C_sAnaCol, C_sAnaTot

                        sFormula = .Cells(i, iEndCol - 1).Address & " / " & .Cells(iEndRow, iEndCol - 1).Address
                        sFormula2 = .Cells(i, iEndCol - 3).Address & " / " & .Cells(iEndRow, iEndCol - 1).Address


                End Select

                With .Cells(i, iEndCol)
                    .Style = "Percent"
                    .NumberFormat = "0.00%"
                    If sFormula <> vbNullString Then .formula = AddPercentage(sFormula)
                End With

                If includeMissing Then

                    sFormula = vbNullString
                    'Add Formula for missing
                    sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarRow, _
                                                OnTotal:=False, sConditionRow:=.Cells(i, iStartCol).Address, _
                                                sVarColumn:=sVarColumn, sConditionColumn:=Chr(34) & Chr(34))

                    If sFormula <> vbNullString Then .Cells(i, iEndCol - 3).FormulaArray = sFormula

                    iMissingColumn = iEndCol - 3

                    'Add the percentage on missing if there is one
                    With .Cells(i, iEndCol - 2)

                        .Style = "Percent"
                        .NumberFormat = "0.00%"
                        If sFormula2 <> vbNullString Then .formula = AddPercentage(sFormula2)

                    End With

                End If

            Else
                'There is no percentage here, just add formulas for total
                If sFormula <> vbNullString Then .Cells(i, iEndCol).FormulaArray = sFormula

                'There is missing but not percentage, add formulas for missing
                If includeMissing Then

                    sFormula = vbNullString

                    sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarRow, _
                                                OnTotal:=False, sConditionRow:=.Cells(i, iStartCol).Address, _
                                                sVarColumn:=sVarColumn, sConditionColumn:=Chr(34) & Chr(34), isFiltered:=True)

                    iMissingColumn = iEndCol - 1

                    If sFormula <> vbNullString Then .Cells(i, iEndCol - 1).FormulaArray = sFormula
                End If
            End If

            i = i + 1
        Loop

        'The EndRow, total Column formula (the right corner) -----------------------------------------------------------------------------------------------------

        sFormula = vbNullString

        Select Case ClearNonPrintableUnicode(sForm)

        Case "COUNT", "COUNT()", "N", "N()"

            sFormula = "= " & "SUM(" & .Cells(iStartRow + 2, iTotalColumn).Address & ":" & .Cells(iEndRow - 1, iTotalColumn).Address & ")"

        Case Else

            sFormula = AnalysisFormula(Wkb:=Wkb, sFormula:=sForm, sVariate:="none", isFiltered:=True)

        End Select

        If sFormula <> vbNullString Then .Cells(iEndRow, iTotalColumn).FormulaArray = sFormula

    'Missings row and columns, with total row and columns -----------------------------------------------------------------------

    Select Case sMiss

        Case C_sAnaRow

        'Missing row and total column
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarRow, _
                                        OnTotal:=True, sConditionRow:=Chr(34) & Chr(34), _
                                        sVarColumn:=sVarColumn, includeMissing:=False, isFiltered:=True)

            If sFormula <> vbNullString Then .Cells(iMissingRow, iTotalColumn).FormulaArray = sFormula

        Case C_sAnaCol

            'Missing column and total row
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarColumn, _
                                        OnTotal:=True, sConditionRow:=Chr(34) & Chr(34), _
                                        sVarColumn:=sVarRow, includeMissing:=False, isFiltered:=True)

             If sFormula <> vbNullString Then .Cells(iEndRow, iMissingColumn).FormulaArray = sFormula

        Case C_sAnaAll

            'Missing row and total column
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarRow, _
                                        OnTotal:=True, sConditionRow:=Chr(34) & Chr(34), _
                                        sVarColumn:=sVarColumn, includeMissing:=True, isFiltered:=True)

            If sFormula <> vbNullString Then .Cells(iMissingRow, iTotalColumn).FormulaArray = sFormula


             'Missing column and total row
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarColumn, _
                                        OnTotal:=True, sConditionRow:=Chr(34) & Chr(34), _
                                        sVarColumn:=sVarRow, includeMissing:=True, isFiltered:=True)

             If sFormula <> vbNullString Then .Cells(iEndRow, iMissingColumn).FormulaArray = sFormula

            'Missing row and missing column
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sVarRow, _
                                        OnTotal:=False, sConditionRow:=Chr(34) & Chr(34), _
                                        sVarColumn:=sVarColumn, sConditionColumn:=Chr(34) & Chr(34), isFiltered:=True)

            If sFormula <> vbNullString Then .Cells(iMissingRow, iMissingColumn).FormulaArray = sFormula

    End Select

    End With


End Sub

'Add missing for univariate analysis
Sub AddUANA(Wkb As Workbook, DictHeaders As BetterArray, _
            sSumFunc As String, sVar As String, _
            sPercent As String, _
            iRow As Long, iStartCol As Long, iEndCol As Long, _
            Optional sInteriorColor As String = "VeryLightGreyBlue", _
            Optional sFontColor As String = "GreyBlue", _
            Optional sNumberFormat As String = "0.00")

    Dim Wksh As Worksheet
    Dim sFormula As String
    Dim sCond As String

    Set Wksh = Wkb.Worksheets(sParamSheetAnalysis)

    'Condition for missing is ""
    sCond = Chr(34) & Chr(34)

    With Wksh

        .Cells(iRow, iStartCol).Value = TranslateLLMsg("MSG_NA")

        FormatARange .Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol)), sFontColor:=sFontColor, _
        sInteriorColor:=sInteriorColor, FontSize:=C_iAnalysisFontSize - 1, isBold:=True, _
        Horiz:=xlHAlignRight

        .Cells(iRow, iStartCol).HorizontalAlignment = xlHAlignLeft

        On Error Resume Next

        sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, _
                                     sForm:=sSumFunc, sVar:=sVar, _
                                     sCondition:=sCond, isFiltered:=True)

        If sFormula <> vbNullString And Len(sFormula) < 255 Then .Cells(iRow, iStartCol + 1).FormulaArray = sFormula

        'Add the percentage
        If sPercent = C_sYes Then
            sFormula = .Cells(iRow, iStartCol + 1).Address & "/" & .Cells(iRow + 1, iStartCol + 1).Address
            sFormula = AddPercentage(sFormula)
            With .Cells(iRow, iEndCol)
                .Style = "Percent"
                .NumberFormat = "0.00%"
                .formula = sFormula
            End With
        End If
        On Error GoTo 0

    End With
End Sub

'Add total for univariate Analysis
Sub AddUATotal(Wkb As Workbook, DictHeaders As BetterArray, sSumFunc As String, sVar As String, sPercent As String, _
               sMiss As String, iRow As Long, iStartCol As Long, iEndCol As Long, _
               Optional sInteriorColor As String = "VeryLightGreyBlue")

    Dim Wksh As Worksheet
    Dim sFormula As String
    Dim sCond As String
    Dim includeMissing As Boolean

    Set Wksh = Wkb.Worksheets(sParamSheetAnalysis)
    sCond = Chr(34) & Chr(34)
    includeMissing = (sMiss = C_sYes)

    With Wksh
        .Cells(iRow, iStartCol).Value = TranslateLLMsg("MSG_Total")

        WriteBorderLines .Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol)), _
        iWeight:=xlHairline, sColor:="DarkBlue"

        FormatARange rng:=.Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol)), isBold:=True, sInteriorColor:=sInteriorColor, _
        FontSize:=C_iAnalysisFontSize + 1, Horiz:=xlHAlignRight

        .Cells(iRow, iStartCol).HorizontalAlignment = xlHAlignLeft

        'Add percentage if required
        If sPercent = C_sYes Then
            sFormula = "=" & .Cells(iRow, iStartCol + 1).Address & "/" & .Cells(iRow, iStartCol + 1).Address
            With .Cells(iRow, iEndCol)
                .formula = sFormula
                .Style = "Percent"
                .NumberFormat = "0.00%"
            End With
        End If
        'Add Formulas for total
        On Error Resume Next

        sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sSumFunc, sVar:=sVar, _
                                     sCondition:=sCond, OnTotal:=True, includeMissing:=includeMissing)
        If sFormula <> vbNullString And Len(sFormula) < 255 Then .Cells(iRow, iStartCol + 1).FormulaArray = sFormula

        On Error GoTo 0

    End With
End Sub

'Add formulas for TimeSeries

Sub AddTimeSeriesFormula(Wkb As Workbook, DictHeaders As BetterArray, _
                         sForm As String, sTimeVar As String, sCondVar As String, _
                         iRow As Long, iStartCol As Long, iEndCol As Long, sPerc As String, _
                         sMiss As String)


    Dim sFirstTimeCond As String
    Dim sSecondTimeCond As String
    Dim sTotalCell As String
    Dim istep As Long
    Dim i As Long
    Dim iInnerEndCol As Long
    Dim includeMissing As Boolean
    Dim sFormula As String
    Dim sCondVal As String

    Dim rng As Range
    Dim Wksh As Worksheet

    Set Wksh = Wkb.Worksheets(sParamSheetTemporalAnalysis)


    With Wksh

        'includeMissing
        includeMissing = (sMiss = C_sYes)

        'Update the end column
        iInnerEndCol = iEndCol
        'update the step
        istep = 1

        If sPerc <> C_sNo Then
            istep = 2
            iInnerEndCol = iInnerEndCol - 1
        End If


        i = iStartCol

        sFirstTimeCond = .Cells(iRow + 2, C_eStartColumnAnalysis).Address(Rowabsolute:=False)
        sSecondTimeCond = .Cells(iRow + 2, C_eStartColumnAnalysis + 1).Address(Rowabsolute:=False)


        Do While (i <= iInnerEndCol)
            sCondVal = .Cells(iRow, i).Address
            sFormula = TimeSeriesFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sTimeVar:=sTimeVar, _
                                         sFirstTimeCond:=sFirstTimeCond, sSecondTimeCond:=sSecondTimeCond, _
                                         sCondVar:=sCondVar, sCondVal:=sCondVal, _
                                         isFiltered:=True)


            If sFormula <> vbNullString Then
                .Cells(iRow + 2, i).FormulaArray = sFormula
                Set rng = .Range(.Cells(iRow + 2, i), .Cells(iRow + 2 + C_iNbTime, i))
                .Cells(iRow + 2, i).AutoFill Destination:=rng, Type:=xlFillValues
            End If

            If sPerc <> C_sNo Then
                Select Case sPerc
                Case C_sAnaRow
                    sTotalCell = .Cells(iRow + 2, iEndCol - 1).Address(Rowabsolute:=False)
                Case C_sAnaCol
                    sTotalCell = .Cells(iRow + 4 + C_iNbTime, i).Address
                Case C_sAnaAll
                    sTotalCell = .Cells(iRow + 4 + C_iNbTime, iEndCol - 1).Address
                End Select

                sFormula = .Cells(iRow + 2, i).Address(Rowabsolute:=False) & "/" & sTotalCell
                .Cells(iRow + 2, i + 1).formula = AddPercentage(sFormula)
                Set rng = .Range(.Cells(iRow + 2, i + 1), .Cells(iRow + 4 + C_iNbTime, i + 1))
                .Cells(iRow + 2, i + 1).AutoFill Destination:=rng, Type:=xlFillValues
                rng.NumberFormat = "0.00 %"
            End If


            'Missing row
            sFormula = BivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVarRow:=sTimeVar, _
                                        sConditionRow:=Chr(34) & Chr(34), _
                                        sVarColumn:=sCondVar, sConditionColumn:=sCondVal)

            If sFormula <> vbNullString Then .Cells(iRow + 3 + C_iNbTime, i).FormulaArray = sFormula

            'Total Row
            sFormula = UnivariateFormula(Wkb, DictHeaders, sForm, sVar:=sCondVar, sCondition:=sCondVal, isFiltered:=True)

            If sFormula <> vbNullString Then .Cells(iRow + 4 + C_iNbTime, i).FormulaArray = sFormula

            i = i + istep
        Loop

        'Missing column
        If sMiss = C_sYes Then


        End If

        'Total column
        sFormula = TimeSeriesFormula(Wkb, DictHeaders, sForm, sTimeVar, sFirstTimeCond, sSecondTimeCond, _
                                     OnTotal:=True, includeMissing:=includeMissing, sCondVar:=sCondVar)


        If sFormula <> vbNullString Then .Cells(iRow + 2, iInnerEndCol).FormulaArray = sFormula
        Set rng = .Range(.Cells(iRow + 2, iInnerEndCol), .Cells(iRow + 4 + C_iNbTime, iInnerEndCol))
        .Cells(iRow + 2, iInnerEndCol).AutoFill Destination:=rng, Type:=xlFillValues

         'Missing Row and Total column

         sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVar:=sTimeVar, _
                                        sCondition:=Chr(34) & Chr(34))

         If sFormula <> vbNullString Then .Cells(iRow + 3 + C_iNbTime, iInnerEndCol).FormulaArray = sFormula

         'Two total columns
         sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sForm, sVar:=sCondVar, isFiltered:=True, OnTotal:=True)
         If sFormula <> vbNullString Then .Cells(iRow + 4 + C_iNbTime, iInnerEndCol).FormulaArray = sFormula

    End With
End Sub

'Format one line for univariate analysis
Sub FormatCell(Wksh As Worksheet, iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long, _
               sPercent As String, _
               Optional sNumberFormat As String = "0.00", _
               Optional sInteriorColor As String = "VeryLightBlue", _
               Optional sFontColor As String = "DarkBlue")


    Dim sFormula As String

    With Wksh
        'Write the lines for each cells
        With .Cells(iStartRow, iStartCol)
            .Interior.Color = Helpers.GetColor(sInteriorColor)
            .Font.Color = Helpers.GetColor(sFontColor)
        End With

        WriteBorderLines .Range(.Cells(iStartRow, iStartCol), .Cells(iStartRow, iEndCol)), iWeight:=xlHairline, sColor:=sFontColor
        'Add the percentage values
        If sPercent = C_sYes Then
            sFormula = .Cells(iStartRow, iStartCol + 1).Address & "/" & .Cells(iEndRow, iStartCol + 1).Address
            sFormula = AddPercentage(sFormula)
            With .Cells(iStartRow, iEndCol)
                .Style = "Percent"
                .NumberFormat = "0.00%"
                .formula = sFormula
            End With
        End If
        'Before the total columns, double lines
        With .Range(.Cells(iEndRow - 1, iStartCol), .Cells(iEndRow - 1, iEndCol))
            With .Borders(xlEdgeBottom)
                .Weight = xlThin
                .LineStyle = xlDouble
                .Color = Helpers.GetColor(sFontColor)
            End With
        End With
    End With
End Sub

'Add formulas for univariate analysis
Function UnivariateFormula(Wkb As Workbook, DictHeaders As BetterArray, _
                           sForm As String, sVar As String, _
                           Optional sCondition As String = "", _
                           Optional isFiltered As Boolean = True, _
                           Optional OnTotal As Boolean = False, _
                           Optional includeMissing As Boolean = False) As String
    Dim sFormula As String

    sFormula = ""

    Select Case Application.WorksheetFunction.Trim(sForm)

    Case "COUNT", "COUNT()", "N", "N()"

        sFormula = AnalysisCount(Wkb, DictHeaders, sVarName:=sVar, sValue:=sCondition, isFiltered:=isFiltered, OnTotal:=OnTotal, includeMissing:=includeMissing)

    Case "SUM", "SUM()"

    Case Else
        If OnTotal And Not includeMissing Then
            sFormula = AnalysisFormula(Wkb, sForm, isFiltered, _
                                       sVariate:="univariate total not missing", sFirstCondVar:=sVar, _
                                       sFirstCondVal:=sCondition)
        ElseIf OnTotal Then
            sFormula = AnalysisFormula(Wkb, sForm, isFiltered, sVariate:="none")
        Else
            sFormula = AnalysisFormula(Wkb, sForm, isFiltered, _
                                       sVariate:="univariate", sFirstCondVar:=sVar, _
                                       sFirstCondVal:=sCondition)
        End If
    End Select

    If sFormula <> vbNullString And Len(sFormula) < 255 Then UnivariateFormula = sFormula
End Function

'Add formulas for bivariate analysis
Function BivariateFormula(Wkb As Workbook, DictHeaders As BetterArray, _
                          sForm As String, sVarRow As String, sVarColumn As String, _
                          Optional sConditionRow As String = "", _
                          Optional sConditionColumn As String = "", _
                          Optional isFiltered As Boolean = True, _
                          Optional OnTotal As Boolean = False, _
                          Optional includeMissing As Boolean = False) As String
    Dim sFormula As String

    sFormula = vbNullString

    Select Case ClearNonPrintableUnicode(sForm)

    Case "COUNT", "COUNT()", "N", "N()"

        sFormula = AnalysisCount(Wkb, DictHeaders, sVarName:=sVarRow, sValue:=sConditionRow, _
                                 sVarName2:=sVarColumn, sValue2:=sConditionColumn, _
                                 isFiltered:=isFiltered, OnTotal:=OnTotal, _
                                 includeMissing:=includeMissing)

    Case "SUM", "SUM()"

    Case Else
        'Working on total (with or without missing)

        If OnTotal And Not includeMissing Then

            sFormula = AnalysisFormula(Wkb, sForm, isFiltered:=isFiltered, _
                                       sVariate:="bivariate total not missing", sFirstCondVar:=sVarRow, _
                                       sFirstCondVal:=sConditionRow, _
                                       sSecondCondVar:=sVarColumn)

        ElseIf OnTotal Then
            'If required, write total on
            sFormula = AnalysisFormula(Wkb, sForm, isFiltered:=isFiltered, _
                                       sVariate:="univariate", sFirstCondVar:=sVarRow, _
                                       sFirstCondVal:=sConditionRow)

        Else

            sFormula = AnalysisFormula(Wkb, sForm, isFiltered:=isFiltered, _
                                       sVariate:="bivariate", sFirstCondVar:=sVarRow, _
                                       sFirstCondVal:=sConditionRow, sSecondCondVar:=sVarColumn, _
                                       sSecondCondVal:=sConditionColumn)

        End If
    End Select

    If sFormula <> vbNullString And Len(sFormula) < 255 Then BivariateFormula = sFormula
End Function

'Add formulas for time series
Function TimeSeriesFormula(Wkb As Workbook, DictHeaders As BetterArray, _
                           sForm As String, sTimeVar As String, _
                           sFirstTimeCond As String, sSecondTimeCond As String, _
                           Optional sCondVar As String, _
                           Optional sCondVal As String, _
                           Optional isFiltered As Boolean = True, _
                           Optional OnTotal As Boolean = False, _
                           Optional includeMissing As Boolean = False) As String
    Dim sFormula As String

    sFormula = vbNullString

    If sCondVar = vbNullString Then

        Select Case ClearNonPrintableUnicode(sForm)

        Case "COUNT", "COUNT()", "N", "N()"
            sFormula = TimeSeriesCount(Wkb, DictHeaders, sVarName:=sTimeVar, sValue1:=sFirstTimeCond, _
                                       sValue2:=sSecondTimeCond, isFiltered:=isFiltered)
        Case "SUM", "SUM()"

        Case Else
            sFormula = AnalysisFormula(Wkb, sForm, isFiltered:=isFiltered, sVariate:="bivariate date unique", _
                                       sSecondCondVar:=sTimeVar, sSecondCondVal:=sFirstTimeCond, _
                                       sThirdCondVal:=sSecondTimeCond)
        End Select

    Else

        Select Case ClearNonPrintableUnicode(sForm)


        Case "COUNT", "COUNT()", "N", "N()"

            sFormula = TimeSeriesCount(Wkb, DictHeaders, sVarName:=sTimeVar, sValue1:=sFirstTimeCond, sValue2:=sSecondTimeCond, _
                                       isFiltered:=isFiltered, sFirstCondVar:=sCondVar, sFirstCondVal:=sCondVal, OnTotal:=OnTotal, _
                                       includeMissing:=includeMissing)


        Case "SUM", "SUM()"


        Case Else

            If OnTotal And Not includeMissing Then

                sFormula = AnalysisFormula(Wkb, sForm, isFiltered:=isFiltered, sVariate:="bivariate date not missing", _
                                           sFirstCondVar:=sCondVar, sSecondCondVar:=sTimeVar, _
                                           sSecondCondVal:=sFirstTimeCond, sThirdCondVal:=sSecondTimeCond)
            ElseIf OnTotal Then

                sFormula = AnalysisFormula(Wkb, sForm, isFiltered:=isFiltered, sVariate:="bivariate date unique", _
                                           sSecondCondVar:=sTimeVar, sSecondCondVal:=sFirstTimeCond, _
                                           sThirdCondVal:=sSecondTimeCond)
            Else

                sFormula = AnalysisFormula(Wkb, sForm, isFiltered:=isFiltered, sVariate:="bivariate date", _
                                           sFirstCondVar:=sCondVar, sFirstCondVal:=sCondVal, sSecondCondVar:=sTimeVar, _
                                           sSecondCondVal:=sFirstTimeCond, sThirdCondVal:=sSecondTimeCond)

            End If

        End Select
    End If
    If sFormula <> vbNullString And Len(sFormula) < 255 Then TimeSeriesFormula = sFormula
End Function

'FUNCTIONS USED TO BUILD TIME SERIES TABLES ===================================================================================================================

Sub AddTimeColumn(Wksh As Worksheet, iStartRow As Long, iCol As Long, _
                  Optional sInteriorColor As String = "VeryLightBlue", _
                  Optional sFontColor As String = "DarkBlue", _
                  Optional sSelectionFontColor As String = "GreyBlue", _
                  Optional sSelectionInteriorColor As String = "VeryLightGreyBlue")

    Dim rng As Range
    Dim iRow As Long
    Dim sAgg As String                           'Aggregate cell
    Dim sMax As String                           'Max Cell

    With Wksh

        'Start Date
        iRow = iStartRow + 2

        FormatARange .Cells(iRow, iCol + 2), isBold:=True, sFontColor:=sFontColor, Horiz:=xlHAlignCenter, _
        sValue:=TranslateLLMsg("MSG_StartDate")
        FormatARange .Cells(iRow + 1, iCol + 2), isBold:=True, sFontColor:=sSelectionFontColor, sInteriorColor:=sSelectionInteriorColor, _
        NumFormat:="dd/mm/yyyy"
        'Information on start date
        FormatARange .Cells(iRow + 2, iCol + 2), sFontColor:=sSelectionFontColor

        'Time Aggregation

        FormatARange .Cells(iRow, iCol + 4), isBold:=True, sFontColor:=sFontColor, Horiz:=xlHAlignLeft, sValue:=TranslateLLMsg("MSG_TimeUnit")
        FormatARange .Cells(iRow + 1, iCol + 4), isBold:=True, sFontColor:=sSelectionFontColor, sInteriorColor:=sSelectionInteriorColor, sValue:=TranslateLLMsg("MSG_Day")
        'Aggregate address
        sAgg = .Cells(iRow + 1, iCol + 4).Address
        'Add validation for time aggregation
        SetValidation .Cells(iRow + 1, iCol + 4), "=" & C_sTimeAgg, 1, TranslateLLMsg("MSG_UnableToAgg")

        'End Date
        FormatARange .Cells(iRow, iCol + 6), isBold:=True, sFontColor:=sFontColor, Horiz:=xlHAlignCenter, _
        sValue:=TranslateLLMsg("MSG_EndDate")
        FormatARange .Cells(iRow + 1, iCol + 6), isBold:=True, sFontColor:=sSelectionFontColor, sInteriorColor:=sSelectionInteriorColor, _
        NumFormat:="dd/mm/yyyy"


        'Minimum date of the analysis
        FormatARange .Cells(iRow, iCol + 8), isBold:=False, sFontColor:=sFontColor, Horiz:=xlHAlignLeft, _
        sValue:=TranslateLLMsg("MSG_MinData"), FontSize:=C_iAnalysisFontSize - 2
        FormatARange .Cells(iRow + 1, iCol + 8), isBold:=False, sFontColor:=sSelectionFontColor, sInteriorColor:=sSelectionInteriorColor, _
        NumFormat:="dd/mm/yyyy", FontSize:=C_iAnalysisFontSize - 2
        .Cells(iRow, iCol + 8).Locked = True


        'Maximum date of the data
        FormatARange .Cells(iRow, iCol + 10), isBold:=False, sFontColor:=sFontColor, Horiz:=xlHAlignLeft, _
        sValue:=TranslateLLMsg("MSG_MaxData"), FontSize:=C_iAnalysisFontSize - 2
        FormatARange .Cells(iRow + 1, iCol + 10), isBold:=False, sFontColor:=sSelectionFontColor, sInteriorColor:=sSelectionInteriorColor, _
        NumFormat:="dd/mm/yyyy", FontSize:=C_iAnalysisFontSize - 2
        .Cells(iRow + 1, iCol + 10).Locked = True

        'Maximum address to be used elsewhere
        sMax = .Cells(iRow + 1, iCol + 10).Address

        'Start Date
        iRow = iRow + 4

        FormatARange .Cells(iRow, iCol + 1), isBold:=False, sFontColor:=sSelectionFontColor, _
                    sInteriorColor:=sSelectionInteriorColor, _
                    NumFormat:="dd/mm/yyyy"

        .Cells(iRow, iCol + 1).Locked = True
        .Cells(iRow, iCol + 1).formula = "=" & "MIN(MAX(" & .Cells(iRow - 2, iCol + 4).Address & "," & _
             .Cells(iRow - 2, iCol + 7).Address & ")," & sMax & ")"


        'The table for the time values
        iRow = iRow + 5
        .Cells(iRow - 1, iCol).Value = TranslateLLMsg("MSG_Period")
        .Cells(iRow, iCol - 2).formula = "= " & .Cells(iRow - 5, iCol + 1).Address
        .Cells(iRow, iCol - 1).formula = "= " & "FindLastDay(" & .Cells(iRow - 7, iCol + 1).Address & ", " & .Cells(iRow, iCol - 2).Address & ")"

        'Next row for autofill
        .Cells(iRow + 1, iCol - 2).formula = "= " & .Cells(iRow, iCol - 1).Address(Rowabsolute:=False, ColumnAbsolute:=False) & "+ 1"
        .Cells(iRow + 1, iCol - 1).formula = "= " & "FindLastDay(" & sAgg & ", " _
                                           & .Cells(iRow + 1, iCol - 2).Address(Rowabsolute:=False, ColumnAbsolute:=False) & ")"

        'Autofill column - 1
        Set rng = .Range(.Cells(iRow + 1, iCol - 1), .Cells(iRow + C_iNbTime, iCol - 1))
        .Cells(iRow + 1, iCol - 1).AutoFill rng

        'Autofill column - 2
        Set rng = .Range(.Cells(iRow + 1, iCol - 2), .Cells(iRow + C_iNbTime, iCol - 2))
        .Cells(iRow + 1, iCol - 2).AutoFill rng

        'Format and AutoFill the Range of values
        .Cells(iRow, iCol).formula = "= " & "FormatDateFromLastDay(" & sAgg & ", " & _
                                     .Cells(iRow, iCol - 1).Address(Rowabsolute:=False, ColumnAbsolute:=False) & "," & sMax & "," & _
                                      .Cells(iRow, iCol - 2).Address(Rowabsolute:=False, ColumnAbsolute:=False) & ")"

        'Format the range of time span (from, to)
        .Cells(iRow - 1, iCol - 2).Value = TranslateLLMsg("MSG_From")
        .Cells(iRow - 1, iCol - 1).Value = TranslateLLMsg("MSG_To")
        Set rng = .Range(.Cells(iRow - 1, iCol - 2), .Cells(iRow + C_iNbTime, iCol - 1))

        'Put the range in white
        FormatARange rng, sFontColor:=vbWhite, NumFormat:="dd-mm-yyyy", FontSize:=10
        rng.Locked = True

        'Format the range for period (with labels)
        Set rng = .Range(.Cells(iRow, iCol), .Cells(iRow + C_iNbTime, iCol))
        .Cells(iRow, iCol).AutoFill rng

        FormatARange rng, sInteriorColor:=sInteriorColor, sFontColor:=sFontColor, isBold:=True
        DrawLines rng, sColor:=sFontColor

        Set rng = .Range(.Cells(iRow - 2, iCol), .Cells(iRow + C_iNbTime + 2, iCol))
        WriteBorderLines rng, sColor:=sFontColor, iWeight:=xlMedium

    End With

End Sub

'FUNCTIONS USED TO BUILD FORMULAS ==============================================================================================================================

'Transform one formula to a formula for analysis.
'Wkb is a workbook where we can find the dictionary, the special character
'data and the name of all 'friendly' functions

Public Function AnalysisFormula(Wkb As Workbook, sFormula As String, _
                                Optional isFiltered As Boolean = False, _
                                Optional sVariate As String = "none", _
                                Optional sFirstCondVar As String = "__all", _
                                Optional sFirstCondVal As String = "__all", _
                                Optional sSecondCondVar As String = "__all", _
                                Optional sSecondCondVal As String = "__all", _
                                Optional sThirdCondVal As String = "__all", _
                                Optional Equal As String = "=") As String



    'Returns a string of cleared formula

    AnalysisFormula = vbNullString

    Dim sFormulaATest As String                  'same formula, with all the spaces replaced with
    Dim sAlphaValue As String                    'Alpha numeric values in a formula
    Dim sLetter As String                        'counter for every letter in one formula

    Dim FormulaAlphaData As BetterArray          'Table of alphanumeric data in one formula
    Dim FormulaData      As BetterArray
    Dim VarNameData  As BetterArray              'List of all variable names
    Dim SpecCharData As BetterArray              'List of Special Characters data
    Dim DictHeaders As BetterArray
    Dim TableNameData As BetterArray

    Dim i As Long
    Dim iPrevBreak As Long
    Dim iNbParentO As Long                       'Number of left parenthesis
    Dim iNbParentF As Long                       'Number of right parenthesis
    Dim icolNumb As Long                         'Column number on one sheet of one column used in a formual

    Dim isError As Boolean
    Dim OpenedQuotes As Boolean                  'Test if the formula has opened some quotes
    Dim QuotedCharacter As Boolean

    Set FormulaAlphaData = New BetterArray       'Alphanumeric values of one formula
    Set FormulaData = New BetterArray
    Set VarNameData = New BetterArray            'The list of all Variable Names
    Set SpecCharData = New BetterArray           'The list of all special characters
    Set DictHeaders = New BetterArray
    Set TableNameData = New BetterArray



    FormulaAlphaData.LowerBound = 1
    VarNameData.LowerBound = 1
    SpecCharData.LowerBound = 1
    DictHeaders.LowerBound = 1

    If sFormula = vbNullString Then Exit Function

    'squish the formula (removing multiple spaces) to avoid problems related to
    'space collapsing and upper/lower cases

    sFormulaATest = "(" & ClearNonPrintableUnicode(sFormula) & ")"


    'Initialisations:

    iNbParentO = 0                               'Number of open brakets
    iNbParentF = 0                               'Number of closed brackets
    iPrevBreak = 1
    OpenedQuotes = False
    QuotedCharacter = False

    i = 1

    Set DictHeaders = GetHeaders(Wkb, C_sParamSheetDict, 1)

    VarNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, 1), DetectLastColumn:=False, DetectLastRow:=True
    FormulaData.FromExcelRange Wkb.Worksheets(C_sSheetFormulas).ListObjects(C_sTabExcelFunctions).ListColumns("ENG").DataBodyRange, DetectLastColumn:=False
    SpecCharData.FromExcelRange Wkb.Worksheets(C_sSheetFormulas).ListObjects(C_sTabASCII).ListColumns("TEXT").DataBodyRange, DetectLastColumn:=False

    'Test if you have variable name in the dictionary

    If DictHeaders.IndexOf(C_sDictHeaderTableName) < 0 Then
        Exit Function
    End If


    TableNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, DictHeaders.IndexOf(C_sDictHeaderTableName)), DetectLastColumn:=False, DetectLastRow:=True

    If VarNameData.Includes(sFormulaATest) Then
        AnalysisFormula = ""                     'We have to aggregate
        Exit Function
    Else

        Do While (i <= Len(sFormulaATest))

            QuotedCharacter = False
            sLetter = Mid(sFormulaATest, i, 1)

            If sLetter = Chr(34) Then
                OpenedQuotes = Not OpenedQuotes
            End If

            If Not OpenedQuotes And SpecCharData.Includes(sLetter) Then 'A special character, not in quotes

                If sLetter = Chr(40) Then
                    iNbParentO = iNbParentO + 1
                End If

                If sLetter = Chr(41) Then
                    iNbParentF = iNbParentF + 1
                End If

                sAlphaValue = Application.WorksheetFunction.Trim(Mid(sFormulaATest, iPrevBreak, i - iPrevBreak))

                If sAlphaValue <> "" Then
                    'It is either a formula or a variable name or a quoted string
                    If Not VarNameData.Includes(sAlphaValue) And Not FormulaData.Includes(UCase(sAlphaValue)) And Not IsNumeric(sAlphaValue) Then

                        'Testing if not opened the quotes
                        If Mid(sAlphaValue, 1, 1) <> Chr(34) Then
                            isError = True
                            Exit Do
                        Else
                            QuotedCharacter = True
                        End If
                    End If

                    If Not isError And Not QuotedCharacter Then

                        'It is either a variable name or a formula
                        If VarNameData.Includes(sAlphaValue) Then 'It is a variable name, I will track its column
                            icolNumb = VarNameData.IndexOf(sAlphaValue)

                            sAlphaValue = BuildVariateFormula(TableNameData.Item(icolNumb), VarNameData.Item(icolNumb), _
                                                              sVariate, sFirstCondVar, sFirstCondVal, sSecondCondVar, sSecondCondVal, _
                                                              sThirdCondVal, isFiltered:=isFiltered)

                        End If

                    End If

                    FormulaAlphaData.Push sAlphaValue, sLetter
                Else

                    'I have a special character, at the value sLetter But nothing between this special character and previous one, just add it
                    FormulaAlphaData.Push sLetter
                End If
                iPrevBreak = i + 1
            End If

            i = i + 1
        Loop
    End If

    If iNbParentO <> iNbParentF Then
        isError = True
    End If

    If Not isError Then
        sAlphaValue = FormulaAlphaData.ToString(Separator:="", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
        AnalysisFormula = Equal & sAlphaValue
    Else

        'MsgBox "Error in analysis formula: " & sFormula

    End If
End Function

'Change / Or adapt the formula for univariate analysis, bivariate analysis or For just summary part

Function BuildVariateFormula(sTableName As String, _
                             sVarName As String, _
                             Optional sVariate As String = "none", _
                             Optional sFirstCondVar As String = "__all", _
                             Optional sFirstCondVal As String = "__all", _
                             Optional sSecondCondVar As String = "__all", _
                             Optional sSecondCondVal As String = "__all", _
                             Optional sThirdCondVal As String = "__all", _
                             Optional isFiltered As Boolean = False) As String



    BuildVariateFormula = vbNullString



    Dim sTable As String                         'The name of the table depends on the fact that we want to filter or not

    Dim sAlphaValue As String


    sTable = sTableName

    If isFiltered Then sTable = C_sFiltered & sTableName

    'Fall back to none if you don't precise the univariate / bivariate values: Those are safeguard

    If (sVariate = "univariate") And _
                                 (sFirstCondVar = "__all" Or sFirstCondVal = "__all") Then sVariate = "none"

    If (sVariate = "bivariate") And _
                                (sFirstCondVar = "__all" Or sFirstCondVal = "__all" Or sSecondCondVar = "__all" Or sSecondCondVal = "__all") Then sVariate = "none"



    Select Case sVariate

    Case "none"

        sAlphaValue = sTable & "[" & sVarName & "]"

    Case "univariate"

        sAlphaValue = "IF(" & sTable & "[" & sFirstCondVar & "]" & "=" _
                    & sFirstCondVal & ", " _
                    & sTable & "[" & sVarName & "]" & ")"

    Case "univariate total not missing"

        sAlphaValue = "IF(" & sTable & "[" & sFirstCondVar & "]" & "<>" _
                    & Chr(34) & Chr(34) & ", " _
                    & sTable & "[" & sVarName & "]" & ")"

    Case "bivariate"

        sAlphaValue = "IF( ((" & sTable & "[" & sFirstCondVar & "]" & "=" _
                    & sFirstCondVal & ") * (" _
                    & sTable & "[" & sSecondCondVar & "]" & "=" _
                    & sSecondCondVal & ")), " _
                    & sTable & "[" & sVarName & "]" & ")"

    Case "bivariate total not missing"

        sAlphaValue = "IF( ((" & sTable & "[" & sFirstCondVar & "]" & "=" _
                    & sFirstCondVal & ") * (" _
                    & sTable & "[" & sSecondCondVar & "]" & "<>" _
                    & Chr(34) & Chr(34) & ")), " _
                    & sTable & "[" & sVarName & "]" & ")"

    Case "bivariate date"

        sAlphaValue = "IF( ((" & sTable & "[" & sFirstCondVar & "]" & "=" & _
                      sFirstCondVal & ") * (" _
                    & sTable & "[" & sSecondCondVar & "]" & " >= " & _
                      sSecondCondVal & ") * (" & sTable & "[" & sSecondCondVar & "]" & " <= " & _
                      sThirdCondVal & ")), " & sTable & "[" & sVarName & "]" & ")"

    Case "bivariate date not missing"

        sAlphaValue = "IF( ((" & sTable & "[" & sFirstCondVar & "]" & " <> " & _
                      Chr(34) & Chr(34) & ") * (" & _
                      sTable & "[" & sSecondCondVar & "]" & " >= " & _
                      sSecondCondVal & ") * (" & sTable & "[" & sSecondCondVar & "]" & " <= " & _
                      sThirdCondVal & ")), " & sTable & "[" & sVarName & "]" & ")"

    Case "bivariate date unique"
        sAlphaValue = "IF( ((" & sTable & "[" & sSecondCondVar & "]" & " >= " & _
                      sSecondCondVal & ") * (" & sTable & "[" & sSecondCondVar & "]" & " <= " & _
                      sThirdCondVal & ")), " & sTable & "[" & sVarName & "]" & ")"

        'By default fall back to simple varname in a table
    Case Else
        sAlphaValue = sTable & "[" & sVarName & "]"
    End Select

    BuildVariateFormula = sAlphaValue

End Function

'Analysis Count

Function AnalysisCount(Wkb As Workbook, DictHeaders As BetterArray, sVarName As String, sValue As String, _
                       Optional sVarName2 As String = "", Optional sValue2 As String = "", Optional isFiltered As Boolean = False, _
                       Optional OnTotal As Boolean = False, _
                       Optional includeMissing As Boolean = False) As String



    Dim VarNameData As BetterArray
    Dim TableNameData As BetterArray
    Dim sTable As String
    Dim sTable2 As String
    Dim sFormula As String

    Set VarNameData = New BetterArray
    Set TableNameData = New BetterArray



    VarNameData.LowerBound = 1
    TableNameData.LowerBound = 1

    VarNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, 1), DetectLastColumn:=False, DetectLastRow:=True
    TableNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, DictHeaders.IndexOf(C_sDictHeaderTableName)), _
        DetectLastColumn:=False, DetectLastRow:=True
    sFormula = vbNullString

    If sVarName2 = vbNullString Then
        'Only one variable, just proceed as before

        If VarNameData.Includes(sVarName) Then

            sTable = TableNameData.Items(VarNameData.IndexOf(sVarName))
            If isFiltered Then sTable = C_sFiltered & sTable

            sFormula = "COUNTIF" & "(" & sTable & "[" & sVarName & "], " & sValue & ")"

            If OnTotal And includeMissing Then
                sFormula = "COUNTA" & "(" & sTable & "[" & sVarName & "]" & ")" & " + " & "COUNTBLANK" & "(" & sTable & "[" & sVarName & "]" & ")"

            ElseIf OnTotal Then
                sFormula = "COUNTA" & "(" & sTable & "[" & sVarName & "]" & ")"
            End If
        End If

    Else

        If VarNameData.Includes(sVarName) And VarNameData.Includes(sVarName2) Then
            sTable = TableNameData.Items(VarNameData.IndexOf(sVarName))
            sTable2 = TableNameData.Items(VarNameData.IndexOf(sVarName2))

            If sTable2 <> sTable Then Exit Function 'Proceed only if variables are in the same Table
            If isFiltered Then sTable = C_sFiltered & sTable

            If OnTotal And Not includeMissing Then

                sFormula = "COUNTIFS" & "(" & sTable & "[" & sVarName & "], " & sValue & ", " & sTable & "[" & sVarName2 & "], " & Chr(34) & "<>" & Chr(34) & ")"

            ElseIf OnTotal Then

                sFormula = "COUNTIFS" & "(" & sTable & "[" & sVarName & "], " & sValue & ")"

            Else

                sFormula = "COUNTIFS" & "(" & sTable & "[" & sVarName & "], " & sValue & "," & sTable & "[" & sVarName2 & "], " & sValue2 & ")"

            End If

        End If
    End If


    AnalysisCount = "=" & sFormula

End Function

Function TimeSeriesCount(Wkb As Workbook, DictHeaders As BetterArray, sVarName As String, sValue1 As String, _
                         sValue2 As String, Optional isFiltered As Boolean = False, _
                         Optional sFirstCondVar As String = "", _
                         Optional sFirstCondVal As String = "", _
                         Optional OnTotal As Boolean = False, _
                         Optional includeMissing As Boolean = True) As String


    Dim VarNameData As BetterArray
    Dim TableNameData As BetterArray
    Dim sTable As String
    Dim sFormula As String

    Set VarNameData = New BetterArray
    Set TableNameData = New BetterArray

    VarNameData.LowerBound = 1
    TableNameData.LowerBound = 1

    VarNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, 1), DetectLastColumn:=False, DetectLastRow:=True
    TableNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, DictHeaders.IndexOf(C_sDictHeaderTableName)), _
        DetectLastColumn:=False, DetectLastRow:=True

    sFormula = vbNullString

    If VarNameData.Includes(sVarName) Then
        sTable = TableNameData.Items(VarNameData.IndexOf(sVarName))
        If isFiltered Then sTable = C_sFiltered & sTable

        If sFirstCondVar = vbNullString Or (OnTotal And includeMissing) Then

            sFormula = "= COUNTIFS" & "(" & sTable & "[" & sVarName & "]," & Chr(34) & ">=" & Chr(34) & "&" & sValue1 & _
                                                                                                      ", " & sTable & "[" & sVarName & "]," & Chr(34) & "<=" & Chr(34) & "&" & sValue2 & ")"
        Else
            If VarNameData.Includes(sFirstCondVar) Then

                'Total without missing

                If OnTotal And Not includeMissing Then

                    sFormula = "= COUNTIFS" & "(" & sTable & "[" & sFirstCondVar & "]" & ", " & _
                               Chr(34) & "<>" & Chr(34) & ", " & sTable & "[" & sVarName & "], " & Chr(34) & ">=" & Chr(34) & "&" & sValue1 & _
                               ", " & sTable & "[" & sVarName & "], " & Chr(34) & "<=" & Chr(34) & "&" & sValue2 & ")"

                Else

                    sFormula = "= COUNTIFS" & "(" & sTable & "[" & sFirstCondVar & "]" & ", " & _
                               sFirstCondVal & ", " & sTable & "[" & sVarName & "], " & Chr(34) & ">=" & Chr(34) & "&" & sValue1 & _
                               ", " & sTable & "[" & sVarName & "], " & Chr(34) & "<=" & Chr(34) & "&" & sValue2 & ")"
                End If
            End If
        End If

    End If

    If sFormula <> vbNullString And Len(sFormula) < 255 Then TimeSeriesCount = sFormula
End Function


'Add percentage taking in account eventual errors
Function AddPercentage(sForm As String) As String
    AddPercentage = "= IF(ISERR(" & sForm & ")," & Chr(34) & Chr(34) & "," & sForm & ")"
End Function


'Function to create a simple bar chart for the univariate analysis part
Public Sub CreateBarChart(Wksh As Worksheet, Left As Integer, Top As Integer, RngSource As Range, Optional chartType As Integer = xlColumnClustered)
    With Wksh
    Dim UAChart As ChartObject
        Set UAChart = .ChartObjects.Add(Left + 450, Top, 100, 180)

        UAChart.Chart.chartType = chartType

        'Add data to the graph
        UAChart.Chart.SeriesCollection.Add Source:=RngSource, RowCol:=xlColumns, SeriesLabels:=True, Categorylabels:=True

    End With

End Sub

