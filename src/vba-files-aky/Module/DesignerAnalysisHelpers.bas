Attribute VB_Name = "DesignerAnalysisHelpers"


Option Explicit




'FUNCTIONS USED TO BUILD UNIVARIATE ANALYSIS ===================================================================================================================

'Create New section
Sub CreateNewSection(Wksh As Worksheet, iRow As Long, iCol As Long, sSection As String, _
                                         Optional sColor As String = "DarkBlue")
        With Wksh
            'New range, format the range
            FormatARange .Cells(iRow, iCol), sValue:=sSection, FontSize:=C_iAnalysisFontSize + 4, _
                    sFontColor:=sColor, Horiz:=xlHAlignLeft

            'Draw a border arround the section
            DrawLines Rng:=Range(.Cells(iRow, iCol), .Cells(iRow, iCol + 6)), iWeight:=xlMedium, sColor:=sColor, At:="Bottom"
        End With
End Sub


'Create Headers for univariate analysis
Sub CreateUAHeaders(Wksh As Worksheet, iRow As Long, iCol As Long, _
                    sMainLab As String, sSummaryLabel As String, _
                    sPercent As String, Optional sColor As String = "DarkBlue")
    With Wksh
        'Variable Label from the dictionary
        FormatARange Rng:=.Cells(iRow, iCol), sValue:=sMainLab, sFontColor:=sColor, isBold:=True, Horiz:=xlHAlignLeft
       'First column on sumary label
        FormatARange Rng:=.Cells(iRow, iCol + 1), sValue:=sSummaryLabel, sFontColor:=sColor, isBold:=True
       'Add Percentage header column if required
        If sPercent = C_sYes Then FormatARange Rng:=.Cells(iRow, iCol + 2), sValue:=TranslateLLMsg("MSG_Percent"), sFontColor:=sColor, isBold:=True
    End With
End Sub

'Create Headers for bivariate Analysis

Sub CreateBAHeaders(Wksh As Worksheet, ColumnsData As BetterArray, _
                    RowsData As BetterArray, _
                    iRow As Long, iCol As Long, _
                    sMainLabRow As String, sMainLabCol As String, _
                    sSummaryLabel As String, _
                    sPercent As String, sMiss As String, _
                    Optional sInteriorColor As String = "VeryLightBlue", _
                    Optional sTotalInteriorColor As String = "VeryLightGreyBlue", _
                    Optional sNAFontColor As String = "GreyBlue", _
                    Optional sColor As String = "DarkBlue")
    Dim i As Long
    Dim iLastCol As Long
    Dim iTotalLastCol As Long
    Dim iTotalFirstCol As Long
    Dim iEndRow As Long
    Dim sArrow As String
    Dim HasPercent As Boolean


     With Wksh
        'Variable Label from the dictionary for Row
        FormatARange Rng:=.Cells(iRow + 2, iCol), sValue:=sMainLabRow, sFontColor:=sColor, isBold:=True

        'Merge the first and second rows of first column of bivariate analysis
        Range(.Cells(iRow + 1, iCol), .Cells(iRow + 2, iCol)).Merge
        .Cells(iRow + 1, iCol).MergeArea.HorizontalAlignment = xlHAlignLeft
        .Cells(iRow + 1, iCol).MergeArea.VerticalAlignment = xlVAlignCenter

        'Variable label from the dictionary for column
        FormatARange Rng:=.Cells(iRow, iCol + 1), sValue:=sMainLabCol, sFontColor:=sColor, isBold:=True, Horiz:=xlHAlignLeft

        'Add the rows Data --------------------------------------------------------------------------

        RowsData.ToExcelRange .Cells(iRow + 3, iCol)
        'EndRow of the table
        iEndRow = iRow + 2 + RowsData.Length

        FormatARange Rng:=Range(.Cells(iRow + 3, iCol), .Cells(iEndRow, iCol)), sFontColor:=sColor, sInteriorColor:=sInteriorColor, Horiz:=xlHAlignLeft

        If sMiss = C_sAnaRow Or sMiss = C_sAnaAll Then
            'Format the last row, just in case we need
            FormatARange Rng:=.Cells(iEndRow + 1, iCol), sFontColor:=sNAFontColor, sInteriorColor:=sTotalInteriorColor, _
                         Horiz:=xlHAlignLeft, sValue:=TranslateLLMsg("MSG_NA")
            iEndRow = iEndRow + 1
        End If

        'Now Add Percentage And Values for the column -----------------------------------------------

        'If you have to add percentage :
        sArrow = vbNullString

        Select Case sPercent

            Case C_sAnaCol
                HasPercent = True
                sArrow = ChrW(8597) 'Arrow is vertical
            Case C_sAnaRow
                HasPercent = True
                sArrow = ChrW(8596) 'Arrow is horizontal
            Case C_sAnaTot
                HasPercent = True
            Case Else
                HasPercent = False
        End Select

        If HasPercent Then

           i = 0

           Do While (i < ColumnsData.Length)
                'There is percentage, we have to add the percentage
                .Cells(iRow + 1, iCol + 2 * i + 1).value = ColumnsData.Items(i + 1)
                .Cells(iRow + 2, iCol + 2 * i + 1).value = sSummaryLabel
                .Cells(iRow + 2, iCol + 2 * i + 2).value = TranslateLLMsg("MSG_Percent") & " " & sArrow
                Range(.Cells(iRow + 1, iCol + 2 * i + 1), .Cells(iRow + 1, iCol + 2 * i + 2)).Merge

                'Write borders arround the different part of the columns
                DrawLines Rng:=Range(.Cells(iRow + 1, iCol + 2 * i + 1), .Cells(iEndRow + 1, iCol + 2 * i + 2)), sColor:=sColor
                DrawLines Rng:=Range(.Cells(iRow + 1, iCol + 2 * i + 1), .Cells(iEndRow + 1, iCol + 2 * i + 1)), At:="Left", iWeight:=xlThin, sColor:=sColor
                i = i + 1
           Loop

           iLastCol = 2 * i 'This is the last column of the table when there is percentage
        Else

            'There is no percentage, only column values
            i = 1

            Do While (i <= ColumnsData.Length)
                .Cells(iRow + 1, iCol + i).value = ColumnsData.Items(i)
                .Cells(iRow + 2, iCol + i).value = sSummaryLabel

               'Draw lines arround all borders
               DrawLines Rng:=Range(.Cells(iRow + 1, iCol + i), .Cells(iEndRow + 1, iCol + i)), sColor:=sColor
               DrawLines Rng:=Range(.Cells(iRow + 1, iCol + i), .Cells(iEndRow + 1, iCol + i)), At:="Left", sColor:=sColor, iWeight:=xlThin
                i = i + 1
            Loop

            iLastCol = i - 1 'Last column of the table without the percentage
        End If

        iLastCol = iCol + iLastCol

        iTotalFirstCol = iLastCol + 1

        'Add Missing for column -------------------------------------------------------------------------------
        If sMiss = C_sAnaCol Or sMiss = C_sAnaAll Then

            'Missing at the end of the column
            .Cells(iRow + 1, iTotalFirstCol).value = TranslateLLMsg("MSG_NA")
            .Cells(iRow + 2, iTotalFirstCol).value = sSummaryLabel

            iTotalFirstCol = iTotalFirstCol + 1

            'Add percentage
            If HasPercent Then
                .Cells(iRow + 2, iTotalFirstCol).value = TranslateLLMsg("MSG_Percent") & " " & sArrow
                Range(.Cells(iRow + 1, iTotalFirstCol - 1), .Cells(iRow + 1, iTotalFirstCol)).Merge

                'Now update the first column for total
                iTotalFirstCol = iTotalFirstCol + 1
            End If

            'Format the missing for column
            DrawLines Rng:=Range(.Cells(iRow + 1, iLastCol + 1), .Cells(iEndRow + 1, iTotalFirstCol - 1)), sColor:=sColor
            FormatARange Rng:=Range(.Cells(iRow + 1, iLastCol + 1), .Cells(iEndRow + 1, iTotalFirstCol - 1)), sInteriorColor:=sTotalInteriorColor, sFontColor:=sNAFontColor
        End If


        'Add Total ------------------------------------------------------------------------------------------------------------------------------------------------

        .Cells(iRow + 1, iTotalFirstCol).value = TranslateLLMsg("MSG_Total")
        .Cells(iRow + 2, iTotalFirstCol).value = sSummaryLabel

        iTotalLastCol = iTotalFirstCol
        'In case it is needed, add percentage for total also
        If HasPercent Then
            .Cells(iRow + 2, iTotalLastCol + 1).value = TranslateLLMsg("MSG_Percent") & " " & sArrow
            Range(.Cells(iRow + 1, iTotalLastCol), .Cells(iRow + 1, iTotalLastCol + 1)).Merge
            iTotalLastCol = iTotalLastCol + 1
        End If

        'Format total
        'Add hairlines between cells
        DrawLines Rng:=Range(.Cells(iRow + 1, iTotalFirstCol), .Cells(iEndRow + 1, iTotalLastCol)), sColor:=sColor
        'Add a left double line
        DrawLines Rng:=Range(.Cells(iRow + 1, iTotalFirstCol), .Cells(iEndRow + 1, iTotalFirstCol)), sColor:=sColor, iLine:=xlDouble, At:="Left"
        'Format all the total range
        FormatARange Rng:=Range(.Cells(iRow + 1, iTotalFirstCol), .Cells(iEndRow + 1, iTotalLastCol)), sInteriorColor:=sTotalInteriorColor, isBold:=True

        'Add Missing for Rows (After total because we need total end column) -------------------------------------------------------------------------------------------------

        If sMiss = C_sAnaRow Or sMiss = C_sAnaAll Then
            FormatARange Rng:=Range(.Cells(iEndRow, iCol + 1), .Cells(iEndRow, iTotalLastCol)), sInteriorColor:=sTotalInteriorColor, sFontColor:=sNAFontColor
        End If

        'Total on the Last line
        FormatARange Rng:=.Cells(iEndRow + 1, iCol), sInteriorColor:=sTotalInteriorColor, isBold:=True, _
                     Horiz:=xlHAlignLeft, sValue:=TranslateLLMsg("MSG_Total")
        FormatARange Rng:=Range(.Cells(iEndRow + 1, iCol + 1), .Cells(iEndRow + 1, iTotalLastCol)), sInteriorColor:=sTotalInteriorColor, isBold:=True

        'Format Table Headers -------------------------------------------------------------------------------------------------------------------------------------------

        'First row with column categories
        FormatARange Rng:=Range(.Cells(iRow + 1, iCol + 1), .Cells(iRow + 1, iLastCol)), sFontColor:=sColor, sInteriorColor:=sInteriorColor

        'Second row with summary label with/without percentage
        FormatARange Rng:=Range(.Cells(iRow + 2, iCol + 1), .Cells(iRow + 2, iLastCol)), sFontColor:=sColor, FontSize:=C_iAnalysisFontSize - 1

        'Draw lines arround the first column of table
        DrawLines Rng:=Range(.Cells(iRow + 1, iCol), .Cells(iEndRow + 1, iCol)), sColor:=sColor

        'Thick line at the header row
        DrawLines Rng:=Range(.Cells(iRow + 2, iCol), .Cells(iRow + 2, iTotalLastCol)), At:="Bottom", iLine:=xlDouble, sColor:=sColor

        'Draw lines for Total
        DrawLines Rng:=Range(.Cells(iEndRow + 1, iCol), .Cells(iEndRow + 1, iTotalLastCol)), sColor:=sColor, iLine:=xlDouble, At:="Top"

        'Drawlines arround all the table
        WriteBorderLines oRange:=Range(.Cells(iRow + 1, iCol), .Cells(iEndRow + 1, iTotalLastCol)), sColor:=sColor, iWeight:=xlThin

    End With
End Sub


'Add interior formulas for the bivariate analysis

Sub AddInnerFormula(Wkb as Workbook, DictHeaders As BetterArray, sForm As String, _
                    iStartRow As Long, iStartCol As Long, iEndRow As Long, iEndCol, _
                    sPercent As String, sMiss As String, sVarRow As String, sVarColumn As String)

    Dim Wksh As Worksheet
    Dim iInnerEndRow As Long
    Dim iInnerEndCol As Long
    Dim i As Long
    Dim j As Long
    Dim istep As Long
    Dim sFormula As String

    Set Wksh = Wkb.Worksheets(C_sSheetAnalysis)

    iInnerEndRow = iEndRow - 1
    'There is a missing line
    If sMiss = C_sAnaRow Or sMiss = C_sAnaAll Then iInnerEndRow = iEndRow - 2

    iInnerEndCol = iEndCol - 1
    'There is missing at the end column
    If sMiss = C_sAnaCol Or sMiss = C_sAnaAll Then iInnerEndCol = iInnerEndCol - 1

    istep = 1
    If sPercent  <> C_sNo Then 'There is precentage.
        iInnerEndCol = iInnerEndCol - 2
        istep = 2
    End If

    i = iStartRow + 2
    j = iStartCol + 1

    With Wksh
        'Add Now the formulas
        Do while( i <= iInnerEndRow)

            Do while(j <= iInnerEndCol)
                sFormula = BivariateFormula(Wkb := Wkb, DictHeaders := DictHeaders, sForm := sForm,  _
                                            sVarRow := sVarRow, sVarColumn := sVarColumn, _
                                            sValue := .Cells(i, iStartCol).Address, sValue2 = .Cells(iStartRow, j).Address, _
                                            isFiltered := True)
                j = j + istep
            Loop

            i = i + 1
        Loop
    End With
End Sub



'Add missing for univariate analysis
Sub AddUANA(Wkb As Workbook, DictHeaders As BetterArray, _
            sSumFunc As String, sVar As String, _
            iRow As Long, iStartCol As Long, iEndCol As Long, _
            Optional sInteriorColor As String = "VeryLightGreyBlue", _
            Optional sFontColor As String = "GreyBlue", _
            Optional sNumberFormat As String = "0.00")

        Dim Wksh As Worksheet
        Dim sFormula As String
        Dim sCond As String

        Set Wksh = Wkb.Worksheets(C_sSheetAnalysis)
        sCond = Chr(34) & Chr(34)

        With Wksh

            .Cells(iRow, iStartCol).value = TranslateLLMsg("MSG_NA")

            FormatARange .Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol)), sFontColor:=sFontColor, _
                    sInteriorColor:=sInteriorColor, FontSize:=C_iAnalysisFontSize - 1, isBold:=True, _
                    NumFormat:=sNumberFormat

            On Error Resume Next

            sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, _
                                            sForm:=sSumFunc, sVar:=sVar, _
                                            sCondition:=sCond, isFiltered:=True)

            If sFormula <> vbNullString And Len(sFormula) < 255 Then .Cells(iRow, iStartCol + 1).FormulaArray = sFormula

            On Error GoTo 0

        End With
End Sub


'Add Missing for Bivariate Analysis

Sub AddBANA(Wkb As Workbook, DictHeaders As BetterArray, _
            sSumFunc As String, sVar As String, _
            iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long, _
            Optional sInteriorColor As String = "VeryLightGreyBlue", _
            Optional sFontColor As String = "GreyBlue", _
            Optional sNumberFormat As String = "0.00")

        Dim Wksh As Worksheet
        Dim sFormula As String
        Dim sCond As String

        Set Wksh = Wkb.Worksheets(C_sSheetAnalysis)
        sCond = Chr(34) & Chr(34)

        'We are on Rows

        With Wksh
          On Error Resume Next

                sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, _
                                            sForm:=sSumFunc, sVar:=sVar, _
                                            sCondition:=sCond, isFiltered:=True)

            If sFormula <> vbNullString And Len(sFormula) < 255 Then .Cells(iRow, iStartCol + 1).FormulaArray = sFormula

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

        Set Wksh = Wkb.Worksheets(C_sSheetAnalysis)
        sCond = Chr(34) & Chr(34)
        includeMissing = (sMiss = C_sYes)

        With Wksh
            .Cells(iRow, iStartCol).value = TranslateLLMsg("MSG_Total")

            WriteBorderLines .Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol)), _
                iWeight:=xlHairline, sColor:="DarkBlue"

            FormatARange Rng:=.Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol)), isBold:=True, sInteriorColor:=sInteriorColor, _
                        FontSize:=C_iAnalysisFontSize + 1

            'Add percentage if required
            If sPercent = C_sYes Then
                sFormula = "=" & .Cells(iRow, iStartCol + 1).Address & "/" & .Cells(iRow, iStartCol + 1).Address
                With .Cells(iRow, iEndCol)
                    .Formula = sFormula
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

        With .Cells(iStartRow, iStartCol + 1)
            .NumberFormat = sNumberFormat
        End With

        WriteBorderLines .Range(.Cells(iStartRow, iStartCol), .Cells(iStartRow, iEndCol)), iWeight:=xlHairline, sColor:=sFontColor
        'Add the percentage values
        If sPercent = C_sYes Then
            sFormula = "=" & .Cells(iStartRow, iStartCol + 1).Address & "/" & .Cells(iEndRow, iStartCol + 1).Address
            With .Cells(iStartRow, iEndCol)
                .Style = "Percent"
                .NumberFormat = "0.00%"
                .Formula = sFormula
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

              sFormula = AnalysisCount(Wkb, DictHeaders, sVarName := sVar, sValue := sCondition, isFiltered := isFiltered, OnTotal := OnTotal, includeMissing := includeMissing)

          Case "SUM", "SUM()"

      Case Else
                If OnTotal And includeMissing Then
                sFormula = AnalysisFormula(Wkb, sForm, isFiltered, _
                                sVariate:="univariate total missing", sFirstCondVar:=sVar, _
                                sFirstCondVal:=sCondition)

                ElseIf OnTotal Then

                    sFormula = AnalysisFormula(Wkb, sForm, isFiltered, _
                                sVariate:="none")

                Else

                    sFormula = AnalysisFormula(Wkb, sForm, isFiltered, _
                                sVariate:="univariate", sFirstCondVar:=sVar, _
                                sFirstCondVal:=sCondition)

                End If
    End Select

        UnivariateFormula = sFormula
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

        sFormula = ""

        Select Case Application.WorksheetFunction.Trim(sForm)

        Case "COUNT", "COUNT()", "N", "N()"

              sFormula = AnalysisCount(Wkb, DictHeaders, sVarName := sVarRow, sValue:=sConditionRow, _
                                      sVarName2 := sVarColumn, sValue2 := sConditionColumn,
                                      isFiltered := isFiltered, OnTotal := OnTotal, _
                                      includeMissing := includeMissing)

        Case "SUM", "SUM()"

        Case Else
                If OnTotal And includeMissing Then
                sFormula = AnalysisFormula(Wkb, sForm, isFiltered := isFiltered, _
                                sVariate:="bivariate total missing", sFirstCondVar:=sVarRow, _
                                sFirstCondVal:=sConditionRow, _
                                sSecondCondVar := sVarColumn, _
                                sSecondCondVal := sConditionColumn)

                ElseIf OnTotal Then

                    sFormula = AnalysisFormula(Wkb, sForm, isFiltered := isFiltered, _
                                sVariate:="bivariate total", sFirstCondVar := sVarRow, _
                                sFirstCondVal := sConditionRow, sSecondCondVar := sVarColumn, _
                                sSecondCondVal := sConditionColumn)

                Else

                    sFormula = AnalysisFormula(Wkb, sForm, isFiltered := isFiltered, _
                                sVariate:="bivariate", sFirstCondVar:=sVarRow, _
                                sFirstCondVal:=sConditionRow, sSecondCondVar := sVarColumn, _
                                sSecondCondVal := sConditionColumn)

                End If
        End Select

        UnivariateFormula = sFormula
 End Function



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
                                Optional sSecondCondVal As String = "__all") As String



    'Returns a string of cleared formula

    AnalysisFormula = vbNullString

    Dim sFormulaATest As String                  'same formula, with all the spaces replaced with
    Dim sAlphaValue As String                    'Alpha numeric values in a formula
    Dim sLetter As String                        'counter for every letter in one formula
    Dim scolAddress As String                    'address of one column used in a formula

    Dim FormulaAlphaData As BetterArray          'Table of alphanumeric data in one formula
    Dim FormulaData      As BetterArray
    Dim VarNameData  As BetterArray              'List of all variable names
    Dim SpecCharData As BetterArray              'List of Special Characters data
    Dim DictHeaders As BetterArray
    Dim TableNameData As BetterArray
    Dim VarMainLabelData As BetterArray

    Dim i As Long
    Dim iPrevBreak As Long
    Dim iNbParentO As Long                    'Number of left parenthesis
    Dim iNbParentF As Long                    'Number of right parenthesis
    Dim icolNumb As Long                      'Column number on one sheet of one column used in a formual

    Dim isError As Boolean
    Dim OpenedQuotes As Boolean                  'Test if the formula has opened some quotes
    Dim QuotedCharacter As Boolean
    Dim NoErrorAndNoEnd As Boolean

    Set FormulaAlphaData = New BetterArray       'Alphanumeric values of one formula
    Set FormulaData = New BetterArray
    Set VarNameData = New BetterArray       'The list of all Variable Names
    Set SpecCharData = New BetterArray       'The list of all special characters
    Set DictHeaders = New BetterArray
    Set VarMainLabelData = New BetterArray
    Set TableNameData = New BetterArray



    FormulaAlphaData.LowerBound = 1
    VarNameData.LowerBound = 1
    SpecCharData.LowerBound = 1
    DictHeaders.LowerBound = 1

    'squish the formula (removing multiple spaces) to avoid problems related to
    'space collapsing and upper/lower cases

    sFormulaATest = "(" & Application.WorksheetFunction.Trim(sFormula) & ")"


    'Initialisations:

    iNbParentO = 0                               'Number of open brakets
    iNbParentF = 0                               'Number of closed brackets
    iPrevBreak = 1
    OpenedQuotes = False
    NoErrorAndNoEnd = True
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
        AnalysisFormula = "" 'We have to aggregate
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
                                                isFiltered)

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
        AnalysisFormula = "=" & sAlphaValue
    Else

    'MsgBox "Error in analysis formula: " & sFormula

    End If

    Set FormulaAlphaData = Nothing  'Alphanumeric values of one formula
    Set VarNameData = Nothing       'The list of all Variable Names
    Set SpecCharData = Nothing      'The list of all special characters
    Set DictHeaders = Nothing
    Set VarMainLabelData = Nothing
    Set TableNameData = Nothing
End Function





'Change / Or adapt the formula for univariate analysis, bivariate analysis or For just summary part

Function BuildVariateFormula(sTableName As String, _
                             sVarName As String, _
                             Optional sVariate As String = "none", _
                             Optional sFirstCondVar As String = "__all", _
                             Optional sFirstCondVal As String = "__all", _
                             Optional sSecondCondVar As String = "__all", _
                             Optional sSecondCondVal As String = "__all", _
                             Optional isFiltered As Boolean = False) As String



    BuildVariateFormula = vbNullString



    Dim sTable As String 'The name of the table depends on the fact that we want to filter or not

    Dim sAlphaValue As String


    sTable = sTableName

    If isFiltered Then sTable = C_sFiltered & sTableName

    'Fall back to none if you don't precise the univariate / bivariate values: Those are safeguard

    If (sVariate = "univariate" Or sVariate = "univariate total") And _
       (sFirstCondVar = "__all" Or sFirstCondVal = "__all") Then sVariate = "none"

    If (sVariate = "bivariate" Or sVariate = "bivariate total") And _
       (sFirstCondVar = "__all" Or sFirstCondVal = "__all" Or sSecondCondVar = "__all" Or sSecondCondVal = "__all") Then sVariate = "none"



    Select Case sVariate

        Case "none"

          sAlphaValue = sTable & "[" & sVarName & "]"

        Case "univariate"

            sAlphaValue = "IF(" & sTable & "[" & sFirstCondVar & "]" & "=" _
                            & sFirstCondVal & ", " _
                           & sTable & "[" & sVarName & "]" & ")"

        Case "univariate total missing"

            sAlphaValue = "IF(OR(" & sTable & "[" & sFirstCondVar & "]" & "=" _
                            & sFirstCondVal & ", " _
                           & sTable & "[" & sFirstCondVar & "]" & "<>" _
                           & Chr(34) & Chr(34) & "), " _
                           & sTable & "[" & sVarName & "]" & ")"

        Case "bivariate"

            sAlphaValue = "IF( AND(" & sTable & "[" & sFirstCondVar & "]" & "=" _
                           & sFirstCondVal & ", " _
                           & sTable & "[" & sSecondCondVar & "]" & "=" _
                            & sSecondCondVal & "), " _
                           & sTable & "[" & sVarName & "]" & ")"

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

            sFormula = "COUNTIFS" &   "(" & sTable & "[" & sVarName & "], " & sValue & "," & sTable & "[" & sVarName2 & "], " & sValue2 & ")"
        End If
    End If


    AnalysisCount = "=" & sFormula

    Set VarNameData = Nothing
    Set TableNameData = Nothing

End Function
