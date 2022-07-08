Attribute VB_Name = "DesignerAnalysis"
Option Explicit


Public Sub BuildAnalysis(Wkb As Workbook, GSData As BetterArray, UAData As BetterArray, BAData As BetterArray, _
                        ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, DictData As BetterArray, _
                        DictHeaders As BetterArray, VarNameData As BetterArray)

    Dim iGoToCol As Long
    Dim prevRef As Byte

    'Add commands Buttons  for filters

     With Wkb.Worksheets(C_sSheetAnalysis)
        .Cells.Font.Size = C_iAnalysisFontSize

        .Rows("1:2").RowHeight = C_iLLButtonsRowHeight
        .Columns(1).ColumnWidth = C_iLLFirstColumnsWidth + 20
         .Columns(1).ColumnWidth = C_iLLFirstColumnsWidth + 20


         'Add command for filtering
        Call AddCmd(Wkb, C_sSheetAnalysis, _
                .Cells(1, 1).Left, _
                .Cells(1, 1).Top, _
                C_sShpFilter, _
                "Calculate on filtered data", _
                C_iCmdWidth, C_iCmdHeight + 10, _
                C_sCmdComputeFilter)
    End With


    'Get the GoTo Column in list_auto
    With Wkb.Worksheets(C_sSheetChoiceAuto)
        iGoToCol = .Cells(C_eStartlinesListAuto, .Columns.Count).End(xlToLeft).Column + 2
    End With


    'Add global summary first column
    AddGlobalSummary Wkb, GSData, iGoToCol

    'Add Univariate Analysis tables
    AddUnivariateAnalysis Wkb, UAData, ChoicesListData, ChoicesLabelsData, DictData, DictHeaders, VarNameData, iGoToCol


    'Build GoTo Area
    BuildGotoArea Wkb:=Wkb, sTableName:=LCase(C_sSheetAnalysis), sSheetName:=C_sSheetAnalysis, iGoToCol:=iGoToCol, iCol:=2


    'Allow text wrap only at the end
    Wkb.Worksheets(C_sSheetAnalysis).Cells.WrapText = True
    Wkb.Worksheets(C_sSheetAnalysis).Cells.EntireRow.AutoFit

    TransferCodeWks Wkb, C_sSheetAnalysis, C_sModLLAnaChange

End Sub





'Helpers Subs and Functions ===========================================================================================================================================================================


Private Sub AddGlobalSummary(Wkb As Workbook, GSData As BetterArray, iGoToCol As Long)

    Dim iSumLength As Integer
    Dim sFormula As String
    Dim sConvertedFormula As String
    Dim sConvertedFilteredFormula As String
    Dim i As Integer 'counter

    iSumLength = GSData.Length


    With Wkb.Worksheets(C_sSheetAnalysis)

        With .Cells(C_eStartLinesAnalysis - 2, C_eStartColumnAnalysis)
            .value = TranslateLLMsg("MSG_GlobalSummary")
            .Font.Size = C_iAnalysisFontSize + 5
            .Font.Bold = True
            .Font.Color = Helpers.GetColor("DarkBlue")
        End With


        With .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1)
            .value = TranslateLLMsg("MSG_AllData")
            .Font.Color = Helpers.GetColor("DarkBlue")
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .Font.Bold = True
            .Font.Size = C_iAnalysisFontSize + 1
        End With

        With .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2)
            .value = TranslateLLMsg("MSG_FilteredData")
            .Font.Color = Helpers.GetColor("DarkBlue")
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .Font.Bold = True
            .Font.Size = C_iAnalysisFontSize + 1
        End With


        On Error Resume Next
        For i = 1 To iSumLength

            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).value = GSData.Items(i, 1)
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).Font.Color = Helpers.GetColor("DarkBlue")
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).Interior.Color = Helpers.GetColor("VeryLightBlue")

            sFormula = GSData.Items(i, 2)

            sConvertedFormula = AnalysisFormula(sFormula, Wkb)
            sConvertedFilteredFormula = AnalysisFormula(sFormula, Wkb, isFiltered:=True)

            If sConvertedFormula <> vbNullString Then
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).FormulaArray = sConvertedFormula
            End If

            If sConvertedFilteredFormula <> vbNullString Then
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2).FormulaArray = sConvertedFilteredFormula
            End If

            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).HorizontalAlignment = xlHAlignRight
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).Font.Size = C_iAnalysisFontSize - 2
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2).HorizontalAlignment = xlHAlignRight
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2).Font.Size = C_iAnalysisFontSize - 2

            'Write boder lines
             WriteBorderLines .Range(.Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis), _
                                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2)), _
                         iWeight:=xlHairline, sColor:="DarkBlue"

        Next
        On Error GoTo 0

        .Columns(C_eStartColumnAnalysis).EntireColumn.AutoFit
        .Columns(C_eStartColumnAnalysis + 1).EntireColumn.AutoFit
        .Columns(C_eStartColumnAnalysis + 2).EntireColumn.AutoFit

        'Write Border Lines

        'Thin lines between formulas
        WriteBorderLines .Range(.Cells(C_eStartLinesAnalysis + 1, C_eStartColumnAnalysis), _
                                .Cells(C_eStartLinesAnalysis + iSumLength, C_eStartColumnAnalysis + 2)), _
                         iWeight:=xlThin, sColor:="DarkBlue"

        'lines on label columns
        WriteBorderLines .Range(.Cells(C_eStartLinesAnalysis + 1, C_eStartColumnAnalysis), _
                                .Cells(C_eStartLinesAnalysis + iSumLength, C_eStartColumnAnalysis)), _
                         iWeight:=xlThin, sColor:="DarkBlue"

        'Lines on the overall table
        WriteBorderLines .Range(.Cells(C_eStartLinesAnalysis + 1, C_eStartColumnAnalysis + 1), _
                                .Cells(C_eStartLinesAnalysis + iSumLength, C_eStartColumnAnalysis + 1)), _
                         iWeight:=xlThin, sColor:="DarkBlue"

    End With

    'Update values of the GoTo Column
    With Wkb.Worksheets(C_sSheetChoiceAuto)

        .Cells(C_eStartlinesListAuto + 1, iGoToCol).value = TranslateLLMsg("MSG_SelectSection") _
                                            & ": " & TranslateLLMsg("MSG_GlobalSummary")

    End With

End Sub


Public Sub AddUnivariateAnalysis(Wkb As Workbook, UAData As BetterArray, ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, _
                                DictData As BetterArray, DictHeaders As BetterArray, VarNameData As BetterArray, iGoToCol As Long)

    Dim sActualSection As String
    Dim sActualGroupBy As String
    Dim sActualSummaryFunction As String
    Dim sActualSummaryLabel As String
    Dim sActualChoice As String
    Dim sActualMainLab As String
    Dim sPreviousSection As String
    Dim sFormula As String
    Dim iCounter As Long
    Dim iSectionRow As Long
    Dim iLength As Long
    Dim iCol As Long
    Dim iRow As Long
    Dim i As Long
    Dim sActualPercentage As String
    Dim sActualMissing As String
    Dim ValidationList As BetterArray

    Set ValidationList = New BetterArray



    iCounter = 1
    sPreviousSection = ""
    With Wkb.Worksheets(C_sSheetAnalysis)


        Do While iCounter <= UAData.Length

            iSectionRow = .Cells(.Rows.Count, C_eStartColumnAnalysis).End(xlUp).Row

            'Actual values in the table of univariate analysis for that line
            sActualSection = UAData.Items(iCounter, 1)
            sActualGroupBy = UAData.Items(iCounter, 2)
            sActualMissing = UAData.Items(iCounter, 3)
            sActualSummaryFunction = UAData.Items(iCounter, 4)
            sActualSummaryLabel = UAData.Items(iCounter, 5)
            sActualPercentage = UAData.Items(iCounter, 6)

            'Set up the different values of the table for one dimensional table
            sActualChoice = DictData.Items(VarNameData.IndexOf(sActualGroupBy), DictHeaders.IndexOf(C_sDictHeaderChoices))
            sActualMainLab = DictData.Items(VarNameData.IndexOf(sActualGroupBy), DictHeaders.IndexOf(C_sDictHeaderMainLab))


            'Where to stop building the table
            iCol = C_eStartColumnAnalysis + 1


            'Set up the sections---------------------------------------------------------------------------------------
            'Value of the section

            If sPreviousSection <> sActualSection Then
                'New Section
                iSectionRow = iSectionRow + 3

                With .Cells(iSectionRow, C_eStartColumnAnalysis)
                    .value = sActualSection
                    .Font.Size = C_iAnalysisFontSize + 3
                    .Font.Color = Helpers.GetColor("DarkBlue")
                End With
                sPreviousSection = sActualSection

                'Build the GoTo column in the list auto sheet
                With Wkb.Worksheets(C_sSheetChoiceAuto)
                    iRow = .Cells(.Rows.Count, iGoToCol).End(xlUp).Row
                    .Cells(iRow + 1, iGoToCol).value = TranslateLLMsg("MSG_SelectSection") & ": " & sActualSection
                End With

                'Draw a border arround the range of section
                With .Range(.Cells(iSectionRow, C_eStartColumnAnalysis), .Cells(iSectionRow, C_eStartColumnAnalysis + 4))
                    With .Borders(xlEdgeBottom)
                        .Weight = xlMedium
                        .LineStyle = xlContinuous
                        .Color = Helpers.GetColor("DarkBlue")
                        .TintAndShade = 0.4
                    End With
                End With
            End If

            'Set up Header of the tables ---------------------------------------
            'Variable Label from the dictionary
            With .Cells(iSectionRow + 3, C_eStartColumnAnalysis)
                .value = sActualMainLab
                .Font.Color = Helpers.GetColor("DarkBlue")
                .HorizontalAlignment = xlHAlignLeft
                .VerticalAlignment = xlVAlignCenter
                .Font.Bold = True
            End With

            'First column on sumary label
            With .Cells(iSectionRow + 3, C_eStartColumnAnalysis + 1)
                .value = sActualSummaryLabel
                .Font.Color = Helpers.GetColor("DarkBlue")
                .Font.Bold = True
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            'Add Percentage header column if required

            If sActualPercentage = C_sYes Then
                With .Cells(iSectionRow + 3, C_eStartColumnAnalysis + 2)
                    .value = TranslateLLMsg("MSG_Percent")
                    .Font.Color = Helpers.GetColor("DarkBlue")
                    .Font.Bold = True
                    .HorizontalAlignment = xlHAlignCenter
                    .VerticalAlignment = xlVAlignCenter
                End With
                iCol = iCol + 1
            End If

            'Add values of the categorical variable ----------------------------
            Set ValidationList = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoice)
            ValidationList.ToExcelRange .Cells(iSectionRow + 4, C_eStartColumnAnalysis)

            'iLength will check the length of the table depending of the number of categorical variables or of add NA
            iLength = iSectionRow + 4 + ValidationList.Length

            'Add NA / Missing if required
            If sActualMissing = C_sYes Then

                .Cells(iLength, C_eStartColumnAnalysis).value = TranslateLLMsg("MSG_NA")

                With .Range(.Cells(iLength, C_eStartColumnAnalysis), .Cells(iLength, iCol))
                    .Font.Color = Helpers.GetColor("GreyBlue")
                    .Interior.Color = Helpers.GetColor("VeryLightGreyBlue")
                    .Font.Size = C_iAnalysisFontSize - 1
                    .Font.Bold = True
                    .NumberFormat = "0.00"
                End With


                'Write formula for first missings
                Select Case Application.WorksheetFunction.Trim(sActualSummaryFunction)
                    Case "COUNT", "COUNT()"

                        'Added + 1 to i because the Validation list starts with index 1
                        sFormula = AnalysisCount(sActualGroupBy, "", Wkb, DictHeaders, isFiltered:=True)

                    Case "SUM", "SUM()"

                    Case Else
                        sFormula = AnalysisFormula(sActualSummaryFunction, Wkb, isFiltered:=True, _
                                sVariate:="univariate", sFirstCondVar:=sActualGroupBy, _
                                sFirstCondVal:="")
                End Select

                If sFormula <> vbNullString Then .Cells(iLength, C_eStartColumnAnalysis + 1).FormulaArray = sFormula

                iLength = iLength + 1
            End If

            'Add Total (Every time) -------------------------------------------------------------------------------

            .Cells(iLength, C_eStartColumnAnalysis).value = TranslateLLMsg("MSG_Total")

            WriteBorderLines .Range(.Cells(iLength, C_eStartColumnAnalysis), .Cells(iLength, iCol)), iWeight:=xlHairline, sColor:="DarkBlue"

            With .Range(.Cells(iLength, C_eStartColumnAnalysis), .Cells(iLength, iCol))
                 .Font.Bold = True
                 .Interior.Color = Helpers.GetColor("VeryLightGreyBlue")
                 .Font.Size = C_iAnalysisFontSize + 1
            End With

            'Add Formulas for total
             Select Case Application.WorksheetFunction.Trim(sActualSummaryFunction)
                    Case "COUNT", "COUNT()"

                        'Added + 1 to i because the Validation list starts with index 1
                        sFormula = AnalysisCount(sActualGroupBy, "", Wkb, DictHeaders, isFiltered:=True, OnTotal:=True)

                    Case "SUM", "SUM()"

                    Case Else
                        sFormula = AnalysisFormula(sActualSummaryFunction, Wkb, isFiltered:=True, sVariate:="none")
            End Select

            If sFormula <> vbNullString Then .Cells(iLength, C_eStartColumnAnalysis + 1).FormulaArray = sFormula

            If sActualPercentage = C_sYes Then
                sFormula = "=" & .Cells(iLength, C_eStartColumnAnalysis + 1).Address & "/" & .Cells(iLength, C_eStartColumnAnalysis + 1).Address
                With .Cells(iLength, C_eStartColumnAnalysis + 2)
                    .Formula = sFormula
                    .Style = "Percent"
                    .NumberFormat = "0.00%"
                End With
            End If


            'Now Work on each category ---------------------------------------------------------------------------------

            For i = 0 To ValidationList.Length - 1

                'Formulas for the first column
                Select Case Application.WorksheetFunction.Trim(sActualSummaryFunction)
                    Case "COUNT", "COUNT()"

                        'Added + 1 to i because the Validation list starts with index 1
                        sFormula = AnalysisCount(sActualGroupBy, ValidationList.Item(i + 1), Wkb, DictHeaders, isFiltered:=True)

                    Case "SUM", "SUM()"

                    Case Else
                        sFormula = AnalysisFormula(sActualSummaryFunction, Wkb, isFiltered:=True, _
                                sVariate:="univariate", sFirstCondVar:=sActualGroupBy, _
                                sFirstCondVal:=ValidationList.Item(i + 1))
                End Select

                If sFormula <> vbNullString Then
                        .Cells(iSectionRow + 4 + i, C_eStartColumnAnalysis + 1).FormulaArray = sFormula
                End If

                'Write the lines for each cells
                With .Cells(iSectionRow + 4 + i, C_eStartColumnAnalysis)
                    .Interior.Color = Helpers.GetColor("VeryLightBlue")
                    .Font.Color = Helpers.GetColor("DarkBlue")
                    .NumberFormat = "0.00"
                End With


                WriteBorderLines .Range(.Cells(iSectionRow + 4 + i, C_eStartColumnAnalysis), .Cells(iSectionRow + 4 + i, iCol)), iWeight:=xlHairline, sColor:="DarkBlue"

                'Add the percentage values

                If sActualPercentage = C_sYes Then
                    sFormula = "=" & .Cells(iSectionRow + 4 + i,  C_eStartColumnAnalysis + 1).Address & "/" & .Cells(iLength, C_eStartColumnAnalysis + 1).Address
                    With .Cells(iSectionRow + 4 + i, iCol)
                        .Style = "Percent"
                        .NumberFormat = "0.00%"
                        .Formula = sFormula
                    End With
                End If

            Next

            'Write borders arround the table

            'Before the total columns
            With .Range(.Cells(iLength - 1, C_eStartColumnAnalysis), .Cells(iLength - 1, iCol))
                With .Borders(xlEdgeBottom)
                    .Weight = xlThin
                    .LineStyle = xlDouble
                    .Color = Helpers.GetColor("DarkBlue")
                End With
            End With

            'On the table outline
            WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), .Cells(iLength, iCol)), iWeight:=xlThin, sColor:="DarkBlue"
            WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), .Cells(iLength, C_eStartColumnAnalysis)), iWeight:=xlThin, sColor:="DarkBlue"
            WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), .Cells(iLength, C_eStartColumnAnalysis + 1)), iWeight:=xlThin, sColor:="DarkBlue"

            iCounter = iCounter + 1
        Loop

    End With
End Sub



