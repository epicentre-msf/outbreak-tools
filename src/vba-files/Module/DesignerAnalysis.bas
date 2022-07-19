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

            sConvertedFormula = AnalysisFormula(Wkb, sFormula)
            sConvertedFilteredFormula = AnalysisFormula(Wkb, sFormula, isFiltered:=True)

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


Public Sub AddUnivariateAnalysis(Wkb As Workbook, UAData As BetterArray, _
                                ChoicesListData As BetterArray, _
                                ChoicesLabelsData As BetterArray, _
                                DictData As BetterArray, _
                                DictHeaders As BetterArray, _
                                VarNameData As BetterArray, _
                                iGoToCol As Long, _
                                Optional sOutlineColor As String = "DarkBlue")



    Dim sActualSection As String
    Dim sActualGroupBy As String
    Dim sActualSummaryFunction As String
    Dim sActualSummaryLabel As String
    Dim sActualChoice As String
    Dim sActualMainLab As String
    Dim sPreviousSection As String
    Dim sFormula As String
    Dim sActualPercentage As String
    Dim sActualMissing As String
    Dim sCondition As String 'Address of the conditions to use in the IF function

        Dim iCounter As Long
    Dim iSectionRow As Long
    Dim iEndRow As Long
    Dim iEndCol As Long
    Dim i As Long
    Dim iRow As Long


    Dim ValidationList As BetterArray
    Dim Wksh As Worksheet

    Set ValidationList = New BetterArray
    Set Wksh = Wkb.Worksheets(C_sSheetAnalysis)

    iCounter = 1

    sPreviousSection = ""
    With Wksh

        Do While iCounter <= UAData.Length

            iSectionRow = .Cells(.Rows.Count, _
            C_eStartColumnAnalysis).End(xlUp).Row

            'values in the table of univariate analysis

            sActualSection = UAData.Items(iCounter, 1)
            sActualGroupBy = UAData.Items(iCounter, 2)
            sActualMissing = UAData.Items(iCounter, 3)
            sActualSummaryFunction = UAData.Items(iCounter, 4)
            sActualSummaryLabel = UAData.Items(iCounter, 5)
            sActualPercentage = UAData.Items(iCounter, 6)

            'Set up the different values

            sActualChoice = DictData.Items(VarNameData.IndexOf(sActualGroupBy), _
                                            DictHeaders.IndexOf(C_sDictHeaderChoices))

            sActualMainLab = DictData. _
            Items(VarNameData.IndexOf(sActualGroupBy), _
            DictHeaders. _
            IndexOf(C_sDictHeaderMainLab))

            'Where to stop building the table

            iEndCol = C_eStartColumnAnalysis + 1

            'Value of the section

            If sPreviousSection <> sActualSection Then
                    'New Section
                iSectionRow = iSectionRow + 3

                'Create a new section
                                CreateNewSection Wkb.Worksheets(C_sSheetAnalysis), iSectionRow, _
                                C_eStartColumnAnalysis, sActualSection

                sPreviousSection = sActualSection

                'Build the GoTo column in the list auto sheet

                With Wkb.Worksheets(C_sSheetChoiceAuto)
                        iRow = .Cells(.Rows.Count, iGoToCol).End(xlUp).Row
                    .Cells(iRow + 1, iGoToCol).value = TranslateLLMsg("MSG_SelectSection") & _
                                                                                           ": " & sActualSection
                End With
            End If

            ' Set up Header of the tables  -------------------------------------------

            ' Then EndColumn iEndCol is a ByRef, to update the ends column

            CreateUAHeaders Wksh, iRow:=iSectionRow + 3, iCol:=C_eStartColumnAnalysis, _
                            sMainLab:=sActualMainLab, sSummaryLabel:=sActualSummaryLabel, _
                            sPercent:=sActualPercentage, iEndCol:=iEndCol

                        ' Update the EndColumn if we have to add percentages
                        If sActualPercentage = C_sYes Then iEndCol = iEndCol + 1

            'Add values of the categorical variable -------------------------------------------

            Set ValidationList = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoice)
            ValidationList.ToExcelRange .Cells(iSectionRow + 4, C_eStartColumnAnalysis)

                        'EndRow of the table.
            iEndRow = iSectionRow + 4 + ValidationList.Length


            'Add NA / Missing if required -----------------------------------------------------

            If sActualMissing = C_sYes Then

                AddUANA Wkb:=Wkb, DictHeaders:=DictHeaders, sSumFunc:=sActualSummaryFunction, _
                sVar:=sActualGroupBy, iRow:=iEndRow, _
                iStartCol:=C_eStartColumnAnalysis, iEndCol:=iEndCol

                iEndRow = iEndRow + 1

            End If

            'Add Total (Every time) ------------------------------------------------------------------------------------

                        AddUATotal Wkb:=Wkb, DictHeaders:=DictHeaders, sSumFunc:=sActualSummaryFunction, _
                                sVar:=sActualGroupBy, iRow:=iEndRow, iStartCol:=C_eStartColumnAnalysis, iEndCol:=iEndCol, _
                                sPercent:=sActualPercentage, sMiss:=sActualMissing


            'Now Work on each category ---------------------------------------------------------------------------------


            For i = 0 To ValidationList.Length - 1


                'Address of the condition to use
                sCondition = .Cells(iSectionRow + 4 + i, C_eStartColumnAnalysis).Address

                'Getting the formulas
                sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sActualSummaryFunction, _
                                       sVar:=sActualGroupBy, sCondition:=sCondition)
                
                On Error Resume Next

                If sFormula <> vbNullString And Len(sFormula) < 255 Then

                        .Cells(iSectionRow + 4 + i, C_eStartColumnAnalysis + 1).FormulaArray = sFormula

                End If

                On Error GoTo 0

                FormatCell Wksh:=Wksh, iStartRow:=iSectionRow + 4 + i, _
                           iEndRow:=iEndRow, iStartCol:=C_eStartColumnAnalysis, _
                           iEndCol:=iEndCol, sPercent:=sActualPercentage

            Next


            'On the table outline ---------------------------------------------------------------------------------

            WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), _
                                             .Cells(iEndRow, iEndCol)), iWeight:=xlThin, sColor:=sOutlineColor

            WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), _
                                             .Cells(iEndRow, C_eStartColumnAnalysis)), iWeight:=xlThin, sColor:=sOutlineColor

            WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), _
                                             .Cells(iEndRow, C_eStartColumnAnalysis + 1)), iWeight:=xlThin, sColor:=sOutlineColor

            iCounter = iCounter + 1

        Loop
    End With

End Sub



