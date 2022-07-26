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

    'Add Bivariate Analysis
     AddBivariateAnalysis Wkb, BAData, ChoicesListData, ChoicesLabelsData, DictData, DictHeaders, VarNameData, iGoToCol


    'Build GoTo Area
    BuildGotoArea Wkb:=Wkb, sTableName:=LCase(C_sSheetAnalysis), sSheetName:=C_sSheetAnalysis, iGoToCol:=iGoToCol, iCol:=2


    'Allow text wrap only at the end
    Wkb.Worksheets(C_sSheetAnalysis).Cells.WrapText = True
    Wkb.Worksheets(C_sSheetAnalysis).Cells.EntireRow.AutoFit
    Wkb.Worksheets(C_sSheetAnalysis).Cells.EntireColumn.ColumnWidth = C_iLLFirstColumnsWidth

    TransferCodeWks Wkb, C_sSheetAnalysis, C_sModLLAnaChange

End Sub



'Helpers Subs and Functions ===========================================================================================================================================================================


Private Sub AddGlobalSummary(Wkb As Workbook, GSData As BetterArray, iGoToCol As Long)

    Dim iSumLength As Integer
    Dim sFormula As String
    Dim sConvertedFormula As String
    Dim sConvertedFilteredFormula As String
    Dim i As Long 'counter

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

        'Formulas for Global Summary
        For i = 2 To iSumLength
            With .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis)
                .value = GSData.Items(i, 1)
                .Font.Color = Helpers.GetColor("DarkBlue")
                .Interior.Color = Helpers.GetColor("VeryLightBlue")
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlHAlignLeft
            End With

            sFormula = GSData.Items(i, 2)
            sConvertedFormula = AnalysisFormula(Wkb, sFormula)
            sConvertedFilteredFormula = AnalysisFormula(Wkb, sFormula, isFiltered:=True)

            If sConvertedFormula <> vbNullString And sConvertedFilteredFormula <> vbNullString Then
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).FormulaArray = sConvertedFormula
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2).FormulaArray = sConvertedFilteredFormula
            End If

            With Range(.Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1), _
                       .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2))
                .HorizontalAlignment = xlHAlignRight
                .Font.Size = C_iAnalysisFontSize - 2
            End With

            'Write boder lines
             WriteBorderLines .Range(.Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis), _
                                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2)), _
                         iWeight:=xlHairline, sColor:="DarkBlue"

        Next

        On Error GoTo 0

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

    iCounter = 2

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
            If VarNameData.Includes(sActualGroupBy) Then

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
                                sPercent:=sActualPercentage

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

            End If

                iCounter = iCounter + 1
        Loop

    End With

End Sub



Public Sub AddBivariateAnalysis(Wkb As Workbook, BAData As BetterArray, _
                                ChoicesListData As BetterArray, _
                                ChoicesLabelsData As BetterArray, _
                                DictData As BetterArray, _
                                DictHeaders As BetterArray, _
                                VarNameData As BetterArray, _
                                iGoToCol As Long, _
                                Optional sOutlineColor As String = "DarkBlue")



    Dim sActualSection As String
    Dim sActualGroupByColumn As String
    Dim sActualGroupByRow As String
    Dim sActualSummaryFunction As String
    Dim sActualSummaryLabel As String
    Dim sActualChoiceRow As String
    Dim sActualChoiceColumn As String
    Dim sActualMainLabRow As String
    Dim sActualMainLabColumn As String
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


    Dim ValidationListRows As BetterArray 'Categories For Rows
    Dim ValidationListColumns As BetterArray 'Categories For Columns
    Dim Wksh As Worksheet

    Set ValidationListRows = New BetterArray
    Set ValidationListColumns = New BetterArray

    Set Wksh = Wkb.Worksheets(C_sSheetAnalysis)

    iCounter = 2

    sPreviousSection = vbNullString

    With Wksh

        Do While iCounter <= BAData.Length

            iSectionRow = .Cells(.Rows.Count, _
            C_eStartColumnAnalysis).End(xlUp).Row

            'values in the table of univariate analysis

            sActualSection = BAData.Items(iCounter, 1)
            sActualGroupByRow = BAData.Items(iCounter, 2)
            sActualGroupByColumn = BAData.Items(iCounter, 3)
            sActualMissing = BAData.Items(iCounter, 4)
            sActualSummaryFunction = BAData.Items(iCounter, 5)
            sActualSummaryLabel = BAData.Items(iCounter, 6)
            sActualPercentage = BAData.Items(iCounter, 7)

            'Set up the different values
            If VarNameData.Includes(sActualGroupByColumn) And VarNameData.Includes(sActualGroupByRow) Then

                sActualChoiceRow = DictData.Items(VarNameData.IndexOf(sActualGroupByRow), _
                                                DictHeaders.IndexOf(C_sDictHeaderChoices))
                sActualChoiceColumn = DictData.Items(VarNameData.IndexOf(sActualGroupByColumn), _
                                                DictHeaders.IndexOf(C_sDictHeaderChoices))

                sActualMainLabRow = DictData. _
                Items(VarNameData.IndexOf(sActualGroupByRow), _
                DictHeaders. _
                IndexOf(C_sDictHeaderMainLab))

                sActualMainLabColumn = DictData. _
                Items(VarNameData.IndexOf(sActualGroupByColumn), _
                DictHeaders. _
                IndexOf(C_sDictHeaderMainLab))

                'Where to stop building the table


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

                Set ValidationListRows = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoiceRow)
                Set ValidationListColumns = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoiceColumn)

                iEndCol = C_eStartColumnAnalysis + ValidationListColumns.Length - 1

                CreateBAHeaders Wksh, iRow:=iSectionRow + 3, ColumnsData:=ValidationListColumns, _
                                RowsData:=ValidationListRows, iCol:=C_eStartColumnAnalysis, _
                                sMainLabRow:=sActualMainLabRow, sMainLabCol:=sActualMainLabColumn, _
                                sSummaryLabel:=sActualSummaryLabel, _
                                sPercent:=sActualPercentage, sMiss:=sActualMissing

                iEndCol = .Cells(iSectionRow + 4, .Columns.Count).End(xlToLeft).Column
                iEndRow = .Cells(.Rows.Count, C_eStartColumnAnalysis).End(xlUp).Row


                'Add NA / Missing if required -----------------------------------------------------

                ' If sActualMissing = C_sYes Then

                '     AddBANA Wkb:=Wkb, DictHeaders:=DictHeaders, sSumFunc:=sActualSummaryFunction, _
                '     sVar:=sActualGroupBy, iRow:=iEndRow, _
                '     iStartCol:=C_eStartColumnAnalysis, iEndCol:=iEndCol

                '     iEndRow = iEndRow + 1

                ' End If
'
'                'Add Total (Every time) ------------------------------------------------------------------------------------
'
'                            AddUATotal Wkb:=Wkb, DictHeaders:=DictHeaders, sSumFunc:=sActualSummaryFunction, _
'                                    sVar:=sActualGroupBy, iRow:=iEndRow, iStartCol:=C_eStartColumnAnalysis, iEndCol:=iEndCol, _
'                                    sPercent:=sActualPercentage, sMiss:=sActualMissing
'
'
'                'Now Work on each category ---------------------------------------------------------------------------------
'
'
'                For i = 0 To ValidationList.Length - 1
'
'
'                    'Address of the condition to use
'                    sCondition = .Cells(iSectionRow + 4 + i, C_eStartColumnAnalysis).Address
'
'                    'Getting the formulas
'                    sFormula = UnivariateFormula(Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sActualSummaryFunction, _
'                                           sVar:=sActualGroupBy, sCondition:=sCondition)
'
'                    On Error Resume Next
'
'                    If sFormula <> vbNullString And Len(sFormula) < 255 Then
'
'                            .Cells(iSectionRow + 4 + i, C_eStartColumnAnalysis + 1).FormulaArray = sFormula
'
'                    End If
'
'                    On Error GoTo 0
'
'                    FormatCell Wksh:=Wksh, iStartRow:=iSectionRow + 4 + i, _
'                               iEndRow:=iEndRow, iStartCol:=C_eStartColumnAnalysis, _
'                               iEndCol:=iEndCol, sPercent:=sActualPercentage
'
'                Next
'
'
'                'On the table outline ---------------------------------------------------------------------------------
'
'                WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), _
'                                                 .Cells(iEndRow, iEndCol)), iWeight:=xlThin, sColor:=sOutlineColor
'
'                WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), _
'                                                 .Cells(iEndRow, C_eStartColumnAnalysis)), iWeight:=xlThin, sColor:=sOutlineColor
'
'                WriteBorderLines .Range(.Cells(iSectionRow + 4, C_eStartColumnAnalysis), _
'                                                 .Cells(iEndRow, C_eStartColumnAnalysis + 1)), iWeight:=xlThin, sColor:=sOutlineColor
'
            End If
'
                iCounter = iCounter + 1
        Loop
'
   End With

End Sub
