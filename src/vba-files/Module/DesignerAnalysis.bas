Attribute VB_Name = "DesignerAnalysis"
Option Explicit
Option Private Module

Public Sub BuildAnalysis(Wkb As Workbook, GSData As BetterArray, UAData As BetterArray, BAData As BetterArray, _
                         TAData As BetterArray, SAData As BetterArray, _
                         ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, dictData As BetterArray, _
                         DictHeaders As BetterArray, VarNameData As BetterArray, TableNameData As BetterArray)

    Dim iGoToColAna As Long
    Dim iGoToColTA  As Long
    Dim iGoToColSA  As Long
    Dim rng As Range

    ' UNIVARIATE AND BIVARIATE ANALYSIS ============================================================================================

    'Add commands Buttons  for filters

    With Wkb.Worksheets(sParamSheetAnalysis)
        .Cells.Font.Size = C_iAnalysisFontSize

        .Rows("1:2").RowHeight = C_iLLButtonsRowHeight
        .Columns(1).ColumnWidth = C_iLLFirstColumnsWidth + 20

        'Add command for filtering
        Call AddCmd(Wkb, sParamSheetAnalysis, _
                    .Cells(1, 1).Left, _
                    .Cells(1, 1).Top, _
                    C_sShpFilter, _
                    "Calculate on filtered data", _
                    C_iCmdWidth, C_iCmdHeight + 10, _
                    C_sCmdComputeFilter)
    End With


    'Get the GoTo Column in list_auto
    With Wkb.Worksheets(C_sSheetChoiceAuto)
        iGoToColAna = .Cells(C_eStartlinesListAuto, .Columns.Count).End(xlToLeft).Column + 2
    End With


    'Add global summary first column
    AddGlobalSummary Wkb, GSData, iGoToColAna

    'Add Univariate Analysis tables
    AddUnivariateAnalysis Wkb, UAData, ChoicesListData, ChoicesLabelsData, dictData, DictHeaders, VarNameData, iGoToColAna

    'Add Bivariate Analysis
    AddBivariateAnalysis Wkb, BAData, ChoicesListData, ChoicesLabelsData, dictData, DictHeaders, VarNameData, iGoToColAna

    'Build GoTo Area for the analysis (univariate and bivariate)
    BuildGotoArea Wkb:=Wkb, sTableName:=C_sTabLLUBA, sSheetName:=sParamSheetAnalysis, iGoToCol:=iGoToColAna, iCol:=2

    'Allow text wrap only at the end
    FormatAnalysisWorksheet Wkb:=Wkb, sSheetName:=sParamSheetAnalysis, sCodeName:=C_sModLLAnaChange

    'TIME SERIES ANALYSIS =============================================================================================================

    'Update the GoTo Column for the time series analysis

    With Wkb.Worksheets(C_sSheetChoiceAuto)

        iGoToColTA = .Cells(C_eStartlinesListAuto, .Columns.Count).End(xlToLeft).Column + 2

        .Cells(C_eStartlinesListAuto, iGoToColTA).Value = C_sTimeAgg
        .Cells(C_eStartlinesListAuto + 1, iGoToColTA).Value = TranslateLLMsg("MSG_Day")
        .Cells(C_eStartlinesListAuto + 2, iGoToColTA).Value = TranslateLLMsg("MSG_Week")
        .Cells(C_eStartlinesListAuto + 3, iGoToColTA).Value = TranslateLLMsg("MSG_Month")
        .Cells(C_eStartlinesListAuto + 4, iGoToColTA).Value = TranslateLLMsg("MSG_Quarter")
        .Cells(C_eStartlinesListAuto + 5, iGoToColTA).Value = TranslateLLMsg("MSG_Year")

        'Define the list object for validation
        Set rng = .Range(.Cells(C_eStartlinesListAuto, iGoToColTA), .Cells(C_eStartlinesListAuto + 5, iGoToColTA))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "lo" & "_" & C_sTimeAgg

        iGoToColTA = iGoToColTA + 2
    End With

    'Add a dynamic name for the times series
    Wkb.Names.Add Name:=C_sTimeAgg, RefersToR1C1:="=" & "lo" & "_" & C_sTimeAgg & "[" & C_sTimeAgg & "]"

    'Add Temporal Analysis
    AddTimeSeriesAnalysis Wkb, TAData, ChoicesListData, ChoicesLabelsData, dictData, DictHeaders, VarNameData, TableNameData, iGoToColTA

    'Build GoTo Area for the Temporal analysis
    BuildGotoArea Wkb:=Wkb, sTableName:=C_sTabLLTA, sSheetName:=sParamSheetTemporalAnalysis, iGoToCol:=iGoToColTA, _
                  iCol:=C_eStartColumnAnalysis + 2, iFontSize:=C_iAnalysisFontSize

    'Format then worksheet for temporal analysis
    FormatAnalysisWorksheet Wkb:=Wkb, sSheetName:=sParamSheetTemporalAnalysis, iColWidth:=C_iLLFirstColumnsWidth - 8, sCodeName:=C_sModLLAnaChange

    'Column witdth of the start column
    With Wkb.Worksheets(sParamSheetTemporalAnalysis)
        .Cells(1, C_eStartColumnAnalysis + 2).EntireColumn.ColumnWidth = C_iLLFirstColumnsWidth
    End With


    'SPATIAL ANALYSIS ================================================================================================================================


End Sub

'Helpers Subs and Functions ===========================================================================================================================


Private Sub AddGlobalSummary(Wkb As Workbook, GSData As BetterArray, iGoToCol As Long)

    Dim iSumLength As Integer
    Dim sFormula As String
    Dim sConvertedFormula As String
    Dim sConvertedFilteredFormula As String
    Dim i As Long                                'counter

    iSumLength = GSData.Length


    With Wkb.Worksheets(sParamSheetAnalysis)

        With .Cells(C_eStartLinesAnalysis - 2, C_eStartColumnAnalysis)
            .Value = TranslateLLMsg("MSG_GlobalSummary")
            .Font.Size = C_iAnalysisFontSize + 5
            .Font.Bold = True
            .Font.Color = Helpers.GetColor("DarkBlue")
        End With


        With .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1)
            .Value = TranslateLLMsg("MSG_AllData")
            .Font.Color = Helpers.GetColor("DarkBlue")
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .Font.Bold = True
            .Font.Size = C_iAnalysisFontSize + 1
        End With

        With .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2)
            .Value = TranslateLLMsg("MSG_FilteredData")
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
                .Value = GSData.Items(i, 1)
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

            With .Range(.Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1), _
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

        .Cells(C_eStartlinesListAuto + 1, iGoToCol).Value = TranslateLLMsg("MSG_SelectSection") _
      & ": " & TranslateLLMsg("MSG_GlobalSummary")

    End With

    'Add Bivariate Graph

End Sub

Public Sub AddUnivariateAnalysis(Wkb As Workbook, UAData As BetterArray, _
                                 ChoicesListData As BetterArray, _
                                 ChoicesLabelsData As BetterArray, _
                                 dictData As BetterArray, _
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
    Dim sCondition As String                     'Address of the conditions to use in the IF function
    Dim sActualChart As String

    Dim iCounter As Long
    Dim iSectionRow As Long
    Dim iEndRow As Long
    Dim iEndCol As Long
    Dim i As Long
    Dim iRow As Long
    Dim rng As Range


    Dim ValidationList As BetterArray
    Dim Wksh As Worksheet

    Set ValidationList = New BetterArray
    Set Wksh = Wkb.Worksheets(sParamSheetAnalysis)

    iCounter = 2

    sPreviousSection = ""
    With Wksh

        Do While iCounter <= UAData.Length

            iSectionRow = .Cells(.Rows.Count, _
                                 C_eStartColumnAnalysis).End(xlUp).row

            'values in the table of univariate analysis

            sActualSection = UAData.Items(iCounter, 1)
            sActualGroupBy = UAData.Items(iCounter, 2)
            sActualMissing = UAData.Items(iCounter, 3)
            sActualSummaryFunction = UAData.Items(iCounter, 4)
            sActualSummaryLabel = UAData.Items(iCounter, 5)
            sActualPercentage = UAData.Items(iCounter, 6)
            sActualChart = UAData.Items(iCounter, 7)

            'Set up the different values
            If VarNameData.Includes(sActualGroupBy) Then

                sActualChoice = dictData.Items(VarNameData.IndexOf(sActualGroupBy), _
                                               DictHeaders.IndexOf(C_sDictHeaderFormula))

                sActualMainLab = dictData. _
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
                    CreateNewSection Wkb.Worksheets(sParamSheetAnalysis), iSectionRow, C_eStartColumnAnalysis, sActualSection

                    sPreviousSection = sActualSection

                    'Build the GoTo column in the list auto sheet

                    With Wkb.Worksheets(C_sSheetChoiceAuto)
                        iRow = .Cells(.Rows.Count, iGoToCol).End(xlUp).row
                        .Cells(iRow + 1, iGoToCol).Value = TranslateLLMsg("MSG_SelectSection") & _
                                                                                               ": " & sActualSection
                    End With
                End If


                ' Set up Header of the tables  ------------------------------------------------------------------------

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
                            sVar:=sActualGroupBy, iRow:=iEndRow, sPercent:=sActualPercentage, _
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
                                dictData As BetterArray, _
                                DictHeaders As BetterArray, _
                                VarNameData As BetterArray, _
                                iGoToCol As Long)



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
    Dim sActualPercentage As String
    Dim sActualMissing As String

    Dim iCounter As Long
    Dim iSectionRow As Long
    Dim iEndRow As Long
    Dim iEndCol As Long
    Dim iRow As Long


    Dim ValidationListRows As BetterArray        'Categories For Rows
    Dim ValidationListColumns As BetterArray     'Categories For Columns
    Dim Wksh As Worksheet

    Set ValidationListRows = New BetterArray
    Set ValidationListColumns = New BetterArray

    Set Wksh = Wkb.Worksheets(sParamSheetAnalysis)

    iCounter = 2

    sPreviousSection = vbNullString

    With Wksh

        Do While iCounter <= BAData.Length

            iSectionRow = .Cells(.Rows.Count, C_eStartColumnAnalysis).End(xlUp).row

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

                sActualChoiceRow = dictData.Items(VarNameData.IndexOf(sActualGroupByRow), _
                                                  DictHeaders.IndexOf(C_sDictHeaderFormula))
                sActualChoiceColumn = dictData.Items(VarNameData.IndexOf(sActualGroupByColumn), _
                                                     DictHeaders.IndexOf(C_sDictHeaderFormula))

                sActualMainLabRow = dictData. _
                                    Items(VarNameData.IndexOf(sActualGroupByRow), _
                                          DictHeaders. _
                                          IndexOf(C_sDictHeaderMainLab))

                sActualMainLabColumn = dictData. _
                                       Items(VarNameData.IndexOf(sActualGroupByColumn), _
                                             DictHeaders. _
                                             IndexOf(C_sDictHeaderMainLab))

                'Where to stop building the table


                'Value of the section

                If sPreviousSection <> sActualSection Or iCounter = 2 Then
                    'New Section
                    iSectionRow = iSectionRow + 3

                    'Create a new section
                    CreateNewSection Wkb.Worksheets(sParamSheetAnalysis), iSectionRow, _
        C_eStartColumnAnalysis, sActualSection

                    sPreviousSection = sActualSection

                    'Build the GoTo column in the list auto sheet

                    With Wkb.Worksheets(C_sSheetChoiceAuto)
                        iRow = .Cells(.Rows.Count, iGoToCol).End(xlUp).row
                        .Cells(iRow + 1, iGoToCol).Value = TranslateLLMsg("MSG_SelectSection") & _
                                                                                               ": " & sActualSection
                    End With
                End If

                ' Set up Header of the tables  -------------------------------------------------------------------------

                ' Then EndColumn iEndCol is a ByRef, to update the ends column

                Set ValidationListRows = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoiceRow)
                Set ValidationListColumns = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoiceColumn)

                iEndCol = C_eStartColumnAnalysis + ValidationListColumns.Length - 1

                CreateBATable Wksh, iRow:=iSectionRow + 3, ColumnsData:=ValidationListColumns, _
                              rowsData:=ValidationListRows, iCol:=C_eStartColumnAnalysis, _
                              sMainLabRow:=sActualMainLabRow, sMainLabCol:=sActualMainLabColumn, _
                              sSummaryLabel:=sActualSummaryLabel, _
                              sPercent:=sActualPercentage, sMiss:=sActualMissing

                iEndCol = .Cells(iSectionRow + 5, .Columns.Count).End(xlToLeft).Column
                iEndRow = .Cells(.Rows.Count, C_eStartColumnAnalysis).End(xlUp).row

                'Add Formulas in the interior of the table
                AddInnerFormula Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sActualSummaryFunction, _
                                iStartRow:=iSectionRow + 4, iStartCol:=C_eStartColumnAnalysis, iEndRow:=iEndRow, _
                                iEndCol:=iEndCol, sVarRow:=sActualGroupByRow, sVarColumn:=sActualGroupByColumn, _
                                sMiss:=sActualMissing, sPercent:=sActualPercentage

                'Add Formulas at the borders of the table
                AddBordersFormula Wkb:=Wkb, DictHeaders:=DictHeaders, sForm:=sActualSummaryFunction, _
                                  iStartRow:=iSectionRow + 4, iStartCol:=C_eStartColumnAnalysis, iEndRow:=iEndRow, _
                                  iEndCol:=iEndCol, sVarRow:=sActualGroupByRow, sVarColumn:=sActualGroupByColumn, _
                                  sMiss:=sActualMissing, sPercent:=sActualPercentage
            End If

            iCounter = iCounter + 1
        Loop

    End With

End Sub

Sub AddTimeSeriesAnalysis(Wkb As Workbook, TAData As BetterArray, _
                          ChoicesListData As BetterArray, _
                          ChoicesLabelsData As BetterArray, _
                          dictData As BetterArray, _
                          DictHeaders As BetterArray, _
                          VarNameData As BetterArray, _
                          TableNameData As BetterArray, _
                          iGoToCol As Long, _
                          Optional sOutlineColor As String = "DarkBlue", _
                          Optional sHeaderFontColor As String = "White", _
                          Optional sHeaderInteriorColor As String = "VeryDarkBlue", _
                          Optional isFiltered As Boolean = True)



    Dim sActualSerie As String
    Dim sActualSection As String
    Dim sPreviousSection As String
    Dim sActualTimeVar As String
    Dim sActualGroupBy As String
    Dim sActualMissing As String
    Dim sActualSummaryFunction As String
    Dim sActualSummaryLabel As String
    Dim sActualPercentage As String
    Dim sActualChoice As String
    Dim sActualMainLabColumn As String
    Dim sActualAddTotal As String
    Dim sMinimumFormula As String
    Dim sMaximumFormula As String
    Dim sPrevTimeVar As String
    Dim sTableName As String
    Dim iRow As Long

    Dim iCounter As Long                         'Counter for the length of the Time Series Data
    Dim iSectionRow As Long
    Dim iPrevCol As Long
    Dim iStartCol As Long
    Dim NewSection As Boolean


    'Temporal analysis worksheet
    Dim Wksh As Worksheet
    'Columns for the group by if there is one
    Dim ValidationListColumns As BetterArray
    Dim rng As Range


    Set Wksh = Wkb.Worksheets(sParamSheetTemporalAnalysis)
    Set ValidationListColumns = New BetterArray


    iCounter = 2

    'By default, the new section is 3
    iSectionRow = 3

    'Previous column
    iPrevCol = C_eStartColumnAnalysis + 2

    'Initialise the newSection
    sPreviousSection = TAData.Items(iCounter, 2)
    sPrevTimeVar = TAData.Items(iCounter, 3)


    With Wksh

        iStartCol = C_eStartColumnAdmData + 2

        Do While iCounter <= TAData.Length

            sActualSerie = TAData.Items(iCounter, 1)
            sActualSection = TAData.Items(iCounter, 2)
            sActualTimeVar = TAData.Items(iCounter, 3)
            sActualGroupBy = TAData.Items(iCounter, 4)
            sActualMissing = TAData.Items(iCounter, 5)
            sActualSummaryFunction = TAData.Items(iCounter, 6)
            sActualSummaryLabel = TAData.Items(iCounter, 7)
            sActualPercentage = TAData.Items(iCounter, 8)
            sActualAddTotal = TAData.Items(iCounter, 9)


            sActualMainLabColumn = vbNullString
            ValidationListColumns.Clear
            NewSection = sPreviousSection <> sActualSection


            'Test if there is a need to enter the process (by testing the time variable)
            If VarNameData.Includes(sActualTimeVar) Then

                sTableName = TableNameData.Items(VarNameData.IndexOf(sActualTimeVar))

                If isFiltered Then sTableName = C_sFiltered & sTableName

                'Build new section
                If NewSection Or iCounter = 2 Then

                    'Update minimum and maximum
                    If iCounter <> 2 Then
                        .Cells(iSectionRow + 3, C_eStartColumnAnalysis + 10).formula = "= MIN(" & sMinimumFormula & ")"
                        .Cells(iSectionRow + 3, C_eStartColumnAnalysis + 12).formula = "= MAX(" & sMaximumFormula & ")"
                    End If

                    iSectionRow = .Cells(.Rows.Count, C_eStartColumnAnalysis + 2).End(xlUp).row + 3
                    iStartCol = C_eStartColumnAnalysis + 2

                    'Create a new section, and new minimum formula
                    .Range(.Cells(iSectionRow, C_eStartColumnAnalysis + 2), .Cells(iSectionRow, C_eStartColumnAnalysis + 3)).Merge
                    CreateNewSection Wksh, iSectionRow, C_eStartColumnAnalysis + 2, sActualSection

                    'Update Previous Section
                    sPreviousSection = sActualSection

                    'Build the GoTo column in the list auto sheet
                    With Wkb.Worksheets(C_sSheetChoiceAuto)
                        iRow = .Cells(.Rows.Count, iGoToCol).End(xlUp).row
                        .Cells(iRow + 1, iGoToCol).Value = TranslateLLMsg("MSG_SelectSection") & ": " & sActualSection
                    End With

                    'Add the start date, Time aggregation, and Time column
                    AddTimeColumn Wksh, iSectionRow, C_eStartColumnAnalysis + 2
                End If

                'Update the minimum formula
                If sPrevTimeVar <> sActualTimeVar And Not NewSection Then
                    'New time variable, no new section

                    sMinimumFormula = sMinimumFormula & ", " & "MIN(" & sTableName & "[" & sActualTimeVar & "]" & ")"
                    sMaximumFormula = sMaximumFormula & ", " & "MAX(" & sTableName & "[" & sActualTimeVar & "]" & ")"

                    sPrevTimeVar = sActualTimeVar

                ElseIf NewSection Or iCounter = 2 Then
                    'New section, new minimum and maximum
                    sMinimumFormula = "MIN(" & sTableName & "[" & sActualTimeVar & "]" & ")"
                    sMaximumFormula = "MAX(" & sTableName & "[" & sActualTimeVar & "]" & ")"

                End If

                'Create a validation list if there it is needed
                If VarNameData.Includes(sActualGroupBy) Then
                    sActualChoice = dictData.Items(VarNameData.IndexOf(sActualGroupBy), DictHeaders.IndexOf(C_sDictHeaderFormula))
                    sActualMainLabColumn = dictData.Items(VarNameData.IndexOf(sActualGroupBy), DictHeaders.IndexOf(C_sDictHeaderMainLab))
                    Set ValidationListColumns = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoice)
                End If

                CreateBATable Wksh, iRow:=iSectionRow + 6, ColumnsData:=ValidationListColumns, _
                              iCol:=iStartCol, sMainLabCol:=sActualMainLabColumn, _
                              sSummaryLabel:=sActualSummaryLabel, _
                              sPercent:=sActualPercentage, sMiss:=sActualMissing, _
                              isTimeSeries:=True
                DoEvents

                iPrevCol = iStartCol + 1
                iStartCol = .Cells(iSectionRow + 8, .Columns.Count).End(xlToLeft).Column

                AddTimeSeriesFormula Wkb, DictHeaders:=DictHeaders, sForm:=sActualSummaryFunction, _
                                     sTimeVar:=sActualTimeVar, sCondVar:=sActualGroupBy, iRow:=iSectionRow + 7, _
                                     iStartCol:=iPrevCol, iEndCol:=iStartCol, sPerc:=sActualPercentage, _
                                     sMiss:=sActualMissing



                Set rng = .Range(.Cells(iSectionRow + 7, iPrevCol), .Cells(iSectionRow + 11 + C_iNbTime, iStartCol))
                WriteBorderLines rng, sColor:=sOutlineColor, iWeight:=xlMedium


                Set rng = .Range(.Cells(iSectionRow + 6, iPrevCol), .Cells(iSectionRow + 6, iStartCol))

                rng.Merge
                FormatARange rng:=rng, sInteriorColor:=sHeaderInteriorColor, sFontColor:=sHeaderFontColor, isBold:=True
                WriteBorderLines rng, sColor:=sOutlineColor, iWeight:=xlMedium

            End If


            iCounter = iCounter + 1

        Loop

        'Add formula at the end for the start date at the end

        On Error Resume Next
        .Cells(iSectionRow + 3, C_eStartColumnAnalysis + 10).formula = "= MIN(" & sMinimumFormula & ")"
        .Cells(iSectionRow + 3, C_eStartColumnAnalysis + 12).formula = "= MAX(" & sMaximumFormula & ")"
        On Error GoTo 0

    End With
End Sub


