Attribute VB_Name = "DesignerAnalysis"
Option Explicit


Public Sub BuildAnalysis(Wkb As Workbook, GlobalSummaryData As BetterArray)

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

    'Add global summary first column
    AddGlobalSummary Wkb, GlobalSummaryData

End Sub





'Helpers Subs and Functions ===========================================================================================================================================================================


Private Sub AddGlobalSummary(Wkb As Workbook, GlobalSummaryData As BetterArray)

    Dim iSumLength As Integer
    Dim sFormula As String
    Dim sConvertedFormula As String
    Dim sConvertedFilteredFormula As String
    Dim i As Integer 'counter

    iSumLength = GlobalSummaryData.Length


    With Wkb.Worksheets(C_sSheetAnalysis)

        With .Cells(C_eStartLinesAnalysis - 2, C_eStartColumnAnalysis)
            .value = TranslateLLMsg("MSG_GlobalSummary")
            .Font.Size = C_iAnalysisFontSize + 12
            .Font.Bold = True
            .Font.Color = Helpers.GetColor("DarkBlue")
        End With


        With .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1)
            .value = TranslateLLMsg("MSG_AllData")
            .Font.Bold = True
            .Font.Color = Helpers.GetColor("DarkBlue")
            .Interior.Color = Helpers.GetColor("LightBlue")
        End With

        With .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2)
            .value = TranslateLLMsg("MSG_FilteredData")
            .Font.Bold = True
            .Font.Color = Helpers.GetColor("DarkBlue")
            .Interior.Color = Helpers.GetColor("LightBlue")
        End With

        WriteBorderLines .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1), _
                         iWeight:=xlThin, sColor:="DarkBlue"
        WriteBorderLines .Cells(C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2), _
                         iWeight:=xlThin, sColor:="DarkBlue"

        On Error Resume Next
        For i = 1 To iSumLength

            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).value = GlobalSummaryData.Items(i, 1)
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).Font.Color = Helpers.GetColor("DarkBlue")
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).Interior.Color = Helpers.GetColor("LightBlue")

            sFormula = GlobalSummaryData.Items(i, 2)

            sConvertedFormula = AnalysisFormula(sFormula, Wkb)
            sConvertedFilteredFormula = AnalysisFormula(sFormula, Wkb, isFiltered:=True)

            If sConvertedFormula <> vbNullString Then
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).Formula = sConvertedFormula
            End If

            If sConvertedFilteredFormula <> vbNullString Then
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 2).Formula = sConvertedFilteredFormula
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
        WriteBorderLines .Range(.Cells(C_eStartLinesAnalysis + 1, C_eStartColumnAdmData), _
                                .Cells(C_eStartLinesAnalysis + iSumLength, C_eStartColumnAdmData + 2)), _
                         iWeight:=xlThin, sColor:="DarkBlue"

        WriteBorderLines .Range(.Cells(C_eStartLinesAnalysis + 1, C_eStartColumnAdmData), _
                                .Cells(C_eStartLinesAnalysis + iSumLength, C_eStartColumnAdmData)), _
                         iWeight:=xlThin, sColor:="DarkBlue"

        WriteBorderLines .Range(.Cells(C_eStartLinesAnalysis + 1, C_eStartColumnAdmData + 1), _
                                .Cells(C_eStartLinesAnalysis + iSumLength, C_eStartColumnAdmData + 1)), _
                         iWeight:=xlThin, sColor:="DarkBlue"

    End With

End Sub



