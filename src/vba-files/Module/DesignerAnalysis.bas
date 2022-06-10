Attribute VB_Name = "DesignerAnalysis"
Option Explicit


Public Sub BuildAnalysis(Wkb As Workbook, GlobalSummaryData As BetterArray)

    'Add commands


    With Wkb

        'Add global summary first column
        AddGlobalSummary Wkb, GlobalSummaryData

    End With



End Sub



















'Helpers Subs and Functions ===========================================================================================================================================================================


Private Sub AddGlobalSummary(Wkb As Workbook, GlobalSummaryData As BetterArray)

    Dim iSumLength As Integer
    Dim sFormula As String
    Dim sConvertedFormula As String
    Dim i As Integer 'counter

    iSumLength = GlobalSummaryData.Length

    With Wkb.Worksheets(C_sSheetAnalysis)

        .Cells.Font.Size = C_iAnalysisFontSize

        With .Cells(C_eStartLinesAnalysis - 1, C_eStartColumnAnalysis)
            .value = TranslateLLMsg("MSG_GlobalSummary")
            .Font.Size = C_iAnalysisFontSize + 12
            .Font.Bold = True
            .Font.Color = Helpers.GetColor("DarkBlue")
        End With

        On Error Resume Next
        For i = 1 To iSumLength

            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).value = GlobalSummaryData.Items(i, 1)
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).Font.Color = Helpers.GetColor("DarkBlue")
            .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis).Interior.Color = Helpers.GetColor("LightBlue")
            
            sFormula = GlobalSummaryData.Items(i, 2)

            sConvertedFormula = AnalysisFormula(sFormula, Wkb)

            If sConvertedFormula <> vbNullString Then
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).Formula = sConvertedFormula
            End If
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).HorizontalAlignment = xlHAlignRight
                .Cells(i + C_eStartLinesAnalysis, C_eStartColumnAnalysis + 1).Font.Size = C_iAnalysisFontSize - 2
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
