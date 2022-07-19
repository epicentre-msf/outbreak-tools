Attribute VB_Name = "DesignerAnalysisHelpers"


Option Explicit




'FUNCTIONS USED TO BUILD UNIVARIATE ANALYSIS ===================================================================================================================

'Create New section
Sub CreateNewSection(Wksh As Worksheet, iRow As Long, iCol As Long, sSection As String, _
                                         Optional sColor As String = "DarkBlue")
        With Wksh
                'New range
                With .Cells(iRow, iCol)
                    .value = sSection
                    .Font.Size = C_iAnalysisFontSize + 3
                    .Font.Color = Helpers.GetColor(sColor)
                End With

                'Draw a border arround the section
                With .Range(.Cells(iRow, iCol), .Cells(iRow, iCol + 4))
                    With .Borders(xlEdgeBottom)
                            .Weight = xlMedium
                            .LineStyle = xlContinuous
                            .Color = Helpers.GetColor(sColor)
                            .TintAndShade = 0.4
                         End With
                 End With
        End With
End Sub


'Create Headers for univariate analysis
Sub CreateUAHeaders(Wksh As Worksheet, iRow As Long, iCol As Long, _
                                        sMainLab As String, sSummaryLabel As String, _
                                        sPercent As String, iEndCol As Long, Optional sColor As String = "DarkBlue")
        With Wksh
                'Variable Label from the dictionary
        With .Cells(iRow, iCol)
            .value = sMainLab
            .Font.Color = Helpers.GetColor(sColor)
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
            .Font.Bold = True
        End With
        'First column on sumary label
        With .Cells(iRow, iCol + 1)
            .value = sSummaryLabel
            .Font.Color = Helpers.GetColor(sColor)
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        'Add Percentage header column if required
        If sPercent = C_sYes Then
            With .Cells(iRow, iCol + 2)
                .value = TranslateLLMsg("MSG_Percent")
                .Font.Color = Helpers.GetColor(sColor)
                .Font.Bold = True
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
        End If
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

      With .Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol))
              .Font.Color = Helpers.GetColor(sFontColor)
              .Interior.Color = Helpers.GetColor(sInteriorColor)
              .Font.Size = C_iAnalysisFontSize - 1
              .Font.Bold = True
              .NumberFormat = sNumberFormat
      End With

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
            
            With .Range(.Cells(iRow, iStartCol), .Cells(iRow, iEndCol))
             .Font.Bold = True
             .Interior.Color = Helpers.GetColor(sInteriorColor)
             .Font.Size = C_iAnalysisFontSize + 1
            End With
            
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

              sFormula = AnalysisCount(Wkb, DictHeaders, sVar, sCondition, isFiltered, OnTotal, includeMissing)

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

        Case "bivariate total"



        'By default fall back to simple varname in a table

        Case Else
             sAlphaValue = sTable & "[" & sVarName & "]"
    End Select

    BuildVariateFormula = sAlphaValue

End Function



'Analysis Count


Function AnalysisCount(Wkb As Workbook, DictHeaders As BetterArray, sVarName As String, sValue As String, Optional isFiltered As Boolean = False, _
                      Optional OnTotal As Boolean = False, Optional includeMissing As Boolean = False) As String



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

    sFormula = ""

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


    AnalysisCount = "=" & sFormula

    Set VarNameData = Nothing
    Set TableNameData = Nothing

End Function
