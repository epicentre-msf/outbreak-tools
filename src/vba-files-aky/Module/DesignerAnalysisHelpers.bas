Attribute VB_Name = "DesignerAnalysisHelpers"


Option Explicit

'Transform one formula to a formula for analysis.
'Wkb is a workbook where we can find the dictionary, the special character
'data and the name of all 'friendly' functions

Public Function AnalysisFormula(sFormula As String, Wkb As Workbook, _
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
                    If Not VarNameData.Includes(LCase(sAlphaValue)) And Not FormulaData.Includes(UCase(sAlphaValue)) And Not IsNumeric(sAlphaValue) Then
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
    If sVariate = "univariate" And (sFirstCondVar = "__all" Or sFirstCondVal = "__all") Then sVariate = "none"
    If sVariate = "bivariate" And (sFirstCondVar = "__all" Or sFirstCondVal = "__all" Or sSecondCondVar = "__all" Or sSecondCondVal = "__all") Then sVariate = "none"

    Select Case sVariate

        Case "none"

           sAlphaValue = sTable & "[" & sVarName & "]"

        Case "univariate"

            sAlphaValue = "IF(" & sTable & "[" & sFirstCondVar & "]" & "=" _
                           & Chr(34) & sFirstCondVal & Chr(34) & ", " _
                           & sTable & "[" & sVarName & "]" & ")"
        Case "bivariate"

            sAlphaValue = "IF( AND(" & sTable & "[" & sFirstCondVar & "]" & "=" _
                           & Chr(34) & sFirstCondVal & Chr(34) & ", " _
                           & sTable & "[" & sSecondCondVar & "]" & "=" _
                           & Chr(34) & sSecondCondVal & Chr(34) & "), " _
                           & sTable & "[" & sVarName & "]" & ")"

        'By default fall back to simple varname in a table
        Case Else
             sAlphaValue = sTable & "[" & sVarName & "]"
    End Select


    BuildVariateFormula = sAlphaValue

End Function



''Analysis Count


Function AnalysisCount(sVarName As String, sValue As String, Wkb As Workbook, DictHeaders As BetterArray, Optional isFiltered As Boolean = False) As String


    Dim VarNameData As BetterArray
    Dim TableNameData As BetterArray
    Dim sTable As String
    Dim sFormula As String

    Set VarNameData = New BetterArray
    Set TableNameData = New BetterArray

    VarNameData.LowerBound = 1
    TableNameData.LowerBound = 1

    VarNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, 1), DetectLastColumn:=False, DetectLastRow:=True
    TableNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, DictHeaders.IndexOf(C_sDictHeaderTableName)), DetectLastColumn:=False, DetectLastRow:=True

    sFormula = ""

    If VarNameData.Includes(sVarName) Then
        sTable = TableNameData.Items(VarNameData.IndexOf(sVarName))
        If isFiltered Then sTable = C_sFiltered & sTable
        sFormula = "COUNTIF" & "(" & sTable & "[" & sVarName & "], " & Chr(34) & sValue & Chr(34) & ")"
    End If

    AnalysisCount = "=" & sFormula

    Set VarNameData = Nothing
    Set TableNameData = Nothing


End Function


