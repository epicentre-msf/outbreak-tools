Attribute VB_Name = "DesignerValidation"
Option Explicit

Private Function BuildDicDic(D_title As Scripting.Dictionary, T_data) As Scripting.Dictionary

    Dim i As Integer
    Dim D_Nom As New Scripting.Dictionary

    D_Nom.RemoveAll
    i = 2
    While i <= UBound(T_data, 2)
        'D_Nom.Add Sheets("Dictionary").Cells(i, 1).Value, i
        D_Nom.Add UCase(T_data(D_title("Variable name") - 1, i)), i
        i = i + 1
    Wend
    Set BuildDicDic = D_Nom

End Function

Private Function BuildFormulaDic() As Scripting.Dictionary

    Dim i As Integer
    Dim T_Form
    Dim D_formule As New Scripting.Dictionary

    T_Form = [T_xlsfonctions]

    i = 1
    D_formule.RemoveAll
    While i <= UBound(T_Form, 1)
        D_formule.Add UCase(T_Form(i, 2)), i
        i = i + 1
    Wend
    Set BuildFormulaDic = D_formule

End Function

Private Function BuildCaractDic() As Scripting.Dictionary

    Dim i As Integer
    Dim T_Carac
    Dim D_CaracSpec As New Scripting.Dictionary

    T_Carac = [T_ascii]

    D_CaracSpec.RemoveAll
    i = 1
    While i <= UBound(T_Carac, 1)
        D_CaracSpec.Add T_Carac(i, 2), T_Carac(i, 2)
        i = i + 1
    Wend
    Set BuildCaractDic = D_CaracSpec

End Function


'Testing the ControlValidationFormula for debugging

Sub TestValidation()

    Dim sFormula As String
    Dim iSheetStartLine As Integer
    Dim VarnameData As New BetterArray
    Dim ColumnIndexData As New BetterArray
    Dim IsValidation As Boolean
    Dim FormulaData As New BetterArray
    Dim SpecCharData As New BetterArray
    
    FormulaData.FromExcelRange SheetFormulas.ListObjects(C_sTabExcelFunctions).ListColumns("ENG").DataBodyRange, DetectLastColumn:=False
    SpecCharData.FromExcelRange SheetFormulas.ListObjects(C_sTabASCII).ListColumns("TEXT").DataBodyRange, DetectLastColumn:=False
    
    VarnameData.Push "date_notification", "var2", "deceased"
    ColumnIndexData.Push 5, 5, 3
    iSheetStartLine = 1
    sFormula = "IF(ISBLANK(date_notification)," & Chr(34) & Chr(34) & ",EPIWEEK(date_notification))"
    IsValidation = False
    
    Debug.Print DesControlValidationFormula(sFormula, VarnameData, ColumnIndexData, FormulaData, SpecCharData)

End Sub

Public Function DesControlValidationFormula(sFormula As String, VarnameData As BetterArray, _
                                         ColumnIndexData As BetterArray, FormulaData As BetterArray, _
                                         SpecCharData As BetterArray) As String
    'Returns a string of cleared formula

    DesControlValidationFormula = ""

    Dim sFormulaATest As String                  'same formula, with all the spaces replaced with
    Dim sAlphaValue As String                    'Alpha numeric values in a formula
    Dim sLetter As String                        'counter for every letter in one formula
    Dim scolAddress As String                    'address of one column used in a formula

    Dim FormulaAlphaData As BetterArray          'Table of alphanumeric data in one formula
    
    Dim i As Integer
    Dim iPrevBreak As Integer
    Dim iNbParentO As Integer                    'Number of left parenthesis
    Dim iNbParentF As Integer                    'Number of right parenthesis
    Dim icolNumb As Integer                      'Column number on one sheet of one column used in a formual
   

    Dim isError As Boolean
    Dim OpenedQuotes As Boolean                  'Test if the formula has opened some quotes
    Dim QuotedCharacter As Boolean
    Dim NoErrorAndNoEnd As Boolean
    Set FormulaAlphaData = New BetterArray       'Alphanumeric values of one formula

    FormulaAlphaData.LowerBound = 1

    'squish the formula (removing multiple spaces) to avoid problems related to
    'space collapsing and upper/lower cases
    sFormulaATest = "(" & Application.WorksheetFunction.Trim(sFormula) & ")"

    iNbParentO = 0                               'Number of open brakets
    iNbParentF = 0                               'Number of closed brackets
    iPrevBreak = 1
    OpenedQuotes = False
    NoErrorAndNoEnd = True
    QuotedCharacter = False
    i = 1

    If VarnameData.Includes(sFormulaATest) Then
        DesControlValidationFormula = sFormulaATest
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
                    If Not VarnameData.Includes(LCase(sAlphaValue)) And Not FormulaData.Includes(UCase(sAlphaValue)) And Not IsNumeric(sAlphaValue) Then
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
                        If VarnameData.Includes(sAlphaValue) Then 'It is a variable name, I will track its column
                            icolNumb = ColumnIndexData.Item(VarnameData.IndexOf(sAlphaValue))
                            sAlphaValue = Cells(C_eStartLinesLLData + 1, icolNumb).Address(False, True)
                        ElseIf FormulaData.Includes(UCase(sAlphaValue)) Then 'It is a formula, excel will do the translation for us
                            sAlphaValue = LetInternationalFormula(sAlphaValue)
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
        DesControlValidationFormula = FormulaAlphaData.ToString(Separator:="", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
    End If
    
    Set FormulaAlphaData = Nothing

End Function

Public Function IsAFunction(sLib As String) As Boolean

    Dim D_Fonction As Scripting.Dictionary

    Set D_Fonction = BuildFormulaDic
    IsAFunction = False
    If D_Fonction.Exists(UCase(Replace(sLib, " ", ""))) Then
        IsAFunction = True
    End If
    Set D_Fonction = Nothing

End Function

Public Function LetInternationalFormula(sFormula As Variant)

    Dim T_Formula
    Dim i As Integer

    T_Formula = SheetFormulas.ListObjects(C_sTabExcelFunctions).DataBodyRange
    i = 1
    While i < UBound(T_Formula, 1) And UCase(T_Formula(i, 2)) <> UCase(sFormula)
        i = i + 1
    Wend
    If UCase(T_Formula(i, 2)) = UCase(sFormula) Then
        Select Case Application.International(xlCountryCode)
        Case 33                                  'FR
            LetInternationalFormula = T_Formula(i, 1)
        Case 44, 1                               'EN US
            LetInternationalFormula = T_Formula(i, 2)
        Case 34                                  'ES
            LetInternationalFormula = T_Formula(i, 3)
        End Select
    End If
    ReDim T_Formula(0)

End Function

