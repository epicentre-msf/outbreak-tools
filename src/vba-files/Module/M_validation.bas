Attribute VB_Name = "M_validation"
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

Public Function ControlValidationFormula(sFormula As String, DictData as betterArray, DictHeaders as betterArray, VarNameData as BetterArray, IsValidation As Boolean)
    'renvoie un tableau de la d�composition de la formule envoy�e

    Dim sFormulaATest As String
    Dim sAlphaValue as string

    Dim i As Integer
    Dim iPrevBreak As Integer
 
    Dim j As Integer

    Dim FormulaData As BetterArray 'Array of formulas
    Dim SpecCharData as BetterArray 'Array of special characters
    Dim VarNameData  As BetterArray 'Array of varnames data
    Dim FormulaAlphaData as betterArray
    Dim FormulaResult as BetterArray

    Dim iNbParentO As Integer 'Number of left parenthesis
    Dim iNbParentF As Integer 'Number of right parenthesis

    Dim sLetter As String

    Dim IsError As Boolean
                              
    IsError = False
    Set FormulaData = new betterArray
    Set SpecCharData = new betterArray
    Set FormulaAlphaData = new betterArray 'Alphanumeric values of one formula
    Set FormulaResult = new betterArray 'Result of formula after going throughout the checking process
    FormulaAlphaData.lowerbound = 1

    FormulaData.fromExcelRange SheetFormulas.ListObjects(C_sTabExcelFunctions).listcolumns("ENG").range, detectlastcolumn:=False
    SpecCharData.fromExcelRange SheetFormulas.listobjects(C_sTabASCII).listcolumns("TEXT").range, detectlastcolumn:=False

    sFormulaATest = Replace(sFormula, " ", "")

    iNbParentO = 0
    iNbParentF = 0
    iPrevBreak = 1
    i = 1

    If VarNameData.Includes(sFormulaATest) Then
        FormulaAlphaData.push sFormulaATest
    Else                                         'pour les formules compos�es
        While i <= Len(sFormulaATest)
            sLetter = UCase(Mid(sFormulaATest, i, 1))
        
            If SpecCharData.includes(sLetter) Then  'si c'est un caratere s�pcial
                If sLetter = Chr(40) Then
                    iNbParentO = iNbParentO + 1
                End If
                If sLetter = Chr(41) Then
                    iNbParentF = iNbParentF + 1
                End If
                If Mid(sFormulaATest, iPrevBreak, i - iPrevBreak) <> "" Then
                    FormulaAlphaData.push Mid(sFormulaATest, iPrevBreak, i - iPrevBreak))
                End If
                iPrevBreak = i + 1
            End If
            i = i + 1
        Wend
    End If

    If iNbParentO <> iNbParentF Then
        IsError = True
    Else
        i = 1
        While i <= FormulaAlphaData.UpperBound
            sAlphaValue = FormulaAlphaData.item(i)
            If Not VarNameData.Includes(sAlphaValue) And Not FormulaData.Includes(UCase(sAlphavalue)) And Not IsNumeric(sAlphaValue) Then
                If sAlphaValue <> "" Then
                    If Asc(sAlphaValue) <> 34 Then 's'il ne s'agit pas d'une chaine de texte a afficher : c'est donc une formule volante non identifi�e
                        IsError = True
                    End If
                End If
            Else
                If VarNameData.Includes(sAlphaValue) Then 'on est pas en validation : on est fonction normale /on ne stock pas le nom de fonction puisque Excel va faire a Translation dans la bonne langue
                    ReDim Preserve T_res(j)
                    T_res(j) = T_String(i)
                    j = j + 1
                ElseIf D_Formula.Exists(T_String(i)) And IsValidation Then
                    ReDim Preserve T_res(j)
                    T_res(j) = T_String(i) & Chr(124) & LetInternationalFormula(T_String(i)) 
                    'on stock la fonction dans le format : Ancienne fonction | fonction traduite
                
                    j = j + 1
                End If
            End If
            i = i + 1
        Wend
    End If

    Set D_Name = Nothing
    Set D_Formula = Nothing
    Set D_CaracSpec = Nothing

    If Not IsError Then
        ControlValidationFormula = T_res
    End If

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

    T_Formula = SheetFormulas.listobject(C_sTabExcelFunctions).Databodyrange
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

