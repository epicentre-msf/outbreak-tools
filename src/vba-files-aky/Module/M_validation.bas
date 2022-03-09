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

Public Function ControlValidationFormula(sFormula As String, T_dataDic, D_TitleDic As Scripting.Dictionary, IsValidation As Boolean)
    'renvoie un tableau de la décomposition de la formule envoyée

    Dim sFormulaATest As String

    Dim i As Integer
    Dim iPrevBreak As Integer
    Dim T_String
    Dim j As Integer

    Dim D_Formula As Scripting.Dictionary
    Dim D_CaracSpec As Scripting.Dictionary
    Dim D_Name As Scripting.Dictionary

    Dim iNbParentO As Integer
    Dim iNbParentF As Integer

    Dim sLetter As String

    Dim IsError As Boolean
    Dim T_res
    Dim T_Formula                                'sert pour la Translation pour les validations

    IsError = False
    Set D_Name = BuildDicDic(D_TitleDic, T_dataDic)
    Set D_Formula = BuildFormulaDic
    Set D_CaracSpec = BuildCaractDic

    sFormulaATest = UCase(Replace(sFormula, " ", ""))

    iNbParentO = 0
    iNbParentF = 0
    If D_Name.Exists(sFormulaATest) Then
        'pour les noms simples
        ReDim T_String(0)
        T_String(0) = sFormulaATest
    Else                                         'pour les formules composées
        j = 0
        ReDim T_String(j)
        i = 1
        iPrevBreak = 1
        While i <= Len(sFormulaATest)
            sLetter = UCase(Mid(sFormulaATest, i, 1))
        
            If D_CaracSpec.Exists(sLetter) Then  'si c'est un caratere sépcial
                If sLetter = Chr(40) Then
                    iNbParentO = iNbParentO + 1
                End If
                If sLetter = Chr(41) Then
                    iNbParentF = iNbParentF + 1
                End If
        
                ReDim Preserve T_String(j)
                If Mid(sFormulaATest, iPrevBreak, i - iPrevBreak) <> "" Then
                    T_String(j) = UCase(Mid(sFormulaATest, iPrevBreak, i - iPrevBreak))
                End If
                iPrevBreak = i + 1
                j = j + 1
            End If
            i = i + 1
        Wend
    End If

    If iNbParentO <> iNbParentF Then
        IsError = True
    Else
        j = 0
        ReDim T_res(j)
        i = 0
        While i <= UBound(T_String)
            If Not D_Name.Exists(T_String(i)) And Not D_Formula.Exists(T_String(i)) And Not IsNumeric(T_String(i)) Then
                If T_String(i) <> "" Then
                    If Asc(T_String(i)) <> 34 Then 's'il ne s'agit pas d'une chaine de texte a afficher : c'est donc une formule volante non identifiée
                        IsError = True
                    End If
                End If
            Else
                If D_Name.Exists(T_String(i)) Then 'on est pas en validation : on est fonction normale /on ne stock pas le nom de fonction puisque Excel va faire a Translation dans la bonne langue
                    ReDim Preserve T_res(j)
                    T_res(j) = T_String(i)
                    j = j + 1
                ElseIf D_Formula.Exists(T_String(i)) And IsValidation Then
                    ReDim Preserve T_res(j)
                    T_res(j) = T_String(i) & Chr(124) & LetInternationalFormula(T_String(i)) 'on stock la fonction dans le format : Ancienne fonction | fonction traduite
                
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

    T_Formula = [T_xlsfonctions]
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

