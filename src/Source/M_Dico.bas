Attribute VB_Name = "M_Dico"
Option Explicit
'
'Const C_NomFeuille = "variables"

Function CreateDicoColVar(xlsApp As Object, sFeuille As String, iLigDebDico As Byte) As Scripting.Dictionary
'stock l'emplacement de chaque colonne

Dim D_Col As New Scripting.Dictionary
Dim i As Integer

With xlsApp.Sheets(sFeuille)
    i = 1
    While .Cells(iLigDebDico, i).Value <> ""
        D_Col(.Cells(iLigDebDico, i).Value) = i
        i = i + 1
    Wend
    Set CreateDicoColVar = D_Col
    Set D_Col = Nothing
End With

End Function

Function CreateTabDataVar(xlsApp As Object, sNomFeuille As String, D_Col As Scripting.Dictionary, iLigDebDico As Byte)
'construit le tableau des variables

Dim i As Integer    'cpt feuille
Dim j As Integer    'cpt tab
Dim iDerLigne As Integer
Dim T_data
'Dim D_col As New Scripting.Dictionary

'Set D_col = CreateDicoColVar("Variables")

With xlsApp.Sheets(sNomFeuille)
    iDerLigne = .Cells(1, 1).End(xlDown).Row
    i = iLigDebDico   'ligne début
    j = 0
    ReDim T_data(D_Col.Count, iDerLigne - i)
    While i <= iDerLigne
        'Name and labels
        T_data(D_Col("name") - 1, j) = .Cells(i, D_Col("name")).Value
        T_data(D_Col("label_1") - 1, j) = .Cells(i, D_Col("label_1")).Value
        T_data(D_Col("label_2") - 1, j) = .Cells(i, D_Col("label_2")).Value
        T_data(D_Col("note") - 1, j) = .Cells(i, D_Col("note")).Value
        
        'Sheets and sections
        T_data(D_Col("form_name") - 1, j) = .Cells(i, D_Col("form_name")).Value
        T_data(D_Col("section_1") - 1, j) = .Cells(i, D_Col("section_1")).Value
        T_data(D_Col("section_2") - 1, j) = .Cells(i, D_Col("section_2")).Value
        
        'Properties/type
        T_data(D_Col("mandatory") - 1, j) = .Cells(i, D_Col("mandatory")).Value
        T_data(D_Col("personal_identifier") - 1, j) = .Cells(i, D_Col("personal_identifier")).Value
        T_data(D_Col("type") - 1, j) = .Cells(i, D_Col("type")).Value
        T_data(D_Col("control") - 1, j) = .Cells(i, D_Col("control")).Value
        T_data(D_Col("formula") - 1, j) = .Cells(i, D_Col("formula")).Value
        T_data(D_Col("choices") - 1, j) = .Cells(i, D_Col("choices")).Value
        T_data(D_Col("visible") - 1, j) = .Cells(i, D_Col("visible")).Value             '??!!
        T_data(D_Col("unique") - 1, j) = .Cells(i, D_Col("unique")).Value
        T_data(D_Col("source") - 1, j) = .Cells(i, D_Col("source")).Value
        T_data(D_Col("hxl") - 1, j) = .Cells(i, D_Col("hxl")).Value
        
        'Exports
        T_data(D_Col("export:MSF") - 1, j) = .Cells(i, D_Col("export:MSF")).Value
        T_data(D_Col("export:MOH") - 1, j) = .Cells(i, D_Col("export:MOH")).Value
        
        'Validation
        T_data(D_Col("min") - 1, j) = .Cells(i, D_Col("min")).Value
        T_data(D_Col("max") - 1, j) = .Cells(i, D_Col("max")).Value
        T_data(D_Col("validation_formula") - 1, j) = .Cells(i, D_Col("validation_formula")).Value
        T_data(D_Col("validation_alert") - 1, j) = .Cells(i, D_Col("validation_alert")).Value
        T_data(D_Col("branching_logic") - 1, j) = .Cells(i, D_Col("branching_logic")).Value
        T_data(D_Col("conditional_formatting") - 1, j) = .Cells(i, D_Col("conditional_formatting")).Value
        
        j = j + 1
        i = i + 1
    Wend

End With
CreateTabDataVar = T_data

'Set D_col = Nothing

End Function

Function RenvoiChaineDecimal(iNbDeci As Integer) As String

'Dim iNbDeci As Integer
Dim i As Integer
Dim sNbDeci As String

RenvoiChaineDecimal = ""

'iNbDeci = Right(T_data(D_entete("type") - 1, i), 1)
i = 0
While i < iNbDeci
    sNbDeci = "0" & sNbDeci
    i = i + 1
Wend
RenvoiChaineDecimal = sNbDeci

End Function

'                                            Choices
Function CreateDicoColChoi(xlsApp As Object, sFeuille As String) As Scripting.Dictionary
'stock l'emplacement de chaque colonne

Dim D_Col As New Scripting.Dictionary
Dim i As Integer

i = 1
With xlsApp.Sheets(sFeuille)
    While .Cells(1, i).Value <> ""
        D_Col.Add .Cells(1, i).Value, i
        i = i + 1
    Wend
    Set CreateDicoColChoi = D_Col
    Set D_Col = Nothing
End With

End Function

Function CreateTabDataChoi(xlsApp As Object, sNomFeuille As String)

Dim i As Integer    'cpt feuille
Dim j As Integer    'cpt tab
Dim iDerLigne As Integer
Dim T_data
Dim D_Col As New Scripting.Dictionary

Set D_Col = CreateDicoColChoi(xlsApp, sNomFeuille)

With xlsApp.Sheets(sNomFeuille)
    iDerLigne = .Cells(1, 1).End(xlDown).Row
    i = 2   'ligne début
    j = 0
    ReDim T_data(D_Col.Count, iDerLigne - i)
    While i <= iDerLigne
        T_data(D_Col("validation") - 1, j) = .Cells(i, D_Col("validation")).Value
        T_data(D_Col("code_num") - 1, j) = .Cells(i, D_Col("code_num")).Value
        T_data(D_Col("code_alpha") - 1, j) = .Cells(i, D_Col("code_alpha")).Value
        T_data(D_Col("label_short") - 1, j) = .Cells(i, D_Col("label_short")).Value
        T_data(D_Col("label") - 1, j) = .Cells(i, D_Col("label")).Value
    
        j = j + 1
        i = i + 1
    Wend
End With

CreateTabDataChoi = T_data
Set D_Col = Nothing
ReDim T_data(0)

End Function

Function GetChaineValidation(T_data, D_Col As Scripting.Dictionary, sValidation) As String      'attention typage sValidation... Excel fait nimp'

Dim i As Integer
'Dim T_data
'Dim D_Col As Scripting.Dictionary
Dim sLabel As String

'Set D_col = CreateDicoColChoi(xlsapp, C_NomFeuilleChoix)
'T_data = CreateTabDataChoi(xlsapp)

i = 0
While i <= UBound(T_data, 2)
    If T_data(D_Col("validation") - 1, i) = sValidation Then
        If sLabel = "" Then
            sLabel = T_data(D_Col("label") - 1, i)
        Else
            sLabel = sLabel & ";" & T_data(D_Col("label") - 1, i)
        End If
    End If
    i = i + 1
Wend

'ReDim T_data(0)
GetChaineValidation = sLabel
'Set D_Col = Nothing

End Function
