Attribute VB_Name = "M_Dico"
Option Explicit
'
'Const C_NomFeuille = "variables"

Function CreateDicoColVar(xlsApp As Excel.Application, sSheet As String, iDicStartLine As Byte) As Scripting.Dictionary
    'stock l'emplacement de chaque colonne

    Dim D_Col As New Scripting.Dictionary
    Dim i As Integer

    D_Col.RemoveAll
    With xlsApp.Sheets(sSheet)
        i = 1
        While .Cells(iDicStartLine, i).value <> ""
            D_Col.Add xlsApp.Sheets(sSheet).Cells(iDicStartLine, i).value, i
            i = i + 1
        Wend
        Set CreateDicoColVar = D_Col
        Set D_Col = Nothing
    End With

End Function

Function CreateTabDataVar(xlsApp As Excel.Application, sSheetName As String, D_Col As Scripting.Dictionary, iDicStartLine As Byte)
    'construit le tableau des variables

    Dim i As Integer                             'cpt feuille
    Dim j As Integer                             'cpt tab
    Dim iLastLine As Integer
    Dim T_data
    'Dim D_col As New scripting.dictionary

    'Set D_col = CreateDicoColVar("Variables")

    With xlsApp.Sheets(sSheetName)
        iLastLine = .Cells(1, 1).End(xlDown).Row
        i = iDicStartLine                        'ligne début
        j = 0
        ReDim T_data(D_Col.Count, iLastLine - i)
        While i <= iLastLine
            'Name and labels
            T_data(D_Col("Variable name") - 1, j) = .Cells(i, D_Col("Variable name")).value
            T_data(D_Col("Main label") - 1, j) = .Cells(i, D_Col("Main label")).value
            T_data(D_Col("Sub-label") - 1, j) = .Cells(i, D_Col("Sub-label")).value
            T_data(D_Col("Note") - 1, j) = .Cells(i, D_Col("Note")).value
        
            'Sheets and sections
            T_data(D_Col("Sheet") - 1, j) = .Cells(i, D_Col("Sheet")).value
            T_data(D_Col("Main section") - 1, j) = .Cells(i, D_Col("Main section")).value
            T_data(D_Col("Sub-section") - 1, j) = .Cells(i, D_Col("Sub-section")).value
        
            'Properties/type
            T_data(D_Col("Status") - 1, j) = .Cells(i, D_Col("Status")).value
            T_data(D_Col("Personal identifier") - 1, j) = .Cells(i, D_Col("Personal identifier")).value
            T_data(D_Col("Type") - 1, j) = .Cells(i, D_Col("Type")).value
            T_data(D_Col("Control") - 1, j) = .Cells(i, D_Col("Control")).value
            T_data(D_Col("Formula") - 1, j) = .Cells(i, D_Col("Formula")).value
            T_data(D_Col("Choices") - 1, j) = .Cells(i, D_Col("Choices")).value
            T_data(D_Col("Unique") - 1, j) = .Cells(i, D_Col("Unique")).value
            T_data(D_Col("Source") - 1, j) = .Cells(i, D_Col("Source")).value
            T_data(D_Col("HXL") - 1, j) = .Cells(i, D_Col("HXL")).value
        
            'Exports
            T_data(D_Col("Export 1") - 1, j) = .Cells(i, D_Col("Export 1")).value
            T_data(D_Col("Export 2") - 1, j) = .Cells(i, D_Col("Export 2")).value
            T_data(D_Col("Export 3") - 1, j) = .Cells(i, D_Col("Export 3")).value
            T_data(D_Col("Export 4") - 1, j) = .Cells(i, D_Col("Export 4")).value
            T_data(D_Col("Export 5") - 1, j) = .Cells(i, D_Col("Export 5")).value
        
            'Validation
            T_data(D_Col("Min") - 1, j) = .Cells(i, D_Col("Min")).value
            T_data(D_Col("Max") - 1, j) = .Cells(i, D_Col("Max")).value
            T_data(D_Col("Message") - 1, j) = .Cells(i, D_Col("Message")).value
            T_data(D_Col("Alert") - 1, j) = .Cells(i, D_Col("Alert")).value
            T_data(D_Col("Branching logic") - 1, j) = .Cells(i, D_Col("Branching logic")).value
            T_data(D_Col("Conditional formatting") - 1, j) = .Cells(i, D_Col("Conditional formatting")).value
        
            j = j + 1
            i = i + 1
        Wend

    End With
    CreateTabDataVar = T_data

    'Set D_col = Nothing

End Function

Function LetDecString(iDecNb As Integer) As String

    'Dim iDecNb As Integer
    Dim i As Integer
    Dim sNbDeci As String

    LetDecString = ""

    'iDecNb = Right(T_data(D_Title("type") - 1, i), 1)
    i = 0
    While i < iDecNb
        sNbDeci = "0" & sNbDeci
        i = i + 1
    Wend
    LetDecString = sNbDeci

End Function

'                                            Choices
Function CreateDicoColChoi(xlsApp As Excel.Application, sSheet As String) As Scripting.Dictionary
    'stock l'emplacement de chaque colonne

    Dim D_Col As New Scripting.Dictionary
    Dim i As Integer

    i = 1
    With xlsApp.Sheets(sSheet)
        While .Cells(1, i).value <> ""
            D_Col.Add .Cells(1, i).value, i
            i = i + 1
        Wend
        Set CreateDicoColChoi = D_Col
        Set D_Col = Nothing
    End With

End Function

Function CreateTabDataChoi(xlsApp As Excel.Application, sSheetName As String)

    Dim i As Integer                             'cpt feuille
    Dim j As Integer                             'cpt tab
    Dim iLastLine As Integer
    Dim T_data
    Dim D_Col As New Scripting.Dictionary

    Set D_Col = CreateDicoColChoi(xlsApp, sSheetName)

    With xlsApp.Sheets(sSheetName)
        iLastLine = .Cells(1, 1).End(xlDown).Row
        i = 2                                    'ligne début
        j = 0
        ReDim T_data(D_Col.Count, iLastLine - i)
        While i <= iLastLine
            T_data(D_Col("list_name") - 1, j) = .Cells(i, D_Col("list_name")).value
            T_data(D_Col("code_num") - 1, j) = .Cells(i, D_Col("code_num")).value
            T_data(D_Col("code_alpha") - 1, j) = .Cells(i, D_Col("code_alpha")).value
            T_data(D_Col("label_short") - 1, j) = .Cells(i, D_Col("label_short")).value
            T_data(D_Col("label") - 1, j) = .Cells(i, D_Col("label")).value
    
            j = j + 1
            i = i + 1
        Wend
    End With

    CreateTabDataChoi = T_data
    Set D_Col = Nothing
    ReDim T_data(0)

End Function

Function GetValidationName(T_data, D_Col As Scripting.Dictionary, sValidation As String) As String 'attention typage sValidation... Excel fait nimp'

    Dim i As Integer
    'Dim T_data
    'Dim D_Col As scripting.dictionary
    Dim sWording As String

    'Set D_col = CreateDicoColChoi(xlsapp, C_SheetNameChoices)
    'T_data = CreateTabDataChoi(xlsapp)

    i = 0
    While i <= UBound(T_data, 2)
        If T_data(D_Col("list_name") - 1, i) = sValidation Then
            If sWording = "" Then
                sWording = T_data(D_Col("label") - 1, i)
            Else
                'sWording = sWording & ";" & T_data(D_Col("label") - 1, i)
                sWording = sWording & Application.International(xlListSeparator) & T_data(D_Col("label") - 1, i)
            End If
        End If
        i = i + 1
    Wend

    'ReDim T_data(0)
    GetValidationName = sWording
    'Set D_Col = Nothing

End Function

