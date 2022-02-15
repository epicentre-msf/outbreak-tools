Attribute VB_Name = "M_Dico"
Option Explicit
'Const C_NomFeuille = "variables"

Function CreateDicoColVar(xlsapp As Excel.Application, ssheet As String, iDicStartLine As Byte) As Scripting.Dictionary
    'extract column names in the dictionnary

    Dim D_Col As New Scripting.Dictionary
    Dim i As Integer

    D_Col.RemoveAll
    With xlsapp.Sheets(ssheet)
        i = 1
        While .Cells(iDicStartLine, i).value <> ""
            D_Col.Add xlsapp.Sheets(ssheet).Cells(iDicStartLine, i).value, i
            i = i + 1
        Wend
        Set CreateDicoColVar = D_Col
        Set D_Col = Nothing
    End With
End Function

'Function to build the data table for the dictionnary sheet in the set-up
Function CreateTabDataVar(xlsapp As Excel.Application, sSheetName As String, D_Col As Scripting.Dictionary, iDicStartLine As Byte)

    Dim i As Integer                             'iterator for the line values
    Dim j As Integer                             'iterator for the column values (values are transposed in the T_data) output
    Dim iLastLine As Integer                     'lastline of the dictionnary sheet
    Dim T_data

    With xlsapp.Sheets(sSheetName)
        iLastLine = .Cells(1, 1).End(xlDown).Row
        i = iDicStartLine                        'StartLine of the dictionnary
        j = 0
        'Each row of the table T_Data contains data on one column of the dictionnary
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
    
    'The output of the function is a variant
    CreateTabDataVar = T_data
End Function

Function LetDecString(iDecNb As Integer) As String
    'Dim iDecNb As Integer
    Dim i As Integer
    Dim sNbDeci As String
    LetDecString = ""
    i = 0
    While i < iDecNb
        sNbDeci = "0" & sNbDeci
        i = i + 1
    Wend
    LetDecString = sNbDeci

End Function

'Function to create dictionnary for choices table (the headers of the table in the choices sheet)
Function CreateDicoColChoi(xlsapp As Excel.Application, ssheet As String) As Scripting.Dictionary

    Dim D_Col As New Scripting.Dictionary
    Dim i As Integer
    i = 1
    With xlsapp.Sheets(ssheet)
        While .Cells(1, i).value <> ""
            D_Col.Add .Cells(1, i).value, i
            i = i + 1
        Wend
        Set CreateDicoColChoi = D_Col
        Set D_Col = Nothing
    End With

End Function

'Function to create the table choices (the values in the choices sheet)
Function CreateTabDataChoi(xlsapp As Excel.Application, sSheetName As String)

    Dim i As Integer                             'cpt feuille
    Dim j As Integer                             'cpt tab
    Dim iLastLine As Integer
    Dim T_data
    Dim D_Col As New Scripting.Dictionary

    Set D_Col = CreateDicoColChoi(xlsapp, sSheetName)

    With xlsapp.Sheets(sSheetName)
        iLastLine = .Cells(1, 1).End(xlDown).Row
        i = 2                                    'starting line
        j = 0                                    'one row of the table is one column in the choice sheet
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
End Function

'Extracts the list of choices for one variable
Function GetValidationName(T_data, D_Col As Scripting.Dictionary, sValidation As String) As String
    'T_Data: the choices data
    'D_Col: dictionnary of choices headers
    'sValidation:  variable
    Dim i As Integer
    Dim sWording As String
    
    i = 0
    While i <= UBound(T_data, 2)
        If T_data(D_Col("list_name") - 1, i) = sValidation Then
            If sWording = "" Then
                sWording = T_data(D_Col("label") - 1, i)
            Else
                sWording = sWording & Application.International(xlListSeparator) & T_data(D_Col("label") - 1, i)
            End If
        End If
        i = i + 1
    Wend
    GetValidationName = sWording

End Function

