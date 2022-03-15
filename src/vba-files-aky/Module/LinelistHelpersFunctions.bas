Attribute VB_Name = "LinelistHelpersFunctions"
Option Explicit

'Clear a String to remove inconsistencies
Public Function ClearString(ByVal sString As String, Optional bremoveHiphen As Boolean = True) As String
    Dim sValue As String
    sValue = sString
    If bremoveHiphen Then
        sValue = Replace(sValue, "?", " ")
        sValue = Replace(sValue, "-", " ")
        sValue = Replace(sValue, "_", " ")
        sValue = Replace(sValue, "/", " ")
    End If
    sValue = Application.WorksheetFunction.Trim(sValue)
    ClearString = LCase(sValue)
End Function

'Get the headers of one sheet from one line
Function GetHeaders(Wkb As Workbook, sSheet As String, StartLine As Byte) As BetterArray
    'Extract column names in one sheet starting from one line
    Dim Headers As BetterArray
    Dim i As Integer
    Set Headers = New BetterArray
    Headers.LowerBound = 1
    Dim sValue As String
    
    With Wkb.Worksheets(sSheet)
        i = 1
        While .Cells(StartLine, i).value <> ""
        'Clear the values in the sheet when adding thems
            sValue = .Cells(StartLine, i).value 'The argument is passed byval to clearstring
            sValue = ClearString(sValue)
            Headers.Push sValue
            i = i + 1
        Wend
        Set GetHeaders = Headers
        'Clear everything
        Set Headers = Nothing
    End With
End Function

'Get the data from one sheet starting from one line
Function GetData(Wkb As Workbook, sSheetName As String, StartLine As Byte) As BetterArray
    Dim Data As BetterArray
    Set Data = New BetterArray
    Data.LowerBound = 1
    Data.FromExcelRange Wkb.Worksheets(sSheetName).Cells(StartLine, 1), DetectLastRow:=True, DetectLastColumn:=True
    'The output of the function is a variant
    Set GetData = Data
    Set Data = Nothing
End Function


'Get the list of validations from the Choices data
Function GetValidationList(ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, sValidation As String) As String

    Dim iChoiceIndex As Integer
    Dim iChoiceLastIndex As Integer
    Dim i As Integer 'iterator to get the values
    Dim sValidationList As String 'Validation List

    sValidationList = ""

    iChoiceIndex = ChoicesListData.IndexOf(sValidation)
    iChoiceLastIndex = ChoicesListData.LastIndexOf(sValidation)

    If (iChoiceIndex > 0) Then
        sValidationList = ChoicesLabelsData.Items(iChoiceIndex)
        For i = iChoiceIndex + 1 To iChoiceLastIndex
            sValidationList = sValidationList & Application.International(xlListSeparator) & ChoicesLabelsData.Items(i)
        Next
    End If

    GetValidationList = sValidationList
End Function


Public Function IsEmptyTable(T_aTest) As Boolean

    Dim test As Variant

    IsEmptyTable = False
    On Error GoTo crash
    test = UBound(T_aTest)
    On Error GoTo 0
    Exit Function

crash:
    IsEmptyTable = True

End Function

Sub QuickSort(T_aTrier, ByVal lngMin As Long, ByVal lngMax As Long)
 
    Dim strMidValue As String
    Dim lngHi As Long
    Dim lngLo As Long
    Dim lngIndex As Long
  
    If lngMin >= lngMax Then Exit Sub
  
    ' Valeur de partionnement
    lngIndex = Int((lngMax - lngMin + 1) * Rnd + lngMin)
    strMidValue = T_aTrier(lngIndex)
 
    ' Echanger les valeurs
    T_aTrier(lngIndex) = T_aTrier(lngMin)
 
    lngLo = lngMin
    lngHi = lngMax
    Do
        ' Chercher,   partir de lngHi, une valeur < strMidValue
        Do While T_aTrier(lngHi) >= strMidValue
            lngHi = lngHi - 1
            If lngHi <= lngLo Then Exit Do
        Loop
        If lngHi <= lngLo Then
            T_aTrier(lngLo) = strMidValue
            Exit Do
        End If
 
        ' Echanger les valeurs lngLo et lngHi
        T_aTrier(lngLo) = T_aTrier(lngHi)
 
        ' Chercher   partir de lngLo une valeur >= strMidValue
        lngLo = lngLo + 1
        Do While T_aTrier(lngLo) < strMidValue
            lngLo = lngLo + 1
            If lngLo >= lngHi Then Exit Do
        Loop
        If lngLo >= lngHi Then
            lngLo = lngHi
            T_aTrier(lngHi) = strMidValue
            Exit Do
        End If
 
        ' Echanger les valeurs lngLo et lngHi
        T_aTrier(lngHi) = T_aTrier(lngLo)
    Loop
 
    ' Trier les 2 sous-T_aTrieres
    QuickSort T_aTrier, lngMin, lngLo - 1
    QuickSort T_aTrier, lngLo + 1, lngMax
    
End Sub

Public Function DesLoadFile(Optional sFilters As String) As String 'lla

    Dim fDialog As Office.FileDialog

    DesLoadFile = ""
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your file"               'MSG_ChooseFile
        .Filters.Clear
        .Filters.Add "Feuille de calcul Excel", sFilters 'MSG_excel file

        If .show = True Then
            DesLoadFile = .SelectedItems(1)
        End If
    End With
    Set fDialog = Nothing
End Function


Public Function DesLoadFolder() As String

    Dim fDialog As Office.FileDialog

    LoadFolder = ""
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your directory"          'MSG_ChooseDir
        .Filters.Clear
    
        If .show = True Then
            LoadFolder = .SelectedItems(1)
        End If
    End With
    Set fDialog = Nothing

End Function

Public Function CleanSpecLettersInName(sName As String) As String 'supp tous les caract sp ciaux du nom

    Dim T_Caract
    Dim i As Integer
    Dim sRes As String

    sRes = sName
    T_Caract = [T_ascii]
    i = 1
    While i <= UBound(T_Caract, 1)
        sName = Replace(sName, T_Caract(i, 2), "")
        i = i + 1
    Wend
    CleanSpecLettersInName = sName

End Function

'                                                                       '
'_________________________ Liste des fonctions _________________________'

Public Function Epiweek(jour As Long) As Long
    
    Dim annee As Long
    
    Dim Jour0_2014, Jour0_2015, Jour0_2016, Jour0_2017, Jour0_2018, Jour0_2019, Jour0_2020, Jour0_2021, Jour0_2022 As Long

    Jour0_2014 = 41638
    Jour0_2015 = 42002
    Jour0_2016 = 42366
    Jour0_2017 = 42730
    Jour0_2018 = 43101
    Jour0_2019 = 43465
    Jour0_2020 = 43829
    Jour0_2021 = 44193
    Jour0_2022 = 44557
    annee = Year(jour)
    
    Select Case annee
    Case 2014
        Epiweek = 1 + Int((jour - Jour0_2014) / 7)
    Case 2015
        Epiweek = 1 + Int((jour - Jour0_2015) / 7)
    Case 2016
        Epiweek = 1 + Int((jour - Jour0_2016) / 7)
    Case 2017
        Epiweek = 1 + Int((jour - Jour0_2017) / 7)
    Case 2018
        Epiweek = 1 + Int((jour - Jour0_2018) / 7)
    Case 2019
        Epiweek = 1 + Int((jour - Jour0_2019) / 7)
    Case 2020
        Epiweek = 1 + Int((jour - Jour0_2020) / 7)
    Case 2021
        Epiweek = 1 + Int((jour - Jour0_2021) / 7)
    Case 2022
        Epiweek = 1 + Int((jour - Jour0_2022) / 7)
    End Select
    
End Function

'This function gets unique values from a table of two dimensions converted to a BetterArray table
' Get unique values from a table on two dimensions (or more)
Public Function GetUnique(ByVal T_table As Variant, Optional ByVal col1 As Integer = -99, Optional ByVal col2 As Integer = -99, Optional ByVal index As Variant) As Variant

    Dim i, k, j As Long                          'for the first line
    Dim outCol As New Collection                 'I will stock pair values here
    Dim bindValues                               'all binded values
    Dim outTable
    Dim indexCols
    
    'If the table is empty, return empty table
    If IsEmptyTable(T_table) Then
        ReDim outTable(0)
        GetUnique = outTable
        Exit Function
    End If
    
    'If you give nothing, I will check on all the columns
    If col1 = -99 And col2 = -99 And IsEmptyTable(index) Then
        ' you can end up here with a one dimensional table, we need to be sure we have two dimensional table
        ReDim indexCols(UBound(T_table, 2))
        i = 1
        While i <= UBound(indexCols)
            indexCols(i) = i
            i = i + 1
        Wend
    ElseIf col2 = -99 And IsEmptyTable(index) Then
        ReDim indexCols(1)
        indexCols(1) = col1
    ElseIf IsEmptyTable(index) Then
        ReDim indexCols(2)
        indexCols(1) = col1
        indexCols(2) = col2
    ElseIf Not IsEmptyTable(index) Then
        indexCols = index
    End If
    
    ' Check the table index is not empty before entering the whole cycle
    If Not IsEmptyTable(indexCols) Then
    
        'Stock elements in a table by binding them I guess the binding character will be most
        'of the time absent from my data. The binding character here is (&123&;
        
        'Bind everything in the indexCols
        
        ReDim bindValues(UBound(T_table))
        i = 1
        While i <= UBound(T_table)
            k = 1
            bindValues(i) = ""
            While k <= UBound(indexCols)
                bindValues(i) = bindValues(i) & "(&123&;" & CStr(T_table(i, indexCols(k)))
                k = k + 1
            Wend
            i = i + 1
        Wend
        
        On Error Resume Next
        'Now quick sort the table first
        Call QuickSort(bindValues, LBound(bindValues), UBound(bindValues))
        On Error GoTo 0
        
        k = 1
        'adding the first unique values
        While k <= UBound(indexCols)
            outCol.Add Split(bindValues(1), "(&123&;")(k)
            'Count the number items
            k = k + 1
        Wend
        
        i = 1
        ' adding the other unique values
        While i < UBound(bindValues)
            If (bindValues(i) <> bindValues(i + 1)) Then
                k = 1
                While k <= UBound(indexCols)
                    outCol.Add Split(bindValues(i + 1), "(&123&;")(k)
                    k = k + 1
                Wend
            End If
            i = i + 1
        Wend
        'Returning the table; (outCol.Cout / UBound(indexCols) is the number of unique values
        ReDim outTable((outCol.Count / UBound(indexCols)), UBound(indexCols))
        
        i = 1                                    'will slice over unique number of values
        j = 1                                    'will slice over all the values of outCol
        While i <= (outCol.Count / UBound(indexCols))
            k = 1
            While (k <= UBound(indexCols))
                outTable(i, k) = outCol.Item(j)
                k = k + 1
                j = j + 1
            Wend
            i = i + 1
        Wend
        GetUnique = outTable
    Else
        'If the column table is empty, return an empty table
        ReDim outTable(0)
        GetUnique = outTable
    End If
End Function

' Function to filter on value of one column, on a two dimensional array
Public Function GetFilter(ByVal T_table As BetterArray, iCol As Integer, sValue As String) As BetterArray

    Dim targetColumn As BetterArray
    Dim fCol As Long                             'First and last columns
    Dim lCol As Long
    Dim i As Long
    Dim filteredTable As New BetterArray
    
    Set targetColumn = New BetterArray
    Set filteredTable = New BetterArray
    
    'target column items
    targetColumn.Items = T_table.ExtractSegment(, ColumnIndex:=iCol)
    targetColumn.Sort
    targetColumn.LowerBound = 1
    fCol = targetColumn.IndexOf(sValue)
    lCol = targetColumn.LastIndexOf(sValue)
    targetColumn.Clear
    Set targetColumn = Nothing
    
    filteredTable.Clear
    filteredTable.LowerBound = 1
    'Extract the lines of table for each of values found
    If fCol > 0 And lCol > 0 Then
        For i = fCol To lCol
            filteredTable.Push T_table.Item(i)
        Next
        Set GetFilter = filteredTable.Clone
    Else
        'return the whole table if you where not able to find a match
        Set GetFilter = T_table.Clone
    End If
End Function


