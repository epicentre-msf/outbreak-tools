Attribute VB_Name = "Helpers"

'Basic Helper functions used in the creation of the linelist and other stuffs
'Most of them are explicit functions. Contains all the ancillary sub/
'Functions used when creating the linelist and also in the linelist
'itself

Option Explicit


Public Function GetColor(sColorCode As String)

    Select Case sColorCode
    Case "BlueEpi"
        GetColor = RGB(45, 85, 158)
    Case "RedEpi"
        GetColor = RGB(252, 228, 214)
    Case "LightBlueTitle"
        GetColor = RGB(217, 225, 242)
    Case "DarkBlueTitle"
        GetColor = RGB(142, 169, 219)
    Case "Grey"
        GetColor = RGB(235, 232, 232)
    Case "Green"
        GetColor = RGB(198, 224, 180)
    Case "Orange"
        GetColor = RGB(248, 203, 173)
    Case "White"
        GetColor = RGB(255, 255, 255)
    Case "MainSecBlue"
        GetColor = RGB(47, 117, 181)
    Case "SubSecBlue"
        GetColor = RGB(221, 235, 247)
    Case "SubLabBlue"
        GetColor = RGB(142, 169, 219)
    End Select

End Function


'This will set the actual application properties to be able to work correctly
Public Sub BeginWork(xlsapp As Excel.Application, Optional bvisbility As Boolean = True)
    xlsapp.ScreenUpdating = False
   ' xlsapp.DisplayAlerts = False
    'xlsapp.Calculation = xlCalculationManual
    'xlsapp.Cursor = xlWait
End Sub


Public Sub EndWork(xlsapp As Excel.Application, Optional bvisbility As Boolean = True)
    xlsapp.ScreenUpdating = True
    'xlsapp.DisplayAlerts = True
    'xlsapp.Cursor = xlDefault
    'xlsapp.Calculation = xlCalculationAutomatic
End Sub

'Load files and folders
Public Function LoadFile(Optional sFilters As String) As String 'lla

    Dim fDialog As Office.FileDialog

    LoadFile = ""
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your file"               'MSG_ChooseFile
        .Filters.Clear
        .Filters.Add "Feuille de calcul Excel", sFilters '"*.xlsx" ', *.xlsm, *.xlsb,  *.xls" 'MSG_ExcelFile'lla

        If .Show = True Then
            LoadFile = .SelectedItems(1)
        End If
    End With
    Set fDialog = Nothing

End Function

Public Function LoadFolder() As String

    Dim fDialog As Office.FileDialog

    LoadFolder = ""
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your directory"          'MSG_ChooseDir
        .Filters.Clear
    
        If .Show = True Then
            LoadFolder = .SelectedItems(1)
        End If
    End With
    Set fDialog = Nothing

End Function

'Get the file extension of a string
'Get the file extension of a file
Private Function GetFileExtension(sString As String) As String
    
    GetFileExtension = ""
    
    Dim iDotPos As Integer
    Dim sExt As String 'extension
    'Find the position of the dot at the end
    iDotPos = InStrRev(sString, ".")
    
    sExt = Right(sString, Len(sString) - iDotPos)
    
    If (sExt <> "") Then
        GetFileExtension = sExt
    End If
    
End Function

'Check if a Workbook is Opened

Public Function IsWkbOpened(sName As String) As Boolean
    Dim oWkb As Workbook                         'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set oWkb = Application.Workbooks.Item(sName)
    IsWkbOpened = (Not oWkb Is Nothing)
    On Error GoTo 0
End Function


'Write lines for borders

Public Sub WriteBorderLines(oRange As Range)

    Dim i As Integer
    For i = 7 To 10
        With oRange.Borders(i)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next

End Sub


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

'Get headers and data from one worksheet of a workbook

'Get the headers of one sheet from one line (probablly the first line)
'The headers are cleaned

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



'Set validation list on a range

Sub SetValidation(oRange As Range, sValidList As String, sAlertType As Byte, sMessage As String)

    With oRange.Validation
        .Delete
        Select Case sAlertType
        Case 1                                   '"error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=sValidList
        Case 2                                   '"warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=sValidList
        Case Else                                'for all the others, add an information alert
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=sValidList
        End Select
        
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = sMessage
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'Get the validation list using Choices data and choices labels
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


Function GetValidationType(sValidationType As String) As Byte

    GetValidationType = 3                    'list of validation info, warning or error
    If sValidationType <> "" Then
        Select Case LCase(sValidationType)
        Case "warning"
            GetValidationType = 2
        Case "error"
            GetValidationType = 1
        End Select
    End If
    
End Function

'Epicemiological week function

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
        ' Chercher, ï¿½ partir de lngHi, une valeur < strMidValue
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
 
        ' Chercher ï¿½ partir de lngLo une valeur >= strMidValue
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

Public Function IsEmptyTable(T_aTest) As Boolean

    Dim Test As Variant

    IsEmptyTable = False
    On Error GoTo crash
    Test = UBound(T_aTest)
    On Error GoTo 0
    Exit Function

crash:
    IsEmptyTable = True

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


'Find the index of sValue on column iCol of a BetterArray T_table
Public Function FindIndex(T_table As BetterArray, iCol As Integer, sValue As String) As Integer
    Dim T_data As BetterArray
    Set T_data = New BetterArray
    T_data.Items = T_table.ExtractSegment(ColumnIndex:=iCol)
    FindIndex = T_data.IndexOf(sValue)
    Set T_data = Nothing
End Function


'Find the value of one variable for one column in a sheet

Public Function FindDicColumnValue(sVarname, sColumn)
    FindDicColumnValue = ""
    Dim VarNameData As BetterArray
    Dim sListObjectName As String
    Set VarNameData = New BetterArray
    
    VarNameData.LowerBound = 2 'Because the first line of dictionary is the header
    sListObjectName = "o" & ClearString(C_sParamSheetDict)
    With ThisWorkbook.Worksheets(C_sParamSheetDict)
        VarNameData.FromExcelRange .ListObjects(sListObjectName).ListColumns(C_sDictHeaderVarName).DataBodyRange, _
                                     DetectLastRow:=True, DetectLastColumn:=False
                If VarNameData.Includes(sVarname) Then
                    FindDicColumnValue = .Cells(VarNameData.IndexOf(sVarname), .ListObjects(sListObjectName).ListColumns(sColumn).index).value
                End If
    End With
    Set VarNameData = Nothing
End Function



Sub StatusBar_Updater(sCpte As Single)
'increase the status progressBar

    Dim CurrentStatus As Integer
    Dim pctDone As Integer

    CurrentStatus = (C_iNumberOfBars) * Round(sCpte / 100, 1)
    Application.StatusBar = "[" & String(CurrentStatus, "|") & Space(C_iNumberOfBars - CurrentStatus) & "]" & " " & CInt(sCpte) & "%"  & TranslateMSg("MSG_BuildLL")

    DoEvents
    
End Sub






