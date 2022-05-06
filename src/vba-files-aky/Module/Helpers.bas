Attribute VB_Name = "Helpers"
'Basic Helper functions used in the creation of the linelist and other stuffs
'Most of them are explicit functions. Contains all the ancillary sub/
'Functions used when creating the linelist and also in the linelist
'itself

Option Explicit
Public DebugMode


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


Public Sub ProtectSheet(Optional pwd As String = C_sLLPassword)
    If Not DebugMode Then
        ActiveSheet.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                         AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                         AllowFormattingColumns:=True
    End If

End Sub


'This will set the actual application properties to be able to work correctly
Public Sub BeginWork(xlsapp As Excel.Application, Optional bstatusbar As Boolean = True)
    xlsapp.ScreenUpdating = False
    xlsapp.DisplayAlerts = False
    xlsapp.Calculation = xlCalculationManual
    xlsapp.DisplayStatusbar = bstatusbar
End Sub


Public Sub EndWork(xlsapp As Excel.Application, Optional bstatusbar As Boolean = True)
    xlsapp.ScreenUpdating = True
    xlsapp.DisplayAlerts = True
    xlsapp.Calculation = xlCalculationAutomatic
    xlsapp.DisplayStatusbar = bstatusbar
End Sub

'Load files and folders
Public Function LoadFile(sFilters As String, sType As String) As String

    Dim fDialog As Office.FileDialog

    LoadFile = ""
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your file"               'MSG_ChooseFile
        .Filters.Clear
        .Filters.Add "Feuille de calcul Excel", sFilters '"*.xlsx" ', *.xlsm, *.xlsb,  *.xls" 'MSG_ExcelFile

        If .show = True Then
            LoadFile = .SelectedItems(1)
            If sType = "Setup" Then Call ImportLanguage(LoadFile)
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

        If .show = True Then
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

Sub SetValidation(oRange As Range, sValidList As String, sAlertType As Byte, Optional sMessage As String = vbNullString)

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
        .errorTitle = ""
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

'Move a plage of data from the setup sheet to the designer sheet
Public Sub MoveData(SourceWkb As Workbook, DestWkb As Workbook, sSheetName As String, sStartCell As Integer)

    Dim sData As BetterArray
    Dim DestWksh As Worksheet
    Dim sheetExists As Boolean

    Set sData = New BetterArray
    sData.FromExcelRange SourceWkb.Worksheets(sSheetName).Range("A" & CStr(sStartCell)), DetectLastRow:=True, DetectLastColumn:=True
    sheetExists = False

    For Each DestWksh In DestWkb.Worksheets
        If DestWksh.Name = sSheetName Then sheetExists = True
    Next

    'Clear the contents if the sheet exists, or create a new sheet if Not
    If sheetExists Then
        DestWkb.Worksheets(sSheetName).Activate
        Cells.Clear
    Else
        DestWkb.Worksheets.Add.Name = sSheetName
    End If

    'Copy the data Now
    sData.ToExcelRange DestWkb.Worksheets(sSheetName).Range("A1")
    DestWkb.Worksheets(sSheetName).Visible = xlSheetHidden
    Set sData = Nothing
End Sub


'Filter a table listobject on one condition and get the values of that table or all the unique values of one column
Public Function FilterLoTable(lo As ListObject, iFiltindex1 As Integer, sValue1 As String, _
                             Optional iFiltindex2 As Integer = 0, Optional sValue2 As String = vbNullString, _
                             Optional iFiltindex3 As Integer = 0, Optional sValue3 As String = vbNullString, _
                             Optional returnIndex As Integer = -99, _
                             Optional bAllData As Boolean = True) As BetterArray
    Dim Rng As Range
    Dim Data As BetterArray
    Dim breturnAllData As Boolean

    With lo.Range

        .AutoFilter Field:=iFiltindex1, Criteria1:=sValue1

        'Add other Filters if required
        If iFiltindex2 > 0 Then
            .AutoFilter Field:=iFiltindex2, Criteria1:=sValue2
        End If

        If iFiltindex3 > 0 Then
            .AutoFilter Field:=iFiltindex3, Criteria1:=sValue3
        End If

    End With

    Set Rng = lo.Range.SpecialCells(xlCellTypeVisible)

    If returnIndex > 0 Then
        breturnAllData = False
    ElseIf bAllData Then
        breturnAllData = True
    Else
        breturnAllData = True
    End If

    'Copy and paste to temp
    With ThisWorkbook.Worksheets(C_sSheetTemp)
            .Visible = xlSheetHidden
            .Cells.Clear

            Rng.Copy Destination:=.Cells(1, 1)

            Set Data = New BetterArray
            Data.LowerBound = 1

            If breturnAllData Then
                Data.FromExcelRange .Cells(2, 1), DetectLastColumn:=True, DetectLastRow:=True
            ElseIf returnIndex > 0 Then
                Data.FromExcelRange .Cells(2, returnIndex), DetectLastColumn:=False, DetectLastRow:=True
            End If

            .Cells.Clear
            .Visible = xlSheetVeryHidden
    End With

    lo.AutoFilter.ShowAllData

    Set FilterLoTable = Data.Clone()
End Function

'Get unique values of one range in a listobject
Function GetUniquelo(lo As ListObject, iIndex As Integer) As BetterArray

    Dim Rng As Range
    Dim Data As BetterArray

    Set Rng = lo.ListColumns(iIndex).DataBodyRange

    'Copy and paste to temp
    With ThisWorkbook.Worksheets(C_sSheetTemp)
            .Visible = xlSheetHidden
            .Cells.Clear

            Rng.Copy Destination:=.Cells(1, 1)

            Set Data = New BetterArray
            Data.LowerBound = 1

            .Range(.Cells(1, 1), .Cells(.Cells(.Rows.Count, 1).End(xlUp).Row, .Cells(1, .Columns.Count).End(xlToLeft).Column)).RemoveDuplicates Columns:=1, Header:=xlNo

            Data.FromExcelRange .Cells(1, 1), DetectLastRow:=True, DetectLastColumn:=True
            .Cells.Clear
            .Visible = xlSheetVeryHidden
    End With

    Set GetUniquelo = Data.Clone()

    Set Data = Nothing
    Set Rng = Nothing

End Function

'Unique of a betteray sorted
Function GetUniqueBA(BA As BetterArray) As BetterArray
Dim sval As String
 Dim i As Integer
   Dim Outable As BetterArray

    BA.Sort


    Set Outable = New BetterArray
    Outable.LowerBound = 1

   sval = BA.Item(BA.LowerBound)
   Outable.Push sval

    If BA.Length > 0 Then
        For i = BA.LowerBound To BA.UpperBound
        If sval <> BA.Item(i) Then
            sval = BA.Item(i)
            Outable.Push sval
        End If
        Next
    End If

    Set GetUniqueBA = Outable.Clone()
    Set Outable = Nothing

End Function

Sub StatusBar_Updater(sCpte As Single)
'increase the status progressBar

    Dim CurrentStatus As Integer
    Dim pctDone As Integer

    CurrentStatus = (C_iNumberOfBars) * Round(sCpte / 100, 1)
    Application.StatusBar = "[" & String(CurrentStatus, "|") & Space(C_iNumberOfBars - CurrentStatus) & "]" & " " & CInt(sCpte) & "%" & TranslateMsg("MSG_BuildLL")

    DoEvents

End Sub








