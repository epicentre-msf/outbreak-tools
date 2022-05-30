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
    xlsapp.DisplayStatusBar = bstatusbar
End Sub

Public Sub RemoveGridLines(Wksh As Worksheet)
    Dim View As WorksheetView

    For Each View In Wksh.Parent.Windows(1).SheetViews
        If View.Sheet.Name = Wksh.Name Then
            View.DisplayGridlines = False
            View.DisplayZeros = False
            Exit Sub
        End If
    Next
End Sub


Public Sub EndWork(xlsapp As Excel.Application, Optional bstatusbar As Boolean = True)
    xlsapp.ScreenUpdating = True
    xlsapp.DisplayAlerts = True
    xlsapp.Calculation = xlCalculationAutomatic
    xlsapp.DisplayStatusBar = bstatusbar
End Sub


'----- File and folder selections depending on the OS

Public Function LoadFolder() As String
    LoadFolder = vbNullString
    If Not Application.OperatingSystem Like "*Mac*" Then
        'We are on windows DOS
        LoadFolder = SelectFolderOnWindows()
    Else
        'We are on Mac, need to test the version of excel running
        If Val(Application.Version) > 14 Then
            LoadFolder = SelectFolderOnMac()
        End If
    End If

End Function

'Load files and folders
Public Function LoadFile(sFilters As String) As String
    LoadFile = vbNullString
    If Not Application.OperatingSystem Like "*Mac*" Then
        'We are on windows DOS
        LoadFile = SelectFileOnWindows(sFilters)
    Else
        'We are on Mac, need to test the version of excel running
        If Val(Application.Version) > 14 Then
            LoadFile = SelectFileOnMac(sFilters)
        End If
    End If
End Function

'The selection process depends on the operating system. Here is a simple
'code for Mac, using applescript:

'----------------------------- FOLDER SELECTION --------------------------------
Private Function SelectFolderOnMac() As String
    Dim FolderPath As String
    Dim RootFolder As String
    Dim Scriptstr As String

    On Error Resume Next

    'Enter the Start Folder, Desktop in this example,
    'Use the second line to enter your own path
    RootFolder = MacScript("return POSIX path of (path to documents folder) as string")

    'Make the path Colon seperated for using in MacScript
    RootFolder = MacScript("return POSIX file (""" & RootFolder & """) as string")
    'Make the Script string
        Scriptstr = "return POSIX path of (choose folder with prompt ""Select the folder""" & _
            " default location alias """ & RootFolder & """) as string"

    'Run the Script
    FolderPath = MacScript(Scriptstr)

    If CInt(Split(Application.Version, ".")(0)) >= 15 Then 'excel 2016 support
        FolderPath = Replace(FolderPath, ":", "/")
        FolderPath = Replace(FolderPath, "Macintosh HD", "", Count:=1)
    End If

    On Error GoTo 0

    If FolderPath <> "" Then
        'Remove the last ":" or "/"
        SelectFolderOnMac = Mid(FolderPath, 1, (Len(FolderPath) - 1))
    End If
End Function

'Now Select a folder on Windows

Private Function SelectFolderOnWindows() As String

    Dim fDialog As Office.FileDialog

    SelectFolderOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your directory"          'MSG_ChooseDir
        .Filters.Clear

        If .show = True Then
            SelectFolderOnWindows = .SelectedItems(1)
        End If
    End With
    Set fDialog = Nothing

End Function


'------------------------------ FILE SELECTION ---------------------------------

Function SelectFileOnMac(sFilter)

    Dim sMacFilter As String
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String


    Select Case sFilter
        Case "*.xls"
            sMacFilter = " {""com.microsoft.Excel.xls""} "
        Case "*.xlsx"
            sMacFilter = " {""org.openxmlformats.spreadsheetml.sheet""} "
        Case "*.xlsb"
            sMacFilter = " {""com.microsoft.Excel.sheet.binary.macroenabled""} "
        Case "*.xlsb, *.xlsx"
            sMacFilter = " {""com.microsoft.excel.xls"",""public.comma-separated-values-text""} "
        Case Else
            sMacFilter = " {""com.microsoft.Excel.xls""} "
    End Select

    SelectFileOnMac = vbNullString
    On Error Resume Next
       MyPath = MacScript("return (path to documents folder) as String")
       MyScript = _
       "set applescript's text item delimiters to "","" " & vbNewLine & _
        "set theFiles to (choose file of type " & _
        sMacFilter & _
        "with prompt ""Please select a file or files"" default location alias """ & _
        MyPath & """ multiple selections allowed false) as string" & vbNewLine & _
        "set applescript's text item delimiters to """" " & vbNewLine & _
        "return theFiles"
    MyFiles = MacScript(MyScript)

      If CInt(Split(Application.Version, ".")(0)) >= 15 Then 'excel 2016 support
        MyFiles = Replace(MyFiles, ":", "/")
        MyFiles = Replace(MyFiles, "Macintosh HD", "", Count:=1)
    End If

    On Error GoTo 0

   SelectFileOnMac = MyFiles

End Function

Function SelectFileOnWindows(sFilters)

    Dim fDialog As Office.FileDialog

    SelectFileOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your file"               'MSG_ChooseFile
        .Filters.Clear
        .Filters.Add "Feuille de calcul Excel", sFilters '"*.xlsx" ', *.xlsm, *.xlsb,  *.xls" 'MSG_ExcelFile

        If .show = True Then
            SelectFileOnWindows = .SelectedItems(1)
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
            sValidationList = sValidationList & "," & ChoicesLabelsData.Items(i)
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


'Move a plage of data from the setup sheet to the designer sheet
Public Sub MoveData(SourceWkb As Workbook, DestWkb As Workbook, sSheetName As String, sStartCell As Integer)

    Dim sData As BetterArray
    Dim DestWksh As Worksheet
    Dim sheetExists As Boolean

    Set sData = New BetterArray
    sData.FromExcelRange SourceWkb.Worksheets(sSheetName).Range("A" & CStr(sStartCell)), DetectLastRow:=True, DetectLastColumn:=True
    sheetExists = False

    For Each DestWksh In DestWkb.Worksheets
        If DestWksh.Name = sSheetName Then
            sheetExists = True
            Exit For
        End If
    Next

    'Clear the contents if the sheet exists, or create a new sheet if Not
    If sheetExists Then
        DestWkb.Worksheets(sSheetName).Cells.Clear
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
    Dim bCurrEvent As Boolean

    bCurrEvent = Application.ScreenUpdating

    Application.ScreenUpdating = True

    CurrentStatus = (C_iNumberOfBars) * Round(sCpte / 100, 1)
    SheetMain.Range(C_sRngUpdate).value = "[" & String(CurrentStatus, "|") & Space(C_iNumberOfBars - CurrentStatus) & "]" & " " & CInt(sCpte) & "% " & TranslateMsg("MSG_BuildLL")

    Application.ScreenUpdating = bCurrEvent

End Sub


'Transform one formula to a formula for analysis.
'Wkb is a workbook where we can find the dictionary, the special character
'data and the name of all 'friendly' functions

Public Function AnalysisFormula(sFormula As String, Wkb As Workbook) As String
    'Returns a string of cleared formula

    AnalysisFormula = ""

    Dim sFormulaATest As String                  'same formula, with all the spaces replaced with
    Dim sAlphaValue As String                    'Alpha numeric values in a formula
    Dim sLetter As String                        'counter for every letter in one formula
    Dim scolAddress As String                    'address of one column used in a formula

    Dim FormulaAlphaData As BetterArray          'Table of alphanumeric data in one formula
    Dim FormulaData      As BetterArray
    Dim VarNameData  As BetterArray              'List of all variable names
    Dim SpecCharData As BetterArray              'List of Special Characters data
    Dim DictHeaders As BetterArray
    Dim SheetNameData As BetterArray
    Dim VarMainLabelData As BetterArray


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
    Set FormulaData = New BetterArray
    Set VarNameData = New BetterArray       'The list of all Variable Names
    Set SpecCharData = New BetterArray       'The list of all special characters
    Set DictHeaders = New BetterArray
    Set VarMainLabelData = New BetterArray
    Set SheetNameData = New BetterArray

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
    If DictHeaders.IndexOf(C_sDictHeaderSheetName) < 0 Or DictHeaders.IndexOf(C_sDictHeaderMainLab) < 0 Then
        Exit Function
    End If


    SheetNameData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, DictHeaders.IndexOf(C_sDictHeaderSheetName)), DetectLastColumn:=False, DetectLastRow:=True
    VarMainLabelData.FromExcelRange Wkb.Worksheets(C_sParamSheetDict).Cells(1, DictHeaders.IndexOf(C_sDictHeaderMainLab)), DetectLastColumn:=False, DetectLastRow:=True


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
                            sAlphaValue = "o" & ClearString(SheetNameData.Item(icolNumb)) & "['" & VarMainLabelData.Item(icolNumb) & "']"
                        ElseIf FormulaData.Includes(UCase(sAlphaValue)) Then 'It is a formula, excel will do the translation for us
                                sAlphaValue = Application.WorksheetFunction.Trim(sAlphaValue)
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
        AnalysisFormula = sAlphaValue
    Else
     MsgBox "Error in analysis formula: " & sFormula
    End If

    Set FormulaAlphaData = Nothing  'Alphanumeric values of one formula
    Set VarNameData = Nothing       'The list of all Variable Names
    Set SpecCharData = Nothing      'The list of all special characters
    Set DictHeaders = Nothing
    Set VarMainLabelData = Nothing
    Set SheetNameData = New BetterArray

End Function




Public Function GetInternationalFormula(sFormula As String, Wksh As Worksheet) As String

    Dim sprevformula As String
    Dim slocalformula As String

    GetInternationalFormula = ""

    'The formula is in English, I need to take the international
    'value of the formula, and avoid using the table of formulas only when I deal with Validations

    If (sFormula <> "") Then
        sprevformula = Wksh.Range("A1").Formula
        'Setting the formula to a range
        Wksh.Range("A1").Formula = "=" & sFormula
        'retrieving the local formula
        GetInternationalFormula = Wksh.Range("A1").FormulaLocal
    End If
        'Reseting the previous formula
    Wksh.Range("A1").Formula = sprevformula
    Set Wksh = Nothing
End Function


Public Function DATE_RANGE(DateRng As Range) As String
    DATE_RANGE = Format(Application.WorksheetFunction.Min(DateRng), "DD/MM/YYYY") & _
     " - " & Format(Application.WorksheetFunction.Max(DateRng), "DD/MM/YYYY")
End Function




