Attribute VB_Name = "Helpers"
Option Private Module

'Basic Helper functions used in the creation of the linelist and other stuffs
'Most of them are explicit functions. Contains all the ancillary sub/
'Functions used when creating the linelist and also in the linelist
'itself

Option Explicit


'FILES, FOLDERS AND OS =========================================================

'Load Folder  and File -----------------------------------------------------

Public Function LoadFolder() As String
    LoadFolder = vbNullString
    If Not Application.OperatingSystem Like "*Mac*" Then
        'We are on windows DOS
        LoadFolder = SelectFolderOnWindows()
    Else
        'We are on Mac, need to test the version of excel running
        If val(Application.Version) > 14 Then
            LoadFolder = SelectFolderOnMac()
        End If
    End If

End Function

Public Function LoadFile(sFilters As String) As String
    LoadFile = vbNullString
    If Not Application.OperatingSystem Like "*Mac*" Then
        'We are on windows DOS
        LoadFile = SelectFileOnWindows(sFilters)
    Else
        'We are on Mac, need to test the version of excel running
        If val(Application.Version) > 14 Then
            LoadFile = SelectFileOnMac(sFilters)
        End If
    End If
End Function

'Just check if it is Mac
Public Function isMac()
    isMac = Application.OperatingSystem Like "*Mac*"
End Function

'Folder selection depending on the OS   ----------------------------------------------------------------------------

'Folder on Mac
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

'Folder on Windows
Private Function SelectFolderOnWindows() As String

    Dim fDialog As Office.FileDialog

    SelectFolderOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .title = "Chose your directory"          'MSG_ChooseDir
        .filters.Clear

        If .Show = -1 Then
            SelectFolderOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'File selection depending on the OS --------------------------------------------

'File on Mac
Private Function SelectFileOnMac(sFilter)

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
        sMacFilter = " {""org.openxmlformats.spreadsheetml.sheet"",""com.microsoft.Excel.sheet.binary.macroenabled""} "
    Case Else
        sMacFilter = " {""org.openxmlformats.spreadsheetml.sheet""} "
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

'File on Windows
Private Function SelectFileOnWindows(sFilters)

    Dim fDialog As Office.FileDialog

    SelectFileOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .title = "Chose your file"               'MSG_ChooseFile
        .filters.Clear
        .filters.Add "Feuille de calcul Excel", sFilters '"*.xlsx" ', *.xlsm, *.xlsb,  *.xls" 'MSG_ExcelFile

        If .Show = True Then
            SelectFileOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'File extension ----------------------------------------------------------------

Public Function GetFileExtension(sString As String) As String

    GetFileExtension = ""

    Dim iDotPos As Integer
    Dim sExt As String                           'extension
    'Find the position of the dot at the end
    iDotPos = InStrRev(sString, ".")

    sExt = Right(sString, Len(sString) - iDotPos)

    If (sExt <> "") Then
        GetFileExtension = sExt
    End If

End Function

'Check if a workbook is opened
Public Function IsWkbOpened(sName As String) As Boolean
    Dim oWkb As Workbook                         'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set oWkb = Application.Workbooks.Item(sName)
    IsWkbOpened = (Not oWkb Is Nothing)
    On Error GoTo 0
End Function

'Check if a Sheet Exists
Public Function SheetExistsInWkb(Wkb As Workbook, sSheetName As String) As Boolean
    SheetExistsInWkb = False
    Dim Wksh As Worksheet                        'Just try to set the workbook if it fails it is closed
    On Error Resume Next
    Set Wksh = Wkb.Worksheets(sSheetName)
    SheetExistsInWkb = (Not Wksh Is Nothing)
    On Error GoTo 0
End Function

'APPLICATION SPEEDUP, WORKSHEET PROTECTION AND RANGE MANAGEMENT =======================================================================================================================================

'Speed up before a work
Public Sub BeginWork(xlsapp As Excel.Application, Optional bstatusbar As Boolean = True)
    xlsapp.ScreenUpdating = False
    xlsapp.DisplayAlerts = False
    xlsapp.Calculation = xlCalculationManual
    xlsapp.EnableAnimations = False
End Sub

'Return previous state
Public Sub EndWork(xlsapp As Excel.Application, Optional bstatusbar As Boolean = True)
    xlsapp.ScreenUpdating = True
    xlsapp.DisplayAlerts = True
    xlsapp.EnableAnimations = True
    xlsapp.DisplayStatusBar = bstatusbar
End Sub

'Remove Gridlines in a worksheet
Public Sub RemoveGridLines(Wksh As Worksheet, Optional DisplayZeros As Boolean = False)
    Dim View As WorksheetView

    For Each View In Wksh.Parent.Windows(1).SheetViews
        If View.Sheet.Name = Wksh.Name Then
            View.DisplayGridlines = False
            View.DisplayZeros = DisplayZeros
            Exit Sub
        End If
    Next
End Sub


'Find The last non empty row of a sheet of type linelist
Function FindLastRow(shLL As Worksheet) As Long

    Dim counter As Long
    Dim LastRow As Long
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim hRng As Range
    Dim controlValue As String
    Dim destRng As Range
    Dim shTemp As Worksheet
    Dim Col As Long 'Column to check the number of rows on

    Set Lo = shLL.ListObjects(1)
    Set hRng = Lo.HeaderRowRange
    Set shTemp = ThisWorkbook.Worksheets("temp__") 'temporary sheet for work

    'First copy the listObject data to the temporary sheet
    'It is really not recommanded to unlist the listobject to get the used range of
    'the worksheet (a lot of formulas rely on this listobject, unlisting will completly break all the links)
    'so we need another approach to find the last row

    '- Copy the range to a temporary sheet
    '- Count the number of rows by removing the formula columns

    shTemp.Cells.Clear
    LastRow = hRng.Row

    Set LoRng = Lo.Range
    Set destRng = shTemp.Range(LoRng.Address)

    'copy the value to the destination range in the temporary worksheet
    destRng.Value = LoRng.Value

    'No need to compute the lastrow if the databodyrange does not exists,
    'in that case the last row is just the headerRow + 1

    If Not Lo.DataBodyRange Is Nothing Then

        For counter = 1 To hRng.Cells.Count

            controlValue = hRng.Cells(1, counter).Offset(-4).Value

            If controlValue <> "formula" And controlValue <> "case_when" And controlValue <> "choice_formula" Then 'case_when is a formula, we should remove them from export

                Col = hRng.Cells(1, counter).Column

                'The test is done only on columns that are not formulas
                If LastRow < shTemp.Cells(Rows.Count, Col).End(xlUp).Row Then LastRow = shTemp.Cells(Rows.Count, Col).End(xlUp).Row

            End If
        Next
    End If

    FindLastRow = LastRow + 1
    shTemp.Cells.Clear
End Function

'STRING AND DATA MANIPULATION =========================================================================================================================================================================
'Safely delete databodyrange of a listobject
Public Sub DeleteLoDataBodyRange(Lo As ListObject)
    If Not Lo.DataBodyRange Is Nothing Then Lo.DataBodyRange.Delete
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

'Clear Unicode Characters and non printable characters in a String

Public Function ClearNonPrintableUnicode(ByVal sString As String) As String
    Dim sValue As String
    sValue = Application.WorksheetFunction.SUBSTITUTE(sString, chr(160), " ")
    sValue = Application.WorksheetFunction.Clean(sValue)
    ClearNonPrintableUnicode = Application.WorksheetFunction.Trim(sValue)
End Function

'Get the headers of one sheet from one line (probablly the first line)
'The headers are cleaned

Public Function GetHeaders(Wkb As Workbook, sSheet As String, startLine As Long, Optional StartColumn As Long = 1) As BetterArray
    'Extract column names in one sheet starting from one line
    Dim Headers As BetterArray
    Dim i As Long
    Set Headers = New BetterArray
    Headers.LowerBound = StartColumn
    Dim sValue As String

    With Wkb.Worksheets(sSheet)
        i = StartColumn
        Do While .Cells(startLine, i).Value <> vbNullString
            'Clear the values in the sheet when adding thems
            sValue = .Cells(startLine, i).Value  'The argument is passed byval to clearstring
            sValue = ClearString(sValue)
            Headers.Push sValue
            i = i + 1
        Loop
        Set GetHeaders = Headers.Clone
        'Clear everything
    End With
End Function

'Get the data from one sheet starting from one line
Public Function GetData(Wkb As Workbook, sSheetName As String, startLine As Long, Optional EndColumn As Long = 0) As BetterArray
    Dim Data As BetterArray
    Dim rng As Range

    Dim iLastRow As Long
    Dim iLastCol As Long
    Set Data = New BetterArray
    Data.LowerBound = 1

    With Wkb.Worksheets(sSheetName)

        iLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        iLastCol = EndColumn
        If EndColumn = 0 Then iLastCol = .Cells(startLine, .Columns.Count).End(xlToLeft).Column
        Set rng = .Range(.Cells(startLine, 1), .Cells(iLastRow, iLastCol))
    End With

    Data.FromExcelRange rng
    'The output of the function is a variant
    Set GetData = Data

End Function

'Get the validation list using Choices data and choices labels
'Get the list of validations from the Choices data
Public Function GetValidationList(ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, sValidation As String) As BetterArray

    Dim iChoiceIndex As Integer
    Dim iChoiceLastIndex As Integer
    Dim ValidationList As BetterArray            'Validation List

    Set ValidationList = New BetterArray
    ValidationList.LowerBound = 1

    iChoiceIndex = ChoicesListData.IndexOf(sValidation)
    iChoiceLastIndex = ChoicesListData.LastIndexOf(sValidation)

    If (iChoiceIndex > 0) Then
        ValidationList.Items = ChoicesLabelsData.Slice(iChoiceIndex, iChoiceLastIndex + 1)
    End If
    Set GetValidationList = ValidationList.Clone()
End Function

'Test if a listobject exists
Public Function ListObjectExists(Wksh As Worksheet, sListObjectName As String) As Boolean
    ListObjectExists = False
    Dim Lo As ListObject
    On Error Resume Next
    Set Lo = Wksh.ListObjects(sListObjectName)
    ListObjectExists = (Not Lo Is Nothing)
    On Error GoTo 0
End Function

'Get validation type
Function GetValidationType(sValidationType As String) As Byte

    GetValidationType = 3                        'list of validation info, warning or error
    If sValidationType <> "" Then
        Select Case LCase(sValidationType)
        Case "warning"
            GetValidationType = 2
        Case "error"
            GetValidationType = 1
        End Select
    End If

End Function

'Filter a table listobject on one condition and get the values of that table or all the unique values of one column
Public Function FilterLoTable(Lo As ListObject, iFiltindex1 As Integer, sValue1 As String, _
                              Optional iFiltindex2 As Integer = 0, Optional sValue2 As String = vbNullString, _
                              Optional iFiltindex3 As Integer = 0, Optional sValue3 As String = vbNullString, _
                              Optional returnIndex As Integer = -99, _
                              Optional bAllData As Boolean = True) As BetterArray
    Dim rng As Range
    Dim Data As BetterArray
    Dim breturnAllData As Boolean

    With Lo.Range

        .AutoFilter Field:=iFiltindex1, Criteria1:=sValue1

        'Add other Filters if required
        If iFiltindex2 > 0 Then
            .AutoFilter Field:=iFiltindex2, Criteria1:=sValue2
        End If

        If iFiltindex3 > 0 Then
            .AutoFilter Field:=iFiltindex3, Criteria1:=sValue3
        End If

    End With

    Set rng = Lo.Range.SpecialCells(xlCellTypeVisible)

    If returnIndex > 0 Then
        breturnAllData = False
    ElseIf bAllData Then
        breturnAllData = True
    Else
        breturnAllData = True
    End If

    'Copy and paste to temp
    With ThisWorkbook.Worksheets(C_sSheetTemp)
        .Cells.Clear

        rng.Copy Destination:=.Cells(1, 1)

        Set Data = New BetterArray
        Data.LowerBound = 1

        If breturnAllData Then
            Data.FromExcelRange .Cells(2, 1), DetectLastColumn:=True, DetectLastRow:=True
        ElseIf returnIndex > 0 Then
            Data.FromExcelRange .Cells(2, returnIndex), DetectLastColumn:=False, DetectLastRow:=True
        End If

        .Cells.Clear
    End With

    Lo.AutoFilter.ShowAllData

    Set FilterLoTable = Data.Clone()
End Function


'Unique of a betteray sorted
Function GetUniqueBA(BA As BetterArray) As BetterArray
    Dim sVal As String
    Dim i As Long
    Dim Outable As BetterArray

    BA.Sort

    Set Outable = New BetterArray
    Outable.LowerBound = 1

    sVal = Application.WorksheetFunction.Trim(BA.Item(BA.LowerBound))

    If sVal <> vbNullString Then
        Outable.Push sVal
    End If

    If BA.Length > 0 Then
        For i = BA.LowerBound To BA.UpperBound
            If sVal <> Application.WorksheetFunction.Trim(BA.Item(i)) And Application.WorksheetFunction.Trim(BA.Item(i)) <> vbNullString Then
                sVal = Application.WorksheetFunction.Trim(BA.Item(i))
                Outable.Push sVal
            End If
        Next
    End If

    Set GetUniqueBA = Outable.Clone()

End Function

'Check if a worksheet name is correct
Public Function SheetNameIsBad(sSheetName As String) As Boolean

    SheetNameIsBad = (sSheetName = C_sSheetGeo Or sSheetName = C_sSheetFormulas Or _
                      sSheetName = C_sSheetPassword Or sSheetName = C_sSheetTemp Or _
                      sSheetName = C_sSheetLLTranslation Or sSheetName = C_sSheetChoiceAuto Or _
                      sSheetName = C_sParamSheetDict Or sSheetName = C_sParamSheetExport Or _
                      sSheetName = C_sParamSheetChoices Or sSheetName = C_sParamSheetTranslation Or _
                      sSheetName = C_sSheetMetadata Or sSheetName = C_sSheetAnalysisTemp Or _
                      sSheetName = C_sSheetImportTemp Or sSheetName = sParamSheetAnalysis Or _
                      sSheetName = sParamSheetTemporalAnalysis Or sSheetName = sParamSheetSpatialAnalysis Or _
                      sSheetName = sParamSheetAdmin)

End Function


Public Function SheetListObjectName(sSheetName As String) As String
    SheetListObjectName = vbNullString
    On Error Resume Next
    SheetListObjectName = ThisWorkbook.Worksheets(sSheetName).ListObjects(1).Name
    On Error GoTo 0
End Function

'FORMULAS AND VALIDATIONS ==============================================================================================================================================================================


