Attribute VB_Name = "DesignerBuildList"
Option Explicit
Option Private Module

'BUILD THE LINELIST ===================================================================================================================================================================================

'Building the linelist from the different input data
'@DictHeaders: The headers of the dictionnary sheet
'@DictData: Dictionnary data
'@ChoicesHeaders: The headers of the Choices sheet
'@ChoicesData: The choices data
'@ExportData: The export data

Private AddedLogo As Boolean                     'Added Logo?

Sub BuildList(DictHeaders As BetterArray, dictData As BetterArray, ExportData As BetterArray, _
              ChoicesHeaders As BetterArray, ChoicesData As BetterArray, _
              TransData As BetterArray, GSData As BetterArray, UAData As BetterArray, BAData As BetterArray, _
              TAData As BetterArray, SAData As BetterArray, sPath As String)


    Dim wkb As Workbook
    Dim LLNbColData             As BetterArray   'Number of columns of a Sheet of type linelist
    Dim LLSheetNameData         As BetterArray   'Names of sheets
    Dim ChoicesListData         As BetterArray   'Choices list
    Dim ChoicesLabelsData       As BetterArray   ' Choices labels
    Dim VarNameData             As BetterArray
    Dim ColumnIndexData         As BetterArray
    Dim SheetsOfTypeLLData      As BetterArray
    Dim ChoiceAutoVarData       As BetterArray
    Dim FormulaData             As BetterArray
    Dim SpecCharData            As BetterArray
    Dim DictVarName             As BetterArray
    Dim DictSheetNames          As BetterArray
    Dim TableNameData           As BetterArray
    Dim LinelistDictionary      As ILLdictionary
    Dim iPastingRow             As Integer
    Dim iNbshifted              As Integer
    'For updating sheet names in the dictionary worksheet
    Dim i                       As Integer       'iterator
    Dim iSheetNameColumn        As Integer
    Dim sFirstSheetName         As String        'Previous sheet names where to copy data to:
    Dim iWindowState            As Integer
    Dim Wksh                    As Worksheet
    Dim iPerc                   As Integer


    Dim iCounterSheet As Integer                 'counter for one Sheet
    Dim iSheetStartLine As Integer               'Counter for starting line of the sheet in the dictionary

    'Instanciating the betterArrays
    Set LLNbColData = New BetterArray
    Set LLSheetNameData = New BetterArray        'Names of sheets of type linelist
    Set ColumnIndexData = New BetterArray
    Set FormulaData = New BetterArray
    Set SpecCharData = New BetterArray
    Set VarNameData = New BetterArray
    Set DictVarName = New BetterArray
    Set SheetsOfTypeLLData = New BetterArray
    Set ChoiceAutoVarData = New BetterArray
    Set DictSheetNames = New BetterArray
    Set TableNameData = New BetterArray


    Set LinelistDictionary = LLdictionary.Create(ThisWorkbook.Worksheets("Dictionary"), 1, 1)
    Set ChoiceAutoVarData = LinelistDictionary.Data.FilterData("control", "list_auto", "control details")

    AddedLogo = False

    BeginWork xlsapp:=Application
    'iWindowState = Application.WindowState

    Application.EnableAnimations = False
    Application.EnableEvents = False
    Application.Cursor = xlDefault

    Set wkb = Workbooks.Add
    iWindowState = Application.WindowState
    Application.WindowState = xlMinimized
    BeginWork xlsapp:=Application

    iUpdateCpt = iUpdateCpt + 5
    StatusBar_Updater (iUpdateCpt)

    'Now Transferring some designers objects (codes, modules) to the workbook we want to create
    Call DesignerBuildListHelpers.TransferDesignerCodes(wkb)

    DoEvents

    'DesignerBuildListHelpers.TransterSheet is for sending worksheets from the actual workbook to the first workbook of the instance
    sFirstSheetName = wkb.Worksheets(1).Name
    TransferSheet wkb, C_sSheetGeo, sFirstSheetName
    TransferSheet wkb, C_sSheetPassword, C_sSheetGeo
    TransferSheet wkb, C_sSheetFormulas, C_sSheetPassword
    TransferSheet wkb, C_sSheetLLTranslation, C_sSheetFormulas

    DoEvents
    iUpdateCpt = iUpdateCpt + 5
    StatusBar_Updater (iUpdateCpt)

    DoEvents

    'Create special characters data
    FormulaData.FromExcelRange SheetFormulas.ListObjects(C_sTabExcelFunctions).ListColumns("ENG").DataBodyRange, DetectLastColumn:=False
    SpecCharData.FromExcelRange SheetFormulas.ListObjects(C_sTabASCII).ListColumns("TEXT").DataBodyRange, DetectLastColumn:=False

    VarNameData.LowerBound = 1
    DictSheetNames.LowerBound = 1
    TableNameData.LowerBound = 1

    VarNameData.Items = dictData.ExtractSegment(ColumnIndex:=DictHeaders.IndexOf(C_sDictHeaderVarName))
    DictSheetNames.Items = dictData.ExtractSegment(ColumnIndex:=DictHeaders.IndexOf(C_sDictHeaderSheetName))
    TableNameData.Items = dictData.ExtractSegment(ColumnIndex:=DictHeaders.IndexOf(C_sDictHeaderTableName))

    'Create all the required Sheets in the workbook (Dictionnary, Export, Password, Geo and other sheets defined by the user)
    Call CreateSheets(wkb, dictData, DictHeaders, ExportData, _
                      ChoicesHeaders, ChoicesData, TransData, _
                      LLNbColData, ColumnIndexData, LLSheetNameData, _
                      bNotHideSheets:=False)
    DoEvents

    'Choices data'Setting the Choices labels and lists
    Set ChoicesListData = New BetterArray
    Set ChoicesLabelsData = New BetterArray

    ChoicesListData.LowerBound = 1
    ChoicesLabelsData.LowerBound = 1
    ChoiceAutoVarData.LowerBound = 1

    'Update the values of the labels and list! here I must make sure my Headers contains those values

    If (ChoicesHeaders.IndexOf(C_sChoiHeaderList) <= 0 Or ChoicesHeaders.IndexOf(C_sChoiHeaderLab) <= 0) Then
        SheetMain.Range("RNG_Edition").Value = "Error 1"
        Exit Sub
    End If

    ChoicesListData.Items = ChoicesData.ExtractSegment(ColumnIndex:=ChoicesHeaders.IndexOf(C_sChoiHeaderList))
    ChoicesLabelsData.Items = ChoicesData.ExtractSegment(ColumnIndex:=ChoicesHeaders.IndexOf(C_sChoiHeaderLab))

    iSheetStartLine = 1
    iNbshifted = 0


    Windows(wkb.Name).Visible = False
    Application.WindowState = iWindowState

    iPerc = 80 - iUpdateCpt

    iPerc = Round(iPerc / LLSheetNameData.Length, 1)


    For iCounterSheet = 1 To LLSheetNameData.UpperBound


        Select Case dictData.Items(iSheetStartLine, DictHeaders.IndexOf(C_sDictHeaderSheetType))
            'On linelist type, build a data entry form
        Case C_sDictSheetTypeLL
            'Create a sheet for data Entry in one sheet of type linelist
            Call CreateSheetLLDataEntry(wkb, LLSheetNameData.Item(iCounterSheet), iSheetStartLine, dictData, _
                                        DictHeaders, LLSheetNameData, LLNbColData, ChoicesListData, ChoicesLabelsData, _
                                        VarNameData, ColumnIndexData, FormulaData, SpecCharData, ChoiceAutoVarData, _
                                        DictSheetNames, iNbshifted)
            DoEvents


            'update the variable names for writing in the dictionary sheet
            i = 1
            With wkb.Worksheets(LLSheetNameData.Item(iCounterSheet))
                Do While (.Cells(C_eStartLinesLLData, i).Value <> "")
                    DictVarName.Push .Cells(C_eStartLinesLLData + 1, i).Value
                    i = i + 1
                Loop
            End With

            'Now writing the data of varnames to the dictionary
            With wkb.Worksheets(C_sParamSheetDict)
                iPastingRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                DictVarName.ToExcelRange Destination:=.Cells(iPastingRow + 1, 1)
                DictVarName.Clear
            End With

            SheetsOfTypeLLData.Push LLSheetNameData.Item(iCounterSheet)

        Case C_sDictSheetTypeAdm

            'Create a sheet of type admin entry
            Call CreateSheetAdmEntry(wkb, LLSheetNameData.Item(iCounterSheet), iSheetStartLine, dictData, _
                                     DictHeaders, LLSheetNameData, LLNbColData, _
                                     ChoicesListData, ChoicesLabelsData)
            i = 0
            With wkb.Worksheets(LLSheetNameData.Item(iCounterSheet))
                Do While (.Cells(C_eStartLinesAdmData + i, C_eStartColumnAdmData + 2).Value <> "")
                    DictVarName.Push .Cells(C_eStartLinesAdmData + i, C_eStartColumnAdmData + 3).Name.Name
                    i = i + 1
                Loop
            End With

            'Now writing the data of varnames to the dictionary
            With wkb.Worksheets(C_sParamSheetDict)
                iPastingRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                DictVarName.ToExcelRange Destination:=.Cells(iPastingRow + 1, 1)
                DictVarName.Clear
            End With
            DoEvents
        End Select

        iSheetStartLine = iSheetStartLine + LLNbColData.Item(iCounterSheet)

        DoEvents
        iUpdateCpt = iUpdateCpt + iPerc
        StatusBar_Updater (iUpdateCpt)
    Next

    'Put the dictionnary in a table format
    With wkb.Worksheets(C_sParamSheetDict)
        .Cells(1, 1).Value = C_sDictHeaderVarName
        'Update values of the Sheet Names with correct spelling
        For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
            iSheetNameColumn = DictHeaders.IndexOf(C_sDictHeaderSheetName)
            .Cells(i, iSheetNameColumn).Value = EnsureGoodSheetName(.Cells(i, iSheetNameColumn).Value)
        Next

        .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(dictData.Length, DictHeaders.Length + 1)), , xlYes).Name = "o" & ClearString(C_sParamSheetDict)
        .ListObjects("o" & ClearString(C_sParamSheetDict)).Resize .ListObjects("o" & ClearString(C_sParamSheetDict)).Range.CurrentRegion
    End With

    SheetMain.Range("RNG_Edition").Value = TranslateMsg("MSG_BuildAna")

    Call DesignerBuildListHelpers.UpdateChoiceAutoHeaders(wkb, ChoiceAutoVarData, DictHeaders)

    '======== Build the Analysis ======================================================================================================================================

    Call BuildAnalysis(wkb, GSData, UAData, BAData, TAData, SAData, ChoicesListData, ChoicesLabelsData, dictData, DictHeaders, VarNameData, TableNameData)


    iUpdateCpt = iUpdateCpt + 2
    StatusBar_Updater (iUpdateCpt)

    #If Mac Then
        'Mac users will have to endure screen flickering, no choice
        Windows(wkb.Name).Visible = True
        Windows(wkb.Name).WindowState = xlMaximized

        For Each Wksh In wkb.Worksheets
            'Unable to write this code as a sub, please keep it in mind because it won't freeze.
            If SheetsOfTypeLLData.Includes(Wksh.Name) Then
                Wksh.Activate
                With ActiveWindow
                    .SplitColumn = C_iLLSplitColumn
                    .SplitRow = C_eStartLinesLLData + 1
                    .FreezePanes = True
                End With

            ElseIf Wksh.Name = sParamSheetAnalysis Then
                Wksh.Activate
                With ActiveWindow
                    .SplitRow = C_eStartLinesAnalysis - 3
                    .FreezePanes = True
                End With

            ElseIf Wksh.Name = sParamSheetTemporalAnalysis Then
                Wksh.Activate
                With ActiveWindow
                    .SplitColumn = C_iLLSplitColumn + 2
                    .SplitRow = C_eStartLinesAnalysis - 3
                    .FreezePanes = True
                End With
            End If
        Next

        wkb.SaveAs FileName:=sPath, fileformat:=xlExcel12, Password:=SheetMain.Range("RNG_LLPwdOpen").Value, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        wkb.Close
    #Else
        'I am on windows, I will save the workbook, reopen it with new instance, put everything as visible in the workbook, hide the instance and do my work on Panes
        wkb.SaveAs SheetMain.Range("RNG_LLDir") & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "Temp", fileformat:=xlExcel12
        wkb.Close
        Dim Myxlsapp As Excel.Application
        Set Myxlsapp = New Excel.Application
        With Myxlsapp
            .Visible = False
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableAnimations = False
            .EnableEvents = False
            Set wkb = .Workbooks.Open(SheetMain.Range("RNG_LLDir") & Application.PathSeparator & "LinelistApp_" & Application.PathSeparator & "Temp.xlsb")
            .Windows(wkb.Name).Visible = True
            For Each Wksh In wkb.Worksheets
                If SheetsOfTypeLLData.Includes(Wksh.Name) Then
                    Wksh.Activate
                    With .ActiveWindow
                        .SplitColumn = C_iLLSplitColumn
                        .SplitRow = C_eStartLinesLLData + 1
                        .FreezePanes = True
                    End With

                ElseIf Wksh.Name = sParamSheetAnalysis Then
                    Wksh.Activate
                    With .ActiveWindow
                        .SplitRow = C_eStartLinesAnalysis - 3
                        .FreezePanes = True
                    End With
                ElseIf Wksh.Name = sParamSheetTemporalAnalysis Then
                    Wksh.Activate
                    With .ActiveWindow
                        .SplitColumn = C_iLLSplitColumn + 2
                        .SplitRow = C_eStartLinesAnalysis - 3
                        .FreezePanes = True
                    End With
                End If
            Next
        End With
        wkb.SaveAs FileName:=sPath, fileformat:=xlExcel12, Password:=SheetMain.Range("RNG_LLPwdOpen").Value, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        wkb.Close

        Myxlsapp.Quit

    #End If

    EndWork xlsapp:=Application

    Application.EnableAnimations = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault

End Sub

'CREATE SHEETS IN A LINELIST ==========================================================================================================================================================================

'Create the required Sheet and Hide some of them

'@Wkb: Workbook
'@DictData: Dictionary Data
'@DictHeaders: Headers of the dictionary
'@ExportData: Export Data
'@LLNbColData: This is a vector that will be updated. It counts for each sheet, the number of columns
'@ColumnIndexData: This a vector that will be updated. It count for each column, the index in the sheet where the column should be
'@LLSheetName: This is a vector that will contain le name of all the sheets
'@bNotHideSheets: For debugging purpose (hide or not dicitonary and Export sheets)

Private Sub CreateSheets(wkb As Workbook, dictData As BetterArray, DictHeaders As BetterArray, _
                         ExportData As BetterArray, ChoicesHeaders As BetterArray, _
                         ChoicesData As BetterArray, TransData As BetterArray, _
                         LLNbColData As BetterArray, ColumnIndexData As BetterArray, _
                         LLSheetNameData As BetterArray, Optional bNotHideSheets As Boolean = False)
    'LLNbColData: Number of columns for a sheet of type linelist
    'LLSheetNameData: Name of a sheet of type linelist

    Dim i As Integer                             'iterators
    Dim j As Integer
    Dim sNewSheetName As String                  'New sheet name
    Dim sPrevSheetName As String                 'Previous sheet name

    ColumnIndexData.LowerBound = 1


    With wkb
        'Workbook already contains Password and formula sheets. Hide them
        .Worksheets(C_sSheetPassword).Visible = xlVeryHidden
        .Worksheets(C_sSheetFormulas).Visible = xlVeryHidden
        .Worksheets(C_sSheetLLTranslation).Visible = xlVeryHidden

        'Creating the dictionnary sheet from setup
        .Worksheets.Add.Name = C_sParamSheetDict
        'Headers of the disctionary
        DictHeaders.ToExcelRange Destination:=.Worksheets(C_sParamSheetDict).Cells(1, 1), TransposeValues:=True
        'Data of the dictionary
        dictData.ToExcelRange Destination:=.Worksheets(C_sParamSheetDict).Cells(2, 1)
        .Worksheets(C_sParamSheetDict).Columns(1).ClearContents
        'Adding the column index to the Dictionary Sheet
        .Worksheets(C_sParamSheetDict).Cells(1, DictHeaders.Length + 1).Value = C_sDictHeaderIndex
        .Worksheets(C_sParamSheetDict).Visible = bNotHideSheets

        'Creating the Choices Sheet
        .Worksheets.Add.Name = C_sParamSheetChoices
        ChoicesHeaders.ToExcelRange Destination:=.Worksheets(C_sParamSheetChoices).Cells(1, 1), TransposeValues:=True
        ChoicesData.ToExcelRange Destination:=.Worksheets(C_sParamSheetChoices).Cells(2, 1)
        .Worksheets(C_sParamSheetChoices).Visible = bNotHideSheets

        '---------- Creating the export sheet
        .Worksheets.Add.Name = C_sParamSheetExport
        ExportData.ToExcelRange Destination:=.Worksheets(C_sParamSheetExport).Cells(1, 1)
        .Worksheets(C_sParamSheetExport).Visible = xlSheetVeryHidden

        '--------- Creating the translation sheet
        .Worksheets.Add.Name = C_sParamSheetTranslation
        TransData.ToExcelRange Destination:=.sheets(C_sParamSheetTranslation).Cells(1, 1)
        .Worksheets(C_sParamSheetTranslation).Visible = xlSheetVeryHidden

        'Add the metadata sheet
        Call DesignerBuildListHelpers.AddMetadataSheet(wkb)

        'Add the temporary sheets for computation and stuffs
        Call DesignerBuildListHelpers.AddTemporarySheets(wkb)

        'Add a Sheet called Admin for buttons and managements
        Call DesignerBuildListHelpers.AddAdminSheet(wkb)

        'Add Analysis sheets
        .Worksheets.Add(after:=.Worksheets(sParamSheetAdmin)).Name = sParamSheetAnalysis
        Call RemoveGridLines(.Worksheets(sParamSheetAnalysis), DisplayZeros:=True)

        'Temporal analysis Sheet
        .Worksheets.Add(after:=.Worksheets(sParamSheetAnalysis)).Name = sParamSheetTemporalAnalysis
        Call RemoveGridLines(.Worksheets(sParamSheetTemporalAnalysis), DisplayZeros:=True)

        'Spatial analysis sheet
        .Worksheets.Add(after:=.Worksheets(sParamSheetTemporalAnalysis)).Name = sParamSheetSpatialAnalysis
        Call RemoveGridLines(.Worksheets(sParamSheetSpatialAnalysis), DisplayZeros:=True)

        '--------------- adding the other the other sheets in the dictionary to the linelist
        i = 1
        sPrevSheetName = sParamSheetAdmin
        j = 0
        'Setting the lower bound before entering the loop
        LLNbColData.LowerBound = 1
        LLSheetNameData.LowerBound = 1
        'i will hep move from one values of dictionnary data to another
        Do While i <= dictData.UpperBound
            sNewSheetName = EnsureGoodSheetName(dictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName)))

            If sPrevSheetName <> sNewSheetName Then
                .Worksheets.Add(after:=.Worksheets(sPrevSheetName)).Name = sNewSheetName

                'Add Filtered Data sheet for filtered data
                .Worksheets.Add(after:=.Worksheets(sNewSheetName)).Name = C_sFiltered & sNewSheetName
                .Worksheets(C_sFiltered & sNewSheetName).Visible = xlSheetVeryHidden

                'Remove the gridlines in this new Sheetname
                Call RemoveGridLines(.Worksheets(sNewSheetName))
                'I am on a new sheet name, I update values
                sPrevSheetName = sNewSheetName

                j = j + 1
                'Here, the column index is the index number of each column in one sheet. I update it when I am on
                'a new sheet
                ColumnIndexData.Item(i) = 1
                LLNbColData.Item(j) = 1
                LLSheetNameData.Push sPrevSheetName

                'Tell the use we have created one sheet
                'adding sheets depending on the type of the sheet
                Select Case dictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetType))
                Case C_sDictSheetTypeAdm
                    'This is a admin Sheet, just add it like that (or maybe do some other stuffs later on)

                Case C_sDictSheetTypeLL
                    'Set the rowheight of the first two rows of a linelist type sheet
                    .Worksheets(sPrevSheetName).Rows("1:4").RowHeight = C_iLLButtonsRowHeight
                    'Now I split at starting lines and freeze the pane
                Case Else
                    SheetMain.Range("RNG_Edition").Value = TranslateMsg(C_sMsgCheckSheetType)
                    Exit Sub
                End Select
            Else
                'I am on a previous sheet name, I will upate in that case the number of columns of the linelist type
                'I will use a select case to anticipate if whe have to deal with another type of sheet
                LLNbColData.Item(j) = LLNbColData.Item(j) + 1
                'Here I need to take in account the Geo for the exact column number in one sheet
            End If

            'Updating the column index data (that is the index of each variable names)
            Select Case ClearString(dictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderControl)))
            Case C_sDictControlGeo
                ColumnIndexData.Item(i + 1) = ColumnIndexData.Items(i) + 4
            Case Else
                ColumnIndexData.Item(i + 1) = ColumnIndexData.Items(i) + 1
            End Select

            i = i + 1
        Loop
    End With

End Sub

'SHEET OF TYPE ADM CREATION (Adaptation from lionel's work) ===========================================================================================================================================

Private Sub CreateSheetAdmEntry(wkb As Workbook, sSheetName As String, iSheetStartLine As Integer, _
                                dictData As BetterArray, DictHeaders As BetterArray, LLSheetNameData As BetterArray, _
                                LLNbColData As BetterArray, ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray)

    Dim sActualMainLab As String                 'Actual Main label
    Dim sActualSubLab As String
    Dim sActualVarName As String                 'Actual Variable Name
    Dim sActualFormula As String                 'Actual Variable Choice
    Dim sActualControl As String
    Dim sActualValidationAlert As String
    Dim sActualValidationMessage As String
    Dim sActualMainSec As String
    Dim sActualSubSec As String

    'Previous sections and sub sections
    Dim sPrevMainSec As String
    Dim sPrevSubSec As String

    Dim iCounterSheetAdmLine As Integer
    Dim iCounterDictSheetLine As Integer
    Dim iTotalSheetAdmColumns As Integer

    Dim iPrevLineSubSec As Integer
    Dim iPrevLineMainSec As Integer



    'Add the logo for the first time
    If Not AddedLogo Then
        'Add the Logo
        With wkb.Worksheets(sParamSheetAdmin)

            On Error Resume Next
            'Logo (copy from the sheet main, copy can fail, you just continue)
            Application.CutCopyMode = False
            SheetMain.Shapes("SHP_Logo").Copy
            .Paste Destination:=wkb.Worksheets(sParamSheetAdmin).Cells(2, 2)
            Application.CutCopyMode = True
            On Error GoTo 0

            AddedLogo = True
            .Protect Password:=(ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value), DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
        End With
    End If

    iPrevLineMainSec = C_eStartLinesAdmData
    iPrevLineSubSec = C_eStartLinesAdmData
    iCounterSheetAdmLine = C_eStartLinesAdmData
    iCounterDictSheetLine = iSheetStartLine
    iTotalSheetAdmColumns = LLNbColData.Items(LLSheetNameData.IndexOf(sSheetName))

    sPrevMainSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
    sPrevSubSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))

    With wkb.Worksheets(sSheetName)

        'FontSizes of Adms
        .Cells.Font.Size = C_iAdmSheetFontSize

        'Updating the values
        Do While (iCounterDictSheetLine <= iSheetStartLine + iTotalSheetAdmColumns - 1)

            sActualMainLab = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainLab))
            sActualSubLab = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubLab))
            sActualVarName = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderVarName))
            sActualFormula = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderFormula))
            sActualControl = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderControl))
            sActualValidationAlert = ClearString(dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderAlert)))
            sActualValidationMessage = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMessage))
            sActualMainSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
            sActualSubSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))


            WriteBorderLines .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 3), iWeight:=xlHairline, sColor:="DarkBlue"

            'Update the previous sub sections and

            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 2).Value = sActualMainLab
            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 2).Interior.color = vbWhite
            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 2).Font.color = Helpers.GetColor("BlueButton")
            WriteBorderLines .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 2), iWeight:=xlHairline, sColor:="DarkBlue"
            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 3).Name = sActualVarName

            'Update values for the first time we have the sections
            If iCounterSheetAdmLine = C_eStartLinesAdmData Then
                .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData).Value = sActualMainSec
                .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 1).Value = sActualSubSec
            End If

            If sPrevSubSec <> sActualSubSec Then
                'I am on a new sub section for the same section
                .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 1).Value = sActualSubSec
                'Merge the sub sections area
                'I have to test I am not on the first column since it is possible that initialized value differed from
                'the actual first value due to changes (taking in account the geo)

                BuildSubSectionVMerge Wksh:=wkb.Worksheets(sSheetName), _
        iColumn:=C_eStartColumnAdmData + 1, iLineFrom:=iPrevLineSubSec, _
        iLineTo:=iCounterSheetAdmLine

                'update previous columns
                sPrevSubSec = sActualSubSec
                iPrevLineSubSec = iCounterSheetAdmLine

            ElseIf sPrevMainSec <> sActualMainSec Then
                'Update sub sections on new Main sections too

                .Cells(iCounterSheetAdmLine, C_eStartLinesAdmData + 1).Value = sActualSubSec
                BuildSubSectionVMerge Wksh:=wkb.Worksheets(sSheetName), _
        iColumn:=C_eStartColumnAdmData + 1, iLineFrom:=iPrevLineSubSec, _
        iLineTo:=iCounterSheetAdmLine

                'update previous columns
                sPrevSubSec = sActualSubSec
                iPrevLineSubSec = iCounterSheetAdmLine

                'Build last section
            ElseIf (iCounterDictSheetLine = iSheetStartLine + iTotalSheetAdmColumns - 1) Then

                BuildSubSectionVMerge Wksh:=wkb.Worksheets(sSheetName), _
        iColumn:=C_eStartColumnAdmData + 1, iLineFrom:=iPrevLineSubSec, _
        iLineTo:=iCounterSheetAdmLine + 1
            End If

            'Do the same for the section
            If sPrevMainSec <> sActualMainSec Then
                'I am on a new Main Section, update the value of the section
                .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData).Value = sActualMainSec

                'Merge the previous area
                BuildMainSectionVMerge Wksh:=wkb.Worksheets(sSheetName), iLineFrom:=iPrevLineMainSec, _
        iColumnFrom:=C_eStartColumnAdmData, iLineTo:=iCounterSheetAdmLine

                'Update the previous columns
                sPrevMainSec = sActualMainSec
                iPrevLineMainSec = iCounterSheetAdmLine
            ElseIf (iCounterDictSheetLine = iSheetStartLine + iTotalSheetAdmColumns - 1) Then

                'I am on the same main section, I will test if I am not on the last column, if it is the case, merge the area
                BuildMainSectionVMerge Wksh:=wkb.Worksheets(sSheetName), _
        iLineFrom:=iPrevLineMainSec, iColumnFrom:=C_eStartColumnAdmData, _
        iLineTo:=iCounterSheetAdmLine + 1
            End If

            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData).EntireColumn.AutoFit
            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 1).EntireColumn.AutoFit
            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 2).EntireColumn.AutoFit
            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 3).ColumnWidth = 30
            .Cells(iCounterSheetAdmLine, C_eStartColumnAdmData + 3).Locked = False


            If sActualControl = C_sDictControlChoice Then
                'Add list if the choice is not empty
                Call AddChoices(wkb, sSheetName, iCounterSheetAdmLine, C_eStartColumnAdmData + 3, _
                                ChoicesListData, ChoicesLabelsData, sActualFormula, _
                                sActualValidationAlert, sActualValidationMessage)
            End If


            'Add the Column index for those variable
            wkb.Worksheets(C_sParamSheetDict).Cells(iCounterDictSheetLine + 1, DictHeaders.Length + 1).Value = iCounterSheetAdmLine '+1 on lines because of headers of the dictionary


            iCounterSheetAdmLine = iCounterSheetAdmLine + 1
            iCounterDictSheetLine = iCounterDictSheetLine + 1
        Loop

        WriteBorderLines .Range(.Cells(C_eStartLinesAdmData, C_eStartColumnAdmData), .Cells(iCounterSheetAdmLine - 1, C_eStartColumnAdmData + 3)), _
        iWeight:=xlThin, sColor:="DarkBlue"

        .Protect Password:=(ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value), DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    End With
End Sub

'SHEET OF TYPE LINELIST CREATION ======================================================================================================================================================================


Private Sub CreateSheetLLDataEntry(wkb As Workbook, sSheetName As String, iSheetStartLine As Integer, _
                                   dictData As BetterArray, DictHeaders As BetterArray, LLSheetNameData As BetterArray, _
                                   LLNbColData As BetterArray, ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, _
                                   VarNameData As BetterArray, ColumnIndexData As BetterArray, FormulaData As BetterArray, _
                                   SpecCharData As BetterArray, ChoiceAutoVarData As BetterArray, AllSheetNamesData As BetterArray, _
                                   ByRef iNbshifted As Integer)

    'DictData: Dictionary data
    'DictHeaders: Dictionary Headers
    'sSheetName: the actual sheet on which we need to do some stuffs
    'iSheetStartLine: Starting line of the sheet in the Dictionnary

    Dim sPrevMainSec            As String        'Previous mainlabel and sub label titles to track if the labels have changed
    Dim sPrevSubSec             As String        'Previous Sub section
    Dim iCounterSheetLLCol      As Integer       'Counter of Columns in one Sheet in the linelist
    Dim iCounterDictSheetLine   As Integer       'Counter of lines in the dictionnary sheet corresponding to values
    Dim iPrevColMainSec         As Integer       'Previous column where the main label stops
    Dim iPrevColSubSec          As Integer       'Previous column where the sub label stops
    Dim iTotalLLSheetColumns    As Integer       'Total number of columns to add on one sheet of type Linelist
    Dim iChoiceCol              As Integer
    Dim iGoToCol                As Long          'Column for the Goto in the choice auto sheet
    Dim iGoToRow                As Long          'Row for the Goto section in the choice auto sheet

    'Those variables are for readability in the future
    Dim sActualMainLab As String                 'Actual main label of a linelist type sheet
    Dim sActualSubLab As String                  'Actual sub label of a linelist type sheet
    Dim sActualVarName As String                 'Actual variable name of a linelist type sheet
    Dim sActualMainSec As String                 'Actual main section the linelist
    Dim sActualSubSec As String                  'Actual sub section of the linelist
    Dim sActualNote As String
    Dim sActualType As String
    Dim sActualControl As String
    Dim sActualChoice As String                  'current choose choice
    Dim sActualMin As String                     'current min
    Dim sActualMax As String                     'current Max
    Dim sActualValidationAlert As String
    Dim sActualValidationMessage As String
    Dim sActualStatus As String
    Dim bCmdGeoExist As Boolean
    Dim sActualFormula As String
    Dim sFormula As String                       'Formula after correcting and cleaning
    Dim sFormulaMin As String                    'Formula for min
    Dim sFormulaMax As String                    'Formula for max
    Dim LoRng As Range                           'Range of the listobject for one table
    Dim rng As Range                             'Range for various headers
    Dim LoFiltRng As Range                       'Range of the listobject in the filtered table
    Dim bLockData As Boolean
    Dim sChoiceAutoName As String

    'The table name of the listobject
    Dim sTableName As String

    'Update the existence of the Geo button
    bCmdGeoExist = False

    If (LLSheetNameData.IndexOf(sSheetName) < 0) Then
        SheetMain.Range("RNG_Edition").Value = "Logging Error 2"
        Exit Sub
    End If

    'Here I am really sure it is a linelist sheet Type before going foward
    iCounterSheetLLCol = 1
    iCounterDictSheetLine = iSheetStartLine
    iPrevColMainSec = 1
    iPrevColSubSec = 1
    iTotalLLSheetColumns = LLNbColData.Items(LLSheetNameData.IndexOf(sSheetName))
    sPrevMainSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
    sPrevSubSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))
    sTableName = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderTableName))

    'Column for the GoTo Section
    With wkb.Worksheets(C_sSheetChoiceAuto)
        iGoToCol = .Cells(C_eStartlinesListAuto, .Columns.Count).End(xlToLeft).Column + 2
        'Rows for the GotTo Section
        iGoToRow = C_eStartlinesListAuto + 1
        .Cells(iGoToRow, iGoToCol).Value = TranslateLLMsg("MSG_SelectSection") & ": " & sPrevMainSec
        .Cells(iGoToRow - 1, iGoToCol).Value = "GoTo" 'This will probably change in the future
    End With

    'Continue adding the columns unless the total number of columns to add is reached
    With wkb.Worksheets(sSheetName)

        'INITIALISATIONS AND ADDING COMMANDS___________________________________________________________________________________________________________________________________________________________

        'Adding required buttons

        'Show Hide Button
        Call DesignerBuildListHelpers.AddCmd(wkb, sSheetName, _
                                             .Cells(1, 1).Left + C_iCmdWidth + 20, _
                                             .Cells(1, 1).Top, _
                                             C_sShpShowHide, _
                                             "Show/Hide", _
                                             C_iCmdWidth, C_iCmdHeight, _
                                             C_sCmdShowHideName)
        'Add 200 Rows Button
        Call DesignerBuildListHelpers.AddCmd(wkb, sSheetName, _
                                             .Cells(2, 1).Left + C_iCmdWidth + 20, _
                                             .Cells(2, 1).Top + 5, _
                                             C_sShpAddRows, _
                                             "Add rows", _
                                             C_iCmdWidth, C_iCmdHeight, _
                                             C_sCmdAddRowsName)

        'Add Command to clear filters
        Call DesignerBuildListHelpers.AddCmd(wkb, sSheetName, _
                                             .Cells(3, 1).Left + C_iCmdWidth + 20, _
                                             .Cells(3, 1).Top + 10, _
                                             C_sShpClearFilters, _
                                             "Add rows", _
                                             C_iCmdWidth, C_iCmdHeight + 5, _
                                             C_sCmdClearFilters)

        'All the cells font size at 9
        .Cells.Font.Size = C_iLLSheetFontSize

        Do While (iCounterDictSheetLine <= iSheetStartLine + iTotalLLSheetColumns - 1)

            wkb.Worksheets(C_sParamSheetDict).Cells(iCounterDictSheetLine + 1 + iNbshifted, DictHeaders.Length + 1).Value = iCounterSheetLLCol '+1 on DictSheetLine because of headers, iNbShifted to take in account Geo

            bLockData = False                    'lock or not the data in one cell

            'First, accessing actual values ussing the dicitonary data and its corrresponding headers
            sActualVarName = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderVarName))
            sActualMainLab = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainLab))
            sActualSubLab = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubLab))
            sActualNote = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderNote))

            sActualMainSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
            sActualSubSec = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))
            sActualStatus = ClearString(dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderStatus)))

            sActualType = ClearString(dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderType)))
            sActualControl = ClearString(dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderControl)), bremoveHiphen:=False)
            sActualFormula = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderFormula))
            'For actual choices, we can tolerate _ or - in the string names
            sActualChoice = ClearString(sActualFormula, bremoveHiphen:=False)


            sActualMin = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMin))
            sActualMax = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMax))
            sActualValidationAlert = ClearString(dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderAlert)))
            sActualValidationMessage = dictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMessage))

            'Adding the control
            .Cells(C_eStartLinesLLMainSec - 1, iCounterSheetLLCol).Value = sActualControl
            .Cells(C_eStartLinesLLMainSec - 1, iCounterSheetLLCol).Font.color = vbWhite
            .Cells(C_eStartLinesLLMainSec - 1, iCounterSheetLLCol).FormulaHidden = True
            .Cells(C_eStartLinesLLMainSec - 1, iCounterSheetLLCol).Locked = True


            'SETTING HEADERS _____________________________________________________________________________________________________________________________________________________________________________

            'Before doing some changes, we need to update the sub-section correspondingly
            'in case whe have the geo control. When the Control is Geo, the subsection label is
            'The main section label if there is no one

            'Geo Titles or Customs --------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Select Case sActualControl
            Case C_sDictControlGeo
                If sActualSubSec = "" Then
                    sActualSubSec = sActualMainLab
                End If
            Case C_sDictControlCustom
                'In case we have custom variables, let the headers as free text for future
                'modifications by the user
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Locked = False
            Case C_sDictControlForm, C_sDictControlCaseWhen
                sActualSubLab = IIf(sActualSubLab <> vbNullString, sActualSubLab & Chr(10) & sCalculatedForm, sCalculatedForm)
            End Select

            'Adding the headers of the table ---------------------------------------------------------------------------------------------------------
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Name = sActualVarName
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Value = DesignerBuildListHelpers.AddSpaceToHeaders(wkb, sActualMainLab, sSheetName, C_eStartLinesLLData)
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).VerticalAlignment = xlTop

            'Adding the sub-label if needed Chr(10) is the return to line character the sublabel is in gray------------------
            If sActualSubLab <> "" Then
                Call DesignerBuildListHelpers.AddSubLab(wkb.Worksheets(sSheetName), C_eStartLinesLLData, _
                                                        iCounterSheetLLCol, sActualMainLab, _
                                                        sActualSubLab)
            End If

            'Adding the notes
            If sActualNote <> "" Then
                Call DesignerBuildListHelpers.AddNotes(wkb.Worksheets(sSheetName), C_eStartLinesLLData, _
                                                       iCounterSheetLLCol, sActualNote)
            End If


            'Adding the sections and sub-section and merging

            'First, update the values the first time encounter thems
            If iCounterSheetLLCol = 1 Then
                .Cells(C_eStartLinesLLSubSec, iCounterSheetLLCol).Value = sActualSubSec
                .Cells(C_eStartLinesLLMainSec, iCounterSheetLLCol).Value = sActualMainSec
            End If


            If sPrevSubSec <> sActualSubSec Then
                'I am on a new sub section for the same section
                .Cells(C_eStartLinesLLSubSec, iCounterSheetLLCol).Value = sActualSubSec
                'Merge the sub sections area
                'I have to test I am not on the first column since it is possible that initialized value differed from
                'the actual first value due to changes (taking in account the geo)

                BuildSubSectionHMerge Wksh:=wkb.Worksheets(sSheetName), iLine:=C_eStartLinesLLSubSec, iColumnFrom:=iPrevColSubSec, _
        iColumnTo:=iCounterSheetLLCol

                'update previous columns
                sPrevSubSec = sActualSubSec
                iPrevColSubSec = iCounterSheetLLCol

            ElseIf sPrevMainSec <> sActualMainSec Then
                'Update sub sections on new Main sections too

                .Cells(C_eStartLinesLLSubSec, iCounterSheetLLCol).Value = sActualSubSec
                BuildSubSectionHMerge Wksh:=wkb.Worksheets(sSheetName), iLine:=C_eStartLinesLLSubSec, iColumnFrom:=iPrevColSubSec, _
        iColumnTo:=iCounterSheetLLCol

                'update previous columns
                sPrevSubSec = sActualSubSec
                iPrevColSubSec = iCounterSheetLLCol
                'Build last Section on last column
            ElseIf iCounterDictSheetLine = iSheetStartLine + iTotalLLSheetColumns - 1 Then
                BuildSubSectionHMerge Wksh:=wkb.Worksheets(sSheetName), iLine:=C_eStartLinesLLSubSec, iColumnFrom:=iPrevColSubSec, _
        iColumnTo:=iCounterSheetLLCol + 1
            End If

            'NEW SECTION
            'Do the same for the section
            If sPrevMainSec <> sActualMainSec Then
                'I am on a new Main Section, update the value of the section
                .Cells(C_eStartLinesLLMainSec, iCounterSheetLLCol).Value = sActualMainSec

                'GOTO : Here I update the list to set as validation for the "GOTO"
                iGoToRow = iGoToRow + 1
                wkb.Worksheets(C_sSheetChoiceAuto).Cells(iGoToRow, iGoToCol).Value = TranslateLLMsg("MSG_SelectSection") & ": " & sActualMainSec

                'Merge the previous area
                BuildMainSectionHMerge Wksh:=wkb.Worksheets(sSheetName), iLineFrom:=C_eStartLinesLLMainSec, _
        iColumnFrom:=iPrevColMainSec, iLineTo:=C_eStartLinesLLSubSec, iColumnTo:=iCounterSheetLLCol

                'Update the previous columns
                sPrevMainSec = sActualMainSec
                iPrevColMainSec = iCounterSheetLLCol
            ElseIf (iCounterDictSheetLine = iSheetStartLine + iTotalLLSheetColumns - 1) Then
                'I am on the same main section, I will test if I am not on the last column, if it is the case, merge the area
                BuildMainSectionHMerge Wksh:=wkb.Worksheets(sSheetName), _
        iLineFrom:=C_eStartLinesLLMainSec, iColumnFrom:=iPrevColMainSec, _
        iColumnTo:=iCounterSheetLLCol + 1, iLineTo:=C_eStartLinesLLSubSec
            End If

            'STATUS, TYPE and CONTROLS ==========================================================================================================================

            'Updating the notes according to the column's Status ---------------------------------------------------------------------------------------------
            Call DesignerBuildListHelpers.AddStatus(wkb.Worksheets(sSheetName), _
                                                    C_eStartLinesLLData, iCounterSheetLLCol, sActualNote, _
                                                    sActualStatus, "Mandatory data")

            'Building the Column Controls ----------------------------------------------------------------------------

            Select Case sActualControl

            Case C_sDictControlChoice

                'Add list if the choice is not emptyy
                If sActualChoice <> vbNullString Then
                    Call DesignerBuildListHelpers.AddChoices(wkb, sSheetName, _
                                                             C_eStartLinesLLData + 2, iCounterSheetLLCol, _
                                                             ChoicesListData, ChoicesLabelsData, sActualChoice, _
                                                             sActualValidationAlert, sActualValidationMessage)
                End If
                'Insert the other columns in case we are with a geo

            Case C_sDictControlGeo
                'First, Geocolumns are in orange
                Call DesignerBuildListHelpers.AddGeo(wkb, dictData, DictHeaders, sSheetName, _
                                                     C_eStartLinesLLData, iCounterSheetLLCol, _
                                                     C_eStartLinesLLSubSec, iCounterDictSheetLine, sActualVarName, sActualValidationMessage, _
                                                     iNbshifted)

                'The geocolumn induce four new columns (I will add 3, keeping the 1 at the end of the loop for next variable)


                iCounterSheetLLCol = iCounterSheetLLCol + 3
                iNbshifted = iNbshifted + 3
                sActualVarName = C_sAdmName & "4" & "_" & sActualVarName

                'Add the GeoButton only one time
                If Not bCmdGeoExist Then
                    Call DesignerBuildListHelpers.AddCmd(wkb, sSheetName, _
                                                         .Cells(1, 1).Left + 5, .Cells(2, 1).Top + 5, _
                                                         C_sShpGeo, _
                                                         "GEO", _
                                                         C_iCmdWidth, C_iCmdHeight, _
                                                         C_sCmdShowGeoApp, "Orange", "Black")
                    bCmdGeoExist = True
                End If

            Case C_sDictControlChoiceAuto

                sChoiceAutoName = C_sDictControlChoiceAuto & "_" & sActualChoice

                If Not ListObjectExists(wkb.Worksheets(C_sSheetChoiceAuto), "o" & sChoiceAutoName) Then
                    'Add the list_auto column in the worksheet list_auto_
                    With wkb.Worksheets(C_sSheetChoiceAuto)
                        iChoiceCol = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1

                        .Cells(C_eStartlinesListAuto, iChoiceCol + 1).Value = sChoiceAutoName

                        Set LoRng = .Range(.Cells(C_eStartlinesListAuto, iChoiceCol + 1), .Cells(C_eStartlinesListAuto + 1, iChoiceCol + 1))
                        .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = "o" & sChoiceAutoName
                        wkb.NAMES.Add Name:=sChoiceAutoName, RefersToR1C1:="=o" & sChoiceAutoName & "[" & sChoiceAutoName & "]"
                    End With
                End If

                'Set the validation for list auto
                Call Helpers.SetValidation(.Cells(C_eStartLinesLLData + 2, iCounterSheetLLCol), "=" & sChoiceAutoName, Helpers.GetValidationType(sActualValidationAlert), sActualValidationMessage)

            Case C_sDictControlHf

                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Interior.color = GetColor("Orange")

            Case C_sDictControlForm, C_sDictControlCaseWhen 'Formulas, are reported to the formula function

                If (sActualFormula <> vbNullString) Then
                    sFormula = sActualFormula

                    If sActualControl = C_sDictControlCaseWhen Then sFormula = ParseCaseWhen(sFormula)

                    sFormula = DesignerBuildListHelpers.ValidationFormula(sFormula, AllSheetNamesData, VarNameData, ColumnIndexData, _
                                                                          FormulaData, SpecCharData, ChoiceAutoVarData, _
                                                                          sActualVarName, wkb.Worksheets(sSheetName), False)

                    'Testing before writing the formula
                    If (sFormula <> vbNullString) Then
                        .Cells(C_eStartLinesLLData + 2, iCounterSheetLLCol).NumberFormat = "General"
                        .Cells(C_eStartLinesLLData + 2, iCounterSheetLLCol).formula = sFormula
                        .Cells(C_eStartLinesLLData + 2, iCounterSheetLLCol).Font.color = GetColor("Grey50")
                        .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Interior.color = GetColor("GreyFormula")
                        bLockData = True         'Lock data for formulas
                    Else
                        'MsgBox "Invalid formula will be ignored : " & sActualFormula & "/" & sActualVarName  'MSG_InvalidFormula
                    End If
                End If

            End Select

            'The type is added after formula validation because we need to take in account the formula before
            'setting the type
            'Formating the Column according to the Column's type -------------------------------------------------------------------------------------------
            Call DesignerBuildListHelpers.AddType(wkb.Worksheets(sSheetName), _
                                                  C_eStartLinesLLData, iCounterSheetLLCol, sActualType)

            'Building Min/Max Validation ----------------------------------------------------------------------------
            If sActualMin <> "" And sActualMax <> "" Then

                'Testing if it is numeric
                sFormulaMin = DesignerBuildListHelpers.ValidationFormula(sActualMin, AllSheetNamesData, VarNameData, ColumnIndexData, _
                                                                         FormulaData, SpecCharData, ChoiceAutoVarData, sActualVarName, wkb.Worksheets(sSheetName), True)
                If sFormulaMin = "" Then
                    'MsgBox "Invalid formula will be ignored : " & sActualMin & " / " & sActualVarName
                Else
                    sFormulaMax = DesignerBuildListHelpers.ValidationFormula(sActualMax, AllSheetNamesData, VarNameData, ColumnIndexData, FormulaData, SpecCharData, _
                                                                             ChoiceAutoVarData, sActualVarName, wkb.Worksheets(sSheetName), True)
                    If sFormulaMax = "" Then
                        'MsgBox "Invalid formula will be ignored : " & sFormulaMax & " / " & sActualVarName
                    End If
                    If (sFormulaMin <> "" And sFormulaMax <> "") Then
                        Call DesignerBuildListHelpers.BuildValidationMinMax(.Cells(C_eStartLinesLLData + 2, iCounterSheetLLCol), _
                                                                            sFormulaMin, sFormulaMax, _
                                                                            GetValidationType(sActualValidationAlert), _
                                                                            sActualType, sActualValidationMessage)
                    End If
                End If
            End If

            'After input every headers, auto fit the columns and unlock data entry part
            .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).Value = sActualVarName

            'List Auto is updated at the end of the buildList process

            .Cells(C_eStartLinesLLData + 2, iCounterSheetLLCol).Locked = bLockData
            Call Helpers.WriteBorderLines(.Range(.Cells(C_eStartLinesLLData, iCounterSheetLLCol), .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol)))


            'Updating the counters
            iCounterSheetLLCol = iCounterSheetLLCol + 1 'Counter of column on one Sheet of type Linelist
            iCounterDictSheetLine = iCounterDictSheetLine + 1 'Counter of lines in the dictionary
            DoEvents
        Loop

        'Formating the variable labels row
        Set rng = .Range(.Cells(C_eStartLinesLLData, 1), .Cells(C_eStartLinesLLData, iCounterSheetLLCol - 1))
        rng.Font.Bold = True
        rng.RowHeight = C_iLLVarLabelHeight

        'Set Column Width of First and Second Column
        .Columns(1).ColumnWidth = C_iLLFirstColumnsWidth
        .Columns(2).ColumnWidth = C_iLLFirstColumnsWidth

        'Set Validation to the Section goto Cell
        Call DesignerBuildListHelpers.BuildGotoArea(wkb, sTableName, sSheetName, iGoToCol)

        'Put the range of variable labels in bold and grey colors
        Set rng = .Range(.Cells(C_eStartLinesLLData + 1, 1), .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol - 1))
        FormatARange rng, sFontColor:="VeryLightGreyBlue", sInteriorColor:="VeryLightGreyBlue"
        rng.Locked = True
        rng.FormulaHidden = True


        'Range of the listobject
        Set LoRng = .Range(.Cells(C_eStartLinesLLData + 1, 1), .Cells(C_eStartLinesLLData + 2, iCounterSheetLLCol - 1))
        'Creating the TableObject that will contain the data entry
        .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = sTableName
        .ListObjects(sTableName).TableStyle = C_sLLTableStyle

        'Set the new range for the table
        Set LoRng = .Range(.Cells(C_eStartLinesLLData + 1, 1), .Cells(C_iNbLinesLLData + C_eStartLinesLLData + 1, iCounterSheetLLCol - 1))
        'Resize for 200 lines entrie
        .ListObjects(sTableName).Resize LoRng
        '   Now Protect the sheet,
        .Protect Password:=(ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value), DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

    End With

    'Tranfert Event code to the worksheet
    TransferCodeWksh wkb:=wkb, sSheetName:=sSheetName, sNameModule:=C_sModLLChange

    'Now on the filtered sheet copy the range of the list object
    With wkb.Worksheets(C_sFiltered & sSheetName)
        Set LoFiltRng = .Range(.Cells(C_eStartLinesLLData + 1, 1), .Cells(C_iNbLinesLLData + C_eStartLinesLLData + 1, iCounterSheetLLCol - 1))
        LoFiltRng.Value = LoRng.Value
        .ListObjects.Add(xlSrcRange, LoFiltRng, , xlYes).Name = C_sFiltered & sTableName
        .ListObjects(C_sFiltered & sTableName).TableStyle = C_sLLTableStyle
    End With
End Sub


