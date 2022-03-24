Attribute VB_Name = "DesignerBuildList"
Option Explicit

'BUILD THE LINELIST ====================================================================================================

'Building the linelist from the different input data
'@DictHeaders: The headers of the dictionnary sheet
'@DictData: Dictionnary data
'@ChoicesHeaders: The headers of the Choices sheet
'@ChoicesData: The choices data
'@ExportData: The export data


Sub BuildList(DictHeaders As BetterArray, DictData As BetterArray, ExportData As BetterArray, _
              ChoicesHeaders As BetterArray, ChoicesData As BetterArray, _
              TransData As BetterArray, sPath As String)

    Dim xlsapp As Excel.Application
    Dim LLNbColData As BetterArray               'Number of columns of a Sheet of type linelist
    Dim LLSheetNameData As BetterArray           'Names of sheets of type linelist
    Dim ChoicesListData As BetterArray           'Choices list
    Dim ChoicesLabelsData As BetterArray         ' Choices labels
    Dim VarnameSheetData As BetterArray
    Dim VarNameData As BetterArray
    Dim ColumnIndexData As BetterArray
    Dim ColumnSheetIndexData As BetterArray
    Dim FormulaData As BetterArray
    Dim SpecCharData As BetterArray
    Dim DictVarName As BetterArray
    Dim iPastingRow As Integer
    Dim sCpte As Integer
    Dim LoRng As Range 'List object's range
    Dim iNbshifted As Integer
    

    Dim iCounterSheet As Integer                'counter for one Sheet
    Dim iSheetStartLine As Integer              'Counter for starting line of the sheet in the dictionary
    Dim i As Integer

    'Instanciating the betterArrays
    Set LLNbColData = New BetterArray
    Set LLSheetNameData = New BetterArray       'Names of sheets of type linelist
    Set ColumnIndexData = New BetterArray
    Set ColumnSheetIndexData = New BetterArray
    Set FormulaData = New BetterArray
    Set SpecCharData = New BetterArray
    Set VarnameSheetData = New BetterArray
    Set VarNameData = New BetterArray
    Set DictVarName = New BetterArray
    
    'lla
    Application.StatusBar = "[" & Space(C_iNumberOfBars) & "]" 'create status ProgressBar
    
    Set xlsapp = New Excel.Application

    With xlsapp 'lla
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Visible = False
        .AutoCorrect.DisplayAutoCorrectOptions = False
        .Workbooks.Add
    End With
    
    DoEvents
    'Now Transferring some designers objects (codes, modules) to the workbook we want to create
    Call DesignerBuildListHelpers.TransferDesignerCodes(xlsapp)
    
    DoEvents
    'DesignerBuildListHelpers.TransterSheet is for sending worksheets from the actual workbook to the first workbook of the instance

    Call DesignerBuildListHelpers.TransferSheet(xlsapp, C_sSheetGeo)
    Call DesignerBuildListHelpers.TransferSheet(xlsapp, C_sSheetPassword)
    Call DesignerBuildListHelpers.TransferSheet(xlsapp, C_sSheetFormulas)
    Call DesignerBuildListHelpers.TransferSheet(xlsapp, C_sSheetLLTranslation)
    

    DoEvents

    'Create special characters data
    FormulaData.FromExcelRange SheetFormulas.ListObjects(C_sTabExcelFunctions).ListColumns("ENG").DataBodyRange, DetectLastColumn:=False
    SpecCharData.FromExcelRange SheetFormulas.ListObjects(C_sTabASCII).ListColumns("TEXT").DataBodyRange, DetectLastColumn:=False
    
    VarNameData.Items = DictData.ExtractSegment(ColumnIndex:=DictHeaders.IndexOf(C_sDictHeaderVarName))


    'Create all the required Sheets in the workbook (Dictionnary, Export, Password, Geo and other sheets defined by the user)
    Call CreateSheets(xlsapp, DictData, DictHeaders, ExportData, _
                      ChoicesHeaders, ChoicesData, TransData, _
                      LLNbColData, ColumnIndexData, LLSheetNameData, _
                      bNotHideSheets:=False)
    DoEvents
    'SheetMain.Range(C_sRngEdition).value = "Created the Sheets"

    'Choices data'Setting the Choices labels and lists
    Set ChoicesListData = New BetterArray
    Set ChoicesLabelsData = New BetterArray
    Set VarnameSheetData = New BetterArray
    
    ChoicesListData.LowerBound = 1
    ChoicesLabelsData.LowerBound = 1

    'Update the values of the labels and list! here I must make sure my Headers contains those values

    If (ChoicesHeaders.IndexOf(C_sChoiHeaderList) <= 0 Or ChoicesHeaders.IndexOf(C_sChoiHeaderLab) <= 0) Then
        SheetMain.Range(C_sRngEdition).value = "Error 1"
        Exit Sub
    End If

    ChoicesListData.Items = ChoicesData.ExtractSegment(ColumnIndex:=ChoicesHeaders.IndexOf(C_sChoiHeaderList))
    ChoicesLabelsData.Items = ChoicesData.ExtractSegment(ColumnIndex:=ChoicesHeaders.IndexOf(C_sChoiHeaderLab))

    iSheetStartLine = 1
    sCpte = 0
    StatusBar_Updater (sCpte)
    
    iNbshifted = 0

    For iCounterSheet = 1 To LLSheetNameData.UpperBound
        sCpte = Round(100 * iCounterSheet / LLSheetNameData.UpperBound, 1)
        'Vector of varnames for one sheet
        VarnameSheetData.Clear
        VarnameSheetData.Items = VarNameData.Slice(iSheetStartLine, iSheetStartLine + LLNbColData.Item(iCounterSheet))
        'Vector of columnIndexes for one sheet (used for the linelist type sheet)
        ColumnSheetIndexData.Clear
        ColumnSheetIndexData.Items = ColumnIndexData.Slice(iSheetStartLine, iSheetStartLine + LLNbColData.Item(iCounterSheet))

        Select Case DictData.Items(iSheetStartLine, DictHeaders.IndexOf(C_sDictHeaderSheetType))
            'On linelist type, build a data entry form
            Case C_sDictSheetTypeLL
                'Create a sheet for data Entry in one sheet of type linelist
                Call CreateSheetLLDataEntry(xlsapp, LLSheetNameData.Item(iCounterSheet), iSheetStartLine, DictData, _
                                         DictHeaders, LLSheetNameData, LLNbColData, ChoicesListData, ChoicesLabelsData, _
                                         VarnameSheetData, ColumnSheetIndexData, FormulaData, SpecCharData, iNbshifted)
                    DoEvents
                    
                    'update the variable names for writing in the dictionary sheet
                    i = 1
                    With xlsapp.worksheets(LLSheetNameData.Item(iCounterSheet))
                        While (.Cells(C_eStartLinesLLData, i).value <> "")
                            DictVarName.Push .Cells(C_eStartLinesLLData, i).Name.Name
                            i = i + 1
                        Wend
                    End With
                    
                     'Now writing the data of varnames to the dictionary
                     With xlsapp.worksheets(C_sParamSheetDict)
                        iPastingRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                        DictVarName.ToExcelRange Destination:=.Cells(iPastingRow + 1, 1)
                        DictVarName.Clear
                     End With
            Case C_sDictSheetTypeAdm
              
                'Create a sheet of type admin entry
                Call CreateSheetAdmEntry(xlsapp, LLSheetNameData.Item(iCounterSheet), iSheetStartLine, DictData, _
                                        DictHeaders, LLSheetNameData, LLNbColData, _
                                        ChoicesListData, ChoicesLabelsData)
                     i = 0
                    With xlsapp.worksheets(LLSheetNameData.Item(iCounterSheet))
                        While (.Cells(C_eStartLinesAdmData + i, 2).value <> "")
                            DictVarName.Push .Cells(C_eStartLinesAdmData + i, 3).Name.Name
                            i = i + 1
                        Wend
                    End With
        
                     'Now writing the data of varnames to the dictionary
                     
                     With xlsapp.worksheets(C_sParamSheetDict)
                        iPastingRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                        DictVarName.ToExcelRange Destination:=.Cells(iPastingRow + 1, 1)
                        DictVarName.Clear
                     End With
            DoEvents
        End Select
        iSheetStartLine = iSheetStartLine + LLNbColData.Item(iCounterSheet)
        
        DoEvents
    StatusBar_Updater (sCpte)
    Next
    
    'Put the dictionnary in a table format
    With xlsapp.worksheets(C_sParamSheetDict)
        .Cells(1, 1).value = C_sDictHeaderVarName
        .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(DictData.Length, DictHeaders.Length + 1)), , xlYes).Name = "o" & ClearString(C_sParamSheetDict)
        .ListObjects("o" & ClearString(C_sParamSheetDict)).Resize .ListObjects("o" & ClearString(C_sParamSheetDict)).Range.CurrentRegion
    End With
    
    Set LLNbColData = Nothing
    Set LLSheetNameData = Nothing      'Names of sheets of type linelist
    Set ColumnIndexData = Nothing
    Set ColumnSheetIndexData = Nothing
    Set FormulaData = Nothing
    Set SpecCharData = Nothing
    Set VarnameSheetData = Nothing
    Set VarNameData = Nothing
    Set ChoicesListData = Nothing
    Set ChoicesLabelsData = Nothing
    Set DictVarName = Nothing
    
    With xlsapp
        '.workSheets("linelist-patient").Select 'lla. On ne sait pas a priori que c'est la feuile linelist-patient. On ne connait pas le nom des feuilles.
        '.workSheets("linelist-patient").Range("A1").Select
        .DisplayAlerts = False
        .ScreenUpdating = False
        '.Visible = True
        .ActiveWindow.DisplayZeros = True
    End With
 
    xlsapp.ActiveWorkbook.SaveAs Filename:=sPath, FileFormat:=xlExcel12, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    xlsapp.Quit
    Set xlsapp = Nothing
    
    Application.StatusBar = "" 'close status ProgressBar
    
End Sub

'CREATE SHEETS IN A LINELIST ============================================================================================

'Create the required Sheet and Hide some of them

'@xlsapp: Excel application
'@DictData: Dictionary Data
'@DictHeaders: Headers of the dictionary
'@ExportData: Export Data
'@LLNbColData: This is a vector that will be updated. It counts for each sheet, the number of columns
'@ColumnIndexData: This a vector that will be updated. It count for each column, the index in the sheet where the column should be
'@LLSheetName: This is a vector that will contain le name of all the sheets
'@bNotHideSheets: For debugging purpose (hide or not dicitonary and Export sheets)

Private Sub CreateSheets(xlsapp As Excel.Application, DictData As BetterArray, DictHeaders As BetterArray, _
                        ExportData As BetterArray, ChoicesHeaders As BetterArray, _
                        ChoicesData As BetterArray, TransData As BetterArray, _
                        LLNbColData As BetterArray, ColumnIndexData As BetterArray, _
                        LLSheetNameData As BetterArray, Optional bNotHideSheets As Boolean = False)
    'LLNbColData: Number of columns for a sheet of type linelist
    'LLSheetNameData: Name of a sheet of type linelist

    Dim i As Integer 'iterators
    Dim j As Integer
    ColumnIndexData.LowerBound = 1

    Dim sPrevSheetName As String 'Previous sheet name

    With xlsapp
        'Workbook already contains Password and formula sheets. Hide them
        .worksheets(C_sSheetPassword).Visible = xlVeryHidden
        .worksheets(C_sSheetFormulas).Visible = xlVeryHidden
        .worksheets(C_sSheetLLTranslation).Visible = xlVeryHidden

        '-------------- Creating the dictionnary sheet from setup
        .worksheets.Add.Name = C_sParamSheetDict
        'Headers of the disctionary
        DictHeaders.ToExcelRange Destination:=.Sheets(C_sParamSheetDict).Cells(1, 1), TransposeValues:=True
        'Data of the dictionary
        DictData.ToExcelRange Destination:=.Sheets(C_sParamSheetDict).Cells(2, 1)
        'Transforming the dictionary in a listobject Table
        .worksheets(C_sParamSheetDict).Columns(1).ClearContents
        .worksheets(C_sParamSheetDict).Visible = bNotHideSheets

        '-------------- Creating the export sheet
        .worksheets.Add.Name = C_sParamSheetExport
        'Headers of the export options
        .worksheets(C_sParamSheetExport).Cells(1, 1).value = "ID"
        .worksheets(C_sParamSheetExport).Cells(1, 2).value = "Lbl"
        .worksheets(C_sParamSheetExport).Cells(1, 3).value = "Pwd"
        .worksheets(C_sParamSheetExport).Cells(1, 4).value = "Actif"
        .worksheets(C_sParamSheetExport).Cells(1, 5).value = "FileName"

        'Adding the data on export parameters
        ExportData.ToExcelRange Destination:=.Sheets(C_sParamSheetExport).Cells(2, 1)
        .worksheets(C_sParamSheetExport).Visible = xlSheetVeryHidden

        '--------- Creating the Choices Sheet
        .worksheets.Add.Name = C_sParamSheetChoices
        ChoicesHeaders.ToExcelRange Destination:=.Sheets(C_sParamSheetChoices).Cells(1, 1), TransposeValues:=True
        ChoicesData.ToExcelRange Destination:=.Sheets(C_sParamSheetChoices).Cells(2, 1)
        .worksheets(C_sParamSheetChoices).Visible = xlSheetVeryHidden

        '--------- Creating the translation sheet
        .worksheets.Add.Name = C_sParamSheetTranslation
        TransData.ToExcelRange Destination:=.Sheets(C_sParamSheetTranslation).Cells(1, 1)
        .worksheets(C_sParamSheetTranslation).Visible = xlSheetVeryHidden
        
        '--------- Adding a temporary sheet for computations
        .worksheets.Add.Name = C_sSheetTemp
        .worksheets(C_sSheetTemp).Visible = xlSheetVeryHidden

        '--------------- adding the other the other sheets in the dictionary to the linelist
        i = 1
        sPrevSheetName = ""
        j = 0
        'Setting the lower bound before entering the loop
        LLNbColData.LowerBound = 1
        LLSheetNameData.LowerBound = 1
        'i will hep move from one values of dictionnary data to another
        While i <= DictData.UpperBound
            If sPrevSheetName <> DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName)) Then

                If sPrevSheetName = "" Then
                    .worksheets(1).Name = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                Else
                    .worksheets.Add(after:=.worksheets(sPrevSheetName)).Name = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                End If
                
                'I am on a new sheet name, I update values
                sPrevSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                
                j = j + 1
                'Here, the column index is the index number of each column in one sheet. I update it when I am on
                'a new sheet
                ColumnIndexData.Item(i) = 1
                LLNbColData.Item(j) = 1
                LLSheetNameData.Push sPrevSheetName
                
                'Tell the use we have created one sheet
                SheetMain.Range(C_sRngEdition).value = TranslateMsg(C_sMsgCreatedSheet) & " " & sPrevSheetName
                'adding sheets depending on the type of the sheet
                Select Case LCase(DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetType)))
                Case C_sDictSheetTypeAdm
                    'This is a admin Sheet, just add it like that (or maybe do some other stuffs later on)
                       
                Case C_sDictSheetTypeLL
                    'Set the rowheight of the first two rows of a linelist type sheet
                    .worksheets(sPrevSheetName).Rows("1:2").RowHeight = C_iLLButtonsRowHeight
                    'Now I split at starting lines and freeze the pane
                    '.Worksheets(sPrevSheetName).Activate
                    .ActiveWindow.DisplayZeros = False
                    .ActiveWindow.SplitColumn = 2
                    .ActiveWindow.SplitRow = C_eStartLinesLLData 'freeze a the starting lines of the linelist data
                    .ActiveWindow.FreezePanes = True
                Case Else
                    SheetMain.Range(C_sRngEdition).value = TranslateMsg(C_sMsgCheckSheetType)
                    Exit Sub
                End Select
            Else
                'I am on a previous sheet name, I will upate in that case the number of columns of the linelist type
                'I will use a select case to anticipate if whe have to deal with another type of sheet
               LLNbColData.Item(j) = LLNbColData.Item(j) + 1
                'Here I need to take in account the Geo for the exact column number in one sheet
            End If

            'Updating the column index data (that is the index of each variable names)
            Select Case ClearString(DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderControl)))
                Case C_sDictControlGeo
                        ColumnIndexData.Item(i + 1) = ColumnIndexData.Items(i) + 4
                Case Else
                        ColumnIndexData.Item(i + 1) = ColumnIndexData.Items(i) + 1
            End Select

            i = i + 1
        Wend

        'Adding the column index to the Dictionary Sheet
        .worksheets(C_sParamSheetDict).Cells(1, DictHeaders.Length + 1).value = C_sDictHeaderIndex
        ColumnIndexData.ToExcelRange .worksheets(C_sParamSheetDict).Cells(2, DictHeaders.Length + 1)
       
    End With
End Sub



'SHEET OF TYPE ADM CREATION (Adaptation from lionel's work) =============================================================================================

 Private Sub CreateSheetAdmEntry(xlsapp As Excel.Application, sSheetName As String, iSheetStartLine As Integer, _
                                 DictData As BetterArray, DictHeaders As BetterArray, LLSheetNameData As BetterArray, _
                                 LLNbColData As BetterArray, ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray)

    Dim sActualMainLab As String 'Actual Main label
    Dim sActualVarName As String 'Actual Variable Name
    Dim sActualChoice As String  'Actual Variable Choice
    Dim sValidationList As String
    Dim sActualControl As String
    Dim sActualValidationAlert As String
    Dim sActualValidationMessage As String
    Dim sActualSubLab As String
    
    
    Dim iCounterSheetAdmLine As Integer
    Dim iCounterDictSheetLine As Integer
    Dim iTotalSheetAdmColumns As Integer
    
    
    iCounterSheetAdmLine = C_eStartLinesAdmData
    iCounterDictSheetLine = iSheetStartLine
    iTotalSheetAdmColumns = LLNbColData.Items(LLSheetNameData.IndexOf(sSheetName))


    With xlsapp.worksheets(sSheetName)
        'Adding the buttons
        
        'Import migration buttons
          Call DesignerBuildListHelpers.AddCmd(xlsapp, sSheetName, _
            .Cells(2, 10).Left, .Cells(2, 1).Top, C_sShpImpMigration, _
            "Import for Migration", _
            C_iCmdWidth + 10, C_iCmdHeight + 20, C_sCmdImportMigration)
        
        
        'Export migration buttons
         Call DesignerBuildListHelpers.AddCmd(xlsapp, sSheetName, _
            .Cells(2, 10).Left + C_iCmdWidth + 20, .Cells(2, 1).Top, C_sShpExpMigration, _
            "Export for Migration", _
            C_iCmdWidth + 10, C_iCmdHeight + 20, C_sCmdExportMigration)

        'Export Button
        Call DesignerBuildListHelpers.AddCmd(xlsapp, sSheetName, _
            .Cells(2, 10).Left + 2 * C_iCmdWidth + 40, .Cells(2, 1).Top, C_sShpExport, _
            "Export", _
            C_iCmdWidth + 10, C_iCmdHeight + 20, C_sCmdExport)

         
        Call DesignerBuildListHelpers.AddCmd(xlsapp, sSheetName, _
            .Cells(2, 10).Left + 3 * C_iCmdWidth + 60, .Cells(2, 1).Top, C_sShpDebug, _
            "Debug", _
            C_iCmdWidth + 10, C_iCmdHeight + 20, C_sCmdDebug, sShpColor:="Orange", sShpTextColor:="Black")
        
        
        'Logo (copy from the sheet main)
        SheetMain.Shapes("SHP_Logo").Copy
        .Select
        xlsapp.ActiveWindow.DisplayGridlines = False
        .Cells(2, 2).Select
        .Paste
        'Validations will not work if don't deselect
        .Cells(1, 1).Select
        xlsapp.CutCopyMode = False

        'FontSizes of Adms
        .Cells.Font.Size = C_iAdmSheetFontSize
        
        'Updating the values
        While (iCounterDictSheetLine <= iSheetStartLine + iTotalSheetAdmColumns - 1)
        
            sActualMainLab = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainLab))
            sActualSubLab = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubLab))
            sActualVarName = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderVarName))
            sActualChoice = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderChoices))
            sActualControl = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderControl))
            sActualValidationAlert = ClearString(DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderAlert)))
            sActualValidationMessage = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMessage))
        
        
            Call WriteBorderLines(.Cells(iCounterSheetAdmLine, 3))

            If sActualControl = C_sDictControlTitle Then

                'Huge title
                .Cells(iCounterSheetAdmLine - 5, 2).value = sActualMainLab
                .Cells(iCounterSheetAdmLine - 5, 2).Font.Size = C_iAdmTitleFontSize
                .Cells(iCounterSheetAdmLine - 5, 2).Font.Bold = True
                
                'Sub title
                .Cells(iCounterSheetAdmLine - 3, 2).value = sActualSubLab
                .Cells(iCounterSheetAdmLine - 3, 2).Font.Size = C_iAdmTitleFontSize - 5
                .Cells(iCounterSheetAdmLine - 3, 2).Font.Italic = True
                iCounterSheetAdmLine = iCounterSheetAdmLine - 1
            Else
                .Cells(iCounterSheetAdmLine, 2).value = sActualMainLab
                .Cells(iCounterSheetAdmLine, 2).Interior.Color = Helpers.GetColor("SubSecBlue")
                .Cells(iCounterSheetAdmLine, 2).Font.Color = Helpers.GetColor("MainSecBlue")
                .Cells(iCounterSheetAdmLine, 3).Name = sActualVarName
            End If
        
            If sActualControl = C_sDictControlChoice Then
                'Add list if the choice is not emptyy
                If sActualChoice <> "" Then
                     sValidationList = Helpers.GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoice)
                    If sValidationList <> "" Then
                       Call Helpers.SetValidation(.Cells(iCounterSheetAdmLine, 3), sValidationList, _
                       GetValidationType(sActualValidationAlert), sActualValidationMessage)
                   End If
                End If
            End If
            .Cells(iCounterSheetAdmLine, 2).EntireColumn.Autofit
            .Cells(iCounterSheetAdmLine, 3).ColumnWidth = 30
            iCounterSheetAdmLine = iCounterSheetAdmLine + 1
            iCounterDictSheetLine = iCounterDictSheetLine + 1
        Wend
    End With
End Sub


'SHEET OF TYPE LINELIST CREATION ==================================================================================================================================

Private Sub CreateSheetLLDataEntry(xlsapp As Excel.Application, sSheetName As String, iSheetStartLine As Integer, _
                                 DictData As BetterArray, DictHeaders As BetterArray, LLSheetNameData As BetterArray, _
                                 LLNbColData As BetterArray, ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, _
                                 VarNameData As BetterArray, ColumnIndexData As BetterArray, FormulaData As BetterArray, _
                                 SpecCharData As BetterArray, ByRef iNbshifted As Integer)

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

    'Those variables are for readability in the future
    Dim sActualMainLab As String                 'Actual main label of a linelist type sheet
    Dim sActualSubLab As String                  'Actual sub label of a linelist type sheet
    Dim sActualVarName As String                 'Actual variable name of a linelist type sheet
    Dim sActualMainSec As String                 'Actual main section the linelist
    Dim sActualSubSec As String                  'Actual sub section of the linelist
    Dim sActualNote As String
    Dim sActualType As String
    Dim sActualControl As String
    Dim iDecType As Integer                      'Decimal type number
    Dim sValidationList As String                'List of validations to use for choices
    Dim sActualChoice As String                  'current choose choice
    Dim sActualMin As String                     'current min
    Dim sActualMax As String                     'current Max
    Dim sActualValidationAlert As String
    Dim sActualValidationMessage As String
    Dim sActualStatus As String
    Dim bCmdGeoExist As Boolean
    Dim sActualFormula As String
    Dim sFormula As String 'Formula after correcting and cleaning
    Dim sFormulaMin As String 'Formula for min
    Dim sFormulaMax As String 'Formula for max
    Dim LoRng As Range 'Range of the listobject for one table
    
    Dim bLockData As Boolean


    'Update the existence of the Geo button
    bCmdGeoExist = False

    If (LLSheetNameData.IndexOf(sSheetName) < 0) Then
        SheetMain.Range(C_sRngEdition).value = "Error 2"
        Exit Sub
    End If

    'Here I am really sure it is a linelist sheet Type before going foward
    iCounterSheetLLCol = 1
    iCounterDictSheetLine = iSheetStartLine
    iPrevColMainSec = 1
    iPrevColSubSec = 1
    iTotalLLSheetColumns = LLNbColData.Items(LLSheetNameData.IndexOf(sSheetName))
    sPrevMainSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
    sPrevSubSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))


    'Continue adding the columns unless the total number of columns to add is reached
    With xlsapp.worksheets(sSheetName)

        'INITIALISATIONS AND ADDING COMMANDS ========================================================================================

        .Select
         xlsapp.ActiveWindow.DisplayGridlines = False
         .Cells(1, 1).Select
         xlsapp.CutCopyMode = False
         
        'Adding required buttons
        
        'Show Hide Button
        Call DesignerBuildListHelpers.AddCmd(xlsapp, sSheetName, _
                                            .Cells(2, 1).Left, _
                                            .Cells(2, 1).Top, _
                                            C_sShpShowHide, _
                                            "Show/Hide", _
                                            C_iCmdWidth, C_iCmdHeight, _
                                            C_sCmdShowHideName)
        'Add 200 Rows Button
        Call DesignerBuildListHelpers.AddCmd(xlsapp, sSheetName, _
                                            .Cells(1, 1).Left + C_iCmdWidth + 10, _
                                            .Cells(1, 2).Top, _
                                            C_sShpAddRows, _
                                             "Add rows", _
                                             C_iCmdWidth, C_iCmdHeight, _
                                             C_sCmdAddRowsName)
 
        'All the cells font size at 9
        .Cells.Font.Size = C_iLLSheetFontSize
        
        While (iCounterDictSheetLine <= iSheetStartLine + iTotalLLSheetColumns - 1)
            bLockData = False 'lock or not the data in one cell
            
            'First, accessing actual values ussing the dicitonary data and its corrresponding headers
            sActualVarName = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderVarName))
            sActualMainLab = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainLab))
            sActualSubLab = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubLab))
            sActualNote = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderNote))

            sActualMainSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
            sActualSubSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))
            sActualStatus = ClearString(DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderStatus)))
            
            sActualType = ClearString(DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderType)))
            sActualControl = ClearString(DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderControl)))
            sActualFormula = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderFormula))
            sActualChoice = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderChoices))

           
            sActualMin = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMin))
            sActualMax = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMax))
            sActualValidationAlert = ClearString(DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderAlert)))
            sActualValidationMessage = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMessage))
            
            
         'SETTING HEADERS ===================================================================================================================
        
            'Before doing some changes, we need to update the sub-section correspondingly
            'in case whe have the geo control. When the Control is Geo, the subsection label is
            'The main section label if there is no one

            'Geo Titles or Customs --------------------------------------------------------------------------
            Select Case sActualControl
                Case C_sDictControlGeo
                    If sActualSubSec = "" Then
                        sActualSubSec = sActualMainLab
                    End If
                Case C_sDictControlCustom
                    'In case we have custom variables, let the headers as free text for future
                    'modifications by the user
                    .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Locked = False
            End Select

            'Adding the headers of the table ------------------------------------------------------------------------
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Name = sActualVarName
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).value = DesignerBuildListHelpers.AddSpaceToHeaders(xlsapp, sActualMainLab, sSheetName, C_eStartLinesLLData)
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).VerticalAlignment = xlTop

            'Adding the sub-label if needed Chr(10) is the return to line character the sublabel is in gray------------------
            If sActualSubLab <> "" Then
                Call DesignerBuildListHelpers.AddSubLab(xlsapp.worksheets(sSheetName), C_eStartLinesLLData, _
                                                   iCounterSheetLLCol, sActualMainLab, _
                                                   sActualSubLab)
            End If

            'Adding the notes
            If sActualNote <> "" Then
                Call DesignerBuildListHelpers.AddNotes(xlsapp.worksheets(sSheetName), C_eStartLinesLLData, _
                                                  iCounterSheetLLCol, sActualNote)
            End If
            
            
            'Adding the sections and sub-section and merging
            
            'First, update the values the first time encounter thems
            If iCounterSheetLLCol = 1 Then
                .Cells(C_eStartLinesLLSubSec, iCounterSheetLLCol).value = sActualSubSec
                .Cells(C_eStartLinesLLMainSec, iCounterSheetLLCol).value = sActualMainSec
            End If
            
            If sPrevSubSec <> sActualSubSec Then
                'I am on a new sub section for the same section
                .Cells(C_eStartLinesLLSubSec, iCounterSheetLLCol).value = sActualSubSec
                'Merge the sub sections area
                'I have to test I am not on the first column since it is possible that initialized value differed from
                'the actual first value due to changes (taking in account the geo)

                If (iCounterSheetLLCol = 1) Then 'The first column is a geoColumn with no value for the sublabel
                    Call DesignerBuildListHelpers.BuildMergeArea(xlsapp.worksheets(sSheetName), _
                                         C_eStartLinesLLSubSec, _
                                        iPrevColSubSec, iCounterSheetLLCol + 1)
                Else
                    'Otherwise to the same as before but mergin only the sub section part
                    Call DesignerBuildListHelpers.BuildMergeArea(xlsapp.worksheets(sSheetName), _
                                        C_eStartLinesLLSubSec, _
                                        iPrevColSubSec, iCounterSheetLLCol)
                End If

                'update previous columns
                sPrevSubSec = sActualSubSec
                iPrevColSubSec = iCounterSheetLLCol
            End If

            'Do the same for the section
            If sPrevMainSec <> sActualMainSec Then
                'I am on a new Main Section, update the value of the section
                .Cells(C_eStartLinesLLMainSec, iCounterSheetLLCol).value = sActualMainSec
                
                'Merge the previous area
                Call DesignerBuildListHelpers.BuildMergeArea(xlsapp.worksheets(sSheetName), C_eStartLinesLLMainSec, iPrevColMainSec, _
                                    iCounterSheetLLCol, C_eStartLinesLLSubSec)
                
                'Update the previous columns
                sPrevMainSec = sActualMainSec
                iPrevColMainSec = iCounterSheetLLCol
            Else
                'I am on the same main section, I will test if I am not on the column, if it is the case, merge the area
                If (iCounterDictSheetLine = iSheetStartLine + iTotalLLSheetColumns - 1) Then
                    Call DesignerBuildListHelpers.BuildMergeArea(xlsapp.worksheets(sSheetName), _
                                         C_eStartLinesLLMainSec, iPrevColMainSec, _
                                         iCounterSheetLLCol + 1, C_eStartLinesLLSubSec)
                End If
            End If

        'STATUS, TYPE and CONTROLS ==================================================================================================
            .Columns(iCounterSheetLLCol).EntireColumn.Autofit

            'Updating the notes according to the column's Status ----------------------------------------------------------------------------
            Call DesignerBuildListHelpers.AddStatus(xlsapp.worksheets(sSheetName), _
                                    C_eStartLinesLLData, iCounterSheetLLCol, sActualNote, _
                                    sActualStatus, "Mandatory data")

            'Building the Column Controls ----------------------------------------------------------------------------
            'For actual choices, we can tolerate _ or - in the string names
            sActualChoice = ClearString(sActualChoice, bremoveHiphen:=False)

            Select Case sActualControl

                Case C_sDictControlChoice
                    'Add list if the choice is not emptyy
                    If sActualChoice <> "" Then
                       Call DesignerBuildListHelpers.AddChoices(xlsapp.worksheets(sSheetName), _
                                        C_eStartLinesLLData, iCounterSheetLLCol, _
                                        ChoicesListData, ChoicesLabelsData, sActualChoice, _
                                        sActualValidationAlert, sActualValidationMessage)
                    End If
                    'Insert the other columns in case we are with a geo
                Case C_sDictControlGeo
                    'First, Geocolumns are in orange
                    DesignerBuildListHelpers.AddGeo xlsapp, DictData, DictHeaders, sSheetName, _
                                        C_eStartLinesLLData, iCounterSheetLLCol, _
                                        C_eStartLinesLLSubSec, iCounterDictSheetLine, sActualVarName, sActualValidationMessage, _
                                        iNbshifted

                    'The geocolumn induce four new columns (I will add 3, keeping the 1 at the end)
                    iCounterSheetLLCol = iCounterSheetLLCol + 3
                    iNbshifted = iNbshifted + 3

                    'Add the GeoButton
                    If Not bCmdGeoExist Then
                        Call DesignerBuildListHelpers.AddCmd(xlsapp, sSheetName, _
                                           .Cells(1, 1).Left, .Cells(1, 1).Top, _
                                             C_sShpGeo, _
                                             "GEO", _
                                             C_iCmdWidth, C_iCmdHeight, _
                                             C_sCmdShowGeoApp, "Orange", "Black")
                        bCmdGeoExist = True
                    End If

                Case C_sDictControlHf
                    .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Interior.Color = GetColor("Orange")
                Case C_sDictControlForm 'Formulas, are reported to the formula function
                    If (sActualFormula <> "") Then
                        sFormula = DesignerBuildListHelpers.ValidationFormula(sActualFormula, VarNameData, ColumnIndexData, _
                                                            FormulaData, SpecCharData, False)
                    End If
                    'Testing before writing the formula
                    If (sFormula <> "") Then
                        .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).NumberFormat = "General"
                        .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).Formula = sFormula
                        bLockData = True  'Lock data for formulas
                    Else
                        'MsgBox "Invalid formula will be ignored : " & sActualFormula & "/" & sActualVarName  'MSG_InvalidFormula
                    End If
            End Select

            'The type is added after formula validation because we need to take in account the formula before
            'setting the type
            'Formating the Column according to the Column's type -------------------------------------------------------------------------------------------
            Call DesignerBuildListHelpers.AddType(xlsapp.worksheets(sSheetName), _
                                    C_eStartLinesLLData, iCounterSheetLLCol, sActualType)

            'Building Min/Max Validation ----------------------------------------------------------------------------
            If sActualMin <> "" And sActualMax <> "" Then

                'Testing if it is numeric
                sFormulaMin = DesignerBuildListHelpers.ValidationFormula(sActualMin, VarNameData, ColumnIndexData, FormulaData, SpecCharData, True)
                If sFormulaMin = "" Then
                       'MsgBox "Invalid formula will be ignored : " & sActualMin & " / " & sActualVarName
                Else
                    sFormulaMax = DesignerBuildListHelpers.ValidationFormula(sActualMax, VarNameData, ColumnIndexData, FormulaData, SpecCharData, True)
                    If sFormulaMax = "" Then
                            'MsgBox "Invalid formula will be ignored : " & sFormulaMax & " / " & sActualVarName
                    End If
                    If (sFormulaMin <> "" And sFormulaMax <> "") Then
                        Call DesignerBuildListHelpers.BuildValidationMinMax(.Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol), _
                                            sFormulaMin, sFormulaMax, _
                                            GetValidationType(sActualValidationAlert), _
                                            sActualType, sActualValidationMessage)
                    End If
                End If
            End If

            'After input every headers, auto fit the columns and unlock data entry part
            .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).Locked = bLockData

            'Updating the counters
            iCounterSheetLLCol = iCounterSheetLLCol + 1 'Counter of column on one Sheet of type Linelist
            iCounterDictSheetLine = iCounterDictSheetLine + 1 'Counter of lines in the dictionary
            DoEvents
        Wend
        
        'Range of the listobject
        Set LoRng = .Range(.Cells(C_eStartLinesLLData, 1), .Cells(C_eStartLinesLLData + 1, .Cells(C_eStartLinesLLData, Columns.Count).End(xlToLeft).Column))
'        'Creating the TableObject that will contain the data entry
        .ListObjects.Add(xlSrcRange, LoRng, , xlYes).Name = "o" & ClearString(sSheetName)
        .ListObjects("o" & ClearString(sSheetName)).TableStyle = C_sLLTableStyle
        
        'Set the new range for the table
        Set LoRng = .Range(.Cells(C_eStartLinesLLData, 1), .Cells(C_iNbLinesLLData + C_eStartLinesLLData, .Cells(C_eStartLinesLLData, Columns.Count).End(xlToLeft).Column))
        'Resize for 200 lines entrie
        .ListObjects("o" & ClearString(sSheetName)).Resize LoRng
     '   Now Protect the sheet,
        .Protect Password:=(C_sLLPassword), DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                         AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
        'Update the custom dictionary
    End With

    'Tranfert Event code to the worksheet
    Call DesignerBuildListHelpers.TransferCodeWks(xlsapp, sSheetName, C_sModLLChange)

End Sub
