Attribute VB_Name = "DesignerBuildLinelist"
Option Explicit

Const C_StartLineTitle1 As Byte = 3
Const C_StartLineTitle2 As Byte = 4
Const C_TitleLine As Byte = 5
Const C_ligneDeb As Byte = 6

Const C_CmdWidht As Byte = 60
Const C_PWD As String = "1234"

'Preliminary functions for building the linelist ==================================================================================================================================

Private Sub TransfertSheet(xlsapp As Excel.Application, sSheetName As String)
    
    'Since We can't move worksheet from one instance to another
    'we need to save as a temporary file and then move it to another instance.
    DesignerWorkbook.Worksheets(sSheetName).Copy
    
    DoEvents
    
    On Error Resume Next
    Kill Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx"
    On Error GoTo 0
    
    ActiveWorkbook.SaveAs Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx"
    ActiveWorkbook.Close
    
    DoEvents

    With xlsapp
        .Workbooks.Open Filename:=Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx", UpdateLinks:=False
        
        .Sheets(sSheetName).Select
        .Sheets(sSheetName).Copy after:=.Workbooks(1).Sheets(1)
        
        DoEvents
        .Workbooks("Temp.xlsx").Close
    End With
    
    DoEvents
    
    Kill Environ("Temp") & Application.PathSeparator & "LinelistApp" & Application.PathSeparator & "Temp.xlsx"

End Sub

'Tranfer some designers modules

Private Sub TransferDesignerCodes(xlsapp As Excel.Application)

    On Error Resume Next
    Kill (Environ("Temp") & Application.PathSeparator & "LinelistApp")
    MkDir (Environ("Temp") & Application.PathSeparator & "LinelistApp") 'create a folder for sending all the data from designer
    On Error GoTo 0
    
    Dim Wkb As Workbook
    Set Wkb = xlsapp.ActiveWorkbook
    
    DoEvents
        
    'Transfert form is for sending forms from the actual excel workbook to another
    Call DesTransferForm(Wkb, C_sFormGeo)
    Call DesTransferForm(Wkb, C_sFormShowHide)
    Call DesTransferForm(Wkb, C_sFormExport)

    'TransferCode is for sending modules  (Modules) or classes (Classes) from actual excel workbook to another excel workbook
    Call DesTransferCode(Wkb, C_sModLinelist, "Module")
    Call DesTransferCode(Wkb, C_sModGeo, "Module")
    Call DesTransferCode(Wkb, C_sModShowHide, "Module")
    Call DesTransferCode(Wkb, C_sModDesHelpers, "Module")
    Call DesTransferCode(Wkb, C_sModLLHelpers, "Module")
    Call DesTransferCode(Wkb, C_sModDesTrans, "Module")
    Call DesTransferCode(Wkb, C_sModMigration, "Module")
    Call DesTransferCode(Wkb, C_sModConstants, "Module")
    Call DesTransferCode(Wkb, C_sClaBA, "Class")
    
    Set Wkb = Nothing
    
End Sub


'Create the required Sheet and put those required to veryHidden
Private Sub CreateSheetsInLL(xlsapp As Excel.Application, DictData As BetterArray, DictHeaders As BetterArray, _
                            ExportData As BetterArray, LLNbColData As BetterArray, ColumnIndexData As BetterArray, _
                            LLSheetNameData As BetterArray, Optional bNotHideSheets As Boolean = False)
    'LLNbColData: Number of columns for a sheet of type linelist
    'LLSheetNameData: Name of a sheet of type linelist

    Dim i As Integer 'iterators
    Dim j As Integer
    ColumnIndexData.LowerBound = 1

    Dim sPrevSheetName As String 'Previous sheet name
        
    With xlsapp
    
        'Workbook already contains the Geo, Password and formula sheets. Hide them
        .Worksheets(C_sSheetPassword).Visible = xlVeryHidden
        .Worksheets(C_sSheetFormulas).Visible = xlVeryHidden
        
        '-------------- Creating the dictionnary sheet
        .Worksheets.Add.Name = C_sParamSheetDict
        'Headers of the disctionary
        DictHeaders.ToExcelRange Destination:=.Sheets(C_sParamSheetDict).Cells(1, 1), TransposeValues:=True
        'Data of the dictionary
        DictData.ToExcelRange Destination:=.Sheets(C_sParamSheetDict).Cells(2, 1)
        .Worksheets(C_sParamSheetDict).Visible = bNotHideSheets
        
        
        '-------------- Creating the export sheet
        .Worksheets.Add.Name = C_sParamSheetExport
        'Headers of the export options
        .Worksheets(C_sParamSheetExport).Cells(1, 1).value = "ID"
        .Worksheets(C_sParamSheetExport).Cells(1, 2).value = "Lbl"
        .Worksheets(C_sParamSheetExport).Cells(1, 3).value = "Pwd"
        .Worksheets(C_sParamSheetExport).Cells(1, 4).value = "Actif"
        .Worksheets(C_sParamSheetExport).Cells(1, 5).value = "FileName"
        
        'Adding the data on export parameters
        ExportData.ToExcelRange Destination:=.Sheets(C_sParamSheetExport).Cells(2, 1)
        .Sheets(C_sParamSheetExport).Visible = bNotHideSheets 'xlSheetVeryHidden
    
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
                    .Worksheets(1).Name = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                Else
                    .Worksheets.Add(after:=.Worksheets(sPrevSheetName)).Name = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
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
                    .Worksheets(sPrevSheetName).Rows("1:2").RowHeight = C_iLLButtonsRowHeight
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
    End With
End Sub


Private Sub CreateSheetDataEntry(xlsapp As Excel.Application, sSheetName As String, iSheetStartLine As Integer, _
                                 DictData As BetterArray, DictHeaders As BetterArray, LLSheetNameData As BetterArray, _
                                 LLNbColData As BetterArray, ChoicesListData As BetterArray, ChoicesLabelsData As BetterArray, _
                                 VarnameData As BetterArray, ColumnIndexData As BetterArray, FormulaData As BetterArray, _
                                 SpecCharData As BetterArray)

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


    'Update the existence of the Geo button
    bCmdGeoExist = False

    If (LLSheetNameData.IndexOf(sSheetName) < 0) Then
        SheetMain.Range(C_sRngEdition).value = "Error 2"
        Exit Sub
    End If

    'Here I am really sure it is a linelist sheet type
    iCounterSheetLLCol = 1
    iCounterDictSheetLine = iSheetStartLine
    iPrevColMainSec = 1
    iPrevColSubSec = 1
    iTotalLLSheetColumns = LLNbColData.Items(LLSheetNameData.IndexOf(sSheetName))
    sPrevMainSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
    sPrevSubSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))
    


    'Continue adding the columns unless the total number of columns to add is reached
    With xlsapp.Worksheets(sSheetName)
         .Select
         xlsapp.ActiveWindow.DisplayGridlines = False
         .Cells(1, 1).Select
         xlsapp.CutCopyMode = False
         
        'Creating the TableObject that will contain the data entry
        .ListObjects.Add(xlSrcRange, .Range(.Cells(C_eStartLinesLLData, 1), .Cells(C_eStartLinesLLData, iTotalLLSheetColumns)), , xlYes).Name = "o" & ClearString(sSheetName)
        .ListObjects("o" & ClearString(sSheetName)).TableStyle = C_sLLTableStyle
         
        'Adding required buttons
        
            'Show Hide column
        Call AddCmd(xlsapp, sSheetName, .Cells(2, 1).Left, .Cells(2, 1).Top, "SHP_NomVisibleApps", "Show/Hide", C_iCmdWidth, C_iCmdHeight)
        .Shapes("SHP_NomVisibleApps").Fill.ForeColor.RGB = DesLetColor("DarkBlueTitle")
        .Shapes("SHP_NomVisibleApps").Fill.BackColor.RGB = DesLetColor("DarkBlueTitle")
        '.Shapes("SHP_NomVisibleApps").Fill.TwoColorGradient msoGradientHorizontal, 1
        .Shapes("SHP_NomVisibleApps").OnAction = "ClicCmdVisibleName"
        
        Call AddCmd(xlsapp, sSheetName, .Cells(1, 1).Left + C_iCmdWidth + 10, .Cells(1, 2).Top, "SHP_Ajout200L", "Add rows", C_iCmdWidth, C_iCmdHeight)
        .Shapes("SHP_Ajout200L").Fill.ForeColor.RGB = DesLetColor("DarkBlueTitle")
        .Shapes("SHP_Ajout200L").Fill.BackColor.RGB = DesLetColor("DarkBlueTitle")
        '.Shapes("SHP_Ajout200L").Fill.TwoColorGradient msoGradientHorizontal, 1
        .Shapes("SHP_Ajout200L").OnAction = "clicAdd200L"
        
        'All the cells font size at 9
        .Cells.Font.Size = C_iLLSheetFontSize
        
        While (iCounterDictSheetLine <= iSheetStartLine + iTotalLLSheetColumns - 1)
        
            'Adding the Headers with sections, Mainlabel and sub labels
            'First, accessing those values ussing the dicitonary data and its corrresponding headers
            sActualMainSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainSec))
            sActualSubSec = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubSec))
            sActualVarName = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderVarName))
            sActualMainLab = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMainLab))
            sActualSubLab = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderSubLab))
            sActualNote = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderNote))
            sActualControl = ClearString(DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderControl)))
            sActualMin = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMin))
            sActualMax = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMax))
            sActualStatus = ClearString(DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderStatus)))
            sActualChoice = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderChoices))
            sActualFormula = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderFormula))
            
            'Befor doing some changes, we need to update the main label or sub-label correspondingly
            'in case whe have the geo control

            'Geo Titles or Customs --------------------------------------------------------------------------
            Select Case sActualControl
            Case C_sDictControlGeo
                If sActualSubSec = "" Then
                    sActualSubSec = sActualMainLab
                End If
            Case C_sDictControlCustom
                'In case we have custom variables, let the headers as free text for future
                'modifications
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Locked = False
            End Select

            'Adding the headers ------------------------------------------------------------------------
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Name = sActualVarName
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).value = LetWordingWithSpace(xlsapp, sActualMainLab, sSheetName)
            .Cells(C_eStartLinesLLData, iCounterSheetLLCol).VerticalAlignment = xlTop

            'Adding the sub-label if needed Chr(10) is the return to line character the sublabel is in gray------------------
            If sActualSubLab <> "" Then
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).value = _
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).value & Chr(10) & sActualSubLab
    
                'Changing the fontsize of the sublabels
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Characters(Start:=Len(sActualMainLab) + 1, _
                       Length:=Len(sActualSubLab) + 1).Font.Size = C_iLLSheetFontSize - 2
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Characters(Start:=Len(sActualMainLab) + 1, _
                       Length:=Len(sActualSubLab) + 1).Font.Color = DesLetColor("NotesBlue")
            End If
        
            'Adding the notes if needed
            If sActualNote <> "" Then
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).AddComment
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Comment.Text Text:=sActualNote
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Comment.Visible = False
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
                    Call BuildMergeArea(xlsapp.Worksheets(sSheetName), C_eStartLinesLLSubSec, _
                                        iPrevColSubSec, iCounterSheetLLCol + 1)
                Else
                    'Otherwise to the same as before but mergin only the sub section part
                    Call BuildMergeArea(xlsapp.Worksheets(sSheetName), C_eStartLinesLLSubSec, _
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
                Call BuildMergeArea(xlsapp.Worksheets(sSheetName), C_eStartLinesLLMainSec, iPrevColMainSec, _
                                    iCounterSheetLLCol, C_eStartLinesLLSubSec)
                
                'Update the previous columns
                sPrevMainSec = sActualMainSec
                iPrevColMainSec = iCounterSheetLLCol
            Else
                'I am on the same main section, I will test if I am not on the column, if it is the case, merge the area
                If (iCounterDictSheetLine = iSheetStartLine + iTotalLLSheetColumns - 1) Then
                    Call BuildMergeArea(xlsapp.Worksheets(sSheetName), C_eStartLinesLLMainSec, iPrevColMainSec, _
                                         iCounterSheetLLCol + 1, C_eStartLinesLLSubSec)
                End If
            End If


            'Updating the notes according to the column's Status ----------------------------------------------------------------------------
            Select Case sActualStatus
            Case C_sDictStatusMan
                If sActualNote <> "" Then
                    'Update the notes to add the Status
                    .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Comment.Text Text:="Mandatory data" & Chr(10) & sActualNote
                Else
                    .Cells(C_eStartLinesLLData, iCounterSheetLLCol).AddComment
                    .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Comment.Text Text:="Mandatory data"
                    .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Comment.Visible = False
                    'Add comment on status
                End If
            Case C_sDictStatusHid
                'Hidden, hid the actual column
                .Columns(iCounterSheetLLCol).EntireColumn.Hidden = True
            Case C_sDictStatusOpt
                'Do nothing for the moment for optional status
            End Select

            'Formating the Column according to the Column's type-------------------------------------------------------------------------------------------

            sActualType = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderType))
            sActualType = ClearString(sActualType)

            'Check to be sure that the actual type contains decimal
            If InStr(1, sActualType, C_sDictTypeDec) Then
                iDecType = CInt(Replace(sActualType, C_sDictTypeDec, ""))
                sActualType = C_sDictTypeDec
            End If

            Select Case sActualType
                'Text Type
            Case C_sDictTypeText
                .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).NumberFormat = "@"
                'Integer Type
            Case C_sDictTypeInt
                .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).NumberFormat = "0"
                'Date Type
            Case C_sDictTypeDate
                .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).NumberFormat = "d-mmm-yyy"
                'Decimal Type
            Case C_sDictTypeDec
                .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).NumberFormat = "0." & LetDecString(iDecType)
            Case Else
                'It it is not in the previous types, put it in text
                .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).NumberFormat = "@"
            End Select

            'Building the Column Controls ----------------------------------------------------------------------------
            sActualValidationAlert = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderAlert))
            sActualValidationAlert = ClearString(sActualValidationAlert)
            sActualValidationMessage = DictData.Items(iCounterDictSheetLine, DictHeaders.IndexOf(C_sDictHeaderMessage))
            'For actual choices, we can tolerate _ or - in the string names
            sActualChoice = ClearString(sActualChoice, bremoveHiphen:=False)

            Select Case sActualControl

            Case C_sDictControlCho
                'Add list if the choice is not emptyy
                If sActualChoice <> "" Then
                    sValidationList = GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoice)
                    If sValidationList <> "" Then
                        Call LetValidationList(.Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol), _
                                               sValidationList, LetValidationLockType(sActualValidationAlert), _
                                              sActualValidationMessage)
                    End If
                End If
                'Insert the other columns in case we are with a geo
            Case C_sDictControlGeo
                'First, Geocolumns are in orange
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Interior.Color = DesLetColor("Orange")
                'update the columns only for the geo
                Call Add4GeoCol(xlsapp, sSheetName, sActualMainLab, sActualVarName, _
                                iCounterSheetLLCol, sActualValidationMessage)
                iCounterSheetLLCol = iCounterSheetLLCol + 3

                'You add a Geo command if it does not exists
                If Not bCmdGeoExist Then
                    Call AddCmd(xlsapp, sSheetName, .Cells(1, 1).Left, .Cells(1, 1).Top, _
                                "SHP_GeoApps", UCase(C_sDictControlGeo), C_iCmdWidth, C_iCmdHeight)
                    With .Shapes("SHP_GeoApps").Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = DesLetColor("Orange")
                        .BackColor.RGB = DesLetColor("Orange")
                        '.TwoColorGradient msoGradientHorizontal, 1
                    End With

                    .Shapes("SHP_GeoApps").OnAction = "ClicCmdGeoApps"
                    bCmdGeoExist = True

                End If

            Case C_sDictControlHf
                .Cells(C_eStartLinesLLData, iCounterSheetLLCol).Interior.Color = DesLetColor("Orange")
            Case C_sDictControlForm 'Formulas, are reported to the formula function
                If (sActualFormula <> "") Then
                    sFormula = DesControlValidationFormula(sActualFormula, VarnameData, ColumnIndexData, _
                                                        FormulaData, SpecCharData)
                End If
                'Testing before writing the formula
                If (sFormula <> "") Then
                    .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).NumberFormat = "General"
                    .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).Formula = "=" & sFormula
                    On Error Resume Next
                    .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).Formula2 = "=" & sFormula 'Seems like formula only induce error on some computers
                    On Error GoTo 0
                    .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).Locked = True
                Else
                    MsgBox "Invalid formula will be ignored : " & sActualFormula & "/" & sActualVarName  'MSG_InvalidFormula
                End If
            End Select

            'Building Min/Max Validation ----------------------------------------------------------------------------
            If sActualMin <> "" And sActualMax <> "" Then

                'Testing if it is numeric
                sFormulaMin = DesControlValidationFormula(sActualMin, VarnameData, ColumnIndexData, FormulaData, SpecCharData)
                If sFormulaMin = "" Then
                        MsgBox "Invalid formula will be ignored : " & sActualMin & " / " & sActualVarName
                Else
                        sFormulaMax = DesControlValidationFormula(sActualMax, VarnameData, ColumnIndexData, FormulaData, SpecCharData)
                        If sFormulaMax = "" Then
                            MsgBox "Invalid formula will be ignored : " & sFormulaMax & " / " & sActualVarName
                        End If
                    If (sFormulaMin <> "" And sFormulaMax <> "") Then
                        sFormulaMin = "=" & sFormulaMin
                        sFormulaMax = "=" & sFormulaMax
                        Call BuildValidationMinMax(.Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol), _
                                            sFormulaMin, sFormulaMax, _
                                            LetValidationLockType(sActualValidationAlert), _
                                            sActualType, sActualValidationMessage)
                    End If
                End If
            End If

            'After input every headers, auto fit the columns and unlock data entry part
            .Columns(iCounterSheetLLCol).EntireColumn.Autofit
            .Cells(C_eStartLinesLLData + 1, iCounterSheetLLCol).Locked = False

            'Updating the counters
            iCounterSheetLLCol = iCounterSheetLLCol + 1
            iCounterDictSheetLine = iCounterDictSheetLine + 1
        Wend
        
        'Resize for 200 lines entry
        .ListObjects("o" & ClearString(sSheetName)).Resize .Range(.Cells(C_eStartLinesLLData, 1), _
        .Cells(C_iNbLinesLLData + C_eStartLinesLLData, .Cells(C_eStartLinesLLData, 1).End(xlToRight).Column))
        
        'Now Protect the sheet
        .Protect Password:=SheetMain.Range(C_sRngLLPassword), DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                         AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
        
    End With

End Sub


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
    
    With xlsapp.Worksheets(sSheetName)
        'Adding the buttons
        
        'Export migration buttons
         Call AddCmd(xlsapp, sSheetName, .Cells(2, 10).Left, .Cells(2, 1).Top, "SHP_ExportMig", _
                    "Export for" & Chr(10) & "migration", C_iCmdWidth + 10, C_iCmdHeight + 10)
        .Shapes("SHP_ExportMig").Fill.ForeColor.RGB = DesLetColor("DarkBlueTitle")
        .Shapes("SHP_ExportMig").Fill.BackColor.RGB = DesLetColor("DarkBlueTitle")
        .Shapes("SHP_ExportMig").OnAction = "clicExportMigration"
                
         'Import migration buttons
         Call AddCmd(xlsapp, sSheetName, .Cells(2, 10).Left + C_iCmdWidth + 20, .Cells(2, 1).Top, "SHP_ImportMig", _
          "Import from" & Chr(10) & "migration", C_iCmdWidth + 10, C_iCmdHeight + 10)
        .Shapes("SHP_ImportMig").Fill.ForeColor.RGB = DesLetColor("DarkBlueTitle")
        .Shapes("SHP_ImportMig").Fill.BackColor.RGB = DesLetColor("DarkBlueTitle")
        .Shapes("SHP_ImportMig").OnAction = "clicImportMigration"
        
        'Export Button
        Call AddCmd(xlsapp, sSheetName, .Cells(2, 10).Left + 2 * C_iCmdWidth + 40, .Cells(2, 1).Top, "SHP_Export", _
         "Export", C_iCmdWidth + 10, C_iCmdHeight + 10)
        .Shapes("SHP_Export").Fill.ForeColor.RGB = DesLetColor("DarkBlueTitle")
        .Shapes("SHP_Export").Fill.BackColor.RGB = DesLetColor("DarkBlueTitle")
        '.Shapes("SHP_Export").Fill.TwoColorGradient msoGradientHorizontal, 1
        .Shapes("SHP_Export").OnAction = "clicExport"
        
        'Logo (copy from the sheet main)
        SheetMain.Shapes("SHP_Logo").Copy
        .Select
        xlsapp.ActiveWindow.DisplayGridlines = False
        .Cells(2, 2).Select
        .Paste
        'Validations will not work if don't deselect
        .Cells(1, 1).Select
        xlsapp.CutCopyMode = False
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
        
        
            .Cells(iCounterSheetAdmLine, 2).value = sActualMainLab
            .Cells(iCounterSheetAdmLine, 2).Interior.Color = DesLetColor("SubLabBlue")
            .Cells(iCounterSheetAdmLine, 2).Font.Color = DesLetColor("MainLabBlue")
            .Cells(iCounterSheetAdmLine, 3).Name = sActualVarName
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
            End If
        
            If sActualControl = C_sDictControlCho Then
                'Add list if the choice is not emptyy
                If sActualChoice <> "" Then
                     sValidationList = GetValidationList(ChoicesListData, ChoicesLabelsData, sActualChoice)
                    If sValidationList <> "" Then
                       Call LetValidationList(.Cells(iCounterSheetAdmLine, 3), sValidationList, _
                       LetValidationLockType(sActualValidationAlert), sActualValidationMessage)
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

Private Sub BuildMergeArea(Wksh As Worksheet, iStartLineOne As Integer, iPrevColumn As Integer, _
                        Optional iActualColumn As Integer = -1, Optional iStartLineTwo As Integer = -1, _
                        Optional sColorMainSec As String = "MainLabBlue", _
                        Optional sColorSubSec As String = "SubLabBlue")

    Dim oCell As Object

    With Wksh

        If iActualColumn = -1 Then
            .Cells(iStartLineOne, iPrevColumn).HorizontalAlignment = xlCenter
            .Cells(iStartLineOne, iPrevColumn).Interior.Color = DesLetColor(sColorSubSec)
            Call WriteBorderLines(.Cells(iStartLineOne, iPrevColumn))
            Exit Sub
        End If

        .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1)).Merge
        .Cells(iStartLineOne, iPrevColumn).MergeArea.HorizontalAlignment = xlCenter

        If (iStartLineTwo <> -1) Then
            With .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1))
                .Interior.Color = DesLetColor(sColorMainSec)
                .Font.Color = vbWhite
                .Font.Bold = True
                .Font.Size = C_iLLMainLabFontSize
            End With
            'For the sub sections, if nothing is mentionned,
            'just put them in white (or the same color as the main sections)
            For Each oCell In .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineTwo, iActualColumn - 1))
                  If oCell.value = "" Then
                    oCell.Interior.Color = vbWhite
                  End If
            Next
            Set oCell = Nothing
            'Write borders to the ranges including the subsection
            Call WriteBorderLines(.Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineTwo, iActualColumn - 1)))
        Else
            With .Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1))
                .Interior.Color = DesLetColor(sColorSubSec)
                .Font.Color = DesLetColor(sColorMainSec)
                .Font.Size = C_iLLSubLabFontSize
            End With
            Call WriteBorderLines(.Range(.Cells(iStartLineOne, iPrevColumn), .Cells(iStartLineOne, iActualColumn - 1)))
        End If
    End With

End Sub



'Building the linelist from the different input data
'DictHeaders: The headers of the dictionnary sheet
'DictData: Dictionnary data
'ChoicesHeaders: The headers of the Choices sheet
'ChoicesData: The choices data
'ExportData: The export data


Sub DesBuildList(DictHeaders As BetterArray, DictData As BetterArray, ChoicesHeaders As BetterArray, ChoicesData As BetterArray, ExportData As BetterArray, sPath As String)

    Dim xlsapp As Excel.Application
    Dim LLNbColData As BetterArray               'Number of columns of a Sheet of type linelist
    Dim LLSheetNameData As BetterArray           'Names of sheets of type linelist
    Dim ChoicesListData As BetterArray           'Choices list
    Dim ChoicesLabelsData As BetterArray         ' Choices labels
    Dim VarnameSheetData As BetterArray
    Dim VarnameData As BetterArray
    Dim ColumnIndexData As BetterArray
    Dim ColumnSheetIndexData As BetterArray
    Dim FormulaData As BetterArray
    Dim SpecCharData As BetterArray


    Dim iCounterSheet As Integer                'counter for one Sheet
    Dim iSheetStartLine As Integer              'Counter for starting line of the sheet in the dictionary

    'Instanciating the betterArrays
    Set LLNbColData = New BetterArray
    Set LLSheetNameData = New BetterArray       'Names of sheets of type linelist
    Set ColumnIndexData = New BetterArray
    Set ColumnSheetIndexData = New BetterArray
    Set FormulaData = New BetterArray
    Set SpecCharData = New BetterArray
    Set VarnameSheetData = New BetterArray
    Set VarnameData = New BetterArray

    Set xlsapp = New Excel.Application
    Application.ScreenUpdating = False

    With xlsapp
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Visible = False
        .AutoCorrect.DisplayAutoCorrectOptions = False
        .Workbooks.Add
    End With
    
    DoEvents
    'Now Transferring some designers objects (codes, modules) to the workbook we want to create
    Call TransferDesignerCodes(xlsapp)
    
    DoEvents
    'TransfertSheet is for sending worksheets from the actual workbook to the first workbook of the instance

    Call TransfertSheet(xlsapp, C_sSheetGeo)
    Call TransfertSheet(xlsapp, C_sSheetPassword)
    Call TransfertSheet(xlsapp, C_sSheetFormulas)

    DoEvents

    'Create special characters data
    FormulaData.FromExcelRange SheetFormulas.ListObjects(C_sTabExcelFunctions).ListColumns("ENG").DataBodyRange, DetectLastColumn:=False
    SpecCharData.FromExcelRange SheetFormulas.ListObjects(C_sTabASCII).ListColumns("TEXT").DataBodyRange, DetectLastColumn:=False
    VarnameData.Items = DictData.ExtractSegment(ColumnIndex:=DictHeaders.IndexOf(C_sDictHeaderVarName))


    'Create all the required Sheets in the workbook (Dictionnary, Export, Password, Geo and other sheets defined by the user)
    Call CreateSheetsInLL(xlsapp, DictData, DictHeaders, ExportData, LLNbColData, _
                          ColumnIndexData, LLSheetNameData, bNotHideSheets:=False)

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
    For iCounterSheet = 1 To LLSheetNameData.UpperBound
        VarnameSheetData.Items = VarnameData.Slice(iSheetStartLine, iSheetStartLine + LLNbColData.Item(iCounterSheet))
        ColumnSheetIndexData.Items = ColumnIndexData.Slice(iSheetStartLine, iSheetStartLine + LLNbColData.Item(iCounterSheet))


        Select Case DictData.Items(iSheetStartLine, DictHeaders.IndexOf(C_sDictHeaderSheetType))
            'On linelist type, build a data entry form
            Case C_sDictSheetTypeLL
                'Create a sheet for data Entry in one sheet of type linelist
                Call CreateSheetDataEntry(xlsapp, LLSheetNameData.Item(iCounterSheet), iSheetStartLine, DictData, _
                                         DictHeaders, LLSheetNameData, LLNbColData, ChoicesListData, ChoicesLabelsData, _
                                         VarnameSheetData, ColumnSheetIndexData, FormulaData, SpecCharData)
            Case C_sDictSheetTypeAdm
                'Create a sheet of type admin entry
                Call CreateSheetAdmEntry(xlsapp, LLSheetNameData.Item(iCounterSheet), iSheetStartLine, DictData, _
                                        DictHeaders, LLSheetNameData, LLNbColData, _
                                        ChoicesListData, ChoicesLabelsData)
        End Select
        iSheetStartLine = iSheetStartLine + LLNbColData.Item(iCounterSheet)
    Next
    
    Set LLNbColData = Nothing
    Set LLSheetNameData = Nothing      'Names of sheets of type linelist
    Set ColumnIndexData = Nothing
    Set ColumnSheetIndexData = Nothing
    Set FormulaData = Nothing
    Set SpecCharData = Nothing
    Set VarnameSheetData = Nothing
    Set VarnameData = Nothing
    Set ChoicesListData = Nothing
    Set ChoicesLabelsData = Nothing
 
    xlsapp.ActiveWorkbook.SaveAs Filename:=sPath, FileFormat:=xlExcel12, ConflictResolution:=xlLocalSessionChanges
    xlsapp.Quit
    Set xlsapp = Nothing
End Sub

'Set the type of a validation of a cell giving the name of the validation in the
'dictionary
Function LetValidationLockType(sValidationLockType As String) As Byte

    LetValidationLockType = 3                    'liste de validation info, warning ou erreur
    If sValidationLockType <> "" Then
        Select Case LCase(sValidationLockType)
        Case "warning"
            LetValidationLockType = 2
        Case "error"
            LetValidationLockType = 1
        End Select
    End If
    
End Function

'adding a validation list in an excel range
Sub LetValidationList(oRange As Range, sValidList As String, sAlertType As Byte, sMessage As String)
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

'Setting the min and the max validation in
Sub BuildValidationMinMax(oRange As Range, iMin As String, iMax As String, iAlertType As Byte, sTypeValidation As String, sMessage As String)

    With oRange.Validation
        .Delete
        Select Case LCase(sTypeValidation)
        Case "integer"                           'numerique
            Select Case iAlertType
            Case 1                               '"error"
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case 2                               '"warning"
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case Else
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            End Select
        Case "date"                              'date
            Select Case iAlertType
            Case 1                               '"error"
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case 2                               '"warning"
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            Case Else
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
            End Select
        Case Else                                'decimal
            If InStr(1, LCase(sTypeValidation), "decimal") > 0 Then
                Select Case iAlertType
                Case 1                           '"error"
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
                Case 2                           '"warning"
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
                Case Else
                    .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=iMin, Formula2:=iMax
                End Select
            End If
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

'Write the borders for one range
Sub WriteBorderLines(oRange As Range)

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

Sub AddCmd(xlsapp As Excel.Application, sSheet As String, iLeft As Integer, iTop As Integer, sName As String, sText As String, iCmdWidth As Integer, iCmdHeight As Integer)

    With xlsapp.Sheets(sSheet)
        .Shapes.AddShape(msoShapeRectangle, iLeft + 3, iTop + 3, iCmdWidth, iCmdHeight).Name = sName
        .Shapes(sName).Placement = xlFreeFloating
        .Shapes(sName).TextFrame2.TextRange.Characters.Text = sText
        .Shapes(sName).TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Shapes(sName).TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Shapes(sName).TextFrame2.WordWrap = msoFalse
        .Shapes(sName).TextFrame2.TextRange.Font.Size = 9
        .Shapes(sName).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbBlack
    End With

End Sub

'The purpose of this procedure is to create the geo columns using the geo data  (its also adds the first dropdowns)
' we shift the columns to the right until we reached the number of columns required
Sub Add4GeoCol(xlsapp As Excel.Application, sSheetName As String, sLib As String, sNameCell As String, iCol As Integer, sMessage As String)

    'sSheetName: Sheet name
    'sNameCell: Name of the cell
    'iCol: Column to start shifting
    'sMessage: message in case of error
    'sLib: header message
    
    Dim i As Byte
    Dim j As Byte
    Dim sTemp As String

    With xlsapp.Sheets(sSheetName)
        i = 4
        While i > 1
            .Columns(iCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            .Cells(C_TitleLine, iCol + 1).value = LetWordingWithSpace(xlsapp, Sheets("GEO").ListObjects("T_ADM" & i).HeaderRowRange.Item(i).value, CStr(sSheetName))
            .Cells(C_TitleLine, iCol + 1).Name = "adm" & i & "_" & sNameCell
            .Cells(C_TitleLine, iCol + 1).Interior.Color = vbWhite
            .Cells(C_TitleLine, iCol + 1).Locked = False
            i = i - 1
        Wend
        .Cells(C_TitleLine, iCol).value = LetWordingWithSpace(xlsapp, Sheets("GEO").ListObjects("T_ADM" & i).HeaderRowRange.Item(1).value, CStr(sSheetName))
        .Range(.Cells(C_StartLineTitle2, iCol), .Cells(C_StartLineTitle2, iCol + 3)).Merge
    
        'ajout des formules de validation
        .Cells(C_TitleLine + 1, iCol).Validation.Delete

        .Cells(C_TitleLine + 1, iCol).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, _
                                                     Formula1:="=GEO!" & xlsapp.Sheets("GEO").Range("T_ADM1").Columns(1).Address
     
        .Cells(C_TitleLine + 1, iCol).Validation.IgnoreBlank = True
        .Cells(C_TitleLine + 1, iCol).Validation.InCellDropdown = True
        .Cells(C_TitleLine + 1, iCol).Validation.InputTitle = ""
        .Cells(C_TitleLine + 1, iCol).Validation.errorTitle = ""
        .Cells(C_TitleLine + 1, iCol).Validation.InputMessage = ""
        .Cells(C_TitleLine + 1, iCol).Validation.ErrorMessage = sMessage
        .Cells(C_TitleLine + 1, iCol).Validation.ShowInput = True
        .Cells(C_TitleLine + 1, iCol).Validation.ShowError = True
    End With

End Sub

Private Function LetColNumberByDataName(xlsapp As Excel.Application, sDataName As String, sSheetName As String) As Integer

    Dim i As Integer

    With xlsapp
        i = 1
        While i <= .Sheets(sSheetName).Cells(C_TitleLine, 1).End(xlToRight).Column And UCase(.Sheets(sSheetName).Cells(C_TitleLine, i).Name.Name) <> sDataName
            i = i + 1
        Wend
        If UCase(.Sheets(sSheetName).Cells(C_TitleLine, i).Name.Name) = sDataName Then
            LetColNumberByDataName = i
        End If
    End With

End Function

Private Function LetWordingWithSpace(xlsapp As Excel.Application, sDataWording As String, sSheetName As String)
    'The goal of this function is to add space to duplicates labels so that excels does not force a unique name with number at the end in
    'the listcolumn
    Dim i As Integer

    LetWordingWithSpace = ""
    With xlsapp
        i = 1
        While i <= .Sheets(sSheetName).Cells(C_TitleLine, 1).End(xlToRight).Column And Replace(UCase(.Sheets(sSheetName).Cells(C_TitleLine, i).value), " ", "") <> Replace(UCase(sDataWording), " ", "")
            i = i + 1
        Wend
        
        If Replace(UCase(xlsapp.Sheets(sSheetName).Cells(C_TitleLine, i).value), " ", "") = Replace(UCase(sDataWording), " ", "") Then
            LetWordingWithSpace = xlsapp.Sheets(sSheetName).Cells(C_TitleLine, i).value & " "
        Else
            LetWordingWithSpace = sDataWording
        End If
    End With

End Function



