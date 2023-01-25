Attribute VB_Name = "LinelistEvents"

Option Explicit
Option Private Module

Public iGeoType As Byte
Public DebugMode As Boolean

Sub ClicCmdGeoApp()

    Dim targetColumn As Integer
    Dim sType As String

    targetColumn = ActiveCell.Column

    If ActiveCell.Row > C_eStartLinesLLData + 1 Then

        sType = ActiveSheet.Cells(C_eStartLinesLLMainSec - 1, targetColumn).Value
        Select Case sType
        Case "geo1"
            iGeoType = 0
            LoadGeo 0

        Case "hf"
            iGeoType = 1
            LoadGeo 1
            
        Case Else
            MsgBox TranslateLLMsg("MSG_WrongCells")
        End Select
    Else
        MsgBox TranslateLLMsg("MSG_WrongCells"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
    End If
End Sub

Sub ClicCmdAddRows()

    Dim oLstobj As Object
    Dim iLastRow As Long
    Dim iLastCol As Long
    Dim LoRng As Range

    On Error GoTo errAddRows

    ActiveSheet.UnProtect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value)
    Application.EnableEvents = False
    Set oLstobj = ActiveSheet.ListObjects(SheetListObjectName(ActiveSheet.Name))

    If Not oLstobj.DataBodyRange Is Nothing Then
        iLastRow = oLstobj.DataBodyRange.Rows.Count + C_eStartLinesLLData + 1 + C_iNbLinesLLData
    Else
        iLastRow = FindLastRow(ActiveSheet) + C_iNbLinesLLData
    End If

    iLastCol = 1

    Do While ActiveSheet.Cells(C_eStartLinesLLData + 1, iLastCol).Value <> vbNullString
        iLastCol = iLastCol + 1
    Loop

    iLastCol = iLastCol - 1

    Set LoRng = Range(Cells(C_eStartLinesLLData + 1, 1), Cells(iLastRow, iLastCol))
    oLstobj.Resize LoRng

    Call ProtectSheet
    Application.EnableEvents = True
    Exit Sub

errAddRows:
    Application.EnableEvents = True
    MsgBox TranslateLLMsg("MSG_ErrAddRows"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
    Exit Sub
End Sub

Sub ClicCmdExport()

    Dim i As Byte
    Dim iHeight As Integer
    Dim Wksh As Worksheet
    Dim iStatus As Byte
    Dim iLabel As Byte
    Dim ExportHeaders As BetterArray
    Const C_CmdHeight As Integer = 40
    Const C_CmdGap As Byte = 10

    Set Wksh = ThisWorkbook.Worksheets(C_sParamSheetExport)
    Set ExportHeaders = GetHeaders(ThisWorkbook, C_sParamSheetExport, 1)
    ExportHeaders.LowerBound = 1
    iStatus = ExportHeaders.IndexOf(C_sExportHeaderStatus)
    iLabel = ExportHeaders.IndexOf(C_sExportHeaderLabelButton)

    iHeight = C_CmdGap

    On Error GoTo errLoadExp

    With F_Export
        i = 1
        Do While i <= 5
            If Not isError(Wksh.Cells(i, iStatus).Value) Then
                'i+1 because the first line is for the headers
                If Wksh.Cells(i + 1, iStatus).Value <> C_sExportActive Then
                    .Controls("CMD_Export" & i).Visible = False
                Else
                    .Controls("CMD_Export" & i).Visible = True
                    .Controls("CMD_Export" & i).Caption = Wksh.Cells(i + 1, iLabel).Value
                    .Controls("CMD_Export" & i).Top = iHeight
                    .Controls("CMD_Export" & i).height = C_CmdHeight
                    .Controls("CMD_Export" & i).width = 160
                    .Controls("CMD_Export" & i).Left = 20
                    iHeight = iHeight + C_CmdHeight + C_CmdGap
                End If
            End If
            i = i + 1
        Loop

        'Height of checks (use filtered data)
        .CHK_ExportFiltered.Top = iHeight + 30
        .CHK_ExportFiltered.Left = 30
        .CHK_ExportFiltered.width = 160

        iHeight = iHeight + 40 + C_CmdHeight + C_CmdGap

        'Height of command for new key
        .CMD_NouvCle.Top = iHeight
        .CMD_NouvCle.height = C_CmdHeight - 10
        .CMD_NouvCle.width = 160
        .CMD_NouvCle.Left = 20

        iHeight = iHeight + C_CmdHeight + C_CmdGap

        'Quit command
        .CMD_Retour.Top = iHeight
        .CMD_Retour.height = C_CmdHeight - 10
        .CMD_Retour.width = 160
        .CMD_Retour.Left = 20

        iHeight = iHeight + C_CmdHeight + C_CmdGap

        'Overall height and width of the form

        .height = iHeight + 50
        .width = 210
    End With


    F_Export.Show
    Exit Sub

errLoadExp:
    MsgBox TranslateLLMsg("MSG_ErrLoadExport"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
    EndWork xlsapp:=Application
    Exit Sub
End Sub

Sub ClicCmdDebug()

    Dim pwd As String
    Dim sh As Worksheet
    Dim SheetsToProtect As BetterArray
    Dim DebugWksh As Worksheet

    BeginWork xlsapp:=Application

    On Error GoTo errDebug
    sPrevSheetName = vbNullString

    Set DebugWksh = Worksheets(ActiveSheet.Name)


    'Unprotect All Sheets
    If Not DebugMode Then
        pwd = InputBox(TranslateLLMsg("MSG_ProvidePassword"), TranslateLLMsg("MSG_DebugMode"), "1234")
        If pwd = ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value Then
            For Each sh In ThisWorkbook.Worksheets
                If sh.ProtectContents = True Then
                    sh.UnProtect pwd
                End If
            Next
            DebugMode = True
            DebugWksh.Shapes(C_sShpDebug).Fill.ForeColor.RGB = Helpers.GetColor("Green")
            DebugWksh.Shapes(C_sShpDebug).Fill.BackColor.RGB = Helpers.GetColor("Green")
            DebugWksh.Shapes(C_sShpDebug).TextFrame2.TextRange.Characters.text = TranslateLLMsg("MSG_Protect")
            DebugWksh.Select
        Else
            MsgBox TranslateLLMsg("MSG_WrongPassword"), vbOK, "DEBUG MODE"
        End If
    Else
        pwd = ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value
        Set SheetsToProtect = New BetterArray
        Set SheetsToProtect = GetDictionaryColumn(C_sDictHeaderSheetName)

        For Each sh In ThisWorkbook.Worksheets
            If SheetsToProtect.Includes(sh.Name) Then
                sh.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                           AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                           AllowFormattingColumns:=True


            End If
        Next

        'Debug Mode is False
        DebugMode = False
        ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value = pwd
        DebugWksh.Shapes(C_sShpDebug).Fill.ForeColor.RGB = Helpers.GetColor("Orange")
        DebugWksh.Shapes(C_sShpDebug).Fill.BackColor.RGB = Helpers.GetColor("Orange")
        DebugWksh.Shapes(C_sShpDebug).TextFrame2.TextRange.Characters.text = TranslateLLMsg("MSG_Debug")
    End If

    Exit Sub

errDebug:
    MsgBox TranslateLLMsg("MSG_ErrDebug"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
    EndWork xlsapp:=Application
    Exit Sub

    EndWork xlsapp:=Application
End Sub

'Protect sheet of type linelist
Public Sub ProtectSheet(Optional sSheetName As String = "_Active")
    Dim pwd As String
    Dim sh As Worksheet

    If sSheetName = "_Active" Then
        Set sh = ActiveSheet
    Else
        Set sh = ThisWorkbook.Worksheets(sSheetName)
    End If

    If Not DebugMode Then
        pwd = ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value
        sh.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                   AllowFormattingColumns:=True
    End If

End Sub

'Trigerring event when the linelist sheet has some values within                                                          -                                                      -
Sub EventValueChangeLinelist(Target As Range)

    Const GOTOSECCODE As String = "go_to_section" 'Go To section constant
    
    Dim T_geo As BetterArray
    Set T_geo = New BetterArray
    Dim varControl As String                   'Control type
    Dim sLabel As String
    Dim varName As String
    Dim varSubLabel As String
    Dim targetColumn As Long 'column of the target range
    Dim rng As Range
    Dim loAdm2 As ListObject
    Dim loAdm3 As ListObject
    Dim loAdm4 As ListObject
    Dim tableName As String
    Dim adminNames As BetterArray
    Dim sh As Worksheet 'Active sheet where the event fires
    Dim geo As ILLGeo
    Dim cellRng As Range
    Dim hRng As Range 'Header Row Range of the listObject
    Dim goToSection As String
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim startLine As Long
    Dim calcRng As Range 'calculate range
    Dim nbOffset As Long 'number of offset from the headerrow range
    
    On Error GoTo errHand
    Set sh = ActiveSheet
    tableName = sh.Cells(1, 4).Value
    Set rng = sh.Range(tableName & "_" & GOTOSECCODE)
    Set geo = LLGeo.Create(ThisWorkbook.Worksheets("Geo"))
    Set hRng = sh.ListObjects(1).HeaderRowRange
    Set adminNames = New BetterArray
    adminNames.LowerBound = 1

    targetColumn = Target.Column
    startLine = sh.Range(tableName & "_START").Row
    varControl = sh.Cells(startLine - 5, targetColumn).Value

    If Target.Row >= startLine Then
        
        nbOffset = Target.Row - hRng.Row
        Set calcRng = hRng.Offset(nbOffset)
        calcRng.Calculate

        If (varControl = "geo1") Or (varControl = "geo2") Or (varControl = "geo3") Or (varControl = "geo4") Then

            Set loAdm2 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin2")
            Set loAdm3 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin3")
            Set loAdm4 = ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects("list_admin4")


            Select Case varControl

            Case "geo1"
                'adm1 has been modified, we will correct and set validation to adm2

                BeginWork xlsapp:=Application

                DeleteLoDataBodyRange loAdm2
                Target.Offset(, 1).Value = vbNullString
                DeleteLoDataBodyRange loAdm3
                Target.Offset(, 2).Value = vbNullString
                DeleteLoDataBodyRange loAdm4
                Target.Offset(, 3).Value = vbNullString

                If Target.Value <> vbNullString Then

                    'Filter on adm1
                    Set T_geo = geo.GeoLevel(LevelAdmin2, CustomTypeGeo, Target.Value)
                    'Build the validation list for adm2
                    T_geo.ToExcelRange loAdm2.Range.Cells(2, 1)
                    T_geo.Clear
                End If


                EndWork xlsapp:=Application

            Case "geo2"

                'Adm2 has been modified, we will correct and filter adm3
                BeginWork xlsapp:=Application

                DeleteLoDataBodyRange loAdm3
                Target.Offset(, 1).Value = vbNullString
                DeleteLoDataBodyRange loAdm4
                Target.Offset(, 2).Value = vbNullString

                If Target.Value <> vbNullString Then
                    adminNames.Push Target.Offset(, -1).Value, Target.Value
                    Set T_geo = geo.GeoLevel(LevelAdmin3, CustomTypeGeo, adminNames)
                    T_geo.ToExcelRange loAdm3.Range.Cells(2, 1)
                    T_geo.Clear
                End If

                EndWork xlsapp:=Application

            Case "geo3"
                'Adm 3 has been modified, correct and filter adm4
                BeginWork xlsapp:=Application

                DeleteLoDataBodyRange loAdm4
                Target.Offset(, 1).Value = vbNullString

                If Target.Value <> vbNullString Then

                    adminNames.Push Target.Offset(, -2).Value, Target.Offset(, -1).Value, Target.Value
                    'Take the adm4 table
                    Set T_geo = geo.GeoLevel(LevelAdmin4, CustomTypeGeo, adminNames)
                    T_geo.ToExcelRange loAdm4.Range.Cells(2, 1)
                    T_geo.Clear
                End If

                EndWork xlsapp:=Application

            End Select
        End If

    End If
    
    'Update the custom control
    If (Target.Row = startLine - 2) And (varControl = "custom") Then
        Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("Dictionary"), 1, 1)
        Set vars = LLVariables.Create(dict)
        'The name of custom variables has been updated, update the dictionary
        varName = sh.Cells(startLine - 1, targetColumn).Value
        varSubLabel = vars.Value(varName:=varName, colName:="sub label")

        sLabel = Replace(Target.Value, varSubLabel, "")
        sLabel = Replace(sLabel, Chr(10), "")

        vars.SetValue varName:=varName, colName:="main label", newValue:=sLabel

    End If
    
    'Update the list auto
    If Target.Row >= startLine And _
       sh.Cells(startLine - 6, targetColumn).Value = "list_auto_origin" And _
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value <> "list_auto_change_yes" Then
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value = "list_auto_change_yes"
    End If

    
    'GoTo section
    If Not Intersect(Target, rng) Is Nothing Then
        goToSection = ThisWorkbook.Worksheets("LinelistTranslation").Range("RNG_GoToSection").Value

        sLabel = Replace(Target.Value, goToSection & ": ", "")
        Set hRng = sh.ListObjects(1).HeaderRowRange
        Set hRng = hRng.Offset(-3)

        Set cellRng = hRng.Find(What:=sLabel, LookAt:=xlWhole, MatchCase:=True)

        If Not cellRng Is Nothing Then cellRng.Activate
    End If

    If Target.Row = startLine - 1 Then
        Target.Value = Target.Offset(-1).Name.Name
        MsgBox "Do not modify the Headers!!!!!"
    End If

errHand:

End Sub

Sub ClicCmdAdvanced()
    'Import exported data into the linelist
    F_Advanced.Show
End Sub

Sub ClicExportMigration()

    Static AfterFirstClicMig As Boolean

    If AfterFirstClicMig Then
        [F_ExportMig].Show
    Else
        'For the first click Thick Migration and Geo and put historic to false
        'For subsequent clicks, just show what have been ticked
        [F_ExportMig].CHK_ExportMigData.Value = True
        [F_ExportMig].CHK_ExportMigGeo.Value = True
        [F_ExportMig].CHK_ExportMigGeoHistoric.Value = True
        [F_ExportMig].Show
        AfterFirstClicMig = True
    End If
End Sub

'Event to update the list_auto when a sheet containing a list_auto is desactivated
Public Sub EventDesactivateLinelist(ByVal sSheetName As String)

    Dim PrevWksh As Worksheet

    On Error GoTo errHand

    If ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value = "list_auto_change_yes" Then

        Set PrevWksh = ThisWorkbook.Worksheets(sSheetName)
        BeginWork xlsapp:=Application

        UpdateListAuto PrevWksh
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).Value = "list_auto_change_no"

        EndWork xlsapp:=Application
        Exit Sub

    End If
errHand:
    EndWork xlsapp:=Application
End Sub

'Update the list Auto of one Sheet

Public Sub UpdateListAuto(Wksh As Worksheet)

    Dim iChoiceCol As Integer
    Dim choiceLo As ListObject
    Dim sVarName As String
    Dim iRow As Long
    Dim i As Long
    Dim arrTable As BetterArray
    Dim listAutoSheet As Worksheet

    Dim rng As Range

    Set arrTable = New BetterArray
    i = 1

    Set listAutoSheet = ThisWorkbook.Worksheets(C_sSheetChoiceAuto)
    With Wksh
        .Calculate
        Do While (.Cells(C_eStartLinesLLData, i) <> vbNullString)
            Select Case .Cells(C_eStartLinesLLMainSec - 2, i).Value
            Case C_sDictControlChoiceAuto & "_origin"
                sVarName = .Cells(C_eStartLinesLLData + 1, i).Value
                If ListObjectExists(listAutoSheet, "list_" & sVarName) Then
                    arrTable.FromExcelRange .Cells(C_eStartLinesLLData + 2, i), DetectLastColumn:=False, DetectLastRow:=True
                    'Unique values (removing the spaces and the Null strings and keeping the case (The remove duplicates doesn't do that))
                    Set arrTable = GetUniqueBA(arrTable)
                    With listAutoSheet
                        Set choiceLo = .ListObjects("list_" & sVarName)
                        iChoiceCol = choiceLo.Range.Column
                        If Not choiceLo.DataBodyRange Is Nothing Then choiceLo.DataBodyRange.Delete
                        arrTable.ToExcelRange .Cells(C_eStartlinesListAuto + 1, iChoiceCol)
                        iRow = .Cells(Rows.Count, iChoiceCol).End(xlUp).Row
                        choiceLo.Resize .Range(.Cells(C_eStartlinesListAuto, iChoiceCol), .Cells(iRow, iChoiceCol))
                        'Sort in descending order
                        Set rng = choiceLo.ListColumns(1).Range
                        With choiceLo.Sort
                            .SortFields.Clear
                            .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, ORDER:=xlDescending
                            .Header = xlYes
                            .Apply
                        End With
                    End With
                End If
            Case Else
            End Select
            i = i + 1
        Loop
    End With

End Sub

'Update data on Filtered values ===================================================================================


Public Sub UpdateFilterTables(Optional Byval calculate As Boolean = True)

    Dim Wksh As Worksheet                        'The actual worksheet
    Dim filtWksh As Worksheet                    'Filtered worksheet
    Dim Lo As ListObject
    Dim rowCounter As Long
    Dim endCol As Long
    Dim destRng As Range
    Dim delRng As Range
    
    On Error GoTo ErrUpdate
    BeginWork xlsapp:=Application

    For Each Wksh In ThisWorkbook.Worksheets
        If Wksh.Cells(1, 3).Value = "HList" Then

            'Unprotect the worksheet
            With Wksh

                'Clean the filtered table list object
                Set Lo = .ListObjects(1)
                endCol = Lo.Range.Columns.Count

                If Not Lo.DataBodyRange Is Nothing Then
                    Set filtWksh = ThisWorkbook.Worksheets(.Cells(1, 5).Value)

                    On Error Resume Next
                        filtWksh.ListObjects(1).DataBodyRange.Delete
                    On Error GoTo ErrUpdate

                    rowCounter = C_eStartLinesLLData + 1 + Lo.DataBodyRange.Rows.Count


                    With filtWksh
                        Set destRng = .Range(.Cells(C_eStartLinesLLData + 1, 1), .Cells(rowCounter, endCol))
                        .ListObjects(1).Resize destRng
                        Set destRng = .Range(.Cells(C_eStartLinesLLData + 2, 1), .Cells(rowCounter, endCol))
                    End With

                    destRng.Value = Lo.DataBodyRange.Value

                    Do While rowCounter > C_eStartLinesLLData + 1

                        If .Rows(rowCounter).Hidden Then
                            With filtWksh
                                If delRng Is Nothing Then
                                    Set delRng = .Range(.Cells(rowCounter, 1), .Cells(rowCounter, endCol))
                                Else
                                    Set delRng = Application.Union(delRng, .Range(.Cells(rowCounter, 1), .Cells(rowCounter, endCol)))
                                End If
                            End With
                        End If
                        rowCounter = rowCounter - 1
                    Loop

                    'Delete the ragne if necessary
                    If Not delRng Is Nothing Then delRng.Delete

                End If
            End With
        End If
    Next

    'caclulate active sheet
    If calculate Then ActiveSheet.Calculate

    EndWork xlsapp:=Application
    Exit Sub

ErrUpdate:
    MsgBox TranslateLLMsg("MSG_ErrUpdate"), vbCritical + vbOKOnly
    EndWork xlsapp:=Application
End Sub

Sub UpdateSpTables()
    Const SPATIALSHEET As String = "spatial_tables__"
    Dim sp As ILLSpatial
    Dim sh As Worksheet
    
    Set sh = ThisWorkbook.Worksheets(SPATIALSHEET)
    Set sp = LLSpatial.Create(sh)

    UpdateFilterTables calculate := False
    sp.Update
End Sub



'Clear All the filters on current sheet =============================================================================

Sub ClearAllFilters()
    Dim Wksh As Worksheet
    Set Wksh = ActiveSheet

    'On Error GoTo errHand

    If Not Wksh.ListObjects(1).AutoFilter Is Nothing Then

        BeginWork xlsapp:=Application

        'Unprotect current worksheet
        Wksh.UnProtect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).Value)
        'remove the filters
        Wksh.ListObjects(1).AutoFilter.ShowAllData
        ProtectSheet Wksh.Name

        EndWork xlsapp:=Application

    End If

    Exit Sub
errHand:
    EndWork xlsapp:=Application
End Sub

'Find the selected column on "GOTO" Area and go to that column
Sub EventValueChangeAnalysis(Target As Range)

    Dim rng As Range
    Dim RngLook As Range
    Dim sLabel As String
    Dim actSh As Worksheet
    Dim analysisType As String
    Dim goToSection As String
    Dim goToHeader As String
    Dim rngName As String


    'Range name if it exists
    On Error Resume Next
        rngName = Target.Name.Name
    On Error GoTo 0

    On Error GoTo Err
    Set actSh = ActiveSheet
    
    analysisType = actSh.Cells(1, 3).Value
    
    Select Case analysisType

    Case "Uni-Bi-Analysis"
        'GoTo section range for univariate and bivariate analysis
        Set rng = actSh.Range("ua_go_to_section")

    Case "TS-Analysis"
        actSh.Calculate
        'Goto section range for time series analysis
        If InStr(rngName, "ts_go_to_section", 1) > 0 Then Set rng = Target
        
    Case "SP-Analysis"
        'GoTo section for spatial analysis
        Set rng = actSh.Range("sp_go_to_section")
        If InStr(rngName, "ADM_DROPDOWN_") Then UpdateSingleTable 


    End Select

    If (Not (Intersect(Target, rng) Is Nothing)) And (Not rng Is Nothing) Then
        goToSection = ThisWorkbook.Worksheets("LinelistTranslation").Range("RNG_GoToSection").Value
        goToHeader = ThisWorkbook.Worksheets("LinelistTranslation").Range("RNG_GoToHeader").Value
        
        sLabel = Replace(Target.Value, goToSection & ": ", "")
        sLabel = Replace(sLabel, goToHeader & ": ", "")

        Set RngLook = ActiveSheet.Cells.Find(What:=sLabel, LookIn:=xlValues, LookAt:=xlWhole, _
                                             MatchCase:=True, SearchFormat:=False)

        If Not RngLook Is Nothing Then RngLook.Activate
    End If


    Exit Sub
Err:
End Sub

Sub EventValueChangeVList(Target As Range)

    Const GOTOSECCODE As String = "go_to_section" 'Go To section constant

    Dim rng As Range
    Dim RngLook As Range
    Dim sLabel As String
    Dim sh As Worksheet
    Dim tableName As String
    Dim goToSection As String
    
    
    On Error GoTo Err
    Set sh = ActiveSheet
    tableName = sh.Cells(1, 4).Value

    'Calculate the range where the values are entered
    Set rng = sh.Range(tableName & "_" & "PLAGEVALUES")
    rng.Calculate
    
    Set rng = sh.Range(tableName & "_" & GOTOSECCODE)
    goToSection = ThisWorkbook.Worksheets("LinelistTranslation").Range("RNG_GoToSection").Value
    
    If Not Intersect(Target, rng) Is Nothing Then
        sLabel = Replace(Target.Value, goToSection & ": ", "")
        Set RngLook = sh.Cells.Find(What:=sLabel, LookAt:=xlWhole, MatchCase:=True)
        If Not RngLook Is Nothing Then RngLook.Activate
    End If

    Exit Sub
Err:
End Sub


