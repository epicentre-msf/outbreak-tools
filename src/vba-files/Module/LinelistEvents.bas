Attribute VB_Name = "LinelistEvents"

Option Explicit
Option Private Module

Public iGeoType As Byte
Dim DebugMode As Boolean

Sub ClicCmdGeoApp()

    Dim iNumCol As Integer
    Dim sType As String

    iNumCol = ActiveCell.Column

    If ActiveCell.Row > C_eStartLinesLLData + 1 Then

        sType = ActiveSheet.Cells(C_eStartLinesLLMainSec - 1, iNumCol).value
        Select Case sType
        Case C_sDictControlGeo
            iGeoType = 0
            Call LoadGeo(iGeoType)

        Case C_sDictControlHf
            iGeoType = 1
            Call LoadGeo(iGeoType)

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

    ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)
    Application.EnableEvents = False
    Set oLstobj = ActiveSheet.ListObjects(SheetListObjectName(ActiveSheet.Name))

    If Not oLstobj.DataBodyRange Is Nothing Then
        iLastRow = oLstobj.DataBodyRange.Rows.Count + C_eStartLinesLLData + 1 + C_iNbLinesLLData
    Else
        iLastRow = FindLastRow(ActiveSheet) + C_iNbLinesLLData
    End If

    iLastCol = 1

    Do While ActiveSheet.Cells(C_eStartLinesLLData + 1, iLastCol).value <> vbNullString
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
            If Not isError(Wksh.Cells(i, iStatus).value) Then
                'i+1 because the first line is for the headers
                If Wksh.Cells(i + 1, iStatus).value <> C_sExportActive Then
                    .Controls("CMD_Export" & i).Visible = False
                Else
                    .Controls("CMD_Export" & i).Visible = True
                    .Controls("CMD_Export" & i).Caption = Wksh.Cells(i + 1, iLabel).value
                    .Controls("CMD_Export" & i).Top = iHeight
                    .Controls("CMD_Export" & i).Height = C_CmdHeight
                    .Controls("CMD_Export" & i).Width = 160
                    .Controls("CMD_Export" & i).Left = 20
                    iHeight = iHeight + C_CmdHeight + C_CmdGap
                End If
            End If
            i = i + 1
        Loop

        'Height of checks (use filtered data)
        .CHK_ExportFiltered.Top = iHeight + 30
        .CHK_ExportFiltered.Left = 30
        .CHK_ExportFiltered.Width = 160

        iHeight = iHeight + 40 + C_CmdHeight + C_CmdGap

        'Height of command for new key
        .CMD_NouvCle.Top = iHeight
        .CMD_NouvCle.Height = C_CmdHeight - 10
        .CMD_NouvCle.Width = 160
        .CMD_NouvCle.Left = 20

        iHeight = iHeight + C_CmdHeight + C_CmdGap

        'Quit command
        .CMD_Retour.Top = iHeight
        .CMD_Retour.Height = C_CmdHeight - 10
        .CMD_Retour.Width = 160
        .CMD_Retour.Left = 20

        iHeight = iHeight + C_CmdHeight + C_CmdGap

        'Overall height and width of the form

        .Height = iHeight + 50
        .Width = 210
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
        If pwd = ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value Then
            For Each sh In ThisWorkbook.Worksheets
                If sh.ProtectContents = True Then
                    sh.Unprotect pwd
                End If
            Next
            DebugMode = True
            DebugWksh.Shapes(C_sShpDebug).Fill.ForeColor.RGB = Helpers.GetColor("Green")
            DebugWksh.Shapes(C_sShpDebug).Fill.BackColor.RGB = Helpers.GetColor("Green")
            DebugWksh.Shapes(C_sShpDebug).TextFrame2.TextRange.Characters.Text = TranslateLLMsg("MSG_Protect")
            DebugWksh.Select
        Else
            MsgBox TranslateLLMsg("MSG_WrongPassword"), vbOK, "DEBUG MODE"
        End If
    Else
        pwd = ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value
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
        ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value = pwd
        DebugWksh.Shapes(C_sShpDebug).Fill.ForeColor.RGB = Helpers.GetColor("Orange")
        DebugWksh.Shapes(C_sShpDebug).Fill.BackColor.RGB = Helpers.GetColor("Orange")
        DebugWksh.Shapes(C_sShpDebug).TextFrame2.TextRange.Characters.Text = TranslateLLMsg("MSG_Debug")
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
        pwd = ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value
        sh.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                   AllowFormattingColumns:=True
    End If

End Sub

'Trigerring event when the linelist sheet has some values within                                                          -
'Trigerring event when the linelist sheet has some values within                                                          -
Sub EventValueChangeLinelist(oRange As Range)

    Dim T_geo As BetterArray
    Set T_geo = New BetterArray
    Dim sControlType As String                   'Control type
    Dim sLabel As String
    Dim sCustomVarName As String
    Dim sNote As String
    Dim sListAutoType As String
    Dim iNumCol As Integer
    Dim Rng As Range

    On Error GoTo errHand
    iNumCol = oRange.Column
    sControlType = ActiveSheet.Cells(C_eStartLinesLLMainSec - 1, iNumCol).value

    If oRange.Row > C_eStartLinesLLData + 1 Then

        Select Case sControlType

        Case C_sDictControlGeo
            ' adm1 has been modified, we will correct and set validation to adm2

            BeginWork xlsapp:=Application
            ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects(C_sTabAdm4 & "_dropdown")
            oRange.Offset(, 1).value = vbNullString
            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects(C_sTabAdm3 & "_dropdown")
            oRange.Offset(, 2).value = vbNullString
            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects(C_sTabAdm2 & "_dropdown")
            oRange.Offset(, 3).value = vbNullString

            If oRange.value <> vbNullString Then

                'Filter on adm1
                Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabAdm2), 1, oRange.value, returnIndex:=2)
                'Build the validation list for adm2
                T_geo.ToExcelRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).Cells(2, 1)
                T_geo.Clear
            End If

            Call ProtectSheet
            EndWork xlsapp:=Application

        Case C_sDictControlGeo & "2"

            'Adm2 has been modified, we will correct and filter adm3
            BeginWork xlsapp:=Application
            ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects(C_sTabAdm3 & "_dropdown")
            oRange.Offset(, 1).value = vbNullString
            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects(C_sTabAdm4 & "_dropdown")
            oRange.Offset(, 2).value = vbNullString

            If oRange.value <> vbNullString Then
                Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabAdm3), 1, oRange.Offset(, -1).value, 2, oRange.value, returnIndex:=3)
                T_geo.ToExcelRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).Cells(2, 3)
                T_geo.Clear
            End If

            Call ProtectSheet
            EndWork xlsapp:=Application

        Case C_sDictControlGeo & "3"
            'Adm 3 has been modified, correct and filter adm4
            BeginWork xlsapp:=Application
            ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

            DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).ListObjects(C_sTabAdm4 & "_dropdown")
            oRange.Offset(, 1).value = vbNullString

            If oRange.value <> vbNullString Then
                'Take the adm4 table
                Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabAdm4), 1, _
                                          oRange.Offset(, -2).value, 2, oRange.Offset(, -1).value, 3, oRange.value, returnIndex:=4)
                T_geo.ToExcelRange ThisWorkbook.Worksheets(C_sSheetChoiceAuto).Cells(2, 5)
                T_geo.Clear
            End If

            Call ProtectSheet
            EndWork xlsapp:=Application

        Case Else

        End Select

    End If

    If oRange.Row = C_eStartLinesLLData And sControlType = C_sDictControlCustom Then
        'The name of custom variables has been updated, update the dictionary
        sCustomVarName = ActiveSheet.Cells(C_eStartLinesLLData + 1, iNumCol).value
        sNote = GetDictColumnValue(sCustomVarName, C_sDictHeaderSubLab)
        sLabel = Replace(oRange.value, sNote, "")
        sLabel = Replace(sLabel, Chr(10), "")

        Call UpdateDictionaryValue(sCustomVarName, C_sDictHeaderMainLab, sLabel)

    End If


    If oRange.Row > C_eStartLinesLLData + 1 And _
       ActiveSheet.Cells(C_eStartLinesLLMainSec - 2, iNumCol).value = C_sDictControlChoiceAuto & "_origin" And _
       ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).value <> "list_auto_change_yes" Then
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).value = "list_auto_change_yes"
    End If


    If oRange.Name.Name = SheetListObjectName(ActiveSheet.Name) & "_" & C_sGotoSection Then
        sLabel = Replace(oRange.value, TranslateLLMsg("MSG_SelectSection") & ": ", "")

        Set Rng = ActiveSheet.Rows(C_eStartLinesLLMainSec).Find(What:=sLabel, _
                                                                LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
                                                                SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)

        If Not Rng Is Nothing Then
            Rng.Activate
        End If

    End If

errHand:

End Sub

Sub ClicImportMigration()
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
        [F_ExportMig].CHK_ExportMigData.value = True
        [F_ExportMig].CHK_ExportMigGeo.value = True
        [F_ExportMig].CHK_ExportMigGeoHistoric.value = True
        [F_ExportMig].Show
        AfterFirstClicMig = True
    End If
End Sub

'Event to update the list_auto when a sheet containing a list_auto is desactivated
Public Sub EventDesactivateLinelist(ByVal sSheetName As String)

    Dim PrevWksh As Worksheet

    On Error GoTo errHand

    If ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).value = "list_auto_change_yes" Then

        Set PrevWksh = ThisWorkbook.Worksheets(sSheetName)
        BeginWork xlsapp:=Application

        UpdateListAuto PrevWksh
        ThisWorkbook.Worksheets(C_sSheetImportTemp).Cells(1, 15).value = "list_auto_change_no"

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

    Dim Rng As Range

    Set arrTable = New BetterArray
    i = 1

    With Wksh


        Do While (.Cells(C_eStartLinesLLData, i) <> vbNullString)
            Select Case .Cells(C_eStartLinesLLMainSec - 2, i).value
            Case C_sDictControlChoiceAuto & "_origin"
                sVarName = .Cells(C_eStartLinesLLData + 1, i).value
                arrTable.FromExcelRange .Cells(C_eStartLinesLLData + 2, i), DetectLastColumn:=False, DetectLastRow:=True
                'Unique values (removing the spaces and the Null strings and keeping the case (The remove duplicates doesn't do that))
                Set arrTable = GetUniqueBA(arrTable)
                With ThisWorkbook.Worksheets(C_sSheetChoiceAuto)
                    Set choiceLo = .ListObjects("o" & C_sDictControlChoiceAuto & "_" & sVarName)
                    iChoiceCol = choiceLo.Range.Column
                    If Not choiceLo.DataBodyRange Is Nothing Then choiceLo.DataBodyRange.Delete
                    arrTable.ToExcelRange .Cells(C_eStartlinesListAuto + 1, iChoiceCol)
                    iRow = .Cells(Rows.Count, iChoiceCol).End(xlUp).Row
                    choiceLo.Resize .Range(.Cells(C_eStartlinesListAuto, iChoiceCol), .Cells(iRow, iChoiceCol))
                    'Sort in descending order
                    Set Rng = choiceLo.ListColumns(1).Range
                    With choiceLo.Sort
                        .SortFields.Clear
                        .SortFields.Add Key:=Rng, SortOn:=xlSortOnValues, Order:=xlDescending
                        .Header = xlYes
                        .Apply
                    End With
                End With
            Case Else
            End Select
            i = i + 1
        Loop
    End With

End Sub

'Update data on Filtered values ===================================================================================

Public Sub UpdateFilterTables()

    Dim Wksh As Worksheet                        'The actual worksheet
    Dim DictHeaders As BetterArray               'Headers of the dictionary
    Dim LLSheets As BetterArray                  'List of all sheets of type linelist
    Dim Rng As Range
    Dim Lo As ListObject
    Dim HiddenColumns As BetterArray
    Dim i As Long
    Dim sActSh As String

    On Error GoTo ErrUpdate
    BeginWork xlsapp:=Application

    sActSh = ActiveSheet.Name

    Set HiddenColumns = New BetterArray
    Set DictHeaders = GetDictionaryHeaders()


    Set LLSheets = FilterLoTable(Lo:=ThisWorkbook.Worksheets(C_sParamSheetDict).ListObjects(1), _
                                 iFiltindex1:=DictHeaders.IndexOf(C_sDictHeaderSheetType), _
                                 sValue1:=C_sDictSheetTypeLL, _
                                 returnIndex:=DictHeaders.IndexOf(C_sDictHeaderSheetName))


    For Each Wksh In ThisWorkbook.Worksheets
        If LLSheets.Includes(Wksh.Name) Then


            HiddenColumns.Clear
            'Unprotect the worksheet
            With Wksh
                .Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

                'Clean the filtered table list object
                DeleteLoDataBodyRange ThisWorkbook.Worksheets(C_sFiltered & .Name).ListObjects(1)

                'Find Hidden Columns in a worksheets
                i = 1
                Do While .Cells(C_eStartLinesLLData + 1, i).value <> vbNullString
                    If .Columns(i).Hidden Then HiddenColumns.Push i
                    i = i + 1
                Loop

                Set Lo = Wksh.ListObjects(1)
                With Lo.DataBodyRange
                    .EntireColumn.AutoFit
                    Set Rng = .SpecialCells(xlCellTypeVisible)
                End With

                Rng.Copy ThisWorkbook.Worksheets(C_sFiltered & .Name).Cells(C_eStartLinesLLData + 2, 1)

                'Bring back hidden columns
                If HiddenColumns.Length > 0 Then
                    For i = 1 To HiddenColumns.Length
                        .Columns(i).Hidden = True
                    Next
                End If

                'Set Column Width of First and Second Column
                .Columns(1).ColumnWidth = C_iLLFirstColumnsWidth
                .Columns(2).ColumnWidth = C_iLLFirstColumnsWidth

                'Reprotect back the worksheet
                ProtectSheet .Name
            End With
        End If
    Next

  

    On Error Resume Next
    ThisWorkbook.Worksheets(sActSh).Activate
    On Error GoTo 0


    EndWork xlsapp:=Application
    Exit Sub

ErrUpdate:
    MsgBox TranslateLLMsg("MSG_ErrUpdate"), vbCritical + vbOKOnly
    EndWork xlsapp:=Application
End Sub

'Clear All the filters on current sheet =====================================================================

Sub ClearAllFilters()
    Dim Wksh As Worksheet
    Set Wksh = ActiveSheet


    If Not Wksh.AutoFilter Is Nothing Then

        BeginWork xlsapp:=Application

        'Unprotect current worksheet
        Wksh.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)
        'remove the filters
        Wksh.AutoFilter.ShowAllData
        ProtectSheet Wksh.Name

        EndWork xlsapp:=Application

    End If


End Sub

'Find the selected column on "GOTO" Area and go to that column
Sub EventValueChangeAnalysis(Target As Range)

    Dim Rng As Range
    Dim RngLook As Range
    Dim sLabel As String

    On Error GoTo Err
    Call SetUserDefineConstants

    Select Case ActiveSheet.Name

        Case sParamSheetAnalysis

            Set Rng = ThisWorkbook.Worksheets(sParamSheetAnalysis).Range(C_sTabLLUBA & "_" & C_sGotoSection)

        Case sParamSheetTemporalAnalysis

            Set Rng = ThisWorkbook.Worksheets(sParamSheetTemporalAnalysis).Range(C_sTabLLTA & "_" & C_sGotoSection)

        Case sParamSheetSpatialAnalysis
    End Select


    If Not Intersect(Target, Rng) Is Nothing Then
        sLabel = Replace(Target.value, TranslateLLMsg("MSG_SelectSection") & ": ", "")

        Set RngLook = ActiveSheet.Cells.Find(What:=sLabel, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True, SearchFormat:=False)

        If Not RngLook Is Nothing Then RngLook.Activate
    End If


    Exit Sub
Err:
End Sub


