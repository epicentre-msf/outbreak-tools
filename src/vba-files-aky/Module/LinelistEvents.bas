Attribute VB_Name = "LinelistEvents"

Option Explicit

Public iGeoType As Byte

Sub ClicCmdGeoApp()

    Dim iNumCol As Integer
    Dim sType As String

    iNumCol = ActiveCell.Column

    If ActiveCell.Row > C_estartlineslldata + 1 Then

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
        MsgBox TranslateLLMsg("MSG_WrongCells"), vbOKOnly + vbCritical, "ERROR"
    End If
End Sub

Sub ClicCmdAddRows()

    Dim oLstobj As Object

    On Error GoTo errAddRows

    ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)
    Application.EnableEvents = False

    For Each oLstobj In ActiveSheet.ListObjects
        oLstobj.Resize Range(Cells(C_estartlineslldata + 1, 1), Cells(oLstobj.DataBodyRange.Rows.Count + C_iNbLinesLLData + C_estartlineslldata + 1, Cells(C_estartlineslldata + 1, 1).End(xlToRight).Column))
    Next

    Call ProtectSheet
    Application.EnableEvents = True

    Exit Sub

errAddRows:
        Application.EnableEvents = True
        MsgBox TranslateLLMsg("MSG_ErrAddRows"), vbOKOnly + vbCritical, "ERROR"
        Exit Sub
End Sub

Sub ClicCmdExport()

    Dim i As Byte
    Dim iHeight As Integer
    Const C_CmdHeight As Integer = 40

    iHeight = 1

    On Error GoTo errLoadExp

    With F_Export
        i = 2
        While i <= 6
            If Not isError(Sheets("Exports").Cells(i, 4).value) Then
                If LCase(Sheets("Exports").Cells(i, 4).value) <> "active" Then
                    .Controls("CMD_Export" & i - 1).Visible = False
                Else
                    .Controls("CMD_Export" & i - 1).Visible = True
                    .Controls("CMD_Export" & i - 1).Caption = Sheets("Exports").Cells(i, 2).value
                    iHeight = iHeight + 24 + C_CmdHeight
                End If
            End If
            i = i + 1
        Wend
        .CMD_NouvCle.Top = iHeight + 5
        '.CMD_NouvCle.Visible = True
        iHeight = iHeight + 24 + C_iCmdHeight

        .CMD_Retour.Top = iHeight + 5
        '.CMD_Retour.Visible = True
        iHeight = .CMD_Retour.Top + .CMD_Retour.Height + 24 + 10
        .Height = iHeight
        .Width = 200
        .Show
    End With

    Exit Sub

errLoadExp:
        MsgBox TranslateLLMsg("MSG_ErrLoadExport"), vbOKOnly + vbCritical, "ERROR"
        EndWork xlsapp:=Application
        Exit Sub


End Sub


Sub ClicCmdDebug()
    'Debug Mode logic
    Static DebugMode As Boolean
    Dim pwd As String
    Dim sh As Worksheet
    Dim SheetsOfTypeLLData As BetterArray
    Dim DictHeaders As BetterArray
    Dim i As Integer
    Dim iNbVar As Integer
    Dim sPrevSheetName As String
    Dim DebugWksh As Worksheet

    BeginWork xlsapp:=Application

    On Error GoTo errDebug

    pwd = InputBox("Provide the debugging password", "DEBUG MODE", "1234")
    Set DebugWksh = Worksheets(ActiveSheet.Name)


    'Unprotect All Sheets
    If Not DebugMode Then
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
        With ThisWorkbook.Worksheets(C_sParamSheetDict)
        'Protect All Sheets of Type LL
            iNbVar = .Cells(Rows.Count, 1).End(xlUp).Row
            Set DictHeaders = GetDictionaryHeaders()
            Set SheetsOfTypeLLData = New BetterArray
            For i = 1 To iNbVar
                If .Cells(i, DictHeaders.IndexOf(C_sDictHeaderSheetType)) = C_sDictSheetTypeLL And .Cells(i, DictHeaders.IndexOf(C_sDictHeaderSheetName)) <> sPrevSheetName Then
                    sPrevSheetName = .Cells(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
                    SheetsOfTypeLLData.Push sPrevSheetName
                End If
            Next
        End With

        For Each sh In ThisWorkbook.Worksheets
            If SheetsOfTypeLLData.Includes(sh.Name) Then

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
        MsgBox TranslateLLMsg("MSG_ErrDebug"), vbOKOnly + vbCritical, "ERROR"
        EndWork xlsapp:=Application
        Exit Sub

    EndWork xlsapp:=Application
End Sub


'Trigerring event when the linelist sheet has some values within                                                          -
'Trigerring event when the linelist sheet has some values within                                                          -
Sub EventValueChangeLinelist(oRange As Range)

    Dim T_geo As BetterArray
    Set T_geo = New BetterArray
    Dim sList As String
    Dim sControlType As String 'Control type
    Dim sLabel As String
    Dim sCustomVarName As String
    Dim sNote As String
    Dim sListAutoType As String
    Dim sVarName As String
    Dim iNumCol As Integer
    Dim iChoiceCol As Integer
    Dim choiceLo As ListObject
    Dim sChoiceAutoType As String
    Dim iRow As Integer
    Dim Rng As Range

    On Error GoTo errHand
    iNumCol = oRange.Column
    sControlType = ActiveSheet.Cells(C_eStartLinesLLMainSec - 1, iNumCol).value

    If oRange.Row > C_estartlineslldata + 1 Then

        Select Case sControlType

            Case C_sDictControlGeo
                ' adm1 has been modified, we will correct and set validation to adm2

                BeginWork xlsapp:=Application
                ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = ""
                oRange.Offset(, 2).Validation.Delete
                oRange.Offset(, 2).value = ""
                oRange.Offset(, 3).Validation.Delete
                oRange.Offset(, 3).value = ""

                If oRange.value <> vbNullString Then

                    'Filter on adm1
                    Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabAdm2), 1, oRange.value, returnIndex:=2)
                    'Build the validation list for adm2
                    sList = T_geo.ToString(Separator:=",", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                    'Set the validation list on adm2
                    Call Helpers.SetValidation(oRange.Offset(, 1), sList, 2)
                    T_geo.Clear
                End If

                Call ProtectSheet
                EndWork xlsapp:=Application

            Case C_sDictControlGeo & "2"

                'Adm2 has been modified, we will correct and filter adm3
                BeginWork xlsapp:=Application
                ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = vbNullString
                oRange.Offset(, 2).Validation.Delete
                oRange.Offset(, 2).value = vbNullString

                If oRange.value <> vbNullString Then
                    Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabAdm3), 1, oRange.Offset(, -1).value, 2, oRange.value, returnIndex:=3)
                    sList = T_geo.ToString(Separator:=",", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                    Call Helpers.SetValidation(oRange.Offset(, 1), sList, 2)
                    T_geo.Clear
                End If

                Call ProtectSheet
                EndWork xlsapp:=Application

            Case C_sDictControlGeo & "3"
                'Adm 3 has been modified, correct and filter adm4
                BeginWork xlsapp:=Application
                ActiveSheet.Unprotect (ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngDebuggingPassWord).value)

                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = vbNullString

                If oRange.value <> vbNullString Then
                    'Take the adm4 table
                    Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabAdm4), 1, _
                                            oRange.Offset(, -2).value, 2, oRange.Offset(, -1).value, 3, oRange.value, returnIndex:=4)
                    sList = T_geo.ToString(Separator:=",", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                    Call Helpers.SetValidation(oRange.Offset(, 1), sList, 2)
                    T_geo.Clear
                End If

                Call ProtectSheet
                EndWork xlsapp:=Application


            Case Else

        End Select

    End If

    If oRange.Row = C_estartlineslldata And sControlType = C_sDictControlCustom Then
        'The name of custom variables has been updated, update the dictionary
        sCustomVarName = ActiveSheet.Cells(C_estartlineslldata + 1, iNumCol).value
        sNote = GetDictColumnValue(sCustomVarName, C_sDictHeaderSubLab)
        sLabel = Replace(oRange.value, sNote, "")
        sLabel = Replace(sLabel, Chr(10), "")

        Call UpdateDictionaryValue(sCustomVarName, C_sDictHeaderMainLab, sLabel)

    End If

    If oRange.Name.Name = ActiveSheet.Name & "_" & C_sGotoSection Then

      Set Rng =  ActiveSheet.Rows(C_eStartLinesLLMainSec).Find(What:=oRange.value, _
       LookIn:= xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
        SearchDirection:= xlNext, MatchCase:=True, SearchFormat:=False)

        If Not Rng is Nothing Then
            Rng.Activate
        End If

    End If

errHand:

End Sub


Sub ClicImportMigration()
'Import exported data into the linelist
    F_ImportMig.Show
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

    Dim iChoiceCol As Integer
    Dim choiceLo As ListObject
    Dim sVarName As String
    Dim iRow As Integer
    Dim i As Integer
    Dim arrTable As BetterArray
    Dim PrevWksh As Worksheet
    Dim Rng As Range

    Set arrTable = New BetterArray

    On Error GoTo errHand

    i = 1

    Set PrevWksh = ThisWorkbook.Worksheets(sSheetName)

    With PrevWksh

        BeginWork xlsapp:=Application

            While (.Cells(C_estartlineslldata, i) <> vbNullString)

                Select Case .Cells(C_eStartLinesLLMainSec - 2, i).value

                    Case C_sDictControlChoiceAuto & "_origin"

                        sVarName = .Cells(C_estartlineslldata + 1, i).value
                        arrTable.FromExcelRange .Cells(C_estartlineslldata + 2, i), DetectLastColumn:=False, DetectLastRow:=True

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
            Wend

        EndWork xlsapp:=Application

    End With

    Set arrTable = Nothing
    Set PrevWksh = Nothing

    Exit Sub

errHand:
        EndWork xlsapp:=Application
        Exit Sub
End Sub
