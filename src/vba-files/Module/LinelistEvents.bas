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
                MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
        End Select
    Else
        MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
    End If
End Sub

Sub ClicCmdAddRows()

    Dim oLstobj As Object

    ActiveSheet.Unprotect (C_sLLPassword)
    Application.EnableEvents = False

    For Each oLstobj In ActiveSheet.ListObjects
        oLstobj.Resize Range(Cells(C_estartlineslldata + 1, 1), Cells(oLstobj.DataBodyRange.Rows.Count + C_iNbLinesLLData + C_estartlineslldata + 1, Cells(C_estartlineslldata + 1, 1).End(xlToRight).Column))
    Next

    Call ProtectSheet
    Application.EnableEvents = True
End Sub

Sub ClicCmdExport()

    Dim i As Byte
    Dim iHeight As Integer
    Const C_CmdHeight As Integer = 6

    iHeight = 1

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
        .Width = 168
        .show
    End With
End Sub


Sub ClicCmdDebug()
    Static DebugMode As Boolean
    Dim pwd As String
    Dim sh As Worksheet
    pwd = Inputbox("Provide the debugging password", "DEBUG MODE", "1234")

    If pwd = C_sLLPassword Then
        For Each sh In ThisWorkbook.Worksheets
            If sh.protectcontents = True Then
                sh.Unprotect pwd
            End If
        Next
    Else
        MsgBox "Wrong Password!", vbok, "DEBUG MODE"
    End If
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
    Dim rng As Range

    On Error GoTo errHand
    iNumCol = oRange.Column
    sControlType = ActiveSheet.Cells(C_eStartLinesLLMainSec - 1, iNumCol).value
    sChoiceAutoType = ActiveSheet.Cells(C_eStartLinesLLMainSec - 2, iNumCol).value

    If oRange.Row > C_estartlineslldata + 1 Then

        Select Case sControlType

            Case C_sDictControlGeo
                ' adm1 has been modified, we will correct and set validation to adm2

                BeginWork xlsapp:=Application
                ActiveSheet.Unprotect (C_sLLPassword)

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
                ActiveSheet.Unprotect (C_sLLPassword)

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
                ActiveSheet.Unprotect (C_sLLPassword)

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

        Select Case sChoiceAutoType

            Case C_sDictControlChoiceAuto & "_origin"
                BeginWork xlsapp:=Application
                sVarName = ActiveSheet.Cells(C_estartlineslldata + 1, iNumCol).value
                With ThisWorkbook.Worksheets(C_sSheetChoiceAuto)
                    Set choiceLo = .ListObjects("o" & C_sDictControlChoiceAuto & "_" & sVarName)

                    iChoiceCol = choiceLo.DataBodyRange.Column
                    iRow = .Cells(Rows.Count, iChoiceCol).End(xlUp).Row

                    'A simple safeguard agains empty values
                    If .Cells(iRow, iChoiceCol).value = vbNullString Then
                        iRow = iRow - 1
                    End If
                    .Cells(iRow + 1, iChoiceCol).value = oRange.value

                    choiceLo.Resize Range(.Cells(C_eStartlinesListAuto, iChoiceCol), .Cells(iRow + 1, iChoiceCol))
                    choiceLo.DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlYes

                    Set rng = choiceLo.ListColumns(1).Range
                    With choiceLo.Sort
                        .SortFields.Clear
                        .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlDescending
                        .Header = xlYes
                        .Apply
                    End With
                End With
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

errHand:

End Sub


Sub ClicImportMigration()
'Import exported data into the linelist
    F_ImportMig.show
End Sub


Sub ClicExportMigration()

    Static AfterFirstClicMig As Boolean

    If AfterFirstClicMig Then
        [F_ExportMig].show
    Else
        'For the first click Thick Migration and Geo and put historic to false
        'For subsequent clicks, just show what have been ticked
        [F_ExportMig].CHK_ExportMigData.value = True
        [F_ExportMig].CHK_ExportMigGeo.value = True
        [F_ExportMig].CHK_ExportMigGeoHistoric.value = True
        [F_ExportMig].show
        AfterFirstClicMig = True
    End If
End Sub
