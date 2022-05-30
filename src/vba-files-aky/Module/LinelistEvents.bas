Attribute VB_Name = "LinelistEvents"

Option Explicit

Public iGeoType As Byte

Sub ClicCmdGeoApp()

    Dim iNumCol As Integer
    Dim sType As String

    iNumCol = ActiveCell.Column

    If ActiveCell.Row > C_eStartlinesLLData + 1 Then

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
        oLstobj.Resize Range(Cells(C_eStartlinesLLData + 1, 1), Cells(oLstobj.DataBodyRange.Rows.Count + C_iNbLinesLLData + C_eStartlinesLLData + 1, Cells(C_eStartlinesLLData + 1, 1).End(xlToRight).Column))
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
Sub EventValueChangeLinelist(oRange As Range)

    Dim T_geo As BetterArray
    Set T_geo = New BetterArray
    Dim sList As String
    Dim sControlType As String 'Control type
    Dim sLabel As String
    Dim sCustomVarName As String
    Dim sNote As String
    Dim iNumCol As Integer

    On Error GoTo errHand
    iNumCol = oRange.Column
    sControlType = ActiveSheet.Cells(C_eStartLinesLLMainSec - 1, iNumCol).value

    If oRange.Row > C_eStartlinesLLData + 1 Then

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
    End If

    If oRange.Row = C_eStartlinesLLData And sControlType = C_sDictControlCustom Then
        'The name of custom variables has been updated, update the dictionary
        sCustomVarName = ActiveSheet.Cells(C_eStartlinesLLData + 1, iNumCol).value
        sNote = GetDictColumnValue(sCustomVarName, C_sDictHeaderSubLab)
        sLabel = Replace(oRange.value, sNote, "")
        sLabel = Replace(sLabel, Chr(10), "")

        Call UpdateDictionaryValue(sCustomVarName, C_sDictHeaderMainLab, sLabel)

    End If

errHand:

End Sub

Sub EventOpenLinelist()
    Dim iNbCols As Integer
    Dim Wksh As Worksheet
    Dim i As Integer
    Dim hasData As Boolean
    Dim LLEnteredData As BetterArray
    Dim ListAutoData As BetterArray
    Dim LoRng As Range
    Dim sVarName As String
    Dim sAutoVariable As String
    Dim sAutoSheetName As String
    Dim iAutoColumn As Integer
    Dim sList As String
    Dim iDataLength As Integer

    Set Wksh = ActiveSheet
    hasData = False
    iDataLength = 0

    With Wksh
        BeginWork xlsapp:=Application
        '.Unprotect (C_sLLPassword)

        iNbCols = .Cells(C_eStartlinesLLData, Columns.Count).End(xlToLeft).Column

        For i = 1 To iNbCols

            If .Cells(C_eStartLinesLLMainSec - 1, i).value = C_sDictControlChoiceAuto Then

                'First take the data in memory (we need it since we don't know the data entered by the operator)
                If Not hasData Then
                    Set LLEnteredData = New BetterArray
                    LLEnteredData.FromExcelRange .Cells(C_eStartlinesLLData + 2, 1), DetectLastRow:=True, DetectLastColumn:=True
                    iDataLength = LLEnteredData.Length
                    'Resize the listObject to the current entered data + 1
                    Set LoRng = .Range(.Cells(C_eStartlinesLLData + 1, 1), .Cells(iDataLength + C_eStartlinesLLData + 2, iNbCols))
                    .ListObjects("o" & ClearString(.Name)).Resize LoRng
                    hasData = True
                End If

                'Do all this only if there is Data in the sheet

                If iDataLength > 2 Then
                    'VarName of the actual sheet
                    sVarName = .Cells(C_eStartlinesLLData + 1, i).value
                    'Get the entered list of the listauto
                    sAutoVariable = GetDictColumnValue(sVarName, C_sDictHeaderChoices)

                    'Get the SheetName of the auto
                    sAutoSheetName = GetDictColumnValue(sAutoVariable, C_sDictHeaderSheetName)

                    'Get the Column of the auto
                    iAutoColumn = GetDictColumnValue(sAutoVariable, C_sDictHeaderIndex)

                    If iAutoColumn > 0 Then
                        'Get the entered values for the auto
                        Set ListAutoData = New BetterArray
                        ListAutoData.FromExcelRange ThisWorkbook.Worksheets(sAutoSheetName).Cells(C_eStartlinesLLData + 2, iAutoColumn), DetectLastColumn:=False, DetectLastRow:=True
                        ListAutoData.Reverse
                        'Get validation list
                        sList = ListAutoData.ToString(Separator:=",", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)

                        'Set the validation list
                        Call Helpers.SetValidation(.Cells(iDataLength + 1, i), sList, 2)
                    End If
                End If

            End If
        Next

        If hasData Then
            'resize because data has been updated
            Set LoRng = .Range(.Cells(C_eStartlinesLLData + 1, 1), .Cells(C_iNbLinesLLData + C_eStartlinesLLData - 1, iNbCols))
            .ListObjects("o" & ClearString(.Name)).Resize LoRng
        End If

        'Call ProtectSheet
        EndWork xlsapp:=Application
    End With
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
