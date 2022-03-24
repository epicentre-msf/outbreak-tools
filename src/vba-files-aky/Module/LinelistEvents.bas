Attribute VB_Name = "LinelistEvents"

Option Explicit

Public iGeoType As Byte

Sub ClicCmdGeoApp()
    
    Dim iNumCol As Integer
    Dim sType As String

    iNumCol = ActiveCell.Column
    ActiveSheet.Unprotect (C_sLLPassword)
    
    'On Error GoTo fin
    If ActiveCell.Row > C_eStartLinesLLData Then
        
        sType = GetDictColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, iNumCol).Name.Name, C_sDictHeaderControl) 'parce qu'un seul .Name ne suffit pas...
        
        Select Case sType
        
        Case C_sDictControlGeo
            iGeoType = 0
            Call LoadGeo(iGeoType)
    
        Case C_sDictControlHf
            iGeoType = 1
            Call LoadGeo(iGeoType)
    
        Case Else
            MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
            Call ProtectSheet

        End Select
    Else
        MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
        Call ProtectSheet

    End If

    Exit Sub
    Call ProtectSheet

fin:
    MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
    Call ProtectSheet
End Sub

Sub ClicCmdAddRows()

    Dim oLstobj As Object

    ActiveSheet.Unprotect (C_sLLPassword)
    Application.EnableEvents = False
    
    For Each oLstobj In ActiveSheet.ListObjects
        oLstobj.Resize Range(Cells(C_eStartLinesLLData, 1), Cells(oLstobj.DataBodyRange.Rows.Count + C_iNbLinesLLData + C_eStartLinesLLData, Cells(C_eStartLinesLLData, 1).End(xlToRight).Column))
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
            If Not isError(Sheets("Exports").Cells(i, 4).value) Then 'lla
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

'Trigerring event when the linelist sheet has some values within                                                          -
Sub EventSheetLineListPatient(oRange As Range)

    Dim T_geo As BetterArray
    Set T_geo = New BetterArray
    Dim sList As String
    
    BeginWork xlsapp:=Application
    ActiveSheet.Unprotect (C_sLLPassword)
        If oRange.Row > C_eStartLinesLLData Then
            On Error GoTo errHand                'if it is not geo for example or something with geo does not work
            If GetDictColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, oRange.Column).Name.Name, C_sDictHeaderControl) = C_sDictControlGeo Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +1
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = ""
                oRange.Offset(, 2).Validation.Delete
                oRange.Offset(, 2).value = ""
                oRange.Offset(, 3).Validation.Delete
                oRange.Offset(, 3).value = ""
                'First Geo adm1
                
                If oRange.value <> "" Then
                    'Filter on adm1
                    Set T_geo = FilterLoTable(ThisWorkbook.worksheets(C_sSheetGeo).ListObjects(C_sTabADM2), 1, oRange.value, returnIndex:=2)
                    'Build the validation list for adm2
                    sList = T_geo.ToString(Separator:=",", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                    Call Helpers.SetValidation(oRange.Offset(, 1), sList, 2)
                    T_geo.Clear
                End If
            ElseIf GetDictColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, oRange.Column - 1).Name.Name, C_sDictHeaderControl) = C_sDictControlGeo Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +2
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = vbNullString
                oRange.Offset(, 2).Validation.Delete
                oRange.Offset(, 2).value = vbNullString
        
                If oRange.value <> vbNullString Then
                    'Take the adm3 table
                    Set T_geo = FilterLoTable(ThisWorkbook.worksheets(C_sSheetGeo).ListObjects(C_sTabADM3), 1, oRange.Offset(, -1).value, 2, oRange.value, returnIndex:=3)
                    sList = T_geo.ToString(Separator:=",", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                    Call Helpers.SetValidation(oRange.Offset(, 1), sList, 2)
                    T_geo.Clear
                End If
        
            ElseIf GetDictColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, oRange.Column - 2).Name.Name, _
                                    C_sDictHeaderControl) = C_sDictControlGeo Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +3
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = vbNullString
        
                If oRange.value <> vbNullString Then
                    'Take the adm4 table
                    Set T_geo = FilterLoTable(ThisWorkbook.worksheets(C_sSheetGeo).ListObjects(C_sTabADM4), 1, _
                                             oRange.Offset(, -2).value, 2, oRange.Offset(, -1).value, 3, oRange.value, returnIndex:=4)

                    sList = T_geo.ToString(Separator:=",", OpeningDelimiter:="", ClosingDelimiter:="", QuoteStrings:=False)
                    Call Helpers.SetValidation(oRange.Offset(, 1), sList, 2)
                    T_geo.Clear
                End If
            End If
errHand:
        
        End If
    
    Call ProtectSheet
    EndWork xlsapp:=Application
End Sub






