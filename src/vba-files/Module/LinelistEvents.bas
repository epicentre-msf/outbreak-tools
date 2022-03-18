Attribute VB_Name = "LinelistEvents"

Option Explicit

Public IsLockedForProcess As Boolean

Sub ClicCmdGeoApp()

    Dim iNumCol As Integer
    Dim sType As String

    iNumCol = ActiveCell.Column
    ActiveSheet.Unprotect (C_sLLPassword)
    'On Error GoTo fin
    If ActiveCell.Row > C_eStartLinesLLData Then
        sType = FindDicColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, iNumCol).Name.Name, C_sDictHeaderControl) 'parce qu'un seul .Name ne suffit pas...
        Debug.Print sType & "1"
        Debug.Print ActiveSheet.Cells(C_eStartLinesLLData, iNumCol).Name.Name
        Select Case LCase(sType)
        Case "geo"
            iGeoType = 0
            Call LoadGeo(iGeoType)
    
        Case "hf"
            iGeoType = 1
            Call LoadGeo(iGeoType)
    
        Case Else
            MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
    
        End Select
    Else
        MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
    End If

    Exit Sub
    ActiveSheet.Protect Password:=C_sLLPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

fin:
    MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
    ActiveSheet.Protect Password:=C_sLLPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                    , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

End Sub

Sub ClicCmdAddRows()

    Dim oLstobj As Object

    ActiveSheet.Unprotect (C_sLLPassword)
    For Each oLstobj In ActiveSheet.ListObjects
        oLstobj.Resize Range(Cells(C_eStartLinesLLData, 1), Cells(oLstobj.DataBodyRange.Rows.Count + C_iNbLinesLLData + C_eStartLinesLLData, Cells(C_eStartLinesLLData, 1).End(xlToRight).Column))
    Next
    ActiveSheet.Protect Password:=C_sLLPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
End Sub

Sub ClicCmdExport()

    Dim i As Byte
    Dim iHeight As Integer
    Const C_CmdHeight As Integer = 6

    iHeight = 1

    With F_Export
        i = 2
        While i <= 6
            If Not isError(Sheets("Export").Cells(i, 4).value) Then 'lla
                If LCase(Sheets("Export").Cells(i, 4).value) <> "active" Then
                    .Controls("CMD_Export" & i - 1).Visible = False
                Else
                    .Controls("CMD_Export" & i - 1).Visible = True
                    .Controls("CMD_Export" & i - 1).Caption = Sheets("Export").Cells(i, 2).value
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
    
    Dim IsGeo As Boolean
    Dim T_geo As BetterArray
    Dim T_list As BetterArray
    Set T_geo = New BetterArray
    Set T_list = New BetterArray
    
    ActiveSheet.Unprotect (C_sLLPassword)
   
    If Not IsLockedForProcess Then
        IsLockedForProcess = True
        
        If oRange.Row > C_eStartLinesLLData Then
            On Error GoTo suivant                'if it is not geo for example or something with geo does not work
            If FindDicColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, oRange.Column).Name.Name, C_sDictHeaderControl) = C_sDictControlGeo Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +1
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = ""
                oRange.Offset(, 2).Validation.Delete
                oRange.Offset(, 2).value = ""
                oRange.Offset(, 3).Validation.Delete
                oRange.Offset(, 3).value = ""
                'First Geo adm1
                If oRange.value <> "" Then
                    'Take the adm2 table
                    T_geo.FromExcelRange Sheets(C_sSheetGeo).ListObjects(C_sTabADM2).DataBodyRange
                    T_geo.Sort SortColumn:=1
                    'Filter on adm1
                    Set T_geo = GetFilter(T_geo, 1, oRange.value)
                    'Build the validation list for adm2
                    T_list.Items = T_geo.ExtractSegment(ColumnIndex:=2)
                    Call BuildListGeo(oRange.Offset(, 1), T_list)
                    T_geo.Clear
                    Set T_geo = Nothing
                End If
            ElseIf FindDicColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, oRange.Column - 1).Name.Name, C_sDictHeaderControl) = C_sDictControlGeo Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +2
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = ""
                oRange.Offset(, 2).Validation.Delete
                oRange.Offset(, 2).value = ""
        
                If oRange.value <> "" Then
                    'Take the adm3 table
                    T_geo.FromExcelRange Sheets(C_sSheetGeo).ListObjects(C_sTabADM3).DataBodyRange
                    'Filter on adm1
                    Set T_geo = GetFilter(T_geo, 1, oRange.Offset(, -1).value)
                    'Filter on adm2
                    Set T_geo = GetFilter(T_geo, 2, oRange.value)
                    'Build the validation list for adm3
                    T_list.Items = T_geo.ExtractSegment(ColumnIndex:=3)
                    Call BuildListGeo(oRange.Offset(, 1), T_list)
                    T_list.Clear
                    T_geo.Clear
                    Set T_geo = Nothing
                    Set T_list = Nothing
                End If
        
            ElseIf FindDicColumnValue(ActiveSheet.Cells(C_eStartLinesLLData, oRange.Column - 2).Name.Name, _
                                    C_sDictHeaderControl) = C_sDictControlGeo Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +3
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = ""
        
                If oRange.value <> "" Then
                    'Take the adm4 table
                    T_geo.FromExcelRange Sheets(C_sSheetGeo).ListObjects(C_sTabADM4).DataBodyRange
                    'Filter on adm1
                    Set T_geo = GetFilter(T_geo, 1, oRange.Offset(, -2).value)
                    'Filter on adm2
                    Set T_geo = GetFilter(T_geo, 2, oRange.Offset(, -1).value)
                    'Filter on adm3
                    Set T_geo = GetFilter(T_geo, 3, oRange.value)
                    'Build the validation list for adm4
                    T_list.Items = T_geo.ExtractSegment(ColumnIndex:=4)
                    Call BuildListGeo(oRange.Offset(, 1), T_list)
                    T_geo.Clear
                    Set T_geo = Nothing
                    T_list.Clear
                    Set T_list = Nothing
                End If
            End If
suivant:
        
        End If
    
        IsLockedForProcess = False
         ActiveSheet.Protect Password:=C_sLLPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                         AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                         AllowFormattingColumns:=True
    End If
End Sub

'Build the dropdown validation list for the geo
Sub BuildListGeo(oRange As Range, T_list As BetterArray) 'sNameTab As String, iLigneDeb As Long, iLigneFin As Long, iCol As Long)
    
    Dim sList As String 'Validation formula list
    T_list.LowerBound = 1
    sList = T_list.Item(1)
    Dim i As Integer 'iterator
    For i = 2 To T_list.UpperBound
     sList = sList & "," & T_list.Item(i)
    Next

    With oRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=sList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub






