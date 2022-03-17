Attribute VB_Name = "M_LineList"
Option Explicit

Public T_typeDico
Public IsLockedForProcess As Boolean
Const C_TitleCol As Byte = 5
Const C_PWD As String = "1234"

Function CreateTitleDic() As Scripting.Dictionary

    Dim i As Integer
    Dim D_temp As New Scripting.Dictionary

    i = 1
    While Sheets("dico").Cells(1, i).value <> ""
        D_temp.Add Sheets("dico").Cells(1, i).value, i
        i = i + 1
    Wend
    Set CreateTitleDic = D_temp
    Set D_temp = Nothing

End Function

Function CreateDataDic(D_title As Scripting.Dictionary)

    Dim i As Integer
    Dim T_temp

    ReDim T_temp(D_title.Count, Sheets("dico").Cells(1, 1).End(xlDown).Row)
    i = 1
    While i < UBound(T_temp, 1)
        T_temp(D_title("Variable name"), i) = Sheets("dico").Cells(i, D_title("Variable name")).value
        T_temp(D_title("Main label"), i) = Sheets("dico").Cells(i, D_title("Main label")).value
        T_temp(D_title("Control"), i) = Sheets("dico").Cells(i, D_title("Control")).value
        T_temp(D_title("Type"), i) = Sheets("dico").Cells(i, D_title("Type")).value

        i = i + 1
    Wend
    CreateDataDic = T_temp

End Function

Function LetDataDic(sName As String, sColName As String) As String '

    Dim i As Integer
    Dim D_title As New Scripting.Dictionary
    Dim T_data

    Set D_title = CreateTitleDic
    T_data = CreateDataDic(D_title)

    If Not IsEmptyTable(T_data) Then
        i = 1
        While i < UBound(T_data) And T_data(D_title("Variable name"), i) <> sName
            i = i + 1
        Wend
        If T_data(D_title("Variable name"), i) = sName Then
            LetDataDic = T_data(D_title(sColName), i)
        End If
    End If
    Set D_title = Nothing

End Function

Sub clicCmdGeoApps()

    Dim iNumCol As Integer
    Dim sType As String

    iNumCol = ActiveCell.Column
    ActiveSheet.Unprotect (C_PWD)
    'On Error GoTo fin
    If ActiveCell.Row > C_TitleCol Then
        sType = LetDataDic(ActiveSheet.Cells(C_TitleCol, iNumCol).Name.Name, "Control") 'parce qu'un seul .Name ne suffit pas...
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
    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

fin:
    MsgBox "Vous n'etes pas sur la bonne cellule" 'MSG_WrongCells
    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

End Sub

Sub clicAdd200L()

    Dim oLstobj As Object

    ActiveSheet.Unprotect (C_PWD)
    For Each oLstobj In ActiveSheet.ListObjects
        oLstobj.Resize Range(Cells(C_TitleCol, 1), Cells(oLstobj.DataBodyRange.Rows.Count + 200 + C_TitleCol, Cells(C_TitleCol, 1).End(xlToRight).Column))
    Next
    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

End Sub

Sub clicExport()

    Dim i As Byte
    Dim iHeight As Integer
    Const C_CmdHeight As Integer = 6

    iHeight = 1

    ActiveSheet.Unprotect (C_PWD)
    With F_Export
        i = 2
        While i <= 6
            If Not IsError(Sheets("Export").Cells(i, 4).value) Then 'lla
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
        iHeight = iHeight + 24 + C_CmdHeight
    
        .CMD_Retour.Top = iHeight + 5
        '.CMD_Retour.Visible = True
        iHeight = .CMD_Retour.Top + .CMD_Retour.Height + 24 + 10
        .Height = iHeight
        .Width = 168
    
        .show
    End With

    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

End Sub

'Trigerring event when the linelist sheet has some values within                                                          -
Sub EventSheetLineListPatient(oRange As Range)
    
    Dim IsGeo As Boolean
    Dim T_geo As BetterArray
    Dim T_list As BetterArray
    Set T_geo = New BetterArray
    Set T_list = New BetterArray
   
    If Not IsLockedForProcess Then
        IsLockedForProcess = True
        ActiveSheet.Unprotect (C_PWD)
    
        If oRange.Row > C_TitleCol Then
            On Error GoTo suivant                'if it is not geo for example or something with geo does not work
            If LCase(LetDataDic(Cells(C_TitleCol, oRange.Column).Name.Name, "Control")) = "geo" Then
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
                    T_geo.FromExcelRange Sheets("GEO").ListObjects("T_ADM2").DataBodyRange
                    T_geo.Sort SortColumn:=1
                    'Filter on adm1
                    Set T_geo = GetFilter(T_geo, 1, oRange.value)
                    'Build the validation list for adm2
                    T_list.Items = T_geo.ExtractSegment(ColumnIndex:=2)
                    Call BuildListGeo(oRange.Offset(, 1), T_list)
                    T_geo.Clear
                    Set T_geo = Nothing
                End If
            ElseIf LCase(LetDataDic(Cells(C_TitleCol, oRange.Column - 1).Name.Name, "Control")) = "geo" Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +2
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = ""
                oRange.Offset(, 2).Validation.Delete
                oRange.Offset(, 2).value = ""
        
                If oRange.value <> "" Then
                    'Take the adm3 table
                    T_geo.FromExcelRange Sheets("GEO").ListObjects("T_ADM3").DataBodyRange
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
        
            ElseIf LCase(LetDataDic(Cells(C_TitleCol, oRange.Column - 2).Name.Name, "Control")) = "geo" Then
                'on controle qu'on a bien ecrit une data geo et remplissage de la colonne +3
                oRange.Offset(, 1).Validation.Delete
                oRange.Offset(, 1).value = ""
        
                If oRange.value <> "" Then
                    'Take the adm4 table
                    T_geo.FromExcelRange Sheets("GEO").ListObjects("T_ADM4").DataBodyRange
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
            'Testing color for dates and numeric values (maybe directly in the validation?)
            If LCase(LetDataDic(Cells(C_TitleCol, oRange.Column).Name.Name, "type")) = "date" Then
                If Not IsDate(oRange.value) Then
                    oRange.Interior.Color = vbRed
                End If
            ElseIf LCase(LetDataDic(Cells(C_TitleCol, oRange.Column).Name.Name, "type")) = "interger" Or InStr(1, LCase(LetDataDic(Cells(C_TitleCol, oRange.Column).Name.Name, "type")), "decimal") > 0 Then
                If Not IsNumeric(oRange.value) Then
                    oRange.Interior.Color = vbRed
                End If
            End If
    
        End If
    
        ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                               , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    
        IsLockedForProcess = False
    End If
    Application.ScreenUpdating = True

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





