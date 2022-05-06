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
    DebugMode = True
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
                    Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabADM2), 1, oRange.value, returnIndex:=2)
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
                    Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabADM3), 1, oRange.Offset(, -1).value, 2, oRange.value, returnIndex:=3)
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
                    Set T_geo = FilterLoTable(ThisWorkbook.Worksheets(C_sSheetGeo).ListObjects(C_sTabADM4), 1, _
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

Sub ClicImportMigration()
'Import exported data into the linelist

    Dim wbkexp As Workbook, wbkLL As Workbook
    Dim shData As Worksheet, shSource As Worksheet, shTarget As Worksheet
    Dim lstobj  As ListObject
    Dim iLastSh As Integer, iLastexp As Integer, i As Integer, j As Integer, iPpos As Integer
    Dim iRows As Integer, iCols As Integer, iStart As Integer, iColExp As Integer, iRowTarget As Integer, iColTarget As Integer
    Dim rgResult As Range
    Dim sLabel As String, sPath As String, sSearchLabel As String, sSearchLabelo As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbkLL = ActiveWorkbook
    
    iLastSh = wbkLL.Sheets.Count
    iLastexp = iLastSh
    
     sPath = LoadFile("*.xlsx", "")
    
    If sPath = "" Then Exit Sub
    
    Workbooks.Open Filename:=sPath, Password:=Range("RNG_PrivateKey").value
    
    Set wbkexp = ActiveWorkbook
    
    For Each shData In wbkexp.Sheets
        
        If shData.Name <> "Dictionary" And shData.Name <> "Choices" And shData.Name <> "Translations" Then
            
            Sheets(shData.Name).Copy After:=wbkLL.Sheets(iLastexp)
            ActiveSheet.Name = shData.Name & "_Exp"

            iLastexp = iLastexp + 1
            wbkexp.Activate
            
        End If
        
    Next shData

    wbkexp.Close
    
    Set wbkexp = Nothing
    Set wbkLL = Nothing
    
    For i = iLastexp To iLastSh + 1 Step -1

        Set shSource = Sheets(Sheets(i).Name)
        Set shTarget = Sheets(Left(Sheets(i).Name, Len(Sheets(i).Name) - 4))
        
        If LCase(shSource.Name) = "admin_exp" Then
        
            shSource.Range("B1:B7").Copy
            shTarget.Select
            Range("C15").Select
            ActiveSheet.Paste

        Else
        
            iCols = shSource.Cells(1, 1).End(xlToRight).Column
            iRows = shSource.Cells(1, 1).End(xlDown).Row
            iStart = shTarget.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
            
            For Each lstobj In shTarget.ListObjects
                If lstobj.Name = "o" & shTarget.Name Then
                    iRowTarget = shTarget.ListObjects(lstobj.Name).ListRows.Count
                    Exit For
                End If
            Next

            shTarget.Select
                            
            Do While (iRowTarget - (iStart - 5)) < iRows
                Call ClicCmdAddRows
                iRowTarget = shTarget.ListObjects(lstobj.Name).ListRows.Count
            Loop
            
            Sheets("admin").Select
            
            j = 1
            

            Do While j <= iCols 'Number of columns in source file
                    
                If shTarget.Cells(5, j).value = "" Then Exit Do
                
                sSearchLabelo = shTarget.Cells(5, j).value
                sSearchLabel = sSearchLabelo
                
                iPpos = InStr(sSearchLabel, Chr(10))
                
                If iPpos > 0 Then
                    sSearchLabel = Left(sSearchLabel, iPpos - 1)
                End If
                
                Set rgResult = Sheets("Dictionary").Columns(2).Find(What:=sSearchLabel, LookAt:=xlValue)
                
                If Not rgResult Is Nothing Then
                    Do
                        If Sheets("Dictionary").Cells(rgResult.Row, 5).value = shTarget.Name Then
                            sLabel = Sheets("Dictionary").Cells(rgResult.Row, 1).value
                            Exit Do
                        Else
                            Set rgResult = Sheets("Dictionary").Columns(2).FindNext(rgResult)
                        End If
        
                    Loop While Not rgResult Is Nothing
                Else
                
                    sLabel = shSource.Cells(1, j).value
                
                End If

                iColExp = shSource.Rows(1).Find(What:=sLabel, LookAt:=xlWhole).Column
                shSource.Select
                shSource.Range(Cells(2, iColExp), Cells(iRows, iColExp)).Copy
                
                shTarget.Select
                
                iColTarget = 0
                iColTarget = shTarget.Rows(5).Find(What:=sSearchLabelo, LookAt:=xlWhole).Column
                
                If iColTarget > 0 And Not shTarget.Cells(6, j).HasFormula Then
                    shTarget.Cells(iStart, iColTarget).Select
                    ActiveSheet.Paste
                    Columns(iColTarget).EntireColumn.AutoFit
                End If
                
                j = j + 1
            Loop
            
        End If
        
        Application.DisplayAlerts = False
        shSource.Delete
        Application.DisplayAlerts = True
        
    Next i
    
    Sheets("admin").Select
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub




