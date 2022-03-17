Attribute VB_Name = "LinelistExport"
'Option Explicit
'
'Const C_eStartLinesLLData As Byte = 1
'Const C_TitleSource As Byte = 5
'Const C_PWD As String = "1234"
'
''on va avoir besoin de CreateDicTitle dans M_LineList
''on fonctionne par exclusion
'
'Private Function GetExportHeaders(iTypeExport As Byte, sSheetName As String, _
'                    DictHeaders As BetterArray, DictData As BetterArray, _
'                    SheetNameData As BetterArray, VarNamesData as BetterArray) as BetterArray
'
'    Dim ExportHeadersData as BetterArray
'    Dim sExportName As String
'    Dim iFirstIndex As Integer                                'Iterator
'    Dim iLastIndex  As Integer
'    Dim i as Integer
'    Dim sActualExportValue As string
'
'
'
'    Set ExportHeadersData = New BetterArray
'
'    'Get the type of the export
'    Select Case iTypeExport
'        Case 1
'         sExportName = C_sDictHeaderExport1
'        Case 2
'         sExportName = C_sDictHeaderExport2
'        Case 3
'         sExportName = C_sDictHeaderExport3
'        Case 4
'         sExportName = C_sDictHeaderExport4
'        Case 5
'         sExportName = C_sDictHeaderExport5
'        Case Else
'         sExportName = C_sDictHeaderExport1
'    End Select
'
'    iFirstIndex = SheetNameData.IndexOf(sSheetName)
'    iLastIndex = SheetNameData.LastIndexOf(sSheetName)
'
'    If iFirstIndex > 0 and iLastIndex > 0 Then
'
'        For i = iFirstIndex To iLastIndex
'
'            sActualExportValue = ClearString(DictData.Items(i, DictHeaders.indexOf(sExportName)))
'
'            If sActualExportValue = C_sYes Then
'                'push adms with the variable names
'                ExportHeadersData.push VarNamesData.items(i)
'            End If
'        Next
'    End If
'
'    Set GetExportHeaders = ExportHeadersData
'    Set ExportHeadersData = Nothing
'End Function
'
'
'Private Function GetVarNames(DictData as BetterArray, DictHeaders As BetterArray) As BetterArray
'
'    Dim VarNamesData as BetterArray
'    Dim i as Integer
'    Dim sActualControl as String
'    Dim sActualVarname as String
'
'    Set VarnameData = New BetterArray
'
'    For i = 1 To DictData.Length
'
'        sActualControl = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderControl))
'        sActualVarname = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderVarName))
'
'        If sActualControl = C_sDictControlGeo  Then
'            'push adms with the variable names
'            VarNamesData.push C_sAdmName & "1" & "_" sActualVarName, _
'             C_sAdmName & "2" & "_" sActualVarName,_
'             C_sAdmName & "3" & "_" sActualVarName, C_sAdmName & "4" & "_" sActualVarName
'        Else
'            VarNamesData.push sActualVarName
'        End If
'    Next
'    Set GetVarNames = VarNamesData
'    Set VarNamesData = Nothing
'End Function
'
'Private Function GetExportIndexes(iTypeExport As Byte, sSheetName As String, _
'                    DictHeaders As BetterArray, DictData As BetterArray, _
'                    SheetNameData As BetterArray, ColumnIndexes as BetterArray) as BetterArray
'
'    Dim ExportIndexes as BetterArray
'    Dim sExportName As String
'    Dim iFirstIndex As Integer                                'Iterator
'    Dim iLastIndex  As Integer
'    Dim i as Integer
'    Dim sActualExportValue As string
'
'    Set ExportIndexes = New BetterArray
'
'    'Get the type of the export
'    Select Case iTypeExport
'        Case 1
'         sExportName = C_sDictHeaderExport1
'        Case 2
'         sExportName = C_sDictHeaderExport2
'        Case 3
'         sExportName = C_sDictHeaderExport3
'        Case 4
'         sExportName = C_sDictHeaderExport4
'        Case 5
'         sExportName = C_sDictHeaderExport5
'        Case Else
'         sExportName = C_sDictHeaderExport1
'    End Select
'
'    iFirstIndex = SheetNameData.IndexOf(sSheetName)
'    iLastIndex = SheetNameData.LastIndexOf(sSheetName)
'
'    If iFirstIndex > 0 and iLastIndex > 0 Then
'        For i = iFirstIndex To iLastIndex
'
'            sActualExportValue = ClearString(DictData.Items(i, DictHeaders.indexOf(sExportName)))
'
'            If sActualExportValue = C_sYes Then
'                'push adms with the variable names
'                ExportIndexes.push ColumnIndex.items(i)
'            End If
'        Next
'    End If
'
'    Set GetExportHeaders = ExportIndexes
'    Set ExportIndexes = Nothing
'End Function
'
'
'
'
'Private Function GetColumnIndexes(DictData as BetterArray, DictHeaders As BetterArray) As BetterArray
'
'    Dim ColumnIndexes as BetterArray
'    Dim i as Integer
'    Dim sActualControl as String
'    Dim iActualIndex as Integer
'
'    Set ColumnIndexes = New BetterArray
'
'    For i = 1 To DictData.Length
'
'        sActualControl = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderControl))
'        iActualIndex = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderIndex))
'
'        If sActualControl = C_sDictControlGeo  Then
'            'push adms with the variable names
'            ColumnIndexes.push iActualIndex, iActualIndex + 1, iActualIndex + 2, iActualIndex + 3
'        Else
'            ColumnIndexes.push iActualIndex
'        End If
'    Next
'    Set GetColumnIndexes = ColumnIndexes
'    Set ColumnIndexes = Nothing
'End Function
'
'
'Private Function GetSheetNames(DictData as BetterArray, DictHeaders As BetterArray) as BetterArray
'    Dim SheetNameData as BetterArray
'    Dim i as integer
'    Dim sActualSheetType as String
'    Dim sActualControl as String
'    Dim sActualVarName as String
'
'    Set SheetNameData = New BetterArray
'    For i = 1 To DictData.Length
'        sActualType = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderType))
'        sActualSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
'        sActualControl = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderControl))
'
'        If sActualSheetType = C_sDictSheetTypeLL Then
'
'            If sActualControl = C_sDictControlGeo Then
'
'                SheetNameData.push sActualSheetName, sActualSheetName, sActualSheetName, sActualSheetName
'            Else
'                SheetNameData.push sActualSheetName
'            End if
'        End if
'    Next
'    Set GetSheetNames = SheetNameData
'    Set SheetNameData = Nothing
'End function
'
'
'Private Function GetExportValues(DictData as BetterArray, DictHeaders As BetterArray) as BetterArray
'    Dim SheetNameData as BetterArray
'    Dim i as integer
'    Dim sActualSheetType as String
'    Dim sActualControl as String
'    Dim sActualVarName as String
'
'    Set SheetNameData = New BetterArray
'    For i = 1 To DictData.Length
'        sActualType = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderType))
'        sActualSheetName = DictData.Items(i, DictHeaders.IndexOf(C_sDictHeaderSheetName))
'        sActualControl = DictData.items(i, DictHeaders.IndexOf(C_sDictHeaderControl))
'
'        If sActualSheetType = C_sDictSheetTypeLL Then
'
'            If sActualControl = C_sDictControlGeo Then
'
'                SheetNameData.push sActualSheetName, sActualSheetName, sActualSheetName, sActualSheetName
'            Else
'                SheetNameData.push sActualSheetName
'            End if
'        End if
'    Next
'    Set GetSheetNames = SheetNameData
'    Set SheetNameData = Nothing
'End function
'
'
'
'
'
'Sub Export(iTypeExport As Byte)
'
'    Dim DictData as BetterArray
'    Dim DictData as BetterArray
'    Dim VarNameData as BetterArray
'    Dim SheetNameData as BetterArray
'    Dim ExportHeadersData as BetterArray
'    Dim ExportColumnIndexes as BetterArray
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim xlsapp As Excel.Application
'    Dim T_data
'    Dim T_dataValid
'    Dim sNameListO As String
'    Dim sSheetName As String
'
'    Dim T_dico
'
'    Dim sPath As String
'    Dim sDirectory As String
'    Dim T_Path
'
'    Dim diaFolder As FileDialog
'    Dim new_path As String
'
'    Dim oCell As Object
'
'    Set xlsapp = New Excel.Application
'
'    BeginWork(xlsapp)
'
'    'Get Dicitionnary Data and headers
'    Set DictData = Helpers.GetData(ThisWorkbook, C_sParamSheetDict, 2)
'    Set DicHeaders = Helpers.GetHeaders(ThisWorkbook, C_sParamSheetDict, 1)
'    Set VarNameData = GetVarNames(DictData, DictHeaders)
'    Set SheetNameData = GetSheetNames(DictData, DictHeaders)
'
'    With xlsapp
'        .ScreenUpdating = False
'        .Visible = False
'        .Workbooks.Add
'
'        'Adding the sheets for export
'        sPrevSheetName = C_sParamSheetDict
'        i = 1
'        .Worksheets(1).Name = C_sParamSheetDict
'
'        'Writing the dictionary data
'        DictHeaders.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(1, 1), Transpose:=True
'        DictData.ToExcelRange   .Worksheets(C_sParamSheetDict).Cells(2, 1)
'
'        While i <= SheetNameData.UpperBound
'
'            .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = SheetNameData.item(i)
'            sPrevSheetName = SheetNameData.Item(i)
'
'            ExportHeadersData.Clear
'            ExportColumnIndexes.Clear
'
'            Set ExportHeadersData = GetExportHeaders(sPrevSheetName, _
'                    DictHeaders, DictData, _
'                    SheetNameData, VarNameData )
'
'            Set ExportColumnIndexes = GetExportIndexes(sPrevSheetName,
'                    DictHeaders, DictData, SheetName, VarNameData)
'
'            'Writing in the sheet
'
'            'Headers
'            ExportHeadersData.ToExcelRange .Worksheets(sPrevSheetName).Range("A1"), Transpose:=True
'
'            'Take values from the linelist and put them in the Sheet
'            For j = 1 To ExportColumnIndexes.UpperBound
'                iColIndex = ExportColumnIndexes.Item(j)
'                TempData.FromExcelRange Sheets(sPrevSheetName).Cells(C_eStartLinesLLData, iColIndex), DetectLastRow := True, DetectLastColumn := True
'                TempData.ToExcelRange .Worksheets(sPrevMainSec).Cells(1, j)
'            Next
'        Wend
'
'
'        'pour le dico
'
'        .Sheets.Add.Name = "Dico"
'        T_dico = CopyDico
'
'        'Set D_dico = CreateDicoName
'        i = 1
'        While i <= UBound(T_dico, 1)
'
'            T_dico(i, 1) = ReplaceCustomDico(CStr(T_dico(i, 0)), CStr(T_dico(i, 1)))
'
'            i = i + 1
'        Wend
'
'        If Not IsEmptyTable(T_dico) Then
'            .Sheets("Dico").Range("A1").Resize(UBound(T_dico, 1), UBound(T_dico, 2)) = T_dico
'        End If
'
'        'l'admin
'        Erase T_dataValid
'        T_dataValid = creationTabChamp(iTypeExport, "admin")
'        .Sheets.Add.Name = "Admin"
'        i = 0
'        j = 1
'        While i <= UBound(T_dataValid, 2)
'            .Sheets("Admin").Cells(j, 2).Name = T_dataValid(1, i)
'            .Sheets("Admin").Cells(j, 1).value = Range(T_dataValid(1, i)).Offset(, -1).value
'            .Sheets("Admin").Cells(j, 2).value = Range(T_dataValid(1, i)).value
'            j = j + 1
'            i = i + 1
'        Wend
'
'        'pour l'enregistrement
'        sPath = Sheets("Export").Cells(iTypeExport + 1, 5).value
'        If sPath <> "" Then
'            T_Path = Split(sPath, "+")
'
'            Set D_dico = CreateDicoName
'            i = 0
'            While i <= UBound(T_Path)
'                If InStr(1, T_Path(i), Chr(34)) = 0 Then
'                    If D_dico.Exists(Trim(T_Path(i))) Then
'                        sPath = Replace(sPath, Trim(T_Path(i)), Range(Trim(T_Path(i))).value)
'                    End If
'                End If
'                i = i + 1
'            Wend
'            Set D_dico = Nothing
'            sPath = Replace(Replace(Replace(sPath & "__" & Range("RNG_PublicKey").value & "__" & Format(Now, "yyyymmdd-HhNn"), " ", ""), "+", "__"), Chr(34), "")
'            sDirectory = LoadFolder
'            If sDirectory <> "" Then
'                sPath = sDirectory & Application.PathSeparator & sPath & ".xlsb"
'
'                i = 0
'                While Len(sPath) >= 255 And i < 3 'MSG_PathTooLong
'                    MsgBox "The path of the export file is too long so the file name gets truncated. Please select a folder higher in the hierarchy to save the export (ex: Desktop, Downloads, Documents etc.)"
'                    sDirectory = LoadFolder
'                    If sDirectory <> "" Then
'                        sPath = sDirectory & Application.PathSeparator & sPath
'                    End If
'                    i = i + 1
'                Wend
'                'on enregistre
'                If i < 3 Then
'                    .ActiveWorkbook.SaveAs Filename:=sPath, FileFormat:=xlExcel12, CreateBackup:=False, Password:=Range("RNG_PrivateKey").value
'                    MsgBox "File saved" & Chr(10) & "Password : " & Range("RNG_PrivateKey").value 'MSG_FileSaved        'MSG_Pass
'                End If
'                .ActiveWorkbook.Close
'            End If
'        Else
'
'        End If
'
'        '    ActiveWindow.WindowState = xlMinimized
'        '    .Sheets(1).Activate
'        '    .Range("A1").Select
'        '    .Visible = True
'        '    .ScreenUpdating = True
'        '    .ActiveWindow.WindowState = xlMaximized
'    End With
'
'    xlsapp.Quit
'    Set xlsapp = Nothing
'
'    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
'                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
'
'End Sub
'
'
'Sub NewKey()
'    '
'
'    Dim nbLigne As Integer
'    Dim T_Cle
'    Dim i As Integer
'
'    Sheets(C_sSheetPassword).Visible = xlSheetHidden
'
'    T_Cle = Sheets(C_sSheetPassword).ListObjects(C_sTabkeys).DataBodyRange
'    nbLigne = UBound(T_Cle, 1)
'
'    Randomize
'    i = Int(nbLigne * Rnd())
'    Sheets(C_sSheetPassword).Range(C_sRngPublickey).value = T_Cle(i, 1)
'    Sheets(C_sSheetPassword).Range(C_sRngPrivatekey).value = T_Cle(i, 2)
'
'    MsgBox "My new password : " & T_Cle(i, 2)    'MSG_NewPass
'
'    Sheets(C_sSheetPassword).Visible = xlSheetVeryHidden
'
'End Sub
'
'Function LetKey(bPriv As Boolean) As Long
'
'    If bPriv Then
'        LetKey = Sheets(C_sSheetPassword).Range("PrivateKey").value
'    Else
'        LetKey = Sheets(C_sSheetPassword).Range("PublicKey").value
'    End If
'
'End Function
'
'
'
