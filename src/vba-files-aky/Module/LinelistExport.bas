Attribute VB_Name = "LinelistExport"

Option Explicit
'Preliminary functions for export ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Function ExportPath(iTypeExport As Byte, iFileNameIndex As Integer) As String

    Dim sPath As String
    Dim sDirectory As String
    Dim sSheetName As String
    Dim i As Long                                'iterator

    Dim iSheetNameIndex As Integer
    Dim iVarIndex As Integer

    Dim PathData As BetterArray
    Dim VarNameData As BetterArray
    Dim Headers As BetterArray
    Dim DictSheet As Worksheet

    sPath = vbNullString
    sPath = ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, iFileNameIndex).value

    Set PathData = New BetterArray
    PathData.LowerBound = 1
    Set VarNameData = GetDictionaryColumn(C_sDictHeaderVarName)
    Set DictSheet = ThisWorkbook.Worksheets(C_sParamSheetDict)

    iSheetNameIndex = GetDictionaryIndex(C_sDictHeaderSheetName)

    If sPath <> vbNullString Then

        PathData.Items = Split(sPath, "+")

        i = 1
        While i <= PathData.UpperBound
            PathData.Item(i) = Replace(Application.WorksheetFunction.Trim(PathData.Items(i)), "+", "")

            If VarNameData.Includes(PathData.Items(i)) Then

                iVarIndex = VarNameData.IndexOf(PathData.Items(i)) 'Index of the varname
                sSheetName = DictSheet.Cells(1 + iVarIndex, iSheetNameIndex) '+1 because first line is for column names
                PathData.Item(i) = ThisWorkbook.Worksheets(sSheetName).Range(PathData.Items(i)).value

            Else
                PathData.Item(i) = Replace(PathData.Items(i), Chr(34), "")
            End If

            i = i + 1
        Wend


        sPath = PathData.ToString(Separator:="-", OpeningDelimiter:=vbNullString, ClosingDelimiter:=vbNullString, QuoteStrings:=False) & _
                                                                                                                                       "__" & ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PublicKey").value & "__" & Format(Now, "yyyymmdd-HhNn")

        'Folder where to write the exports
        sDirectory = Helpers.LoadFolder

        If sDirectory <> vbNullString Then
            sPath = sDirectory & Application.PathSeparator & sPath
            i = 0
            While Len(sPath) >= 255 And i < 3

                MsgBox TranslateLLMsg("MSG_PathTooLong")
                sDirectory = LoadFolder

                If sDirectory <> vbNullString Then

                    sPath = sDirectory & Application.PathSeparator & sPath

                End If
                i = i + 1
            Wend

            If i > 3 Then
                sPath = vbNullString
                Exit Function
            End If
        End If
    End If


    ExportPath = sPath
    Set PathData = Nothing
    Set VarNameData = Nothing
    Set DictSheet = Nothing
End Function


Private Function AddExportLLSheet(Wkb As Workbook, sSheetName As String, sPrevSheetName As String, _
                          DictExportData As BetterArray, i As Long, iSheetNameIndex As Integer, _
                          iVarnameIndex As Integer) As Long

    Dim k As Long
    Dim iLastRow As Long

    Dim sVarName As String
    Dim sNewSheetName As String
    Dim src As Range
    Dim dest As Range

    'Add the new worksheet
    Wkb.Worksheets.Add(after:=Wkb.Worksheets(sPrevSheetName)).Name = sSheetName

    k = i
    Do While DictExportData.Items(k, iSheetNameIndex) = sSheetName

        sVarName = DictExportData.Items(k, iVarnameIndex)

        With ThisWorkbook.Worksheets(sSheetName)
            Set src = .ListObjects(SheetListObjectName(sSheetName)).ListColumns(sVarName).Range
        End With

        iLastRow = src.Rows.Count

        With Wkb.Worksheets(sSheetName)
            Set dest = .Range(.Cells(1, k - i + 1), .Cells(iLastRow, k - i + 1))
        End With

        dest.value = src.value

        k = k + 1
        If k > DictExportData.Length Then Exit Do
    Loop

    AddExportLLSheet = k - 1
End Function

Private Function AddExportAdmSheet(Wkb As Workbook, sSheetName As String, sPrevSheetName As String, _
                           DictExportData As BetterArray, i As Long, iSheetNameIndex As Integer, _
                           iVarnameIndex As Integer) As Long

    Dim k As Long
    Dim sVarName As String
    Dim srcWksh As Worksheet


    'Add the new worksheet
    Wkb.Worksheets.Add(after:=Wkb.Worksheets(sPrevSheetName)).Name = sSheetName
    Set srcWksh = ThisWorkbook.Worksheets(sSheetName)

    k = i
    Do While DictExportData.Items(k, iSheetNameIndex) = sSheetName

        sVarName = DictExportData.Items(k, iVarnameIndex)
        With Wkb.Worksheets(sSheetName)
            .Cells(k - i + 2, 2).value = srcWksh.Range(sVarName).value
            .Cells(k - i + 2, 1).value = sVarName
        End With

        k = k + 1
    Loop

    'Add variable and value
    Wkb.Worksheets(sSheetName).Cells(1, 1).value = C_sVariable
    Wkb.Worksheets(sSheetName).Cells(1, 2).value = C_sValue

    AddExportAdmSheet = k - 1
End Function

'Export Function --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Export(iTypeExport As Byte)

    Dim Wkb             As Workbook
    Dim DictExportData  As BetterArray           'Values of the dictionary
    Dim ExportHeader    As BetterArray
    Dim ExportData      As BetterArray
    Dim DictLo          As ListObject

    Dim i As Long                                'Iterator
    Dim istep As Long                            'Step in the loop
    Dim iExportIndex As Integer
    Dim iSheetTypeIndex As Integer
    Dim iSheetNameIndex As Integer
    Dim iVarnameIndex As Integer
    Dim fileformat As Byte

    Dim sPrevSheetName As String
    Dim sSheetName As String
    Dim sPath As String
    Dim sFirstSheetName As String

    Set DictExportData = New BetterArray
    Set ExportData = New BetterArray
    Set ExportHeader = New BetterArray


    On Error GoTo exportErrHandExport

    Set ExportHeader = Helpers.GetHeaders(ThisWorkbook, C_sParamSheetExport, 1)

    'Get Export Path
    sPath = ExportPath(iTypeExport, ExportHeader.IndexOf(C_sExportHeaderFileName))

    'Creating the data for the exports
    On Error GoTo exportErrHandData

    If sPath <> vbNullString Then

        BeginWork xlsapp:=Application

        Set DictLo = ThisWorkbook.Worksheets(C_sParamSheetDict).ListObjects("o" & ClearString(C_sParamSheetDict))

        'Get the type of the export
        Select Case iTypeExport
        Case 1
            iExportIndex = GetDictionaryIndex(C_sDictHeaderExport1)
        Case 2
            iExportIndex = GetDictionaryIndex(C_sDictHeaderExport2)
        Case 3
            iExportIndex = GetDictionaryIndex(C_sDictHeaderExport3)
        Case 4
            iExportIndex = GetDictionaryIndex(C_sDictHeaderExport4)
        Case 5
            iExportIndex = GetDictionaryIndex(C_sDictHeaderExport5)
        Case Else
            iExportIndex = GetDictionaryIndex(C_sDictHeaderExport1)
        End Select

        Set DictExportData = FilterLoTable(DictLo, iExportIndex, C_sYes)

        'Here I have the list of all the variables to Export Just go on

        Set Wkb = Workbooks.Add

        With Wkb

            'Adding the worksheets
            sPrevSheetName = .Worksheets(1).Name
            sFirstSheetName = sPrevSheetName

            'Add Translation
            i = ExportHeader.IndexOf(C_sExportHeaderTranslation)

            If (ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value = C_sYes) Then
                Set ExportData = GetTransData()
                .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetTranslation
                ExportData.ToExcelRange .Worksheets(C_sParamSheetTranslation).Cells(1, 1)
                sPrevSheetName = C_sParamSheetTranslation
                ExportData.Clear
            End If

            i = ExportHeader.IndexOf(C_sExportHeaderMetadata)

            'Add Choice
            If (ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value = C_sYes) Then

                Set ExportData = GetChoicesData()
                'Choices Sheet
                .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetChoices
                ExportData.ToExcelRange .Worksheets(C_sParamSheetChoices).Cells(1, 1)
                sPrevSheetName = C_sParamSheetChoices
                ExportData.Clear

                'Add Dictionary
                .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sParamSheetDict
                ExportData.FromExcelRange DictLo.Range
                ExportData.ToExcelRange .Worksheets(C_sParamSheetDict).Cells(1, 1)
                sPrevSheetName = C_sParamSheetDict
                ExportData.Clear

                'Add Metadata
                .Worksheets.Add(before:=.Worksheets(sPrevSheetName)).Name = C_sSheetMetadata
                ExportData.FromExcelRange ThisWorkbook.Worksheets(C_sSheetMetadata).Cells(1, 1), DetectLastColumn:=True, DetectLastRow:=True
                ExportData.ToExcelRange .Worksheets(C_sSheetMetadata).Cells(1, 1)
                sPrevSheetName = C_sSheetMetadata
                ExportData.Clear
            End If
        End With

        'Adding the others sheets (Admin, linelist)


        'write all the sheets of type linelist
        i = 1
        istep = 1

        iSheetNameIndex = GetDictionaryIndex(C_sDictHeaderSheetName)
        iSheetTypeIndex = GetDictionaryIndex(C_sDictHeaderSheetType)
        iVarnameIndex = GetDictionaryIndex(C_sDictHeaderVarName)

        While i <= DictExportData.Length

            sSheetName = DictExportData.Items(i, iSheetNameIndex)

            Select Case DictExportData.Items(i, iSheetTypeIndex)

            Case C_sDictSheetTypeLL

                istep = AddExportLLSheet(Wkb, sSheetName, sPrevSheetName, DictExportData, i, iSheetNameIndex, iVarnameIndex)

                'You were thingking about using a function to speed the steps
            Case C_sDictSheetTypeAdm

                istep = AddExportAdmSheet(Wkb, sSheetName, sPrevSheetName, DictExportData, i, iSheetNameIndex, iVarnameIndex)

            End Select

            'refresh previous sheet name
            sPrevSheetName = sSheetName

            i = istep + 1
        Wend
    End If

    Wkb.Worksheets(sFirstSheetName).Delete

    'Now writing on the choosen directory
    On Error GoTo exportErrHandWrite

    'Handling the file format
    i = ExportHeader.IndexOf(C_sExportHeaderFileFormat)

    Select Case ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value
    Case "xlsx"
        fileformat = xlOpenXMLWorkbook
    Case "xlsb"
        fileformat = xlExcel12
    Case Else
        fileformat = xlExcel12
    End Select

    'Handling password

    i = ExportHeader.IndexOf(C_sExportHeaderPassword)

    Select Case ClearString(ThisWorkbook.Worksheets(C_sParamSheetExport).Cells(iTypeExport + 1, i).value)
    Case C_sYes
        Wkb.SaveAs Filename:=sPath, fileformat:=fileformat, CreateBackup:=False, Password:=ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value, _
        ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        MsgBox "File saved" & Chr(10) & "Password : " & ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value
    Case C_sNo
        Wkb.SaveAs Filename:=sPath, fileformat:=fileformat, CreateBackup:=False, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        MsgBox "File saved" & Chr(10) & "No Password"
    Case Else
        Wkb.SaveAs Filename:=sPath, fileformat:=fileformat, CreateBackup:=False, Password:=ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value, _
        ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        MsgBox "File saved" & Chr(10) & "Password : " & ThisWorkbook.Worksheets(C_sSheetPassword).Range("RNG_PrivateKey").value
    End Select

    Wkb.Close
    F_Export.Hide
    EndWork xlsapp:=Application

    Set Wkb = Nothing
    Set DictExportData = Nothing
    Set ExportData = Nothing
    Set ExportHeader = Nothing
    Set DictLo = Nothing

    Exit Sub

exportErrHandExport:
        MsgBox TranslateLLMsg("MSG_ErrHandExport"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
        Exit Sub
exportErrHandData:
        MsgBox TranslateLLMsg("MSG_exportErrHandData"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
        Exit Sub
exportErrHandWrite:
        MsgBox TranslateLLMsg("MSG_exportErrHandWrite"), vbOKOnly + vbCritical, TranslateLLMsg("MSG_Error")
        Exit Sub
End Sub

'Password functions ===================================================================================================================================================================================

Sub NewKey()
    Dim nbLigne As Integer
    Dim T_Cle
    Dim i As Integer

    ThisWorkbook.Worksheets(C_sSheetPassword).Visible = xlSheetHidden

    T_Cle = ThisWorkbook.Worksheets(C_sSheetPassword).ListObjects(C_sTabkeys).DataBodyRange
    nbLigne = UBound(T_Cle, 1)

    'Randomize
    i = Int(nbLigne * Rnd())
    ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngPublickey).value = T_Cle(i, 1)
    ThisWorkbook.Worksheets(C_sSheetPassword).Range(C_sRngPrivatekey).value = T_Cle(i, 2)

    MsgBox TranslateLLMsg("MSG_Password") & T_Cle(i, 2) 'MSG_NewPass

    ThisWorkbook.Worksheets(C_sSheetPassword).Visible = xlSheetVeryHidden

End Sub

Function LetKey(bPriv As Boolean) As Long

    If bPriv Then
        LetKey = ThisWorkbook.Worksheets(C_sSheetPassword).Range("PrivateKey").value
    Else
        LetKey = ThisWorkbook.Worksheets(C_sSheetPassword).Range("PublicKey").value
    End If

End Function


