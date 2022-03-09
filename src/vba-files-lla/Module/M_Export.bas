Attribute VB_Name = "M_Export"
Option Explicit

Const C_TitleCol As Byte = 1
Const C_TitleSource As Byte = 5
Const C_PWD As String = "1234"

'on va avoir besoin de CreateDicTitle dans M_LineList
'on fonctionne par exclusion

Private Function creationTabChamp(iTypeExport As Byte, sSheetname As String)

    Dim i As Integer
    Dim j As Integer
    Dim D_title As Scripting.Dictionary
    Dim T_data
    Dim sDataNameExport As String
    Dim k As Byte                                'pour la geo

    Set D_title = CreateDicTitle

    sDataNameExport = "export " & iTypeExport

    ReDim T_data(1, 0)
    j = 0
    i = 1
    While i <= Sheets("Dico").Cells(1, 1).End(xlDown).Row
        If LCase(Sheets("Dico").Cells(i, D_title("Export " & iTypeExport)).value) = "yes" And Sheets("Dico").Cells(i, D_title("Sheet")).value = sSheetname Then
            ReDim Preserve T_data(1, j)
            T_data(0, j) = LetPosDataName(Sheets("Dico").Cells(i, D_title("Variable name")).value)
            'T_data(1, j) = Sheets("Dico").Cells(i, D_Title("Variable name")).value
            'If LCase(Sheets("Dico").Cells(i, D_Title("Sheet")).value) <> "admin" Then
            If LCase(Sheets("Dico").Cells(i, D_title("Control")).value) = "geo" Then
                T_data(1, j) = "adm1_" & Sheets("Dico").Cells(i, D_title("Variable name")).value
            
                k = 1
                While k <= 3
                    ReDim Preserve T_data(1, j + k)
                    T_data(0, j + k) = T_data(0, j) + k
                    T_data(1, j + k) = "adm" & k + 1 & "_" & Sheets("Dico").Cells(i, D_title("Variable name")).value
                    k = k + 1
                Wend
                j = j + k
                '        ElseIf LCase(Sheets("Dico").Cells(i, D_Title("Control")).value) = "custom" Then
                '            T_data(1, j) = LetCustomWording(Sheets("Dico").Cells(i, D_Title("Variable name")).value)
                '            j = j + 1
            Else
                T_data(1, j) = Sheets("Dico").Cells(i, D_title("Variable name")).value
                j = j + 1
            End If
            '        Else
            '            T_data(1, j) = "admin|" & Sheets("Dico").Cells(i, D_Title("Variable name")).value
            '
            '        End If
        End If
        i = i + 1
    Wend
    creationTabChamp = T_data
    Set D_title = Nothing

End Function

'Private Function LetCustomWording(sDataName As String) As String
'
'Dim i As Integer
'
'i = 1
'While i <= ActiveSheet.Cells(C_TitleSource, 1).End(xlToRight).Column And LCase(ActiveSheet.Cells(C_TitleSource, i).Name.Name) <> LCase(sDataName)
'    i = i + 1
'Wend
'If LCase(ActiveSheet.Cells(C_TitleSource, i).Name.Name) = LCase(sDataName) Then
'    LetCustomWording = ActiveSheet.Cells(C_TitleSource, i).value
'End If
'
'End Function

Private Function ReplaceCustomDico(sDataName As String, sWording As String)

    Dim i As Integer

    On Error GoTo fin
    i = 1
    While i <= ActiveSheet.Cells(C_TitleSource, 1).End(xlToRight).Column And LCase(ActiveSheet.Cells(C_TitleSource, i).Name.Name) <> LCase(sDataName)
        i = i + 1
    Wend
    If LCase(ActiveSheet.Cells(C_TitleSource, i).Name.Name) = LCase(sDataName) Then
        If InStr(1, ActiveSheet.Cells(C_TitleSource, i).value, sWording) > 0 Then
            ReplaceCustomDico = sWording
        Else
            ReplaceCustomDico = ActiveSheet.Cells(C_TitleSource, i).value
        End If
    End If

fin:

End Function

Private Function LetPosDataName(sDataName As String) As Integer

    Dim i As Integer

    On Error GoTo fin:
    i = 1
    While i <= ActiveSheet.Cells(C_TitleSource, 16320).End(xlToLeft).Column + 1 And ActiveSheet.Cells(5, i).Name.Name <> sDataName
        i = i + 1
    Wend
    If ActiveSheet.Cells(C_TitleSource, i).Name.Name = sDataName Then
        LetPosDataName = i
    End If

fin:

End Function

Sub Export(iTypeExport As Byte)

    Dim i As Integer
    Dim j As Integer
    Dim xlsapp As New Excel.Application
    Dim T_data
    Dim T_dataValid
    Dim sNameListO As String
    Dim sSheetname As String

    Dim T_dico
    Dim D_dico As Scripting.Dictionary

    Dim sPath As String
    Dim sDirectory As String
    Dim T_Path

    Dim diaFolder As FileDialog
    Dim new_path As String

    Dim oCell As Object

    ActiveSheet.Unprotect (C_PWD)
    With xlsapp
        .ScreenUpdating = False
        .Visible = False
        .Workbooks.Add
        sSheetname = ActiveSheet.Name
        .Sheets(1).Name = sSheetname
        'pour la feuille a exporter
        If IsValidSheetForExport(ActiveSheet.Name) Then
            T_dataValid = creationTabChamp(iTypeExport, sSheetname)
            If Not IsEmptyTable(T_dataValid) Then
                sNameListO = "o" & Replace(sSheetname, "-", "_")
                i = 1
                While i <= UBound(T_dataValid, 2)
                    T_data = ActiveSheet.ListObjects("olinelist-patient").ListColumns(T_dataValid(0, i)).Range
                    .Sheets(sSheetname).Cells(C_TitleCol, i).Resize(UBound(T_data)) = T_data
                    .Sheets(sSheetname).Cells(C_TitleCol, i).value = T_dataValid(1, i)
           
                    i = i + 1
                Wend
            End If
        End If
    
        'pour le dico
    
        .Sheets.Add.Name = "Dico"
        T_dico = CopyDico
    
        'Set D_dico = CreateDicoName
        i = 1
        While i <= UBound(T_dico, 1)

            T_dico(i, 1) = ReplaceCustomDico(CStr(T_dico(i, 0)), CStr(T_dico(i, 1)))

            i = i + 1
        Wend
    
        If Not IsEmptyTable(T_dico) Then
            .Sheets("Dico").Range("A1").Resize(UBound(T_dico, 1), UBound(T_dico, 2)) = T_dico
        End If
    
        'l'admin
        Erase T_dataValid
        T_dataValid = creationTabChamp(iTypeExport, "admin")
        .Sheets.Add.Name = "Admin"
        i = 0
        j = 1
        While i <= UBound(T_dataValid, 2)
            .Sheets("Admin").Cells(j, 2).Name = T_dataValid(1, i)
            .Sheets("Admin").Cells(j, 1).value = Range(T_dataValid(1, i)).Offset(, -1).value
            .Sheets("Admin").Cells(j, 2).value = Range(T_dataValid(1, i)).value
            j = j + 1
            i = i + 1
        Wend
    
        'pour l'enregistrement
        sPath = Sheets("Export").Cells(iTypeExport + 1, 5).value
        If sPath <> "" Then
            T_Path = Split(sPath, "+")
        
            Set D_dico = CreateDicoName
            i = 0
            While i <= UBound(T_Path)
                If InStr(1, T_Path(i), Chr(34)) = 0 Then
                    If D_dico.Exists(Trim(T_Path(i))) Then
                        sPath = Replace(sPath, Trim(T_Path(i)), Range(Trim(T_Path(i))).value)
                    End If
                End If
                i = i + 1
            Wend
            Set D_dico = Nothing
            sPath = Replace(Replace(Replace(sPath & "__" & Range("RNG_PublicKey").value & "__" & Format(Now, "yyyymmdd-HhNn"), " ", ""), "+", "__"), Chr(34), "")
            sDirectory = LoadFolder
            If sDirectory <> "" Then
                sPath = sDirectory & Application.PathSeparator & sPath & ".xlsb"
            
                i = 0
                While Len(sPath) >= 255 And i < 3 'MSG_PathTooLong
                    MsgBox "The path of the export file is too long so the file name gets truncated. Please select a folder higher in the hierarchy to save the export (ex: Desktop, Downloads, Documents etc.)"
                    sDirectory = LoadFolder
                    If sDirectory <> "" Then
                        sPath = sDirectory & Application.PathSeparator & sPath
                    End If
                    i = i + 1
                Wend
                'on enregistre
                If i < 3 Then
                    .ActiveWorkbook.SaveAs Filename:=sPath, FileFormat:=xlExcel12, CreateBackup:=False, Password:=Range("RNG_PrivateKey").value
                    MsgBox "File saved" & Chr(10) & "Password : " & Range("RNG_PrivateKey").value 'MSG_FileSaved        'MSG_Pass
                End If
                .ActiveWorkbook.Close
            End If
        Else
    
        End If
    
        '    ActiveWindow.WindowState = xlMinimized
        '    .Sheets(1).Activate
        '    .Range("A1").Select
        '    .Visible = True
        '    .ScreenUpdating = True
        '    .ActiveWindow.WindowState = xlMaximized
    End With

    xlsapp.Quit
    Set xlsapp = Nothing
        
    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

End Sub

Private Function CreateDicoName() As Scripting.Dictionary

    Dim i As Integer
    Dim D_dic As New Scripting.Dictionary
    Dim D_dicTitle As Scripting.Dictionary

    Set D_dicTitle = CreateDicTitle

    i = 2
    D_dic.RemoveAll
    While i <= Sheets("dico").Cells(1, 1).End(xlDown).Row
        D_dic.Add Sheets("dico").Cells(i, D_dicTitle("Variable name")).value, Sheets("dico").Cells(i, D_dicTitle("Control")).value
        i = i + 1
    Wend
    Set CreateDicoName = D_dic
    Set D_dic = Nothing

End Function

Private Function CreateDicTitle() As Scripting.Dictionary

    Dim i As Integer
    Dim D_dic As New Scripting.Dictionary

    D_dic.RemoveAll
    i = 1
    While i <= Sheets("dico").Cells(1, 1).End(xlToRight).Column
        D_dic.Add Sheets("dico").Cells(1, i).value, i
        i = i + 1
    Wend
    Set CreateDicTitle = D_dic
    Set D_dic = Nothing

End Function

Private Function IsValidSheetForExport(sName As String) As Boolean

    Dim i As Integer
    Dim D_title As Scripting.Dictionary

    IsValidSheetForExport = True
    Set D_title = CreateDicTitle

    i = 2
    While i <= Sheets("dico").Cells(1, 1).End(xlDown).Row
        If Sheets("dico").Cells(i, D_title("Sheet")) = sName Then
            IsValidSheetForExport = True
            Exit Function
        End If
        i = i + 1
    Wend

    Set D_title = Nothing

End Function

Public Function CopyDico()

    Dim i As Integer
    Dim j As Integer
    Dim iMax As Integer
    Dim jMax As Integer
    Dim T_res

    iMax = Sheets("dico").Cells(1, 1).End(xlDown).Row
    jMax = Sheets("dico").Cells(1, 1).End(xlToRight).Column
    ReDim T_res(iMax, jMax)

    i = 1
    While i <= iMax
        j = 1
        While j <= jMax
            T_res(i - 1, j - 1) = Sheets("dico").Cells(i, j).value
            j = j + 1
        Wend
        i = i + 1
    Wend
    CopyDico = T_res

End Function

Sub NewKey()
    '
                    
    Dim nbLigne As Integer
    Dim T_Cle
    Dim i As Integer
                    
    Sheets("PASSWORD").Visible = xlSheetHidden
                    
    T_Cle = [T_Keys]
    nbLigne = UBound(T_Cle, 1)
    
    Randomize
    i = Int(nbLigne * Rnd())
    Sheets("PASSWORD").Range("RNG_PublicKey").value = T_Cle(i, 1)
    Sheets("PASSWORD").Range("RNG_PrivateKey").value = T_Cle(i, 2)
    
    MsgBox "My new password : " & T_Cle(i, 2)    'MSG_NewPass
    
    Sheets("PASSWORD").Visible = xlSheetVeryHidden
    
End Sub

Function LetKey(bPriv As Boolean) As Long
    
    If bPriv Then
        LetKey = Sheets("PASSWORD").Range("PrivateKey").value
    Else
        LetKey = Sheets("PASSWORD").Range("PublicKey").value
    End If
    
End Function


