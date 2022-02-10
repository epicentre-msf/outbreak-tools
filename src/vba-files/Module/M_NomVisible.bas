Attribute VB_Name = "M_NomVisible"
Option Explicit

Const C_PWD As String = "1234"
Const C_TitleCol As Byte = 5
Public bLockActu As Boolean

Function CreateDicTitle() As Scripting.Dictionary

    Dim i As Integer
    Dim D_temp As New Scripting.Dictionary

    D_temp.RemoveAll
    i = 1
    While i <= Sheets("Dico").Cells(1, 1).End(xlToRight).Column

        D_temp.Add Sheets("Dico").Cells(1, i).value, i
        i = i + 1
    Wend
    If Not D_temp.Exists("Champ Visible") Then
        D_temp.Add "Champ Visible", i
    End If
    Set CreateDicTitle = D_temp

    Set D_temp = Nothing

End Function

Sub ClicCmdVisibleName()

    Dim i As Integer
    Dim j As Integer
    Dim T_dataName
    Dim sSheetActive As String
    Dim D_title As New Scripting.Dictionary

    ActiveSheet.Unprotect (C_PWD)

    sSheetActive = ActiveSheet.Name
    Set D_title = CreateDicTitle

    Sheets("Dico").Cells(1, D_title.Count).value = "Champ Visible"

    'chargement de la liste
    'ReDim T_dataName(Sheets("Dico").Cells(1, 1).End(xlDown).Row - 1, 2)
    ReDim T_dataName(2, 0)

    i = 2
    j = 0
    While i <= Sheets("Dico").Cells(1, 1).End(xlDown).Row
        If sSheetActive = Sheets("Dico").Cells(i, D_title("Sheet")) Then
            If LCase(Sheets("Dico").Cells(i, D_title("Status")).value) <> "hidden" Then
                ReDim Preserve T_dataName(2, j)
                T_dataName(0, j) = Sheets("Dico").Cells(i, D_title("Main label")).value
                T_dataName(1, j) = Sheets("Dico").Cells(i, D_title("Variable name")).value
                If LCase(Sheets("Dico").Cells(i, D_title("Status")).value) = "mandatory" Then
                    T_dataName(2, j) = "Mandatory"
                ElseIf LCase(Sheets("Dico").Cells(i, D_title("Champ Visible")).value) = "" Then
                    T_dataName(2, j) = "Shown"
                ElseIf LCase(Sheets("Dico").Cells(i, D_title("Champ Visible")).value) = "0" Then
                    T_dataName(2, j) = ""
                End If
                j = j + 1
            End If
        End If
        i = i + 1
    Wend
    Application.EnableEvents = False
    F_NomVisible.LST_NomChamp.List = Application.Transpose(T_dataName)
    Application.EnableEvents = True

    F_NomVisible.Width = 427
    F_NomVisible.Height = 269

    F_NomVisible.CMD_Fermer.SetFocus
    F_NomVisible.Show

    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

End Sub

Sub IsVisibleDataName(sDataName As String)

    Dim i As Integer
    Dim j As Integer
    Dim D_title As Scripting.Dictionary
    Dim sSheetActive As String

    bLockActu = True
    sSheetActive = ActiveSheet.Name
    Set D_title = CreateDicTitle

    i = 2
    While i <= Sheets("Dico").Cells(1, 1).End(xlDown).Row And Sheets("Dico").Cells(i, D_title("Variable name")).value <> sDataName
        i = i + 1
    Wend
    If Sheets("Dico").Cells(i, D_title("Variable name")).value = sDataName Then
        If LCase(Sheets("Dico").Cells(i, D_title("Status")).value) <> "mandatory" Then
            F_NomVisible.OPT_Masque.Visible = True
            Select Case Sheets("Dico").Cells(i, D_title("Champ Visible")).value
            Case "0"                             'apparait masqué
                F_NomVisible.OPT_Affiche.value = 0
                F_NomVisible.OPT_Masque.value = 1
            Case Else                            'apparait visible
                F_NomVisible.OPT_Affiche.value = 1
                F_NomVisible.OPT_Masque.value = 0
            End Select
        Else
            F_NomVisible.OPT_Masque.Visible = False
        End If
    End If
    Set D_title = Nothing
    bLockActu = False

End Sub

Sub ShowDataCol(sWording As String)

    Dim i As Integer
    Dim j As Integer
    Dim iLineNumber As Integer
    Dim sSheetActive As String
    Dim T_dataName
    Dim D_title As Scripting.Dictionary


    bLockActu = True
    ActiveSheet.Unprotect (C_PWD)
    sSheetActive = ActiveSheet.Name
    Application.ScreenUpdating = False
    Set D_title = CreateDicTitle
    With Sheets(sSheetActive)
    
        iLineNumber = 1
        While iLineNumber <= Sheets("dico").Cells(1, 1).End(xlDown).Row And Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value <> sWording
            iLineNumber = iLineNumber + 1
        Wend
        If Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value = sWording Then
            On Error Resume Next
            i = 1
            While i <= .Cells(C_TitleCol, 16320).End(xlToLeft).Column + 1 And .Cells(C_TitleCol, i).Name.Name <> Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value
                i = i + 1
            Wend
            On Error GoTo 0
            If Sheets(sSheetActive).Cells(C_TitleCol, i).Name.Name = Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value Then
                .Columns(i).Hidden = False
                Sheets("dico").Cells(iLineNumber, D_title("Champ Visible")).value = ""
                If LCase(Sheets("Dico").Cells(iLineNumber, D_title("Control")).value) = "geo" Then
                    .Columns(i + 1).Hidden = False
                    .Columns(i + 2).Hidden = False
                    .Columns(i + 3).Hidden = False
                End If
            End If
        End If
    End With
    '
    ReDim T_dataName(2, 0)
    i = 2
    j = 0
    While i <= Sheets("Dico").Cells(1, 1).End(xlDown).Row
        If sSheetActive = Sheets("Dico").Cells(i, D_title("Sheet")).value Then
            ReDim Preserve T_dataName(2, j)
            T_dataName(0, j) = Sheets("Dico").Cells(i, D_title("Main label"))
            T_dataName(1, j) = Sheets("Dico").Cells(i, D_title("Variable name"))
            If LCase(Sheets("Dico").Cells(i, D_title("Status")).value) = "mandatory" Then
                T_dataName(2, j) = "Mandatory"
            ElseIf Sheets("Dico").Cells(i, D_title("Champ Visible")).value = "0" Then
                T_dataName(2, j) = ""
            ElseIf Sheets("Dico").Cells(i, D_title("Champ Visible")).value = "" Then
                T_dataName(2, j) = "Shown"
            End If
            j = j + 1
        End If
        i = i + 1
    Wend

    F_NomVisible.LST_NomChamp.List = Application.Transpose(T_dataName)
    Application.ScreenUpdating = True
    Set D_title = Nothing
    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    bLockActu = False

End Sub

Sub HideDataCol(sWording As String)

    Dim i As Integer
    Dim j As Integer
    Dim iLineNumber As Integer
    Dim sSheetActive As String
    Dim T_dataName
    Dim D_title As Scripting.Dictionary

    bLockActu = True
    ActiveSheet.Unprotect (C_PWD)
    sSheetActive = ActiveSheet.Name
    Set D_title = CreateDicTitle
    Application.ScreenUpdating = False
    With Sheets(sSheetActive)
        iLineNumber = 1
        While iLineNumber <= Sheets("dico").Cells(1, 1).End(xlDown).Row And Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value <> sWording
            iLineNumber = iLineNumber + 1
        Wend
        If Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value = sWording Then
            i = 1
            On Error Resume Next
            While i <= .Cells(C_TitleCol, 16320).End(xlToLeft).Column + 1 And .Cells(C_TitleCol, i).Name.Name <> Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value
                i = i + 1
            Wend
            On Error GoTo 0
            If Sheets(sSheetActive).Cells(C_TitleCol, i).Name.Name = Sheets("dico").Cells(iLineNumber, D_title("Variable name")).value Then
                Sheets(sSheetActive).Columns(i).Hidden = True
                Sheets("dico").Cells(iLineNumber, D_title("Champ Visible")).value = "0"
                If LCase(Sheets("Dico").Cells(iLineNumber, D_title("Control")).value) = "geo" Then
                    .Columns(i + 1).Hidden = True
                    .Columns(i + 2).Hidden = True
                    .Columns(i + 3).Hidden = True
                End If
            End If
        End If
    End With

    '
    ReDim T_dataName(2, 0)
    i = 2
    j = 0
    While i <= Sheets("Dico").Cells(1, 1).End(xlDown).Row
        If sSheetActive = Sheets("Dico").Cells(i, D_title("Sheet")).value Then
            ReDim Preserve T_dataName(2, j)
            T_dataName(0, j) = Sheets("Dico").Cells(i, D_title("Main label"))
            T_dataName(1, j) = Sheets("Dico").Cells(i, D_title("Variable name"))
            If LCase(Sheets("Dico").Cells(i, D_title("Status")).value) = "mandatory" Then
                T_dataName(2, j) = "Mandatory"
            ElseIf Sheets("Dico").Cells(i, D_title("Champ Visible")).value = "0" Then
                T_dataName(2, j) = ""
            ElseIf Sheets("Dico").Cells(i, D_title("Champ Visible")).value = "" Then
                T_dataName(2, j) = "Shown"
            End If
            j = j + 1
        End If
        i = i + 1
    Wend

    F_NomVisible.LST_NomChamp.List = Application.Transpose(T_dataName)
    Application.ScreenUpdating = True

    Set D_title = Nothing
    ActiveSheet.Protect Password:=C_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                                                                           , AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    bLockActu = False

End Sub


