Attribute VB_Name = "M_Migration"
Option Explicit

Const C_StartLineTitle1 As Byte = 3
Const C_StartLineTitle2 As Byte = 4
Const C_TitleLine As Byte = 5

Sub clicExportMigration()

    Dim xlsapp As New Excel.Application

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    Dim T_title
    Dim T_dataLL
    Dim T_dico

    Dim T_Histo
    Dim T_HistoF

    Dim T_admin

    Dim D_title As Scripting.Dictionary
    Dim bSheetAdminExist As Boolean

    Dim sSheetName As String

    With xlsapp
        .Visible = True
        .ScreenUpdating = True
        .Workbook.Add
    
        'dico
        .Sheets.Add.Name = "Dico"
        T_dico = CopyDico                        'M_Export
    
        If Not IsEmptyTable(T_dico) Then
            .Sheets("Dico").Range("A1").Resize(UBound(T_dico, 1), UBound(T_dico, 2)) = T_dico
        End If

        'histo
        .Sheets.Add.Name = "Histo"
        T_Histo = Sheets("geo").[T_HistoGeo]
        T_HistoF = Sheets("geo").[T_HistoFacil]

        If Not IsEmptyTable(T_Histo) Then
            Sheets("histo").Range("A1").Resize(UBound(T_Histo, 1), UBound(T_Histo, 2)) = T_Histo
        End If
        If Not IsEmptyTable(T_HistoF) Then
            Sheets("histo").Range("D1").Resize(UBound(T_HistoF, 1), UBound(T_HistoF, 2)) = T_HistoF
        End If
    
        'admin
        i = 0
        j = 0
        sSheetName = ""
        Set D_title = createDicoTitle
        While i <= UBound(T_dico, 1)
            If LCase(T_dico(i, D_title("sheet"))) = "admin" Then
                If Not bSheetAdminExist Then
                    .Sheets.Add.Name = "Admin"
                
                    ReDim T_admin(2, 0)
                
                    bSheetAdminExist = True
                End If
                ReDim Preserve T_admin(2, i)
                T_admin(0, j) = T_dico(i, D_title("Variable name")) 'cle
                T_admin(1, j) = T_dico(i, D_title("Main label")) 'lib
                T_admin(2, j) = Sheets("Admin").Range(T_dico(i, D_title("Variable name"))) 'value
                j = j + 1
        
            Else                                 'on est pas sur la feuille admin
                'sSheetName =
        
            
        
        
            End If
    
            i = i + 1
        Wend

        If Not IsEmptyTable(T_HistoF) Then
            .Sheets("admin").Range("A1").Resize(UBound(T_admin, 1), UBound(T_admin, 2)) = T_admin
        End If

    End With

End Sub

Private Function createDicoTitle() As Scripting.Dictionary

    Dim i As Integer
    Dim D_title As New Scripting.Dictionary

    i = 1
    While i <= Sheets("dico").Cells(1, 1).End(xlToRight).Column
        D_title.Add Sheets("dico").Cells(1, i).value, i
        i = i + 1
    Wend
    Set createDicoTitle = D_title
    Set D_title = Nothing

End Function

