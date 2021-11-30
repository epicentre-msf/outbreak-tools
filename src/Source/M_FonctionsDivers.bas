Attribute VB_Name = "M_FonctionsDivers"
Option Explicit

Sub balanceTonCode(xlsApp As Application, sNomModule As String)

Dim oNouvM As Object
Dim sNouvCode As String


    With ThisWorkbook.VBProject.VBComponents(sNomModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With
    
    Set oNouvM = xlsApp.ActiveWorkbook.VBProject.VBComponents.Add(1)
    oNouvM.Name = sNomModule
    With xlsApp.ActiveWorkbook.VBProject.VBComponents(oNouvM.Name).CodeModule
        .DeleteLines 1, .CountOfLines
        DoEvents
        .AddFromString sNouvCode
    End With

End Sub

Sub balanceTonFrm(xlsApp As Application, sNomFormulaire As String)

DoEvents

ThisWorkbook.VBProject.VBComponents(sNomFormulaire).Export "C:\LineListeApp\CopieUsf.frm"
xlsApp.ActiveWorkbook.VBProject.VBComponents.Import "C:\LineListeApp\CopieUsf.frm"

DoEvents

Kill "C:\LineListeApp\CopieUsf.frm"
Kill "C:\LineListeApp\CopieUsf.frx"

End Sub

Sub EcrireEventOuverture(xlsApp As Application)

    With xlsApp.ActiveWorkbook.VBProject.VBComponents(xlsApp.ActiveWorkbook.CodeName).CodeModule
        .InsertLines Line:=.CreateEventProc("Open", "Workbook") + 1, _
        String:=vbCrLf & "call Ouverture"
    End With
    
End Sub

Function retourneCouleur(sCodeCouleur As String)

Select Case sCodeCouleur
Case "BleuEpi"
    retourneCouleur = RGB(45, 85, 158)
Case "RougeEpi"
    retourneCouleur = RGB(240, 64, 66)
Case "BleuClairTitre"
    retourneCouleur = RGB(217, 225, 242)
Case "BleuFonceTitre"
    retourneCouleur = RGB(142, 169, 219)
Case "Gris"
    retourneCouleur = RGB(128, 128, 128)
End Select

End Function

