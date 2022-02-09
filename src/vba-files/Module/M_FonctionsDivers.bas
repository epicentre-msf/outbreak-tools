Attribute VB_Name = "M_FonctionsDivers"
Option Explicit

Sub TransferCode(xlsApp As Application, sNameModule As String)

Dim oNouvM As Object
Dim sNouvCode As String

    With ThisWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With
    
    Select Case Left(sNameModule, 1)
    Case "M"
        Set oNouvM = xlsApp.ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    Case "C"
        Set oNouvM = xlsApp.ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    End Select
    oNouvM.Name = sNameModule
    With xlsApp.ActiveWorkbook.VBProject.VBComponents(oNouvM.Name).CodeModule
        .DeleteLines 1, .CountOfLines
        DoEvents
        .AddFromString sNouvCode
    End With

End Sub

Sub TransferCodeWks(xlsApp As Excel.Application, sSheetName As String)

Dim sCodeName As String

'sCodeName = CleanSpecLettersInName(CStr(sSheetName))

'With xlsApp.Worksheets(xlsApp.Sheets(sSheetName).Index)
'    .Parent.VBProject.VBComponents(.CodeName) _
'            .Properties("_CodeName") = sCodeName
'End With

'With xlsApp.ActiveWorkbook.VBProject.VBComponents(sCodeName).CodeModule
With xlsApp.ActiveWorkbook.VBProject.VBComponents(17).CodeModule
'With xlsApp.ActiveWorkbook.VBProject.VBComponents("feuil6").CodeModule

    .InsertLines 1, "Private Sub Worksheet_Change(ByVal Target As Range)"
    '.InsertLines 2, "call EventFeuille" & sCodeName & "(target)"
    .InsertLines 2, "call EventSheetLineListPatient(target)"
    .InsertLines 3, "End Sub"
End With

End Sub

Sub TransferForm(xlsApp As Application, sFormName As String)

DoEvents

ThisWorkbook.VBProject.VBComponents(sFormName).Export "C:\LineListeApp\CopieUsf.frm"
xlsApp.ActiveWorkbook.VBProject.VBComponents.Import "C:\LineListeApp\CopieUsf.frm"

'ThisWorkbook.VBProject.VBComponents(sFormName).Export ThisWorkbook.Path & "/CopieUsf.frm"
'xlsApp.ActiveWorkbook.VBProject.VBComponents.Import ThisWorkbook.Path & "/CopieUsf.frm"

DoEvents

Kill "C:\LineListeApp\CopieUsf.frm"
Kill "C:\LineListeApp\CopieUsf.frx"

'Kill ThisWorkbook.Path & "/CopieUsf.frm"
'Kill ThisWorkbook.Path & "/CopieUsf.frx"

End Sub

'Sub AddOpeningEvent(xlsApp As Excel.Application)
'
'    With xlsApp.ActiveWorkbook.VBProject.VBComponents(xlsApp.ActiveWorkbook.CodeName).CodeModule
'        .InsertLines Line:=.CreateEventProc("Open", "Workbook") + 1, _
'        String:=vbCrLf & "call OuGreenure"
'    End With
'
'End Sub

Function LetColor(sColorCode As String)

Select Case sColorCode
Case "BlueEpi"
    LetColor = RGB(45, 85, 158)
Case "RedEpi"
    LetColor = RGB(240, 64, 66)
Case "LightBlueTitle"
    LetColor = RGB(217, 225, 242)
Case "DarkBlueTitle"
    LetColor = RGB(142, 169, 219)
Case "Grey"
    LetColor = RGB(128, 128, 128)
Case "Green"
    LetColor = RGB(198, 224, 180)
Case "Orange"
    LetColor = RGB(248, 203, 173)
End Select

End Function
