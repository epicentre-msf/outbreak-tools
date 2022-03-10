Attribute VB_Name = "M_FonctionsDivers"
Option Explicit

Sub TransferCode(xlsapp As Application, sNameModule As String, sType As String)

    Dim oNouvM As Object
    Dim sNouvCode As String

    With ThisWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With
    
    Select Case sType
    Case "M"
        Set oNouvM = xlsapp.ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    Case "C"
        Set oNouvM = xlsapp.ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    End Select
    
    oNouvM.Name = sNameModule
    With xlsapp.ActiveWorkbook.VBProject.VBComponents(oNouvM.Name).CodeModule
        .DeleteLines 1, .CountOfLines
        DoEvents
        .AddFromString sNouvCode
    End With

End Sub

'Transfert code from one module to a worksheet to trigger some events
' sSheetName the sheet name we want to transfer to
' sNameModule the name of the module we want to copy code from
Sub TransferCodeWks(xlsapp As Excel.Application, sSheetname As String, sNameModule As String)

    Dim sNouvCode As String                      'a string to contain code to add
    Dim sheetComp As String
    Dim vbProj As Object                         'component, project and modules
    Dim vbComp As Object
    Dim codeMod As Object
    
    'save the code module in the string sNouvCode
    With ThisWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With
    
    With xlsapp
        Set vbProj = .ActiveWorkbook.VBProject
        Set vbComp = vbProj.VBComponents(.Sheets(sSheetname).CodeName)
        Set codeMod = vbComp.CodeModule
    End With
    
    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        DoEvents
        .AddFromString sNouvCode
    End With
    
    'With xlsApp.ActiveWorkbook.VBProject.VBComponents(sCodeName).CodeModule
    'With xlsApp.ActiveWorkbook.VBProject.VBComponents(17).CodeModule
    'With xlsApp.ActiveWorkbook.VBProject.VBComponents("feuil6").CodeModule

    '.InsertLines 1, "Private Sub Worksheet_Change(ByVal Target As Range)"
    '.InsertLines 2, "call EventFeuille" & sCodeName & "(target)"
    '.InsertLines 2, "call EventSheetLineListPatient(target)"
    '.InsertLines 3, "End Sub"
    'End With

End Sub

Sub TransferForm(xlsapp As Application, sFormName As String)
    
    'The form is sent to the C:\LinelisteApp folder
    DoEvents

    ThisWorkbook.VBProject.VBComponents(sFormName).Export "C:\LineListeApp\CopieUsf.frm"
    xlsapp.ActiveWorkbook.VBProject.VBComponents.Import "C:\LineListeApp\CopieUsf.frm"
    
    DoEvents

    Kill "C:\LineListeApp\CopieUsf.frm"
    Kill "C:\LineListeApp\CopieUsf.frx"
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
    Case "White"
        LetColor = RGB(255, 255, 255) 'lla
    End Select

End Function


