Attribute VB_Name = "DesignerHelpersFunctions"
Option Explicit

Sub TransferCode(xlsapp As Application, sNameModule As String, sType As String)

    Dim oNouvM As Object 'New module name
    Dim sNouvCode As String 'New module code

    'get all the values within the actual module to transfer
    With ThisWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With
    
    'create to code or module if needed
    Select Case sType
    Case "Module"
        Set oNouvM = xlsapp.ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    Case "Class"
        Set oNouvM = xlsapp.ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    End Select

    'keep the name and add the codes
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


Function DesLetColor(sColorCode As String)

    Select Case sColorCode
    Case "BlueEpi"
        DesLetColor = RGB(45, 85, 158)
    Case "RedEpi"
        DesLetColor = RGB(252, 228, 214)
    Case "LightBlueTitle"
        DesLetColor = RGB(217, 225, 242)
    Case "DarkBlueTitle"
        DesLetColor = RGB(142, 169, 219)
    Case "Grey"
        DesLetColor = RGB(235, 232, 232)
    Case "Green"
        DesLetColor = RGB(198, 224, 180)
    Case "Orange"
        DesLetColor = RGB(248, 203, 173)
    Case "White"
        DesLetColor = RGB(255, 255, 255)
    End Select

End Function


