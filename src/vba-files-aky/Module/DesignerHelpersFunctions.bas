Attribute VB_Name = "DesignerHelpersFunctions"
Option Explicit

Sub DesTransferCode(Wkb As Workbook, sNameModule As String, sType As String)

    Dim oNouvM As Object 'New module name
    Dim sNouvCode As String 'New module code

    'get all the values within the actual module to transfer
    With DesignerWorkbook.VBProject.VBComponents(sNameModule).CodeModule
        sNouvCode = .Lines(1, .CountOfLines)
    End With
    
    'create to code or module if needed
    Select Case sType
    Case "Module"
        Set oNouvM = Wkb.VBProject.VBComponents.Add(vbext_ct_StdModule)
    Case "Class"
        Set oNouvM = Wkb.VBProject.VBComponents.Add(vbext_ct_ClassModule)
    End Select

    'keep the name and add the codes
    oNouvM.Name = sNameModule
    With Wkb.VBProject.VBComponents(oNouvM.Name).CodeModule
        .DeleteLines 1, .CountOfLines
         DoEvents
        .AddFromString sNouvCode
    End With

End Sub

'Transfert code from one module to a worksheet to trigger some events
' sSheetName the sheet name we want to transfer to
' sNameModule the name of the module we want to copy code from
Sub DesTransferCodeWks(xlsapp As Excel.Application, sSheetName As String, sNameModule As String)

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
        Set vbComp = vbProj.VBComponents(.Sheets(sSheetName).CodeName)
        Set codeMod = vbComp.CodeModule
    End With
    
    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        DoEvents
        .AddFromString sNouvCode
    End With

End Sub

Sub DesTransferForm(Wkb As Workbook, sFormName As String)
    
    'The form is sent to the LinelisteApp folder
    On Error Resume Next
    Kill (Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm")
    On Error GoTo 0
    
    DoEvents
    DesignerWorkbook.VBProject.VBComponents(sFormName).Export Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm"
    Wkb.VBProject.VBComponents.Import Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm"
    DoEvents

    Kill (Environ("Temp") & Application.PathSeparator & "LinelistApp" & "CopieUsf.frm")
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
    Case "MainLabBlue"
        DesLetColor = RGB(47, 117, 181)
    Case "SubLabBlue"
        DesLetColor = RGB(221, 235, 247)
    Case "NotesBlue"
        DesLetColor = RGB(142, 169, 219)
    End Select

End Function


