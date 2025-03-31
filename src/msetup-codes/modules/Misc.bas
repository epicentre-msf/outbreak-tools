Attribute VB_Name="Misc"

Option Explicit

'@IgnoreModule EmptyMethod

Public Sub TransferCodeWksh(ByVal sheetName As String)

   Const CHANGEMODULENAME As String = "EventsSheetChange"
   Const WBMODULENAME As String = "EventsWorkbook"

   Dim codeContent As String                    'a string to contain code to add
   Dim vbProj As Object                         'component, project and modules
   Dim vbComp As Object
   Dim codeMod As Object
   Dim modName As String
   Dim currwb As Workbook

   Set currwb = ThisWorkbook

    modName = IIf(sheetName = "__WorkbookLevel", WBMODULENAME, CHANGEMODULENAME)
    'save the code module in the string sNouvCode
    With currwb.VBProject.VBComponents(modName).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With
    With currwb
        Set vbProj = .VBProject
        'The component could be the workbook code name for workbook related transfers
        If sheetName = "__WorkbookLevel" Then
            Set vbComp = vbProj.VBComponents(.codeName)
        Else
            Set vbComp = vbProj.VBComponents(.sheets(sheetName).codeName)
        End If
        Set codeMod = vbComp.CodeModule
    End With
    'Adding the code
    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub




Public Sub Compare()

End Sub