Attribute VB_Name = "FormLogicExport"

'@Folder("Linelist Forms")
'@ModuleDescription("Complete code-behind of F_Export -- key commands and export button setup")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@depends ExportButton, LLExport, TranslationObject, LLTranslation, Passwords, BetterArray

' This module is the complete code-behind of the F_Export form and is copied
' into the form at deployment. EventsLinelistButtons.ClickExport reaches
' SetupExportForm and TeardownExportForm through F_Export, the qualified
' route, so the standard-module copy of this code is never compiled and its
' form references cost nothing there.
'
' The form owns the numbered export buttons. It adds them, it holds the
' ExportButton that answers each click, and it takes them off again when the
' form goes down. The holder is the module-level collection below: a binding
' lives as long as the object holding it, and a collection in the caller of
' Show stays alive only while that caller sits on the stack.

Option Explicit

Private Const LLSHEET As String = "LinelistTranslation"
Private Const PASSWORDSHEET As String = "__pass"
Private Const NAMETAG As String = "CMDExport"
Private Const BUTTON_WIDTH As Single = 160
Private Const BUTTON_LEFT As Single = 20

Private tradform As TranslationObject 'Translation of forms
Private tradmess As TranslationObject 'Translation of messages
Private pass As Passwords

' One ExportButton per export button on the open form. Each holds the
' WithEvents binding of its button, so the button answers a click for as long
' as this collection holds it.
Private exportButtons As Collection


'@section Initialization and control callbacks
'===============================================================================

'Initialize the translation objects and the passwords
Private Sub InitializeTrads()
    Dim lltrads As LLTranslation
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set lltrads = LLTranslation.Create(wb.Worksheets(LLSHEET))
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set pass = Passwords.Create(wb.Worksheets(PASSWORDSHEET))
End Sub

'Generate a new private key for the exports
Private Sub CMD_NewKey_Click()
    InitializeTrads
    pass.GenerateKey tradmess
End Sub

'Show the private key currently in use
Private Sub CMD_ShowKey_Click()
    InitializeTrads
    pass.DisplayPrivateKey tradmess
End Sub

'Leave the export form. Hide keeps the instance and every button on it, and
'Controls.Add refuses a name already in use, so the buttons come off first.
Private Sub CMD_Back_Click()
    TeardownExportForm
    Me.Hide
End Sub

'Translate the form. The sizes come from ClickExport, which positions the
'controls around the buttons SetupExportForm adds.
Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
End Sub

'Take the export buttons off a form the user closes with the window box.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    TeardownExportForm
End Sub


'@section Export button setup
'===============================================================================

' @description Build one export button per active export on the form.
' Each button is named CMDExport<n> after the export it runs, and the
' ExportButton that answers its clicks is kept in the module-level
' collection. The buttons of an earlier open come off first: Me.Hide keeps
' the form instance and everything on it, and Controls.Add refuses a name
' already in use.
'
' Every button reads the one CHK_ExportFiltered box. Two exports never run at
' the same time, so one box carries the answer for the export being run.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
' @param exportList LLExport. The export definitions of the workbook.
' @param topPosition Single. Where the first button goes.
' @param buttonHeight Single. Height of one button.
' @param buttonGap Single. Space between two buttons.
' @return Single. The free position under the last button, for the caller to
' lay the fixed controls out from.
Public Function SetupExportForm(ByVal sourceWkb As Workbook, _
                                ByVal trads As TranslationObject, _
                                ByVal exportList As LLExport, _
                                ByVal topPosition As Single, _
                                ByVal buttonHeight As Single, _
                                ByVal buttonGap As Single) As Single

    Dim activeExports As BetterArray
    Dim counter As Long
    Dim exportNumber As Long
    Dim cmdBtn As MSForms.CommandButton
    Dim nextTop As Single

    TeardownExportForm

    Set exportButtons = New Collection
    Set activeExports = exportList.ActiveExportNumbers()
    nextTop = topPosition

    For counter = activeExports.LowerBound To activeExports.UpperBound
        exportNumber = activeExports.Item(counter)

        Set cmdBtn = Me.Controls.Add("Forms.CommandButton.1", _
                                     NAMETAG & CStr(exportNumber), True)
        cmdBtn.Caption = exportList.ColumnValue(exportNumber, "label button")
        cmdBtn.Top = nextTop
        cmdBtn.Height = buttonHeight
        cmdBtn.Width = BUTTON_WIDTH
        cmdBtn.Left = BUTTON_LEFT
        cmdBtn.WordWrap = True

        exportButtons.Add ExportButton.Create(sourceWkb, trads, cmdBtn, _
                                              Me.CHK_ExportFiltered)

        nextTop = nextTop + buttonHeight + buttonGap
    Next counter

    SetupExportForm = nextTop
End Function

' @description Take the export buttons off the form and drop what listens to
' them. The bindings go first, so a control comes off once nothing is
' listening to it. The sweep then walks the controls of the form and matches
' their names, which also clears a form a failed setup left half built.
Public Sub TeardownExportForm()
    Dim counter As Long
    Dim controlName As String

    Set exportButtons = Nothing

    For counter = Me.Controls.Count - 1 To 0 Step -1
        controlName = Me.Controls(counter).Name
        If StrComp(Left$(controlName, Len(NAMETAG)), NAMETAG, vbTextCompare) = 0 Then
            'A control the form was built with stays on it, and Remove says so.
            On Error Resume Next
            Me.Controls.Remove counter
            On Error GoTo 0
        End If
    Next counter
End Sub
