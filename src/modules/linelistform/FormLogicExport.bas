Attribute VB_Name = "FormLogicExport"

'@Folder("Linelist Forms")
'@ModuleDescription("Complete code-behind of F_Export -- key commands and export button setup")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@depends ExportButton, LLExport, TranslationObject, LLTranslation, Passwords

' This module is the complete code-behind of the F_Export form and is copied
' into the form at deployment. The numbered export buttons are added and
' bound by EventsLinelistButtons.ClickExport before the form shows, and
' ExportButton answers their clicks, so the callbacks here are the three
' fixed commands: the two key commands and the back command.

Option Explicit

Private Const LLSHEET As String = "LinelistTranslation"
Private Const PASSWORDSHEET As String = "__pass"

Private tradform As TranslationObject 'Translation of forms
Private tradmess As TranslationObject 'Translation of messages
Private pass As Passwords

' Collection of ExportButton instances (kept alive for WithEvents)
Private buttons As Collection


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

'Leave the export form
Private Sub CMD_Back_Click()
    Me.Hide
End Sub

'Translate the form. The sizes come from ClickExport, which positions the
'controls around the buttons it adds.
Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
End Sub


'@section Export button setup
'===============================================================================

' @description Initialize the general export form.
' Creates one ExportButton per configured export on the form.
' @param frm Object. The export form (F_Export).
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
' @param numberOfExports Long. Total configured exports.
Public Sub SetupExportForm(ByVal frm As Object, ByVal sourceWkb As Workbook, _
                           ByVal trads As TranslationObject, _
                           ByVal numberOfExports As Long)

    Dim btn As ExportButton
    Dim cmdBtn As MSForms.CommandButton
    Dim chkBtn As MSForms.CheckBox
    Dim counter As Long
    Dim btnName As String
    Dim chkName As String

    Set buttons = New Collection

    For counter = 1 To numberOfExports
        btnName = "CMDExport" & CStr(counter)
        chkName = "CHKFilter" & CStr(counter)

        Set cmdBtn = Nothing
        Set chkBtn = Nothing

        On Error Resume Next
        Set cmdBtn = frm.Controls(btnName)
        Set chkBtn = frm.Controls(chkName)
        On Error GoTo 0

        If Not cmdBtn Is Nothing Then
            Set btn = ExportButton.Create(sourceWkb, trads, cmdBtn, chkBtn)
            buttons.Add btn
        End If
    Next counter
End Sub

' @description Tear down export form state.
' Releases all ExportButton instances and their event bindings.
Public Sub TeardownExportForm()
    Set buttons = Nothing
End Sub
