Attribute VB_Name = "FormLogicExport"

'@Folder("DataIO")
'@ModuleDescription("Setup and teardown for the general export form")
'@depends ExportButton, ILLExport, LLExport, ITranslationObject

Option Explicit

' Collection of ExportButton instances (kept alive for WithEvents)
Private buttons As Collection


' @description Initialize the general export form.
' Creates one ExportButton per configured export on the form.
' @param frm Object. The export form (F_Export).
' @param sourceWkb Workbook. The linelist workbook.
' @param trads ITranslationObject. Translations for messages.
' @param numberOfExports Long. Total configured exports.
Public Sub SetupExportForm(ByVal frm As Object, ByVal sourceWkb As Workbook, _
                           ByVal trads As ITranslationObject, _
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
