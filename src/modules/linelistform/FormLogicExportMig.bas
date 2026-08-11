Attribute VB_Name = "FormLogicExportMig"

'@Folder("Linelist Forms")
'@ModuleDescription("Migration, analysis, geo and other-linelist export workflows")
'@depends LLExporter, ApplicationState, OSFiles, TranslationObject

' The handlers of the F_ExportMig form. The form's own event stubs stay thin:
' CMD_ExportMig_Click calls HandleExportClick with the form, the workbook and
' the message translations; CHK_OtherLinelist_Click calls
' HandleOtherLinelistChecked when the box goes on and HandleOtherLinelistUnchecked
' followed by tradform.TranslateForm when it goes off, which is what puts the
' hint captions back on the two labels; LBL_OtherPath_Click and
' LBL_OtherPass_Click call PromptOtherLinelistPath and
' PromptOtherLinelistPassword; CMD_ExportMigQuit_Click hides the form and
' LBL_Previous_Click hides it and shows F_Advanced.
'
' The success boxes show the saved file paths under MSG_FileSaved: the path the
' export answered is what the user needs, and MSG_FileSaved is in the linelist
' message table in five languages.

Option Explicit

' The linelist chosen for the other-linelist export. The labels of the form
' show these two, and this pair is what the export reads. The module state
' dies on any unhandled error while the form outlives it; the export then
' finds an empty path and asks for the file again.
Private otherLinelistPath As String
Private otherLinelistPassword As String


'@section The export click
'===============================================================================

' @description Run the exports the form's checkboxes ask for.
' The other-linelist box excludes the five current-linelist boxes: it exports
' everything of another linelist, so combining the two makes no export to name,
' and the user is asked to fix the boxes. The current-linelist walk exports the
' data, the geobase and the historic geobase into one chosen folder; the
' other-linelist walk confirms, then exports everything of that file, geobase
' included, into one chosen folder.
' @param frm Object. The F_ExportMig form.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
Public Sub HandleExportClick(ByVal frm As Object, _
                             ByVal sourceWkb As Workbook, _
                             ByVal trads As TranslationObject)

    Dim wantData As Boolean
    Dim wantGeo As Boolean
    Dim wantHistoric As Boolean
    Dim wantOther As Boolean
    Dim includeShowHide As Boolean
    Dim keepLabels As Boolean

    On Error GoTo ErrHand

    wantData = frm.CHK_ExportMigData.Value
    includeShowHide = frm.CHK_ExportMigShowHide.Value
    keepLabels = frm.CHK_ExportMigEditableLabel.Value
    wantGeo = frm.CHK_ExportMigGeo.Value
    wantHistoric = frm.CHK_ExportMigGeoHistoric.Value
    wantOther = frm.CHK_OtherLinelist.Value

    If wantOther And (wantData Or wantGeo Or wantHistoric Or _
                      includeShowHide Or keepLabels) Then
        MsgBox trads.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               trads.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    If wantOther Then
        OtherLinelistWalk frm, sourceWkb, trads
    ElseIf wantData Or wantGeo Or wantHistoric Then
        CurrentLinelistWalk frm, sourceWkb, trads, wantData, wantGeo, _
                            wantHistoric, includeShowHide, keepLabels
    End If
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
End Sub


' @description Export the current linelist: the migration file, the geobase
' and the historic geobase, whichever of the three the boxes ask for, into
' one folder the user picks once.
Private Sub CurrentLinelistWalk(ByVal frm As Object, _
                                ByVal sourceWkb As Workbook, _
                                ByVal trads As TranslationObject, _
                                ByVal wantData As Boolean, _
                                ByVal wantGeo As Boolean, _
                                ByVal wantHistoric As Boolean, _
                                ByVal includeShowHide As Boolean, _
                                ByVal keepLabels As Boolean)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim folderPath As String
    Dim savedPaths As String

    On Error GoTo ErrHand

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    Set exporter = LLExporter.Create(sourceWkb)

    If wantData Then _
        savedPaths = exporter.ExportMigration(folderPath, includeShowHide, keepLabels)
    If wantGeo Then _
        savedPaths = JoinPath(savedPaths, exporter.ExportGeo(folderPath, onlyHistoric:=False))
    If wantHistoric Then _
        savedPaths = JoinPath(savedPaths, exporter.ExportGeo(folderPath, onlyHistoric:=True))

    appState.Restore

    MsgBox savedPaths, vbOKOnly + vbInformation, _
           trads.TranslatedValue("MSG_FileSaved")

    If MsgBox(trads.TranslatedValue("MSG_FinishedExports"), _
              vbQuestion + vbYesNo, _
              trads.TranslatedValue("MSG_Migration")) = vbYes Then frm.Hide
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Sub


' @description Export everything of another linelist: its whole migration
' file, show/hide state and editable labels included, and its full geobase.
' The user confirms first, then picks one folder for the two files. The other
' linelist is opened read-only under the busy state, so its open events stay
' quiet, and it is closed without saving once the files are written.
Private Sub OtherLinelistWalk(ByVal frm As Object, _
                              ByVal sourceWkb As Workbook, _
                              ByVal trads As TranslationObject)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim folderPath As String
    Dim savedPaths As String

    On Error GoTo ErrHand

    If LenB(Dir(otherLinelistPath)) = 0 Then
        MsgBox trads.TranslatedValue("MSG_ProvideLLPath"), _
               vbExclamation + vbOKOnly, _
               trads.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    If LCase$(otherLinelistPath) = LCase$(sourceWkb.FullName) Then
        MsgBox trads.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               trads.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    If MsgBox(trads.TranslatedValue("MSG_ConfirmExportOther") & _
              vbNewLine & otherLinelistPath, _
              vbQuestion + vbYesNo, _
              trads.TranslatedValue("MSG_Confirm")) = vbNo Then Exit Sub

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    'A path or password the file refuses lands here, and the message names
    'that failure rather than the export one.
    On Error GoTo ErrOpen
    Set exporter = LLExporter.CreateFromFile(otherLinelistPath, otherLinelistPassword)
    On Error GoTo ErrHand

    savedPaths = exporter.ExportMigration(folderPath, includeShowHide:=True, _
                                          keepLabels:=True)
    savedPaths = JoinPath(savedPaths, exporter.ExportGeo(folderPath, onlyHistoric:=False))

    'Closes the other linelist, which this exporter opened
    exporter.CloseAll

    appState.Restore

    MsgBox savedPaths, vbOKOnly + vbInformation, _
           trads.TranslatedValue("MSG_FileSaved")

    If MsgBox(trads.TranslatedValue("MSG_FinishedExports"), _
              vbQuestion + vbYesNo, _
              trads.TranslatedValue("MSG_Migration")) = vbYes Then frm.Hide
    Exit Sub

ErrOpen:
    On Error Resume Next
    If Not appState Is Nothing Then appState.Restore
    MsgBox trads.TranslatedValue("MSG_ErrOpenOther"), _
           vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    Exit Sub

ErrHand:
    On Error Resume Next
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
    MsgBox trads.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
End Sub


'@section The other-linelist choice
'===============================================================================

' @description React to the other-linelist box going on. The five
' current-linelist boxes conflict with it: when one of them is on, a message
' asks the user to fix the boxes and the box goes back off. Otherwise the user
' picks the linelist file at once, then gives its password; cancelling the
' file picker puts the box back off.
' @param frm Object. The F_ExportMig form.
' @param trads TranslationObject. Translations for messages.
Public Sub HandleOtherLinelistChecked(ByVal frm As Object, _
                                      ByVal trads As TranslationObject)

    If frm.CHK_ExportMigData.Value Or frm.CHK_ExportMigShowHide.Value Or _
       frm.CHK_ExportMigEditableLabel.Value Or frm.CHK_ExportMigGeo.Value Or _
       frm.CHK_ExportMigGeoHistoric.Value Then
        MsgBox trads.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               trads.TranslatedValue("MSG_Migration")
        frm.CHK_OtherLinelist.Value = False
        Exit Sub
    End If

    PromptOtherLinelistPath frm, trads
    If LenB(otherLinelistPath) = 0 Then
        frm.CHK_OtherLinelist.Value = False
        Exit Sub
    End If

    PromptOtherLinelistPassword frm, trads
End Sub

' @description Forget the chosen linelist when the box goes off. The form stub
' translates the form right after, which puts the hint captions back on the
' two labels.
Public Sub HandleOtherLinelistUnchecked()
    otherLinelistPath = vbNullString
    otherLinelistPassword = vbNullString
End Sub

' @description Ask for the linelist file and write it on the path label.
' Quiet while the other-linelist box is off. Choosing the current linelist is
' refused with the conflict message, since the current-linelist boxes are the
' way to export it.
' @param frm Object. The F_ExportMig form.
' @param trads TranslationObject. Translations for messages.
Public Sub PromptOtherLinelistPath(ByVal frm As Object, _
                                   ByVal trads As TranslationObject)

    Dim io As OSFiles

    If Not frm.CHK_OtherLinelist.Value Then Exit Sub

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb"
    If Not io.HasValidFile Then Exit Sub

    If LCase$(io.File()) = LCase$(ThisWorkbook.FullName) Then
        MsgBox trads.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               trads.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    otherLinelistPath = io.File()
    frm.LBL_OtherPath.Caption = otherLinelistPath
End Sub

' @description Ask for the password of the chosen linelist and write it on the
' password label. Quiet while the other-linelist box is off. An empty answer
' stands for a linelist with no password.
' @param frm Object. The F_ExportMig form.
' @param trads TranslationObject. Translations for messages.
Public Sub PromptOtherLinelistPassword(ByVal frm As Object, _
                                       ByVal trads As TranslationObject)

    If Not frm.CHK_OtherLinelist.Value Then Exit Sub

    otherLinelistPassword = InputBox( _
        trads.TranslatedValue("MSG_ProvideLLPassword"), _
        trads.TranslatedValue("MSG_Migration"))
    frm.LBL_OtherPass.Caption = otherLinelistPassword
End Sub


'@section Analysis export
'===============================================================================

' @description Export analysis worksheets only.
' Shows a folder picker, creates the analysis export, and handles errors.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
Public Sub HandleAnalysisExport(ByVal sourceWkb As Workbook, _
                                ByVal trads As TranslationObject)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim folderPath As String
    Dim filePath As String

    On Error GoTo ErrHand

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    Set exporter = LLExporter.Create(sourceWkb)
    filePath = exporter.ExportAnalysis(folderPath)

    appState.Restore

    MsgBox filePath, vbOKOnly + vbInformation, _
           trads.TranslatedValue("MSG_FileSaved")
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Sub


'@section Helpers
'===============================================================================

' @description Ask the user for the folder the export files land in.
' @return String. The chosen folder, or an empty string on cancel.
Private Function PickExportFolder() As String
    Dim io As OSFiles

    Set io = OSFiles.Create()
    io.LoadFolder
    If io.HasValidFolder Then PickExportFolder = io.Folder()
End Function

' @description Stack a new file path under the ones already collected.
' @param collected String. The paths so far, possibly empty.
' @param newPath String. The path to add.
' @return String. The paths on one line each.
Private Function JoinPath(ByVal collected As String, _
                          ByVal newPath As String) As String
    If LenB(collected) = 0 Then
        JoinPath = newPath
    Else
        JoinPath = collected & vbNewLine & newPath
    End If
End Function
