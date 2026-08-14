Attribute VB_Name = "FormLogicExportMig"

'@Folder("Linelist Forms")
'@ModuleDescription("Complete code-behind of F_ExportMig -- migration, geo and other-linelist exports")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@depends LLExporter, ApplicationState, OSFiles, TranslationObject, LLTranslation, LLLog

' This module is the complete code-behind of the F_ExportMig form and is
' copied into the form at deployment, so every control callback, the
' translation build and the export walks live here. EventsLinelistButtons
' reaches HandleAnalysisExport through F_ExportMig.HandleAnalysisExport, the
' qualified route, so the standard-module copy of this code is never compiled
' and its form references cost nothing there.
'
' The success boxes show the saved file paths under MSG_FileSaved: the path
' the export answered is what the user needs, and MSG_FileSaved is in the
' linelist message table in five languages.

Option Explicit


Private tradform As TranslationObject
Private tradmess As TranslationObject

' The linelist chosen for the other-linelist export. The two labels of the
' form show this pair, and this pair is what the export reads. The module
' state dies on any unhandled error while the form outlives it; the export
' then finds an empty path and asks for the file again.
Private otherLinelistPath As String
Private otherLinelistPassword As String


'@section Initialization
'===============================================================================

' @description Build the two translation objects when they are missing. Every
' entry point rebuilds what it reads when the state is gone.
Private Sub InitializeTrads()
    Dim linelistEvents As EventLinelist
    Dim lltrads As LLTranslation

    If Not (tradform Is Nothing Or tradmess Is Nothing) Then Exit Sub

    ' The helper is the one EventLinelist holds. This module used to build its
    ' own, and LLTranslation.Create validates all five translation tables per
    ' build.
    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If Not linelistEvents Is Nothing Then Set lltrads = linelistEvents.Translation()

    If lltrads Is Nothing Then _
        Err.Raise ProjectError.ObjectNotInitialized, "FormLogicExportMig", _
                  "This linelist carries no usable translation sheet"

    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
End Sub

' @description Translate the form.
Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
    Me.Width = 250
    Me.Height = 450
End Sub


' @description The user log the event service holds. A workbook whose log
' cannot be built answers Nothing and every log line below stays quiet.
Private Function UserLogOf() As LLLog
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    Set UserLogOf = linelistEvents.UserLog()
End Function


' @description Write the success line of a finished export walk. The write
' is guarded so a log fault never takes down the walk it records.
Private Sub LogSuccessLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSuccess action, detail
    On Error GoTo 0
End Sub


' @description Write the failure line of a walk that ended at its error label.
Private Sub LogFailureLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogFailure action, detail
    On Error GoTo 0
End Sub


' @description The saved paths on one line, for the log detail.
Private Function PathsOnOneLine(ByVal savedPaths As String) As String
    PathsOnOneLine = Replace(savedPaths, vbNewLine, ", ")
End Function


'@section The other-linelist choice
'===============================================================================

' @description React to the other-linelist box. The five current-linelist
' boxes conflict with it: when one of them is on, a message asks the user to
' fix the boxes and the box goes back off. Otherwise the user picks the
' linelist file at once, then gives its password; cancelling the file picker
' puts the box back off. The box going off forgets the chosen pair and the
' translation pass puts the hint captions back on the two labels.
Private Sub CHK_OtherLinelist_Click()
    InitializeTrads

    If Not Me.CHK_OtherLinelist.Value Then
        otherLinelistPath = vbNullString
        otherLinelistPassword = vbNullString
        tradform.TranslateForm Me
        Exit Sub
    End If

    If Me.CHK_ExportMigData.Value Or Me.CHK_ExportMigShowHide.Value Or _
       Me.CHK_ExportMigEditableLabel.Value Or Me.CHK_ExportMigGeo.Value Or _
       Me.CHK_ExportMigGeoHistoric.Value Then
        MsgBox tradmess.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               tradmess.TranslatedValue("MSG_Migration")
        Me.CHK_OtherLinelist.Value = False
        Exit Sub
    End If

    PromptOtherLinelistPath
    If LenB(otherLinelistPath) = 0 Then
        Me.CHK_OtherLinelist.Value = False
        Exit Sub
    End If

    PromptOtherLinelistPassword
End Sub

Private Sub LBL_OtherPath_Click()
    InitializeTrads
    PromptOtherLinelistPath
End Sub

Private Sub LBL_OtherPass_Click()
    InitializeTrads
    PromptOtherLinelistPassword
End Sub

' @description Ask for the linelist file and write it on the path label.
' Quiet while the other-linelist box is off. Choosing the current linelist is
' refused with the conflict message, since the current-linelist boxes are the
' way to export it.
Private Sub PromptOtherLinelistPath()
    Dim io As OSFiles

    If Not Me.CHK_OtherLinelist.Value Then Exit Sub

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb"
    If Not io.HasValidFile Then Exit Sub

    If LCase$(io.File()) = LCase$(ThisWorkbook.FullName) Then
        MsgBox tradmess.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               tradmess.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    otherLinelistPath = io.File()
    Me.LBL_OtherPath.Caption = otherLinelistPath
End Sub

' @description Ask for the password of the chosen linelist and write it on
' the password label. Quiet while the other-linelist box is off. An empty
' answer stands for a linelist with no password.
Private Sub PromptOtherLinelistPassword()
    If Not Me.CHK_OtherLinelist.Value Then Exit Sub

    otherLinelistPassword = InputBox( _
        tradmess.TranslatedValue("MSG_ProvideLLPassword"), _
        tradmess.TranslatedValue("MSG_Migration"))
    Me.LBL_OtherPass.Caption = otherLinelistPassword
End Sub


'@section The export click
'===============================================================================

' @description Run the exports the checkboxes ask for. The other-linelist box
' excludes the five current-linelist boxes: it exports everything of another
' linelist, so combining the two makes no export to name, and the user is
' asked to fix the boxes. The current-linelist walk exports the data, the
' geobase and the historic geobase into one chosen folder; the other-linelist
' walk confirms, then exports everything of that file, geobase included, into
' one chosen folder.
Private Sub CMD_ExportMig_Click()
    Dim wantData As Boolean
    Dim wantGeo As Boolean
    Dim wantHistoric As Boolean
    Dim wantOther As Boolean
    Dim includeShowHide As Boolean
    Dim keepLabels As Boolean
    Dim failDetail As String

    On Error GoTo ErrHand

    InitializeTrads

    wantData = Me.CHK_ExportMigData.Value
    includeShowHide = Me.CHK_ExportMigShowHide.Value
    keepLabels = Me.CHK_ExportMigEditableLabel.Value
    wantGeo = Me.CHK_ExportMigGeo.Value
    wantHistoric = Me.CHK_ExportMigGeoHistoric.Value
    wantOther = Me.CHK_OtherLinelist.Value

    If wantOther And (wantData Or wantGeo Or wantHistoric Or _
                      includeShowHide Or keepLabels) Then
        MsgBox tradmess.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               tradmess.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    If wantOther Then
        OtherLinelistWalk
    ElseIf wantData Or wantGeo Or wantHistoric Then
        CurrentLinelistWalk wantData, wantGeo, wantHistoric, _
                            includeShowHide, keepLabels
    End If
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "export-migration", failDetail
    MsgBox tradmess.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, tradmess.TranslatedValue("MSG_Error")
End Sub

' @description Export the current linelist: the migration file, the geobase
' and the historic geobase, whichever of the three the boxes ask for, into
' one folder the user picks once.
Private Sub CurrentLinelistWalk(ByVal wantData As Boolean, _
                                ByVal wantGeo As Boolean, _
                                ByVal wantHistoric As Boolean, _
                                ByVal includeShowHide As Boolean, _
                                ByVal keepLabels As Boolean)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim folderPath As String
    Dim savedPaths As String
    Dim failDetail As String

    On Error GoTo ErrHand

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    Set exporter = LLExporter.Create(ThisWorkbook)

    If wantData Then _
        savedPaths = exporter.ExportMigration(folderPath, includeShowHide, keepLabels)
    If wantGeo Then _
        savedPaths = JoinPath(savedPaths, exporter.ExportGeo(folderPath, onlyHistoric:=False))
    If wantHistoric Then _
        savedPaths = JoinPath(savedPaths, exporter.ExportGeo(folderPath, onlyHistoric:=True))

    appState.Restore
    LogSuccessLine "export-migration", PathsOnOneLine(savedPaths)

    MsgBox savedPaths, vbOKOnly + vbInformation, _
           tradmess.TranslatedValue("MSG_FileSaved")

    If MsgBox(tradmess.TranslatedValue("MSG_FinishedExports"), _
              vbQuestion + vbYesNo, _
              tradmess.TranslatedValue("MSG_Migration")) = vbYes Then Me.Hide
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "export-migration", failDetail
    MsgBox tradmess.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, tradmess.TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Sub

' @description Export everything of another linelist: its whole migration
' file, show/hide state and editable labels included, and its full geobase.
' The user confirms first, then picks one folder for the two files. The other
' linelist is opened read-only under the busy state, so its open events stay
' quiet, and it is closed without saving once the files are written.
Private Sub OtherLinelistWalk()
    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim folderPath As String
    Dim savedPaths As String
    Dim failDetail As String

    On Error GoTo ErrHand

    If LenB(Dir(otherLinelistPath)) = 0 Then
        MsgBox tradmess.TranslatedValue("MSG_ProvideLLPath"), _
               vbExclamation + vbOKOnly, _
               tradmess.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    If LCase$(otherLinelistPath) = LCase$(ThisWorkbook.FullName) Then
        MsgBox tradmess.TranslatedValue("MSG_ExportMigConflict"), _
               vbExclamation + vbOKOnly, _
               tradmess.TranslatedValue("MSG_Migration")
        Exit Sub
    End If

    If MsgBox(tradmess.TranslatedValue("MSG_ConfirmExportOther") & _
              vbNewLine & otherLinelistPath, _
              vbQuestion + vbYesNo, _
              tradmess.TranslatedValue("MSG_Confirm")) = vbNo Then Exit Sub

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
    LogSuccessLine "export-other", PathsOnOneLine(savedPaths)

    MsgBox savedPaths, vbOKOnly + vbInformation, _
           tradmess.TranslatedValue("MSG_FileSaved")

    If MsgBox(tradmess.TranslatedValue("MSG_FinishedExports"), _
              vbQuestion + vbYesNo, _
              tradmess.TranslatedValue("MSG_Migration")) = vbYes Then Me.Hide
    Exit Sub

ErrOpen:
    On Error Resume Next
    LogFailureLine "export-other", "open refused: " & otherLinelistPath
    If Not appState Is Nothing Then appState.Restore
    MsgBox tradmess.TranslatedValue("MSG_ErrOpenOther"), _
           vbOKOnly + vbCritical, tradmess.TranslatedValue("MSG_Error")
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "export-other", failDetail
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
    MsgBox tradmess.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, tradmess.TranslatedValue("MSG_Error")
End Sub


'@section Navigation
'===============================================================================

Private Sub CMD_ExportMigQuit_Click()
    Me.Hide
End Sub

Private Sub LBL_Previous_Click()
    Me.Hide
    F_Advanced.Show
End Sub


'@section Analysis export
'===============================================================================

' @description Export analysis worksheets only. Shows a folder picker,
' creates the analysis export, and handles errors. No control of the form
' takes part: EventsLinelistButtons.ClickExportAnalysis calls this through
' F_ExportMig.HandleAnalysisExport with its own workbook and translations.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
Public Sub HandleAnalysisExport(ByVal sourceWkb As Workbook, _
                                ByVal trads As TranslationObject)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim folderPath As String
    Dim filePath As String
    Dim failDetail As String

    On Error GoTo ErrHand

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    Set exporter = LLExporter.Create(sourceWkb)
    filePath = exporter.ExportAnalysis(folderPath)

    appState.Restore
    LogSuccessLine "export-analysis", PathsOnOneLine(filePath)

    MsgBox filePath, vbOKOnly + vbInformation, _
           trads.TranslatedValue("MSG_FileSaved")
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "export-analysis", failDetail
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
