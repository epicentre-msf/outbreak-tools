Attribute VB_Name = "FormLogicExportMig"

'@Folder("Linelist Forms")
'@ModuleDescription("Complete code-behind of F_ExportMig -- migration, geo and other-linelist exports")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@depends LinelistRun, OSFiles, TranslationObject, LLTranslation, LLLog, LinelistEventsManager, EventLinelist, Messenger

' This module is the complete code-behind of the F_ExportMig form and is
' copied into the form at deployment, so every control callback and the
' translation build live here. EventsLinelistButtons reaches
' HandleAnalysisExport through F_ExportMig.HandleAnalysisExport, the qualified
' route, so the standard-module copy of this code is never compiled and its
' form references cost nothing there.
'
' THE FOUR EXPORT WALKS LEFT THIS MODULE. They live in LinelistRun, a standard
' module that is compiled everywhere and can be tested, and that RunExport
' drives with no form in the way. What stayed here is everything that reads a
' control or moves the form: the five checkboxes, the folder picker, the two
' other-linelist labels, the password InputBox and the question that puts the
' form away.
'
' The success boxes show the saved file paths under MSG_FileSaved: the path
' the export answered is what the user needs, and MSG_FileSaved is in the
' linelist message table in five languages. They are shown by the walks in
' LinelistRun, through the messenger, so a scripted run swallows them.

Option Explicit


Private tradform As TranslationObject
Private tradmess As TranslationObject

' The linelist chosen for the other-linelist export. The two labels of the
' form show this pair, and this pair is what the export reads. The pair
' outlives an untick, so a user who unticks the box and ticks it back finds
' the file already chosen; the module state dies on any unhandled error while
' the form outlives it, and the export then asks for the file to be chosen
' again.
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

' @description Translate the form. The two other-linelist labels start hidden,
' because the box that owns them starts off; ticking the box is what brings
' them up.
Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
    ShowOtherLinelistChoice Me.CHK_OtherLinelist.Value
    Me.Width = 250
    Me.Height = 480
End Sub


' @description The user log the event service holds. A workbook whose log
' cannot be built answers Nothing and every log line below stays quiet.
Private Function UserLogOf() As LLLog
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    Set UserLogOf = linelistEvents.UserLog()
End Function


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


'@section The other-linelist choice
'===============================================================================

' @description React to the other-linelist box. The box says where the export
' lands, not what it carries: the five export boxes are read the same way for
' either target, so no box conflicts with this one. Ticking brings up the path
' and password labels, unticking puts them away, and nothing else happens
' here -- the file is chosen by double-clicking the labels, and the export
' runs on the export button.
Private Sub CHK_OtherLinelist_Click()
    ShowOtherLinelistChoice Me.CHK_OtherLinelist.Value
End Sub

' @description Put the path and password labels up or down.
' @param shown Boolean. True brings the two labels up.
Private Sub ShowOtherLinelistChoice(ByVal shown As Boolean)
    Me.LBL_OtherPath.Visible = shown
    Me.LBL_OtherPass.Visible = shown
End Sub

' @description Choose the linelist file. A single click is left alone, so the
' label answers the second click only: a Click handler would open the dialog
' before the second click could ever arrive.
Private Sub LBL_OtherPath_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    InitializeTrads
    PromptOtherLinelistPath
End Sub

' @description Give the password of the chosen linelist, on a double click.
Private Sub LBL_OtherPass_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    InitializeTrads
    PromptOtherLinelistPassword
End Sub

' @description Ask for the linelist file and write it on the path label.
' Quiet while the other-linelist box is off. Choosing the current linelist is
' refused with the conflict message, since the box is off is how the current
' linelist is exported.
Private Sub PromptOtherLinelistPath()
    Dim io As OSFiles

    If Not Me.CHK_OtherLinelist.Value Then Exit Sub

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb"
    If Not io.HasValidFile Then Exit Sub

    If LCase$(io.File()) = LCase$(ThisWorkbook.FullName) Then
        Messenger.Show tradmess.TranslatedValue("MSG_ExportMigConflict"), vbOK, _
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

' @description Run the exports the checkboxes ask for. The five export boxes
' say what the export carries -- the migration file, the geobase, the historic
' geobase, and for the migration file the show/hide state and the editable
' labels. The other-linelist box says which linelist those five are read from:
' off, the current one; on, the file chosen on the labels. Either walk exports
' into one folder the user picks once.
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

    If Not (wantData Or wantGeo Or wantHistoric) Then Exit Sub

    If wantOther Then
        OtherLinelistWalk wantData, wantGeo, wantHistoric, _
                          includeShowHide, keepLabels
    Else
        CurrentLinelistWalk wantData, wantGeo, wantHistoric, _
                            includeShowHide, keepLabels
    End If
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "export-migration", failDetail
    Messenger.Show tradmess.TranslatedValue("MSG_ErrHandExport"), vbOK, _
                   vbOKOnly + vbCritical, tradmess.TranslatedValue("MSG_Error")
End Sub

' @description Export the current linelist: the migration file, the geobase
' and the historic geobase, whichever of the three the boxes ask for, into
' one folder the user picks once. The walk itself is
' LinelistRun.HandleExportMigration, which shows the saved paths; this asks for
' the folder first and puts the form away after.
' @param wantData Boolean. True exports the migration file.
' @param wantGeo Boolean. True exports the geobase.
' @param wantHistoric Boolean. True exports the historic geobase.
' @param includeShowHide Boolean. True carries the show/hide state into the migration file.
' @param keepLabels Boolean. True carries the editable labels into the migration file.
Private Sub CurrentLinelistWalk(ByVal wantData As Boolean, _
                                ByVal wantGeo As Boolean, _
                                ByVal wantHistoric As Boolean, _
                                ByVal includeShowHide As Boolean, _
                                ByVal keepLabels As Boolean)

    Dim folderPath As String
    Dim savedPaths As String

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    savedPaths = LinelistRun.HandleExportMigration(ThisWorkbook, tradmess, folderPath, _
                                                   wantData, wantGeo, wantHistoric, _
                                                   includeShowHide, keepLabels)

    'A walk that wrote nothing was refused or failed, and it has already said
    'so. The question below is about the form, so it is asked only when there
    'was an export to finish.
    If LenB(savedPaths) = 0 Then Exit Sub

    If MsgBox(tradmess.TranslatedValue("MSG_FinishedExports"), _
              vbQuestion + vbYesNo, _
              tradmess.TranslatedValue("MSG_Migration")) = vbYes Then Me.Hide
End Sub

' @description Export another linelist, the one chosen on the two labels,
' with the same three files and the same two migration switches the boxes ask
' of the current linelist. The walk is LinelistRun.HandleExportOther, which
' confirms the file, opens it read-only, writes the files and closes it again
' without saving. This asks for the folder first and puts the form away after.
' @param wantData Boolean. True exports the migration file.
' @param wantGeo Boolean. True exports the geobase.
' @param wantHistoric Boolean. True exports the historic geobase.
' @param includeShowHide Boolean. True carries the show/hide state into the migration file.
' @param keepLabels Boolean. True carries the editable labels into the migration file.
Private Sub OtherLinelistWalk(ByVal wantData As Boolean, _
                              ByVal wantGeo As Boolean, _
                              ByVal wantHistoric As Boolean, _
                              ByVal includeShowHide As Boolean, _
                              ByVal keepLabels As Boolean)

    Dim folderPath As String
    Dim savedPaths As String

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    savedPaths = LinelistRun.HandleExportOther(tradmess, folderPath, _
                                               otherLinelistPath, otherLinelistPassword, _
                                               wantData, wantGeo, wantHistoric, _
                                               includeShowHide, keepLabels)

    If LenB(savedPaths) = 0 Then Exit Sub

    If MsgBox(tradmess.TranslatedValue("MSG_FinishedExports"), _
              vbQuestion + vbYesNo, _
              tradmess.TranslatedValue("MSG_Migration")) = vbYes Then Me.Hide
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

Private Sub CHK_ExportMigData_Click()
    CHK_ExportMigEditableLabel.Value = CHK_ExportMigData.Value
    CHK_ExportMigShowHide.Value = CHK_ExportMigData.Value
End Sub


'@section Analysis export
'===============================================================================

' @description Export analysis worksheets only. The picker is here and the
' walk is LinelistRun.HandleExportAnalysis, which writes the file and shows the
' path. No control of the form takes part: EventsLinelistButtons
' .ClickExportAnalysis calls this through F_ExportMig.HandleAnalysisExport with
' its own workbook and translations.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
Public Sub HandleAnalysisExport(ByVal sourceWkb As Workbook, _
                                ByVal trads As TranslationObject)

    Dim folderPath As String

    folderPath = PickExportFolder()
    If LenB(folderPath) = 0 Then Exit Sub

    LinelistRun.HandleExportAnalysis sourceWkb, trads, folderPath
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
