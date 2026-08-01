Attribute VB_Name = "FormLogicExportMig"

'@Folder("Linelist Forms")
'@ModuleDescription("Migration, analysis, and geo export workflows")
'@depends LLExporter, ApplicationState, OSFiles

Option Explicit

' The three boxes below used to be built from MSG_Export, MSG_ExportSuccess and
' MSG_ExportGeoSuccess. None of the three is in either translation workbook and
' TranslateTag answers the tag itself on a miss, so the user read the raw key.
' MSG_FileSaved is in the linelist message table in five languages, and the path
' the export answered is what the user actually needs.


' @description Export all data for migration to another linelist.
' Shows a folder picker, creates the migration export, and handles errors.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
' @param includeShowHide Boolean. Include show/hide state with column widths.
' @param keepLabels Boolean. Mark editable labels for update on import.
Public Sub HandleMigrationExport(ByVal sourceWkb As Workbook, _
                                 ByVal trads As TranslationObject, _
                                 ByVal includeShowHide As Boolean, _
                                 ByVal keepLabels As Boolean)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim io As OSFiles
    Dim folderPath As String
    Dim filePath As String

    On Error GoTo ErrHand

    ' Select export folder
    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

    ' Busy state
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    ' Export
    Set exporter = LLExporter.Create(sourceWkb)
    filePath = exporter.ExportMigration(folderPath, includeShowHide, keepLabels)

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


' @description Export analysis worksheets only.
' Shows a folder picker, creates the analysis export, and handles errors.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
Public Sub HandleAnalysisExport(ByVal sourceWkb As Workbook, _
                                ByVal trads As TranslationObject)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim io As OSFiles
    Dim folderPath As String
    Dim filePath As String

    On Error GoTo ErrHand

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

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


' @description Export geobase data to a separate workbook.
' Shows a folder picker, creates the geo export, and handles errors.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
' @param onlyHistoric Boolean. When True, exports only historic geobase data.
Public Sub HandleGeoExport(ByVal sourceWkb As Workbook, _
                           ByVal trads As TranslationObject, _
                           Optional ByVal onlyHistoric As Boolean = False)

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim io As OSFiles
    Dim folderPath As String
    Dim filePath As String

    On Error GoTo ErrHand

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    Set exporter = LLExporter.Create(sourceWkb)
    filePath = exporter.ExportGeo(folderPath, onlyHistoric)

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
