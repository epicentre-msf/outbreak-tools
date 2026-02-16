Attribute VB_Name = "FormLogicExportMig"

'@Folder("DataIO")
'@ModuleDescription("Migration, analysis, and geo export workflows")
'@depends ILLExporter, LLExporter, IApplicationState, ApplicationState, IOSFiles, OSFiles

Option Explicit


' @description Export all data for migration to another linelist.
' Shows a folder picker, creates the migration export, and handles errors.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads ITranslationObject. Translations for messages.
' @param includeShowHide Boolean. Include show/hide state with column widths.
' @param keepLabels Boolean. Mark editable labels for update on import.
Public Sub HandleMigrationExport(ByVal sourceWkb As Workbook, _
                                 ByVal trads As ITranslationObject, _
                                 ByVal includeShowHide As Boolean, _
                                 ByVal keepLabels As Boolean)

    Dim exporter As ILLExporter
    Dim appState As IApplicationState
    Dim io As IOSFiles
    Dim folderPath As String

    On Error GoTo ErrHand

    ' Select export folder
    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

    ' Busy state
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState True, False, xlWait, False

    ' Export
    Set exporter = LLExporter.Create(sourceWkb)
    exporter.ExportMigration folderPath, includeShowHide, keepLabels

    appState.Restore

    MsgBox trads.TranslatedValue("MSG_ExportSuccess"), _
           vbOKOnly + vbInformation, trads.TranslatedValue("MSG_Export")
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
' @param trads ITranslationObject. Translations for messages.
Public Sub HandleAnalysisExport(ByVal sourceWkb As Workbook, _
                                ByVal trads As ITranslationObject)

    Dim exporter As ILLExporter
    Dim appState As IApplicationState
    Dim io As IOSFiles
    Dim folderPath As String

    On Error GoTo ErrHand

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState True, False, xlWait, False

    Set exporter = LLExporter.Create(sourceWkb)
    exporter.ExportAnalysis folderPath

    appState.Restore

    MsgBox trads.TranslatedValue("MSG_ExportSuccess"), _
           vbOKOnly + vbInformation, trads.TranslatedValue("MSG_Export")
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
' @param trads ITranslationObject. Translations for messages.
' @param onlyHistoric Boolean. When True, exports only historic geobase data.
Public Sub HandleGeoExport(ByVal sourceWkb As Workbook, _
                           ByVal trads As ITranslationObject, _
                           Optional ByVal onlyHistoric As Boolean = False)

    Dim exporter As ILLExporter
    Dim appState As IApplicationState
    Dim io As IOSFiles
    Dim folderPath As String

    On Error GoTo ErrHand

    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState True, False, xlWait, False

    Set exporter = LLExporter.Create(sourceWkb)
    exporter.ExportGeo folderPath, onlyHistoric

    appState.Restore

    MsgBox trads.TranslatedValue("MSG_ExportGeoSuccess"), _
           vbOKOnly + vbInformation, trads.TranslatedValue("MSG_Export")
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Sub
