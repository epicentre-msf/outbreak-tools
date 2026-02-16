Attribute VB_Name = "FormLogicAdvanced"

'@Folder("DataIO")
'@ModuleDescription("Import data, import geobase, and clear data workflows")
'@depends ILLImporter, LLImporter, IApplicationState, ApplicationState, IOSFiles, OSFiles

Option Explicit


' @description Import data from a migration workbook.
' Shows file picker, checks language, imports data and metadata, shows report.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads ITranslationObject. Translations for messages.
' @param pasteAtBottom Boolean. When True, appends data below existing rows.
Public Sub HandleImportData(ByVal sourceWkb As Workbook, _
                            ByVal trads As ITranslationObject, _
                            ByVal pasteAtBottom As Boolean)

    Dim imp As ILLImporter
    Dim appState As IApplicationState
    Dim io As IOSFiles
    Dim filePath As String
    Dim impwb As Workbook
    Dim actsh As Worksheet

    On Error GoTo ErrHand

    ' Select import file
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile Then Exit Sub
    filePath = io.File()

    ' Confirm import
    If MsgBox(trads.TranslatedValue("MSG_ImportConfirm"), _
              vbOKCancel, trads.TranslatedValue("MSG_Confirm")) = vbCancel Then
        GoTo EndImport
    End If

    ' Busy state
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState True, True, xlWait, False
    Set actsh = ActiveSheet

    ' Open import workbook
    Set impwb = Workbooks.Open(filePath)
    ActiveWindow.WindowState = xlMinimized

    ' Create importer and check language
    Set imp = LLImporter.Create(sourceWkb)

    If Not imp.HasSameLanguage(impwb) Then
        If MsgBox(trads.TranslatedValue("MSG_NoLanguage"), _
                  vbExclamation + vbYesNo, _
                  trads.TranslatedValue("MSG_Imports")) = vbNo Then
            GoTo EndImport
        End If
    End If

    ' Import all data
    imp.ImportData impwb, pasteAtBottom
    imp.ImportCustomDropdown impwb, pasteAtBottom
    imp.FinalizeReport

    ' Import migration metadata
    imp.ImportShowHide impwb
    imp.ImportEditableLabels impwb
    imp.ImportSingleValues impwb

    ' Close import workbook
    impwb.Close savechanges:=False
    Set impwb = Nothing

    actsh.Activate
    appState.Restore

    ' Show result
    If imp.NeedReport Then
        MsgBox trads.TranslatedValue("MSG_FinishImportRep"), _
               vbQuestion + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    Else
        MsgBox trads.TranslatedValue("MSG_FinishImport"), _
               vbOKOnly, trads.TranslatedValue("MSG_Imports")
    End If
    Exit Sub

EndImport:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_AbortImport"), _
           vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not actsh Is Nothing Then actsh.Activate
    If Not appState Is Nothing Then appState.Restore
    On Error GoTo 0
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrorImport"), _
           vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not actsh Is Nothing Then actsh.Activate
    If Not appState Is Nothing Then appState.Restore
End Sub


' @description Import a geobase from an external workbook.
' Shows file picker, imports geobase data, optionally updates headers and dictionary.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads ITranslationObject. Translations for messages.
' @param histoOnly Boolean. When True, imports only historic geobase data.
Public Sub HandleImportGeobase(ByVal sourceWkb As Workbook, _
                               ByVal trads As ITranslationObject, _
                               Optional ByVal histoOnly As Boolean = False)

    Dim imp As ILLImporter
    Dim appState As IApplicationState
    Dim io As IOSFiles
    Dim filePath As String
    Dim impwb As Workbook

    On Error GoTo ErrHand

    ' Select geobase file
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile Then Exit Sub
    filePath = io.File()

    ' Busy state
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState True, True, xlWait, False

    ' Open geobase workbook
    Set impwb = Workbooks.Open(filePath)
    ActiveWindow.WindowState = xlMinimized

    ' Import geobase
    Set imp = LLImporter.Create(sourceWkb)
    imp.ImportGeobase impwb, histoOnly

    impwb.Close savechanges:=False
    Set impwb = Nothing

    appState.Restore

    MsgBox trads.TranslatedValue("MSG_FinishImportGeo"), _
           vbOKOnly, trads.TranslatedValue("MSG_Imports")
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrImportGeo"), _
           vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not appState Is Nothing Then appState.Restore
End Sub


' @description Clear all entered data from the linelist.
' Prompts the user for workbook name confirmation before deleting.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads ITranslationObject. Translations for messages.
Public Sub HandleClearData(ByVal sourceWkb As Workbook, _
                           ByVal trads As ITranslationObject)

    Dim imp As ILLImporter
    Dim appState As IApplicationState
    Dim proceed As Long
    Dim inputName As String
    Dim goodName As Boolean

    On Error GoTo ErrHand

    ' Confirm deletion
    proceed = MsgBox(trads.TranslatedValue("MSG_DeleteAllData"), _
                     vbExclamation + vbYesNo, _
                     trads.TranslatedValue("MSG_Delete"))
    If proceed <> vbYes Then
        MsgBox trads.TranslatedValue("MSG_DelCancel"), _
               vbOKOnly, trads.TranslatedValue("MSG_Delete")
        Exit Sub
    End If

    ' Require workbook name confirmation
    goodName = False
    Do While Not goodName
        inputName = InputBox(trads.TranslatedValue("MSG_LLName"), _
                             trads.TranslatedValue("MSG_Delete"), _
                             trads.TranslatedValue("MSG_EnterWkbName"))

        If StrPtr(inputName) = 0 Then
            ' User cancelled
            MsgBox trads.TranslatedValue("MSG_DelCancel"), _
                   vbOKOnly, trads.TranslatedValue("MSG_Delete")
            Exit Sub

        ElseIf inputName = Replace(sourceWkb.Name, ".xlsb", vbNullString) Then
            goodName = True

        Else
            If MsgBox(trads.TranslatedValue("MSG_BadLLNameQ"), _
                      vbExclamation + vbYesNo, _
                      trads.TranslatedValue("MSG_Delete")) = vbNo Then
                Exit Sub
            End If
        End If
    Loop

    ' Proceed with deletion
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState True, False, xlWait, False

    Set imp = LLImporter.Create(sourceWkb)
    imp.ClearData

    appState.Restore
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrClearData"), _
           vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Error")
    If Not appState Is Nothing Then appState.Restore
End Sub
