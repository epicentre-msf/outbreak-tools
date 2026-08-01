Attribute VB_Name = "FormLogicAdvanced"

'@Folder("Linelist Forms")
'@ModuleDescription("Import data, import geobase, and clear data workflows")
'@depends LLImporter, ImportMetadata, ApplicationState, OSFiles

Option Explicit

' What an import does with the rows the linelist already holds
Private Const PASTING_RULE_EMPTY As Byte = 0
Private Const PASTING_RULE_BOTTOM As Byte = 1
Private Const PASTING_RULE_STOP As Byte = 2


' @description Import data from a migration workbook.
' Shows file picker, asks what to do with data already entered, checks language,
' imports data and metadata, shows report.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
Public Sub HandleImportData(ByVal sourceWkb As Workbook, _
                            ByVal trads As TranslationObject)

    Dim impObj As LLImporter
    Dim meta As ImportMetadata
    Dim appState As ApplicationState
    Dim io As OSFiles
    Dim filePath As String
    Dim impwb As Workbook
    Dim actsh As Worksheet
    Dim pastingRule As Byte
    Dim pasteAtBottom As Boolean
    Dim sameLanguage As Boolean

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

    ' Ask what happens to the rows this linelist already holds. The question is
    ' asked before the busy state goes on, so the user answers a live workbook.
    Set impObj = LLImporter.Create(sourceWkb)
    pastingRule = PastingRuleFor(impObj, trads)
    If pastingRule = PASTING_RULE_STOP Then GoTo EndImport
    pasteAtBottom = (pastingRule = PASTING_RULE_BOTTOM)

    ' Busy state
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=True, _
                            busyCursor:=xlWait, blockSecurity:=False
    Set actsh = ActiveSheet

    ' Open import workbook
    Set impwb = Workbooks.Open(filePath)
    ActiveWindow.WindowState = xlMinimized

    ' Read what the file says about itself, once, and read the file over
    Set meta = ImportMetadata.Create(impwb)
    If Not impObj.CheckImportFile(impwb, meta) Then GoTo RefusedImport

    ' Check the import is in the language of this linelist
    sameLanguage = impObj.HasSameLanguage(meta)
    If Not sameLanguage Then
        If Not KeepGoingOnLanguage(meta, impObj, trads) Then GoTo EndImport
    End If

    ' Import all data
    impObj.ImportData impwb, pasteAtBottom, meta
    impObj.ImportCustomDropdown impwb, pasteAtBottom
    impObj.CompareWithImportFile impwb
    impObj.FinalizeReport

    ' Import migration metadata. These three read the file's own dictionary and
    ' labels, so they run only when the two files are in the same language.
    If sameLanguage Then
        impObj.ImportShowHide impwb, meta
        impObj.ImportEditableLabels impwb, meta
        impObj.ImportSingleValues meta
    End If

    ' Close import workbook
    impwb.Close savechanges:=False
    Set impwb = Nothing

    actsh.Activate
    appState.Restore

    ' Show result
    If impObj.NeedReport Then
        MsgBox trads.TranslatedValue("MSG_FinishImportRep"), _
               vbQuestion + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    Else
        MsgBox trads.TranslatedValue("MSG_FinishImport"), _
               vbOKOnly, trads.TranslatedValue("MSG_Imports")
    End If

    ' Everything the import found, on one worksheet the user can close
    ImportChecking.ShowImportCheckings sourceWkb, impObj.CheckingValues
    Exit Sub

RefusedImport:
    ' The file cannot be read at all. The reason is on the checking worksheet,
    ' because it is too long for a message box and it names what to do about it.
    On Error Resume Next
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    Set impwb = Nothing
    If Not actsh Is Nothing Then actsh.Activate
    If Not appState Is Nothing Then appState.Restore
    On Error GoTo 0
    MsgBox trads.TranslatedValue("MSG_AbortImport"), _
           vbExclamation + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    ImportChecking.ShowImportCheckings sourceWkb, impObj.CheckingValues
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


' @description Ask the user whether to go on when the languages do not match,
' and say which of the three things happened.
'
' The three used to give one message, MSG_NoLanguage, so a user reading "unable
' to find the language" could be looking at a file whose language was found and
' was French against English. The four keys the other two messages need have
' been translated in the workbook all along with nothing reading them.
' @param meta ImportMetadata. What the file being imported says about itself.
' @param impObj LLImporter. The importer, for the language this linelist is in.
' @param trads TranslationObject. Translations for messages.
' @return Boolean. True when the user wants the import to go on.
Private Function KeepGoingOnLanguage(ByVal meta As ImportMetadata, _
                                     ByVal impObj As LLImporter, _
                                     ByVal trads As TranslationObject) As Boolean

    Dim message As String

    ' The file carries no Metadata sheet at all
    If Not meta.Exists Then
        KeepGoingOnLanguage = ( _
            MsgBox(trads.TranslatedValue("MSG_NoMetadata"), _
                   vbExclamation + vbYesNo, _
                   trads.TranslatedValue("MSG_Imports")) = vbNo)
        Exit Function
    End If

    ' The Metadata sheet is there and names no language
    If LenB(meta.Language) = 0 Then
        KeepGoingOnLanguage = ( _
            MsgBox(trads.TranslatedValue("MSG_NoLanguage"), _
                   vbExclamation + vbYesNo, _
                   trads.TranslatedValue("MSG_Imports")) = vbYes)
        Exit Function
    End If

    ' Both languages are known and they differ. Show the user both.
    message = trads.TranslatedValue("MSG_LanguageDifferent") & vbNewLine & _
              trads.TranslatedValue("MSG_ActualLanguage") & " " & _
              impObj.CurrentLanguage & vbNewLine & _
              trads.TranslatedValue("MSG_ImportLanguage") & " " & _
              meta.Language & vbNewLine & vbNewLine & _
              trads.TranslatedValue("MSG_QuitImports")

    KeepGoingOnLanguage = ( _
        MsgBox(message, vbExclamation + vbYesNo, _
               trads.TranslatedValue("MSG_Imports")) = vbYes)
End Function


' @description What an import does with the rows the linelist already holds.
' A linelist holding no user data takes the import from the first row and the
' user is asked nothing. A linelist holding data is asked, and the three answers
' are the three rules: delete everything first, add the import under what is
' there, or stop.
'
' The question used to be asked and the answer used to decide this. The whole
' decision went away in the restructure and False was passed as a literal
' instead, so every import blanked the tables and started at row 1. A user with
' three weeks of entered cases lost them with no warning.
' @param impObj LLImporter. The importer bound to the linelist.
' @param trads TranslationObject. Translations for messages.
' @return Byte. One of the three PASTING_RULE_ values.
Private Function PastingRuleFor(ByVal impObj As LLImporter, _
                                ByVal trads As TranslationObject) As Byte

    Dim answer As Long

    If Not impObj.HasData Then
        PastingRuleFor = PASTING_RULE_EMPTY
        Exit Function
    End If

    answer = MsgBox(trads.TranslatedValue("MSG_DeleteForImport"), _
                    vbExclamation + vbYesNoCancel, _
                    trads.TranslatedValue("MSG_Imports"))

    Select Case answer
    Case vbYes
        PastingRuleFor = PASTING_RULE_EMPTY
    Case vbNo
        PastingRuleFor = PASTING_RULE_BOTTOM
    Case Else
        PastingRuleFor = PASTING_RULE_STOP
    End Select
End Function


' @description Import a geobase from an external workbook.
' Shows file picker, imports geobase data, optionally updates headers and dictionary.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
' @param histoOnly Boolean. When True, imports only historic geobase data.
Public Sub HandleImportGeobase(ByVal sourceWkb As Workbook, _
                               ByVal trads As TranslationObject, _
                               Optional ByVal histoOnly As Boolean = False)

    Dim impObj As LLImporter
    Dim appState As ApplicationState
    Dim io As OSFiles
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
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=True, _
                            busyCursor:=xlWait, blockSecurity:=False

    ' Open geobase workbook
    Set impwb = Workbooks.Open(filePath)
    ActiveWindow.WindowState = xlMinimized

    ' Import geobase
    Set impObj = LLImporter.Create(sourceWkb)
    impObj.ImportGeobase impwb, histoOnly

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
' @param trads TranslationObject. Translations for messages.
Public Sub HandleClearData(ByVal sourceWkb As Workbook, _
                           ByVal trads As TranslationObject)

    Dim impObj As LLImporter
    Dim appState As ApplicationState
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
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    Set impObj = LLImporter.Create(sourceWkb)
    impObj.ClearData

    appState.Restore
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox trads.TranslatedValue("MSG_ErrClearData"), _
           vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Error")
    If Not appState Is Nothing Then appState.Restore
End Sub
