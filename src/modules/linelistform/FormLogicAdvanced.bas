Attribute VB_Name = "FormLogicAdvanced"

'@Folder("Linelist Forms")
'@ModuleDescription("Complete code-behind of F_Advanced -- imports, clears, reset and saved layouts")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@depends LLImporter, ImportMetadata, ApplicationState, OSFiles, LLdictionary, ShowHide, ShowHideLayout, ShowHideStore, HiddenNames, Passwords, LLGeo, LLTranslation, TranslationObject, LLLog

' This module is the complete code-behind of the F_Advanced form and is
' copied into the form at deployment, so the control callbacks, the
' translation build and the Handle* walks all live here. Callers outside the
' form reach the walks through the form, the qualified route, so the
' standard-module copy of this code is never compiled and its form
' references cost nothing there: EventsLinelistButtons calls
' F_Advanced.HandleImportData, F_Advanced.HandleImportGeobase and
' F_Advanced.HandleResetColumns, and the F_ShowHideSave code-behind calls
' F_Advanced.HandleSaveShowHideLayout and
' F_Advanced.HandleRestoreShowHideLayout.

Option Explicit

Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const LLSHEET As String = "LinelistTranslation"
Private Const PRINT_PREFIX As String = "print_"
Private Const CRF_PREFIX As String = "crf_"

' What an import does with the rows the linelist already holds
Private Const PASTING_RULE_EMPTY As Byte = 0
Private Const PASTING_RULE_BOTTOM As Byte = 1
Private Const PASTING_RULE_STOP As Byte = 2

Private tradform As TranslationObject 'Translation of forms
Private tradmess As TranslationObject 'Translation of messages
Private currwb As Workbook


'@section Initialization and control callbacks
'===============================================================================

'Initialize the two translation objects
Private Sub InitializeTrads()
    Dim lltrads As LLTranslation

    Set currwb = ThisWorkbook
    Set lltrads = LLTranslation.Create(currwb.Worksheets(LLSHEET))
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
End Sub

'Open the export data form for exports
Private Sub CMD_ExportData_Click()
    Me.Hide
    ClickExportMigration
End Sub

'Import the historic of the geobase alone
Private Sub CMD_ImportGeoHistoric_Click()
    Me.Hide
    HandleImportGeobase currwb, tradmess, histoOnly:=True
End Sub

'Show the report of the last import
Private Sub CMD_ImportMigRep_Click()
    Me.Hide
    F_ImportRep.Show
End Sub

'Clear all the data in the current workbook
Private Sub CMD_ClearData_Click()
    Me.Hide
    HandleClearData currwb, tradmess
End Sub

'Clear the historic of the geobase
Private Sub CMD_ClearGeo_Click()
    Dim geoObj As LLGeo

    'The event service holds the one geobase manager of the workbook. It
    'answers Nothing when the build failed, where LLGeo.Create used to raise,
    'so the failure is reported through the box this form already shows.
    Set geoObj = GeoOf()
    If geoObj Is Nothing Then
        MsgBox tradmess.TranslatedValue("MSG_ErrImportGeo"), _
               vbCritical + vbOKOnly, _
               tradmess.TranslatedValue("MSG_DeleteHistoric")
        Exit Sub
    End If

    If MsgBox(tradmess.TranslatedValue("MSG_HistoricDelete"), _
              vbExclamation + vbYesNo, _
              tradmess.TranslatedValue("MSG_DeleteHistoric")) = vbYes Then

        geoObj.ClearHistoric

        MsgBox tradmess.TranslatedValue("MSG_Done"), _
               vbInformation, _
               tradmess.TranslatedValue("MSG_DeleteHistoric")
    End If
End Sub

'Put the show/hide state of every worksheet back to the state at creation.
'ClickResetColumns holds the busy state and hands the walk back to
'HandleResetColumns below.
Private Sub CMD_ResetCols_Click()
    Me.Hide
    ClickResetColumns
End Sub

'Show the user log sheet
Private Sub CMD_OpenLog_Click()
    Me.Hide
    HandleOpenLog currwb
End Sub

'Leave the advanced form
Private Sub CMD_ImportMigQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.Width = 250
    Me.Height = 550
End Sub


' @description Drop the held managers of the event service.
' The walks below rewrite the worksheets those managers were built over: an
' import rewrites the data, the dropdowns and the dictionary metadata, a
' geobase import the Geo sheet, a clear the data tables. The service builds
' fresh managers on the next event. Called on the error paths too, because a
' walk that failed midway has already rewritten part of what the managers read.
Private Sub ResetEventCaches()
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Sub

    linelistEvents.ResetCaches
End Sub


' @description The one geobase manager of the workbook, held by the event
' service and dropped by ResetCaches. The answer lives in a procedure-local:
' a module field here would hold a manager an import leaves stale.
Private Function GeoOf() As LLGeo
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set GeoOf = linelistEvents.GeoManager()
End Function


' @description The user log the event service holds. A workbook whose log
' cannot be built answers Nothing and every log line below stays quiet.
Private Function UserLogOf() As LLLog
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    Set UserLogOf = linelistEvents.UserLog()
End Function


' @description The password manager the event service holds. A workbook with
' no usable keys answers Nothing and the visibility walk below runs bare.
Private Function PasswordManagerOf() As Passwords
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set PasswordManagerOf = linelistEvents.PasswordManager()
End Function


' @description Write the success line of a finished walk. The write is
' guarded so a log fault never takes down the walk it records.
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


' @description Write the warning line of a refused or cancelled walk.
Private Sub LogWarningLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogWarning action, detail
    On Error GoTo 0
End Sub


' @description Write the warning line of one sheet whose layout ended its
' walk with refused writes, so a protected sheet or a position that is gone
' shows up in the log with its count.
Private Sub LogRefusedWritesLine(ByVal action As String, _
                                 ByVal layout As ShowHideLayout, _
                                 ByVal sheetName As String)
    If layout Is Nothing Then Exit Sub
    If layout.FailureCount = 0 Then Exit Sub

    LogWarningLine action, sheetName & ": " & layout.FailureCount & " refused writes"
End Sub


' @description The file name at the end of a picked path, for the log detail.
' Both separators are tried, so the helper answers the same on every host.
Private Function FileNameOf(ByVal filePath As String) As String
    Dim sepAt As Long

    sepAt = InStrRev(filePath, "/")
    If sepAt = 0 Then sepAt = InStrRev(filePath, "\")

    FileNameOf = Mid$(filePath, sepAt + 1)
End Function


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
    Dim failDetail As String

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

    ' The import rewrote the sheets the held managers were built over
    ResetEventCaches
    LogSuccessLine "import-data", FileNameOf(filePath)

    ' Show result. MSG_FinishImportRep asks whether the user wants to see a
    ' report, and it used to be asked with an OK button, so there was no way to
    ' answer and nothing behind it either.
    If impObj.NeedReport Then
        If MsgBox(trads.TranslatedValue("MSG_FinishImportRep"), _
                  vbQuestion + vbYesNo, _
                  trads.TranslatedValue("MSG_Imports")) = vbYes Then
            F_ImportRep.Show
        End If
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
    LogWarningLine "import-data", "file refused: " & FileNameOf(filePath)
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
    LogWarningLine "import-data", "cancelled"
    MsgBox trads.TranslatedValue("MSG_AbortImport"), _
           vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not actsh Is Nothing Then actsh.Activate
    If Not appState Is Nothing Then appState.Restore
    On Error GoTo 0
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "import-data", failDetail
    MsgBox trads.TranslatedValue("MSG_ErrorImport"), _
           vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not actsh Is Nothing Then actsh.Activate
    If Not appState Is Nothing Then appState.Restore
    ResetEventCaches
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
    Dim failDetail As String

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

    ' The import rewrote the Geo sheet the held geo manager was built over
    ResetEventCaches
    LogSuccessLine "import-geobase", _
                   FileNameOf(filePath) & IIf(histoOnly, " (historic only)", vbNullString)

    MsgBox trads.TranslatedValue("MSG_FinishImportGeo"), _
           vbOKOnly, trads.TranslatedValue("MSG_Imports")
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "import-geobase", failDetail
    MsgBox trads.TranslatedValue("MSG_ErrImportGeo"), _
           vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not appState Is Nothing Then appState.Restore
    ResetEventCaches
End Sub


' @description Put the show/hide state of every worksheet back where the
' dictionary started it. Each layered worksheet - data entry, printed, VList
' and CRF - gets a fresh entry list, ShowHide.ResetToAuthored applies the
' authored state and the printed header directions, and the saved choices are
' overwritten, so the reset survives the workbook closing. A workbook with no
' show/hide worksheet resets the sheets alone.
' @param sourceWkb Workbook. The linelist workbook.
' @param pass Passwords. The protection keys, or Nothing.
Public Sub HandleResetColumns(ByVal sourceWkb As Workbook, _
                              ByVal pass As Passwords)

    Dim sh As Worksheet
    Dim dict As LLdictionary
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim store As ShowHideStore

    Set dict = LLdictionary.Create(sourceWkb.Worksheets(DICTIONARY_SHEET), 1, 1)

    ' An older linelist carries no show/hide worksheet. The reset still walks
    ' the sheets, and there is simply nothing to overwrite.
    On Error Resume Next
    Set store = ShowHideStore.Create(sourceWkb)
    On Error GoTo 0

    For Each sh In sourceWkb.Worksheets
        If LayerContextOf(sh, dict, pass, entries, layout) Then
            entries.ResetToAuthored layout
            If Not store Is Nothing Then store.Save entries, layout
            LogRefusedWritesLine "reset-columns", layout, sh.Name
        End If
    Next
End Sub


' @description Show the very hidden log sheet and land the user on it, so the
' record of the past runs can be read in place. The workbook structure guards
' sheet visibility, so the walk unprotects the workbook around the change and
' protects it again on both paths. The ribbon close button hides the sheet
' again through ClickCloseSheet. A workbook whose log cannot be built has no
' sheet to show and the walk leaves quietly.
' @param sourceWkb Workbook. The linelist workbook.
Public Sub HandleOpenLog(ByVal sourceWkb As Workbook)

    Dim logStore As LLLog
    Dim pass As Passwords
    Dim failDetail As String

    On Error GoTo ErrHand

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    Set pass = PasswordManagerOf()
    If Not pass Is Nothing Then pass.UnProtect sourceWkb

    logStore.Wksh.Visible = xlSheetVisible
    logStore.Wksh.Activate

    If Not pass Is Nothing Then pass.Protect sourceWkb
    LogSuccessLine "open-log"
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "open-log", failDetail
    If Not pass Is Nothing Then pass.Protect sourceWkb
End Sub


' @description Save what the user sees on every layered worksheet as one named
' layout: visibility, sizes and printed header directions, over all four
' layers. The sheet is the record, so each entry list adopts its sheet before
' the save. The caller holds the busy state.
' @param sourceWkb Workbook. The linelist workbook.
' @param pass Passwords. The protection keys, or Nothing.
' @param layoutName String. The name to save under.
' @return Boolean. True when the layout was saved. False on an empty name, and
' on a new name when the store already holds its maximum of saved layouts.
Public Function HandleSaveShowHideLayout(ByVal sourceWkb As Workbook, _
                                         ByVal pass As Passwords, _
                                         ByVal layoutName As String) As Boolean

    Dim sh As Worksheet
    Dim dict As LLdictionary
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim store As ShowHideStore

    layoutName = Trim$(layoutName)
    If LenB(layoutName) = 0 Then Exit Function

    Set store = ShowHideStore.Create(sourceWkb)
    If Not store.HasLayout(layoutName) Then
        If store.LayoutCount() >= store.MaxSavedLayouts Then Exit Function
    End If

    Set dict = LLdictionary.Create(sourceWkb.Worksheets(DICTIONARY_SHEET), 1, 1)

    For Each sh In sourceWkb.Worksheets
        If LayerContextOf(sh, dict, pass, entries, layout) Then
            entries.Adopt layout
            store.Save entries, layout, layoutName
            LogRefusedWritesLine "layout-save", layout, sh.Name
        End If
    Next

    LogSuccessLine "layout-save", layoutName
    HandleSaveShowHideLayout = True
End Function


' @description Put every layered worksheet in the state one saved layout
' describes, and record that state as the current one, so it survives the
' workbook closing. A sheet the layout has no rows for keeps its state. The
' caller holds the busy state.
' @param sourceWkb Workbook. The linelist workbook.
' @param pass Passwords. The protection keys, or Nothing.
' @param layoutName String. The saved layout to restore.
' @return Long. How many stored rows landed on an entry, 0 when the name is
' unknown.
Public Function HandleRestoreShowHideLayout(ByVal sourceWkb As Workbook, _
                                            ByVal pass As Passwords, _
                                            ByVal layoutName As String) As Long

    Dim sh As Worksheet
    Dim dict As LLdictionary
    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim store As ShowHideStore
    Dim matched As Long
    Dim total As Long

    layoutName = Trim$(layoutName)
    If LenB(layoutName) = 0 Then Exit Function

    Set store = ShowHideStore.Create(sourceWkb)
    If Not store.HasLayout(layoutName) Then Exit Function

    Set dict = LLdictionary.Create(sourceWkb.Worksheets(DICTIONARY_SHEET), 1, 1)

    For Each sh In sourceWkb.Worksheets
        If LayerContextOf(sh, dict, pass, entries, layout) Then
            matched = store.Load(entries, layout, layoutName)

            If matched > 0 Then
                entries.Apply layout
                store.Save entries, layout
                total = total + matched
                LogRefusedWritesLine "layout-restore", layout, sh.Name
            End If
        End If
    Next

    LogSuccessLine "layout-restore", layoutName & ": " & total & " rows"
    HandleRestoreShowHideLayout = total
End Function


' @description Build the entry list and the layout of one worksheet. A sheet
' outside the four show/hide layers answers False and both stay Nothing.
' @param sh Worksheet. The worksheet to read.
' @param dict LLdictionary. The dictionary of the workbook.
' @param pass Passwords. The protection keys, or Nothing.
' @param entries ShowHide. Set to the sheet's entry list.
' @param layout ShowHideLayout. Set to the sheet's layout.
' @return Boolean. True when the sheet carries a layer.
Private Function LayerContextOf(ByVal sh As Worksheet, _
                                ByVal dict As LLdictionary, _
                                ByVal pass As Passwords, _
                                ByRef entries As ShowHide, _
                                ByRef layout As ShowHideLayout) As Boolean

    Dim shNames As HiddenNames
    Dim shType As String
    Dim layer As Byte

    Set entries = Nothing
    Set layout = Nothing

    Set shNames = SheetNamesOf(sh)
    If shNames Is Nothing Then Exit Function

    shType = shNames.ValueAsString("sheet_type")
    layer = ShowHideLayerOf(shType)
    If layer = 0 Then Exit Function

    Set entries = ShowHide.Create(dict, layer, BaseSheetNameOf(sh, shType))
    Set layout = ShowHideLayout.Create(sh, layer, pass, _
                                       BaseTableNameOf(shNames.ValueAsString("table_name")))

    LayerContextOf = True
End Function


' @description The held hidden names of one worksheet, through the event
' service, so the walk shares the stores every button already uses. A sheet
' whose names cannot be read answers Nothing and the walk passes it by.
' @param sh Worksheet. The worksheet whose names are wanted.
' @return HiddenNames. The held store, or Nothing.
Private Function SheetNamesOf(ByVal sh As Worksheet) As HiddenNames
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    Set SheetNamesOf = linelistEvents.SheetNames(sh)
End Function


' @description The show/hide layer of one sheet tag. A sheet outside the four
' layers answers 0 and the reset leaves it alone.
' @param shType String. The sheet_type hidden name value.
' @return Byte. A ShowHideWorksheetLayer value, or 0.
Private Function ShowHideLayerOf(ByVal shType As String) As Byte
    Select Case shType
    Case "HList"
        ShowHideLayerOf = ShowHideLayerHList
    Case "HList Print"
        ShowHideLayerOf = ShowHideLayerPrinted
    Case "VList"
        ShowHideLayerOf = ShowHideLayerVList
    Case "HList CRF"
        ShowHideLayerOf = ShowHideLayerCRF
    Case Else
        ShowHideLayerOf = 0
    End Select
End Function


' @description The base sheet name the dictionary knows, with the print_ or
' crf_ prefix of a companion sheet cut off the front.
' @param sh Worksheet. The worksheet being reset.
' @param shType String. The sheet_type hidden name value.
' @return String. The base sheet name.
Private Function BaseSheetNameOf(ByVal sh As Worksheet, ByVal shType As String) As String
    Select Case shType
    Case "HList Print"
        BaseSheetNameOf = Mid$(sh.Name, Len(PRINT_PREFIX) + 1)
    Case "HList CRF"
        BaseSheetNameOf = Mid$(sh.Name, Len(CRF_PREFIX) + 1)
    Case Else
        BaseSheetNameOf = sh.Name
    End Select
End Function


' @description The table name the PRINTSTART anchor is named after. A printed
' sheet stores its own table name with the print_ prefix in front, and the
' anchor carries none.
' @param tabName String. The table_name hidden name value.
' @return String. The base table name.
Private Function BaseTableNameOf(ByVal tabName As String) As String
    If InStr(1, tabName, PRINT_PREFIX, vbTextCompare) = 1 Then
        tabName = Mid$(tabName, Len(PRINT_PREFIX) + 1)
    End If

    BaseTableNameOf = tabName
End Function


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
    Dim failDetail As String

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

    ' The clear emptied the tables the held managers were built over
    ResetEventCaches
    LogSuccessLine "clear-data"
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "clear-data", failDetail
    MsgBox trads.TranslatedValue("MSG_ErrClearData"), _
           vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Error")
    If Not appState Is Nothing Then appState.Restore
    ResetEventCaches
End Sub
