Attribute VB_Name = "FormLogicAdvanced"

'@Folder("Linelist Forms")
'@ModuleDescription("Complete code-behind of F_Advanced -- imports, clears, reset and saved layouts")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@depends LLImporter, ApplicationState, OSFiles, LLdictionary, ShowHide, ShowHideLayout, ShowHideStore, HiddenNames, Passwords, LLGeo, LLTranslation, TranslationObject, LLLog, LinelistEventsManager, EventLinelist, LinelistRun

' This module is the complete code-behind of the F_Advanced form and is
' copied into the form at deployment, so the control callbacks, the
' translation build and the Handle* walks all live here. Callers outside the
' form reach the walks through the form, the qualified route, so the
' standard-module copy of this code is never compiled and its form
' references cost nothing there: EventsLinelistButtons calls
' F_Advanced.HandleResetColumns, and the F_ShowHideSave code-behind calls
' F_Advanced.HandleSaveShowHideLayout and
' F_Advanced.HandleRestoreShowHideLayout.
'
' THE TWO IMPORT WALKS LEFT THIS MODULE. HandleImportData and
' HandleImportGeobase live in LinelistRun, a standard module that is compiled
' everywhere and can be tested. They were reachable only through the form while
' they sat here, so a script had to open a form to drive an import and no suite
' could ever reach them.

Option Explicit

Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const PRINT_PREFIX As String = "print_"
Private Const CRF_PREFIX As String = "crf_"

' How many clicks on the form surface open the debugging password prompt. The
' setup asks for five on its own import form; the owner asked for seven here.
Private Const DEBUG_CLICKS As Long = 7
Private Const DEBUG_TITLE As String = "Debugging Password"

Private tradform As TranslationObject 'Translation of forms
Private tradmess As TranslationObject 'Translation of messages
Private currwb As Workbook
Private numberOfClicks As Long


'@section Initialization and control callbacks
'===============================================================================

'Initialize the two translation objects.
'
'The translation helper is the one EventLinelist holds. This module used to
'build its own on every open, and LLTranslation.Create validates all five
'translation tables per build.
Private Sub InitializeTrads()
    Dim linelistEvents As EventLinelist
    Dim lltrads As LLTranslation

    Set currwb = ThisWorkbook

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If Not linelistEvents Is Nothing Then Set lltrads = linelistEvents.Translation()

    If lltrads Is Nothing Then _
        Err.Raise ProjectError.ObjectNotInitialized, "FormLogicAdvanced", _
                  "This linelist carries no usable translation sheet"

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
    LinelistRun.HandleImportGeobase currwb, tradmess, histoOnly:=True
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

'Write the user log out as a text file the user can send on
Private Sub CMD_ExportLog_Click()
    Me.Hide
    HandleExportLog currwb
End Sub

'Leave the advanced form
Private Sub CMD_ImportMigQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.Width = 200
    Me.Height = 450
End Sub

' The counter starts again on every open. Initialize runs once per form
' instance and this form is shown on the predeclared instance and hidden
' rather than unloaded, so six clicks left behind by an earlier open would
' otherwise carry over. The caption is put back here too, because the hint
' below overwrites it.
Private Sub UserForm_Activate()
    numberOfClicks = 0
    Me.Caption = tradform.TranslatedValue(Me.Name)
End Sub


'@section Debug mode
'===============================================================================

' Seven clicks on the form surface, then the debugging password, and every
' protection in the workbook comes off. This is the setup's ImportForm.LabPath
' walk on the linelist side, with three differences forced by what is here:
'
' - The clicks land on the FORM, not on a label. F_Advanced carries eight
'   command buttons and nothing else, so there is no LabProgress to write the
'   hint into. The title bar carries it and UserForm_Activate puts the
'   translated caption back.
' - The password manager is the one the event service holds, so opening this
'   form does not build a second one over the same worksheet.
' - The busy state goes through LinelistEventsManager, which counts its
'   nesting, rather than the setup's EventsManager.
'
' The strings are English literals. The workbook carries no translation keys
' for any of this, and the setup states its own the same way.
'
' THERE IS NO WAY BACK OUT BY HAND, and that is the design. Debug mode ends
' when the workbook closes: Passwords.EnsureDebugExitHandler injects
' LeaveDebugModeOnClose into the output workbook and calls it from
' Workbook_BeforeClose, so the protection matrix is reapplied whatever the
' user did in between. Only the sheets the protection table lists come back
' protected -- a sheet protected by hand while in debug mode does not.
Private Sub UserForm_Click()

    Dim pass As Passwords
    Dim answer As Variant
    Dim expected As String
    Dim failDetail As String

    numberOfClicks = numberOfClicks + 1

    If numberOfClicks = (DEBUG_CLICKS - 1) Then
        Me.Caption = "Click the form once more for debug mode"
        Exit Sub
    End If

    If numberOfClicks < DEBUG_CLICKS Then Exit Sub

    numberOfClicks = 0
    Me.Caption = tradform.TranslatedValue(Me.Name)

    ' A workbook with no usable keys has no debugging password to check
    ' against, and the seven clicks do nothing at all.
    Set pass = PasswordManagerOf()
    If pass Is Nothing Then Exit Sub

    If pass.IsInDebugMode() Then
        MsgBox "This linelist is already in debug mode.", _
               vbInformation + vbOKOnly, DEBUG_TITLE
        Exit Sub
    End If

    expected = pass.Value("debuggingpassword")

    answer = Application.InputBox("Enter the debugging password.", _
                                  DEBUG_TITLE, Type:=2)

    ' A cancelled InputBox answers the Boolean False, and a typed password is
    ' always a String, so the type alone says the user backed out.
    If VarType(answer) = vbBoolean Then Exit Sub

    If StrComp(CStr(answer), expected, vbBinaryCompare) <> 0 Then
        LogWarningLine "enter-debug-mode", "wrong password"
        MsgBox "Incorrect password.", vbExclamation + vbOKOnly, DEBUG_TITLE
        Exit Sub
    End If

    ' EnterDebugMode walks every worksheet in the workbook unprotecting it,
    ' and the busy state is what keeps that walk off the screen.
    On Error GoTo DebugFailed
    LinelistEventsManager.LLEnterBusyState
    pass.EnterDebugMode currwb
    LinelistEventsManager.LLExitBusyState
    On Error GoTo 0

    LogWarningLine "enter-debug-mode", "protections removed"
    MsgBox "The linelist is in debug mode. Every protection comes back " & _
           "when the workbook is closed.", _
           vbInformation + vbOKOnly, DEBUG_TITLE
    Me.Hide
    Exit Sub

DebugFailed:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LinelistEventsManager.LLExitBusyState
    Application.Cursor = xlNorthwestArrow
    LogFailureLine "enter-debug-mode", failDetail
    On Error GoTo 0
    MsgBox "Unable to enter debug mode.", vbCritical + vbOKOnly, DEBUG_TITLE
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

' @description Recalculate the cells whose formulas read the geobase. The
'              walk itself lives on EventLinelist, where the harness measures
'              it; this keeps the form side to one guarded call.
Private Sub RecalculateGeoCells()
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Sub

    linelistEvents.RecalculateGeoColumns
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


' @description Open the user log's stopwatch on a layout walk. Guarded, and
' quiet on a workbook with no log, like every log line here.
Private Sub StartLayoutWalk()
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.StartWalk
    On Error GoTo 0
End Sub


' @description Name the step the layout walk is on.
Private Sub MarkLayoutStep(ByVal stepName As String)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.MarkStep stepName
    On Error GoTo 0
End Sub


' @description Write the step line of the layout walk and close it. A walk that
' was never opened writes nothing.
Private Sub LogLayoutSteps(ByVal action As String)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSteps action
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


' @description Write the user log out as a plain text file, so a user who hits
' a problem can send the file on instead of describing what happened. The file
' carries the whole Metadata worksheet above the entries, so it says which
' linelist it came from, which designer built it and on which platform, and
' every entry carries the platform it was written on.
' The walk holds no busy state: a folder picker needs the screen, and the write
' is one small file.
' @param sourceWkb Workbook. The linelist workbook.
Public Sub HandleExportLog(ByVal sourceWkb As Workbook)

    Dim logStore As LLLog
    Dim folderPath As String
    Dim filePath As String
    Dim failDetail As String

    On Error GoTo ErrHand

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    folderPath = PickFolder()
    ' A cancelled picker is an ordinary end, and it is worth a line: a user
    ' who says the export did nothing has the record of their own cancel.
    If LenB(folderPath) = 0 Then
        LogWarningLine "export-log", "no folder was picked"
        Exit Sub
    End If

    filePath = logStore.ExportText(folderPath, BaseNameOf(sourceWkb))

    ' Logged before the box, so the line is on the sheet whatever the user
    ' does with the box. It is NOT in the file just written, which is the
    ' honest order: the file was closed before this line existed.
    LogSuccessLine "export-log", FileNameOf(filePath)

    If Not tradmess Is Nothing Then
        MsgBox filePath, vbInformation, tradmess.TranslatedValue("MSG_Done")
    End If
    Exit Sub

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "export-log", failDetail
End Sub


' @description Ask the user for the folder a written file lands in.
' @return String. The chosen folder, or an empty string on cancel.
Private Function PickFolder() As String
    Dim io As OSFiles

    Set io = OSFiles.Create()
    io.LoadFolder
    If io.HasValidFolder Then PickFolder = io.Folder()
End Function


' @description The workbook name without its extension, for a written file to
' be named after. A workbook that is Nothing answers a plain name rather than
' ending the walk, because the file is worth having either way.
' @param wkb Workbook. The linelist workbook.
' @return String. The name to build a file name on.
Private Function BaseNameOf(ByVal wkb As Workbook) As String
    Dim wkbName As String
    Dim dotAt As Long

    BaseNameOf = "linelist"
    If wkb Is Nothing Then Exit Function

    wkbName = wkb.Name
    dotAt = InStrRev(wkbName, ".")
    If dotAt > 1 Then wkbName = Left$(wkbName, dotAt - 1)
    If LenB(Trim$(wkbName)) = 0 Then Exit Function

    BaseNameOf = wkbName
End Function


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

    'One step per worksheet, named before LayerContextOf, because that call
    'reads the hidden names of EVERY sheet to find its layer and the log is
    'where that cost shows. A sheet with no layer is a short step.
    StartLayoutWalk

    For Each sh In sourceWkb.Worksheets
        MarkLayoutStep sh.Name
        If LayerContextOf(sh, dict, pass, entries, layout) Then
            'The sizes come from the saved layout, so they are written
            matched = store.Load(entries, layout, layoutName, writeSizes:=True)

            If matched > 0 Then
                entries.Apply layout
                store.Save entries, layout
                total = total + matched
                LogRefusedWritesLine "layout-restore", layout, sh.Name
            End If
        End If
    Next

    LogSuccessLine "layout-restore", layoutName & ": " & total & " rows"
    LogLayoutSteps "layout-restore"
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
