Attribute VB_Name = "LinelistRun"
Attribute VB_Description = "The two import walks of a linelist, and the entry points a script calls them through"

Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("The two import walks of a linelist, and the entry points a script calls them through")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed
'@depends LLImporter, ImportMetadata, ApplicationState, OSFiles, ImportChecking, LinelistEventsManager, EventLinelist, LLTranslation, LLLog, TranslationObject, Messenger, CustomTable, Passwords, HiddenNames

'WHAT IS IN HERE
'The two import walks of a running linelist -- HandleImportData and
'HandleImportGeobase -- and the three entry points a script calls them through:
'RunImportGeobase, RunImportData and LinelistLastSummary. The wrappers take
'their paths as strings, open no picker and no box, and answer "OK" or
'"ERROR <number>: <text>". LinelistLastSummary reads the last run back.
'
'The buttons of the linelist are unchanged. ClickImportData and
'ClickImportGeobase call the same two walks with no path and no rule, and a
'person clicking still gets every picker and every box they always had.
'
'THE WALKS USED TO LIVE IN A FORM
'-------------------------------------------------------------------------------
'They sat in FormLogicAdvanced, which is the code-behind of F_Advanced and is
'merged into that form at build time. So the only way to reach them was through
'the form -- F_Advanced.HandleImportData -- and that cost two things:
'
' - A script had to open a form to drive an import.
' - No suite could ever reach them. A form module is compiled into the delivered
'   linelist and into nothing the test harness drives, which is why the thirteen
'   message boxes of these two walks had no test at all.
'
'They are a standard module now, so they compile everywhere and can be tested.
'
'THIS MODULE NAMES NO FORM
'-------------------------------------------------------------------------------
'A form reference is a COMPILE time reference: a workbook that does not carry
'the form stops compiling altogether, and this module is imported into the test
'driver workbook, which carries no forms. The one form these walks show --
'F_ImportRep, when a person answers yes to "open the report now?" -- is reached
'by name through UserForms.Add. That is the route
'EventsLinelistButtons.ClickOpenShowHideSections already takes.
'
'IT CANNOT CALL EventsLinelistButtons EITHER, for the same reason: that module
'names seven forms. So the pre-import walk of ClickImportData is copied below
'rather than called.
'
'EVERY PATH OF A WRAPPER DISARMS
'-------------------------------------------------------------------------------
'CloseRun is the one exit of both wrappers: it reads the swallowed boxes off the
'messenger, disarms, writes the summary file and answers the outcome. A run that
'refuses at its guard goes through it too, so a bad path cannot leave the next
'caller's boxes swallowed -- and that next caller is a person clicking a button.
'
'THE SUMMARY IS WRITTEN TWICE
'-------------------------------------------------------------------------------
'LinelistLastSummary reads the run back through a second Application.Run, and
'that is the reading lost whenever the Apple Event transport gives up, which
'happens on runs Excel finished green. So the same text is written to
'<folder>/<name>-obt-summary.txt beside the file the wrapper touched.
'
'A RUN THAT IMPORTED SAVES THE LINELIST AND CLOSES IT
'-------------------------------------------------------------------------------
'The import is the whole point of the call, and a script has nobody to press
'save. So CloseRun saves the workbook and closes it once the outcome reads OK.
'
'Closing the workbook ends the code running inside it, so that run answers
'Application.Run nothing at all and LinelistLastSummary can no longer be read
'back. THE SUMMARY FILE IS THE READING after an import, which is why it is
'written before the save rather than after.
'
'A run that failed leaves the workbook open. Nothing was imported, the caller
'may want to look at it or try again, and its summary file is already on disk.

'THE MODULE LEVEL DECLARATIONS
'A module-level declaration has to sit in this section, above every procedure.
'VBA registers no declaration written between two procedures, and every use of
'it then reads as an undefined variable under Option Explicit.

' What an import does with the rows the linelist already holds
Private Const PASTING_RULE_EMPTY As Byte = 0
Private Const PASTING_RULE_BOTTOM As Byte = 1
Private Const PASTING_RULE_STOP As Byte = 2

'The two words a caller writes for what an import does with the rows the
'linelist already holds. An empty word means append, which is the answer that
'keeps data.
Private Const RULE_REPLACE As String = "replace"
Private Const RULE_APPEND As String = "append"

'The two words a caller writes for force. Yes is what lets a scripted import
'push past the three language warnings, and the caller has to ask for it by name.
Private Const FORCE_YES As String = "Yes"
Private Const FORCE_NO As String = "No"

'What a wrapper answers when the run went through.
Private Const OUTCOME_OK As String = "OK"

'What an outcome that failed opens with.
Private Const OUTCOME_ERROR_LEAD As String = "ERROR "

'What the summary file is called, after the name of the file the run touched.
Private Const SUMMARY_SUFFIX As String = "-obt-summary.txt"

'The marker the free text starts after, the shape HeadlessBuild already answers.
Private Const REPORT_MARKER As String = "--report--"

'The sheet type of a data entry table the import writes rows into.
Private Const HLIST_TAG As String = "HList"

'What the last wrapper run read, wrote and said. LinelistLastSummary answers
'these, and they are module level because that reading is a second
'Application.Run.
Private lastOutcome As String
Private lastGeobaseFile As String
Private lastImportFile As String
Private lastExportFile As String
Private lastMessages As String


'@section The two import walks
'===============================================================================
'The bodies that moved here out of the F_Advanced code-behind. The buttons call
'them with no path and no rule; the wrappers below call them with both.

' @description Import data from a migration workbook.
' Shows file picker, asks what to do with data already entered, checks language,
' imports data and metadata, writes what the import found.
'
' ClickImportData is the button, RunImportData below is the script. Both call
' this, and the two arguments at the end are what tells them apart.
' Everything the import found goes onto the import checking worksheet in ONE
' write, placed after the import work and before any box: the report form is
' offered after that write, so its Open Log button has a sheet to show. A
' refused file is the only path that puts the sheet on screen. A cancel and an
' error write no worksheet at all - the workbook log already carries the line.
'
' A CALLER CAN HAND IN THE FILE AND THE RULE
' Empty chosenPath means "ask", which is what the button passes. Filled means
' "use this file", and the picker never opens. chosenRule is the same shape for
' the question about the rows already entered: empty asks, "replace" and
' "append" answer.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
' @param chosenPath String. Optional. The file to import. Empty opens the picker.
' @param chosenRule String. Optional. "replace" or "append". Empty asks.
' @return Boolean. True when the import ran through and wrote its data. Every
' other way out answers False, which is what tells a scripted caller it failed.
Public Function HandleImportData(ByVal sourceWkb As Workbook, _
                                 ByVal trads As TranslationObject, _
                                 Optional ByVal chosenPath As String = vbNullString, _
                                 Optional ByVal chosenRule As String = vbNullString) As Boolean

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
    Dim refused As Boolean
    Dim failDetail As String

    On Error GoTo ErrHand

    ' Select import file. A caller that named one is taken at its word: the
    ' picker also proved the file was there, and a path handed in has proved
    ' nothing, so RunImportData checks the file before it calls.
    If LenB(chosenPath) = 0 Then
        Set io = OSFiles.Create()
        io.LoadFile "*.xlsx"
        If Not io.HasValidFile Then Exit Function
        filePath = io.File()
    Else
        filePath = chosenPath
    End If

    ' Confirm import. A script asked for this import, so confirming it is the
    ' silent answer: the box is vbOKCancel and the answer is vbOK.
    If Messenger.Show(trads.TranslatedValue("MSG_ImportConfirm"), vbOK, _
                      vbOKCancel, _
                      trads.TranslatedValue("MSG_Confirm")) = vbCancel Then
        GoTo EndImport
    End If

    ' Ask what happens to the rows this linelist already holds. The question is
    ' asked before the busy state goes on, so the user answers a live workbook.
    Set impObj = LLImporter.Create(sourceWkb)
    pastingRule = PastingRuleFor(impObj, trads, chosenRule)
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
    refused = Not impObj.CheckImportFile(impwb, meta)

    ' A refused file is read no further, and the walk carries on to the single
    ' write below rather than leaving through its own exit: the reason the file
    ' was turned away is itself something the import found, and it belongs on
    ' the same worksheet as everything else it found.
    If Not refused Then

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

        ' Import migration metadata. These three read the file's own dictionary
        ' and labels, so they run only when the two files are in the same
        ' language.
        If sameLanguage Then
            impObj.ImportShowHide impwb, meta
            impObj.ImportEditableLabels impwb, meta
            impObj.ImportSingleValues meta
        End If
    End If

    ' Close import workbook
    impwb.Close savechanges:=False
    Set impwb = Nothing

    actsh.Activate
    appState.Restore

    ' Everything the import found, written onto the worksheet and left hidden.
    ' ONE write, and it happens HERE: right after the import work, before any
    ' box and before the report form. The Open Log button of that form shows
    ' this worksheet, so a sheet written after the form was dismissed is a
    ' button that does nothing.
    ImportChecking.WriteImportCheckings sourceWkb, impObj.CheckingValues

    ' A REFUSAL IS SHOWN, unlike a finished import. Nothing else tells the user
    ' why the file was turned away: the box carries one sentence, the import
    ' report form is never offered on this path, and the worksheet holds the
    ' reason and what to do about it. This is the only place the sheet is put on
    ' screen by an import.
    If refused Then
        LogWarningLine "import-data", "file refused: " & FileNameOf(filePath)
        Messenger.Show trads.TranslatedValue("MSG_AbortImport"), vbOK, _
                       vbExclamation + vbOKOnly, _
                       trads.TranslatedValue("MSG_Imports")
        ImportChecking.ShowReportSheet sourceWkb
        Exit Function
    End If

    ' The import rewrote the sheets the held managers were built over
    ResetEventCaches
    LogSuccessLine "import-data", FileNameOf(filePath)

    ' The one line that answers True, and it sits after the data is written.
    HandleImportData = True

    ' Show result. MSG_FinishImportRep asks whether the user wants to see a
    ' report, and it used to be asked with an OK button, so there was no way to
    ' answer and nothing behind it either.
    ' The report is a worksheet on screen and nobody is there to read it while
    ' a script drives, so the silent answer to the question is vbNo.
    If impObj.NeedReport Then
        If Messenger.Show(trads.TranslatedValue("MSG_FinishImportRep"), vbNo, _
                          vbQuestion + vbYesNo, _
                          trads.TranslatedValue("MSG_Imports")) = vbYes Then
            ShowImportReportForm
        End If
    Else
        Messenger.Show trads.TranslatedValue("MSG_FinishImport"), vbOK, _
                       vbOKOnly, trads.TranslatedValue("MSG_Imports")
    End If
    Exit Function

EndImport:
    On Error Resume Next
    LogWarningLine "import-data", "cancelled"
    Messenger.Show trads.TranslatedValue("MSG_AbortImport"), vbOK, _
                   vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not actsh Is Nothing Then actsh.Activate
    If Not appState Is Nothing Then appState.Restore
    On Error GoTo 0
    Exit Function

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "import-data", failDetail
    Messenger.Show trads.TranslatedValue("MSG_ErrorImport"), vbOK, _
                   vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not actsh Is Nothing Then actsh.Activate
    If Not appState Is Nothing Then appState.Restore
    ResetEventCaches
End Function


' @description Show the import report form the user asked for.
' The form is reached by name rather than written into the code. Naming
' F_ImportRep here would be a COMPILE time reference, and this module is
' imported into workbooks that carry no form at all -- the test driver among
' them. UserForms.Add resolves the name when the user answers yes, and a
' workbook without the form simply shows nothing. The precedent is
' EventsLinelistButtons.ClickOpenShowHideSections, which reaches
' F_ShowHideSections the same way.
'
' This runs only when a person clicked Yes. A script never reaches it: the
' silent answer to that box is vbNo, because nobody is there to read a sheet.
Private Sub ShowImportReportForm()
    Dim reportForm As Object

    On Error Resume Next
    Set reportForm = UserForms.Add("F_ImportRep")
    On Error GoTo 0

    If reportForm Is Nothing Then Exit Sub

    reportForm.Show
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

    ' The file carries no Metadata sheet at all. MSG_NoMetadata asks whether to
    ' QUIT, and the two boxes below ask whether to CONTINUE. Same warning, read
    ' from opposite ends, so this one is answered by QuitUnlessForced.
    If Not meta.Exists Then
        KeepGoingOnLanguage = ( _
            Messenger.Show(trads.TranslatedValue("MSG_NoMetadata"), _
                           QuitUnlessForced(), _
                           vbExclamation + vbYesNo, _
                           trads.TranslatedValue("MSG_Imports")) = vbNo)
        Exit Function
    End If

    ' The Metadata sheet is there and names no language. The box asks whether
    ' to import anyway, so Messenger.CarryOn answers it: a forced run imports,
    ' every other run stops and the wrapper says why.
    If LenB(meta.Language) = 0 Then
        KeepGoingOnLanguage = ( _
            Messenger.Show(trads.TranslatedValue("MSG_NoLanguage"), _
                           Messenger.CarryOn(), _
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

    ' The box asks whether to continue, so Messenger.CarryOn answers it.
    KeepGoingOnLanguage = ( _
        Messenger.Show(message, Messenger.CarryOn(), _
                       vbExclamation + vbYesNo, _
                       trads.TranslatedValue("MSG_Imports")) = vbYes)
End Function


' @description The silent answer to a warning box that asks whether to QUIT.
' Messenger.CarryOn answers a box that asks whether to CONTINUE: vbNo on a
' normal run, vbYes on a forced one. MSG_NoMetadata asks the same question the
' other way round, so its answer is the other way round too. A scripted import
' quits on the missing metadata unless the caller asked for force, which is the
' same rule the other two language warnings follow.
' @return VbMsgBoxResult. vbYes on a normal run, vbNo on a forced one.
Private Function QuitUnlessForced() As VbMsgBoxResult
    QuitUnlessForced = vbYes
    If Messenger.CarryOn() = vbYes Then QuitUnlessForced = vbNo
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
'
' A CALLER CAN ANSWER IT INSTEAD. chosenRule carries the word RunImportData was
' given, and a word is only ever read on a linelist that holds data: an empty
' linelist takes the import from the first row whoever asked for it, so
' "append" and "replace" mean the same thing there and neither is worth acting
' on. ResolvedRule has already checked the word against these two, so an
' unknown word cannot reach the Else below.
' @param impObj LLImporter. The importer bound to the linelist.
' @param trads TranslationObject. Translations for messages.
' @param chosenRule String. "replace" or "append", or empty to ask.
' @return Byte. One of the three PASTING_RULE_ values.
Private Function PastingRuleFor(ByVal impObj As LLImporter, _
                                ByVal trads As TranslationObject, _
                                ByVal chosenRule As String) As Byte

    Dim answer As Long

    If Not impObj.HasData Then
        PastingRuleFor = PASTING_RULE_EMPTY
        Exit Function
    End If

    If LenB(chosenRule) > 0 Then
        If StrComp(chosenRule, RULE_REPLACE, vbTextCompare) = 0 Then
            PastingRuleFor = PASTING_RULE_EMPTY
        Else
            PastingRuleFor = PASTING_RULE_BOTTOM
        End If
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
'
' ClickImportGeobase is the button, RunImportGeobase below is the script, and
' the CMD_ImportGeoHistoric button of F_Advanced is the third caller.
'
' A CALLER CAN HAND IN THE FILE. Empty chosenPath means "ask", which is what
' the two buttons pass. Filled means "use this file", and the picker never
' opens.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads TranslationObject. Translations for messages.
' @param histoOnly Boolean. When True, imports only historic geobase data.
' @param chosenPath String. Optional. The geobase to read. Empty opens the picker.
' @return Boolean. True when the geobase was read in. Every other way out
' answers False, which is what tells a scripted caller it failed.
Public Function HandleImportGeobase(ByVal sourceWkb As Workbook, _
                                    ByVal trads As TranslationObject, _
                                    Optional ByVal histoOnly As Boolean = False, _
                                    Optional ByVal chosenPath As String = vbNullString) As Boolean

    Dim impObj As LLImporter
    Dim appState As ApplicationState
    Dim io As OSFiles
    Dim filePath As String
    Dim impwb As Workbook
    Dim failDetail As String

    On Error GoTo ErrHand

    ' Select geobase file. A caller that named one is taken at its word, the
    ' same way HandleImportData above takes the path it is given.
    If LenB(chosenPath) = 0 Then
        Set io = OSFiles.Create()
        io.LoadFile "*.xlsx"
        If Not io.HasValidFile Then Exit Function
        filePath = io.File()
    Else
        filePath = chosenPath
    End If

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

    ' The workbook runs on manual calculation, so the p-code and concat
    ' columns of the data entry sheets, and the admin level labels of the
    ' spatial dropdowns, keep the values of the geobase before until they are
    ' recalculated. A historic-only import leaves the level tables alone.
    If Not histoOnly Then RecalculateGeoCells
    LogSuccessLine "import-geobase", _
                   FileNameOf(filePath) & IIf(histoOnly, " (historic only)", vbNullString)

    ' The one line that answers True, and it sits after the geobase is read in.
    HandleImportGeobase = True

    Messenger.Show trads.TranslatedValue("MSG_FinishImportGeo"), vbOK, _
                   vbOKOnly, trads.TranslatedValue("MSG_Imports")
    Exit Function

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    LogFailureLine "import-geobase", failDetail
    Messenger.Show trads.TranslatedValue("MSG_ErrImportGeo"), vbOK, _
                   vbCritical + vbOKOnly, trads.TranslatedValue("MSG_Imports")
    If Not impwb Is Nothing Then impwb.Close savechanges:=False
    If Not appState Is Nothing Then appState.Restore
    ResetEventCaches
End Function

'@section The entry points a script calls
'===============================================================================

'@sub-title Read a geobase into this linelist, with no picker and no box.
'@details
'The body of ClickImportGeobase, with the file handed in instead of picked. The
'picker also proved the file was there; a path from a script has proved nothing,
'so the file is checked before any of the work starts.
'
'A historic-only import is the CMD_ImportGeoHistoric button and is not offered
'here: a script asking for a geobase wants the whole geobase.
'
'A run that read the geobase in saves this linelist and closes it, so this
'answers "OK" to nobody. Read the summary file. A run that failed answers its
'refusal and leaves the workbook open.
'@param geobasePath String. Full path of the .xlsx to read.
'@return String. "OK", or "ERROR <number>: <text>". A run that worked returns
'no answer at all, because the workbook it was called in is closed by then.
'@EntryPoint
Public Function RunImportGeobase(ByVal geobasePath As String) As String
    Dim trads As TranslationObject
    Dim sourcePath As String
    Dim outcome As String
    Dim imported As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    Messenger.Arm ThisWorkbook
    ResetRunRecord

    sourcePath = Trim$(geobasePath)

    If Not FileIsThere(sourcePath) Then
        RunImportGeobase = CloseRun(RunError(0, "no file at " & sourcePath), _
                                    ThisWorkbook.Path, BaseNameOf(ThisWorkbook.Name))
        Exit Function
    End If

    On Error GoTo Handler

    Set trads = MessagesTranslator()
    If trads Is Nothing Then
        RunImportGeobase = CloseRun(RunError(0, "this linelist carries no usable translation sheet"), _
                                    ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
        Exit Function
    End If

    imported = HandleImportGeobase(ThisWorkbook, trads, False, sourcePath)

    If imported Then
        lastGeobaseFile = sourcePath
        outcome = OUTCOME_OK
    Else
        outcome = RunError(0, "the geobase at " & sourcePath & " was not read in")
    End If

    RunImportGeobase = CloseRun(outcome, ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
    Exit Function

Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    Debug.Print "RunImportGeobase: "; errNumber; errDescription

    RunImportGeobase = CloseRun(RunError(errNumber, errDescription), _
                                ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
End Function

'@sub-title Read a migration workbook into this linelist, with no picker and no box.
'@details
'The body of ClickImportData, with the file and the pasting rule handed in
'instead of asked for. The events stay off across the whole walk, the data entry
'tables are trimmed of their blank rows first and the automatic lists are rebuilt
'after, exactly as the button does.
'
'FORCE IS WHAT PUSHES A RUN PAST A WARNING IT CANNOT JUDGE. Three boxes of the
'import warn that the file may not match this linelist -- no metadata, no
'language recorded, a different language. A person can look at the file and
'judge; a script cannot. So an unforced run stops on all three and this answers
'why, and a caller that has already judged the file writes force:="Yes".
'A run that imported saves this linelist and closes it, so this answers "OK" to
'nobody. Read the summary file. A run that failed answers its refusal and leaves
'the workbook open.
'@param importPath String. Full path of the .xlsx to read.
'@param pastingRule String. "replace" wipes the rows this linelist holds and
'starts at row 1; "append" puts the import under them. Empty means "append".
'@param force String. Optional, "Yes" or "No", default "No".
'@return String. "OK", or "ERROR <number>: <text>". A run that worked returns
'no answer at all, because the workbook it was called in is closed by then.
'@EntryPoint
Public Function RunImportData(ByVal importPath As String, _
                              ByVal pastingRule As String, _
                              Optional ByVal force As String = FORCE_NO) As String
    Dim trads As TranslationObject
    Dim keys As Passwords
    Dim sourcePath As String
    Dim chosenRule As String
    Dim outcome As String
    Dim forced As Boolean
    Dim forceIsAWord As Boolean
    Dim imported As Boolean
    Dim quietStateOn As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    'The force word is read before the messenger is armed, because arming is
    'what carries it. A word that is neither Yes nor No is refused below rather
    'than read as No: a typo must never quietly turn force off.
    forced = (StrComp(Trim$(force), FORCE_YES, vbTextCompare) = 0)
    forceIsAWord = forced Or (StrComp(Trim$(force), FORCE_NO, vbTextCompare) = 0)

    Messenger.Arm ThisWorkbook, force:=forced
    ResetRunRecord

    sourcePath = Trim$(importPath)

    If Not FileIsThere(sourcePath) Then
        RunImportData = CloseRun(RunError(0, "no file at " & sourcePath), _
                                 ThisWorkbook.Path, BaseNameOf(ThisWorkbook.Name))
        Exit Function
    End If

    chosenRule = ResolvedRule(pastingRule)
    If LenB(chosenRule) = 0 Then
        RunImportData = CloseRun(RunError(0, "pasting rule """ & pastingRule & _
                                             """ is neither " & RULE_REPLACE & " nor " & RULE_APPEND), _
                                 ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
        Exit Function
    End If

    If Not forceIsAWord Then
        RunImportData = CloseRun(RunError(0, "force """ & force & """ is neither " & _
                                             FORCE_YES & " nor " & FORCE_NO), _
                                 ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
        Exit Function
    End If

    On Error GoTo Handler

    Set trads = MessagesTranslator()
    If trads Is Nothing Then
        RunImportData = CloseRun(RunError(0, "this linelist carries no usable translation sheet"), _
                                 ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
        Exit Function
    End If

    'Events stay off across the whole import, through the events manager, which
    'is the one owner of that flag in a running linelist. The trim below writes
    'to every data entry sheet and the import writes thousands of rows after it.
    LinelistEventsManager.LLEnterQuietState
    quietStateOn = True

    Set keys = PasswordManagerOf()
    TrimDataTables ThisWorkbook, keys

    imported = HandleImportData(ThisWorkbook, trads, sourcePath, chosenRule)

    LinelistEventsManager.UpdateAllListAuto

    LinelistEventsManager.LLExitQuietState
    quietStateOn = False

    If imported Then
        lastImportFile = sourcePath
        outcome = OUTCOME_OK
    Else
        outcome = RunError(0, "the file at " & sourcePath & " was not imported")
    End If

    RunImportData = CloseRun(outcome, ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
    Exit Function

Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    Debug.Print "RunImportData: "; errNumber; errDescription

    'Opens with a suppression because the events have to come back on whatever
    'else fails here. A linelist left with events off answers no worksheet event
    'at all, and the user reads it as the dropdowns having stopped working.
    On Error Resume Next
        If quietStateOn Then LinelistEventsManager.LLExitQuietState
    On Error GoTo 0

    RunImportData = CloseRun(RunError(errNumber, errDescription), _
                             ParentFolderOf(sourcePath), BaseNameOf(sourcePath))
End Function

'@sub-title What the last wrapper run read, wrote and said.
'@details
'key=value lines, then the REPORT_MARKER, then the boxes the run swallowed, one
'per line. The shape HeadlessBuild.LastBuildSummary and SetupLastSummary already
'answer.
'
'outcome= leads the block. The whole reason this text is also written to a file
'is that the answer of Application.Run is lost when the transport gives up, and
'the outcome is the first thing that reading loses, so the file has to carry it.
'
'PATHS ONLY, NO COUNTS. How many variables and how many sheets a linelist holds
'belongs to the generation run, which keeps its own record in the designer's
'GenerationLog and answers it through DesignerLastSummary.
'
'export= is answered by the export wrapper, which is its own piece of work. It
'is empty until that wrapper writes it, the same way geobase= is empty after an
'import run.
'@return String. The summary of the last run, empty keys and all.
'@EntryPoint
Public Function LinelistLastSummary() As String
    LinelistLastSummary = "outcome=" & lastOutcome & vbLf & _
                          "geobase=" & lastGeobaseFile & vbLf & _
                          "import=" & lastImportFile & vbLf & _
                          "export=" & lastExportFile & vbLf & _
                          REPORT_MARKER & vbLf & lastMessages
End Function


'@section What the wrappers share
'===============================================================================

'@sub-title Forget the run before this one.
'@details
'Messenger.Arm empties its own record; this empties what LinelistLastSummary
'reads beside it, so a refused run never answers the paths of the run before it.
Private Sub ResetRunRecord()
    lastOutcome = vbNullString
    lastGeobaseFile = vbNullString
    lastImportFile = vbNullString
    lastExportFile = vbNullString
    lastMessages = vbNullString
End Sub

'@sub-title The one exit of both wrappers.
'@details
'Reads the swallowed boxes, disarms the messenger, writes the summary beside the
'file the run touched, and answers the outcome. The file write is deliberately
'quiet: a run that worked is not turned into a failure because the folder it was
'given cannot be written to.
'@param outcome String. "OK", or an "ERROR ..." line.
'@param folderPath String. Where the summary file goes.
'@param baseName String. What the summary file is named after.
'@return String. The outcome it was given.
Private Function CloseRun(ByVal outcome As String, _
                          ByVal folderPath As String, _
                          ByVal baseName As String) As String
    lastOutcome = outcome
    lastMessages = Messenger.Messages()
    Messenger.Disarm

    On Error Resume Next
        WriteSummaryFile folderPath, baseName
    On Error GoTo 0

    CloseRun = outcome

    'Last of all, and only when the import happened. The line below ends this
    'procedure where it stands, so everything the caller needs has to be on disk
    'before it runs.
    If outcome = OUTCOME_OK Then SealHostWorkbook
End Function

'@sub-title Save the linelist and close it.
'@details
'A script has nobody to press save, and the workbook it drove has to be on disk
'and shut before the caller reads it.
'
'THIS ENDS THE CODE RUNNING INSIDE THE WORKBOOK. Closing the workbook a macro
'is running in stops that macro where it stands, so the wrapper answers
'Application.Run nothing and LinelistLastSummary can no longer be called. The
'summary file is written before this runs for exactly that reason.
'
'It is guarded so a workbook that refuses to close -- one Excel is still busy
'with, one a BeforeClose handler stops -- leaves the run as it was rather than
'turning a finished import into a raise.
Private Sub SealHostWorkbook()
    On Error Resume Next
        ThisWorkbook.Close savechanges:=True
    On Error GoTo 0
End Sub

'@sub-title Write the summary beside the file the run touched.
'@param folderPath String. The folder to write into.
'@param baseName String. The name the file is built on.
Private Sub WriteSummaryFile(ByVal folderPath As String, ByVal baseName As String)
    Dim filePath As String
    Dim fileNumber As Long

    If LenB(folderPath) = 0 Then Exit Sub
    If LenB(baseName) = 0 Then Exit Sub

    filePath = folderPath
    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
    filePath = filePath & baseName & SUMMARY_SUFFIX

    fileNumber = FreeFile

    On Error GoTo CloseFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, LinelistLastSummary()
    Close #fileNumber
    Exit Sub

CloseFile:
    On Error Resume Next
        Close #fileNumber
    On Error GoTo 0
End Sub

'@sub-title Build the "ERROR <number>: <text>" line a failed run answers.
'@param errNumber Long. The error number, 0 for a refusal of the wrapper's own.
'@param message String. What went wrong.
'@return String. The outcome line.
Private Function RunError(ByVal errNumber As Long, ByVal message As String) As String
    RunError = OUTCOME_ERROR_LEAD & CStr(errNumber) & ": " & message
End Function

'@sub-title The pasting rule word this run acts on.
'@details
'An empty word answers append, which is the reading that keeps the rows the
'linelist already holds. A word that is neither answers empty, and the caller is
'told so rather than having its data wiped on a guess.
'@param pastingRule String. What the caller wrote.
'@return String. "replace", "append", or empty when the word is neither.
Private Function ResolvedRule(ByVal pastingRule As String) As String
    Dim word As String

    word = Trim$(pastingRule)
    If LenB(word) = 0 Then
        ResolvedRule = RULE_APPEND
        Exit Function
    End If

    If StrComp(word, RULE_REPLACE, vbTextCompare) = 0 Then ResolvedRule = RULE_REPLACE
    If StrComp(word, RULE_APPEND, vbTextCompare) = 0 Then ResolvedRule = RULE_APPEND
End Function

'@sub-title Whether a file sits at that path.
'@param filePath String. The path to look at.
'@return Boolean. True when the path names a file that is there.
Private Function FileIsThere(ByVal filePath As String) As Boolean
    If LenB(filePath) = 0 Then Exit Function

    On Error Resume Next
        FileIsThere = (LenB(Dir$(filePath)) > 0)
    On Error GoTo 0
End Function

'@sub-title The folder a path sits in.
'@param filePath String. A full file path.
'@return String. Everything before the last separator, empty when there is none.
Private Function ParentFolderOf(ByVal filePath As String) As String
    Dim cutAt As Long

    cutAt = InStrRev(filePath, Application.PathSeparator)
    If cutAt <= 1 Then Exit Function

    ParentFolderOf = Left$(filePath, cutAt - 1)
End Function

'@sub-title The name of a file with its folder and its extension taken off.
'@param filePath String. A file name or a full path.
'@return String. The bare name.
Private Function BaseNameOf(ByVal filePath As String) As String
    Dim bareName As String
    Dim cutAt As Long

    bareName = filePath

    cutAt = InStrRev(bareName, Application.PathSeparator)
    If cutAt > 0 Then bareName = Mid$(bareName, cutAt + 1)

    cutAt = InStrRev(bareName, ".")
    If cutAt > 1 Then bareName = Left$(bareName, cutAt - 1)

    BaseNameOf = bareName
End Function

'@sub-title Take the blank rows off every data entry table before an import.
'@details
'The pre-import walk of ClickImportData, copied rather than called: naming
'EventsLinelistButtons here would drag its seven forms into this module's
'compile closure and the module could then carry no registry row.
'
'A table padded with blank rows takes an appended import under the padding, so
'this runs whatever the pasting rule is. A workbook whose keys could not be
'built hands back Nothing and the walk is skipped rather than raising.
'@param sourceWkb Workbook. The linelist workbook.
'@param keys Passwords. The protection keys, or Nothing.
Private Sub TrimDataTables(ByVal sourceWkb As Workbook, ByVal keys As Passwords)
    Dim sh As Worksheet
    Dim csTab As CustomTable
    Dim Lo As ListObject
    Dim nbBlank As Long

    If keys Is Nothing Then Exit Sub

    For Each sh In sourceWkb.Worksheets
        If SheetTag(sh) = HLIST_TAG Then
            nbBlank = BlankRowCountOf(sh)
            Set Lo = sh.ListObjects(1)
            Set csTab = CustomTable.Create(Lo)
            keys.UnProtect sh.Name
            On Error Resume Next
                If Not (Lo.AutoFilter Is Nothing) Then Lo.AutoFilter.ShowAllData
                csTab.RemoveRows totalCount:=nbBlank
            On Error GoTo 0
            keys.Protect sh.Name
        End If
    Next
End Sub


'@section What the linelist hands back
'===============================================================================
'Each of these reaches the one EventLinelist the events manager holds. The
'answers live in procedure-locals: a module field here would hold a manager an
'import leaves stale.

'@sub-title The messages translator of the running linelist.
'@details
'A workbook with no usable translation sheet answers Nothing, and the wrapper
'says so rather than raising into a transport that would read the raise as no
'answer at all.
'@return TranslationObject. The messages translator, or Nothing.
Private Function MessagesTranslator() As TranslationObject
    Dim linelistEvents As EventLinelist
    Dim lltrads As LLTranslation

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set lltrads = linelistEvents.Translation()
    If lltrads Is Nothing Then Exit Function

    Set MessagesTranslator = lltrads.TransObject()
End Function

'@sub-title The password manager the event service holds.
'@return Passwords. The protection keys, or Nothing.
Private Function PasswordManagerOf() As Passwords
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set PasswordManagerOf = linelistEvents.PasswordManager()
End Function

'@sub-title The user log the event service holds.
'@details
'A workbook whose log cannot be built answers Nothing and every log line of
'this module stays quiet.
'@return LLLog. The log store, or Nothing.
Private Function UserLogOf() As LLLog
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set UserLogOf = linelistEvents.UserLog()
End Function

'@sub-title The hidden name store of one worksheet.
'@details
'The service holds one store per sheet and drops it when that sheet raises a
'change, so the two readers below share one walk.
'@param sh Worksheet. The sheet to read.
'@return HiddenNames. The store, or Nothing.
Private Function SheetStoreOf(ByVal sh As Worksheet) As HiddenNames
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set SheetStoreOf = linelistEvents.SheetNames(sh)
End Function

'@sub-title The sheet type tag of one worksheet.
'@param sh Worksheet. The sheet to read.
'@return String. The tag, empty when the sheet holds no store.
Private Function SheetTag(ByVal sh As Worksheet) As String
    Dim shHn As HiddenNames

    Set shHn = SheetStoreOf(sh)
    If shHn Is Nothing Then Exit Function

    SheetTag = shHn.ValueAsString("sheet_type")
End Function

'@sub-title The number of filled cells an untouched data row of a sheet carries.
'@details
'LLDataEntry writes it when it makes the table, and a row holding more filled
'cells than this is a row the user has typed into.
'@param sh Worksheet. The sheet to read.
'@return Long. The count, 0 when the sheet holds no store.
Private Function BlankRowCountOf(ByVal sh As Worksheet) As Long
    Dim shHn As HiddenNames

    Set shHn = SheetStoreOf(sh)
    If shHn Is Nothing Then Exit Function

    BlankRowCountOf = shHn.ValueAsLong("blank_row_count")
End Function

'@sub-title Drop the held managers of the event service.
'@details
'The walks above rewrite the worksheets those managers were built over: an
'import rewrites the data, the dropdowns and the dictionary metadata, a geobase
'import the Geo sheet. The service builds fresh managers on the next event.
'Called on the error paths too, because a walk that failed midway has already
'rewritten part of what the managers read.
Private Sub ResetEventCaches()
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Sub

    linelistEvents.ResetCaches
End Sub

'@sub-title Recalculate the cells whose formulas read the geobase.
'@details
'The walk itself lives on EventLinelist, where the harness measures it; this
'keeps the caller side to one guarded call.
Private Sub RecalculateGeoCells()
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Sub

    linelistEvents.RecalculateGeoColumns
End Sub

'@sub-title Write the success line of a finished walk.
'@details
'The write is guarded so a log fault never takes down the walk it records.
'@param action String. What the line is about.
'@param detail String. Optional. What to write beside it.
Private Sub LogSuccessLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSuccess action, detail
    On Error GoTo 0
End Sub

'@sub-title Write the failure line of a walk that ended at its error label.
'@param action String. What the line is about.
'@param detail String. Optional. What to write beside it.
Private Sub LogFailureLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogFailure action, detail
    On Error GoTo 0
End Sub

'@sub-title Write the warning line of a refused or cancelled walk.
'@param action String. What the line is about.
'@param detail String. Optional. What to write beside it.
Private Sub LogWarningLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogWarning action, detail
    On Error GoTo 0
End Sub

'@sub-title The file name at the end of a picked path, for the log detail.
'@details
'Both separators are tried, so the helper answers the same on every host.
'@param filePath String. A full file path.
'@return String. Everything after the last separator.
Private Function FileNameOf(ByVal filePath As String) As String
    Dim sepAt As Long

    sepAt = InStrRev(filePath, "/")
    If sepAt = 0 Then sepAt = InStrRev(filePath, "\")

    FileNameOf = Mid$(filePath, sepAt + 1)
End Function
