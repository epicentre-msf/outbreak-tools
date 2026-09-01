Attribute VB_Name = "LinelistRun"
Attribute VB_Description = "The import and export walks of a linelist, and the entry points a script calls them through"

Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("The import and export walks of a linelist, and the entry points a script calls them through")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed
'@depends LLImporter, LLExporter, LLExport, FilteredData, ImportMetadata, ApplicationState, OSFiles, ImportChecking, LinelistEventsManager, EventLinelist, LLTranslation, LLLog, TranslationObject, Messenger, CustomTable, Passwords, HiddenNames

'WHAT IS IN HERE
'The import and export walks of a running linelist -- HandleImportData,
'HandleImportGeobase, HandleExportMigration, HandleExportOther,
'HandleExportAnalysis and HandleExportCustom -- and the four entry points a
'script calls them through: RunImportGeobase, RunImportData, RunExport and
'LinelistLastSummary. The wrappers take their paths as strings, open no picker,
'no prompt and no box, and answer "OK" or "ERROR <number>: <text>".
'LinelistLastSummary reads the last run back.
'
'The buttons of the linelist are unchanged. ClickImportData, ClickImportGeobase,
'ClickExportMigration and ClickExportAnalysis call the same walks with nothing
'handed in, and a person clicking still gets every picker, every prompt and
'every box they always had.
'
'THE WALKS USED TO LIVE IN A FORM
'-------------------------------------------------------------------------------
'The two import walks sat in FormLogicAdvanced and the four export walks in
'FormLogicExportMig, which are the code-behind of F_Advanced and F_ExportMig and
'are merged into those forms at build time. So the only way to reach them was
'through the form -- F_Advanced.HandleImportData -- and that cost two things:
'
' - A script had to open a form to drive an import or an export.
' - No suite could ever reach them. A form module is compiled into the delivered
'   linelist and into nothing the test harness drives, which is why the thirteen
'   message boxes of the two import walks had no test at all.
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
'rather than called. HandleExportCustom is a second copy of that kind: the body
'it repeats is ExportButton.RunExport, and ExportButton is a class holding a
'WithEvents binding to a form control, so it cannot be called from here either.
'
'THE IMPORT WALK IS TIMED
'-------------------------------------------------------------------------------
'HandleImportData names each call of its sequence to the stopwatch LLLog
'holds, and leaves the whole walk on one info line of the workbook log:
'each step with its seconds and a total at the end. The stopwatch opens
'after the picker and the pasting question, so the line times the work and
'not the two questions a person answered, and it is written on every exit,
'the refusal and the error label included.
'The four export walks are timed the same way, inside LLExporter.
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
'A RUN THAT WORKED CLOSES THE LINELIST, AND ONLY AN IMPORT SAVES IT
'-------------------------------------------------------------------------------
'The import is the whole point of an import call, and a script has nobody to
'press save. So CloseRun saves the workbook and closes it once the outcome reads
'OK.
'
'AN EXPORT SAVES NOTHING. It writes its own file and reads this linelist, so the
'workbook is closed with savechanges:=False and whatever the session touched
'stays untouched on disk. That is the saveOnClose argument of CloseRun, and
'RunExport is the one caller that passes it False.
'
'Closing the workbook ends the code running inside it, so that run answers
'Application.Run nothing at all and LinelistLastSummary can no longer be read
'back. THE SUMMARY FILE IS THE READING after any run that worked, which is why
'it is written before the close rather than after.
'
'A run that failed leaves the workbook open. Nothing was written, the caller may
'want to look at it or try again, and its summary file is already on disk.

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

'The words a caller writes for the export to run. An empty word means the
'migration export, which is the file another linelist reads back in.
Private Const EXPORT_MIGRATION As String = "migration"
Private Const EXPORT_GEO As String = "geo"
Private Const EXPORT_HISTORIC As String = "historic"
Private Const EXPORT_ANALYSIS As String = "analysis"

'What ResolvedExport answers for a word that is a custom export number rather
'than one of the four above. It is not a word a caller writes: a caller asking
'for a custom export writes its number.
Private Const EXPORT_CUSTOM As String = "custom"

'The sheet the custom export definitions are listed on. The number a caller
'writes is read off this sheet, and a number that is not active there is
'refused rather than run.
Private Const EXPORTSHEET As String = "Exports"

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

    ' Busy state. The stopwatch opens here rather than at the top of the
    ' walk: the picker and the pasting question wait for a person, and a
    ' walk timed across them would say more about the reader than the work.
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=True, _
                            busyCursor:=xlWait, blockSecurity:=False
    Set actsh = ActiveSheet
    StartStepWatch

    ' Open import workbook. The file is read and closed with savechanges:=False,
    ' so it is opened the way the export side opens one (LLExporter.cls:219):
    ' read-only, and with no link refresh. A workbook whose formulas point at
    ' other files spends the whole refresh before the walk starts, and it can
    ' put a dialog on the screen of a run nobody is watching.
    MarkWalkStep "opening the import file"
    Set impwb = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=0)
    ActiveWindow.WindowState = xlMinimized

    ' Read what the file says about itself, once, and read the file over
    MarkWalkStep "reading what the file says about itself"
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

        ' Import all data. Each call is one step of the stopwatch, so the
        ' log line says which of the seven the walk spent its time in.
        MarkWalkStep "importing the data"
        impObj.ImportData impwb, pasteAtBottom, meta

        MarkWalkStep "importing the custom dropdowns"
        impObj.ImportCustomDropdown impwb, pasteAtBottom

        MarkWalkStep "comparing with the import file"
        impObj.CompareWithImportFile impwb

        MarkWalkStep "finishing the report"
        impObj.FinalizeReport

        ' Import migration metadata. These three read the file's own dictionary
        ' and labels, so they run only when the two files are in the same
        ' language.
        If sameLanguage Then
            MarkWalkStep "importing the show/hide choices"
            impObj.ImportShowHide impwb, meta

            MarkWalkStep "importing the editable labels"
            impObj.ImportEditableLabels impwb, meta

            MarkWalkStep "importing the single values"
            impObj.ImportSingleValues meta
        End If
    End If

    ' Close import workbook
    MarkWalkStep "closing the import file"
    impwb.Close savechanges:=False
    Set impwb = Nothing

    actsh.Activate
    appState.Restore

    ' Everything the import found, written onto the worksheet and left hidden.
    ' ONE write, and it happens HERE: right after the import work, before any
    ' box and before the report form. The Open Log button of that form shows
    ' this worksheet, so a sheet written after the form was dismissed is a
    ' button that does nothing.
    MarkWalkStep "writing what the import found"
    ImportChecking.WriteImportCheckings sourceWkb, impObj.CheckingValues

    ' A REFUSAL IS SHOWN, unlike a finished import. Nothing else tells the user
    ' why the file was turned away: the box carries one sentence, the import
    ' report form is never offered on this path, and the worksheet holds the
    ' reason and what to do about it. This is the only place the sheet is put on
    ' screen by an import.
    If refused Then
        LogWalkSteps "import-data"
        LogWarningLine "import-data", "file refused: " & FileNameOf(filePath)
        Messenger.Show trads.TranslatedValue("MSG_AbortImport"), vbOK, _
                       vbExclamation + vbOKOnly, _
                       trads.TranslatedValue("MSG_Imports")
        ImportChecking.ShowReportSheet sourceWkb
        Exit Function
    End If

    ' The import rewrote the sheets the held managers were built over
    ResetEventCaches
    LogWalkSteps "import-data"
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
    LogWalkSteps "import-data"
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
    LogWalkSteps "import-data"
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

    ' Open geobase workbook, read-only and with no link refresh, the same shape
    ' HandleImportData above uses. A geobase is read and closed unchanged.
    Set impwb = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=0)
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

'@section The four export walks
'===============================================================================
'The bodies that moved here out of the F_ExportMig code-behind, and one that
'repeats ExportButton.RunExport. Each takes the folder it writes into as an
'argument and answers the paths it wrote, one per line, empty when it wrote
'nothing.
'
'THE FORM KEPT ITS PICKER, ITS PROMPT AND ITS Me.Hide. Everything that reads a
'control or moves the form stayed behind in FormLogicExportMig: the folder
'picker, the two other-linelist labels, the password InputBox and the question
'that puts the form away. What moved is the work and the boxes that report it.
'
'EVERY BOX HERE GOES THROUGH THE MESSENGER, so a scripted run swallows it and
'reads it back off LinelistLastSummary. Session 114 did this for the two import
'walks; these four were out of reach then, because a form module is compiled
'into nothing the harness drives.

'@sub-title Export this linelist: the migration file, the geobase, the historic geobase.
'@details
'The body of CurrentLinelistWalk. The three switches say which of the three
'files are written and the two after them say what the migration file carries.
'The busy state covers the whole walk and is restored on both ways out.
'
'CloseAll runs on the way out of a walk that worked as well as one that failed.
'A finished export leaves nothing open -- LLExporter.SaveWorkbook closes the
'output workbook and drops its reference -- so the call costs nothing there and
'catches the workbook an export that stopped midway left behind.
'@param sourceWkb Workbook. The linelist to export.
'@param trads TranslationObject. Translations for messages.
'@param folderPath String. The folder the files land in.
'@param wantData Boolean. True writes the migration file.
'@param wantGeo Boolean. True writes the geobase.
'@param wantHistoric Boolean. True writes the historic geobase.
'@param includeShowHide Boolean. True carries the show/hide state into the migration file.
'@param keepLabels Boolean. True carries the editable labels into the migration file.
'@return String. The paths written, one per line. Empty when nothing was written.
Public Function HandleExportMigration(ByVal sourceWkb As Workbook, _
                                      ByVal trads As TranslationObject, _
                                      ByVal folderPath As String, _
                                      ByVal wantData As Boolean, _
                                      ByVal wantGeo As Boolean, _
                                      ByVal wantHistoric As Boolean, _
                                      ByVal includeShowHide As Boolean, _
                                      ByVal keepLabels As Boolean) As String

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim savedPaths As String
    Dim failDetail As String

    If LenB(folderPath) = 0 Then Exit Function
    If Not (wantData Or wantGeo Or wantHistoric) Then Exit Function

    On Error GoTo ErrHand

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

    exporter.CloseAll
    appState.Restore
    LogSuccessLine "export-migration", PathsOnOneLine(savedPaths)

    Messenger.Show savedPaths, vbOK, vbOKOnly + vbInformation, _
                   trads.TranslatedValue("MSG_FileSaved")

    HandleExportMigration = savedPaths
    Exit Function

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    failDetail = ExporterDetail(exporter, failDetail)
    LogFailureLine "export-migration", failDetail
    Messenger.Show trads.TranslatedValue("MSG_ErrHandExport"), vbOK, _
                   vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Function

'@sub-title Export another linelist, the same three files with the same two switches.
'@details
'The body of OtherLinelistWalk. The file is confirmed first, then opened
'read-only under the busy state so its open events stay quiet, and closed
'without saving once the files are written. CloseAll is what closes it, and it
'runs on both ways out.
'
'THE CONFIRMATION BOX ANSWERS vbYes WHEN THE BOXES ARE OFF, and it is the one
'box of this block that does. It reads back the file the caller chose, and it
'guards a misclick on the form's path label. A caller that wrote the path as an
'argument has already made that choice, so there is no misclick to guard.
'
'The two refusals above it answer vbOK and the walk answers empty: no file at
'that path, and a path naming this very linelist, which is the export the
'migration walk above already is.
'@param trads TranslationObject. Translations for messages.
'@param folderPath String. The folder the files land in.
'@param otherPath String. Full path of the linelist to export.
'@param otherPassword String. The password that file opens with, empty for none.
'@param wantData Boolean. True writes the migration file.
'@param wantGeo Boolean. True writes the geobase.
'@param wantHistoric Boolean. True writes the historic geobase.
'@param includeShowHide Boolean. True carries the show/hide state into the migration file.
'@param keepLabels Boolean. True carries the editable labels into the migration file.
'@return String. The paths written, one per line. Empty when nothing was written.
Public Function HandleExportOther(ByVal trads As TranslationObject, _
                                  ByVal folderPath As String, _
                                  ByVal otherPath As String, _
                                  ByVal otherPassword As String, _
                                  ByVal wantData As Boolean, _
                                  ByVal wantGeo As Boolean, _
                                  ByVal wantHistoric As Boolean, _
                                  ByVal includeShowHide As Boolean, _
                                  ByVal keepLabels As Boolean) As String

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim savedPaths As String
    Dim failDetail As String

    If LenB(folderPath) = 0 Then Exit Function
    If Not (wantData Or wantGeo Or wantHistoric) Then Exit Function

    On Error GoTo ErrHand

    If Not FileIsThere(otherPath) Then
        Messenger.Show trads.TranslatedValue("MSG_ProvideLLPath"), vbOK, _
                       vbExclamation + vbOKOnly, _
                       trads.TranslatedValue("MSG_Migration")
        LogWarningLine "export-other", "no linelist file at " & otherPath
        Exit Function
    End If

    If StrComp(otherPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Messenger.Show trads.TranslatedValue("MSG_ExportMigConflict"), vbOK, _
                       vbExclamation + vbOKOnly, _
                       trads.TranslatedValue("MSG_Migration")
        LogWarningLine "export-other", "that path is this linelist: " & otherPath
        Exit Function
    End If

    If Messenger.Show(trads.TranslatedValue("MSG_ConfirmExportOther") & _
                      vbNewLine & otherPath, _
                      vbYes, vbQuestion + vbYesNo, _
                      trads.TranslatedValue("MSG_Confirm")) = vbNo Then Exit Function

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    'A path or password the file refuses lands here, and the message names that
    'failure rather than the export one.
    On Error GoTo ErrOpen
    Set exporter = LLExporter.CreateFromFile(otherPath, otherPassword)
    On Error GoTo ErrHand

    If wantData Then _
        savedPaths = exporter.ExportMigration(folderPath, includeShowHide, keepLabels)
    If wantGeo Then _
        savedPaths = JoinPath(savedPaths, exporter.ExportGeo(folderPath, onlyHistoric:=False))
    If wantHistoric Then _
        savedPaths = JoinPath(savedPaths, exporter.ExportGeo(folderPath, onlyHistoric:=True))

    'Closes the other linelist, which this exporter opened
    exporter.CloseAll
    appState.Restore
    LogSuccessLine "export-other", PathsOnOneLine(savedPaths)

    Messenger.Show savedPaths, vbOK, vbOKOnly + vbInformation, _
                   trads.TranslatedValue("MSG_FileSaved")

    HandleExportOther = savedPaths
    Exit Function

ErrOpen:
    On Error Resume Next
    LogFailureLine "export-other", "open refused: " & otherPath
    If Not appState Is Nothing Then appState.Restore
    Messenger.Show trads.TranslatedValue("MSG_ErrOpenOther"), vbOK, _
                   vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    Exit Function

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    failDetail = ExporterDetail(exporter, failDetail)
    LogFailureLine "export-other", failDetail
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
    Messenger.Show trads.TranslatedValue("MSG_ErrHandExport"), vbOK, _
                   vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
End Function

'@sub-title Export the analysis worksheets into a workbook of their own.
'@details
'The body of HandleAnalysisExport. No data and no metadata sheet goes with
'them; the file carries the analysis tables and their graphs.
'@param sourceWkb Workbook. The linelist to export.
'@param trads TranslationObject. Translations for messages.
'@param folderPath String. The folder the file lands in.
'@return String. The path written, empty when nothing was written.
Public Function HandleExportAnalysis(ByVal sourceWkb As Workbook, _
                                     ByVal trads As TranslationObject, _
                                     ByVal folderPath As String) As String

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim filePath As String
    Dim failDetail As String

    If LenB(folderPath) = 0 Then Exit Function

    On Error GoTo ErrHand

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    Set exporter = LLExporter.Create(sourceWkb)
    filePath = exporter.ExportAnalysis(folderPath)

    exporter.CloseAll
    appState.Restore
    LogSuccessLine "export-analysis", PathsOnOneLine(filePath)

    Messenger.Show filePath, vbOK, vbOKOnly + vbInformation, _
                   trads.TranslatedValue("MSG_FileSaved")

    HandleExportAnalysis = filePath
    Exit Function

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    failDetail = ExporterDetail(exporter, failDetail)
    LogFailureLine "export-analysis", failDetail
    Messenger.Show trads.TranslatedValue("MSG_ErrHandExport"), vbOK, _
                   vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Function

'@sub-title Run one numbered custom export off the Exports sheet.
'@details
'The body of ExportButton.RunExport, with the folder picker and the filter
'question replaced by arguments. It is a second copy of that body rather than a
'call into it: ExportButton holds a WithEvents binding to an MSForms control, so
'it is built by a form and reached from nowhere else.
'
'The filtered companions are synced first when the export reads them, and a
'sheet the sync skipped stops the export -- the file would otherwise carry a
'companion that no longer matches its table.
'
'THE PASSWORD GOES IN THE BOX, and it has to. A password-protected export is
'encrypted with the linelist private key, which sits in a very hidden
'worksheet, so a caller who is not told it here has no way back into the file.
'A scripted run reads the same text off LinelistLastSummary.
'@param sourceWkb Workbook. The linelist to export.
'@param trads TranslationObject. Translations for messages.
'@param folderPath String. The folder the file lands in.
'@param exportNumber Long. Which export definition to run, 1-based.
'@param useFilter Boolean. True exports the filtered rows of each HList.
'@return String. The path written, empty when nothing was written.
Public Function HandleExportCustom(ByVal sourceWkb As Workbook, _
                                   ByVal trads As TranslationObject, _
                                   ByVal folderPath As String, _
                                   ByVal exportNumber As Long, _
                                   ByVal useFilter As Boolean) As String

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim filePath As String
    Dim failDetail As String

    If LenB(folderPath) = 0 Then Exit Function
    If exportNumber < 1 Then Exit Function

    On Error GoTo ErrHand

    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    If useFilter Then SyncFilteredCompanions sourceWkb, trads

    Set exporter = LLExporter.Create(sourceWkb)
    filePath = exporter.ExportCustom(exportNumber, folderPath, useFilter)

    exporter.CloseAll
    appState.Restore
    LogSuccessLine "export-custom", PathsOnOneLine(filePath)

    ReportSavedExport filePath, exporter.LastExportPassword, trads

    HandleExportCustom = filePath
    Exit Function

ErrHand:
    ' Err is read before the Resume Next below clears it.
    failDetail = Err.Description
    On Error Resume Next
    failDetail = ExporterDetail(exporter, failDetail)
    LogFailureLine "export-custom", failDetail
    Messenger.Show trads.TranslatedValue("MSG_ErrHandExport") & ": " & failDetail, _
                   vbOK, vbOKOnly + vbCritical, trads.TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
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

'@sub-title Run one export of this linelist, with no picker, no prompt and no box.
'@details
'The four export walks above, reached by a word instead of by five checkboxes
'on a form. The folder is handed in rather than picked, and so is the password
'the other-linelist export used to ask for through an InputBox.
'
'WHAT THE WORD SAYS
'
'| exportName | What runs |
'| --- | --- |
'| empty, or "migration" | the migration file another linelist reads back in |
'| "geo" | the geobase |
'| "historic" | the historic geobase |
'| "analysis" | the analysis worksheets, in a workbook of their own |
'| a number, "1", "2", ... | that custom export off the Exports sheet |
'
'Any other word is refused and named back to the caller. A number that is not
'an active export on the Exports sheet is refused the same way, rather than run
'as export number one.
'
'ANOTHER LINELIST IS EXPORTED BY NAMING IT. otherLinelist is the file to export
'from and otherPassword is what it opens with; both empty means this linelist.
'The file is opened read-only and closed without saving, which is what
'LLExporter.CloseAll does. Only the first three words take it -- the analysis
'export and the custom exports read this linelist and nothing else -- and a
'call that pairs the two is refused rather than quietly exporting the wrong
'workbook.
'
'A run that exported closes this linelist WITHOUT SAVING, so this answers "OK"
'to nobody. Read the summary file. An export writes its own file and only reads
'this one, so there is nothing here to save. A run that failed answers its
'refusal and leaves the workbook open.
'@param exportName String. The export to run. Empty means the migration export.
'@param outputFolder String. The folder the files land in.
'@param otherPassword String. Optional. The password the other linelist opens with.
'@param otherLinelist String. Optional. Full path of the linelist to export from.
'@return String. "OK", or "ERROR <number>: <text>". A run that worked returns no
'answer at all, because the workbook it was called in is closed by then.
'@EntryPoint
Public Function RunExport(ByVal exportName As String, _
                          ByVal outputFolder As String, _
                          Optional ByVal otherPassword As String = vbNullString, _
                          Optional ByVal otherLinelist As String = vbNullString) As String
    Dim trads As TranslationObject
    Dim folderPath As String
    Dim otherPath As String
    Dim runWord As String
    Dim outcome As String
    Dim savedPaths As String
    Dim exportNumber As Long
    Dim errNumber As Long
    Dim errDescription As String

    Messenger.Arm ThisWorkbook
    ResetRunRecord

    folderPath = Trim$(outputFolder)
    otherPath = Trim$(otherLinelist)

    If Not FolderIsThere(folderPath) Then
        RunExport = CloseExportRun(RunError(0, "no folder at " & folderPath))
        Exit Function
    End If

    runWord = ResolvedExport(exportName)
    If LenB(runWord) = 0 Then
        RunExport = CloseExportRun(RunError(0, "export """ & exportName & _
                                              """ is no export this linelist runs"), _
                                   folderPath)
        Exit Function
    End If

    If LenB(otherPath) > 0 And Not WordTakesAnotherLinelist(runWord) Then
        RunExport = CloseExportRun(RunError(0, "the " & runWord & _
                                              " export reads this linelist only, so it takes no other linelist"), _
                                   folderPath)
        Exit Function
    End If

    On Error GoTo Handler

    'Read before the translator, because it is the last of the guards on the
    'arguments and a bad argument is answered before any of the linelist is.
    If runWord = EXPORT_CUSTOM Then
        exportNumber = CustomExportNumber(exportName)
        If Not CustomExportIsActive(ThisWorkbook, exportNumber) Then
            RunExport = CloseExportRun(RunError(0, "export number " & CStr(exportNumber) & _
                                                  " is not an active export on the " & _
                                                  EXPORTSHEET & " sheet"), _
                                       folderPath)
            Exit Function
        End If
    End If

    Set trads = MessagesTranslator()
    If trads Is Nothing Then
        RunExport = CloseExportRun(RunError(0, "this linelist carries no usable translation sheet"), _
                                   folderPath)
        Exit Function
    End If

    savedPaths = ExportForWord(runWord, trads, folderPath, otherPath, otherPassword, exportNumber)

    If LenB(savedPaths) > 0 Then
        lastExportFile = PathsOnOneLine(savedPaths)
        outcome = OUTCOME_OK
    Else
        outcome = RunError(0, "the " & runWord & " export wrote no file")
    End If

    RunExport = CloseExportRun(outcome, folderPath)
    Exit Function

Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    Debug.Print "RunExport: "; errNumber; errDescription

    RunExport = CloseExportRun(RunError(errNumber, errDescription), folderPath)
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
'export= carries every path RunExport wrote, comma separated, because one word
'can write more than one file. It is empty after an import run, the same way
'import= is empty after an export one.
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
'@param saveOnClose Boolean. Optional, default True. False closes the workbook
'without saving, which is what an export does: it wrote its own file and only
'read this one.
'@return String. The outcome it was given.
Private Function CloseRun(ByVal outcome As String, _
                          ByVal folderPath As String, _
                          ByVal baseName As String, _
                          Optional ByVal saveOnClose As Boolean = True) As String
    lastOutcome = outcome
    lastMessages = Messenger.Messages()
    Messenger.Disarm

    On Error Resume Next
        WriteSummaryFile folderPath, baseName
    On Error GoTo 0

    CloseRun = outcome

    'Last of all, and only when the run worked. The line below ends this
    'procedure where it stands, so everything the caller needs has to be on disk
    'before it runs.
    If outcome = OUTCOME_OK Then SealHostWorkbook saveOnClose
End Function

'@sub-title Close the linelist, saving it when the run wrote into it.
'@details
'A script has nobody to press save, and the workbook it drove has to be shut
'before the caller reads it.
'
'AN IMPORT SAVES AND AN EXPORT DOES NOT. An import rewrote the data, the
'dropdowns and the dictionary metadata, and that work is the whole point of the
'call. An export wrote a file of its own and only read this workbook, so it
'closes with nothing saved and leaves the file on disk as it was.
'
'THIS ENDS THE CODE RUNNING INSIDE THE WORKBOOK. Closing the workbook a macro
'is running in stops that macro where it stands, so the wrapper answers
'Application.Run nothing and LinelistLastSummary can no longer be called. The
'summary file is written before this runs for exactly that reason.
'
'It is guarded so a workbook that refuses to close -- one Excel is still busy
'with, one a BeforeClose handler stops -- leaves the run as it was rather than
'turning a finished run into a raise.
'@param keepChanges Boolean. Optional, default True. True saves on the way out.
Private Sub SealHostWorkbook(Optional ByVal keepChanges As Boolean = True)
    On Error Resume Next
        ThisWorkbook.Close savechanges:=keepChanges
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

'@sub-title Open the stopwatch of a walk on the workbook log.
'@details
'The log holds the stopwatch, so nothing is kept in this module, and a
'workbook that will not take a log simply goes untimed. One log instance
'times one walk: this module reads the one the event service holds, while
'LLExporter times its four export walks on a log of its own.
Private Sub StartStepWatch()
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.StartWalk
    On Error GoTo 0
End Sub

'@sub-title Name the step a walk is about to take, and time the one before it.
'@param stepName String. What the walk is about to do.
Private Sub MarkWalkStep(ByVal stepName As String)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.MarkStep stepName
    On Error GoTo 0
End Sub

'@sub-title Write the step times of a walk, and close its stopwatch.
'@details
'Called on every exit of a timed walk, the refusals and the error label
'included: the line is as much wanted when a walk stops early as when it
'finishes, and it then carries the step it stopped in with the seconds it
'had spent. A walk with no stopwatch open writes nothing.
'@param action String. The action code the walk logs under.
Private Sub LogWalkSteps(ByVal action As String)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSteps action, "LinelistRun.HandleImportData"
    On Error GoTo 0
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


'@section What the export walks and RunExport share
'===============================================================================

'@sub-title The one exit of RunExport.
'@details
'CloseRun with the two answers an export always gives: the summary is named
'after this linelist, because an export writes several files and no one of them
'is the file the run touched; and the workbook is closed WITHOUT SAVING,
'because an export only ever read it.
'
'A refused run has no folder to write the summary into, so the folder of this
'workbook stands in.
'@param outcome String. "OK", or an "ERROR ..." line.
'@param folderPath String. Optional. Where the summary file goes.
'@return String. The outcome it was given.
Private Function CloseExportRun(ByVal outcome As String, _
                                Optional ByVal folderPath As String = vbNullString) As String
    Dim summaryFolder As String

    summaryFolder = folderPath
    If LenB(summaryFolder) = 0 Then summaryFolder = ThisWorkbook.Path

    CloseExportRun = CloseRun(outcome, summaryFolder, BaseNameOf(ThisWorkbook.Name), _
                              saveOnClose:=False)
End Function

'@sub-title Run the walk the word names.
'@details
'The one place the five checkboxes of F_ExportMig are answered for a scripted
'run. A word naming one of the three migration files ticks that box and leaves
'the other two alone; the migration file itself carries the show/hide state and
'the editable labels, which is what ClickExportMigration ticks for a person.
'@param runWord String. The word ResolvedExport answered.
'@param trads TranslationObject. Translations for messages.
'@param folderPath String. The folder the files land in.
'@param otherPath String. The other linelist, empty for this one.
'@param otherPassword String. The password that file opens with.
'@param exportNumber Long. The custom export to run, 0 for the other words.
'@return String. The paths written, one per line.
Private Function ExportForWord(ByVal runWord As String, _
                               ByVal trads As TranslationObject, _
                               ByVal folderPath As String, _
                               ByVal otherPath As String, _
                               ByVal otherPassword As String, _
                               ByVal exportNumber As Long) As String

    Dim wantData As Boolean
    Dim wantGeo As Boolean
    Dim wantHistoric As Boolean

    Select Case runWord
    Case EXPORT_ANALYSIS
        ExportForWord = HandleExportAnalysis(ThisWorkbook, trads, folderPath)
        Exit Function
    Case EXPORT_CUSTOM
        ExportForWord = HandleExportCustom(ThisWorkbook, trads, folderPath, _
                                           exportNumber, useFilter:=False)
        Exit Function
    Case EXPORT_MIGRATION
        wantData = True
    Case EXPORT_GEO
        wantGeo = True
    Case EXPORT_HISTORIC
        wantHistoric = True
    End Select

    If LenB(otherPath) > 0 Then
        ExportForWord = HandleExportOther(trads, folderPath, otherPath, otherPassword, _
                                          wantData, wantGeo, wantHistoric, _
                                          includeShowHide:=True, keepLabels:=True)
        Exit Function
    End If

    ExportForWord = HandleExportMigration(ThisWorkbook, trads, folderPath, _
                                          wantData, wantGeo, wantHistoric, _
                                          includeShowHide:=True, keepLabels:=True)
End Function

'@sub-title The export word this run acts on.
'@details
'An empty word answers the migration export, which is the file the whole
'migration path exists for. A word that is none of the four and no export
'number answers empty, and the caller is told the word back rather than having
'some other export run on a guess.
'@param exportName String. What the caller wrote.
'@return String. One of the four words, EXPORT_CUSTOM, or empty for none of them.
Private Function ResolvedExport(ByVal exportName As String) As String
    Dim word As String

    word = Trim$(exportName)
    If LenB(word) = 0 Then
        ResolvedExport = EXPORT_MIGRATION
        Exit Function
    End If

    If StrComp(word, EXPORT_MIGRATION, vbTextCompare) = 0 Then ResolvedExport = EXPORT_MIGRATION
    If StrComp(word, EXPORT_GEO, vbTextCompare) = 0 Then ResolvedExport = EXPORT_GEO
    If StrComp(word, EXPORT_HISTORIC, vbTextCompare) = 0 Then ResolvedExport = EXPORT_HISTORIC
    If StrComp(word, EXPORT_ANALYSIS, vbTextCompare) = 0 Then ResolvedExport = EXPORT_ANALYSIS
    If CustomExportNumber(word) > 0 Then ResolvedExport = EXPORT_CUSTOM
End Function

'@sub-title Whether the word names an export that can read another linelist.
'@details
'The three migration files can. The analysis export and the custom exports
'read the running linelist and nothing else.
'@param runWord String. The word ResolvedExport answered.
'@return Boolean. True when otherLinelist means something for that word.
Private Function WordTakesAnotherLinelist(ByVal runWord As String) As Boolean
    WordTakesAnotherLinelist = (runWord = EXPORT_MIGRATION Or _
                                runWord = EXPORT_GEO Or _
                                runWord = EXPORT_HISTORIC)
End Function

'@sub-title The custom export number a word carries.
'@details
'Every character is checked against the digits rather than handed to IsNumeric
'or Val. This box reads as en_FR, where the decimal separator is a COMMA, so
'"3,5" passes IsNumeric here and Val answers 3 from it. A caller that wrote
'anything but digits meant no export number, and the word is refused above.
'@param exportName String. What the caller wrote.
'@return Long. The number, 0 when the word is not one.
Private Function CustomExportNumber(ByVal exportName As String) As Long
    Dim word As String
    Dim oneChar As String
    Dim counter As Long

    word = Trim$(exportName)
    If LenB(word) = 0 Then Exit Function
    If Len(word) > 9 Then Exit Function

    For counter = 1 To Len(word)
        oneChar = Mid$(word, counter, 1)
        If oneChar < "0" Or oneChar > "9" Then Exit Function
    Next counter

    CustomExportNumber = CLng(word)
End Function

'@sub-title Whether a number names an active export on the Exports sheet.
'@details
'A workbook with no Exports sheet, and a number past the end of the table,
'both answer False. RunExport then names the number back to the caller rather
'than running export number one in its place.
'@param sourceWkb Workbook. The linelist to read.
'@param exportNumber Long. The number to look for.
'@return Boolean. True when that export is there and marked active.
Private Function CustomExportIsActive(ByVal sourceWkb As Workbook, _
                                      ByVal exportNumber As Long) As Boolean
    Dim expsh As Worksheet
    Dim expObj As LLExport

    If exportNumber < 1 Then Exit Function

    On Error Resume Next
        Set expsh = sourceWkb.Worksheets(EXPORTSHEET)
    On Error GoTo 0
    If expsh Is Nothing Then Exit Function

    On Error Resume Next
        Set expObj = LLExport.Create(expsh)
        If Not expObj Is Nothing Then CustomExportIsActive = expObj.IsActive(exportNumber)
        Err.Clear
    On Error GoTo 0
End Function

'@sub-title Sync the filtered companions a filtered export reads.
'@details
'The body of ExportButton.SyncFilteredData. A sheet the sync skipped means the
'export would read a companion that no longer matches its table, so the skipped
'sheets are joined into one message and raised into the walk's handler.
'@param sourceWkb Workbook. The linelist to sync.
'@param trads TranslationObject. Translations for messages.
'@throws ProjectError.SomethingWentWrong When the sync skipped a sheet.
Private Sub SyncFilteredCompanions(ByVal sourceWkb As Workbook, _
                                   ByVal trads As TranslationObject)
    Dim filtered As FilteredData
    Dim failures As BetterArray
    Dim failureText As String
    Dim counter As Long

    Set filtered = FilteredData.Create(sourceWkb)
    filtered.Sync

    Set failures = filtered.FailedSheets
    If failures.Length = 0 Then Exit Sub

    For counter = failures.LowerBound To failures.UpperBound
        failureText = failureText & vbNewLine & failures.Item(counter)
    Next counter

    Err.Raise ProjectError.SomethingWentWrong, "LinelistRun", _
              trads.TranslatedValue("MSG_ErrUpdate") & failureText
End Sub

'@sub-title Tell the caller where a custom export went and what opens it.
'@details
'A password-protected export is encrypted with the linelist private key, which
'sits in a very hidden worksheet. A caller who is not told the password here has
'no way back into the file, and neither has the colleague they send it to.
'
'Nothing is shown when the export answered no path, which is what a mode that
'stopped early returns.
'@param filePath String. The saved file path.
'@param password String. The password applied to the file, empty for none.
'@param trads TranslationObject. Translations for messages.
Private Sub ReportSavedExport(ByVal filePath As String, _
                              ByVal password As String, _
                              ByVal trads As TranslationObject)
    Dim message As String

    If LenB(filePath) = 0 Then Exit Sub

    message = filePath & vbNewLine & vbNewLine

    If LenB(password) > 0 Then
        message = message & trads.TranslatedValue("MSG_Password") & " " & password
    Else
        message = message & trads.TranslatedValue("MSG_NoPassword")
    End If

    Messenger.Show message, vbOK, vbOKOnly + vbInformation, _
                   trads.TranslatedValue("MSG_FileSaved")
End Sub

'@sub-title Stack a new file path under the ones already collected.
'@param collected String. The paths so far, possibly empty.
'@param newPath String. The path to add.
'@return String. The paths, one per line.
Private Function JoinPath(ByVal collected As String, _
                          ByVal newPath As String) As String
    If LenB(collected) = 0 Then
        JoinPath = newPath
    Else
        JoinPath = collected & vbNewLine & newPath
    End If
End Function

'@sub-title The saved paths on one line, for the log detail and the summary.
'@param savedPaths String. The paths, one per line.
'@return String. The same paths, comma separated.
Private Function PathsOnOneLine(ByVal savedPaths As String) As String
    PathsOnOneLine = Replace(savedPaths, vbNewLine, ", ")
End Function

'@sub-title The detail line of a failed export walk.
'@details
'An error raised inside the exporter reaches the handler as the name of the
'method and nothing else, so the exporter's own account of the failure is
'preferred whenever it has one.
'@param exporter LLExporter. The exporter of the walk, Nothing when it never got made.
'@param errDetail String. The description the walk's handler read.
'@return String. What to write into the log.
Private Function ExporterDetail(ByVal exporter As LLExporter, _
                                ByVal errDetail As String) As String
    ExporterDetail = errDetail
    If exporter Is Nothing Then Exit Function
    If LenB(exporter.LastFailure) > 0 Then ExporterDetail = exporter.LastFailure
End Function

'@sub-title Whether a folder sits at that path.
'@details
'GetAttr rather than Dir. Dir on a folder answers the first file inside it on
'some hosts and the folder name on others, and it also resets the walk any
'other Dir loop is in the middle of.
'@param folderPath String. The path to look at.
'@return Boolean. True when the path names a folder that is there.
Private Function FolderIsThere(ByVal folderPath As String) As Boolean
    Dim folderAttr As Long

    If LenB(folderPath) = 0 Then Exit Function

    On Error Resume Next
        folderAttr = GetAttr(folderPath)
        If Err.Number = 0 Then FolderIsThere = ((folderAttr And vbDirectory) = vbDirectory)
        Err.Clear
    On Error GoTo 0
End Function
