Attribute VB_Name = "BuildSteps"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("The linelist build as steps: each one answers an outcome string and keeps its state for the next.")
'@depends LinelistSpecs, Linelist, LLDataEntry, LLSheets, AnalysisOutput, DropdownLists, DesignerEntry, GenerationLog, Checking, BetterArray, TemporaryRepos, InitTransfer
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@description
'ONE BUILD, NINE STEPS
'-------------------------------------------------------------------------------
'The body of a linelist generation, cut at its milestones: BuildBegin
'(the specifications and the transfer), BuildLinelist, BuildSheetCount,
'one BuildSheet per data entry sheet, BuildDropdowns, BuildAnalyses,
'BuildSave. The specifications, the linelist and the sheet list live in
'module-level fields between the calls, so a driver runs the steps one
'after the other and keeps the loop on its own side.
'
'Every step answers a string: "OK", or "ERROR <number> (<source>):
'<description>", the shape HeadlessBuild answers. BuildSheetCount answers
'the count, BuildCheckings the bundles, BuildCounts the two totals. A step
'raises nothing: a driver in another Excel instance reaches this module
'through Application.Run, where a raise arrives as a bare 1004, so the
'outcome string is the one faithful error channel.
'
'A STEP THAT FAILS KEEPS THE FILE
'-------------------------------------------------------------------------------
'Every step wraps its whole body in a handler. A step that fails aborts
'the build itself before it answers: the unfinished output workbook is
'saved as __temp.xlsb in the temporary repository, closed, and the path
'rides the outcome as "| kept <path>". BuildAbort is the same exit, for a
'driver that stops a build on its own. Nothing here opens a dialog.
'
'THE CHECKINGS TRAVEL AFTER EVERY STEP
'-------------------------------------------------------------------------------
'The Checking bundles a step files are queued as text, through
'GenerationLog.BundleText, and BuildCheckings answers the queue and empties
'it. The driver pulls it after every step, a failed step included, and
'hands it to GenerationLog.CollectText on its own side. That is what keeps
'the promise of the report: a build that dies still names the phases that
'finished.
'
'THE DESIGNER OF THE BUILD
'-------------------------------------------------------------------------------
'BuildBegin takes the designer workbook the entries sit on. A driver in the
'same instance hands it over; a driver in another instance cannot pass an
'object, so it calls BuildBeginEntries with the entries as text, one
'"tag=value" per line, and the steps read ThisWorkbook, which is the
'designer copy they run inside. The copy's Main takes the entries first.
'
'EVERY PUBLIC FUNCTION TRAPS ITS OWN RAISE
'-------------------------------------------------------------------------------
'Application.Run opens a fresh call stack, so a raise that escapes a
'function reached through it reaches no handler of the caller: in place it
'shows the runtime error box, across the processes the hidden instance
'shows the VBE box and un-hides itself. BuildCheckings, BuildCounts and
'BuildAbort answer plain values and carry a handler too, for that reason.


'@section Constants
'===============================================================================

Private Const MODULE_NAME As String = "BuildSteps"
Private Const SHEET_MAIN As String = "Main"
Private Const OUTCOME_OK As String = "OK"
Private Const OUTCOME_ERROR_LEAD As String = "ERROR "

'The mark that leads the kept path at the end of an outcome
Private Const KEPT_MARK As String = " | kept "

'The name of the unfinished workbook a stopped build leaves behind, the one
'Linelist.DiscardBuild writes. BuildBegin looks for the file of an earlier
'run under this name to say it is overwritten.
Private Const KEPT_FILE_NAME As String = "__temp.xlsb"

'The separators of the entries text BuildBegin takes: one entry per
'vertical tab, the tag before the first equals sign.
Private Const ENTRY_SEP As String = vbVerticalTab
Private Const ENTRY_EQUALS As String = "="

'The separator of the two totals BuildCounts answers
Private Const COUNTS_SEP As String = "|"

'The title of the bundle this module files on its own
Private Const TITLE_EARLIER_FILE As String = "build start"


'@section State
'===============================================================================

'The build under way. BuildBegin resets everything; the failure exit and
'BuildAbort drop the objects and keep the counts, the queue and the kept
'path for the driver to read.
Private designerBook As Workbook
Private specs As LinelistSpecs
Private ll As Linelist
Private sheetLists As BetterArray
Private sheetInfo As LLSheets

'The bundle texts queued since the last BuildCheckings
Private pendingText As String

'What the build has written so far
Private builtSheets As Long
Private builtVariables As Long

'The path of the kept file, set by the abort
Private keptPath As String


'@section The steps
'===============================================================================

'@Description("Prepare the specifications and fill the output workbook from the setup and the designer.")
'@details
'The first step. Resets the state of the last build, writes the entries
'when some are handed over, notes the __temp.xlsb of an earlier failed run
'when there is one, then runs LinelistSpecs.Prepare, which creates the
'output workbook and fills it through InitTransfer. The bundles the
'specification collaborators and the transfer filed are queued.
'@param targetDesigner Optional Workbook. The designer the entries sit on. Nothing reads ThisWorkbook.
'@param entriesText Optional String. Entries to write on Main first, one "tag=value" per line.
'@return String. "OK" or "ERROR <number> (<source>): <description>".
Public Function BuildBegin(Optional ByVal targetDesigner As Workbook = Nothing, _
                           Optional ByVal entriesText As String = vbNullString) As String
    Dim entry As DesignerEntry

    On Error GoTo Failed
    ResetState

    Set designerBook = targetDesigner
    If designerBook Is Nothing Then Set designerBook = ThisWorkbook

    Set entry = DesignerEntry.Create(designerBook.Worksheets(SHEET_MAIN))
    WriteEntries entry, entriesText

    NoteEarlierKeptFile entry.ValueOf("lldir")

    'Prepare creates the output workbook and hands it to InitTransfer,
    'which fills it from the setup file and from the designer.
    Set specs = LinelistSpecs.Create(designerBook)
    specs.Prepare entry.ValueOf("setuppath")

    'The specification checkings (dictionary, choices, exports and the
    'rest), then the transfer record. A setup whose translations table is
    'missing is filed there.
    QueueBundles GenerationLog.SpecificationBundles(specs)
    If InitTransfer.HasCheckings() Then QueueBundle InitTransfer.CheckingValues()

    BuildBegin = OUTCOME_OK
    Exit Function

Failed:
    BuildBegin = FailureOutcome(Err.Number, Err.Source, Err.Description, "BuildBegin")
End Function

'@Description("BuildBegin for a driver in another instance: the entries arrive as text, the designer is the copy.")
'@details
'A workbook cannot cross Application.Run between two Excel processes, and
'the copy's Main holds the entries of the moment the copy was written. A
'driver that builds several rows over one copy hands each row's entries
'over here, and the step writes them on the copy's Main before it reads
'them. The text is the shape WriteEntries reads: one "tag=value" per
'vertical tab.
'@param entriesText String. The entries to write on Main first. Empty writes nothing.
'@return String. "OK" or "ERROR <number> (<source>): <description>".
Public Function BuildBeginEntries(ByVal entriesText As String) As String
    BuildBeginEntries = BuildBegin(Nothing, entriesText)
End Function

'@Description("Build the output workbook: sheets, temporary sheets, admin sheet, code transfer.")
'@details
'Linelist.Prepare over the prepared specifications. The sheet list and
'the shared sheet information manager are kept for the sheet steps. The
'checkings of the code transfer are queued: a component the output
'workbook already carried was replaced by the designer's copy, and this
'is where the report names it.
'@return String. "OK" or "ERROR <number> (<source>): <description>".
Public Function BuildLinelist() As String
    On Error GoTo Failed
    RequireSpecifications "BuildLinelist"

    Set ll = Linelist.Create(specs)
    ll.Prepare

    If ll.HasCheckings Then QueueBundle ll.CheckingValues

    'The sheet names Linelist.Prepare already walked the dictionary for, and
    'the LLSheets it holds, so every row resolution is computed once.
    Set sheetLists = ll.SheetNames
    Set sheetInfo = ll.SheetInfoManager

    BuildLinelist = OUTCOME_OK
    Exit Function

Failed:
    BuildLinelist = FailureOutcome(Err.Number, Err.Source, Err.Description, "BuildLinelist")
End Function

'@Description("The number of data entry sheets the build will make.")
'@details
'Answered after BuildLinelist, so a driver sizes its loop and its bar.
'@return String. The count as text, or "ERROR <number> (<source>): <description>".
Public Function BuildSheetCount() As String
    On Error GoTo Failed
    RequireLinelist "BuildSheetCount"

    BuildSheetCount = CStr(sheetLists.Length)
    Exit Function

Failed:
    BuildSheetCount = FailureOutcome(Err.Number, Err.Source, Err.Description, "BuildSheetCount")
End Function

'@Description("Build one data entry sheet by its position in the sheet list.")
'@details
'A sheet of neither layout is skipped and the step still answers OK: the
'dictionary decides what a sheet is and this step only builds. The sheet's
'build checkings are queued for the worksheet and its per-variable
'milestones for the record alone, since those are a few hundred entries.
'@param position Long. The 1-based position among the sheets BuildSheetCount counted.
'@return String. "OK" or "ERROR <number> (<source>): <description>".
Public Function BuildSheet(ByVal position As Long) As String
    Dim sheetName As String
    Dim sheetType As String
    Dim layer As Byte
    Dim listBld As LLDataEntry

    On Error GoTo Failed
    RequireLinelist "BuildSheet"

    If position < 1 Or position > sheetLists.Length Then
        ThrowError ProjectError.InvalidArgument, _
                   "There is no data entry sheet at position " & CStr(position)
    End If

    sheetName = CStr(sheetLists.Item(sheetLists.LowerBound + position - 1))
    sheetType = sheetInfo.SheetInfo(sheetName)

    If sheetType = "vlist1D" Then
        layer = LLDataEntryLayerVList
    ElseIf sheetType = "hlist2D" Then
        layer = LLDataEntryLayerHList
    Else
        BuildSheet = OUTCOME_OK
        Exit Function
    End If

    Set listBld = LLDataEntry.Create(layer, sheetName, ll, sheetInfo)
    listBld.Build

    If listBld.HasCheckings Then QueueBundle listBld.CheckingValues
    If listBld.HasMilestones Then QueueBundle listBld.MilestoneValues, True

    builtSheets = builtSheets + 1
    builtVariables = builtVariables + listBld.VariablesWritten

    BuildSheet = OUTCOME_OK
    Exit Function

Failed:
    BuildSheet = FailureOutcome(Err.Number, Err.Source, Err.Description, "BuildSheet")
End Function

'@Description("Flush the two dropdown stores of the output workbook.")
'@details
'The standard store first, the custom store second, each one queuing its
'own checkings.
'@return String. "OK" or "ERROR <number> (<source>): <description>".
Public Function BuildDropdowns() As String
    Dim store As DropdownLists

    On Error GoTo Failed
    RequireLinelist "BuildDropdowns"

    Set store = ll.Dropdown(1)
    If store.HasCheckings Then QueueBundle store.CheckingValues

    Set store = ll.Dropdown(2)
    If store.HasCheckings Then QueueBundle store.CheckingValues

    BuildDropdowns = OUTCOME_OK
    Exit Function

Failed:
    BuildDropdowns = FailureOutcome(Err.Number, Err.Source, Err.Description, "BuildDropdowns")
End Function

'@Description("Write the four analysis sheets.")
'@details
'The analyses' own entries are queued even when the stage raises.
'AnalysisOutput logs the scope it reached and the table that refused, and
'a failure that took those entries with it left the report saying only
'"Failed: <description>", which names nothing.
'@return String. "OK" or "ERROR <number> (<source>): <description>".
Public Function BuildAnalyses() As String
    Dim anaOut As AnalysisOutput
    Dim errNumber As Long
    Dim errSource As String
    Dim errDesc As String

    On Error GoTo Failed
    RequireLinelist "BuildAnalyses"

    Set anaOut = AnalysisOutput.Create(specs.AnalysisObject.Wksh(), ll)
    anaOut.WriteAnalysis AnalysisBuildStageAll

    QueueAnalysisBundle anaOut

    BuildAnalyses = OUTCOME_OK
    Exit Function

Failed:
    'The fault is read before anything else runs: the queue below has its
    'own On Error, and an On Error statement clears Err.
    errNumber = Err.Number
    errSource = Err.Source
    errDesc = Err.Description

    QueueAnalysisBundle anaOut
    BuildAnalyses = FailureOutcome(errNumber, errSource, errDesc, "BuildAnalyses")
End Function

'@Description("Save the linelist as .xlsb with its password and close it.")
'@details
'Linelist.SaveLL, which also empties the temporary repository. The state
'objects are dropped here: the build is over, and a driver that wants the
'totals reads BuildCounts, which survives.
'@return String. "OK" or "ERROR <number> (<source>): <description>".
Public Function BuildSave() As String
    On Error GoTo Failed
    RequireLinelist "BuildSave"

    ll.SaveLL

    DropBuildObjects
    BuildSave = OUTCOME_OK
    Exit Function

Failed:
    BuildSave = FailureOutcome(Err.Number, Err.Source, Err.Description, "BuildSave")
End Function

'@Description("Answer the bundles queued since the last call, and empty the queue.")
'@details
'The joined texts of GenerationLog.BundleText, ready for
'GenerationLog.CollectText. Empty when nothing was queued.
'@return String. The queued bundle texts, or "ERROR <number> (<source>): <description>".
Public Function BuildCheckings() As String
    On Error GoTo Failed
    BuildCheckings = pendingText
    pendingText = vbNullString
    Exit Function

Failed:
    BuildCheckings = ErrorText(Err.Number, Err.Source, Err.Description, "BuildCheckings")
End Function

'@Description("Answer the sheets built and the variables written by the current build.")
'@details
'Two numbers on one line, "<sheets>|<variables>". They count from
'BuildBegin and survive the failure exit and BuildSave, so a driver adds
'them up once per build, whatever way the build ended.
'@return String. "<sheets>|<variables>", or "ERROR <number> (<source>): <description>".
Public Function BuildCounts() As String
    On Error GoTo Failed
    BuildCounts = CStr(builtSheets) & COUNTS_SEP & CStr(builtVariables)
    Exit Function

Failed:
    BuildCounts = ErrorText(Err.Number, Err.Source, Err.Description, "BuildCounts")
End Function

'@Description("Stop the build: keep the unfinished workbook as __temp.xlsb, close it, drop the state.")
'@details
'The exit every failed step takes on its own, offered to a driver that
'stops a build for a reason of its own. With no output workbook open, or
'after a step already aborted, there is nothing to keep and the answer is
'a bare OK.
'@return String. "OK", or "OK | kept <path>" when a file was kept, or "ERROR <number> (<source>): <description>".
Public Function BuildAbort() As String
    Dim kept As String

    On Error GoTo Failed
    kept = AbortBuild()

    BuildAbort = OUTCOME_OK
    If LenB(kept) > 0 Then BuildAbort = BuildAbort & KEPT_MARK & kept
    Exit Function

Failed:
    BuildAbort = ErrorText(Err.Number, Err.Source, Err.Description, "BuildAbort")
End Function


'@section The failure exit
'===============================================================================

'@Description("Abort the build and shape the outcome of a failed step.")
'@details
'Called from the handler of every step, with the fault already read off
'Err by the caller. The abort runs under its own error handling, so a
'refusal inside it never reaches the handler that is calling.
'@param errNumber Long. The error number of the fault.
'@param errSource String. The source of the fault. Empty takes the step name.
'@param errDesc String. The description of the fault.
'@param stepName String. The step that failed.
'@return String. "ERROR <number> (<source>): <description>", with "| kept <path>" when a file was kept.
Private Function FailureOutcome(ByVal errNumber As Long, _
                                ByVal errSource As String, _
                                ByVal errDesc As String, _
                                ByVal stepName As String) As String
    Dim kept As String
    Dim sourceText As String

    kept = AbortBuild()

    FailureOutcome = ErrorText(errNumber, errSource, errDesc, stepName)
    If LenB(kept) > 0 Then FailureOutcome = FailureOutcome & KEPT_MARK & kept
End Function

'@Description("Shape the error outcome of a fault, with nothing aborted.")
'@details
'The shape every failed answer of this module takes. The functions that
'answer a value and abort nothing (BuildCheckings, BuildCounts, BuildAbort)
'use it on their own; the steps reach it through FailureOutcome.
'@param errNumber Long. The error number of the fault.
'@param errSource String. The source of the fault. Empty takes the function name.
'@param errDesc String. The description of the fault.
'@param functionName String. The function that failed.
'@return String. "ERROR <number> (<source>): <description>".
Private Function ErrorText(ByVal errNumber As Long, _
                           ByVal errSource As String, _
                           ByVal errDesc As String, _
                           ByVal functionName As String) As String
    Dim sourceText As String

    sourceText = errSource
    If LenB(sourceText) = 0 Then sourceText = MODULE_NAME & "." & functionName

    ErrorText = OUTCOME_ERROR_LEAD & CStr(errNumber) & " (" & sourceText & "): " & errDesc
End Function

'@Description("Keep the unfinished output workbook, close it and drop the build objects.")
'@details
'Linelist.DiscardBuild does the keeping: it saves the workbook as
'__temp.xlsb in the temporary repository, marks the file to survive the
'drop of the repository, closes the workbook and answers the path. A
'build that stopped between Prepare and Linelist.Create still has the
'output workbook on the specifications, so a linelist is made over them
'for the discard alone. Everything runs under On Error Resume Next: this
'is the exit of a failure.
'@return String. The kept path, or empty when nothing was kept.
Private Function AbortBuild() As String
    Dim discarded As Linelist

    On Error Resume Next
    If ll Is Nothing Then
        If Not specs Is Nothing Then
            If Not specs.LLWorkbook Is Nothing Then Set discarded = Linelist.Create(specs)
        End If
    Else
        Set discarded = ll
    End If

    If Not discarded Is Nothing Then keptPath = discarded.DiscardBuild()
    On Error GoTo 0

    DropBuildObjects
    AbortBuild = keptPath
End Function

'@Description("Drop the objects of the build. The counts, the queue and the kept path stay.")
'@details
'Dropping the specifications drops the temporary repository with them,
'which empties the working folder of everything except the kept file.
Private Sub DropBuildObjects()
    Set sheetInfo = Nothing
    Set sheetLists = Nothing
    Set ll = Nothing
    Set specs = Nothing
    Set designerBook = Nothing
End Sub

'@Description("Put every field back to the state before a build.")
Private Sub ResetState()
    DropBuildObjects
    pendingText = vbNullString
    builtSheets = 0
    builtVariables = 0
    keptPath = vbNullString
End Sub


'@section The queue of checkings
'===============================================================================

'@Description("Queue one bundle as text.")
'@param checks Checking. The bundle. Nothing and an empty bundle queue nothing.
'@param recordOnly Optional Boolean. True keeps the bundle off the report worksheet.
Private Sub QueueBundle(ByVal checks As Checking, _
                        Optional ByVal recordOnly As Boolean = False)
    If checks Is Nothing Then Exit Sub
    pendingText = pendingText & GenerationLog.BundleText(checks, recordOnly)
End Sub

'@Description("Queue every bundle of a list.")
'@param bundles BetterArray. Checking bundles.
Private Sub QueueBundles(ByVal bundles As BetterArray)
    Dim index As Long

    If bundles Is Nothing Then Exit Sub
    For index = bundles.LowerBound To bundles.UpperBound
        QueueBundle bundles.Item(index)
    Next index
End Sub

'@Description("Queue the analyses' own bundle, quietly.")
'@details
'Called on the success path and from the failure handler of BuildAnalyses.
'On the failure path a read that raises here would replace the fault the
'step is about to answer, so the read is shielded.
'@param anaOut AnalysisOutput. The analysis writer. Nothing queues nothing.
Private Sub QueueAnalysisBundle(ByVal anaOut As AnalysisOutput)
    If anaOut Is Nothing Then Exit Sub

    On Error Resume Next
    If anaOut.HasCheckings Then QueueBundle anaOut.CheckingValues
    On Error GoTo 0
End Sub

'@Description("Note the __temp.xlsb of an earlier run when the output folder holds one.")
'@details
'The next run on the same output folder overwrites the kept file of the
'last failed one, and the report of that run says so. The repository
'built here to find the path is dropped before Prepare makes its own,
'and the drop empties the folder: the earlier file is scratch from the
'moment this run starts.
'@param outputFolder String. The output folder entry. Empty notes nothing.
Private Sub NoteEarlierKeptFile(ByVal outputFolder As String)
    Dim probe As TemporaryRepos
    Dim earlierPath As String
    Dim found As String
    Dim note As Checking

    If LenB(Trim$(outputFolder)) = 0 Then Exit Sub

    On Error Resume Next
    Set probe = TemporaryRepos.Create(outputFolder)
    earlierPath = probe.RootPath() & KEPT_FILE_NAME
    found = Dir$(earlierPath)
    Set probe = Nothing
    On Error GoTo 0

    If LenB(found) = 0 Then Exit Sub

    Set note = Checking.Create(TITLE_EARLIER_FILE)
    note.Add "earlier file", "The unfinished linelist of an earlier run was found at " & _
                             earlierPath & " and is overwritten by this run", checkingInfo
    QueueBundle note
End Sub


'@section Entries and guards
'===============================================================================

'@Description("Write entries handed over as text onto the Main sheet.")
'@details
'One entry per line, the tag before the first equals sign, the value
'after it. An empty text writes nothing. The tags are the ones
'DesignerEntry.AddInfo knows: setuppath, geopath, lldir, llname,
'llpassword, debugpassword, setuplang, lllang, epiweekstart, design,
'temppath.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@param entriesText String. The entries, or empty.
Private Sub WriteEntries(ByVal entry As DesignerEntry, ByVal entriesText As String)
    Dim entryLines() As String
    Dim index As Long
    Dim lineText As String
    Dim splitAt As Long

    If LenB(entriesText) = 0 Then Exit Sub

    entryLines = Split(entriesText, ENTRY_SEP)
    For index = LBound(entryLines) To UBound(entryLines)
        lineText = entryLines(index)
        splitAt = InStr(lineText, ENTRY_EQUALS)
        If splitAt > 1 Then
            entry.AddInfo Mid$(lineText, splitAt + 1), Left$(lineText, splitAt - 1)
        End If
    Next index
End Sub

'@Description("Raise when BuildBegin has not run.")
'@param stepName String. The step asking.
Private Sub RequireSpecifications(ByVal stepName As String)
    If specs Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   stepName & " needs BuildBegin to have run first"
    End If
End Sub

'@Description("Raise when BuildLinelist has not run.")
'@param stepName String. The step asking.
Private Sub RequireLinelist(ByVal stepName As String)
    RequireSpecifications stepName
    If ll Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   stepName & " needs BuildLinelist to have run first"
    End If
End Sub

'@Description("Raise a typed project error.")
'@param errorCode Long. A ProjectError member.
'@param messageText String. What went wrong, in words.
Private Sub ThrowError(ByVal errorCode As Long, ByVal messageText As String)
    Err.Raise errorCode, MODULE_NAME, messageText
End Sub
