' build-linelist.vbs
' ============================================================================
' Windows trigger for the headless linelist build — the sibling of
' scripts/headless/macos/build-linelist.applescript, and the counterpart of
' scripts/tests/windows/run-tests.vbs. On Windows Excel exposes a full COM
' object model, so VBScript creates the Application directly. The macOS path has
' to go through Apple Events, which is the only real difference between the two.
'
' Invoked by scripts/headless/build-linelist.R:
'   cscript //nologo build-linelist.vbs "<workbook>" "<designer>" "<setup>" _
'           "<sourceRoot>" "<formsFolder>" "<outFolder>" "<outName>" _
'           "<options>" "<reportPath>" "<obtHome>"
'
' <obtHome> is the ONE folder every other path sits under. It exists for the
' macOS sandbox and does nothing here: Windows has no sandbox, OBTGrantRoot
' answers "not available on this host", and the run reads what it likes. The
' argument is still passed and still called, so the two triggers drive the same
' entry points in the same order and a difference in what gets built stays a
' difference in the VBA.
'
' Runs NO test. It calls HeadlessBuild.BuildLinelistFromSetup directly with the
' eight paths R resolved, in the same order and with the same arguments as the
' AppleScript, so the two platforms drive one identical build.
'
' ORDER MATTERS, and it is the order the test trigger uses for its own reasons:
' OBTBuildCodeTables reads the run dir's .generated intermediates and rebuilds
' the Codes import tables from them, then OBTSilentImport reads those tables to
' decide what to pull from the staged source. R has already filtered the
' intermediates down to the classes and modules alone, so what lands in the
' driver is the build closure with no test module in it.
'
' WHY THE SUMMARY IS ONE CALL
' ----------------------------------------------------------------------------
' The six things worth knowing about a run — the path written, the log, the
' sheet, variable and component counts, and the narrative — are exposed on
' HeadlessBuild as Property Get, and Application.Run cannot reach a Property
' Get on either platform. HeadlessBuild.LastBuildSummary is the Function that
' can, and it answers all six in one string.
'
' THE REPORT FILE IS WRITTEN IN THE HOST ANSI CODE PAGE
' ----------------------------------------------------------------------------
' The AppleScript writes UTF-8; this writes what FileSystemObject writes, which
' is the host code page, and R reads its own platform's native encoding on each
' side. A path carrying characters outside the host code page is the known
' limit of that choice. Writing BOM-less UTF-8 from VBScript needs the
' ADODB.Stream binary-copy dance, and it buys nothing until a path needs it.
'
' STATUS: never executed against a real Windows Excel, exactly like
' run-tests.vbs beside it. The Windows path is a trigger-file watcher (a
' scheduled task polling the shared repo for a .trigger, then calling this
' script); host->guest SSH and prlctl are out. Left as a faithful sibling so
' the contract — open copy -> refresh -> build tables -> import -> build ->
' report file -> quit — is written down for whoever implements the guest side.
' ============================================================================

Option Explicit

Dim Arg
Set Arg = WScript.Arguments

If Arg.Count < 10 Then
    WScript.Echo "usage: cscript //nologo build-linelist.vbs <workbook> " & _
                 "<designer> <setup> <sourceRoot> <formsFolder> <outFolder> " & _
                 "<outName> <options> <reportPath> <obtHome>"
    WScript.Quit 2
End If

Dim wbPath, designerPath, setupPath, sourceRoot, formsFolder
Dim outFolder, outName, buildOptions, reportPath, obtHome

wbPath       = Arg(0)
designerPath = Arg(1)
setupPath    = Arg(2)
sourceRoot   = Arg(3)
formsFolder  = Arg(4)
outFolder    = Arg(5)
outName      = Arg(6)
buildOptions = Arg(7)
reportPath   = Arg(8)
obtHome      = Arg(9)

Dim outcomeText, summaryText, grantText
outcomeText = ""
summaryText = ""
grantText = ""

Dim xlsApp, Wkb, wbName
' Initialised so Cleanup can test them with Is Nothing even when Open never ran.
Set xlsApp = Nothing
Set Wkb = Nothing

Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.DisplayAlerts = False
xlsApp.ScreenUpdating = False

' Any failure below leaves Excel running as an invisible orphan that wedges the
' NEXT run, so trap from here on and always reach the cleanup.
On Error Resume Next

' Read/write: the run imports components into the copy. Arg 3 = ReadOnly.
Set Wkb = xlsApp.Workbooks.Open(wbPath, , False)
If Err.Number <> 0 Then Fail "Workbooks.Open"
wbName = Wkb.Name

' 0) the sandbox grant, first because every call after it reads a file. It is a
'    no-op on this host and is called anyway, so the two triggers stay in step.
grantText = xlsApp.Run(wbName & "!" & "OBTGrantRoot", obtHome)
If Err.Number <> 0 Then
    grantText = "GRANT CALL FAILED " & Err.Number & ": " & Err.Description
    Err.Clear
End If

' 1) refresh the harness modules from the run dir, so OBTImport can be iterated
'    without a manual VBE re-import.
xlsApp.Run wbName & "!" & "OBTRefreshHarness"
If Err.Number <> 0 Then Fail "OBTRefreshHarness"

' 2) rebuild the Codes tables from the filtered intermediates. Always, never
'    optional: the driver may be holding the tables a test run left behind, and
'    those carry the test folders too.
xlsApp.Run wbName & "!" & "OBTBuildCodeTables"
If Err.Number <> 0 Then Fail "OBTBuildCodeTables"

' 3) pull the staged source into the driver
xlsApp.Run wbName & "!" & "OBTSilentImport"
If Err.Number <> 0 Then Fail "OBTSilentImport"

' 4) the build. A raise here is RECORDED rather than fatal, so the report file
'    still says how far the run got instead of leaving only an exit code.
outcomeText = xlsApp.Run(wbName & "!" & "HeadlessBuild.BuildLinelistFromSetup", _
                         designerPath, setupPath, sourceRoot, formsFolder, _
                         outFolder, outName, buildOptions, obtHome)
If Err.Number <> 0 Then
    outcomeText = "TRIGGER ERROR " & Err.Number & ": " & Err.Description
    Err.Clear
End If

' 5) what the build recorded. Read it whatever the outcome was: a run that
'    stopped halfway still holds the narrative up to the stop, and that
'    narrative is the whole diagnostic.
summaryText = xlsApp.Run(wbName & "!" & "HeadlessBuild.LastBuildSummary")
If Err.Number <> 0 Then
    summaryText = ""
    Err.Clear
End If

WriteReport reportPath, outcomeText, grantText, summaryText

On Error GoTo 0
Cleanup
WScript.Quit 0

' ----------------------------------------------------------------------------

' The outcome is the one field the summary cannot carry: it is what the build
' ANSWERED, and a build that raised never returned to say it.
Sub WriteReport(path, outcomeValue, grantValue, summaryValue)
    On Error Resume Next

    Dim fso, handle
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ForWriting = 2, create if absent. Truncates the file in place, which is
    ' what the macOS side does and for a reason that only bites there: a grant
    ' is tied to the file identity, so a file swept and recreated is a new file
    ' the grant no longer covers. Same call here keeps the two in step.
    Set handle = fso.OpenTextFile(path, 2, True)
    If Err.Number = 0 Then
        handle.WriteLine "outcome=" & outcomeValue
        handle.WriteLine "grant=" & grantValue
        handle.WriteLine summaryValue
        handle.Close
    End If

    Err.Clear
    On Error GoTo 0
End Sub

' Report which entry point raised, close Excel down, and exit non-zero so the
' caller can tell a broken run from a build that reported its own fault.
Sub Fail(stepName)
    WScript.Echo "build-linelist.vbs: " & stepName & " failed: " & _
                 Err.Number & " " & Err.Description
    Err.Clear
    WriteReport reportPath, "TRIGGER ERROR at " & stepName, grantText, ""
    Cleanup
    WScript.Quit 1
End Sub

' Close without prompting and quit. Nothing here needs the driver saved: the
' build writes its own files and the driver is a copy.
Sub Cleanup()
    On Error Resume Next
    If Not Wkb Is Nothing Then Wkb.Close False
    If Not xlsApp Is Nothing Then xlsApp.Quit
    Set Wkb = Nothing
    Set xlsApp = Nothing
    Set Arg = Nothing
End Sub
