' run-tests.vbs
' ============================================================================
' Windows trigger for the test harness — the sibling of the macOS
' AppleScript trigger, and of scripts/devtools/rundesigner.vbs. On Windows Excel
' exposes a full COM object model, so VBScript creates the Application directly.
' The macOS path has to go through Apple Events.
'
' Invoked as:
'   cscript //nologo run-tests.vbs "<workbook-copy-path>" [build|nobuild]
'
'   Arg(0)  absolute path to the per-run workbook COPY (never the original)
'   Arg(1)  "build" -> also run OBTBuildCodeTables; else just import + run
'
' Runs the SAME parameterless OBT* entry points in the SAME order as
' scripts/tests/macos/run-tests.applescript, so the two platforms exercise one
' identical registry. The point is to catch Windows/Mac behavioural drift (for
' example the Byte-enum rule).
'
' ORDER MATTERS: OBTBuildCodeTables runs BEFORE OBTSilentImport, because the
' silent import reads the freshly rebuilt Codes tables to decide what to pull
' from src/. The chain is refresh-harness -> build (optional) -> import -> run.
'
' The entry points carry no underscores. In a VBA class or document module
' Foo_Bar parses as event Bar of object Foo, so the project keeps procedure
' names underscore-free everywhere.
'
' STATUS: Phase D, not wired up, and never executed against a real Windows
' Excel. The intended Windows path is a trigger-file watcher (a scheduled task
' polling the shared repo for a .trigger, then calling this script). Host->guest
' SSH and prlctl are out: Parallels is Standard edition, and SSH is ruled out by
' preference. See src/tests/automated-testing-macos.md for the contract this
' mirrors. Left as a faithful stub so the Mac orchestrator's contract
' (open copy -> OBT* -> CSV beside copy -> quit) is documented for whoever
' implements the guest side.
' ============================================================================

Option Explicit

Dim Arg, wbPath, doBuild
Set Arg = WScript.Arguments

If Arg.Count < 1 Then
    WScript.Echo "usage: cscript //nologo run-tests.vbs <workbook-copy-path> [build|nobuild]"
    WScript.Quit 2
End If

wbPath  = Arg(0)
doBuild = False
If Arg.Count >= 2 Then
    If LCase(Arg(1)) = "build" Then doBuild = True
End If

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

' The loop MUTATES the workbook (imports components via Development, rebuilds
' the Codes tables, writes results) and Saves before returning, so open it
' read/write. Arg 3 = ReadOnly.
Set Wkb = xlsApp.Workbooks.Open(wbPath, , False)
If Err.Number <> 0 Then Fail "Workbooks.Open"
wbName = Wkb.Name

' 0) refresh the harness modules (OBTImport/OBTHeadless) from the run dir, so
'    their code can be iterated without a manual VBE re-import.
xlsApp.Run wbName & "!" & "OBTRefreshHarness"
If Err.Number <> 0 Then Fail "OBTRefreshHarness"

' 1) optionally rebuild Codes tables + ModulesForTesting from the registry
If doBuild Then
    xlsApp.Run wbName & "!" & "OBTBuildCodeTables"
    If Err.Number <> 0 Then Fail "OBTBuildCodeTables"
End If

' 2) refresh workbook code from src/ (no dialogs) via Development
xlsApp.Run wbName & "!" & "OBTSilentImport"
If Err.Number <> 0 Then Fail "OBTSilentImport"

' 3) run every registered module, serialize testsOutputs to CSV beside the copy, Save
xlsApp.Run wbName & "!" & "OBTRunAllTests"
If Err.Number <> 0 Then Fail "OBTRunAllTests"

On Error GoTo 0
Cleanup
WScript.Quit 0

' ----------------------------------------------------------------------------

' Report which entry point raised, close Excel down, and exit non-zero so the
' caller can tell a broken run from a run with failing tests.
Sub Fail(stepName)
    WScript.Echo "run-tests.vbs: " & stepName & " failed: " & _
                 Err.Number & " " & Err.Description
    Err.Clear
    Cleanup
    WScript.Quit 1
End Sub

' Close without prompting and quit. VBA has already Saved on the happy path;
' discarding changes on the failure path keeps the copy reusable.
Sub Cleanup()
    On Error Resume Next
    If Not Wkb Is Nothing Then Wkb.Close False
    If Not xlsApp Is Nothing Then xlsApp.Quit
    Set Wkb = Nothing
    Set xlsApp = Nothing
    Set Arg = Nothing
End Sub
