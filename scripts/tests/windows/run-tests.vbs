' run-tests.vbs
' ============================================================================
' Windows trigger for the test harness — the sibling of the macOS
' AppleScript trigger, and of scripts/devtools/rundesigner.vbs. On Windows,
' Excel exposes a full COM object model, so (unlike macOS) VBScript can create
' the Application directly rather than going through Apple Events.
'
' Invoked as:
'   cscript //nologo run-tests.vbs "<workbook-copy-path>" [build|nobuild]
'
'   Arg(0)  absolute path to the per-run workbook COPY (never the original)
'   Arg(1)  "build" -> also run OBT_BuildCodeTables; else just import + run
'
' Runs the SAME parameterless OBT_* entry points as the Mac path, so the two
' platforms exercise one identical registry — the whole point of Phase D is to
' catch Windows/Mac behavioural drift (e.g. the Byte-enum rule).
'
' STATUS: Phase D, not wired up. The intended Windows path is a trigger-file
' watcher (a scheduled task polling the shared repo for a .trigger, then
' calling this script) — NOT host->guest SSH/prlctl (Parallels is Standard
' edition; SSH is out by preference). See .obt/plans/automated-testing.md
' Phase D and .obt/plans/test-scripts-status.md. Left as a faithful stub so the
' Mac orchestrator's contract (open copy -> OBT_* -> CSV beside copy -> quit)
' is documented for whoever implements the guest side.
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
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = False
xlsApp.DisplayAlerts = False
xlsApp.ScreenUpdating = False

Set Wkb = xlsApp.Workbooks.Open(wbPath)
wbName = Wkb.Name

' 1) refresh workbook code from src/ (no dialogs)
xlsApp.Run wbName & "!" & "OBT_SilentImport"

' 2) optionally rebuild Codes tables + ModulesForTesting from the registry
If doBuild Then
    xlsApp.Run wbName & "!" & "OBT_BuildCodeTables"
End If

' 3) run every registered module, serialize testsOutputs to CSV beside the copy, Save
xlsApp.Run wbName & "!" & "OBT_RunAllTests"

' VBA has already Saved; close without prompting and quit.
Wkb.Close False
xlsApp.Quit

Set Wkb = Nothing
Set xlsApp = Nothing
Set Arg = Nothing
