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
' STATUS: run against a real Windows Excel for the first time on 2026-08-17,
' and green. Excel 16.0.20228 on Windows 11, the measles setup, 2 data entry
' sheets, 128 variables, the geobase imported, zero error lines in the
' generation log, 2,797,064 bytes delivered in about 60 seconds. The contract
' it drives is open copy -> refresh -> build tables -> import -> build ->
' report file -> quit.
' ============================================================================

Option Explicit

Dim Arg
Set Arg = WScript.Arguments

If Arg.Count < 11 Then
    WScript.Echo "usage: cscript //nologo build-linelist.vbs <workbook> " & _
                 "<designer> <setup> <sourceRoot> <formsFolder> <outFolder> " & _
                 "<outName> <options> <reportPath> <obtHome> <needPick> [<rows>]"
    WScript.Quit 2
End If

Dim wbPath, designerPath, setupPath, sourceRoot, formsFolder
Dim outFolder, outName, buildOptions, reportPath, obtHome, needPick, rowsSpec

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
needPick     = Arg(10)

' THE TWELFTH ARGUMENT PICKS THE ENTRY POINT.
' ----------------------------------------------------------------------------
' Empty, or absent, and this drives BuildLinelistFromSetup over <setup>, which
' is every run this script did before. Filled, it carries the rows of a
' multiple generation and the run goes through BuildMultipleFromTable instead:
' the designer's own T_Multi loop, one linelist per row, over a designer copy.
' <setup> is then unread, because each row names its own.
'
' One trigger for both, so the two entry points share the four steps around
' them -- refresh, tables, import, report -- and a difference between a single
' build and a multi build stays a difference in the VBA.
rowsSpec = ""
If Arg.Count >= 12 Then rowsSpec = Arg(11)

Dim outcomeText, summaryText, grantText
outcomeText = ""
summaryText = ""
grantText = ""

Dim xlsApp, Wkb, wbName
' Initialised so Cleanup can test them with Is Nothing even when Open never ran.
Set xlsApp = Nothing
Set Wkb = Nothing

' EXCEL STAYS ON SCREEN, AND THE BUILD NEEDS IT THERE.
' ----------------------------------------------------------------------------
' A hidden Excel refuses Window.FreezePanes with error 1004, "Unable to set the
' FreezePanes property of the Window class". LLDataEntry freezes the header of
' every data entry sheet, so the run stopped on the first one and delivered
' nothing.
'
' Measured 2026-08-17 on Excel 16.0.20228, same source, same setup, one line
' changed: hidden, the trace ends at `building data entry sheet Info`; on
' screen, the same build writes 2 sheets, 128 variables and a 2,797,064 byte
' linelist with the geobase in it.
'
' A plain Workbooks.Add in a hidden Excel takes the freeze without complaint,
' so the refusal needs the build's own conditions to show up and a small probe
' will say everything is fine. Trust the build.
'
' The macOS sibling has always left Excel on screen -- AppleScript's `open
' workbook` shows it -- so this also keeps the two triggers driving one Excel
' in one state, which is the whole reason they are written as siblings.
'
' The cost is that the run takes over the screen for its minute. Excel is its
' own instance, so anything the operator has open is untouched.
Set xlsApp = CreateObject("Excel.Application")
xlsApp.Visible = True
xlsApp.DisplayAlerts = False
xlsApp.ScreenUpdating = False

' Any failure below leaves Excel running as an invisible orphan that wedges the
' NEXT run, so trap from here on and always reach the cleanup.
On Error Resume Next

' Read/write: the run imports components into the copy. Arg 3 = ReadOnly.
Set Wkb = xlsApp.Workbooks.Open(wbPath, , False)
If Err.Number <> 0 Then Fail "Workbooks.Open"
wbName = Wkb.Name

' 0) the folder pick, first because every call after it reads a file. Windows
'    has no sandbox and MacScript is absent, so this never runs here: R only
'    sets needPick on macOS. Kept so the two triggers read the same.
If needPick = "1" Then
    xlsApp.Run wbName & "!" & "OBTGrantAccess"
    If Err.Number <> 0 Then
        grantText = "PICKER FAILED " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        grantText = "folder picker was shown for " & obtHome
    End If
Else
    grantText = "step 0 skipped, no folder picker was opened"
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
If rowsSpec = "" Then
    outcomeText = xlsApp.Run(wbName & "!" & "HeadlessBuild.BuildLinelistFromSetup", _
                             designerPath, setupPath, sourceRoot, formsFolder, _
                             outFolder, outName, buildOptions, obtHome)
Else
    outcomeText = xlsApp.Run(wbName & "!" & "HeadlessBuild.BuildMultipleFromTable", _
                             designerPath, rowsSpec, sourceRoot, formsFolder, _
                             outFolder, outName, buildOptions, obtHome)
End If
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
