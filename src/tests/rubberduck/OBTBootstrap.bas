Attribute VB_Name = "OBTBootstrap"
Attribute VB_Description = "Stable bootstrap: refreshes the harness modules from the run dir before each run"

Option Explicit

'@Folder("Rubberduck")
'@ModuleDescription("Stable bootstrap: refreshes the harness modules from the run dir before each run")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

' =============================================================================
' OBTRefreshHarness -- re-imports the harness modules (OBTImport, OBTHeadless)
' from the assembled run dir so their code can be ITERATED without a manual VBE
' re-import every time. run-tests.R copies the current harness sources into
' <workbook folder>/bootstrap/; the AppleScript trigger calls OBTRefreshHarness
' as the FIRST macro of each run -- a separate `run VB macro` call, so the
' refreshed modules are fully loaded before OBTBuildCodeTables / OBTSilentImport
' / OBTRunAllTests run.
'
' OBTGrantRoot -- asks the macOS sandbox for ONE folder, and it runs BEFORE
' OBTRefreshHarness because OBTRefreshHarness is itself a file read. It lives
' here rather than in OBTImport for exactly that reason: the module that grants
' access cannot be one of the modules that has to be read off disk first.
'
' Keep THIS module minimal and correct: it is the one piece that still needs a
' manual re-import if it ever changes. It never refreshes itself (a module
' cannot replace itself while its code is on the call stack), and it must not
' depend on OBTImport/OBTHeadless (the very modules it swaps out).
'
' Requires "Trust access to the VBA project object model" = ON (already needed
' for @TestMethod discovery).
' =============================================================================

'@EntryPoint
'@fun-title Grant Excel persistent sandbox access to one folder and its whole tree.
'@details
'Excel for Mac is SANDBOXED: a VBA file read on a path it holds no
'security-scoped grant for pops a dialog, and a dialog in a headless run is a
'hang. GrantAccessToMultipleFiles is the API for it.
'
'CALL IT FIRST, and that is the part this fixes with certainty. Everything else
'in the trigger reads or writes under the root, OBTRefreshHarness included, and
'OBTRefreshHarness runs immediately after this. Nothing used to grant the run
'dir at all, so the harness modules prompted on every single run.
'
'One path rather than a list because the launcher now stages the whole run
'under one root. Whether that reaches ZERO prompts also needs a folder grant to
'PERSIST across an Excel restart and to CASCADE to files created in the tree
'afterwards. Neither is established -- see .obt/gotchas/macos-sandbox-grant.md,
'which records a probe that appeared to prove both and proved nothing.
'
'The call is late-bound through Object because the member exists only in the
'Mac type library, and one identifier from a missing library costs the whole
'project its compile. On Windows there is no sandbox and there is nothing to do.
'@param rootPath String. The folder to grant, as a full path.
'@return String. What happened, for the trigger to put in the run report.
Public Function OBTGrantRoot(ByVal rootPath As String) As String
    Dim host As Object
    Dim granted As Boolean

    If LenB(Trim$(rootPath)) = 0 Then
        OBTGrantRoot = "no root given - expect prompts"
        Exit Function
    End If

    'Answered rather than swallowed. The earlier wrapper reported nothing at
    'all, so no run ever recorded whether a grant had been made, and the
    'question stayed open for weeks.
    On Error Resume Next
        Set host = Application
        granted = host.GrantAccessToMultipleFiles(Array(rootPath))
        If Err.Number <> 0 Then
            OBTGrantRoot = "not available on this host (" & Err.Description & ")"
            Err.Clear
            Exit Function
        End If
    On Error GoTo 0

    If granted Then
        OBTGrantRoot = "granted " & rootPath
    Else
        OBTGrantRoot = "REFUSED " & rootPath & " - expect prompts"
    End If
End Function

'@EntryPoint
'@sub-title Re-import the harness modules from the run dir (run first, before the loop).
Public Sub OBTRefreshHarness()
    Dim base As String
    base = ThisWorkbook.Path & Application.PathSeparator & "bootstrap"

    ReimportComponent "OBTImport", base & Application.PathSeparator & "OBTImport.bas"
    ReimportComponent "OBTHeadless", base & Application.PathSeparator & "OBTHeadless.bas"
End Sub

'@sub-title Remove any existing component of this name, then import the file if present.
'@param compName String. The VBA component to replace.
'@param filePath String. The .bas source to import from the run dir.
Private Sub ReimportComponent(ByVal compName As String, ByVal filePath As String)
    If Dir$(filePath) = vbNullString Then Exit Sub

    Dim proj As Object
    Set proj = ThisWorkbook.VBProject

    On Error Resume Next
        proj.VBComponents.Remove proj.VBComponents(compName)
    On Error GoTo 0

    proj.VBComponents.Import filePath
End Sub
