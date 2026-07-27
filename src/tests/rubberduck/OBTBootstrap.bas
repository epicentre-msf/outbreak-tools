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
' Keep THIS module minimal and correct: it is the one piece that still needs a
' manual re-import if it ever changes. It never refreshes itself (a module
' cannot replace itself while its code is on the call stack), and it must not
' depend on OBTImport/OBTHeadless (the very modules it swaps out).
'
' Requires "Trust access to the VBA project object model" = ON (already needed
' for @TestMethod discovery).
' =============================================================================

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
