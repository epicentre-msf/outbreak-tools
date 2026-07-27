Attribute VB_Name = "OBTGrantAccess"
Attribute VB_Description = "One-time interactive helper: grant Excel persistent sandbox access to the test tree"

Option Explicit

'@Folder("Rubberduck")
'@ModuleDescription("One-time interactive helper: grant Excel persistent sandbox access to the test tree")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

' =============================================================================
' OBTGrantAccess -- pops a macOS folder picker so the operator can grant Excel
' PERSISTENT sandbox access to the test tree (pick the outbreak-tools repo root,
' or at least the untracked working area). Excel for Mac is a SANDBOXED app: VBA file
' reads (Open / Dir / VBComponents.Import) need a security-scoped grant, and
' Full Disk Access does NOT provide it -- only a folder-pick through a dialog
' does (the same mechanism clickDevFolder / OSFiles.LoadFolder use). One pick of
' a parent folder covers every file under it, for any number of test files.
'
' RUN THIS ONCE, BY HAND (VBE -> F5), then pick the repo root. It is NOT part of
' the automated loop. MsgBox feedback is intentional (interactive use only).
'
' Uses MacScript + AppleScript `choose folder`, exactly as OSFiles does. On a
' non-Mac host MacScript is unavailable; the macro reports that and exits.
' =============================================================================

'@EntryPoint
'@sub-title Prompt for a folder and grant Excel persistent read access to that tree.
Public Sub OBTGrantAccess()
    Dim script As String
    Dim chosenPath As String

    script = _
        "set chosenFolder to (choose folder with prompt " & _
        Chr(34) & "Select the outbreak-tools repo folder to grant test access" & Chr(34) & ")" & vbNewLine & _
        "return POSIX path of chosenFolder"

    On Error Resume Next
        chosenPath = MacScript(script)
    On Error GoTo 0

    If LenB(chosenPath) = 0 Then
        MsgBox "No folder was selected -- access was not granted." & vbNewLine & _
               "Run OBTGrantAccess again and pick the outbreak-tools folder.", _
               vbExclamation, "OBT Grant Access"
        Exit Sub
    End If

    MsgBox "Granted Excel access to:" & vbNewLine & vbNewLine & _
           chosenPath & vbNewLine & vbNewLine & _
           "The headless test loop can now read files under this folder " & _
           "without per-file prompts.", _
           vbInformation, "OBT Grant Access"
End Sub
