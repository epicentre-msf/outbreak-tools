Attribute VB_Name = "OBTProbeSandbox"
Attribute VB_Description = "Settle whether ONE folder grant persists and cascades to files made later"

Option Explicit

'@Folder("Rubberduck")
'@ModuleDescription("Settle whether ONE folder grant persists and cascades to files made later")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

' =============================================================================
' OBTProbeSandbox -- the experiment that decides the headless layout.
'
' THE QUESTION
' -----------------------------------------------------------------------------
' If Excel is granted ONE folder, does that grant
'   (a) survive Excel being quit and relaunched, and
'   (b) cover files created inside that folder AFTER the grant was made?
'
' Both must hold for "one grant ask, ever" to work. Nothing in the project
' confirms either, and elapsed time and exit codes cannot see a grant panel.
' The person at the machine is the only oracle. Watch the screen.
'
' HOW TO RUN IT
' -----------------------------------------------------------------------------
'   1. Paste this module into any workbook's VBE.
'   2. F5 on OBTProbeStepOne. ONE grant panel is expected -- click Grant.
'   3. Quit Excel from its own File menu. Not pkill: a bookmark is flushed on
'      normal termination and a killed Excel may never write it.
'   4. Relaunch Excel, reopen the workbook, F5 on OBTProbeStepTwo.
'
'      NO PANEL in step two  -> the grant persists AND cascades. One granted
'                               root is enough and the headless run moves into
'                               one folder.
'      A PANEL in step two   -> report which paths it listed. That tells us
'                               whether persistence or cascade is the one that
'                               failed, and the layout changes accordingly.
'
'   5. OBTProbeCleanUp removes the probe tree when you are done.
'
' The probe root sits in the home folder because that is what a deployed
' package folder looks like. It is granted from INSIDE Excel's process, which
' is the only place a grant can be made at all.
' =============================================================================

' The root is SPELT OUT, and that is the whole correction to this probe.
'
' The first version derived it from Environ$("HOME"). Inside a sandboxed Excel
' that answers the CONTAINER home, ~/Library/Containers/com.microsoft.Excel/Data,
' and everything under the container is reachable with no grant of any kind. So
' the probe wrote where a grant was never needed, saw no panel, and proved
' nothing about persistence or cascade. A run that cannot fail is not a test.
'
' This path is outside the container and outside every folder Excel is granted
' by default, which is the only place the question can be asked.
Private Const PROBE_ROOT As String = "/Users/komlaviamevoin/OBTProbe"
Private Const PROBE_SUB As String = "made-later"


'@EntryPoint
'@sub-title Create the root, grant it, and record what the grant answered.
Public Sub OBTProbeStepOne()
    Dim root As String
    Dim granted As Boolean
    Dim host As Object

    root = ProbeRoot()

    'The root and its subfolder exist BEFORE the grant. A folder that does not
    'exist cannot be bookmarked, and TemporaryRepos-style late creation is the
    'trap this project already fell into once.
    MakeFolder root
    MakeFolder root & PROBE_SUB

    'One file written before the grant. Step two compares its readability
    'against a file written after, which is the whole cascade question.
    WriteLine root & PROBE_SUB & "/before-grant.txt", "written before the grant"

    On Error Resume Next
        Set host = Application
        granted = host.GrantAccessToMultipleFiles(Array(root))
        Err.Clear
    On Error GoTo 0

    MsgBox "Granted the ROOT only:" & vbNewLine & vbNewLine & _
           root & vbNewLine & vbNewLine & _
           "GrantAccessToMultipleFiles answered: " & CStr(granted) & vbNewLine & vbNewLine & _
           "Now quit Excel FROM ITS OWN FILE MENU, relaunch, and run " & _
           "OBTProbeStepTwo.", _
           vbInformation, "OBT sandbox probe 1 of 2"
End Sub


'@EntryPoint
'@sub-title After a relaunch, touch the tree with NO grant call and see what asks.
Public Sub OBTProbeStepTwo()
    Dim root As String
    Dim madeLater As String
    Dim readBack As String
    Dim outcome As String

    root = ProbeRoot()
    madeLater = root & PROBE_SUB & "/after-grant.txt"

    'Deliberately NO grant call anywhere in this Sub. Everything below leans on
    'the bookmark step one asked for.

    outcome = "existing file (written before the grant):" & vbNewLine & "  "
    outcome = outcome & ReadFirstLine(root & PROBE_SUB & "/before-grant.txt")

    'The cascade test: a file that did not exist when the grant was made.
    WriteLine madeLater, "written after the grant, in a later Excel session"
    readBack = ReadFirstLine(madeLater)

    outcome = outcome & vbNewLine & vbNewLine & _
              "file created AFTER the grant, this session:" & vbNewLine & "  " & readBack

    'Dir$ over the folder, which is what the import path does before it reads.
    outcome = outcome & vbNewLine & vbNewLine & _
              "Dir$ over the subfolder answered:" & vbNewLine & "  " & _
              FirstEntry(root & PROBE_SUB)

    MsgBox outcome & vbNewLine & vbNewLine & _
           "Did a grant panel appear during this Sub? That answer, not this " & _
           "text, is the result of the probe.", _
           vbInformation, "OBT sandbox probe 2 of 2"
End Sub


'@EntryPoint
'@sub-title Remove the probe tree.
Public Sub OBTProbeCleanUp()
    Dim root As String

    root = ProbeRoot()

    On Error Resume Next
        Kill root & PROBE_SUB & "/*.txt"
        RmDir root & PROBE_SUB
        RmDir root
    On Error GoTo 0

    MsgBox "Probe tree removed from " & root, vbInformation, "OBT sandbox probe"
End Sub


'@Description("The probe root, with a trailing separator.")
'@return String. The spelt-out root. See the note on PROBE_ROOT.
Private Function ProbeRoot() As String
    ProbeRoot = PROBE_ROOT
    If Right$(ProbeRoot, 1) <> "/" Then ProbeRoot = ProbeRoot & "/"
End Function

'@EntryPoint
'@sub-title Report where Excel thinks HOME is, which is what voided the first probe.
Public Sub OBTProbeWhereIsHome()
    MsgBox "Environ$(""HOME"") inside this Excel is:" & vbNewLine & vbNewLine & _
           Environ$("HOME") & vbNewLine & vbNewLine & _
           "A path under Library/Containers/com.microsoft.Excel means the " & _
           "sandbox is answering with the container, and anything written " & _
           "there needs no grant at all.", _
           vbInformation, "OBT sandbox probe - where is home"
End Sub


'@Description("Make one folder, saying nothing when it is already there.")
'@param path String. The folder to create.
Private Sub MakeFolder(ByVal path As String)
    On Error Resume Next
        MkDir path
    On Error GoTo 0
End Sub


'@Description("Write one line to a file, truncating whatever was there.")
'@param path String. The file to write.
'@param text String. The line.
Private Sub WriteLine(ByVal path As String, ByVal text As String)
    Dim fileNum As Integer

    On Error Resume Next
        fileNum = FreeFile
        Open path For Output As #fileNum
        Print #fileNum, text
        Close #fileNum
    On Error GoTo 0
End Sub


'@Description("Read the first line of a file, or say why it could not be read.")
'@param path String. The file to read.
'@return String. The line, or "COULD NOT READ - <error>".
Private Function ReadFirstLine(ByVal path As String) As String
    Dim fileNum As Integer
    Dim buffer As String

    On Error GoTo Failed

    fileNum = FreeFile
    Open path For Input As #fileNum
    Line Input #fileNum, buffer
    Close #fileNum

    ReadFirstLine = buffer
    Exit Function

Failed:
    ReadFirstLine = "COULD NOT READ - " & Err.Description
    On Error Resume Next
        Close #fileNum
    On Error GoTo 0
End Function


'@Description("The first entry Dir$ finds in a folder.")
'@param path String. The folder to list.
'@return String. The entry name, or a report of what went wrong.
Private Function FirstEntry(ByVal path As String) As String
    Dim found As String

    On Error GoTo Failed

    found = Dir$(path & "/*.txt")
    If LenB(found) = 0 Then
        FirstEntry = "nothing (Dir$ found no file)"
    Else
        FirstEntry = found
    End If
    Exit Function

Failed:
    FirstEntry = "COULD NOT LIST - " & Err.Description
End Function
