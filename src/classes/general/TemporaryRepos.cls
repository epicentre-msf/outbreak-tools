VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemporaryRepos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manages temporary files for Linelist preparation with automatic cleanup"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Manages temporary files for Linelist preparation with automatic cleanup")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends ProjectError

Option Explicit

Implements ITemporaryRepos

'@class TemporaryRepos
'@description
'Manages a temporary folder on disk for linelist preparation artefacts.
'Creates, sanitises, and cleans up files within a dedicated repository
'directory. Automatically deletes all content on termination.
'@depends ProjectError
'@version 1.0 (2026-02-09)


'@section State and constants
'===============================================================================
Private Type TTempService
    rootPath As String
    isReady As Boolean
End Type

Private this As TTempService
Private Const DEFAULT_FOLDER_NAME As String = "OBTApp_"


'@section Factory helpers
'===============================================================================
'@label:create
'@sub-title Build and initialise a temporary repository service
'@param baseDirectory String. Optional custom root location.
'@return ITemporaryRepos. Ready for use.
'@export
Public Function Create(Optional ByVal baseDirectory As String) As ITemporaryRepos
    Dim service As TemporaryRepos
    Set service = New TemporaryRepos
    service.Initialise baseDirectory
    Set Create = service
End Function

'@label:initialise
'@sub-title Set the root path and mark the repository as not ready
'@param baseDirectory String. Optional custom root location.
Friend Sub Initialise(Optional ByVal baseDirectory As String)
    this.rootPath = ResolveRootPath(baseDirectory)
    this.isReady = False
End Sub


'@section Lifecycle
'===============================================================================
'@label:ensure-ready
'@sub-title Prepare the repository root if it has not already been created
'@export
Public Sub EnsureReady()
    If this.isReady Then Exit Sub

    EnsureDirectoryExists this.rootPath
    this.isReady = True
End Sub

'@label:reset
'@sub-title Remove all temporary artefacts and mark repository as uninitialised
'@export
Public Sub Reset()
    DeleteAllInternal
    this.isReady = False
End Sub


'@section Paths
'===============================================================================
'@label:root-path
'@prop-title Fully-qualified repository root
'@return String. Absolute path ending with a path separator.
'@export
Public Property Get RootPath() As String
    RootPath = this.rootPath
End Property

'@label:create-file-path
'@sub-title Generate a safe path for a new temporary file within the repository
'@param fileName String. Requested file name (sanitised when required).
'@return String. Absolute path ready for file creation.
'@export
Public Function CreateFilePath(ByVal fileName As String) As String
    EnsureReady
    CreateFilePath = this.rootPath & SanitizeFileName(fileName)
End Function


'@section File operations
'===============================================================================
'@label:delete-file
'@sub-title Delete a single file; silently ignore missing paths
'@param path String. Absolute file path.
'@export
Public Sub DeleteFile(ByVal path As String)
    If LenB(path) = 0 Then Exit Sub
    On Error Resume Next
        Kill path
    On Error GoTo 0
End Sub

'@label:delete-all
'@sub-title Remove every file within the repository, then recreate the root folder
'@export
Public Sub DeleteAll()
    EnsureReady
    DeleteAllInternal
    EnsureDirectoryExists this.rootPath
End Sub


'@section Internal helpers
'===============================================================================
'@label:resolve-root-path
'@sub-title Determine the root directory used for temporary storage
'@param baseDirectory String. Optional path supplied by the caller.
'@return String. Fully-qualified path ending with the repository folder.
Private Function ResolveRootPath(ByVal baseDirectory As String) As String
    Dim root As String

    If LenB(Trim$(baseDirectory)) = 0 Then
        root = Application.DefaultFilePath
        If LenB(root) = 0 Then root = CurDir$
        If LenB(root) = 0 Then
            ThrowError ProjectError.ErrorUnexpectedState, _
                      "Unable to resolve default directory for temporary files"
        End If
    Else
        root = baseDirectory
    End If

    If Right$(root, 1) <> Application.PathSeparator Then
        root = root & Application.PathSeparator
    End If

    ResolveRootPath = root & DEFAULT_FOLDER_NAME & Application.PathSeparator
End Function

'@label:ensure-directory-exists
'@sub-title Create the supplied folder when missing and validate existing paths
'@param folderPath String. Path that should represent a directory.
Private Sub EnsureDirectoryExists(ByVal folderPath As String)
    Dim attr As Long

    'If it exists, ensure it is a directory
    On Error Resume Next
    attr = GetAttr(folderPath)
        If Err.Number = 0 Then
            If (attr And vbDirectory) <> vbDirectory Then
                ThrowError ProjectError.ErrorUnexpectedState, _
                        "Path exists but is not a directory: " & folderPath
            End If
            Exit Sub ' already a folder, we’re done
        End If
    On Error GoTo 0

    'Create the directory tree safely
    CreateDirectoryTree folderPath
End Sub

'@label:create-directory-tree
'@sub-title Create a directory and all missing parents
'@param folderPath String. Fully-qualified target folder.
Private Sub CreateDirectoryTree(ByVal folderPath As String)
    'Creates all missing parents. Treat "already exists" as success.
    Dim i As Long
    Dim start As Long
    Dim sep As String
    Dim parts() As String 
    Dim build As String

    sep = Application.PathSeparator ' "\" on Windows, ":" on old Mac, "/" also works on modern Mac
    parts = Split(NormalizePath(folderPath, sep), sep)

    ' Handle drive / UNC root
    If Left$(folderPath, 2) = "\\" Then
        ' UNC: \\server\share\...
        build = "\\" & parts(2) & sep & parts(3)
        start = 4
    ElseIf InStr(parts(0), ":") > 0 Then
        ' Drive letter root, e.g. C:
        build = parts(0)
        start = 1
    Else
        ' Relative or POSIX-like
        build = IIf(parts(0) = "", sep, parts(0)) ' keep leading "/" if present
        start = IIf(parts(0) = "", 1, 1)
    End If

    For i = start To UBound(parts)
        If Len(build) > 0 And Right$(build, 1) <> sep Then build = build & sep
        build = build & parts(i)
        If Len(build) = 0 Then GoTo NextPart

        'Attempt to create the current segment; tolerate already-existing folders.
        On Error Resume Next
        MkDir build
        Select Case Err.Number
            Case 0
                ' created successfully
            Case 75
                ' Path/File access error — often means it already exists; verify it's a directory
                Err.Clear
                If Not IsDirectory(build) Then
                    On Error GoTo 0
                    ThrowError ProjectError.ErrorUnexpectedState, _
                              "A file already exists where a folder is expected: " & build
                End If
            Case 76
                ' Path not found — parent missing; try to create on next iterations (loop will handle)
                ' (This case usually won’t occur because we’re building parents in order.)
                ' Just continue; next iteration will attempt parent then child.
                Err.Clear
            Case Else
                ' Any other error is surfaced with context
                Dim msg As String
                msg = "Unable to create folder: " & build & " (" & Err.Description & ")"
                On Error GoTo 0
                ThrowError ProjectError.ErrorUnexpectedState, msg
        End Select
        On Error GoTo 0
NextPart:
    Next i

    ' Final verification
    If Not IsDirectory(folderPath) Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Unable to ensure directory exists: " & folderPath
    End If
End Sub


'@label:is-directory
'@sub-title Determine whether the supplied path points to a folder
'@param path String. Path to test.
'@return Boolean. True when the path is an existing directory.
Private Function IsDirectory(ByVal path As String) As Boolean
    Dim a As Long
    On Error Resume Next
    a = GetAttr(path)
    If Err.Number = 0 Then
        IsDirectory = ((a And vbDirectory) = vbDirectory)
    Else
        IsDirectory = False
    End If
    On Error GoTo 0
End Function


'@label:normalize-path
'@sub-title Remove trailing separators and normalise the supplied path string
'@param path String. Raw path.
'@param sep String. Platform path separator.
'@return String. Cleaned path.
Private Function NormalizePath(ByVal path As String, ByVal sep As String) As String
    ' Trim whitespace and remove trailing separators (except root)
    Dim p As String: p = Trim$(path)
    Do While Len(p) > 1 And Right$(p, 1) = sep
        p = Left$(p, Len(p) - 1)
    Loop
    NormalizePath = p
End Function


'@label:sanitize-file-name
'@sub-title Replace characters that are invalid on common file systems
'@param fileName String. Candidate file name.
'@return String. Safe to append to the repository root.
Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim result As String
    Dim index As Long

    If LenB(Trim$(fileName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Temporary file name cannot be empty"
    End If

    invalidChars = Array(":", "*", "?", "\", "/", "|", "<", ">", Chr$(34), ",")
    result = fileName

    For index = LBound(invalidChars) To UBound(invalidChars)
        result = Replace$(result, invalidChars(index), "_")
    Next index

    SanitizeFileName = result
End Function

'@label:delete-all-internal
'@sub-title Wipe every file inside the repository without recreating the folder
Private Sub DeleteAllInternal()
    Dim filePath As String

    On Error Resume Next
        filePath = Dir$(this.rootPath & "*")
        Do While LenB(filePath) > 0
            If filePath <> "." And filePath <> ".." Then
                Kill this.rootPath & filePath
            End If
            filePath = Dir$
        Loop
    On Error GoTo 0
End Sub

'@label:class-terminate
'@sub-title Clean up temporary files when the instance is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
        DeleteAllInternal
    On Error GoTo 0
End Sub

'@label:throw-error
'@sub-title Raise a typed project error
'@param errNumber ProjectError. The error code.
'@param message String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, "TemporaryRepos", message
End Sub


'@section Interface
'===============================================================================
'ITemporaryRepos delegation wiring.

Private Sub ITemporaryRepos_EnsureReady()
    EnsureReady
End Sub

Private Sub ITemporaryRepos_Reset()
    Reset
End Sub

Private Property Get ITemporaryRepos_RootPath() As String
    ITemporaryRepos_RootPath = RootPath
End Property

Private Function ITemporaryRepos_CreateFilePath(ByVal fileName As String) As String
    ITemporaryRepos_CreateFilePath = CreateFilePath(fileName)
End Function

Private Sub ITemporaryRepos_DeleteFile(ByVal path As String)
    DeleteFile path
End Sub

Private Sub ITemporaryRepos_DeleteAll()
    DeleteAll
End Sub
