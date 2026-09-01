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

'@class TemporaryRepos
'@description
'Owns one folder on disk that a linelist build writes its working files into.
'Create takes the base directory, or falls back to the default one. RootPath
'answers the folder, EnsureReady creates it, and CreateFilePath builds a path
'inside it from a file name it cleans first. DeleteFile removes one file,
'DeleteAll empties the folder and creates it again, Reset empties it and
'marks the object as needing another EnsureReady, and Remove empties it and
'deletes the folder itself. Terminating an instance empties the folder.
'
'THE KEPT FILE
'-------------------------------------------------------------------------------
'A build that fails saves its unfinished workbook in this folder and drops
'the instance right after. KeepFile marks one file name to survive that
'drop: Class_Terminate deletes every other file and leaves the marked ones.
'Reset, DeleteAll and Remove are explicit calls and wipe the folder whole,
'marks included, so the next run that saves its linelist starts clean.
'@depends ProjectError
'@version 1.2 (2026-09-01)


'@section State and constants
'===============================================================================
Private Type TTempService
    rootPath As String
    isReady As Boolean
    'The file names KeepFile marked, each wrapped in the separator so a
    'name is matched whole: |__temp.xlsb|
    keptNames As String
End Type

Private this As TTempService
Private Const DEFAULT_FOLDER_NAME As String = "OBTApp_"
Private Const KEPT_SEPARATOR As String = "|"


'@section Factory helpers
'===============================================================================
'@label:create
'@sub-title Build and initialise a temporary repository service
'@param baseDirectory String. Optional custom root location.
'@return TemporaryRepos. Ready for use.
'@export
Public Function Create(Optional ByVal baseDirectory As String) As TemporaryRepos
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
    this.keptNames = vbNullString
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
'@details
'The kept marks go with the files: a reset repository starts clean.
'@export
Public Sub Reset()
    DeleteAllInternal
    this.keptNames = vbNullString
    this.isReady = False
End Sub

'@label:remove
'@sub-title Remove every file and the folder itself
'@details
'What Reset does, and then the folder goes too. A build that reached its
'save calls this, so the working folder is not left behind next to the
'linelist it wrote.
'
'The folder is removed only when it is empty. A removal the file system
'refuses -- a file this class cannot see, a folder another program holds
'open -- is left alone: the files this class wrote are gone either way,
'and the next EnsureReady creates the folder again.
'@export
Public Sub Remove()
    Reset

    If LenB(this.rootPath) = 0 Then Exit Sub

    On Error Resume Next
        RmDir NormalizePath(this.rootPath, Application.PathSeparator)
    On Error GoTo 0
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
'@details
'The kept marks go with the files, the way Reset drops them.
'@export
Public Sub DeleteAll()
    EnsureReady
    DeleteAllInternal
    this.keptNames = vbNullString
    EnsureDirectoryExists this.rootPath
End Sub

'@label:keep-file
'@sub-title Mark one file of the repository to survive the drop of the instance
'@details
'The name is cleaned the way CreateFilePath cleans it, so the mark matches
'the file that call wrote. Marking the same name twice is one mark. The
'mark holds until Reset or DeleteAll wipes the folder.
'@param fileName String. The file name inside the repository.
'@export
Public Sub KeepFile(ByVal fileName As String)
    Dim wrapped As String

    wrapped = KEPT_SEPARATOR & SanitizeFileName(fileName) & KEPT_SEPARATOR
    If InStr(1, this.keptNames, wrapped, vbTextCompare) > 0 Then Exit Sub

    this.keptNames = this.keptNames & wrapped
End Sub

'@label:is-kept
'@sub-title Whether a file name carries the keep mark
'@param fileName String. The bare file name.
'@return Boolean. True when KeepFile marked it.
Private Function IsKept(ByVal fileName As String) As Boolean
    If LenB(this.keptNames) = 0 Then Exit Function
    IsKept = InStr(1, this.keptNames, _
                   KEPT_SEPARATOR & fileName & KEPT_SEPARATOR, vbTextCompare) > 0
End Function


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
    Dim readError As Long

    'A statement that succeeds leaves Err holding whatever it held before, so
    'the number is cleared first. Without that, an error raised earlier in the
    'call made this read as "the folder is missing" and the create ran anyway.
    On Error Resume Next
        Err.Clear
        attr = GetAttr(folderPath)
        readError = Err.Number
    On Error GoTo 0

    If readError = 0 Then
        If (attr And vbDirectory) <> vbDirectory Then
            ThrowError ProjectError.ErrorUnexpectedState, _
                       "Path exists but is not a directory: " & folderPath
        End If
        Exit Sub
    End If

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

    sep = Application.PathSeparator ' "\" on Windows, "/" on modern Mac
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
                ' Path or file access error. It usually means the folder is
                ' already there, so check that it is a folder.
                Err.Clear
                If Not IsDirectory(build) Then
                    On Error GoTo 0
                    ThrowError ProjectError.ErrorUnexpectedState, _
                              "A file already exists where a folder is expected: " & build
                End If
            Case 76
                ' Path not found: the parent is missing. The loop builds the
                ' parents in order, so the next turn creates it.
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
'@sub-title Wipe every file inside the repository, keeping the folder itself
'@details
'The whole listing is read before the first file is deleted. Dir$ walks one
'folder through state the host holds, and deleting an entry part way through
'that walk makes the rest of the listing unreliable, so files were left behind.
'The listing is small: this folder holds the working files of one generation.
'
'A file KeepFile marked is left alone when skipKept is True, which is what
'the instance drop asks for. The explicit wipes pass False.
'@param skipKept Optional Boolean. True leaves the marked files on disk.
Private Sub DeleteAllInternal(Optional ByVal skipKept As Boolean = False)
    Dim fileName As String
    Dim names As Collection
    Dim entry As Variant

    If LenB(this.rootPath) = 0 Then Exit Sub

    Set names = New Collection

    On Error Resume Next
        fileName = Dir$(this.rootPath & "*")
        Do While LenB(fileName) > 0
            If fileName <> "." And fileName <> ".." Then names.Add fileName
            fileName = Dir$
        Loop
    On Error GoTo 0

    For Each entry In names
        If Not (skipKept And IsKept(CStr(entry))) Then
            On Error Resume Next
                Kill this.rootPath & CStr(entry)
            On Error GoTo 0
        End If
    Next entry
End Sub

'@label:class-terminate
'@sub-title Clean up temporary files when the instance is destroyed
'@details
'The files KeepFile marked stay: the unfinished linelist of a failed build
'is saved here right before the instance is dropped.
Private Sub Class_Terminate()
    On Error Resume Next
        DeleteAllInternal skipKept:=True
    On Error GoTo 0
End Sub

'@label:throw-error
'@sub-title Raise a typed project error
'@param errNumber ProjectError. The error code.
'@param message String. Descriptive error message.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, "TemporaryRepos", message
End Sub
