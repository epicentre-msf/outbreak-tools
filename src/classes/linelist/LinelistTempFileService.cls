VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistTempFileService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Manages temporary files for Linelist preparation with automatic cleanup"

'@Folder("Linelist")
'@ModuleDescription("Manages temporary files for Linelist preparation with automatic cleanup")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistTempFileService

'@section State and constants
'===============================================================================
Private Type TTempService
    rootPath As String
    isReady As Boolean
End Type

Private this As TTempService
Private Const CLASSNAME As String = "LinelistTempFileService"
Private Const DEFAULT_FOLDER_NAME As String = "LinelistApp_"


'@section Factory helpers
'===============================================================================
Public Function Create(Optional ByVal baseDirectory As String) As ILinelistTempFileService
    Dim service As LinelistTempFileService
    Set service = New LinelistTempFileService
    service.Initialise baseDirectory
    Set Create = service
End Function

Friend Sub Initialise(Optional ByVal baseDirectory As String)
    this.rootPath = ResolveRootPath(baseDirectory)
    this.isReady = False
End Sub


'@section Lifecycle
'===============================================================================
Public Sub EnsureReady()
    If this.isReady Then Exit Sub

    EnsureDirectoryExists this.rootPath
    this.isReady = True
End Sub

Public Sub Reset()
    DeleteAllInternal
    this.isReady = False
End Sub


'@section Paths
'===============================================================================
Public Property Get RootPath() As String
    RootPath = this.rootPath
End Property

Public Function CreateFilePath(ByVal fileName As String) As String
    EnsureReady
    CreateFilePath = this.rootPath & SanitizeFileName(fileName)
End Function


'@section File operations
'===============================================================================
Public Sub DeleteFile(ByVal path As String)
    If LenB(path) = 0 Then Exit Sub
    On Error Resume Next
        Kill path
    On Error GoTo 0
End Sub

Public Sub DeleteAll()
    EnsureReady
    DeleteAllInternal
    EnsureDirectoryExists this.rootPath
End Sub


'@section Internal helpers
'===============================================================================
Private Function ResolveRootPath(ByVal baseDirectory As String) As String
    Dim root As String

    If LenB(Trim$(baseDirectory)) = 0 Then
        root = Application.DefaultFilePath
        If LenB(root) = 0 Then root = Environ$("TEMP")
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

    ResolveRootPath = root & DEFAULT_FOLDER_NAME
End Function

Private Sub EnsureDirectoryExists(ByVal folderPath As String)
    On Error GoTo CreateFolder
        Dim attr As Long
        attr = GetAttr(folderPath)
        If (attr And vbDirectory) = 0 Then
            ThrowError ProjectError.ErrorUnexpectedState, _
                      "Temporary path exists but is not a directory: " & folderPath
        End If
        Exit Sub

CreateFolder:
    Err.Clear
    MkDir folderPath
    If Err.Number <> 0 Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Unable to create temporary folder: " & folderPath & " (" & Err.Description & ")"
    End If
End Sub

Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim result As String
    Dim index As Long

    If LenB(Trim$(fileName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Temporary file name cannot be empty"
    End If

    invalidChars = Array(":", "*", "?", "\", "/", "|", "<", ">", Chr$(34))
    result = fileName

    For index = LBound(invalidChars) To UBound(invalidChars)
        result = Replace$(result, invalidChars(index), "_")
    Next index

    SanitizeFileName = result
End Function

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

Private Sub Class_Terminate()
    On Error Resume Next
        DeleteAllInternal
    On Error GoTo 0
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistTempFileService_EnsureReady()
    EnsureReady
End Sub

Private Sub ILinelistTempFileService_Reset()
    Reset
End Sub

Private Property Get ILinelistTempFileService_RootPath() As String
    ILinelistTempFileService_RootPath = RootPath
End Property

Private Function ILinelistTempFileService_CreateFilePath(ByVal fileName As String) As String
    ILinelistTempFileService_CreateFilePath = CreateFilePath(fileName)
End Function

Private Sub ILinelistTempFileService_DeleteFile(ByVal path As String)
    DeleteFile path
End Sub

Private Sub ILinelistTempFileService_DeleteAll()
    DeleteAll
End Sub
