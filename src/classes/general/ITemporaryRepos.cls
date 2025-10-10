VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITemporaryRepos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides temporary folder and file utilities for the Linelist workflow"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Provides temporary folder and file utilities for the Linelist workflow")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@section Lifecycle
'===============================================================================
'@label:EnsureReady
'@sub-title Prepare the repository so that file operations can be performed.
Public Sub EnsureReady()
End Sub

'@label:Reset
'@sub-title Clear the repository contents and mark it as uninitialised.
Public Sub Reset()
End Sub

'@section Paths
'===============================================================================
'@label:RootPath
'@sub-title Provide the active repository root folder.
'@return String absolute path.
Public Property Get RootPath() As String
End Property

'@label:CreateFilePath
'@sub-title Generate a sanitised file path within the repository.
'@param fileName String requested file name.
'@return String fully-qualified path.
Public Function CreateFilePath(ByVal fileName As String) As String
End Function

'@section File operations
'===============================================================================
'@label:DeleteFile
'@sub-title Remove a single file from the repository if it exists.
'@param path String absolute file path.
Public Sub DeleteFile(ByVal path As String)
End Sub

'@label:DeleteAll
'@sub-title Remove every file created within the repository.
Public Sub DeleteAll()
End Sub
