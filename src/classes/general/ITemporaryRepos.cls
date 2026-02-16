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
'@Folder("General")
'@ModuleDescription("Provides temporary folder and file utilities for the Linelist workflow")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@class ITemporaryRepos
'@description
'Defines the public contract for the TemporaryRepos class, exposing
'temporary folder lifecycle, path generation, and file cleanup
'operations used during linelist preparation.


'@section Lifecycle
'===============================================================================
'@description Readiness and cleanup.

'@jump:ensure-ready
'@sub-title Prepare the repository so that file operations can be performed
Public Sub EnsureReady()
End Sub

'@jump:reset
'@sub-title Clear the repository contents and mark it as uninitialised
Public Sub Reset()
End Sub


'@section Paths
'===============================================================================
'@description Path accessors and generation.

'@jump:root-path
'@prop-title Active repository root folder
'@return String. Fully-qualified absolute path.
Public Property Get RootPath() As String
End Property

'@jump:create-file-path
'@sub-title Generate a sanitised file path within the repository
'@param fileName String. Requested file name.
'@return String. Fully-qualified path.
Public Function CreateFilePath(ByVal fileName As String) As String
End Function


'@section File Operations
'===============================================================================
'@description Deletion utilities.

'@jump:delete-file
'@sub-title Remove a single file from the repository if it exists
'@param path String. Absolute file path.
Public Sub DeleteFile(ByVal path As String)
End Sub

'@jump:delete-all
'@sub-title Remove every file created within the repository
Public Sub DeleteAll()
End Sub
