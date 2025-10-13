VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOSFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of OSFiles"

'@Interface
'@Folder("GeneralInterface")
'@ModuleDescription("Interface of OSFiles")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'The main advantage of OSFiles is that it could work on both Macos and Windows
'The allowed filters are the following: "*.xls", "*.xlsx", "*.xlsb", "*.xlsb, *.xlsx"
'If you want to use more filters, please remember to add them on Mac in the
'SelectFileOnMac Sub of OSFiles, new filters like *.csv may be
'accessible on windows, but if the filter is not in the previous list, it will 
'default to xlsx files on Mac.

'Ideas based on https://macexcel.com/examples/filesandfolders/selectfiles/

'@section Elements
'===============================================================================

'@jump:File
'Return back the loaded File/Folder
Public Property Get File() As String: End Property

'@jump:Files
'Return back all loaded files when multi-selection is used
Public Property Get Files() As Variant: End Property

'@jump:Folder
'Return back the loaded folder
Public Property Get Folder() As String: End Property

'@jump:Folders
'Return back all loaded folders when multi-selection is used
Public Property Get Folders() As Variant: End Property

'@section Selecting
'===============================================================================


'@jump:LoadFile
'Load a file by adding filtes
Public Sub LoadFile(ByVal filters As String, _ 
                    Optional ByVal customMacFilter As String = vbNullString)
End Sub

'@jump:LoadFiles
'Load multiple files when the platform supports it
Public Sub LoadFiles(ByVal filters As String, _ 
                     Optional ByVal customMacFilter As String = vbNullString)
End Sub

'@jump:LoadFolder
'Load a folder
Public Sub LoadFolder(): End Sub

'@jump:LoadFolders
'Load multiple folders when the platform supports it
Public Sub LoadFolders(): End Sub

'@jump:ResetFilesIterator
'Reset the iteration cursor over selected files
Public Sub ResetFilesIterator(): End Sub

'@jump:HasNextFile
'Check if there is another file available in the iterator
Public Function HasNextFile() As Boolean: End Function

'@jump:NextFile
'Retrieve the next file from the iterator
Public Function NextFile() As String: End Function

'@jump:ResetFoldersIterator
'Reset the iteration cursor over selected folders
Public Sub ResetFoldersIterator(): End Sub

'@jump:HasNextFolder
'Check if there is another folder available in the iterator
Public Function HasNextFolder() As Boolean: End Function

'@jump:NextFolder
'Retrieve the next folder from the iterator
Public Function NextFolder() As String: End Function

'@section Validation
'===============================================================================

'@jump:HasValidFile
'Test if the loaded file is valid (shoud be after loading)
Public Function HasValidFile() As Boolean: End Function

'@jump:HasValidFiles
'Test if the loaded files array is valid
Public Function HasValidFiles() As Boolean: End Function

'@jump:HasValidFolder
'Test if the loaded folder is valid
Public Function HasValidFolder() As Boolean: End Function

'@jump:HasValidFolders
'Test if the loaded folders array is valid
Public Function HasValidFolders() As Boolean: End Function
