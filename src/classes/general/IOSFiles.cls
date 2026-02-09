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
'@Folder("General")
'@ModuleDescription("Interface of OSFiles")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IOSFiles
'@description
'Defines the public contract for the OSFiles class, exposing
'cross-platform file and folder selection dialogs (Windows FileDialog
'and macOS AppleScript), iteration over multi-selections, and
'validation helpers.


'@section Elements
'===============================================================================
'@description File and folder path accessors.

'@jump:file
'@prop-title Current loaded file path
'@return String. The first selected file path.
Public Property Get File() As String: End Property

'@jump:files
'@prop-title All loaded file paths from a multi-selection
'@return Variant. Zero-based string array of file paths.
Public Property Get Files() As Variant: End Property

'@jump:folder
'@prop-title Current loaded folder path
'@return String. The first selected folder path.
Public Property Get Folder() As String: End Property

'@jump:folders
'@prop-title All loaded folder paths from a multi-selection
'@return Variant. Zero-based string array of folder paths.
Public Property Get Folders() As Variant: End Property


'@section Selecting
'===============================================================================
'@description File and folder picker operations.

'@jump:load-file
'@sub-title Show a single-file picker dialog
'@param filters String. File type filters (e.g. "*.xlsx", "*.xlsb, *.xlsx").
'@param customMacFilter String. Optional UTI override for macOS. Defaults to vbNullString.
Public Sub LoadFile(ByVal filters As String, _
                    Optional ByVal customMacFilter As String = vbNullString)
End Sub

'@jump:load-files
'@sub-title Show a multi-file picker dialog
'@param filters String. File type filters.
'@param customMacFilter String. Optional UTI override for macOS. Defaults to vbNullString.
Public Sub LoadFiles(ByVal filters As String, _
                     Optional ByVal customMacFilter As String = vbNullString)
End Sub

'@jump:load-folder
'@sub-title Show a single-folder picker dialog
Public Sub LoadFolder(): End Sub

'@jump:load-folders
'@sub-title Show a multi-folder picker dialog
Public Sub LoadFolders(): End Sub

'@jump:reset-files-iterator
'@sub-title Reset the iteration cursor over selected files
Public Sub ResetFilesIterator(): End Sub

'@jump:has-next-file
'@sub-title Check if another file is available in the iterator
'@return Boolean. True when the cursor has not reached the end.
Public Function HasNextFile() As Boolean: End Function

'@jump:next-file
'@sub-title Retrieve the next file path and advance the cursor
'@return String. The next file path.
Public Function NextFile() As String: End Function

'@jump:reset-folders-iterator
'@sub-title Reset the iteration cursor over selected folders
Public Sub ResetFoldersIterator(): End Sub

'@jump:has-next-folder
'@sub-title Check if another folder is available in the iterator
'@return Boolean. True when the cursor has not reached the end.
Public Function HasNextFolder() As Boolean: End Function

'@jump:next-folder
'@sub-title Retrieve the next folder path and advance the cursor
'@return String. The next folder path.
Public Function NextFolder() As String: End Function


'@section Validation
'===============================================================================
'@description Selection validity checks.

'@jump:has-valid-file
'@sub-title Whether the loaded file path is valid
'@return Boolean. True when a non-empty file path exists.
Public Function HasValidFile() As Boolean: End Function

'@jump:has-valid-files
'@sub-title Whether the loaded files array contains entries
'@return Boolean. True when at least one file was selected.
Public Function HasValidFiles() As Boolean: End Function

'@jump:has-valid-folder
'@sub-title Whether the loaded folder path is valid
'@return Boolean. True when a non-empty folder path exists.
Public Function HasValidFolder() As Boolean: End Function

'@jump:has-valid-folders
'@sub-title Whether the loaded folders array contains entries
'@return Boolean. True when at least one folder was selected.
Public Function HasValidFolders() As Boolean: End Function
