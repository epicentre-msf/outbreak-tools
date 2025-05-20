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
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'The main advantage of OSFiles is that it could work on both Macos and Windows
'The allowed filters are the following: "*.xls", "*.xlsx", "*.xlsb", "*.xlsb, *.xlsx"
'If you want to use more filters, please remember to add them on Mac in the
'SelectFileOnMac Sub of OSFiles, new filters like *.csv may be
'accessible on windows, but if the filter is not in the previous list, it will 
'default to xlsx files on Mac.

'Ideas based on https://macexcel.com/examples/filesandfolders/selectfiles/

'@section Elements

'@jump:File
'Return back the loaded File/Folder
Public Property Get File() As String: End Property

'@jump:Folder
'Return back the loaded folder
Public Property Get Folder() As String: End Property

'@section Selecting

'@jump:LoadFile
'Load a file by adding filtes
Public Sub LoadFile(ByVal filters As String, _ 
                    Optional ByVal customMacFilter As String = vbNullString)
End Sub

'@jump:LoadFolder
'Load a folder
Public Sub LoadFolder(): End Sub

'@section Validation

'@jump:HasValidFile
'Test if the loaded file is valid (shoud be after loading)
Public Function HasValidFile() As Boolean: End Function

'@jump:HasValidFolder
'Test if the loaded folder is valid
Public Function HasValidFolder() As Boolean: End Function
