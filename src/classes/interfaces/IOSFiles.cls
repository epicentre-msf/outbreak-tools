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

Option Explicit

'The main advantage of OSFiles is that it could work on both Macos and Windows
'The allowed filters are the following: "*.xls", "*.xlsx", "*.xlsb", "*.xlsb, *.xlsx"
'If you want to use more filters, please remember to add them on Mac in the
'SelectFileOnMac Sub of OSFiles, new filters like *.csv may be
'accessible on windows, but if the filter is not in the previous list, it will 
'default to xlsx files on Mac.

'Ideas based on https://macexcel.com/examples/filesandfolders/selectfiles/

'Return back the loaded File/Folder
Public Property Get File() As String: End Property
Public Property Get Folder() As String: End Property
'Load a file by adding filtes
Public Sub LoadFile(ByVal filters As String): End Sub
'Load a folder
Public Sub LoadFolder(): End Sub
'Test if the loaded file is valid (shoud be after loading)
Public Function HasValidFile() As Boolean: End Function
'Test if the loaded folder is valid
Public Function HasValidFolder() As Boolean: End Function
