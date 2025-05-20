VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OSFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Importing files from the operating system"


'@Folder("OS")
'@ModuleDescription("Importing files from the operating system")
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation


Option Explicit
Implements IOSFiles

Private Type TOSFiles
    currentOS As String
    currentLoadedFile As String
    currentLoadedFolder As String
End Type

'Exposed methods
Private this As TOSFiles

'@label:Create
'@sub-title Create an OSFile object
'@details
'OSFile is a class to import files from different operation systems 
'(Mac or Windows). 
'The main advantage of OSFiles is that it could work on both Macos and Windows
'The allowed filters are the following: "*.xls", "*.xlsx", "*.xlsb", "*.xlsb, *.xlsx"
'If you want to use more filters, please remember to update the customMacFilter 
'argument of `LoadFile` or `LoadFolder`. New filters like .csv may be accessible on 
'windows, but if the filter is not in the previous list it will default to xlsx files
'on Mac.
' Here is the list of filters you can use on Mac:
' - xls : com.microsoft.Excel.xls
' - xlsx : org.openxmlformats.spreadsheetml.sheet
' - xlsm : org.openxmlformats.spreadsheetml.sheet.macroenabled
' - xlsb : com.microsoft.Excel.sheet.binary.macroenabled
' - csv : public.comma-separated-values-text
' - doc : com.microsoft.word.doc
' - docx : org.openxmlformats.wordprocessingml.document
' - docm : org.openxmlformats.wordprocessingml.document.macroenabled
' - ppt : com.microsoft.powerpoint.ppt
' - pptx : org.openxmlformats.presentationml.presentation
' - pptm : org.openxmlformats.presentationml.presentation.macroenabled
' - txt : public.plain-text
' - pdf : com.adobe.pdf
' - jpg : public.jpeg
' - png : public.png
' - QIF : com.apple.traditional-mac-plain-text
' - htm : public.html
'
'Based on https://macexcel.com/examples/filesandfolders/selectfiles/.
'
'@return OSFile Object
'@export
Public Function Create() As IOSFiles

    'Will run on Mac or windows operating system
    Dim currentOS As String
    currentOS = IIf(Application.OperatingSystem Like "*Mac*", "Mac", "Windows")
  
    With New OSFiles
        .OS = currentOS
        Set Create = .Self
    End With
End Function

'@label:elements
'@section OSFiles elements

Public Property Get Self() As IOSFiles
    Set Self = Me
End Property

Public Property Get OS() As String
    OS = this.currentOS
End Property

Public Property Let OS(ByVal currentOS As String)
    this.currentOS = currentOS
End Property

'@label:File
'@prop-title Get the current loaded file path
'@return A string, the file path
'@export
Private Property Get File() As String
    File = this.currentLoadedFile
End Property

'@label:Folder
'@prop-title Get the current loaded folder path
'@return A String, the folder path
'@export
Private Property Get Folder() As String
    Folder = this.currentLoadedFolder
End Property

Private Property Let File(ByVal ff As String)
    this.currentLoadedFile = ff
End Property

Private Property Let Folder(ByVal ff As String)
    this.currentLoadedFolder = ff
End Property

'Functions to select a folder


'@prop-title Select a folder on MacOS
Private Function SelectFolderOnMac() As String
    Dim FolderPath As String
    Dim RootFolder As String
    Dim Scriptstr As String

    On Error Resume Next

    'Enter the Start Folder, Desktop in this example,
    'Use the second line to enter your own path
    RootFolder = MacScript("return POSIX path of (path to documents folder) as string")

    'Make the path Colon seperated for using in MacScript
    RootFolder = MacScript("return POSIX file (""" & RootFolder & """) as string")
    'Make the Script string
    Scriptstr = "return POSIX path of (choose folder with prompt ""Select the folder""" & _
                " default location alias """ & RootFolder & """) as string"

    'Run the Script
    FolderPath = MacScript(Scriptstr)

    If CInt(Split(Application.Version, ".")(0)) >= 15 Then 'excel 2016 support
        FolderPath = Replace(FolderPath, ":", "/")
        FolderPath = Replace(FolderPath, "Macintosh HD", "", Count:=1)
    End If

    On Error GoTo 0

    If FolderPath <> "" Then
        'Remove the last ":" or "/"
        SelectFolderOnMac = Mid(FolderPath, 1, (Len(FolderPath) - 1))
    End If
End Function

'@prop-title Select a folder on windows OS
'@return A String, the path to the folder on windows
Private Function SelectFolderOnWindows() As String

    Dim fDialog As Office.FileDialog

    SelectFolderOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .title = "Chose your directory"          'MSG_ChooseDir
        .filters.Clear

        If .Show = -1 Then
            SelectFolderOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'@prop-title Select a file on MacOS
'@param filters A String, filters to use during selection. See Create
'@param customMacFilter Optional, a string. filter to use in case you want a specific filter on MacOS. See Create
'@return A String, the path to the selected file on MacOS
Private Function SelectFileOnMac(ByVal filters As String, _ 
                                 Optional ByVal customMacFilter As String = vbNullString) As String

    Dim sMacFilter As String
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String


    If customMacFilter = vbNullString Then 
        Select Case filters
        Case "*.xls"
            sMacFilter = " {""com.microsoft.Excel.xls""} "
        Case "*.xlsx"
            sMacFilter = " {""org.openxmlformats.spreadsheetml.sheet""} "
        Case "*.xlsb"
            sMacFilter = " {""com.microsoft.Excel.sheet.binary.macroenabled""} "
        Case "*.xlsb, *.xlsx"
            sMacFilter = " {""org.openxmlformats.spreadsheetml.sheet"",""com.microsoft.Excel.sheet.binary.macroenabled""} "
        Case Else
            sMacFilter = " {""org.openxmlformats.spreadsheetml.sheet""} "
        End Select
    Else
        sMacFilter = Chr(34) & " {" & Chr(34) & Chr(34) &  customMacFilter & Chr(34) & Chr(34) & "} "
    End If

    SelectFileOnMac = vbNullString
    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")
    MyScript = _
             "set applescript's text item delimiters to "","" " & vbNewLine & _
                                                           "set theFiles to (choose file of type " & _
                                                           sMacFilter & _
                                                           "with prompt ""Please select a file or files"" default location alias """ & _
                                                           MyPath & """ multiple selections allowed false) as string" & vbNewLine & _
                                                                                                           "set applescript's text item delimiters to """" " & vbNewLine & _
                                                                                                           "return theFiles"
    MyFiles = MacScript(MyScript)

    If CInt(Split(Application.Version, ".")(0)) >= 15 Then 'excel 2016 support
        MyFiles = Replace(MyFiles, ":", "/")
        MyFiles = Replace(MyFiles, "Macintosh HD", "", Count:=1)
    End If

    On Error GoTo 0

    SelectFileOnMac = MyFiles
End Function

'@prop-title Select a file on Windows OS
'@param filters A String, filters to use during selection. See Create
'@return A String, the path to the file on Windows
Private Function SelectFileOnWindows(ByVal filters As String) As String

    Dim fDialog As Office.FileDialog

    SelectFileOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .title = "Chose your file"               'MSG_ChooseFile
        .filters.Clear
        .filters.Add "Excel Sheet", filters      '"*.xlsx" ', *.xlsm, *.xlsb,  *.xls" 'MSG_ExcelFile

        If .Show = True Then
            SelectFileOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'@label:selecting
'@section selecting

'@label:LoadFolder
'@sub-title Select a folder
'@return A String, The file path to the folder
'@export
Private Sub LoadFolder()
    If OS() = "Windows" Then
        'We are on windows DOS
        Folder = SelectFolderOnWindows()
    ElseIf OS() = "Mac" Then
        'We are on Mac, need to test the version of excel running
        If val(Application.Version) > 14 Then
            Folder = SelectFolderOnMac()
        End If
    End If
End Sub


'@label:LoadFile
'@sub-title Select a file
'@details
'If the user aborts the selection process or something go wrong during the 
'interaction, the internal object will have an empty file. There is a property
'validate the selection process and make sur the file exists.
'
'@param filters A String, filters to use during selection. See Create
'@param customMacFilter Optional, a string. filter to use in case you want a specific filter on MacOS. See Create
'@export
Private Sub LoadFile(ByVal filters As String, _ 
                     Optional ByVal customMacFilter As String = vbNullString)
    If OS() = "Windows" Then
        'We are on windows DOS
        File = SelectFileOnWindows(filters)
    ElseIf OS() = "Mac" Then
        'We are on Mac, need to test the version of excel running
        If val(Application.Version) > 14 Then
            File = SelectFileOnMac(filters, customMacFilter)
        End If
    End If
End Sub

'@label:validation
'@section Validation

'@label:HasValidFile
'@prop-title Test if the internal loaded file is valid
'@return A Boolean
'@export
Public Function HasValidFile() As Boolean
    HasValidFile = (File() <> vbNullString)
End Function

'@label:HasValidFolder
'@prop-title Test if the internal loaded folder is valid
'@return A Boolean
'@export
Public Function HasValidFolder() As Boolean
    HasValidFolder = (Folder() <> vbNullString)
End Function


'@jump:Create
'@section Interface

'@jump:elements
'IOSFile elements

'@jump:File
Public Property Get IOSFiles_File() As String
    IOSFiles_File = File()
End Property

'@jump:Folder
Public Property Get IOSFiles_Folder() As String
    IOSFiles_Folder = Folder()
End Property


'@jump:selecting
'Selecting------------------------------------------------------

'@jump:LoadFolder
Public Sub IOSFiles_LoadFolder()
    LoadFolder
End Sub

'@jump:LoadFile
Public Sub IOSFiles_LoadFile(ByVal filters As String, _ 
                             Optional ByVal customMacFilter As String = vbNullString)
    LoadFile filters:=filters, customMacFilter:=customMacFilter
End Sub

'@jump:validation
'Validate---------------------------------------------------------

'@jump:HasValidFile
Public Function IOSFiles_HasValidFile() As Boolean
    IOSFiles_HasValidFile = HasValidFile()
End Function

'@jump:HasValidFolder
Public Function IOSFiles_HasValidFolder() As Boolean
    IOSFiles_HasValidFolder = HasValidFolder()
End Function


 
