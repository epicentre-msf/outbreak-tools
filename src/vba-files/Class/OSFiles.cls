VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OSFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("OS")
'@ModuleDescription("I/O from operating system to the linelist or designer")
'@IgnoreModule

Option Explicit
Implements IOSFiles

Private Type TOSFiles
    currentOS As String
    currentLoadedFile As String
    currentLoadedFolder As String
End Type

'Exposed methods
Private this As TOSFiles

Public Function Create() As IOSFiles

    'Will run on Mac or windows operating system
    Dim currentOS As String
    currentOS = IIf(Application.OperatingSystem Like "*Mac*", "Mac", "Windows")
  
    With New OSFiles
        .OS = currentOS
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IOSFiles
    Set Self = Me
End Property

Public Property Get OS() As String
    OS = this.currentOS
End Property

Public Property Let OS(ByVal currentOS As String)
    this.currentOS = currentOS
End Property

Private Property Get File() As String
    File = this.currentLoadedFile
End Property

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

'Folder on Mac
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

'Folder on Windows
Private Function SelectFolderOnWindows() As String

    Dim fDialog As Office.FileDialog

    SelectFolderOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your directory"          'MSG_ChooseDir
        .filters.Clear

        If .Show = -1 Then
            SelectFolderOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'Functions to select a file

'File on Mac
Private Function SelectFileOnMac(ByVal filters As String) As String

    Dim sMacFilter As String
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String


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

'File on Windows
Private Function SelectFileOnWindows(ByVal filters As String) As String

    Dim fDialog As Office.FileDialog

    SelectFileOnWindows = vbNullString

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Chose your file"               'MSG_ChooseFile
        .filters.Clear
        .filters.Add "Excel Sheet", filters      '"*.xlsx" ', *.xlsm, *.xlsb,  *.xls" 'MSG_ExcelFile

        If .Show = True Then
            SelectFileOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'Load file and load folders
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

Private Sub LoadFile(ByVal filters As String)
    If OS() = "Windows" Then
        'We are on windows DOS
        File = SelectFileOnWindows(filters)
    ElseIf OS() = "Mac" Then
        'We are on Mac, need to test the version of excel running
        If val(Application.Version) > 14 Then
            File = SelectFileOnMac(filters)
        End If
    End If
End Sub

Public Function HasValidFile() As Boolean
    HasValidFile = (File() <> vbNullString)
End Function

Public Function HasValidFolder() As Boolean
    HasValidFolder = (Folder() <> vbNullString)
End Function

'Exposed Methods

Public Sub IOSFiles_LoadFile(ByVal filters As String)
    LoadFile filters
End Sub

Public Sub IOSFiles_LoadFolder()                 'Load a folder and store the path
    LoadFolder
End Sub

Public Function IOSFiles_HasValidFile() As Boolean
    IOSFiles_HasValidFile = HasValidFile()
End Function

Public Function IOSFiles_HasValidFolder() As Boolean
    IOSFiles_HasValidFolder = HasValidFolder()
End Function

Public Property Get IOSFiles_File() As String
    IOSFiles_File = File()
End Property

Public Property Get IOSFiles_Folder() As String
    IOSFiles_Folder = Folder()
End Property


