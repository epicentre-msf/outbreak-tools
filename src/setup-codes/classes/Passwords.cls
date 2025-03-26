VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Passwords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage passwords in the setup file"

'@PredeclaredId
'@Folder("setup file)
'@ModuleDescription("Manage passwords in the setup file")

Option Explicit

'Exposed methods
Implements IPasswords

Private Type TPassWord
    sh As Worksheet
End Type

Private this As TPassWord

Private Const CLASSNAME As String = "Passwords"

Public Function Create(ByVal sh As Worksheet) As IPasswords
    CheckRequirements sh

    With New Passwords
        Set .Wksh = sh
        Set Create = .Self()
    End With
End Function

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get Self() As IPasswords
    Set Self = Me
End Property

Private Property Get Password() As String
    Dim sh As Worksheet
    Set sh = Wksh()
    Password = sh.Range("RNG_DebuggingPassword").Value
End Property

Private Property Get InDebugMode() As Boolean
    Dim sh As Worksheet
    Set sh = Wksh()
    InDebugMode = (sh.Range("RNG_DebugMode").Value = "yes")
End Property

'protect
Private Sub Protect(Optional ByVal sheetName As String = "_active", _
                    Optional ByVal strictProtect As Boolean = False, _
                    Optional ByVal deleteColumns As Boolean = False)

    'Strict protect restricts from deleting / inserting rows and formatting columns
    'and cells.

    Dim pwd As String
    Dim sh As Worksheet
    Dim wb As Workbook
    Set wb = Wksh.Parent

    If sheetName = "_active" Then
        Set sh = ActiveSheet
    Else
        On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
        On Error GoTo 0
    End If
    'Test if the sheet Exists
    If sh Is Nothing Then Exit Sub
    If Not InDebugMode() Then
        pwd = Password()
        sh.Protect Password:=pwd, DrawingObjects:=False, Contents:=True, Scenarios:=True, _
                   AllowInsertingRows:=(Not strictProtect), AllowSorting:=True, AllowFiltering:=True, _
                   AllowFormattingCells:=True, AllowDeletingRows:=(Not strictProtect), _
                   AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                   AllowDeletingColumns:=deleteColumns
    End If

End Sub

'Protect workbook
Private Sub ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
    Dim pwd As String
    pwd = Password()
    If Not InDebugMode() Then wb.Protect pwd, Structure:=protectSheetName
End Sub


'UnProptect the workbook
Private Sub UnProtectWkb(ByVal wb As Workbook)
    Dim pwd As String
    pwd = Password()
    On Error Resume Next
    wb.UnProtect pwd
    On Error GoTo 0
End Sub


Private Sub UnProtect(Optional ByVal sheetName As String = "_active")

    Dim pwd As String
    Dim sh As Worksheet
    Dim wb As Workbook
    Set wb = Wksh.Parent
    If sheetName = "_active" Then
        Set sh = ActiveSheet
    Else
        On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
        On Error GoTo 0
    End If
    'Test if the sheet Exists
    If sh Is Nothing Then Exit Sub
    pwd = Password()
    If sh.ProtectContents Then sh.UnProtect pwd
End Sub

'Requirements for instanciation
Private Sub CheckRequirements(ByVal sh As Worksheet)

    If sh Is Nothing Then _
     ThrowError "Unable to instanciate Password class because there is no sheet given"

    'Ranges in the linelist translation worksheet
    RangeExists rngName:="RNG_DebuggingPassword", sh:=sh
    RangeExists rngName:="RNG_DebugMode", sh:=sh
End Sub

'Check if listobject exists in a worksheet
Private Sub RangeExists(ByVal rngName As String, ByVal sh As Worksheet)

    Dim rng As Range
    On Error Resume Next
    Set rng = sh.Range(rngName)
    On Error GoTo 0

    If rng Is Nothing Then ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name
End Sub

Private Sub LoExists(ByVal loname As String, sh As Worksheet)
    Dim Lo As ListObject
    On Error Resume Next
    Set Lo = sh.ListObjects(loname)
    On Error GoTo 0
    If Lo Is Nothing Then ThrowError "Table " & loname & " does not exists in Worksheet " & sh.Name
End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Property Get IPasswords_PassWord() As String
    IPasswords_PassWord = Password()
End Property

Private Sub IPasswords_UnProtect(Optional ByVal sheetName As String = "_active")
    UnProtect sheetName
End Sub

Private Sub IPasswords_Protect(Optional ByVal sheetName As String = "_active", _
                              Optional ByVal strictProtect As Boolean = False, _
                              Optional ByVal deleteColumns As Boolean = False)
    Protect sheetName, strictProtect, deleteColumns
End Sub

    'Protect and unprotect the workbook to avoid sheet name changes
Private Sub IPasswords_ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
    ProtectWkb wb, protectSheetName
End Sub

Private Sub IPasswords_UnProtectWkb(ByVal wb As Workbook)
    UnProtectWkb wb
End Sub
