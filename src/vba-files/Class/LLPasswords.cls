VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist Passwords")
'@ModuleDescription("Manage passwords in the linelist")
'@IgnoreModule

Option Explicit

'Exposed methods
Implements ILLPasswords

Private Type TPassWord
    sh As Worksheet
End Type

Private this As TPassWord

Private Const CLASSNAME As String = "LLPasswords"
Private Const TABKEYS   As String = "T_keys"

Public Function Create(ByVal sh As Worksheet) As ILLPasswords
    CheckRequirements sh

    With New LLPasswords
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

Public Property Get Self() As ILLPasswords
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
Private Sub Protect(Optional ByVal sheetName As String = "_active", Optional Byval allowShapes As Boolean = False)

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
        sh.protect Password:=pwd, DrawingObjects:=(Not allowShapes), Contents:=True, Scenarios:=allowShapes, _
                   AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
                   AllowFormattingCells:=True, AllowDeletingRows:=True
                   
        If Not allowShapes Then sh.EnableSelection = xlUnlockedCells
    End If

End Sub

Private Sub UnProtect(Optional ByVal sheetName As String = "_active")

    Dim pwd As String
    Dim sh As Worksheet
    Dim wb As Workbook
    Set wb = Wksh.Parent
    If sheetName = "_Active" Then
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

    If sh Is Nothing Then ThrowError "Unable to instanciate, no password sheet"

    'Required listObjects fo translations
    LoExists TABKEYS, sh

    'Ranges in the linelist translation worksheet
    RangeExists rngName:="RNG_DebuggingPassword", sh:=sh
    RangeExists rngName:="RNG_PublicKey", sh:=sh
    RangeExists rngName:="RNG_LabPublicKey", sh:=sh
    RangeExists rngName:="RNG_PrivateKey", sh:=sh
    RangeExists rngName:="RNG_LabPrivateKey", sh:=sh
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

Private Sub LoExists(ByVal loName As String, sh As Worksheet)
    Dim Lo As ListObject
    On Error Resume Next
    Set Lo = sh.ListObjects(loName)
    On Error GoTo 0
    If Lo Is Nothing Then ThrowError "Table " & loName & " does not exists in Worksheet " & sh.Name
End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

Private Sub Export(ByVal toWkb As Workbook)

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim Lo As ListObject
    Dim expAddress As String
    Dim expRange As Range
    Dim rngTable As BetterArray
    Dim counter As Long
    Dim rngName As String

    'Test if a sheet named as the geo sheet already exists
    Set sh = Wksh()

    On Error Resume Next
    Set outsh = toWkb.Worksheets(sh.Name)
    On Error GoTo 0

    If Not outsh Is Nothing Then ThrowError "Unable to export, already a sheet named " & sh.Name & " in the workbook"

    'Now move every listObjects
    toWkb.Worksheets.Add.Name = sh.Name
    Set outsh = toWkb.Worksheets(sh.Name)

    'Export the listObjects
    For Each Lo In sh.ListObjects
        expAddress = Lo.Range.Address
        Set expRange = outsh.Range(expAddress)
        expRange.Value = Lo.Range.Value
        'Add the listObject
        With outsh
            .ListObjects.Add(xlSrcRange, expRange, , xlYes).Name = Lo.Name
            .ListObjects(Lo.Name).TableStyle = Lo.TableStyle
        End With
    Next

    'Export all the ranges
    Set rngTable = New BetterArray
    rngTable.Push "RNG_DebuggingPassword", "RNG_PublicKey", "RNG_LabPublicKey", "RNG_PrivateKey", _
                  "RNG_LabPrivateKey", "RNG_DebugMode"

    For counter = rngTable.LowerBound To rngTable.UpperBound
        rngName = rngTable.Item(counter)
        expAddress = sh.Range(rngName).Address
        Set expRange = outsh.Range(expAddress)
        expRange.Name = rngName
        expRange.Value = sh.Range(rngName).Value
        If sh.Range(rngName).formula <> vbNullString Then
            On Error Resume Next
            expRange.formula = sh.Range(rngName).formula
            On Error GoTo 0
        End If
    Next

    outsh.Visible = xlSheetVeryHidden
End Sub

Public Property Get ILLPasswords_PassWord() As String
    ILLPasswords_PassWord = Password()
End Property

Public Sub ILLPasswords_UnProtect(Optional ByVal sheetName As String = "_active")
    UnProtect sheetName
End Sub

Public Sub ILLPasswords_Protect(Optional ByVal sheetName As String = "_active", Optional Byval allowShapes As Boolean = False)
    protect sheetName, allowShapes
End Sub

Public Sub ILLPasswords_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub


