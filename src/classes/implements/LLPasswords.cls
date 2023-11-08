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
'@IgnoreModule UseMeaningfulName, HungarianNotation

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

Private Property Get Value(ByVal tagName As String) As String
    Dim sh As Worksheet
    Dim retVal As String

    Set sh = Wksh()
    
    Select Case tagName
    Case "debuggingpassword"
        retVal = sh.Range("RNG_DebuggingPassword").Value
    Case "publickey"
        retVal = sh.Range("RNG_PublicKey").Value
    Case "privatekey"
        retVal = sh.Range("RNG_PrivateKey").Value
    Case "version"
        retVal = sh.Range("RNG_Version").Value
    End Select

    Value = retVal
End Property

Private Property Get TableRange() As Range

    Dim sh As Worksheet

    Set sh = Wksh()
    Set TableRange = sh.ListObjects(TABKEYS).DataBodyRange

End Property

'Generate a new key for the file to be saved
Private Sub GenerateKey(ByVal trads As ITranslation)
    Dim nbRows As Long
    Dim randNum As Integer
    Dim LoKeys As listObject 'ListObject of keys
    Dim LoRng As Range
    Dim sh As Worksheet

    Set sh = Wksh()
    Set LoKeys = sh.ListObjects(TABKEYS)

    'Don't try to generate new keys if the databodyrange does not exists
    If LoKeys.DataBodyRange Is Nothing Then Exit Sub
    Set LoRng = LoKeys.DataBodyRange
    
    'Number of Rows of the databodyrange of the listobject
    nbRows = LoRng.Rows.Count

    'Generate a random number
    randNum = Int(nbRows * Rnd())
    
    'Private and Public key values
    sh.Range("RNG_PublicKey").Value = LoRng.Cells(randNum, 1).Value
    sh.Range("RNG_PrivateKey").Value = LoRng.Cells(randNum, 2).Value

    'Display the Private key to the user
    DisplayPrivateKey trads
End Sub

'Display the private key
Private Sub DisplayPrivateKey(ByVal trads As ITranslation)
    Dim sh As Worksheet
    Dim privKey As String
    Set sh = Wksh()
    privKey = sh.Range("RNG_PrivateKey").Value
    MsgBox trads.TranslatedValue("MSG_Password") & privKey
End Sub

'protect
Private Sub Protect(ByVal sheetObject As Variant, Optional ByVal allowShapes As Boolean = False, _ 
                    Optional ByVal allowDeletingRows As Boolean = True)

    Dim pwd As String
    Dim sh As Worksheet
    Dim wb As Workbook

    Set wb = Wksh.Parent

    'If the sheetObject is a worksheet, unProtect it, if it is a string

    If (TypeName(sheetObject) = "String") Then
        If sheetObject = "_active" Then
            Set sh = ActiveSheet
        Else
            On Error Resume Next
            Set sh = wb.Worksheets(sheetObject)
            On Error GoTo 0
        End If

    ElseIf (TypeName(sheetObject) = "Worksheet") Then
        Set sh = sheetObject
    End If

    'Test if the sheet Exists
    If sh Is Nothing Then Exit Sub
    If Not InDebugMode() Then

        pwd = Password()
        
        'Contents: Protect graphs data, and Locked cells in one sheet
        'Scenarios: Protect scenarios
        'DrawingObjects: You can change graphs dimensions, move them and / or add new shapes
        'You can remove alowing scenario and unable selection of lock cells
        
        sh.Protect Password:=pwd, _ 
                   DrawingObjects:=(Not allowShapes), _ 
                   Contents:=True, _ 
                   Scenarios:=True, _
                   AllowInsertingRows:=True, _ 
                   AllowSorting:=False, _ 
                   AllowFiltering:=True, _
                   AllowFormattingCells:=True, _ 
                   AllowDeletingRows:=allowDeletingRows, _ 
                   AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True
    End If
End Sub

'Protect a worksheet that is not in current workbook, using the LLPassword
'class

'Protect workbook
Private Sub ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
    Dim pwd As String
    pwd = Password()
    wb.Protect pwd, Structure:=protectSheetName
End Sub


'UnProptect the workbook
Private Sub UnProtectWkb(ByVal wb As Workbook)
    Dim pwd As String
    pwd = Password()
    On Error Resume Next
    wb.UnProtect pwd
    On Error GoTo 0
End Sub

'UnProtect a workbook
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

'Enter debug mode in the linelist (when something crashes)
Private Sub EnterDebugMode()

End Sub

'Leave debug mode in the linelist (when closing the linelist, or after debugging)
Private Sub LeaveDebugMode()


End Sub


Private Sub Export(ByVal toWkb As Workbook)

    Dim sh As Worksheet
    Dim outsh As Worksheet
    Dim Lo As listObject
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
                  "RNG_LabPrivateKey", "RNG_DebugMode", "RNG_Version"

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

Private Sub Import(ByVal pass As ILLPasswords)

    Dim sh As Worksheet
    Dim impTabRng As Range
    Dim actTabRng As Range
    Dim Lo As listObject

    Set sh = Wksh()
    
    'Set values of ranges
    sh.Range("RNG_PublicKey").Value = pass.Value("publickey")
    sh.Range("RNG_PrivateKey").Value = pass.Value("privatekey")
    sh.Range("RNG_Version").Value = pass.Value("version")
    
    Set impTabRng = pass.TableRange()
    '
    If impTabRng.Columns.Count <> 2 Then _
    ThrowError "Error when importing, import keys table doesn't have 2 columns"
    
    Set Lo = sh.ListObjects(TABKEYS)
    Set actTabRng = Lo.DataBodyRange

    'Resize the actual Range
    actTabRng.ClearContents
    Set actTabRng = sh.Range(actTabRng.Cells(1, 1), actTabRng.Cells(impTabRng.Rows.Count, 2))
    actTabRng.Value = impTabRng.Value

    'Resize the keys ListObject
    Set actTabRng = sh.Range(actTabRng.Cells(0, 1), actTabRng.Cells(actTabRng.Rows.Count, 2))
    Lo.Resize actTabRng
End Sub

'Requirements for instanciation
Private Sub CheckRequirements(ByVal sh As Worksheet)

    If sh Is Nothing Then ThrowError "Unable to instanciate, no password sheet"

    'Required listObjects fo translations
    LoExists TABKEYS, sh
    
    If sh.ListObjects(TABKEYS).DataBodyRange Is Nothing Then ThrowError "Unable to instanciate password, the keys table is empty"

    'Ranges in the linelist translation worksheet
    RangeExists rngName:="RNG_DebuggingPassword", sh:=sh
    RangeExists rngName:="RNG_PublicKey", sh:=sh
    RangeExists rngName:="RNG_LabPublicKey", sh:=sh
    RangeExists rngName:="RNG_PrivateKey", sh:=sh
    RangeExists rngName:="RNG_LabPrivateKey", sh:=sh
    RangeExists rngName:="RNG_DebugMode", sh:=sh
    RangeExists rngName:="RNG_Version", sh:=sh

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
    Dim Lo As listObject
    On Error Resume Next
    Set Lo = sh.ListObjects(loName)
    On Error GoTo 0
    If Lo Is Nothing Then ThrowError "Table " & loName & " does not exists in Worksheet " & sh.Name
End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub


Private Property Get ILLPasswords_PassWord() As String
    ILLPasswords_PassWord = Password()
End Property

Private Sub ILLPasswords_UnProtect(Optional ByVal sheetName As String = "_active")
    UnProtect sheetName
End Sub

Private Sub ILLPasswords_Protect(Optional ByVal sheetObject As Variant, _
                                 Optional ByVal allowShapes As Boolean = False, _ 
                                 Optional ByVal allowDeletingRows As Boolean = True)
    Protect sheetObject, allowShapes, allowDeletingRows
End Sub

Private Sub ILLPasswords_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub

    'Protect and unprotect the workbook to avoid sheet name changes
Private Sub ILLPasswords_ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
    ProtectWkb wb, protectSheetName
End Sub

Private Sub ILLPasswords_UnProtectWkb(ByVal wb As Workbook)
    UnProtectWkb wb
End Sub

Private Property Get ILLPasswords_Value(ByVal tagName As String) As String
    ILLPasswords_Value = Value(tagName)
End Property

Private Sub ILLPasswords_DisplayPrivateKey(ByVal trads As ITranslation)
    DisplayPrivateKey trads
End Sub

Private Sub ILLPasswords_GenerateKey(ByVal trads As ITranslation)
    GenerateKey trads
End Sub

Private Sub ILLPasswords_EnterDebugMode()
    EnterDebugMode
End Sub

Private Sub ILLPasswords_LeaveDebugMode()
    LeaveDebugMode
End Sub

Private Sub ILLPasswords_Import(ByVal pass As ILLPasswords)
    Import pass
End Sub

Private Property Get ILLPasswords_TableRange() As Range
    Set ILLPasswords_TableRange = TableRange()
End Property
