VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Passwords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Lock/unlock workbooks and worksheets"

'@Folder("General")
'@ModuleDescription("Lock/unlock workbooks and worksheets")
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
Implements IPasswords

Private Type TPassWord
    sh As Worksheet
    hasChecks As Boolean
End Type

Private this As TPassWord

Private Const CLASSNAME As String = "Passwords"
Private Const TABKEYS   As String = "T_keys"
Private Const TABPROTECTION As String = "T_ProtectedSheets"


'checking objects
Private checkCounter As Long
Private internalChecks As IChecking

'@label:Create
'@prop-title create a password object
'@details
'The password object is a worksheet with one listObject table in it called
'password. The listObject default name should be T_Keys. The table key
'has basically two column: one for private key and another one for public key.
'Private keys are used for locking the exports, public keys are the corresponding
'key use to find private keys when opening automatically the files. The password
'object also contains the debugging password which is the password used for
'locking/unlocking the worksheets and workbook where the password is.
'@return A IPasswords Object
'@export
Public Function Create(ByVal sh As Worksheet) As IPasswords
    CheckRequirements sh

    With New Passwords
        Set .Wksh = sh
        Set Create = .Self()
    End With
End Function

'@label:elements
'@section Password Elements
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get Self() As IPasswords
    Set Self = Me
End Property

'@prop-title Name of the password object
'@return A String
'@export
Private Property Get Name() As String
    Name = Wksh.Name
End Property

'@prop-title Get the current debugging password
'@return A string, the password in the DebuggingPassword range
'@return A String
Private Property Get Password() As String
    Dim sh As Worksheet
    Set sh = Wksh()
    Password = sh.Range("RNG_DebuggingPassword").Value
End Property


'@label:Value
'@prop-title Get specific value of ranges in the password worksheet
'
'@details
'In case the values are not found, a vbNullString is returned. Three tag names
'are actually availaible:
'- debuggingpassword to retrieve the debugging password
'- publickey to retrieve the public key
'- version to retrieve the version
'
'@tagName String, the tag of the value to look for.
'@return A String
'@export
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

'@label:TableRange
'@prop-title Get one table of the password worksheet
'@param tabName Optional, String. The name of the table to return
'@export
'@return A Range Object
Private Property Get TableRange(Optional ByVal tabName As String = TABKEYS) As Range

    Dim sh As Worksheet
    Dim outRng As Range

    Set sh = Wksh()

    On Error Resume Next
    Set outRng = sh.ListObjects(tabName).DataBodyRange
    On Error GoTo 0

    If outRng Is Nothing Then
        'come here for logging because the table is empty
    End If

    Set TableRange = outRng
End Property

'@label:protections
'@section Protections


'@label:InDebugMode
'@prop-title Check if the current workbook is in debug mode
'return A Boolean
'@export
Private Property Get InDebugMode() As Boolean
    Dim sh As Worksheet
    Set sh = Wksh()
    InDebugMode = (sh.Range("RNG_DebugMode").Value = "yes")
End Property

'@label:EnterDebugMode
'@prop-title Enter a debug mode
'@details
'Debug Mode is a mode where all the worksheets of the current workbook are
'unprotected as well as the workbook. This will give the current user all the
'power to modify, use and delete data in the current workbook. This is
'usually useful in case something go wrong in the workbook on in a debugging
'process
'
'@export
Private Sub EnterDebugMode(Optional ByVal wb As Object = Nothing)

    Dim targetWb As Workbook
    Dim psh As Worksheet
    Dim sh As Worksheet

    Set psh = Wksh()


    'The default targeted workbook is the
    'current workbook (where the password)
    If( TypeName(wb) = "Nothing") Then
        Set targetWb = psh.Parent
    ElseIf (TypeName(wb) = "Workbook") Then
        Set targetwb = wb
    End If


    'Unprotect the current workbook
    UnProtectWkb targetwb, useTable:=True

    For Each sh In targetWb
        If IsProtected(sh.Name) Then
            UnProtect sh.Name, useTable:=True
        End If
    Next

    psh.Range("RNG_DebugMode").Value = "yes"

    ShowDebug
    LogInfo "The current workbook is in debug mode", checkingWarning
End Sub


'@label:LeaveDebugMode
'@prop-title Leave the debug mode
'@details
'Debug Mode is a mode where all the worksheets of the current workbook are
'unprotected as well as the workbook. This will give the current user all the
'power to modify, use and delete data in the current workbook. This is
'usually useful in case something go wrong in the workbook on in a debugging
'process
'
'@export
Private Sub LeaveDebugMode(Optional ByVal wb As Object = Nothing)

    Dim targetWb As Workbook
    Dim psh As Worksheet
    Dim sh As Worksheet

    Set psh = Wksh()

    'The default targeted workbook is the
    'current workbook (where the password)
    If( TypeName(wb) = "Nothing") Then
        Set targetWb = psh.Parent
    ElseIf (TypeName(wb) = "Workbook") Then
        Set targetwb = wb
    End If

    'Unprotect the current workbook
    ProtectWkb targetwb, useTable:=True

    For Each sh In targetWb
        If Not IsProtected(sh.Name) Then
            Protect sh.Name, useTable:=True
        End If
    Next

    psh.Range("RNG_DebugMode").Value = "no"
End Sub

'@label:GenerateKey
'@sub-title Generate a new public/private key
'@details
'Shuffle the table and select one pair of public/private key after shuffling.
'@export
Private Sub GenerateKey(ByVal trads As ITranslation)
    Dim nbRows As Long
    Dim randNum As Integer
    Dim LoKeys As listObject 'ListObject of keys
    Dim LoRng As Range
    Dim sh As Worksheet

    Set sh = Wksh()
    Set LoKeys = sh.ListObjects(TABKEYS)

    'Don't try to generate new keys if the databodyrange does not exists
    If LoKeys.DataBodyRange Is Nothing Then
        LogInfo "Table for private/public key is empty"
        ShowDebug
        Exit Sub
    End If

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


'@label:DisplayPrivateKey
'@sub-title Display the value of the private key
'@param trads A Translation Object used to show messages
'@export
Private Sub DisplayPrivateKey(ByVal trads As ITranslation)
    Dim sh As Worksheet
    Dim privKey As String
    Set sh = Wksh()
    privKey = sh.Range("RNG_PrivateKey").Value
    MsgBox trads.TranslatedValue("MSG_Password") & privKey
End Sub

'@label:Protect
'@sub-title Protect one worksheet
'@details
'The protect sub can have a worksheet, or a worksheet name as argument. It
'can also use a "_active" tag to protect the current active worksheet. If useTables
'is true, protected worksheets are added to the current protected sheets table.
'@param sheetObject A Variant; either the worksheet or the worksheet name
'@param allowShapes Optional, Boolean. Allow drowing objects. Used
'for graphs. Default is False.
'@param allowDeletingRow Optional, Boolean. Default is True
'@param useTable Optional, Boolean. Use the Protected table parameters for locking. Default is False.
'@export
Private Sub Protect(ByVal sheetObject As Variant, _
                    Optional ByVal allowShapes As Boolean = False, _
                    Optional ByVal allowDeletingRows As Boolean = True, _
                    Optional ByVal useTable As Boolean = False)

    Dim pwd As String
    Dim sh As Worksheet
    Dim wb As Workbook
    Dim drawObjects As Boolean
    Dim deleteRow As Boolean

    Set wb = Wksh.Parent

    'If the sheetObject is a worksheet, unProtect it, if it is a string
    If Not useTable Then
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

        drawObjects = (Not allowShapes)
        deleteRow = allowDeletingRows

    Else
        drawObjects = TableValue(sheetObject, "drawObject") <> "no")
        deleteRows = TableValue(sheetObject, "deleteRows") <> "no")
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
                   DrawingObjects:=drawObjects, _
                   Contents:=True, _
                   Scenarios:=True, _
                   AllowInsertingRows:=True, _
                   AllowSorting:=False, _
                   AllowFiltering:=True, _
                   AllowFormattingCells:=True, _
                   AllowDeletingRows:=deleteRows _
                   AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True

        If useTable Then AddValue sheetObject, drawObjects, deleteRows
    End If
End Sub


'Extract some properties from the table where we have informations on worksheets
'protected
Private Function TableValue(ByVal sheetObject As Variant, _
                            Optional ByVal colIndex As Byte = 1) As String

    Dim sh As Worksheet
    Dim idkey As String
    Dim idRng As Range
    Dim cellRng As Range
    Dim outValue As String
    Dim Lo As ListObject

    Set sh = Wksh()
    Set Lo = sh.ListObjects(TABPROTECTION)

    If (TypeName(sheetObject) = "Worksheet") Then
      TableValue sheetObject.Name, colIndex
    End If

    If (TypeName(sheetObject) <> "String") Then
        LogInfo "Tried to add values with something that is not a worksheet", checkingWarning
        ShowDebug
        Exit Sub
    End If
    idKey = sheetObject
    idRng = Lo.ListColumns(1).DataBodyRange

    'Look for the worksheet/Workbook value and retrieve it
    On Error Resume Next
        Set cellRng = idRng.Find(What:=idKey, LooAt:=xlWhole, MatchCase:=True, _ 
                                 after:=idRng.Cells(idRng.Rows.Count, 1))
    On Error GoTo 0
    
    'in case nothing is found, return vbNullString
    If Not (cellRng Is Nothing) Then
        outValue = cellRng.Cells(1, colIndex).Value
    End If

    TableValue = outValue
End Function


Private Sub AddValue(ByVal sheetObject As Variant, _
                     Optional ByVal drawObjects As Boolean = True, _
                     Optional ByVal deleteRows As Boolean = True)


    Dim Lo As ListObject
    Dim IdRng As Range
    Dim cellRng As Range
    Dim idValue As String
    Dim resizeLo As Boolean
    Dim sh As Worksheet

    If (TypeName(sheetObject) = "Worksheet") Then
        AddValue sheetObject.Name, allowShape, allowDeletingRows
    End If

    If (TypeName(sheetObject) <> "String") Then
        LogInfo "Tried to add values with something that is not a worksheet", checkingWarning
        ShowDebug
        Exit Sub
    End If

    idValue = sheetObject

    Set sh = Wksh()
    Set Lo = sh.ListObjects(TABPROTECTION)
    Set idRng = Lo.ListColumns(1).DataBodyRange

    On Error Resume Next
        Set cellRng = idRng.Find(What:=idValue, MatchCase:=True, LookAt:=xlWhole, _
                                after:=idRng.Cells(idRng.Rows.Count, 1))
    On Error GoTo 0

    If (cellRng Is Nothing) Then
        '
        Set cellRng = idRng.Cells(1, 1)
        Do While (cellRng.Value <> vbNullString)
            Set cellRng = cellRng.Offset(1)
        Loop
        resizeLo = (cellRng.Row > idRng.cells(idRng.Rows.Count, 1).Row)
    End If

    cellRng.Value = idValue
    cellRng.Offset(, 1).Value = IIf(allowShape, "yes", "no")
    cellRng.Offset(, 2).Value = IIf(allowDeletingRows, "yes", "no")

    'Resize the listObject if necessary
    If resizeLo Then
        With sh
            Set cellRng = .Range(idRng.Cells(0, 1), cellRng.Cells(1, 3))
        End With
        Lo.Resize cellRng
    End If
End Sub



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

Private Sub Import(ByVal pass As IPasswords)

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

'@label:checkings
'@section Checkings

@label:HasChecking
'@prop-title Check if the object has checkings
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Get the cheking values
'@export
Private Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property


'@sub-title Add elements to the checking object for future notes
'@param Label String, the message to add.
'@param scope Byte. The scope of the checking object.
'@export
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "DropdownList " & Name() & " object trace"

    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)

    'Add new informations to the checking object
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'Requirements for instanciation
Private Sub CheckRequirements(ByVal sh As Worksheet)

    If sh Is Nothing Then ThrowError "Unable to instanciate, no password sheet"

    'Required listObjects fo translations
    LoExists TABKEYS, sh
    LoExists TABPROTECTION, sh

    If sh.ListObjects(TABKEYS).DataBodyRange Is Nothing Then ThrowError "Unable to instanciate password, the keys table is empty"

    'Ranges in the password worksheets that should be present
    RangeExists rngName:="RNG_DebuggingPassword", sh:=sh
    RangeExists rngName:="RNG_PublicKey", sh:=sh
    RangeExists rngName:="RNG_LabPublicKey", sh:=sh
    RangeExists rngName:="RNG_PrivateKey", sh:=sh
    RangeExists rngName:="RNG_LabPrivateKey", sh:=sh
    RangeExists rngName:="RNG_DebugMode", sh:=sh
    RangeExists rngName:="RNG_Version", sh:=sh

End Sub


Private Sub ShowDebug()
    Debug.Print "Check the informations of dropdownlist object " & Name()
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



'@jump:Create
'@section Interface

'@jump:elements
'Password object elements ------------------------------------------------------

'password object name
Private Property Get IPasswords_Name() As String
    IPasswords_Name = Name()
End Property

'debugging password
Private Property Get IPasswords_PassWord() As String
    IPasswords_PassWord = Password()
End Property

'table range
Private Property Get IPasswords_TableRange(ByVal tagName As String) As Range
    Set IPasswords_TableRange = TableRange(tabName)
End Property

'one value of a range in the password object
Private Property Get IPasswords_Value(ByVal tagName As String) As String
    IPasswords_Value = Value(tagName)
End Property

'@jump:protections
'protections -------------------------------------------------------------------

'Enter and leave degub mode

'@jump:EnterDebugMode
Private Sub IPasswords_EnterDebugMode()
    EnterDebugMode
End Sub

'@jump:LeaveDebugMode
Private Sub IPasswords_LeaveDebugMode()
    LeaveDebugMode
End Sub

'@jump:GenerateKey
Private Sub IPasswords_GenerateKey(ByVal trads As ITranslation)
    GenerateKey trads
End Sub

'@jump:DisplayPrivateKey
Private Sub IPasswords_DisplayPrivateKey(ByVal trads As ITranslation)
    DisplayPrivateKey trads
End Sub



'@jump:importexport
'Import/export -----------------------------------------------------------------


Private Sub IPasswords_UnProtect(Optional ByVal sheetName As String = "_active")
    UnProtect sheetName
End Sub

Private Sub IPasswords_Protect(Optional ByVal sheetObject As Variant, _
                                 Optional ByVal allowShapes As Boolean = False, _
                                 Optional ByVal allowDeletingRows As Boolean = True)
    Protect sheetObject, allowShapes, allowDeletingRows
End Sub

Private Sub IPasswords_Export(ByVal toWkb As Workbook)
    Export toWkb
End Sub

    'Protect and unprotect the workbook to avoid sheet name changes
Private Sub IPasswords_ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
    ProtectWkb wb, protectSheetName
End Sub

Private Sub IPasswords_UnProtectWkb(ByVal wb As Workbook)
    UnProtectWkb wb
End Sub





Private Sub IPasswords_Import(ByVal pass As IPasswords)
    Import pass
End Sub


'@jump:checkings
'Checkings ---------------------------------------------------------------------

'@jump:HasChekings
'Test if the object has checkings
Private Property Get IPasswords_HasCheckings() As Boolean
 IPasswords_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
'Get the checking values
Private Property Get IPasswords_CheckingValues() As Object
    Set IPasswords_CheckingValues = CheckingValues()
End Property