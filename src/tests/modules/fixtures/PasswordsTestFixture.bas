Attribute VB_Name = "PasswordsTestFixture"
Attribute VB_Description = "Shared password fixture for tests"

Option Explicit

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Password worksheet fixture seeded from CSV drafts")

'@section Constants
'===============================================================================

Private Const BASE_PATH As String = "\src\tests\"
Private Const FILE_KEYS_RANGES As String = "draft3.csv"
Private Const FILE_KEYS_TABLE As String = "draft4.csv"
Private Const FILE_PROTECTED_TABLE As String = "draft5.csv"

Private Const NAME_DEBUG_PASSWORD As String = "RNG_DebuggingPassword"
Private Const NAME_PUBLIC_KEY As String = "RNG_PublicKey"
Private Const NAME_LAB_PUBLIC_KEY As String = "RNG_LabPublicKey"
Private Const NAME_PRIVATE_KEY As String = "RNG_PrivateKey"
Private Const NAME_LAB_PRIVATE_KEY As String = "RNG_LabPrivateKey"
Private Const NAME_DEBUG_MODE As String = "RNG_DebugMode"
Private Const NAME_VERSION As String = "RNG_Version"
Private Const NAME_PROTECTED_SHEETS As String = "Passwords_ProtectedSheets"

Private Const TABLE_KEYS As String = "T_keys"
Private Const TABLE_PROTECTED As String = "T_ProtectedSheets"

Private Const ADDRESS_PUBLIC_KEY As String = "A1"
Private Const ADDRESS_PRIVATE_KEY As String = "A2"
Private Const ADDRESS_DEBUG_PASSWORD As String = "A3"
Private Const ADDRESS_DEBUG_MODE As String = "A4"
Private Const ADDRESS_VERSION As String = "A5"
Private Const ADDRESS_LAB_PUBLIC As String = "A6"
Private Const ADDRESS_LAB_PRIVATE As String = "A7"

Private Const KEY_TABLE_START As String = "A10"
Private Const PROTECTED_TABLE_START As String = "D10"

'@section Public API
'===============================================================================

'@description Prepare a worksheet containing password data built from the draft CSV files.
'@param sheetName String. Worksheet name to create/reset.
'@param targetBook Optional Workbook into which the fixture should be loaded. Defaults to ThisWorkbook.
Public Sub PreparePasswordsFixture(ByVal sheetName As String, Optional ByVal targetBook As Workbook)

    Dim wb As Workbook
    Dim sh As Worksheet

    Set wb = ResolveWorkbook(targetBook)
    Set sh = TestHelpers.EnsureWorksheet(sheetName, wb)

    SeedNamedRanges sh
    SeedKeysTable sh
    SeedProtectedTable sh
    EnsureProtectedSheetsName wb, sh.ListObjects(TABLE_PROTECTED)
End Sub

'@section Worksheet Seeding Helpers
'===============================================================================

Private Sub SeedNamedRanges(ByVal sh As Worksheet)

    AddNamedRange sh, ADDRESS_PUBLIC_KEY, NAME_PUBLIC_KEY, "1234"
    AddNamedRange sh, ADDRESS_PRIVATE_KEY, NAME_PRIVATE_KEY, "1234"
    AddNamedRange sh, ADDRESS_DEBUG_PASSWORD, NAME_DEBUG_PASSWORD, "1234"
    AddNamedRange sh, ADDRESS_DEBUG_MODE, NAME_DEBUG_MODE, "No"
    AddNamedRange sh, ADDRESS_VERSION, NAME_VERSION, "d0099"
    AddNamedRange sh, ADDRESS_LAB_PUBLIC, NAME_LAB_PUBLIC_KEY, vbNullString
    AddNamedRange sh, ADDRESS_LAB_PRIVATE, NAME_LAB_PRIVATE_KEY, vbNullString
End Sub

Private Sub SeedKeysTable(ByVal sh As Worksheet)

    Dim startCell As Range
    Dim dataRange As Range
    Dim lo As ListObject
    Dim keysTable As Variant

    Set startCell = sh.Range(KEY_TABLE_START)
    keysTable = KeysTableMatrix()
    TestHelpers.WriteMatrix startCell, keysTable
    Set dataRange = startCell.Resize(UBound(keysTable, 1), UBound(keysTable, 2))

    Set lo = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = TABLE_KEYS
    lo.TableStyle = "TableStyleMedium2"
End Sub

Private Sub SeedProtectedTable(ByVal sh As Worksheet)

    Dim startCell As Range
    Dim dataRange As Range
    Dim lo As ListObject
    Dim protectedTable As Variant

    Set startCell = sh.Range(PROTECTED_TABLE_START)
    protectedTable = ProtectedTableMatrix()
    TestHelpers.WriteMatrix startCell, protectedTable
    Set dataRange = startCell.Resize(UBound(protectedTable, 1), UBound(protectedTable, 2))

    Set lo = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = TABLE_PROTECTED
    lo.TableStyle = "TableStyleMedium3"
End Sub

Private Sub EnsureProtectedSheetsName(ByVal wb As Workbook, ByVal protectedTable As ListObject)

    Dim refersTo As String
    Dim existing As Name

    refersTo = "=" & protectedTable.Range.Address(True, True, xlA1, True)

    On Error Resume Next
        Set existing = wb.Names(NAME_PROTECTED_SHEETS)
    On Error GoTo 0

    If existing Is Nothing Then
        wb.Names.Add Name:=NAME_PROTECTED_SHEETS, RefersTo:=refersTo
    Else
        existing.RefersTo = refersTo
    End If
End Sub

'@section Static Fixture Data
'===============================================================================

Private Function KeysTableMatrix() As Variant

    Dim rows As Variant

    rows = Array(_
        Array("PublicKeys", "PrivateKeys"), _
        Array("1234", "1234"), _
        Array("6789", "6789"))

    KeysTableMatrix = TestHelpers.RowsToMatrix(rows)
End Function

Private Function ProtectedTableMatrix() As Variant

    Dim rows As Variant

    rows = Array(_
        Array("ID", "DrawObjects", "DeleteRows"), _
        Array("", "", ""))

    ProtectedTableMatrix = TestHelpers.RowsToMatrix(rows)
End Function

Private Sub AddNamedRange(ByVal sh As Worksheet, _
                          ByVal address As String, _
                          ByVal rangeName As String, _
                          ByVal value As String)

    Dim wb As Workbook

    Set wb = sh.Parent
    sh.Range(address).Value = value

    On Error Resume Next
        wb.Names(rangeName).Delete
        sh.Names(rangeName).Delete
    On Error GoTo 0

    wb.Names.Add Name:=rangeName, RefersTo:=sh.Range(address)
End Sub

Private Function ResolveWorkbook(ByVal candidate As Workbook) As Workbook
    If candidate Is Nothing Then
        Set ResolveWorkbook = ThisWorkbook
    Else
        Set ResolveWorkbook = candidate
    End If
End Function
