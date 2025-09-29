Attribute VB_Name = "TestPasswords"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private PasswordSubject As IPasswords
Private FixtureSheet As Worksheet
Private ProtectedSheet As Worksheet
Private FixtureWorkbook As Workbook

Private Const PASSWORD_SHEET As String = "PasswordsFixture"
Private Const PROTECTED_SHEET_NAME As String = "PasswordsProtectedFixture"
Private Const TABLE_KEYS As String = "T_keys"
Private Const TABLE_PROTECTED As String = "T_ProtectedSheets"
Private Const NAME_DEBUG_PASSWORD As String = "RNG_DebuggingPassword"
Private Const NAME_PUBLIC_KEY As String = "RNG_PublicKey"
Private Const NAME_LAB_PUBLIC_KEY As String = "RNG_LabPublicKey"
Private Const NAME_PRIVATE_KEY As String = "RNG_PrivateKey"
Private Const NAME_LAB_PRIVATE_KEY As String = "RNG_LabPrivateKey"
Private Const NAME_DEBUG_MODE As String = "RNG_DebugMode"
Private Const NAME_VERSION As String = "RNG_Version"
Private Const NAME_PROTECTED_SHEETS As String = "Passwords_ProtectedSheets"
Private Const DEFAULT_BOOL_YES As String = "yes"
Private Const DEFAULT_BOOL_NO As String = "no"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET, FixtureWorkbook
    Set FixtureSheet = FixtureWorkbook.Worksheets(PASSWORD_SHEET)
    Set ProtectedSheet = TestHelpers.EnsureWorksheet(PROTECTED_SHEET_NAME, FixtureWorkbook)
    Set PasswordSubject = Passwords.Create(FixtureSheet)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    On Error Resume Next
    If Not FixtureWorkbook Is Nothing Then
        FixtureWorkbook.Names(NAME_PROTECTED_SHEETS).Delete
    End If
    On Error GoTo 0

    If Not FixtureWorkbook Is Nothing Then
        TestHelpers.DeleteWorkbook FixtureWorkbook
        Set FixtureWorkbook = Nothing
    End If

    TestHelpers.RestoreApp
    Set PasswordSubject = Nothing
    Set FixtureSheet = Nothing
    Set ProtectedSheet = Nothing
End Sub

'@TestMethod("Passwords")
Private Sub TestCreateInitialisesTagLookup()
    Assert.AreEqual "debugPwd", PasswordSubject.TagValue(PasswordTagDebug), _
                     "TagValue should expose the debugging password value"
End Sub

'@TestMethod("Passwords")
Private Sub TestProtectSheetPersistsSettings()
    PasswordSubject.ProtectSheet ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=True, useTable:=False

    Assert.IsTrue ProtectedSheet.ProtectContents, "ProtectSheet should apply protection"

    Dim settings As Range
    Set settings = PasswordSubject.TableRange(TABLE_PROTECTED, includeHeaders:=False)
    Assert.IsNotNothing settings, "Persisted settings should populate the protected sheets table"

    Dim found As Range
    On Error Resume Next
        Set found = settings.Columns(1).Find(What:=ProtectedSheet.Name, LookAt:=xlWhole, MatchCase:=True)
    On Error GoTo 0

    Assert.IsNotNothing found, "ProtectSheet should record the sheet name in the protection table"
    Assert.AreEqual DEFAULT_BOOL_YES, CStr(found.Offset(0, 2).Value), _
                     "AllowDeletingRows should be stored as 'yes'"
End Sub

'@TestMethod("Passwords")
Private Sub TestEnterAndLeaveDebugModeRestoresProtections()
    PasswordSubject.ProtectSheet ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=False, useTable:=False
    PasswordSubject.EnterDebugMode

    Assert.AreEqual DEFAULT_BOOL_YES, CStr(FixtureSheet.Range(NAME_DEBUG_MODE).Value), _
                     "Debug mode flag should switch to 'yes'"
    Assert.IsFalse FixtureWorkbook.ProtectStructure, "Workbook structure should be unlocked in debug mode"
    Assert.IsFalse ProtectedSheet.ProtectContents, "Protected sheet should be unprotected in debug mode"

    PasswordSubject.LeaveDebugModeIfActive

    Assert.AreEqual DEFAULT_BOOL_NO, CStr(FixtureSheet.Range(NAME_DEBUG_MODE).Value), _
                     "Debug mode flag should revert to 'no'"
    Assert.IsTrue FixtureWorkbook.ProtectStructure, "Workbook structure should be protected after leaving debug mode"
    Assert.IsTrue ProtectedSheet.ProtectContents, "Protected sheet should be protected after leaving debug mode"
End Sub

'@TestMethod("Passwords")
Private Sub TestEnsureProtectedSheetsNameAddsWorkbookName()
    PasswordSubject.EnsureProtectedSheetsName

    Dim nameObj As Name
    On Error Resume Next
        Set nameObj = FixtureWorkbook.Names(NAME_PROTECTED_SHEETS)
    On Error GoTo 0

    Assert.IsNotNothing nameObj, "EnsureProtectedSheetsName should register the workbook name"
    Assert.IsTrue nameObj.RefersTo Like "=*" & TABLE_PROTECTED & "*", _
                 "Workbook name should reference the protected sheets table"
End Sub

'@TestMethod("Passwords")
Private Sub TestCloneToProducesIndependentSheet()
    Dim cloneSheet As Worksheet
    Set cloneSheet = FixtureWorkbook.Worksheets.Add(After:=FixtureWorkbook.Worksheets(FixtureWorkbook.Worksheets.Count))
    cloneSheet.Name = "PasswordsClone" & Format(Timer, "000")

    Dim cloned As IPasswords
    Set cloned = PasswordSubject.CloneTo(cloneSheet)

    Assert.AreEqual PasswordSubject.TagValue(PasswordTagPublic), cloned.TagValue(PasswordTagPublic), _
                     "Clone should copy public key"

    Application.DisplayAlerts = False
    cloneSheet.Delete
    Application.DisplayAlerts = True
End Sub

'@TestMethod("Passwords")
Private Sub TestEnsureDebugExitHandlerInjectsCode()
    Dim tempWb As Workbook
    Set tempWb = Workbooks.Add

    On Error GoTo InjectionAccessDenied

    Dim targetSheet As Worksheet
    Set targetSheet = tempWb.Worksheets(1)
    targetSheet.Name = "PwdTemp"

    Dim cloned As IPasswords
    Set cloned = PasswordSubject.CloneTo(targetSheet)

    cloned.EnsureDebugExitHandler tempWb

    Dim codeModule As Object
    Set codeModule = tempWb.VBProject.VBComponents(tempWb.CodeName).CodeModule

    Dim lines As String
    lines = codeModule.Lines(1, codeModule.CountOfLines)

    Assert.IsTrue InStr(1, lines, "LeaveDebugModeOnClose", vbTextCompare) > 0, _
                  "Workbook module should now expose LeaveDebugModeOnClose"
    Assert.IsTrue InStr(1, lines, "Workbook_BeforeClose", vbTextCompare) > 0, _
                  "Workbook module should contain Workbook_BeforeClose handler"

InjectionCleanup:
    If Not tempWb Is Nothing Then tempWb.Close SaveChanges:=False
    Exit Sub

InjectionAccessDenied:
    If Err.Number = 1004 Or Err.Number = 91 Then
        Assert.Inconclusive "VBProject access is disabled; skipping injection test"
    Else
        Assert.Fail "Unexpected failure during debug handler injection: " & Err.Number & " - " & Err.Description
    End If
    Resume InjectionCleanup
End Sub
