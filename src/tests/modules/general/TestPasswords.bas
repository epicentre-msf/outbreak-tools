Attribute VB_Name = "TestPasswords"

'@ModuleDescription Tests for the Passwords class, covering factory creation, worksheet and workbook
'   protection, debug mode lifecycle, key generation, import/export operations, cloning,
'   and VBProject code injection for the debug exit handler.
'
' @description
'   This test module exercises the full public surface of the IPasswords interface as
'   implemented by the Passwords class. Tests are organised into logical groups:
'
'     - Creation and value retrieval: verifies that Passwords.Create correctly
'       populates named ranges and exposes key values.
'     - Protection: validates Protect/UnProtect on worksheets and workbooks, including
'       settings persistence in the T_ProtectedSheets table.
'     - Debug mode: ensures EnterDebugMode and LeaveDebugMode toggle protection state
'       correctly and record diagnostic log entries via the Checking interface.
'     - Import/Export/Clone: tests data transfer between IPasswords instances across
'       worksheets and workbooks.
'     - Key generation: confirms GenerateKey selects a valid key pair, updates named
'       ranges, and captures the private key prompt.
'     - VBProject injection: verifies EnsureDebugExitHandler injects the
'       LeaveDebugModeOnClose routine into a workbook's ThisWorkbook module, preserves
'       existing Workbook_BeforeClose code, and avoids duplicate injection.
'     - Roundtrip: full save/reopen cycle to confirm the debug exit handler fires
'       correctly on workbook close.
'
'   The module relies on PasswordsTestFixture to build a consistent fixture sheet with
'   known named ranges and tables. A separate FixtureWorkbook is created per test to
'   ensure full isolation. Tests that modify VBProject components require VBA project
'   access to be enabled in the host application's Trust Center; when access is denied,
'   those tests are skipped with an explicit log message.
'
' @depends Passwords, IPasswords, TranslationObject, ITranslationObject,
'   ApplicationState, IApplicationState, HiddenNames, IHiddenNames, Checking,
'   IChecking, BetterArray, CustomTest, TestHelpers, PasswordsTestFixture

Option Explicit



'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private PasswordSubject As IPasswords
Private FixtureSheet As Worksheet
Private ProtectedSheet As Worksheet
Private FixtureWorkbook As Workbook

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const DEFAULTPASSWORDSHEET As String = "PasswordsFixture"
Private Const PROTECTEDSHEETNAME As String = "PasswordsProtectedFixture"
Private Const TABLEKEYS As String = "T_keys"
Private Const TABLEPROTECTED As String = "T_ProtectedSheets"
Private Const NAMEDEBUGPASSWORD As String = "RNG_DebuggingPassword"
Private Const NAMEPUBLICKEY As String = "RNG_PublicKey"
Private Const NAMEPRIVATEKEY As String = "RNG_PrivateKey"
Private Const NAMEDEBUGMODE As String = "RNG_DebugMode"
Private Const NAMEVERSION As String = "RNG_Version"
Private Const NAMEPROTECTEDSHEETS As String = "Passwords_ProtectedSheets"
Private Const DEFAULTBOOLYES As String = "yes"
Private Const DEFAULTBOOLNO As String = "no"
Private Const DEFAULTTRANSLATIONSHEET As String = "PasswordsTranslations"
Private Const TRANSLATIONTABLE As String = "T_PasswordTranslations"
Private Const TRANSLATIONLANGUAGE As String = "English"
Private Const ERR_VBPROJECT_ACCESS_DENIED As Long = 1004
Private Const ERR_VBPROJECT_NOT_SET As Long = 91

'@section Helper builders
'===============================================================================

' @sub-title Build an in-memory TranslationObject for password prompt tests
Private Function CreatePasswordTranslator() As ITranslationObject
    Dim translationSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim translationTable As ListObject

    'Set up an in-memory translation table so message strings resolve without additional fixtures.
    Set translationSheet = TestHelpers.EnsureWorksheet(DEFAULTTRANSLATIONSHEET, FixtureWorkbook)

    headerMatrix = TestHelpers.RowsToMatrix(Array(Array("tag", TRANSLATIONLANGUAGE)))
    TestHelpers.WriteMatrix translationSheet.Cells(1, 1), headerMatrix

    dataMatrix = TestHelpers.RowsToMatrix(Array( _
        Array("MSG_Password", "Password:"), _
        Array("MSG_Title", "Credentials")))
    TestHelpers.WriteMatrix translationSheet.Cells(2, 1), dataMatrix

    Set translationTable = translationSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                           Source:=translationSheet.Range("A1").CurrentRegion, _
                                                           XlListObjectHasHeaders:=xlYes)
    translationTable.Name = TRANSLATIONTABLE

    Set CreatePasswordTranslator = TranslationObject.Create(translationTable, TRANSLATIONLANGUAGE)
End Function

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestPasswords"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    BusyApp
    PasswordsTestFixture.PreparePasswordsFixture DEFAULTPASSWORDSHEET, FixtureWorkbook
    Set FixtureSheet = FixtureWorkbook.Worksheets(DEFAULTPASSWORDSHEET)
    Set ProtectedSheet = TestHelpers.EnsureWorksheet(PROTECTEDSHEETNAME, FixtureWorkbook)
    Set PasswordSubject = Passwords.Create(FixtureSheet)
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then
            FixtureWorkbook.Names(NAMEPROTECTEDSHEETS).Delete
        End If
    On Error GoTo 0

    BusyApp
    If Not FixtureWorkbook Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook FixtureWorkbook
        Set FixtureWorkbook = Nothing
    End If

    Set PasswordSubject = Nothing
    Set FixtureSheet = Nothing
    Set ProtectedSheet = Nothing
End Sub

' @sub-title Verify that Passwords.Create initialises named values and returns a valid PasswordSheet
' @details Tests the factory method Passwords.Create by passing a prepared fixture sheet.
'   Arrange: TestInitialize creates FixtureWorkbook and calls PreparePasswordsFixture to
'   populate named ranges, then creates PasswordSubject via Passwords.Create. Act/Assert:
'   verifies that Value("debuggingpassword") returns the expected fixture value "1234" and
'   that PasswordSheet returns a Worksheet reference (not Nothing). This confirms that the
'   factory correctly wires up the internal named-range lookup and sheet reference.
'@TestMethod("Passwords")
Public Sub TestCreateInitialisesNamedValues()
    CustomTestSetTitles Assert, "Passwords", "TestCreateInitialisesNamedValues"
    Assert.AreEqual "1234", PasswordSubject.Value("debuggingpassword"), _
                     "Create should expose the debugging password value through Value()"
    Assert.IsTrue TypeName(PasswordSubject.PasswordSheet) = "Worksheet", _
                   "PasswordSheet must return a worksheet reference"
End Sub

' @sub-title Verify that Value exposes laboratory public and private keys
' @details Tests that Value() correctly retrieves laboratory key values from the fixture.
'   Arrange: the fixture sheet is pre-populated with labpublickey "LABPUB123" and
'   labprivatekey "LABPRIV456" in named ranges. Act: calls Value("labpublickey") and
'   Value("labprivatekey"). Assert: confirms both return the expected fixture strings,
'   validating that the keys table lookup resolves multiple named key entries correctly.
'@TestMethod("Passwords")
Public Sub TestValueExposesLabKeys()
    CustomTestSetTitles Assert, "Passwords", "TestValueExposesLabKeys"
    Assert.AreEqual "LABPUB123", PasswordSubject.Value("labpublickey"), _
                     "Value should expose the laboratory public key"
    Assert.AreEqual "LABPRIV456", PasswordSubject.Value("labprivatekey"), _
                     "Value should expose the laboratory private key"
End Sub

' @sub-title Verify that Protect applies worksheet protection and persists settings in the table
' @details Tests the Protect method by applying protection to a worksheet with specific
'   options and then checking persistence. Arrange: ProtectedSheet is an unprotected
'   fixture worksheet. Act: calls Protect with allowShapes=False, allowDeletingRows=True,
'   and registerState=True. Assert: checks that ProtectContents is True on the worksheet,
'   that the T_ProtectedSheets table now contains a data body range, that the sheet name
'   appears in column 1, and that the allowShapes and allowDeletingRows flags are persisted
'   as "no" and "yes" respectively. This validates both the runtime protection and the
'   settings round-trip for the LeaveDebugMode restoration logic.
'@TestMethod("Passwords")
Public Sub TestProtectPersistsSettings()
    CustomTestSetTitles Assert, "Passwords", "TestProtectPersistsSettings"
    PasswordSubject.Protect ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=True, registerState:=True

    Assert.IsTrue ProtectedSheet.ProtectContents, "Protect should apply worksheet protection"

    Dim settings As Range
    Set settings = PasswordSubject.TableRange(TABLEPROTECTED, includeHeaders:=False)
    Assert.ObjectExists settings, "Range", "Protection settings should populate the protected sheets table"

    Dim found As Range
    On Error Resume Next
        Set found = settings.Columns(1).Find(What:=ProtectedSheet.Name, LookAt:=xlWhole, MatchCase:=True)
    On Error GoTo 0

    Assert.ObjectExists found, "Range", "Protect should record the sheet name in the protection table"
    Assert.AreEqual DEFAULTBOOLNO, CStr(found.Offset(0, 1).Value), _
                     "AllowShapes preference should be stored as 'no'"
    Assert.AreEqual DEFAULTBOOLYES, CStr(found.Offset(0, 2).Value), _
                     "AllowDeletingRows preference should be stored as 'yes'"
End Sub

' @sub-title Verify that UnProtect reverses worksheet protection
' @details Tests the UnProtect method as the inverse of Protect. Arrange: the test first
'   calls Protect on ProtectedSheet with registerState=False to apply protection without
'   table registration. Act: calls UnProtect on the same sheet name. Assert: confirms
'   that ProtectContents is now False, verifying that UnProtect correctly removes the
'   worksheet protection that was previously applied.
'@TestMethod("Passwords")
Public Sub TestUnProtectReleasesWorksheet()
    CustomTestSetTitles Assert, "Passwords", "TestUnProtectReleasesWorksheet"
    PasswordSubject.Protect ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=False, registerState:=False
    PasswordSubject.UnProtect ProtectedSheet.Name

    Assert.IsFalse ProtectedSheet.ProtectContents, "UnProtect should release worksheet protection"
End Sub

' @sub-title Verify that the "_wbactive" token routes Protect/UnProtect to the workbook level
' @details Tests the special "_wbactive" token that dispatches protection to the workbook
'   structure rather than a worksheet. Arrange: starts by calling UnProtect("_wbactive")
'   to ensure a clean baseline. Act: calls Protect("_wbactive") then UnProtect("_wbactive")
'   in sequence. Assert: checks ProtectStructure is False initially, True after Protect,
'   and False again after UnProtect. This validates the TypeName dispatch path inside the
'   Passwords class that routes workbook-level protection through the same Protect/UnProtect
'   interface.
'@TestMethod("Passwords")
Public Sub TestProtectWorkbookUsingActiveToken()
    CustomTestSetTitles Assert, "Passwords", "TestProtectWorkbookUsingActiveToken"
    PasswordSubject.UnProtect "_wbactive"
    Assert.IsFalse FixtureWorkbook.ProtectStructure, "Workbook should start unprotected"

    PasswordSubject.Protect "_wbactive"
    Assert.IsTrue FixtureWorkbook.ProtectStructure, "Protect should lock workbook structure"

    PasswordSubject.UnProtect "_wbactive"
    Assert.IsFalse FixtureWorkbook.ProtectStructure, "UnProtect should unlock workbook structure"
End Sub

' @sub-title Verify that EnterDebugMode and LeaveDebugMode toggle protection and restore state
' @details Tests the full debug mode lifecycle. Arrange: protects ProtectedSheet via Protect
'   so it is registered in the protected sheets table. Act: calls EnterDebugMode, which
'   should set the debug flag, unlock the workbook structure, and unprotect all tracked
'   sheets; then calls LeaveDebugMode, which should reverse all of those changes. Assert:
'   after EnterDebugMode checks that the debug flag is "yes", workbook structure is
'   unprotected, and the tracked sheet is unprotected; after LeaveDebugMode checks that
'   the debug flag is "no", workbook structure is protected, and the tracked sheet is
'   re-protected. This is the core test for the debug mode contract.
'@TestMethod("Passwords")
Public Sub TestEnterAndLeaveDebugModeRestoresProtections()
    CustomTestSetTitles Assert, "Passwords", "TestEnterAndLeaveDebugModeRestoresProtections"
    PasswordSubject.Protect ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=False
    PasswordSubject.EnterDebugMode

    Assert.AreEqual DEFAULTBOOLYES, CStr(FixtureSheet.Range(NAMEDEBUGMODE).Value), _
                     "EnterDebugMode should set the debug mode flag"
    Assert.IsFalse FixtureWorkbook.ProtectStructure, "EnterDebugMode should unlock workbook structure"
    Assert.IsFalse ProtectedSheet.ProtectContents, "EnterDebugMode should unprotect tracked sheets"

    PasswordSubject.LeaveDebugMode

    Assert.AreEqual DEFAULTBOOLNO, CStr(FixtureSheet.Range(NAMEDEBUGMODE).Value), _
                     "LeaveDebugMode should clear the debug mode flag"
    Assert.IsTrue FixtureWorkbook.ProtectStructure, "LeaveDebugMode should protect workbook structure"
    Assert.IsTrue ProtectedSheet.ProtectContents, "LeaveDebugMode should reapply worksheet protection"
End Sub

' @sub-title Verify that EnsureProtectedSheetsName creates a workbook-level name with structured reference
' @details Tests the creation of a workbook-scoped Name object that references the protected
'   sheets table. Arrange: the fixture workbook has the T_ProtectedSheets table already
'   populated by TestInitialize. Act: calls EnsureProtectedSheetsName. Assert: retrieves
'   the Name object by constant NAMEPROTECTEDSHEETS from the workbook's Names collection,
'   verifies it exists, and checks that its RefersTo value equals "=T_ProtectedSheets[#All]".
'   This structured reference is required by the linelist builder at deployment time.
'@TestMethod("Passwords")
Public Sub TestEnsureProtectedSheetsNameUsesStructuredReference()
    CustomTestSetTitles Assert, "Passwords", "TestEnsureProtectedSheetsNameUsesStructuredReference"
    PasswordSubject.EnsureProtectedSheetsName

    Dim nameObj As Name
    On Error Resume Next
        Set nameObj = FixtureWorkbook.Names(NAMEPROTECTEDSHEETS)
    On Error GoTo 0

    Assert.ObjectExists nameObj, "Name", "EnsureProtectedSheetsName should create the workbook-level name"
    Assert.AreEqual "=" & TABLEPROTECTED & "[#All]", CStr(nameObj.RefersTo), _
                     "Workbook-level name should reference the protected table using structured syntax"
End Sub

' @sub-title Verify that HasCheckings and CheckingValues report diagnostics after debug transitions
' @details Tests the diagnostic logging contract. Arrange: verifies that a freshly created
'   PasswordSubject has no checkings. Act: performs a full debug mode cycle by calling
'   EnterDebugMode followed by LeaveDebugMode. Assert: confirms HasCheckings is now True,
'   that CheckingValues returns a valid Checking object, extracts the first key from the
'   Checking's ListOfKeys, and verifies the log message contains the expected "Workbook
'   entered debug mode" substring. This validates that internal state transitions are
'   captured for downstream diagnostic display.
'@TestMethod("Passwords")
Public Sub TestHasCheckingsAfterDebugTransition()
    CustomTestSetTitles Assert, "Passwords", "TestHasCheckingsAfterDebugTransition"
    Assert.IsFalse PasswordSubject.HasCheckings, "Password handler should start without checkings"

    PasswordSubject.EnterDebugMode
    PasswordSubject.LeaveDebugMode

    Assert.IsTrue PasswordSubject.HasCheckings, "Debug mode transitions should add a checking entry"
    Assert.ObjectExists PasswordSubject.CheckingValues, "Checking", "CheckingValues should expose collected diagnostics"

    Dim checkingKeys As BetterArray
    Dim firstKey As String
    Set checkingKeys = PasswordSubject.CheckingValues.ListOfKeys
    firstKey = CStr(checkingKeys.Item(checkingKeys.LowerBound))

    Dim firstLogMessage As String
    firstLogMessage = PasswordSubject.CheckingValues.ValueOf(firstKey, checkingLabel)
    Assert.IsTrue InStr(1, firstLogMessage, "Workbook entered debug mode", vbTextCompare) > 0, _
                  "Debug transition log should capture the entry message"
End Sub

' @sub-title Verify that TableRange returns the data body range of a named table
' @details Tests the TableRange method, which provides access to ListObject data ranges
'   by table name. Arrange: the fixture sheet contains T_keys with four data rows
'   populated by PasswordsTestFixture. Act: calls TableRange(TABLEKEYS, includeHeaders:=False).
'   Assert: checks that the returned range has exactly four rows and that the first cell
'   contains "1234", matching the fixture data. This confirms the table lookup and header
'   exclusion logic.
'@TestMethod("Passwords")
Public Sub TestTableRangeReturnsDataBody()
    CustomTestSetTitles Assert, "Passwords", "TestTableRangeReturnsDataBody"
    Dim keysBody As Range
    Set keysBody = PasswordSubject.TableRange(TABLEKEYS, includeHeaders:=False)

    Assert.AreEqual 4, keysBody.Rows.Count, "Fixture keys table should expose four data rows"
    Assert.AreEqual "1234", CStr(keysBody.Cells(1, 1).Value), "First public key should match fixture data"
End Sub

' @sub-title Verify that ExportToWorkbook copies the password sheet as a hidden clone
' @details Tests the ExportToWorkbook method, which transfers the password sheet to another
'   workbook. Arrange: creates a new empty destination workbook. Act: calls ExportToWorkbook
'   with the destination workbook. Assert: retrieves the cloned sheet by its expected name
'   (DEFAULTPASSWORDSHEET) and verifies that its Visible property is xlSheetVeryHidden,
'   ensuring the passwords remain hidden from end users. Cleanup: closes the destination
'   workbook without saving.
'@TestMethod("Passwords")
Public Sub TestExportToWorkbookCreatesHiddenClone()
    CustomTestSetTitles Assert, "Passwords", "TestExportToWorkbookCreatesHiddenClone"
    Dim destination As Workbook
    Set destination = Workbooks.Add

    PasswordSubject.ExportToWorkbook destination

    Dim clonedSheet As Worksheet
    Set clonedSheet = destination.Worksheets(DEFAULTPASSWORDSHEET)
    Assert.AreEqual xlSheetVeryHidden, clonedSheet.Visible, "Exported password sheet should be hidden in destination"

    destination.Close SaveChanges:=False
End Sub

' @sub-title Verify that ImportFrom copies key values from a source Passwords instance
' @details Tests the ImportFrom method by confirming that key values are transferred between
'   two independent IPasswords instances. Arrange: creates a clone of PasswordSubject on a
'   new worksheet, then overwrites the clone's public key, private key, and keys table with
'   placeholder values. Act: calls clone.ImportFrom(PasswordSubject). Assert: checks that
'   the clone's publickey value now matches the original source value. This validates the
'   key transfer pathway used when updating an existing linelist with new credentials.
'@TestMethod("Passwords")
Public Sub TestImportFromCopiesKeys()
    CustomTestSetTitles Assert, "Passwords", "TestImportFromCopiesKeys"
    Dim destination As Worksheet
    Set destination = FixtureWorkbook.Worksheets.Add(After:=FixtureWorkbook.Worksheets(FixtureWorkbook.Worksheets.Count))
    destination.Name = "PwdImport" & Format(Timer, "000")

    Dim clone As IPasswords
    Set clone = PasswordSubject.CloneToWorksheet(destination)

    clone.PasswordSheet.Range(NAMEPUBLICKEY).Value = "placeholder"
    clone.PasswordSheet.Range(NAMEPRIVATEKEY).Value = "placeholder"

    Dim keysTable As ListObject
    Set keysTable = clone.PasswordSheet.ListObjects(TABLEKEYS)
    If Not keysTable.DataBodyRange Is Nothing Then
        keysTable.DataBodyRange.Value = "placeholder"
    End If

    clone.ImportFrom PasswordSubject

    Assert.AreEqual PasswordSubject.Value("publickey"), clone.Value("publickey"), _
                     "ImportFrom should copy public key value"
End Sub

' @sub-title Verify that ImportFrom from an empty source clears the target table and logs a checking
' @details Tests the edge case where the source keys table is empty. Arrange: creates a
'   clone on a new worksheet and places a "stale" value in the target keys table, then
'   deletes all rows from the source keys table. Act: calls target.ImportFrom(PasswordSubject)
'   where PasswordSubject now has an empty keys table. Assert: verifies that the target
'   keys table data body is either Nothing or has zero non-empty cells (CountA = 0), that
'   HasCheckings is True, and that the first log message contains the expected skip
'   reasoning. This guards against stale key data persisting when the source is empty.
'@TestMethod("Passwords")
Public Sub TestImportFromEmptySourceClearsTable()
    CustomTestSetTitles Assert, "Passwords", "TestImportFromEmptySourceClearsTable"
    Dim destinationSheet As Worksheet
    Set destinationSheet = FixtureWorkbook.Worksheets.Add(After:=FixtureWorkbook.Worksheets(FixtureWorkbook.Worksheets.Count))
    destinationSheet.Name = "PwdImportEmpty" & Format(Timer, "000")

    Dim target As IPasswords
    Set target = PasswordSubject.CloneToWorksheet(destinationSheet)

    Dim targetKeys As ListObject
    Set targetKeys = target.PasswordSheet.ListObjects(TABLEKEYS)
    If Not targetKeys.DataBodyRange Is Nothing Then
        targetKeys.DataBodyRange.Cells(1, 1).Value = "stale"
    End If

    Dim sourceKeys As ListObject
    Set sourceKeys = FixtureSheet.ListObjects(TABLEKEYS)
    Do While sourceKeys.ListRows.Count > 0
        sourceKeys.ListRows(sourceKeys.ListRows.Count).Delete
        If sourceKeys.DataBodyRange Is Nothing Then Exit Do
    Loop

    target.ImportFrom PasswordSubject

    Dim dataRange As Range
    Set dataRange = targetKeys.DataBodyRange
    If Not dataRange Is Nothing Then
        Assert.AreEqual 0, Application.WorksheetFunction.CountA(dataRange), _
                         "ImportFrom should clear destination keys table when source is empty"
    End If

    Assert.IsTrue target.HasCheckings, "ImportFrom should log a checking when source is empty"

    Dim logKeys As BetterArray
    Dim firstKey As String
    Set logKeys = target.CheckingValues.ListOfKeys
    firstKey = CStr(logKeys.Item(logKeys.LowerBound))

    Dim logMessage As String
    logMessage = target.CheckingValues.ValueOf(firstKey, checkingLabel)
    Assert.IsTrue InStr(1, logMessage, "Import skipped because source keys table is empty", vbTextCompare) > 0, _
                  "ImportFrom should record reasoning for skipped import"
End Sub

' @sub-title Verify that CloneToWorksheet creates an independent copy on a given sheet
' @details Tests CloneToWorksheet by cloning the password data to a new worksheet within
'   the same workbook. Arrange: creates a fresh worksheet named with a timestamp suffix.
'   Act: calls CloneToWorksheet to produce a new IPasswords instance backed by the new
'   sheet. Assert: confirms the cloned instance's publickey matches the original's value,
'   proving that the clone received a faithful copy of the keys data. Cleanup: deletes
'   the cloned worksheet to avoid polluting subsequent tests.
'@TestMethod("Passwords")
Public Sub TestCloneToWorksheetProducesIndependentSheet()
    CustomTestSetTitles Assert, "Passwords", "TestCloneToWorksheetProducesIndependentSheet"
    Dim cloneSheet As Worksheet
    Set cloneSheet = FixtureWorkbook.Worksheets.Add(After:=FixtureWorkbook.Worksheets(FixtureWorkbook.Worksheets.Count))
    cloneSheet.Name = "PasswordsClone" & Format(Timer, "000")

    Dim cloned As IPasswords
    Set cloned = PasswordSubject.CloneToWorksheet(cloneSheet)

    Assert.AreEqual PasswordSubject.Value("publickey"), cloned.Value("publickey"), _
                     "CloneToWorksheet should copy the public key"

    Application.DisplayAlerts = False
    cloneSheet.Delete
    Application.DisplayAlerts = True
End Sub

' @sub-title Verify that CloneToWorkbook creates a new workbook-hosted Passwords handler
' @details Tests CloneToWorkbook by cloning the password data into a separate workbook.
'   Arrange: creates a new empty workbook. Act: calls CloneToWorkbook, which should copy
'   the password sheet into the target workbook and return a new IPasswords instance.
'   Assert: verifies the cloned handler's PasswordSheet.Name matches DEFAULTPASSWORDSHEET,
'   confirming the sheet was created with the correct name. Cleanup: closes the temporary
'   workbook without saving.
'@TestMethod("Passwords")
Public Sub TestCloneToWorkbookProducesHandler()
    CustomTestSetTitles Assert, "Passwords", "TestCloneToWorkbookProducesHandler"
    Dim tempWb As Workbook
    Set tempWb = Workbooks.Add

    Dim cloned As IPasswords
    Set cloned = PasswordSubject.CloneToWorkbook(tempWb)

    Assert.AreEqual DEFAULTPASSWORDSHEET, cloned.PasswordSheet.Name, _
                     "CloneToWorkbook should create the password sheet in the target workbook"

    tempWb.Close SaveChanges:=False
End Sub


' @sub-title Verify that GenerateKey raises an error when the translator argument is Nothing
' @details Tests the input guard on GenerateKey. Arrange: PasswordSubject is created
'   normally. Act: calls GenerateKey(Nothing). Assert: expects the call to raise
'   ProjectError.ObjectNotInitialized. If no error is raised, the test logs a failure.
'   The On Error GoTo pattern captures the error number and compares it to the expected
'   value, validating that the guard clause fires before any key generation logic runs.
'@TestMethod("Passwords")
Public Sub TestGenerateKeyWithoutTranslatorRaises()
    CustomTestSetTitles Assert, "Passwords", "TestGenerateKeyWithoutTranslatorRaises"
    On Error GoTo ExpectError
        PasswordSubject.GenerateKey Nothing
        Assert.LogFailure "GenerateKey should raise when translator is missing"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "GenerateKey should guard against missing translator"
End Sub

' @sub-title Verify that DisplayPrivateKey raises an error when the translator argument is Nothing
' @details Tests the input guard on DisplayPrivateKey. Arrange: PasswordSubject is created
'   normally. Act: calls DisplayPrivateKey(Nothing). Assert: expects the call to raise
'   ProjectError.ObjectNotInitialized. This mirrors the GenerateKey guard test and ensures
'   both key-display methods enforce the translator precondition consistently.
'@TestMethod("Passwords")
Public Sub TestDisplayPrivateKeyWithoutTranslatorRaises()
    CustomTestSetTitles Assert, "Passwords", "TestDisplayPrivateKeyWithoutTranslatorRaises"
    On Error GoTo ExpectError
        PasswordSubject.DisplayPrivateKey Nothing
        Assert.LogFailure "DisplayPrivateKey should raise when translator is missing"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "DisplayPrivateKey should guard against missing translator"
End Sub

' @sub-title Verify that DisplayPrivateKey in silent mode captures the prompt text without showing a dialog
' @details Tests the DisplayPrompts=False path, which stores the private key prompt instead
'   of displaying it in a message box. Arrange: creates a translator via
'   CreatePasswordTranslator, sets DisplayPrompts to False. Act: calls DisplayPrivateKey.
'   Assert: checks that LastPrivatePrompt contains "Password:" concatenated with the
'   private key value, that LastPrivatePromptTitle equals "Credentials", that HasCheckings
'   is True, and that the last checking log entry contains the expected prompt text. This
'   is essential for automated testing since message boxes cannot be dismissed by code.
'@TestMethod("Passwords")
Public Sub TestDisplayPrivateKeySilentModeCapturesPrompt()
    CustomTestSetTitles Assert, "Passwords", "TestDisplayPrivateKeySilentModeCapturesPrompt"

    Dim translator As ITranslationObject
    Dim expectedPrompt As String
    Dim expectedTitle As String
    Dim logs As IChecking
    Dim logKeys As BetterArray
    Dim lastKey As String
    Dim loggedMessage As String

    Set translator = CreatePasswordTranslator()
    PasswordSubject.DisplayPrompts = False
    PasswordSubject.DisplayPrivateKey translator

    expectedPrompt = "Password:" & PasswordSubject.Value("privatekey")
    expectedTitle = "Credentials"

    Assert.AreEqual expectedPrompt, PasswordSubject.LastPrivatePrompt, _
                     "DisplayPrivateKey should expose the prompt text when prompts are suppressed"
    Assert.AreEqual expectedTitle, PasswordSubject.LastPrivatePromptTitle, _
                     "DisplayPrivateKey should expose the prompt title when prompts are suppressed"

    Assert.IsTrue PasswordSubject.HasCheckings, "DisplayPrivateKey should record a checking entry"
    Set logs = PasswordSubject.CheckingValues
    Assert.ObjectExists logs, "Checking", "CheckingValues should be available after displaying the key"

    Set logKeys = logs.ListOfKeys
    lastKey = CStr(logKeys.Item(logKeys.UpperBound))
    loggedMessage = logs.ValueOf(lastKey, checkingLabel)

    Assert.IsTrue InStr(1, loggedMessage, expectedPrompt, vbTextCompare) > 0, _
                  "Checking log should contain the prompt text"
End Sub

' @sub-title Verify that GenerateKey selects a key pair, updates named ranges, and captures the prompt
' @details Tests the full GenerateKey workflow. Arrange: records the current public and
'   private key values, overwrites the keys table with four known pairs, creates a
'   translator, and sets DisplayPrompts to False. Act: calls GenerateKey. Assert: confirms
'   the new public and private key values differ from the originals, that the selected pair
'   matches one of the four configured pairs (using a Collection-based lookup), and that
'   LastPrivatePrompt and LastPrivatePromptTitle contain the expected prompt text. This
'   covers the random selection, range update, and silent prompt capture in one test.
'@TestMethod("Passwords")
Public Sub TestGenerateKeyUpdatesRanges()
    CustomTestSetTitles Assert, "Passwords", "TestGenerateKeyUpdatesRanges"
    Dim translator As ITranslationObject
    Dim publicValue As String
    Dim privateValue As String
    Dim keysTable As ListObject
    Dim updatedPublic As String
    Dim updatedPrivate As String
    Dim expectedPrompt As String
    Dim newKeys As Variant
    Dim allowedPairs As Collection
    Dim idx As Long
    Dim pairKey As String

    publicValue = CStr(FixtureSheet.Range(NAMEPUBLICKEY).Value)
    privateValue = CStr(FixtureSheet.Range(NAMEPRIVATEKEY).Value)

    Set keysTable = FixtureSheet.ListObjects(TABLEKEYS)
    newKeys = TestHelpers.RowsToMatrix(Array( _
        Array("PUB-ALPHA", "PRIV-OMEGA"), _
        Array("PUB-BETA", "PRIV-GAMMA"), _
        Array("PUB-DELTA", "PRIV-LAMBDA"), _
        Array("PUB-OMEGA", "PRIV-ALPHA")))
    keysTable.DataBodyRange.Value = newKeys

    Set allowedPairs = New Collection
    For idx = LBound(newKeys, 1) To UBound(newKeys, 1)
        allowedPairs.Add newKeys(idx, 1) & "|" & newKeys(idx, 2)
    Next idx

    Set translator = CreatePasswordTranslator()
    PasswordSubject.DisplayPrompts = False
    PasswordSubject.GenerateKey translator

    updatedPublic = CStr(FixtureSheet.Range(NAMEPUBLICKEY).Value)
    updatedPrivate = CStr(FixtureSheet.Range(NAMEPRIVATEKEY).Value)

    Assert.IsTrue updatedPublic <> publicValue, "GenerateKey should update the public key value"
    Assert.IsTrue updatedPrivate <> privateValue, "GenerateKey should update the private key value"

    pairKey = updatedPublic & "|" & updatedPrivate
    For idx = 1 To allowedPairs.Count
        If allowedPairs(idx) = pairKey Then
            Exit For
        End If
    Next idx
    Assert.IsTrue idx <= allowedPairs.Count, _
                  "GenerateKey should select one of the configured key pairs"

    expectedPrompt = "Password:" & updatedPrivate
    Assert.AreEqual expectedPrompt, PasswordSubject.LastPrivatePrompt, _
                     "GenerateKey should prepare the private key prompt when prompts are suppressed"
    Assert.AreEqual "Credentials", PasswordSubject.LastPrivatePromptTitle, _
                     "GenerateKey should capture the prompt title when prompts are suppressed"
End Sub


' @sub-title Verify that EnsureDebugExitHandler injects the LeaveDebugModeOnClose routine into a workbook
' @details Tests VBProject code injection by calling EnsureDebugExitHandler on a temporary
'   workbook. Arrange: creates a new workbook, clones PasswordSubject into it. Act: calls
'   EnsureDebugExitHandler. Assert: reads the code from the workbook's ThisWorkbook module
'   and confirms it contains "LeaveDebugModeOnClose", "Workbook_BeforeClose", and the
'   conditional invocation "If Not Cancel Then LeaveDebugModeOnClose". If VBProject access
'   is denied (error 1004 or 91), the test is skipped with an explicit log message rather
'   than failing silently. Cleanup: deletes injected code lines and closes the temp workbook.
'@TestMethod("Passwords")
Public Sub TestEnsureDebugExitHandlerInjectsCode()
    CustomTestSetTitles Assert, "Passwords", "TestEnsureDebugExitHandlerInjectsCode"
    Dim tempWb As Workbook
    Set tempWb = Workbooks.Add

    On Error GoTo InjectionAccessDenied

    Dim cloned As IPasswords
    Set cloned = PasswordSubject.CloneToWorkbook(tempWb)
    cloned.EnsureDebugExitHandler tempWb

    Dim codeModule As Object
    Set codeModule = tempWb.VBProject.VBComponents(tempWb.CodeName).CodeModule

    Dim lines As String
    lines = codeModule.Lines(1, codeModule.CountOfLines)

    Assert.IsTrue InStr(1, lines, "LeaveDebugModeOnClose", vbTextCompare) > 0, _
                  "Workbook module should expose LeaveDebugModeOnClose"
    Assert.IsTrue InStr(1, lines, "Workbook_BeforeClose", vbTextCompare) > 0, _
                  "Workbook module should expose Workbook_BeforeClose"
    Assert.IsTrue InStr(1, lines, "If Not Cancel Then LeaveDebugModeOnClose", vbTextCompare) > 0, _
                  "Workbook_BeforeClose should call LeaveDebugModeOnClose without overriding Cancel"

InjectionCleanup:
    On Error Resume Next
        If Not codeModule Is Nothing Then
            If codeModule.CountOfLines > 0 Then codeModule.DeleteLines 1, codeModule.CountOfLines
        End If
        If Not tempWb Is Nothing Then tempWb.Close SaveChanges:=False
    Exit Sub

InjectionAccessDenied:
    If Err.Number = 1004 Or Err.Number = 91 Then
        Debug.Print "VBProject access is disabled; skipping injection test"
        Assert.LogFailure "VBProject access is disabled; skipping injection test"
    Else
        Debug.Print "Unexpected failure during debug handler injection: "
        Assert.LogFailure "Unexpected failure during debug handler injection: " & Err.Number & " - " & Err.Description
    End If
    Resume InjectionCleanup
End Sub

' @sub-title Verify that EnsureDebugExitHandler preserves existing Workbook_BeforeClose code and avoids duplicates
' @details Tests the merge behaviour when a Workbook_BeforeClose handler already exists.
'   Arrange: creates a temp workbook, clones passwords into it, and manually injects a
'   baseline Workbook_BeforeClose with a Debug.Print statement and error handling. Act:
'   calls EnsureDebugExitHandler twice in succession to test idempotency. Assert: reads
'   the merged procedure text and verifies it contains "If Not Cancel Then LeaveDebugModeOnClose"
'   (the injected call), that the original "Debug.Print ""Closing""" statement is preserved,
'   and that "LeaveDebugModeOnClose" appears exactly once (no duplicate injection). If
'   VBProject access is denied, the test is skipped. Cleanup: deletes code lines via
'   SafeDeleteLines and closes the temp workbook.
'@TestMethod("Passwords")
Public Sub TestEnsureDebugExitHandlerPreservesExistingBeforeCloseCode()
    CustomTestSetTitles Assert, "Passwords", "TestEnsureDebugExitHandlerPreservesExistingBeforeCloseCode"

    Dim tempWb As Workbook
    Dim cloned As IPasswords
    Dim codeModule As Object
    Dim baseLines As String
    Dim procStart As Long
    Dim procLines As Long
    Dim procText As String
    Dim firstCall As Long
    Dim app As IApplicationState

    On Error GoTo InjectionAccessDenied
        Set app = ApplicationState.Create(Application)
        app.ApplyBusyState suppressEvents:=True, calculateOnSave:=True

        Set tempWb = Workbooks.Add
        Set cloned = PasswordSubject.CloneToWorkbook(tempWb)

        Set codeModule = tempWb.VBProject.VBComponents(tempWb.CodeName).CodeModule

        baseLines = "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbNewLine & _
                    "    On Error GoTo Restore" & vbNewLine & _
                    "    Debug.Print ""Closing""" & vbNewLine & _
                    "Restore:" & vbNewLine & _
                    "End Sub"
        codeModule.InsertLines 3, baseLines
        DoEvents

        cloned.EnsureDebugExitHandler tempWb
        DoEvents

        cloned.EnsureDebugExitHandler tempWb
        DoEvents

        Debug.Print "Sent the two debug handlers"

        procStart = codeModule.ProcStartLine("Workbook_BeforeClose", 0)
        DoEvents

        procLines = codeModule.ProcCountLines("Workbook_BeforeClose", 0)
        DoEvents
        Debug.Print "Got the procLines"

        procText = codeModule.Lines(procStart, procLines)
        DoEvents
        Debug.Print "Got the procText"

        Assert.IsTrue InStr(1, procText, "If Not Cancel Then LeaveDebugModeOnClose", vbTextCompare) > 0, _
                      "Existing Workbook_BeforeClose should call LeaveDebugModeOnClose"

        Assert.IsTrue InStr(1, procText, "Debug.Print ""Closing""", vbBinaryCompare) > 0, _
                      "Existing Workbook_BeforeClose statements must remain intact"

        firstCall = InStr(1, procText, "LeaveDebugModeOnClose", vbTextCompare)
        Assert.IsTrue InStr(firstCall + 1, procText, "LeaveDebugModeOnClose", vbTextCompare) = 0, _
                      "LeaveDebugModeOnClose should be injected only once"
    GoTo InjectionCleanup

InjectionAccessDenied:
    If Err.Number = 1004 Or Err.Number = 91 Then
        Assert.LogFailure "VBProject access is disabled; skipping existing handler merge test"
    Else
        Assert.LogFailure "Unexpected failure during existing handler merge test: " & Err.Number & " - " & Err.Description
    End If
    Resume InjectionCleanup

InjectionCleanup:
    On Error Resume Next
        SafeDeleteLines codeModule, 2, codeModule.CountOfLines
        DoEvents
        codeModule.InsertLines codeModule.CountOfLines + 1, vbCrLf & "' EOF pad (Mac)"
        If Not (tempWb Is Nothing) Then
            tempWb.Close SaveChanges:=False
        End If
        app.Restore
    On Error GoTo 0
End Sub


' @sub-title Verify end-to-end debug exit handler roundtrip across save and reopen
' @details Tests the full lifecycle: clone passwords into a temp workbook, inject the debug
'   exit handler, enter debug mode, protect a sheet, leave debug mode, save to disk, close,
'   and reopen the file. Arrange: creates a temp workbook and export folder, clones
'   PasswordSubject, injects the handler, imports required VBA components (Passwords,
'   Checking, BetterArray, TranslationObject, etc.) so the handler code can execute on
'   reopen, protects a guard sheet, and enters debug mode. Act: calls LeaveDebugMode,
'   saves the workbook to disk, closes it, then reopens it. Assert: verifies that the
'   reopened workbook has ProtectStructure=True, the guard sheet has ProtectContents=True,
'   and the debug flag is "no". This confirms the Workbook_BeforeClose handler fired
'   correctly during the close event. If VBProject access is denied, the test is skipped.
'   Cleanup: restores Application.EnableEvents, closes the reopened workbook, deletes
'   exported files and the temp folder.
'@TestMethod("Passwords")
Public Sub TestDebugExitHandlerRoundtripPersistsProtections()
    CustomTestSetTitles Assert, "Passwords", "TestDebugExitHandlerRoundtripPersistsProtections"

    Dim tempWb As Workbook
    Dim cloned As IPasswords
    Dim guardSheet As Worksheet
    Dim guardSheetName As String
    Dim workbookPath As String
    Dim exportFolder As String
    Dim exportedFiles As Collection
    Dim previousEventState As Boolean
    Dim reopened As Workbook
    Dim debugFlagCell As Range

    exportFolder = TestHelpers.BuildTempFolder(ThisWorkbook, "PasswordTests")
    Set exportedFiles = New Collection

    previousEventState = Application.EnableEvents

    On Error GoTo AccessDenied
        Set tempWb = Workbooks.Add

        Set cloned = PasswordSubject.CloneToWorkbook(tempWb)
        cloned.EnsureDebugExitHandler tempWb

        ImportPasswordsComponents tempWb, exportFolder, exportedFiles

        DoEvents

        Set guardSheet = FirstWorkbookSheet(tempWb, DEFAULTPASSWORDSHEET)
        If guardSheet Is Nothing Then
            Set guardSheet = tempWb.Worksheets.Add
            guardSheet.Name = "PwdGuard" & Format$(Timer, "000")
        End If
        guardSheetName = guardSheet.Name

        cloned.Protect guardSheetName, allowShapes:=False, allowDeletingRows:=False
        cloned.EnterDebugMode tempWb

        DoEvents

        Application.EnableEvents = True

        workbookPath = TestHelpers.BuildWorkbookPath(exportFolder, "passwords_debug_roundtrip")
        If Dir$(workbookPath) <> vbNullString Then Kill workbookPath

        cloned.LeaveDebugMode tempWb

        tempWb.SaveAs Filename:=workbookPath, FileFormat:=xlExcel12

        tempWb.Close SaveChanges:=True

        DoEvents

        Set reopened = Workbooks.Open(Filename:=workbookPath)

        DoEvents

        Assert.IsTrue reopened.ProtectStructure, _
                      "Workbook structure should be protected after reopening"
        Assert.IsTrue reopened.Worksheets(guardSheetName).ProtectContents, _
                      "Tracked worksheet should remain protected after debug handler runs"

        Set debugFlagCell = reopened.Worksheets(DEFAULTPASSWORDSHEET).Range(NAMEDEBUGMODE)

        Assert.AreEqual DEFAULTBOOLNO, CStr(debugFlagCell.Value), _
                     "Debug mode flag should be cleared after workbook close handler runs"

Cleanup:
        On Error Resume Next
            Application.EnableEvents = previousEventState
            If Not reopened Is Nothing Then reopened.Close SaveChanges:=False
            If LenB(workbookPath) > 0 Then
                If Dir$(workbookPath) <> vbNullString Then Kill workbookPath
            End If
            If Not tempWb Is Nothing Then tempWb.Close SaveChanges:=False
            TestHelpers.CleanupExportedFiles exportedFiles
            Kill exportFolder
        On Error GoTo 0
        Exit Sub

AccessDenied:
        If Err.Number = ERR_VBPROJECT_ACCESS_DENIED Or Err.Number = ERR_VBPROJECT_NOT_SET Then
            Assert.LogFailure "VBProject access is disabled; skipping workbook roundtrip test"
        Else
            Assert.LogFailure "Unexpected failure during workbook roundtrip test: " & Err.Number & " - " & Err.Description
        End If
        Resume Cleanup
End Sub

'@section Helpers
'===============================================================================

' @sub-title Import required VBA components into a target workbook for roundtrip tests
Private Sub ImportPasswordsComponents(ByVal targetWorkbook As Workbook, _
                                      ByVal exportFolder As String, _
                                      ByVal exportedFiles As Collection)

    Dim components As Variant
    Dim idx As Long
    Dim exportPath As String

    components = Array("BetterArray", "IChecking", "Checking", "IPasswords", "Passwords", "TranslationObject", "ITranslationObject")

    For idx = LBound(components) To UBound(components)
        exportPath = TestHelpers.ExportComponentToFolder(ThisWorkbook, CStr(components(idx)), exportFolder)
        exportedFiles.Add exportPath
        targetWorkbook.VBProject.VBComponents.Import exportPath
    Next idx
End Sub

' @sub-title Return the first worksheet in a workbook whose name does not match the excluded name
Private Function FirstWorkbookSheet(ByVal wb As Workbook, ByVal excludedName As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, excludedName, vbBinaryCompare) <> 0 Then
            Set FirstWorkbookSheet = sh
            Exit Function
        End If
    Next sh
    Set FirstWorkbookSheet = Nothing
End Function


' @sub-title Safely delete lines from a VBE CodeModule, respecting Attribute and Option lines on Mac
Private Sub SafeDeleteLines(cm As Object, ByVal startLine As Long, ByVal count As Long)
    On Error GoTo Fail

    If count < 1 Then Exit Sub
    Dim total As Long: total = cm.CountOfLines

    ' Clamp to valid range
    If startLine < 1 Then startLine = 1
    If startLine > total Then Exit Sub
    If startLine + count - 1 > total Then count = total - startLine + 1
    If count < 1 Then Exit Sub

#If Mac Then
    ' Make VBE visible helps on some Mac builds
    On Error Resume Next
    Application.VBE.MainWindow.Visible = True
    DoEvents
    On Error GoTo Fail
#End If

    ' Avoid deleting ATTRIBUTES/Option Explicit accidentally on Mac
    Dim l As Long, firstReal As Long: firstReal = startLine
    For l = startLine To startLine + count - 1
        Dim t As String: t = LCase$(Trim$(cm.Lines(l, 1)))
        If Left$(t, 9) <> "attribute" And Left$(t, 6) <> "option" Then
            firstReal = l: Exit For
        End If
    Next l
    If firstReal > startLine Then
        count = count - (firstReal - startLine)
        startLine = firstReal
        If count < 1 Then Exit Sub
    End If

    ' Delete in small chunks to avoid sporadic failures
    Const chunk As Long = 200
    Do While count > 0
        Dim n As Long: n = IIf(count > chunk, chunk, count)
        cm.DeleteLines startLine, n
        DoEvents
        count = count - n
        ' After deletion, subsequent lines shift up; keep same startLine
    Loop
    Exit Sub
Fail:
    Err.Raise Err.Number, "SafeDeleteLines", Err.Description
End Sub
