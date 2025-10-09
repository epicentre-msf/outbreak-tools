Attribute VB_Name = "TestPasswords"

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
        Array("MSG_Password", "Password: "), _
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

'@TestMethod("Passwords")
Public Sub TestCreateInitialisesNamedValues()
    CustomTestSetTitles Assert, "Passwords", "TestCreateInitialisesNamedValues"
    Assert.AreEqual "1234", PasswordSubject.Value("debuggingpassword"), _
                     "Create should expose the debugging password value through Value()"
    Assert.IsTrue TypeName(PasswordSubject.PasswordSheet) = "Worksheet", _
                   "PasswordSheet must return a worksheet reference"
End Sub

'@TestMethod("Passwords")
Public Sub TestValueExposesLabKeys()
    CustomTestSetTitles Assert, "Passwords", "TestValueExposesLabKeys"
    Assert.AreEqual "LABPUB123", PasswordSubject.Value("labpublickey"), _
                     "Value should expose the laboratory public key"
    Assert.AreEqual "LABPRIV456", PasswordSubject.Value("labprivatekey"), _
                     "Value should expose the laboratory private key"
End Sub

'@TestMethod("Passwords")
Public Sub TestProtectPersistsSettings()
    CustomTestSetTitles Assert, "Passwords", "TestProtectPersistsSettings"
    PasswordSubject.Protect ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=True, registerState:=False

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

'@TestMethod("Passwords")
Public Sub TestUnProtectReleasesWorksheet()
    CustomTestSetTitles Assert, "Passwords", "TestUnProtectReleasesWorksheet"
    PasswordSubject.Protect ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=False, registerState:=False
    PasswordSubject.UnProtect ProtectedSheet.Name

    Assert.IsFalse ProtectedSheet.ProtectContents, "UnProtect should release worksheet protection"
End Sub

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

'@TestMethod("Passwords")
Public Sub TestEnterAndLeaveDebugModeRestoresProtections()
    CustomTestSetTitles Assert, "Passwords", "TestEnterAndLeaveDebugModeRestoresProtections"
    PasswordSubject.Protect ProtectedSheet.Name, allowShapes:=False, allowDeletingRows:=False, registerState:=False
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

'@TestMethod("Passwords")
Public Sub TestTableRangeReturnsDataBody()
    CustomTestSetTitles Assert, "Passwords", "TestTableRangeReturnsDataBody"
    Dim keysBody As Range
    Set keysBody = PasswordSubject.TableRange(TABLEKEYS, includeHeaders:=False)

    Assert.AreEqual 4, keysBody.Rows.Count, "Fixture keys table should expose four data rows"
    Assert.AreEqual "1234", CStr(keysBody.Cells(1, 1).Value), "First public key should match fixture data"
End Sub

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

    exportFolder = TestHelpers.ResolveExportFolder(ThisWorkbook, "PasswordTests")
    Set exportedFiles = New Collection

    previousEventState = Application.EnableEvents

    On Error GoTo AccessDenied
        Set tempWb = Workbooks.Add

        Set cloned = PasswordSubject.CloneToWorkbook(tempWb)
        cloned.EnsureDebugExitHandler tempWb

        ImportPasswordsComponents tempWb, exportFolder, exportedFiles

        Set guardSheet = FirstWorkbookSheet(tempWb, DEFAULTPASSWORDSHEET)
        If guardSheet Is Nothing Then
            Set guardSheet = tempWb.Worksheets.Add
            guardSheet.Name = "PwdGuard" & Format$(Timer, "000")
        End If
        guardSheetName = guardSheet.Name

        cloned.Protect guardSheetName, allowShapes:=False, allowDeletingRows:=False, registerState:=False
        cloned.EnterDebugMode tempWb

        Application.EnableEvents = True

        workbookPath = TestHelpers.BuildWorkbookPath(exportFolder, "passwords_debug_roundtrip")
        If Dir$(workbookPath) <> vbNullString Then Kill workbookPath

        'cloned.LeaveDebugMode tempWb

        tempWb.SaveAs Filename:=workbookPath, FileFormat:=xlExcel12
        tempWb.Close SaveChanges:=True

        Set reopened = Workbooks.Open(Filename:=workbookPath)

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

    expectedPrompt = "Password: " & PasswordSubject.Value("privatekey")
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

    expectedPrompt = "Password: " & updatedPrivate
    Assert.AreEqual expectedPrompt, PasswordSubject.LastPrivatePrompt, _
                     "GenerateKey should prepare the private key prompt when prompts are suppressed"
    Assert.AreEqual "Credentials", PasswordSubject.LastPrivatePromptTitle, _
                     "GenerateKey should capture the prompt title when prompts are suppressed"
End Sub

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
