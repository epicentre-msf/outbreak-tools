Attribute VB_Name = "TestDevelopment"

Option Explicit

Private Const DEV_SHEET_NAME As String = "TestDevelopmentDevs"
Private Const CODE_SHEET_NAME As String = "TestDevelopmentCodes"
Private Const NAMED_MODULES As String = "ModulesCodes"
Private Const NAMED_CLASSES As String = "ClassesImplementation"
Private Const NAMED_TESTS As String = "TestsCodes"
Private Const GENERAL_FOLDER As String = "general"
Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Manager As IDevelopment
Private TestBook As Workbook
Private DevSheet As Worksheet
Private CodeSheet As Worksheet

Private TempRoot As String
Private ModulesPath As String
Private ClassesPath As String
Private TestsPath As String


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDevelopment"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults "testsOutputs"
    End If
    Set Assert = Nothing
    RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Private Sub TestInitialize()
    Set TestBook = TestHelpers.NewWorkbook
    Set DevSheet = TestHelpers.EnsureWorksheet(DEV_SHEET_NAME, TestBook)
    Set CodeSheet = TestHelpers.EnsureWorksheet(CODE_SHEET_NAME, TestBook)
    PrepareNamedRanges

    Set Manager = Development.Create(DevSheet, CodeSheet)
    Manager.DisplayPrompts = False
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Manager Is Nothing Then
        Set Manager = Nothing
    End If

    If Not TestBook Is Nothing Then
        TestHelpers.DeleteWorkbook TestBook
        Set TestBook = Nothing
    End If

    CleanupFolder TempRoot
    TempRoot = vbNullString

    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Tests
'===============================================================================
'@TestMethod("Development")
Public Sub TestAddClassTableIncrementsCounters()
    CustomTestSetTitles Assert, "Development", "TestAddClassTableIncrementsCounters"
    On Error GoTo TestFail

    Dim firstTable As ListObject
    Set firstTable = Manager.AddClassTable()

    Dim secondTable As ListObject
    Set secondTable = Manager.AddClassTable()

    Assert.AreEqual "ClassesLo1", firstTable.Name, "First classes table should register as ClassesLo1"
    Assert.AreEqual "ClassesLo2", secondTable.Name, "Second classes table should increment the counter"
    Assert.AreEqual firstTable.Range.Column + firstTable.ListColumns.Count + 1, _
                     secondTable.Range.Column, _
                     "Next classes table should be positioned one column after the previous block"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddClassTableIncrementsCounters", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestAddTableCreatesTestTag()
    CustomTestSetTitles Assert, "Development", "TestAddTableCreatesTestTag"
    On Error GoTo TestFail

    Dim testModules As ListObject
    Set testModules = Manager.AddModuleTable(True)

    Assert.AreEqual "tests modules", LCase$(CStr(testModules.Range.Cells(0, 1).Value)), _
                     "Adding a test modules table should tag it as tests modules"

    Dim classModules As ListObject
    Set classModules = Manager.AddClassTable(True)
    Assert.AreEqual "tests classes", LCase$(CStr(classModules.Range.Cells(0, 1).Value)), _
                    "Adding a class table should tag it as test classes"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddTableCreatesTestTag", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestImportAllLoadsModulesAndInterfaces()
    CustomTestSetTitles Assert, "Development", "TestImportAllLoadsModulesAndInterfaces"
    On Error GoTo TestFail

    PrepareGeneralFolders

    Dim moduleSource As String
    Dim classSource As String
    Dim interfaceSource As String

    moduleSource = JoinPath(ModulesPath, GENERAL_FOLDER, "DevModuleSample.bas")
    classSource = JoinPath(ClassesPath, GENERAL_FOLDER, "DevClassSample.cls")
    interfaceSource = JoinPath(ClassesPath, GENERAL_FOLDER, "IDevClassSample.cls")

    WriteTextFile moduleSource, ModuleSourceCode("DevModuleSample")
    WriteTextFile classSource, ClassSourceCode("DevClassSample")
    WriteTextFile interfaceSource, ClassSourceCode("IDevClassSample")

    Dim modulesTable As ListObject
    Set modulesTable = Manager.AddModuleTable()
    modulesTable.Range.Cells(-1, 1).Value = GENERAL_FOLDER
    modulesTable.ListRows.Add
    modulesTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value = "DevModuleSample"

    Dim classesTable As ListObject
    Set classesTable = Manager.AddClassTable()
    classesTable.Range.Cells(-1, 1).Value = GENERAL_FOLDER
    classesTable.ListRows.Add
    classesTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value = "DevClassSample"
    classesTable.ListColumns(2).DataBodyRange.Cells(1, 1).Value = "Yes"

    RemoveComponentIfExists "DevModuleSample"
    RemoveComponentIfExists "DevClassSample"
    RemoveComponentIfExists "IDevClassSample"

    Manager.ImportAll

    AssertComponentExists "DevModuleSample", "Module should be imported into workbook"
    AssertComponentExists "DevClassSample", "Class should be imported into workbook"
    AssertComponentExists "IDevClassSample", "Interface should be imported when flagged"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestImportAllLoadsModulesAndInterfaces", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestExportAllWritesFiles()
    CustomTestSetTitles Assert, "Development", "TestExportAllWritesFiles"
    On Error GoTo TestFail

    PrepareGeneralFolders

    Dim moduleComponent As Object
    Dim classComponent As Object
    Dim interfaceComponent As Object

    Set moduleComponent = TestBook.VBProject.VBComponents.Add(1)
    DoEvents
    moduleComponent.Name = "ExportModuleSample"
    moduleComponent.CodeModule.AddFromString ModuleSourceCode("ExportModuleSample", False)

    Set classComponent = TestBook.VBProject.VBComponents.Add(2)
    DoEvents
    classComponent.Name = "ExportClassSample"
    classComponent.CodeModule.AddFromString ClassSourceCode("ExportClassSample", False)

    Set interfaceComponent = TestBook.VBProject.VBComponents.Add(2)
    DoEvents
    interfaceComponent.Name = "IExportClassSample"
    interfaceComponent.CodeModule.AddFromString ClassSourceCode("IExportClassSample", False)

    Dim modulesTable As ListObject
    Set modulesTable = Manager.AddModuleTable()
    modulesTable.Range.Cells(-1, 1).Value = GENERAL_FOLDER
    modulesTable.ListRows.Add
    modulesTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value = "ExportModuleSample"

    Dim classesTable As ListObject
    Set classesTable = Manager.AddClassTable()
    classesTable.Range.Cells(-1, 1).Value = GENERAL_FOLDER
    classesTable.ListRows.Add
    classesTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value = "ExportClassSample"
    classesTable.ListColumns(2).DataBodyRange.Cells(1, 1).Value = "Yes"

    Dim moduleTarget As String
    Dim classTarget As String
    Dim interfaceTarget As String

    moduleTarget = JoinPath(ModulesPath, GENERAL_FOLDER, "ExportModuleSample.bas")
    classTarget = JoinPath(ClassesPath, GENERAL_FOLDER, "ExportClassSample.cls")
    interfaceTarget = JoinPath(ClassesPath, GENERAL_FOLDER, "IExportClassSample.cls")

    DeleteFileIfExists moduleTarget
    DeleteFileIfExists classTarget
    DeleteFileIfExists interfaceTarget

    Manager.ExportAll

    Assert.IsTrue FileExists(moduleTarget), "Module export should create .bas file"
    Assert.IsTrue FileExists(classTarget), "Class export should create .cls file"
    Assert.IsTrue FileExists(interfaceTarget), "Interface export should create .cls file"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestExportAllWritesFiles", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestAddFormsCodesCopiesContent()
    CustomTestSetTitles Assert, "Development", "TestAddFormsCodesCopiesContent"
    On Error GoTo TestFail

    Dim sourceComponent As Object
    Dim targetComponent As Object

    Set sourceComponent = TestBook.VBProject.VBComponents.Add(1)
    DoEvents
    sourceComponent.Name = "FormLogicSource"
    sourceComponent.CodeModule.AddFromString "Public Sub Execute()" & vbNewLine & "    Debug.Print ""source""" & vbNewLine & "End Sub"

    Set targetComponent = TestBook.VBProject.VBComponents.Add(3)
    DoEvents
    targetComponent.Name = "FormLogicTarget"
    targetComponent.CodeModule.AddFromString "Public Sub Execute()" & vbNewLine & "    Debug.Print ""target""" & vbNewLine & "End Sub"

    'Capture source code now: CopyModuleCode removes the source component,
    'making the reference stale and its CodeModule inaccessible afterwards.
    Dim expectedCode As String
    expectedCode = sourceComponent.CodeModule.Lines(1, sourceComponent.CodeModule.CountOfLines)

    Dim formsTable As ListObject
    Set formsTable = Manager.AddFormsTable()
    formsTable.ListRows.Add
    formsTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value = "FormLogicSource"
    formsTable.ListColumns(2).DataBodyRange.Cells(1, 1).Value = "FormLogicTarget"

    DoEvents
    Manager.AddFormsCodes

    Assert.AreEqual expectedCode, _
                     targetComponent.CodeModule.Lines(1, targetComponent.CodeModule.CountOfLines), _
                     "Target component should mirror source code after AddFormsCodes"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddFormsCodesCopiesContent", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestTablesFallbackToDevSheetWhenCodeWorksheetMissing()
    CustomTestSetTitles Assert, "Development", "TestTablesFallbackToDevSheetWhenCodeWorksheetMissing"
    On Error GoTo TestFail

    RemoveSheetName DevSheet, "Development_CodeSheet"
    Set Manager = Development.Create(DevSheet)
    Manager.DisplayPrompts = False

    Dim fallbackTable As ListObject
    Set fallbackTable = Manager.AddModuleTable()

    Assert.AreEqual DevSheet.Name, fallbackTable.Parent.Name, _
                    "When no code worksheet is registered, tables should be created on the Dev sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTablesFallbackToDevSheetWhenCodeWorksheetMissing", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestAddCodeSheetsRegistersWorksheet()
    CustomTestSetTitles Assert, "Development", "TestAddCodeSheetsRegistersWorksheet"
    On Error GoTo TestFail

    RemoveSheetName DevSheet, "Development_CodeSheet"
    Set Manager = Development.Create(DevSheet)
    Manager.DisplayPrompts = False

    Dim registered As Worksheet
    Set registered = Manager.AddCodeSheets(CODE_SHEET_NAME)

    Assert.IsNotNothing registered, "AddCodeSheets should return the registered worksheet"
    Assert.AreEqual CODE_SHEET_NAME, registered.Name, "Code worksheet should match requested name"
    Assert.AreEqual CODE_SHEET_NAME, Manager.CodeWorksheet.Name, "Manager should retain registered code worksheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAddCodeSheetsRegistersWorksheet", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestDeployHidesCodeSheetAndSetsFlag()
    CustomTestSetTitles Assert, "Development", "TestDeployHidesCodeSheetAndSetsFlag"
    On Error GoTo TestFail

    Dim sourceComponent As Object
    Dim targetComponent As Object

    Set sourceComponent = TestBook.VBProject.VBComponents.Add(1)
    DoEvents
    sourceComponent.Name = "DeploySource"
    sourceComponent.CodeModule.AddFromString "Public Sub Execute()" & vbNewLine & _
                                           "    Debug.Print ""deploy source""" & vbNewLine & _
                                           "End Sub"

    Set targetComponent = TestBook.VBProject.VBComponents.Add(2)
    DoEvents
    targetComponent.Name = "DeployTarget"
    targetComponent.CodeModule.AddFromString "Public Sub Execute()" & vbNewLine & _
                                           "    Debug.Print ""deploy target""" & vbNewLine & _
                                           "End Sub"

    'Capture source code now: Deploy calls CopyModuleCode which removes the
    'source component, making the reference stale afterwards.
    Dim expected As String
    expected = sourceComponent.CodeModule.Lines(1, sourceComponent.CodeModule.CountOfLines)

    Dim formsTable As ListObject
    Set formsTable = Manager.AddFormsTable()
    formsTable.ListRows.Add
    formsTable.ListColumns(1).DataBodyRange.Cells(1, 1).Value = "DeploySource"
    formsTable.ListColumns(2).DataBodyRange.Cells(1, 1).Value = "DeployTarget"

    Manager.AddProtectedSheet DevSheet.Name

    Dim pass As IPasswords
    Set pass = New LinelistPasswordStub

    DoEvents
    Manager.Deploy pass

    Assert.AreEqual expected, _
                     targetComponent.CodeModule.Lines(1, targetComponent.CodeModule.CountOfLines), _
                     "Deploy should synchronise form modules before protecting"

    Assert.AreEqual xlSheetVeryHidden, CodeSheet.Visible, _
                     "Deploy should hide the registered code worksheet"

    Dim deploymentName As Name
    Set deploymentName = TestBook.Names("inDeployment")
    Assert.AreEqual "=""Yes""", deploymentName.RefersTo, _
                     "Deploy should mark workbook as in deployment via name value"
    Assert.IsTrue Manager.InDeployment, "InDeployment helper should reflect workbook flag after deployment"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDeployHidesCodeSheetAndSetsFlag", Err.Number, Err.Description
End Sub

'@TestMethod("Development")
Public Sub TestInDeploymentFlag()
    CustomTestSetTitles Assert, "Development", "TestInDeploymentFlag"
    On Error GoTo TestFail

    RemoveWorkbookName "inDeployment"
    Assert.IsFalse Manager.InDeployment, "InDeployment should be False when workbook flag is absent"

    TestBook.Names.Add Name:="inDeployment", RefersTo:="=""Yes"""
    Assert.IsTrue Manager.InDeployment, "InDeployment should detect workbook flag value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestInDeploymentFlag", Err.Number, Err.Description
End Sub


'@section Helpers
'===============================================================================
Private Sub PrepareNamedRanges()
    TempRoot = TestHelpers.BuildTempFolder(ThisWorkbook, "DevelopmentTests")

    ModulesPath = JoinPath(TempRoot, "src", "modules")
    ClassesPath = JoinPath(TempRoot, "src", "classes")
    TestsPath = JoinPath(TempRoot, "src", "tests")

    TestHelpers.EnsureFolder ModulesPath
    TestHelpers.EnsureFolder ClassesPath
    TestHelpers.EnsureFolder TestsPath

    BindNamedRange DevSheet, NAMED_MODULES, DevSheet.Range("A1"), ModulesPath
    BindNamedRange DevSheet, NAMED_CLASSES, DevSheet.Range("A2"), ClassesPath
    BindNamedRange DevSheet, NAMED_TESTS, DevSheet.Range("A3"), TestsPath
End Sub

Private Sub PrepareGeneralFolders()
    TestHelpers.EnsureFolder JoinPath(ModulesPath, GENERAL_FOLDER)
    TestHelpers.EnsureFolder JoinPath(ClassesPath, GENERAL_FOLDER)
    TestHelpers.EnsureFolder JoinPath(TestsPath, "modules")
    TestHelpers.EnsureFolder JoinPath(TestsPath, "classes")
End Sub

Private Sub BindNamedRange(ByVal sheet As Worksheet, _
                           ByVal nameId As String, _
                           ByVal targetCell As Range, _
                           ByVal assignedValue As String)

    RemoveSheetName sheet, nameId
    sheet.Names.Add Name:=nameId, _
                    RefersTo:="=" & targetCell.Address(True, True, xlA1, True)
    targetCell.Value = assignedValue
End Sub

Private Sub RemoveSheetName(ByVal sheet As Worksheet, ByVal nameId As String)
    Dim idx As Long

    For idx = sheet.Names.Count To 1 Step -1
        If StrComp(sheet.Names(idx).Name, sheet.Name & "!" & nameId, vbTextCompare) = 0 _
           Or StrComp(sheet.Names(idx).Name, nameId, vbTextCompare) = 0 Then
            sheet.Names(idx).Delete
        End If
    Next idx
End Sub

Private Sub RemoveWorkbookName(ByVal nameId As String)
    If TestBook Is Nothing Then Exit Sub

    Dim idx As Long
    For idx = TestBook.Names.Count To 1 Step -1
        If StrComp(TestBook.Names(idx).Name, nameId, vbTextCompare) = 0 Then
            TestBook.Names(idx).Delete
        End If
    Next idx
End Sub

Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim fileNum As Integer
    TestHelpers.EnsureFolder TestHelpers.ParentFolder(filePath)
    fileNum = FreeFile()
    Open filePath For Output As #fileNum
        Print #fileNum, content
    Close #fileNum
End Sub

'Returns source code for a .bas module. When includeHeaders is True (default),
'the Attribute VB_Name header is prepended for writing to disk. When False,
'only the code body is returned for use with CodeModule.AddFromString.
Private Function ModuleSourceCode(ByVal moduleName As String, _
                                  Optional ByVal includeHeaders As Boolean = True) As String
    Dim body As String
    body = "Option Explicit" & vbNewLine & _
           "Public Sub Execute()" & vbNewLine & _
           "    Debug.Print ""module""" & vbNewLine & _
           "End Sub"

    If includeHeaders Then
        ModuleSourceCode = "Attribute VB_Name = """ & moduleName & """" & vbNewLine & body
    Else
        ModuleSourceCode = body
    End If
End Function

'Returns source code for a .cls class. When includeHeaders is True (default),
'the VERSION/BEGIN/END block and Attribute headers are prepended for writing to
'disk. When False, only the code body is returned for use with
'CodeModule.AddFromString.
Private Function ClassSourceCode(ByVal className As String, _
                                 Optional ByVal includeHeaders As Boolean = True) As String
    Dim body As String
    body = "Option Explicit" & vbNewLine & _
           "Public Sub Execute()" & vbNewLine & _
           "    Debug.Print ""class""" & vbNewLine & _
           "End Sub"

    If includeHeaders Then
        ClassSourceCode = "VERSION 1.0 CLASS" & vbNewLine & _
                          "BEGIN" & vbNewLine & _
                          "  MultiUse = -1  'True" & vbNewLine & _
                          "END" & vbNewLine & _
                          "Attribute VB_Name = """ & className & """" & vbNewLine & _
                          "Attribute VB_GlobalNameSpace = False" & vbNewLine & _
                          "Attribute VB_Creatable = False" & vbNewLine & _
                          "Attribute VB_PredeclaredId = False" & vbNewLine & _
                          "Attribute VB_Exposed = False" & vbNewLine & _
                          body
    Else
        ClassSourceCode = body
    End If
End Function

Private Sub RemoveComponentIfExists(ByVal componentName As String)
    Dim vbProj As Object
    Dim components As Object

    Set vbProj = TestBook.VBProject
    Set components = vbProj.VBComponents

    On Error Resume Next
        components.Remove components(componentName)
    On Error GoTo 0
End Sub

Private Sub AssertComponentExists(ByVal componentName As String, ByVal messageText As String)
    Dim component As Object

    On Error Resume Next
        Set component = TestBook.VBProject.VBComponents(componentName)
    On Error GoTo 0

    Assert.IsFalse component Is Nothing, messageText
End Sub

Private Sub DeleteFileIfExists(ByVal filePath As String)
    If LenB(filePath) = 0 Then Exit Sub
    On Error Resume Next
        If Dir$(filePath) <> vbNullString Then Kill filePath
    On Error GoTo 0
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    If LenB(filePath) = 0 Then Exit Function
    FileExists = (Dir$(filePath) <> vbNullString)
End Function

Private Sub CleanupFolder(ByVal folderPath As String)
    If LenB(folderPath) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) = vbNullString Then Exit Sub

    Dim sep As String
    sep = Application.PathSeparator

    Dim entry As String
    entry = Dir$(folderPath & sep & "*", vbDirectory Or vbNormal Or vbHidden Or vbSystem)

    Do While LenB(entry) > 0
        If entry <> "." And entry <> ".." Then
            Dim fullPath As String
            fullPath = folderPath & sep & entry

            If (GetAttr(fullPath) And vbDirectory) = vbDirectory Then
                CleanupFolder fullPath
            Else
                On Error Resume Next
                    Kill fullPath
                On Error GoTo 0
            End If
        End If

        entry = vbNullString

        On Error Resume Next
            entry = Dir$()
        On Error GoTo 0
    Loop

    On Error Resume Next
        RmDir folderPath
    On Error GoTo 0
End Sub
