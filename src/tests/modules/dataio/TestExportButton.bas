Attribute VB_Name = "TestExportButton"
Attribute VB_Description = "Unit tests for ExportButton"

'@Folder("Tests.DataIO")
'@ModuleDescription("Unit tests for ExportButton")
'@TestModule
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument

Option Explicit
Option Private Module

Private Assert As ICustomTest
Private testSheet As Worksheet


'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CustomTest.Create()
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    '
End Sub

'@TestCleanup
Public Sub TestCleanup()
    CleanupTestSheet
End Sub


'@section Helpers
'===============================================================================

Private Function CreateTestSheet() As Worksheet
    Set testSheet = ThisWorkbook.Worksheets.Add
    Set CreateTestSheet = testSheet
End Function

Private Sub CleanupTestSheet()
    If Not testSheet Is Nothing Then
        Application.DisplayAlerts = False
        testSheet.Delete
        Application.DisplayAlerts = True
        Set testSheet = Nothing
    End If
End Sub

Private Function CreateButton(ByVal sh As Worksheet, _
                               ByVal buttonName As String) As MSForms.CommandButton
    Dim ole As OLEObject
    Set ole = sh.OLEObjects.Add(ClassType:="Forms.CommandButton.1")
    Dim btn As MSForms.CommandButton
    Set btn = ole.Object
    btn.Name = buttonName
    Set CreateButton = btn
End Function

Private Function CreateCheckBox(ByVal sh As Worksheet) As MSForms.CheckBox
    Dim ole As OLEObject
    Set ole = sh.OLEObjects.Add(ClassType:="Forms.CheckBox.1")
    Set CreateCheckBox = ole.Object
End Function

Private Function CreateTranslationStub() As ITranslationObject
    Dim stub As New LinelistSpecsTranslationStub
    stub.Initialise "ExportTestStub"
    Set CreateTranslationStub = stub
End Function


'@section Factory Validation
'===============================================================================

'@TestMethod("ExportButton")
Public Sub FactoryCreatesWithValidArgs()
    CustomTestSetTitles Assert, "ExportButton", "FactoryCreatesWithValidArgs"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.IsNotNothing sut, "Factory should return a valid object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesWithValidArgs", Err.Number, Err.Description
End Sub

'@TestMethod("ExportButton")
Public Sub FactoryRejectsNothingWorkbook()
    CustomTestSetTitles Assert, "ExportButton", "FactoryRejectsNothingWorkbook"
    On Error GoTo TestFail

    Dim sut As ExportButton
    On Error Resume Next
    Set sut = ExportButton.Create(Nothing, CreateTranslationStub(), Nothing)
    Assert.IsTrue Err.Number <> 0, "Should raise error for Nothing workbook"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingWorkbook", Err.Number, Err.Description
End Sub

'@TestMethod("ExportButton")
Public Sub FactoryRejectsNothingTranslations()
    CustomTestSetTitles Assert, "ExportButton", "FactoryRejectsNothingTranslations"
    On Error GoTo TestFail

    Dim sut As ExportButton
    On Error Resume Next
    Set sut = ExportButton.Create(ThisWorkbook, Nothing, Nothing)
    Assert.IsTrue Err.Number <> 0, "Should raise error for Nothing translations"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingTranslations", Err.Number, Err.Description
End Sub

'@TestMethod("ExportButton")
Public Sub FactoryRejectsNothingButton()
    CustomTestSetTitles Assert, "ExportButton", "FactoryRejectsNothingButton"
    On Error GoTo TestFail

    Dim sut As ExportButton
    On Error Resume Next
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), Nothing)
    Assert.IsTrue Err.Number <> 0, "Should raise error for Nothing button"
    On Error GoTo 0

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryRejectsNothingButton", Err.Number, Err.Description
End Sub


'@section ExportNumber
'===============================================================================

'@TestMethod("ExportButton")
Public Sub ExportNumberParsesButtonName()
    CustomTestSetTitles Assert, "ExportButton", "ExportNumberParsesButtonName"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport3")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.AreEqual 3&, sut.ExportNumber, _
                    "ExportNumber should parse '3' from CMDExport3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "ExportNumberParsesButtonName", Err.Number, Err.Description
End Sub


'@section UseFilter
'===============================================================================

'@TestMethod("ExportButton")
Public Sub UseFilterFalseWithoutCheckbox()
    CustomTestSetTitles Assert, "ExportButton", "UseFilterFalseWithoutCheckbox"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)
    Assert.IsFalse sut.UseFilter, _
                   "UseFilter should be False when no checkbox bound"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterFalseWithoutCheckbox", Err.Number, Err.Description
End Sub

'@TestMethod("ExportButton")
Public Sub UseFilterReadsCheckboxValue()
    CustomTestSetTitles Assert, "ExportButton", "UseFilterReadsCheckboxValue"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim chk As MSForms.CheckBox
    Set chk = CreateCheckBox(sh)
    chk.Value = True

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn, chk)
    Assert.IsTrue sut.UseFilter, _
                  "UseFilter should reflect checkbox True value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterReadsCheckboxValue", Err.Number, Err.Description
End Sub

'@TestMethod("ExportButton")
Public Sub UseFilterLetUpdatesCheckbox()
    CustomTestSetTitles Assert, "ExportButton", "UseFilterLetUpdatesCheckbox"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport1")

    Dim chk As MSForms.CheckBox
    Set chk = CreateCheckBox(sh)
    chk.Value = True

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn, chk)

    sut.UseFilter = False
    Assert.IsFalse chk.Value, _
                   "Setting UseFilter to False should uncheck the checkbox"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "UseFilterLetUpdatesCheckbox", Err.Number, Err.Description
End Sub


'@section Interface
'===============================================================================

'@TestMethod("ExportButton")
Public Sub InterfaceExposesExportNumber()
    CustomTestSetTitles Assert, "ExportButton", "InterfaceExposesExportNumber"
    On Error GoTo TestFail

    Dim sh As Worksheet
    Set sh = CreateTestSheet()

    Dim btn As MSForms.CommandButton
    Set btn = CreateButton(sh, "CMDExport2")

    Dim sut As ExportButton
    Set sut = ExportButton.Create(ThisWorkbook, CreateTranslationStub(), btn)

    Dim iface As IExportButton
    Set iface = sut
    Assert.AreEqual 2&, iface.ExportNumber, _
                    "IExportButton.ExportNumber should delegate to ExportNumber"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "InterfaceExposesExportNumber", Err.Number, Err.Description
End Sub
