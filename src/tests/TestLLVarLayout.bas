Attribute VB_Name = "TestLLVarLayout"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object

Private Const HORIZONTAL_SHEET As String = "LLVarLayoutHorizontal"
Private Const HORIZONTAL_PRINT_SHEET As String = "LLVarLayoutHorizontalPrint"
Private Const VERTICAL_SHEET As String = "LLVarLayoutVertical"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet HORIZONTAL_SHEET
    TestHelpers.DeleteWorksheet HORIZONTAL_PRINT_SHEET
    TestHelpers.DeleteWorksheet VERTICAL_SHEET
    Set Assert = Nothing
End Sub

'@TestMethod("LLVarLayout")
Private Sub TestHorizontalLayoutPositions()
    Dim baseSheet As Worksheet
    Dim printSheet As Worksheet
    Dim layout As ILLVarLayout

    Set baseSheet = TestHelpers.EnsureWorksheet(HORIZONTAL_SHEET)
    Set printSheet = TestHelpers.EnsureWorksheet(HORIZONTAL_PRINT_SHEET)

    Set layout = New LLVarLayoutHorizontal
    layout.Configure baseSheet, 3, printSheet

    Assert.AreEqual "$C$9", layout.ValueCell.Address(True, True), _
                     "Horizontal layout should anchor values on row 9"
    Assert.AreEqual "$C$7", layout.LabelCell.Address(True, True), _
                     "Horizontal layout should place the main label two rows above"
    Assert.AreEqual "$C$8", layout.NameCell.Address(True, True), _
                     "Horizontal layout should place the variable name one row above"
    Assert.AreEqual "$C$4", layout.ControlCell.Address(True, True), _
                     "Horizontal layout control metadata resides five rows above"
    Assert.AreEqual "$C$3", layout.AutoOriginCell.Address(True, True), _
                     "Horizontal layout auto origin metadata resides six rows above"

    Assert.IsTrue layout.SupportsPrintedLayout, _
                  "Horizontal layout should advertise printed support when provided"
    Assert.AreEqual "$C$9", layout.PrintedValueCell.Address(True, True), _
                     "Printed layout should mirror the base column"
End Sub

'@TestMethod("LLVarLayout")
Private Sub TestHorizontalLayoutResetClearsState()
    Dim baseSheet As Worksheet
    Dim layout As ILLVarLayout

    Set baseSheet = TestHelpers.EnsureWorksheet(HORIZONTAL_SHEET)
    Set layout = New LLVarLayoutHorizontal
    layout.Configure baseSheet, 2

    layout.Reset

    Dim errNumber As Long
    On Error Resume Next
        Dim result As String
        result = layout.LayoutKey
        result = layout.ValueCell.Address(False, False)
        errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual ProjectError.ObjectNotInitialized, errNumber, _
                     "Accessing value cell after reset should raise an initialisation error"
End Sub

'@TestMethod("LLVarLayout")
Private Sub TestVerticalLayoutPositions()
    Dim baseSheet As Worksheet
    Dim layout As ILLVarLayout

    Set baseSheet = TestHelpers.EnsureWorksheet(VERTICAL_SHEET)

    Set layout = New LLVarLayoutVertical
    layout.Configure baseSheet, 12

    Assert.AreEqual "$E$12", layout.ValueCell.Address(True, True), _
                     "Vertical layout should anchor values on column E"
    Assert.AreEqual "$D$12", layout.LabelCell.Address(True, True), _
                     "Vertical layout should use the previous column for labels"
    Assert.AreEqual "$E$12", layout.NameCell.Address(True, True), _
                     "Vertical layout keeps the variable name on the anchor cell"
    Assert.AreEqual "$F$12", layout.ControlCell.Address(True, True), _
                     "Vertical layout stores control metadata one column to the right"
    Assert.IsFalse layout.SupportsPrintedLayout, _
                   "Vertical layout should not expose a printed variant"
End Sub
