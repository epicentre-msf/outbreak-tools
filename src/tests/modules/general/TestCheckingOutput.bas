Attribute VB_Name = "TestCheckingOutput"

Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@TestModule
'@Folder("Tests")

Private Const DEFAULTCHECKINGSHEET As String = "CheckingOutputFixture"
Private Const DEFAULTFILTERCELL As String = "C1"
Private Const DEFAULTEVENTMAKER As String = "CheckingOutputEventInstalled"

Private Assert As Object
Private Fakes As Object
Private OutputWriter As ICheckingOutput
Private PrimaryCheck As IChecking
Private SecondaryCheck As IChecking

'@section Helpers
'===============================================================================

Private Function BuildChecking(ByVal heading As String, ByVal subHeading As String, _
                               ParamArray entries() As Variant) As IChecking
    Dim checkingInstance As IChecking
    Dim index As Long

    Set checkingInstance = Checking.Create(heading, subHeading)
    For index = LBound(entries) To UBound(entries)
        checkingInstance.Add entries(index)(0), entries(index)(1), entries(index)(2)
    Next index

    Set BuildChecking = checkingInstance
End Function

Private Function OutputSheet() As Worksheet
    Set OutputSheet = EnsureWorksheet(DEFAULTCHECKINGSHEET)
End Function

Private Function CountOccurrences(ByVal sheet As Worksheet, ByVal textValue As String) As Long
    Dim usedRange As Range
    Dim cell As Range

    Set usedRange = sheet.UsedRange
    For Each cell In usedRange
        If StrComp(CStr(cell.Value), textValue, vbTextCompare) = 0 Then
            CountOccurrences = CountOccurrences + 1
        End If
    Next cell
End Function

Private Function GetCustomPropertyValue(ByVal sheet As Worksheet, ByVal propertyName As String) As String
    Dim prop As CustomProperty

    For Each prop In sheet.CustomProperties
        If StrComp(prop.Name, propertyName, vbTextCompare) = 0 Then
            GetCustomPropertyValue = prop.Value
            Exit Function
        End If
    Next prop
End Function

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet DEFAULTCHECKINGSHEET
    RestoreApp
    Application.EnableEvents = True
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ClearWorksheet OutputSheet
    Set OutputWriter = CheckingOutput.Create(OutputSheet)
    Set PrimaryCheck = BuildChecking("Data Validation Summary", "Core checks", _
                                     Array("key-1", "Missing identifier", checkingError), _
                                     Array("key-2", "Inconsistent dates", checkingWarning))
    Set SecondaryCheck = BuildChecking("Data Validation Summary", "Extended checks", _
                                       Array("key-3", "Optional comment", checkingNote), _
                                       Array("key-4", "Informative remark", checkingInfo))
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set OutputWriter = Nothing
    Set PrimaryCheck = Nothing
    Set SecondaryCheck = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("CheckingOutput")
Private Sub TestPrintOutputWritesRows()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sheet As Worksheet

    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.PrintOutput checks

    Set sheet = OutputWriter.Worksheet

    Assert.AreEqual "Show only:", sheet.Cells(1, 2).Value, "Header cell should be initialised"
    Assert.AreEqual "All", sheet.Range(DEFAULTFILTERCELL).Value, "Filter should default to All"
    Assert.AreEqual xlValidateList, sheet.Range(DEFAULTFILTERCELL).Validation.Type, "Filter cell should contain list validation"
    Assert.AreEqual "True", GetCustomPropertyValue(sheet, DEFAULTEVENTMAKER), "Worksheet change handler marker should be stored"
    Assert.AreEqual 1, CountOccurrences(sheet, "Data Validation Summary"), "Title should be written only once"
    Assert.AreEqual "Core checks", sheet.Cells(3, 2).Value, "Subtitle should follow the title"
    Assert.AreEqual "Error", sheet.Cells(4, 2).Value, "First data row should include type caption"
    Assert.AreEqual "Missing identifier", sheet.Cells(4, 3).Value, "First data row should include label"
    Assert.AreEqual RGB(255, 0, 0), sheet.Cells(4, 2).Font.Color, "Error rows should use red font"
    Assert.AreEqual "Informative remark", sheet.Cells(sheet.UsedRange.Rows.Count, 3).Value, "Last row should match final entry"
    Assert.AreEqual RGB(112, 48, 160), sheet.Cells(7, 2).Font.Color, "Note rows should use purple font"
    Assert.AreEqual RGB(244, 236, 255), sheet.Cells(7, 2).Interior.Color, "Note rows should use purple fill"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestPrintOutputWritesRows"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestPrintOutputRejectsInvalidItems()
    Dim invalidChecks As BetterArray
    Dim raisedError As Boolean

    Set invalidChecks = BetterArrayFromList("invalid entry")

    On Error Resume Next
        OutputWriter.PrintOutput invalidChecks
        raisedError = (Err.Number <> 0)
    On Error GoTo 0

    Assert.IsTrue raisedError, "PrintOutput should raise when items are not IChecking"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestFilterDropdownHidesRows()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sheet As Worksheet

    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.PrintOutput checks

    Set sheet = OutputWriter.Worksheet
    Application.EnableEvents = True

    sheet.Range(DEFAULTFILTERCELL).Value = "Warnings"
    Assert.IsTrue sheet.Rows(4).Hidden, "Rows with non matching types should be hidden"
    Assert.IsFalse sheet.Rows(5).Hidden, "Rows with matching type should remain visible"
    Assert.IsFalse sheet.Rows(3).Hidden, "Subtitles should always be visible"

    sheet.Range(DEFAULTFILTERCELL).Value = "All"
    Assert.IsFalse sheet.Rows(4).Hidden, "All rows should be visible after resetting filter"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFilterDropdownHidesRows"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestFilterWorksheetMethodAcceptsExplicitSelection()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sheet As Worksheet

    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.PrintOutput checks

    Set sheet = OutputWriter.Worksheet
    OutputWriter.FilterWorksheet "Warnings"

    Assert.IsTrue sheet.Rows(4).Hidden, "Explicit warning filter should hide error rows"
    Assert.IsFalse sheet.Rows(5).Hidden, "Explicit warning filter should keep warning rows visible"
    Assert.IsTrue sheet.Rows(7).Hidden, "Explicit warning filter should hide note rows"
    Assert.IsTrue sheet.Rows(8).Hidden, "Explicit warning filter should hide info rows"
    Assert.IsFalse sheet.Rows(6).Hidden, "Section headers should remain visible"

    OutputWriter.FilterWorksheet "All"
    Assert.IsFalse sheet.Rows(4).Hidden, "All filter should reveal previously hidden rows"
    Assert.IsFalse sheet.Rows(5).Hidden, "All filter should keep warning rows visible"
    Assert.IsFalse sheet.Rows(7).Hidden, "All filter should restore note rows"
    Assert.IsFalse sheet.Rows(8).Hidden, "All filter should restore info rows"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFilterWorksheetMethodAcceptsExplicitSelection"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestFilterWorksheetUsesCellValueWhenParameterMissing()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sheet As Worksheet
    Dim previousEventState As Boolean

    previousEventState = Application.EnableEvents
    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.PrintOutput checks

    Set sheet = OutputWriter.Worksheet
    Application.EnableEvents = False
    sheet.Range(DEFAULTFILTERCELL).Value = "Notes"
    Application.EnableEvents = previousEventState

    OutputWriter.FilterWorksheet

    Assert.IsTrue sheet.Rows(4).Hidden, "Notes filter should hide error rows"
    Assert.IsTrue sheet.Rows(5).Hidden, "Notes filter should hide warning rows"
    Assert.IsFalse sheet.Rows(7).Hidden, "Notes filter should keep note rows visible"
    Assert.IsTrue sheet.Rows(8).Hidden, "Notes filter should hide info rows"
    Assert.IsFalse sheet.Rows(6).Hidden, "Section headers should remain visible when filtering by notes"
    Exit Sub

Fail:
    Application.EnableEvents = previousEventState
    FailUnexpectedError Assert, "TestFilterWorksheetUsesCellValueWhenParameterMissing"
End Sub
