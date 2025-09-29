Attribute VB_Name = "TestSectionWritersIntegration"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object

Private Const H_BASE As String = "HSectionFixture"
Private Const V_BASE As String = "VSectionFixture"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet H_BASE & "_Data"
    TestHelpers.DeleteWorksheet V_BASE & "_Data"
    TestHelpers.DeleteWorksheet "HSection"
    TestHelpers.DeleteWorksheet "VSection"
End Sub

'@TestMethod("SectionWriters")
Private Sub TestHorizontalSectionWriterFormatsHeaders()
    Dim design As LLFormatLogStub
    Dim context As ILLSectionContext
    Dim writer As ILLSectionWriter

    Set design = New LLFormatLogStub
    Set context = SectionTestFixture.CreateSectionContext(H_BASE, SectionTestFixture.HorizontalSectionRows(), 2, design)

    Set writer = HListSectionWriter.Create(context)
    writer.WriteSection 2, 4

    Assert.AreEqual 1, design.ScopeCount(HListSection), _
                     "Horizontal section header should use HListSection scope"
    Assert.AreEqual 2, design.ScopeCount(HListSubSection), _
                     "Subsections should format headers with HListSubSection scope"
    Assert.AreEqual 1, design.ScopeCount(HListCRFSection), _
                     "CRF section formatting should be applied once"
    Assert.AreEqual 1, design.ScopeCount(HListCRFSubSection), _
                     "CRF subsection formatting should mirror header scopes"

    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets("HSection")
    Assert.AreEqual "Section H", sheet.Cells(5, 4).Value, _
                     "Section header should populate the worksheet"
    Assert.AreEqual "Sub H1", sheet.Cells(6, 4).Value, _
                     "First subsection header should be written"
    Assert.AreEqual "Sub H3", sheet.Cells(6, 8).Value, _
                     "Second subsection header should be written"
End Sub

'@TestMethod("SectionWriters")
Private Sub TestVerticalSectionWriterFormatsHeaders()
    Dim design As LLFormatLogStub
    Dim context As ILLSectionContext
    Dim writer As ILLSectionWriter

    Set design = New LLFormatLogStub
    Set context = SectionTestFixture.CreateSectionContext(V_BASE, SectionTestFixture.VerticalSectionRows(), 2, design)

    Set writer = VListSectionWriter.Create(context)
    writer.WriteSection 2, 3

    Assert.AreEqual 1, design.ScopeCount(VListSection), _
                     "Vertical section header should use VListSection scope"
    Assert.AreEqual 1, design.ScopeCount(VListSubSection), _
                     "Vertical subsection header should use VListSubSection scope"

    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets("VSection")
    Assert.AreEqual "Section V", sheet.Cells(10, 2).Value, _
                     "Vertical section header should be written"
    Assert.AreEqual "Sub V1", sheet.Cells(10, 3).Value, _
                     "Vertical subsection header should be written"
End Sub

