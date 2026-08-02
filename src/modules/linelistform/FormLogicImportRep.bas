Attribute VB_Name = "FormLogicImportRep"
Attribute VB_Description = "Form code-behind for F_ImportRep"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ImportRep")

'@description
'The code behind F_ImportRep. The form shows the report of the last import,
'read straight off the four ListObjects ImportReport keeps on import_rep__.
'
'Nothing is passed in. The report is a worksheet store, so opening the form
'days after the import shows the same four lists, and the advanced form's
'CMD_ImportMigRep button works whether or not an import ran in this session.
'
'The four lists were computed and thrown away for as long as this form has
'existed: the arrays lived in memory and died with the import object, and the
'box that asked the user whether they wanted a report was shown with an OK
'button, so there was nothing to answer.

Option Explicit

Private Const LLSHEET As String = "LinelistTranslation"

Private tradform As TranslationObject


Private Sub InitializeTrads()
    Dim lltrads As LLTranslation

    Set lltrads = LLTranslation.Create(ThisWorkbook.Worksheets(LLSHEET))
    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

'Fill the four list controls from the import report store.
Private Sub LoadReportLists()

    Dim store As ImportReport

    Set store = Nothing
    On Error Resume Next
    Set store = ImportReport.Create(ThisWorkbook)
    On Error GoTo 0
    If store Is Nothing Then Exit Sub

    FillOneColumnList Me.LST_ImpRepSheet, _
                      store.SheetNames(ImportReportNotImported)
    FillOneColumnList Me.LST_ImpLLSheet, _
                      store.SheetNames(ImportReportNotTouched)
    FillTwoColumnList Me.LST_ImpRepVarImp, _
                      store.VariableEntries(ImportReportNotImported)
    FillTwoColumnList Me.LST_ImpRepVarLL, _
                      store.VariableEntries(ImportReportNotTouched)
End Sub

'Put a list of names into one list control.
Private Sub FillOneColumnList(ByVal listControl As Object, ByVal namesList As BetterArray)

    Dim counter As Long

    listControl.Clear
    If namesList Is Nothing Then Exit Sub

    For counter = namesList.LowerBound To namesList.UpperBound
        listControl.AddItem CStr(namesList.Item(counter))
    Next counter
End Sub

'Put a list of variable and sheet pairs into one two-column list control.
Private Sub FillTwoColumnList(ByVal listControl As Object, ByVal entriesList As BetterArray)

    Dim entry As Variant
    Dim counter As Long
    Dim rowIndex As Long

    listControl.Clear
    If entriesList Is Nothing Then Exit Sub

    listControl.ColumnCount = 2

    For counter = entriesList.LowerBound To entriesList.UpperBound
        entry = entriesList.Item(counter)
        listControl.AddItem CStr(entry(LBound(entry)))
        listControl.List(rowIndex, 1) = CStr(entry(LBound(entry) + 1))
        rowIndex = rowIndex + 1
    Next counter
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.Width = 550
    Me.Height = 450

    LoadReportLists
End Sub

Private Sub CMD_ImpRepQuit_Click()
    Me.Hide
End Sub

Private Sub LBL_Previous_Click()
    Me.Hide
    F_Advanced.Show
End Sub
