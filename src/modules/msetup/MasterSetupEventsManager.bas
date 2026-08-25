Attribute VB_Name = "MasterSetupEventsManager"
Option Explicit

'@Folder("Msetup")
'@ModuleDescription("The one place master setup events are handled and shared services live.")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, ExcelMemberMayReturnNothing, UseMeaningfulName

'The workbook code-behind (EventMasterSetupWorkbook) forwards every event
'here, the way the setup workbook forwards to EventsManager. This module owns
'one EventMasterSetup service for the life of the workbook, guards the
'Application state around its own writes, and keeps the worksheet function
'caches honest.
'
'What the change handlers do:
'- Variables sheet: a pick in the Default Choice column fills the Choices
'  Values cell of the line, joined with " | ", in the language the labels are
'  typed in (the first translation language).
'- Disease sheet: a pick in the Variable Name column fills the section, the
'  default choice and the default status of the line. The label and the
'  choice values follow through the two worksheet functions.
'- A pick in the language cell of a disease sheet is stored in the sheet's
'  hidden names, so exports read the language the user chose.
'- An edit on the Translations sheet drops the translation caches.

'Column order of a disease table, as DiseaseSheet builds it.
Private Const DISEASE_NAME_COLUMN As Long = 2
Private Const DISEASE_SECTION_COLUMN As Long = 3
Private Const DISEASE_CHOICE_COLUMN As Long = 5
Private Const DISEASE_STATUS_COLUMN As Long = 7

Private Const LANGUAGE_CELL_ROW As Long = 2
Private Const LANGUAGE_CELL_COLUMN As Long = 2
Private Const NAME_DISLANG As String = "__Var_DISLANG"

Private Const HEADER_VARIABLE_NAME As String = "Variable Name"
Private Const HEADER_DEFAULT_CHOICE As String = "Default Choice"
Private Const CHOICES_DROPDOWN As String = "__lst_choices"

Private eventService As EventMasterSetup

'@section Service
'===============================================================================

'@sub-title The one EventMasterSetup service of this workbook, made on demand.
Public Function MasterSetupService() As EventMasterSetup
    If eventService Is Nothing Then
        Set eventService = EventMasterSetup.Create(ThisWorkbook)
    End If
    Set MasterSetupService = eventService
End Function

'@sub-title Drop the service and every cache derived from it.
Public Sub DisposeMasterSetup()
    Set eventService = Nothing
    ResetMasterSetupFunctionCaches
End Sub

'@section Workbook events
'===============================================================================

'@sub-title Prepare the workbook once when it opens.
Public Sub MsWorkbookOpened()
    ResetMasterSetupFunctionCaches
    MasterSetupService.OnWorkbookOpen Application
End Sub

'@sub-title Remember the ribbon for later invalidations.
Public Sub RibbonLoaded(ByVal ribbon As IRibbonUI)
    MasterSetupService.OnRibbonLoad ribbon
End Sub

'@sub-title Route a committed edit to the sheet's own handler.
Public Sub MsSheetChanged(ByVal sh As Worksheet, ByVal target As Range)
    Dim scope As ApplicationState

    If sh Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error GoTo Handler

    Set scope = ApplicationState.Create(Application)
    scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    If MasterSetupHelpers.IsMasterDiseaseSheet(sh) Then
        HandleDiseaseSheetChange sh, target
    ElseIf SheetIs(sh, "vars") Then
        HandleVariablesSheetChange sh, target
    ElseIf SheetIs(sh, "choi") Then
        HandleChoicesSheetChange sh, target
        'The choice values feed the worksheet functions; their caches are
        'stale now.
        ResetMasterSetupFunctionCaches
    ElseIf SheetIs(sh, "trans") Then
        ResetMasterSetupFunctionCaches
        MasterSetupService.RefreshTranslations
    End If

Cleanup:
    'Shielded: Handler is still armed here, and a raise from Restore
    'would come straight back to this label and raise again.
    On Error Resume Next
    If Not scope Is Nothing Then scope.Restore
    Exit Sub

Handler:
    Debug.Print "SheetChanged - "; sh.Name; " error "; Err.Number; " "; Err.Description
    Resume Cleanup
End Sub

'@section Sheet handlers
'===============================================================================

'@sub-title Fill the line of a disease table when its variable is picked.
Private Sub HandleDiseaseSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    Dim table As ListObject
    Dim nameCells As Range
    Dim changedCell As Range
    Dim languageCell As Range

    'The language pick of the sheet is stored beside it, so exports read the
    'language the user chose.
    Set languageCell = sh.Cells(LANGUAGE_CELL_ROW, LANGUAGE_CELL_COLUMN)
    If Not Intersect(target, languageCell) Is Nothing Then
        StoreDiseaseLanguage sh, CStr(languageCell.Value)
    End If

    If sh.ListObjects.Count = 0 Then Exit Sub
    Set table = sh.ListObjects(1)
    If table.DataBodyRange Is Nothing Then Exit Sub

    Set nameCells = Intersect(target, table.DataBodyRange.Columns(DISEASE_NAME_COLUMN))
    If nameCells Is Nothing Then Exit Sub

    MasterSetupHelpers.UnProtectMasterSetupSheet sh, "disease"

    For Each changedCell In nameCells.Cells
        FillDiseaseLine table, changedCell
    Next changedCell

    MasterSetupHelpers.ProtectMasterSetupSheet sh, "disease"
End Sub

'@sub-title Fill one line of the disease table from the Variables sheet.
'@details A cleared name cleans the whole line: the filled cells empty and
'the choice dropdown the fill added goes with them. A picked name fills the
'section, the default choice and the default status, and puts the choices
'dropdown on the Choice cell so another choice is one pick away.
Private Sub FillDiseaseLine(ByVal table As ListObject, ByVal nameCell As Range)
    Dim variables As MasterSetupVariables
    Dim choiceCell As Range
    Dim variableName As String
    Dim rowIndex As Long

    rowIndex = nameCell.Row - table.DataBodyRange.Row + 1
    variableName = Trim$(CStr(nameCell.Value))
    Set choiceCell = table.DataBodyRange.Cells(rowIndex, DISEASE_CHOICE_COLUMN)

    If LenB(variableName) = 0 Then
        table.DataBodyRange.Cells(rowIndex, DISEASE_SECTION_COLUMN).Value = vbNullString
        choiceCell.Value = vbNullString
        table.DataBodyRange.Cells(rowIndex, DISEASE_STATUS_COLUMN).Value = vbNullString
        'A cell with no validation raises on Delete; the clean line matters
        'more than the raise.
        On Error Resume Next
        choiceCell.Validation.Delete
        On Error GoTo 0
        Exit Sub
    End If

    Set variables = MasterSetupService.Variables
    If variables Is Nothing Then Exit Sub

    'An unknown name fills nothing and clears nothing: the user may still be
    'typing it.
    On Error Resume Next
    table.DataBodyRange.Cells(rowIndex, DISEASE_SECTION_COLUMN).Value = variables.SectionFor(variableName)
    choiceCell.Value = variables.DefaultChoiceFor(variableName)
    table.DataBodyRange.Cells(rowIndex, DISEASE_STATUS_COLUMN).Value = variables.DefaultStatusFor(variableName)
    MasterSetupService.Dropdowns.SetValidation choiceCell, CHOICES_DROPDOWN
    On Error GoTo 0
End Sub

'@sub-title Recalculate the label column when a translated label changes.
'@details The same rule the setup applies on its Choices sheet: the label
'column carries formulas over the translated and untranslated labels, and a
'committed edit in the translated column recalculates them.
Private Sub HandleChoicesSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    Dim wrapper As CustomTable
    Dim translatedRange As Range
    Dim labelRange As Range

    If target.ListObject Is Nothing Then Exit Sub

    Set wrapper = CustomTable.Create(target.ListObject)

    Set translatedRange = wrapper.DataRange(colName:="translated label", includeHeaders:=False, strictSearch:=True, matchCase:=False)
    If translatedRange Is Nothing Then Exit Sub
    If Intersect(target, translatedRange) Is Nothing Then Exit Sub

    Set labelRange = wrapper.DataRange(colName:="label", includeHeaders:=False, strictSearch:=True, matchCase:=False)
    If labelRange Is Nothing Then Exit Sub

    'A refused calculate on a hidden or protected view is logged by the sheet
    'itself; the edit stands either way.
    On Error Resume Next
        labelRange.Calculate
    On Error GoTo 0
End Sub

'@sub-title Refresh the choices values when a default choice is picked.
Private Sub HandleVariablesSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    Dim table As ListObject
    Dim choiceColumn As ListColumn
    Dim nameColumn As ListColumn
    Dim changedCells As Range
    Dim changedCell As Range
    Dim variableName As String
    Dim rowIndex As Long

    If sh.ListObjects.Count = 0 Then Exit Sub
    Set table = sh.ListObjects(1)
    If table.DataBodyRange Is Nothing Then Exit Sub

    Set choiceColumn = FindColumn(table, HEADER_DEFAULT_CHOICE)
    Set nameColumn = FindColumn(table, HEADER_VARIABLE_NAME)
    If choiceColumn Is Nothing Or nameColumn Is Nothing Then Exit Sub

    Set changedCells = Intersect(target, choiceColumn.DataBodyRange)
    If changedCells Is Nothing Then Exit Sub

    MasterSetupHelpers.UnProtectMasterSetupSheet sh, "variables"

    For Each changedCell In changedCells.Cells
        rowIndex = changedCell.Row - table.DataBodyRange.Row + 1
        variableName = Trim$(CStr(nameColumn.DataBodyRange.Cells(rowIndex, 1).Value))
        If LenB(variableName) > 0 Then
            'A half-filled line is normal while the user types; it fills
            'nothing and raises nothing at the user.
            On Error Resume Next
            MasterSetupService.Variables.RefreshChoices variableName, MasterSetupService.Choices
            On Error GoTo 0
        End If
    Next changedCell

    MasterSetupHelpers.ProtectMasterSetupSheet sh, "variables"
End Sub

'@section Internal helpers
'===============================================================================

Private Sub StoreDiseaseLanguage(ByVal sh As Worksheet, ByVal languageTag As String)
    Dim store As HiddenNames

    On Error Resume Next
    Set store = HiddenNames.Create(sh)
    store.EnsureName NAME_DISLANG, languageTag, HiddenNameTypeString
    On Error GoTo 0
End Sub

Private Function SheetIs(ByVal sh As Worksheet, ByVal sheetKey As String) As Boolean
    SheetIs = (StrComp(sh.Name, MasterSetupHelpers.ResolveMasterSetupSheetName(sheetKey), vbTextCompare) = 0)
End Function

Private Function FindColumn(ByVal table As ListObject, ByVal headerName As String) As ListColumn
    Dim column As ListColumn

    For Each column In table.ListColumns
        If StrComp(Trim$(CStr(column.Name)), headerName, vbTextCompare) = 0 Then
            Set FindColumn = column
            Exit Function
        End If
    Next column
End Function
