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
'  typed in (the first translation language), and the new default follows
'  onto every disease line carrying that variable.
'- Disease sheet: a pick in the Variable Name column fills the section, the
'  default choice and the default status of the line. The label and the
'  choice values follow through the two worksheet functions.
'- Choices sheet: an edit recalculates the label column, refreshes the joined
'  values on the Variables sheet, empties every reference to a choice that is
'  gone, and recalculates the disease lines.
'- A pick in the language cell of a disease sheet is stored in the sheet's
'  hidden names, so exports read the language the user chose.
'- An edit on the Translations sheet drops the translation caches.

'Column order of a disease table, as DiseaseSheet builds it.
Private Const DISEASE_NAME_COLUMN As Long = 2
Private Const DISEASE_SECTION_COLUMN As Long = 3
Private Const DISEASE_LABEL_COLUMN As Long = 4
Private Const DISEASE_CHOICE_COLUMN As Long = 5
Private Const DISEASE_VALUES_COLUMN As Long = 6
Private Const DISEASE_STATUS_COLUMN As Long = 7

Private Const LANGUAGE_CELL_ROW As Long = 2
Private Const LANGUAGE_CELL_COLUMN As Long = 2
Private Const NAME_DISLANG As String = "__Var_DISLANG"

Private Const HEADER_VARIABLE_NAME As String = "Variable Name"
Private Const HEADER_DEFAULT_CHOICE As String = "Default Choice"
Private Const HEADER_CHOICES_VALUES As String = "Choices Values"
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
        'The choice values feed the worksheet functions; the caches drop
        'before anything recalculates over them, and the service drops its
        'choices helper so the cascade reads the edited block.
        ResetMasterSetupFunctionCaches
        MasterSetupService.RefreshTranslations
        HandleChoicesSheetChange sh, target
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

'@sub-title Carry a choices edit into every sheet reading the choices.
'@details The label column recalculates the way the setup does it; then the
'Variables sheet refreshes its joined values, references to a choice that is
'gone empty out, and the disease lines recalculate. Speed rule: the edited
'and existing choice names are gathered ONCE into keyed Collections, every
'column is read in one crossing, and a sheet is only unprotected and written
'when one of its rows actually moves.
Private Sub HandleChoicesSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    Dim editedChoices As Collection
    Dim existingChoices As Collection

    RecalculateChoiceLabels target

    Set editedChoices = EditedChoiceNames(target)
    Set existingChoices = ExistingChoiceNames()
    If existingChoices Is Nothing Then Exit Sub

    CascadeChoicesToVariables editedChoices, existingChoices
    CascadeChoicesToDiseases editedChoices, existingChoices
End Sub

'@sub-title Collect the list names of the edited choices rows, keyed.
Private Function EditedChoiceNames(ByVal target As Range) As Collection
    Dim table As ListObject
    Dim wrapper As CustomTable
    Dim listRange As Range
    Dim touched As Range
    Dim rowCells As Range
    Dim nameValue As String

    Set EditedChoiceNames = New Collection

    Set table = target.ListObject
    If table Is Nothing Then Exit Function
    If table.DataBodyRange Is Nothing Then Exit Function

    Set wrapper = CustomTable.Create(table)
    Set listRange = wrapper.DataRange(colName:="list name", includeHeaders:=False, strictSearch:=True, matchCase:=False)
    If listRange Is Nothing Then Exit Function

    Set touched = Intersect(target.EntireRow, listRange)
    If touched Is Nothing Then Exit Function

    For Each rowCells In touched.Cells
        nameValue = Trim$(CStr(rowCells.Value))
        If LenB(nameValue) > 0 Then AddKey EditedChoiceNames, nameValue
    Next rowCells
End Function

'@sub-title Collect every list name the Choices sheet carries, keyed.
Private Function ExistingChoiceNames() As Collection
    Dim choices As LLChoices
    Dim names As BetterArray
    Dim idx As Long

    Set choices = MasterSetupService.Choices
    If choices Is Nothing Then Exit Function

    Set ExistingChoiceNames = New Collection

    Set names = choices.AllChoices
    If names Is Nothing Then Exit Function

    For idx = names.LowerBound To names.UpperBound
        AddKey ExistingChoiceNames, Trim$(CStr(names.Item(idx)))
    Next idx
End Function

'@sub-title Recalculate the label column when a translated label changes.
Private Sub RecalculateChoiceLabels(ByVal target As Range)
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

'@sub-title Refresh the Variables sheet after a choices edit.
'@details A default choice the Choices sheet no longer carries empties with
'its joined values; a default choice whose list was edited gets its values
'rejoined. Both columns are read in one crossing each, and the sheet is only
'unprotected when a row actually moves.
Private Sub CascadeChoicesToVariables(ByVal editedChoices As Collection, ByVal existingChoices As Collection)
    Dim variablesSheet As Worksheet
    Dim table As ListObject
    Dim nameColumn As ListColumn
    Dim choiceColumn As ListColumn
    Dim valuesColumn As ListColumn
    Dim choices As LLChoices
    Dim nameData As Variant
    Dim choiceData As Variant
    Dim rowIndex As Long
    Dim variableName As String
    Dim choiceValue As String
    Dim unprotected As Boolean

    Set variablesSheet = MasterSetupHelpers.ResolveMasterVariablesSheet()
    If variablesSheet Is Nothing Then Exit Sub
    If variablesSheet.ListObjects.Count = 0 Then Exit Sub

    Set table = variablesSheet.ListObjects(1)
    If table.DataBodyRange Is Nothing Then Exit Sub

    Set nameColumn = FindColumn(table, HEADER_VARIABLE_NAME)
    Set choiceColumn = FindColumn(table, HEADER_DEFAULT_CHOICE)
    Set valuesColumn = FindColumn(table, HEADER_CHOICES_VALUES)
    If nameColumn Is Nothing Or choiceColumn Is Nothing Then Exit Sub

    Set choices = MasterSetupService.Choices
    If choices Is Nothing Then Exit Sub

    'One crossing per column; the loop below runs in memory.
    nameData = ColumnBlock(nameColumn.DataBodyRange)
    choiceData = ColumnBlock(choiceColumn.DataBodyRange)

    For rowIndex = 1 To UBound(choiceData, 1)
        choiceValue = Trim$(CStr(choiceData(rowIndex, 1)))
        If LenB(choiceValue) > 0 Then
            If Not HasKey(existingChoices, choiceValue) Then
                'The choice is gone from the Choices sheet; every reference
                'to it empties.
                EnsureUnprotected variablesSheet, "variables", unprotected
                choiceColumn.DataBodyRange.Cells(rowIndex, 1).Value = vbNullString
                If Not valuesColumn Is Nothing Then
                    valuesColumn.DataBodyRange.Cells(rowIndex, 1).Value = vbNullString
                End If
            ElseIf HasKey(editedChoices, choiceValue) Then
                variableName = Trim$(CStr(nameData(rowIndex, 1)))
                If LenB(variableName) > 0 Then
                    EnsureUnprotected variablesSheet, "variables", unprotected
                    'A half-filled line is normal while the user types; it
                    'fills nothing and raises nothing at the user.
                    On Error Resume Next
                    MasterSetupService.Variables.RefreshChoices variableName, choices
                    On Error GoTo 0
                End If
            End If
        End If
    Next rowIndex

    If unprotected Then MasterSetupHelpers.ProtectMasterSetupSheet variablesSheet, "variables"
End Sub

'@sub-title Refresh the disease lines a choices edit reaches.
'@details Each disease table's choice column is read in one crossing. A line
'referencing a choice that is gone empties; a table is recalculated only when
'one of its lines references an edited or a missing choice.
Private Sub CascadeChoicesToDiseases(ByVal editedChoices As Collection, ByVal existingChoices As Collection)
    Dim sh As Worksheet
    Dim table As ListObject
    Dim choiceData As Variant
    Dim rowIndex As Long
    Dim choiceValue As String
    Dim needsRecalc As Boolean
    Dim unprotected As Boolean

    For Each sh In ThisWorkbook.Worksheets
        If MasterSetupHelpers.IsMasterDiseaseSheet(sh) Then
            If sh.ListObjects.Count > 0 Then
                Set table = sh.ListObjects(1)
                If Not table.DataBodyRange Is Nothing Then
                    choiceData = ColumnBlock(table.DataBodyRange.Columns(DISEASE_CHOICE_COLUMN))
                    needsRecalc = False
                    unprotected = False

                    For rowIndex = 1 To UBound(choiceData, 1)
                        choiceValue = Trim$(CStr(choiceData(rowIndex, 1)))
                        If LenB(choiceValue) > 0 Then
                            If Not HasKey(existingChoices, choiceValue) Then
                                EnsureUnprotected sh, "disease", unprotected
                                table.DataBodyRange.Cells(rowIndex, DISEASE_CHOICE_COLUMN).Value = vbNullString
                                needsRecalc = True
                            ElseIf HasKey(editedChoices, choiceValue) Then
                                needsRecalc = True
                            End If
                        End If
                    Next rowIndex

                    If needsRecalc Then
                        EnsureUnprotected sh, "disease", unprotected
                        RecalculateDiseaseColumns table
                    End If
                    If unprotected Then MasterSetupHelpers.ProtectMasterSetupSheet sh, "disease"
                End If
            End If
        End If
    Next sh
End Sub

'@sub-title Write a new default choice onto every disease line of a variable.
'@details The name column is read in one crossing per table; a sheet with no
'line carrying the variable is left untouched, protected and uncalculated.
Private Sub PropagateDefaultChoice(ByVal variableName As String, ByVal choiceValue As String)
    Dim sh As Worksheet
    Dim table As ListObject
    Dim nameData As Variant
    Dim rowIndex As Long
    Dim unprotected As Boolean

    If LenB(variableName) = 0 Then Exit Sub

    For Each sh In ThisWorkbook.Worksheets
        If MasterSetupHelpers.IsMasterDiseaseSheet(sh) Then
            If sh.ListObjects.Count > 0 Then
                Set table = sh.ListObjects(1)
                If Not table.DataBodyRange Is Nothing Then
                    nameData = ColumnBlock(table.DataBodyRange.Columns(DISEASE_NAME_COLUMN))
                    unprotected = False

                    For rowIndex = 1 To UBound(nameData, 1)
                        If StrComp(Trim$(CStr(nameData(rowIndex, 1))), variableName, vbTextCompare) = 0 Then
                            EnsureUnprotected sh, "disease", unprotected
                            table.DataBodyRange.Cells(rowIndex, DISEASE_CHOICE_COLUMN).Value = choiceValue
                        End If
                    Next rowIndex

                    If unprotected Then
                        RecalculateDiseaseColumns table
                        MasterSetupHelpers.ProtectMasterSetupSheet sh, "disease"
                    End If
                End If
            End If
        End If
    Next sh
End Sub

'@sub-title Unprotect a sheet the first time a write needs it.
Private Sub EnsureUnprotected(ByVal sh As Worksheet, ByVal sheetTag As String, ByRef unprotected As Boolean)
    If unprotected Then Exit Sub
    MasterSetupHelpers.UnProtectMasterSetupSheet sh, sheetTag
    unprotected = True
End Sub

'@sub-title Read a one-column range as a 2D block, whatever its height.
Private Function ColumnBlock(ByVal columnRange As Range) As Variant
    Dim values As Variant
    Dim scalar(1 To 1, 1 To 1) As Variant

    values = columnRange.Value
    If IsArray(values) Then
        ColumnBlock = values
    Else
        scalar(1, 1) = values
        ColumnBlock = scalar
    End If
End Function

'@sub-title Add a key once; a duplicate add is a no-op.
Private Sub AddKey(ByVal keys As Collection, ByVal keyValue As String)
    ' A Collection refuses a duplicate key with a raise; the scoped probe is
    ' the cheapest existence test. Restored right after.
    On Error Resume Next
    keys.Add keyValue, keyValue
    On Error GoTo 0
End Sub

'@sub-title Answer whether a keyed Collection carries the key.
Private Function HasKey(ByVal keys As Collection, ByVal keyValue As String) As Boolean
    Dim probe As Variant

    If keys Is Nothing Then Exit Function

    On Error Resume Next
    probe = keys.Item(keyValue)
    HasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

'@sub-title Recalculate the two formula columns of a disease table.
Private Sub RecalculateDiseaseColumns(ByVal table As ListObject)
    'The label and the choice values columns carry the two worksheet
    'functions; a refused calculate leaves the old text standing.
    On Error Resume Next
        table.DataBodyRange.Columns(DISEASE_LABEL_COLUMN).Calculate
        table.DataBodyRange.Columns(DISEASE_VALUES_COLUMN).Calculate
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

            'The new default follows onto every disease line carrying the
            'variable.
            PropagateDefaultChoice variableName, Trim$(CStr(changedCell.Value))
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
