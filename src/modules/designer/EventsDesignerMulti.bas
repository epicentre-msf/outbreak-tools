Attribute VB_Name = "EventsDesignerMulti"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Ribbon callbacks for the Multi group on the designer workbook.")
'@depends CustomTable, ICustomTable, ApplicationState, IApplicationState, OSFiles, IOSFiles, HiddenNames, IHiddenNames, BetterArray
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Ribbon callbacks for the Multi group manage the T_Multi ListObject on
'the GenerateMultiple worksheet. Each callback follows the established
'pattern: show dialogs before entering busy state, wrap work in
'On Error GoTo Cleanup, and restore application state on exit.

Private Const SHEET_GENERATE_MULTIPLE As String = "GenerateMultiple"
Private Const TABLE_MULTI As String = "T_Multi"
Private Const PROMPT_TITLE As String = "Designer"

'Column names on T_Multi
Private Const COL_SETUPS As String = "setups"
Private Const COL_GEOBASES As String = "geobases"
Private Const COL_OUTPUT_FOLDERS As String = "output folders"
Private Const COL_LANG_DICTIONARY As String = "language of the dictionary"

'Setup language extraction
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const SETUP_LANGUAGES_TAG As String = "__SetupTranslationsLanguages__"
Private Const ID_HEADER As String = "ID"
Private Const ID_PREFIX As String = "Operation-"


'@section Multi group callbacks
'===============================================================================

'@Description("Load files or folder into the active T_Multi column (setups, geobases, output folders).")
'@EntryPoint
Public Sub clickFolderMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim colName As String
    Dim io As IOSFiles
    Dim appScope As IApplicationState

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then Exit Sub

    colName = ActiveCellColumnName(lo)

    'Validate that the active cell is on a supported column
    Select Case LCase$(colName)
    Case LCase$(COL_SETUPS), LCase$(COL_GEOBASES), LCase$(COL_OUTPUT_FOLDERS)
        'valid column, continue
    Case Else
        MsgBox "Please place the cursor on the " & Chr(34) & COL_SETUPS & Chr(34) & _
               ", " & Chr(34) & COL_GEOBASES & Chr(34) & ", or " & Chr(34) & _
               COL_OUTPUT_FOLDERS & Chr(34) & " column.", _
               vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End Select

    'Show the appropriate file/folder dialog before entering busy state
    Set io = OSFiles.Create()

    Select Case LCase$(colName)
    Case LCase$(COL_SETUPS)
        io.LoadFiles "*.xlsb;*.xlsx"
        If Not io.HasValidFiles() Then Exit Sub
    Case LCase$(COL_GEOBASES)
        io.LoadFiles "*.xlsx"
        If Not io.HasValidFiles() Then Exit Sub
    Case LCase$(COL_OUTPUT_FOLDERS)
        io.LoadFolder
        If Not io.HasValidFolder() Then Exit Sub
    End Select

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Select Case LCase$(colName)
    Case LCase$(COL_SETUPS)
        LoadSetupFiles lo, io
    Case LCase$(COL_GEOBASES)
        LoadGeobaseFiles lo, io
    Case LCase$(COL_OUTPUT_FOLDERS)
        LoadOutputFolder lo, io
    End Select

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickFolderMulti: "; errNumber; errDesc
        MsgBox "Unable to load files: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Duplicate the active row in T_Multi with the same values.")
'@EntryPoint
Public Sub clickDupMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim appScope As IApplicationState
    Dim relPos As Long
    Dim sourceRow As Range
    Dim destRow As Range

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then Exit Sub

    'Verify the active cell is inside the table data body
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Application.ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox "Please place the cursor inside the table data area.", _
               vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    'Compute the relative row position (1-based within ListRows)
    relPos = Application.ActiveCell.Row - lo.HeaderRowRange.Row

    'Insert a new row immediately below the current one
    If relPos >= lo.ListRows.Count Then
        lo.ListRows.Add
    Else
        lo.ListRows.Add Position:=relPos + 1
    End If

    'Copy values from the source row to the new row
    Set sourceRow = lo.ListRows(relPos).Range
    Set destRow = lo.ListRows(relPos + 1).Range
    destRow.Value = sourceRow.Value

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickDupMulti: "; errNumber; errDesc
        MsgBox "Unable to duplicate row: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Add rows to the T_Multi table.")
'@EntryPoint
Public Sub clickAddRowsMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim table As ICustomTable
    Dim appScope As IApplicationState

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set table = CustomTable.Create(lo, ID_HEADER, ID_PREFIX)
    table.AddRows nbRows:=10, insertShift:=False, includeIds:=True, 

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickAddRowsMulti: "; errNumber; errDesc
        MsgBox "Unable to add rows: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Resize the T_Multi table by removing empty rows.")
'@EntryPoint
Public Sub clickResizeMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim table As ICustomTable
    Dim appScope As IApplicationState

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set table = CustomTable.Create(lo, ID_HEADER, ID_PREFIX)
    table.RemoveRows totalCount:=0, includeIds:=False, forceShift:=False

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickResizeMulti: "; errNumber; errDesc
        MsgBox "Unable to resize table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Import T_Multi data from another workbook.")
'@EntryPoint
Public Sub clickImpMulti(ByRef control As IRibbonControl)
    Dim io As IOSFiles
    Dim appScope As IApplicationState
    Dim importBook As Workbook
    Dim sourceLo As ListObject
    Dim targetLo As ListObject
    Dim sourceTable As ICustomTable
    Dim targetTable As ICustomTable

    'Show file picker before entering busy state
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb;*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=True)

    'Try to find T_Multi on the GenerateMultiple sheet first
    On Error Resume Next
    Set sourceLo = importBook.Worksheets(SHEET_GENERATE_MULTIPLE).ListObjects(TABLE_MULTI)
    On Error GoTo Cleanup

    'Fallback: use the first ListObject on the first worksheet
    If sourceLo Is Nothing Then
        If importBook.Worksheets(1).ListObjects.Count > 0 Then
            Set sourceLo = importBook.Worksheets(1).ListObjects(1)
        End If
    End If

    If sourceLo Is Nothing Then
        importBook.Close saveChanges:=False
        Set importBook = Nothing
        MsgBox "No table found in the selected workbook.", _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        GoTo Cleanup
    End If

    Set targetLo = ResolveMultiTable()
    If targetLo Is Nothing Then
        importBook.Close saveChanges:=False
        Set importBook = Nothing
        GoTo Cleanup
    End If

    Set sourceTable = CustomTable.Create(sourceLo, ID_HEADER, ID_PREFIX)
    Set targetTable = CustomTable.Create(targetLo, ID_HEADER, ID_PREFIX)
    targetTable.Import sourceTable

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not importBook Is Nothing Then
        importBook.Close saveChanges:=False
    End If
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickImpMulti: "; errNumber; errDesc
        MsgBox "Unable to import table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Export the T_Multi table to a new workbook in a user-selected folder.")
'@EntryPoint
Public Sub clickExportMulti(ByRef control As IRibbonControl)
    Dim io As IOSFiles
    Dim appScope As IApplicationState
    Dim lo As ListObject
    Dim table As ICustomTable
    Dim exportBook As Workbook
    Dim exportSheet As Worksheet
    Dim folderPath As String
    Dim exportPath As String

    'Show folder picker before entering busy state
    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then GoTo Cleanup

    Set table = CustomTable.Create(lo, ID_HEADER, ID_PREFIX)

    'Create a new workbook and export the table
    Set exportBook = Workbooks.Add
    Set exportSheet = exportBook.Worksheets(1)
    table.Export sh:=exportSheet, startLine:=1, startColumn:=1, addListObject:=True

    'Build the export file path with timestamp
    folderPath = io.Folder()
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
    exportPath = folderPath & TABLE_MULTI & "_export_" & _
                 Format$(Now, "yyyymmdd\_hhnnss") & ".xlsx"

    exportBook.SaveAs Filename:=exportPath, FileFormat:=xlOpenXMLWorkbook
    exportBook.Close saveChanges:=False
    Set exportBook = Nothing

    appScope.Restore
    Set appScope = Nothing
    MsgBox "Exported to: " & exportPath, vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not exportBook Is Nothing Then
        exportBook.Close saveChanges:=False
    End If
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickExportMulti: "; errNumber; errDesc
        MsgBox "Unable to export table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section Internal helpers
'===============================================================================

'@Description("Resolve the T_Multi ListObject from the GenerateMultiple worksheet.")
'@return ListObject. The T_Multi ListObject, or Nothing when not found.
Private Function ResolveMultiTable() As ListObject
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(SHEET_GENERATE_MULTIPLE)
    On Error GoTo 0

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Set ResolveMultiTable = sh.ListObjects(TABLE_MULTI)
    On Error GoTo 0
End Function

'@Description("Return the T_Multi column header matching the active cell position.")
'@param lo ListObject. The T_Multi ListObject.
'@return String. Column header name, or vbNullString when outside the table.
Private Function ActiveCellColumnName(ByVal lo As ListObject) As String
    Dim colOffset As Long

    If Intersect(Application.ActiveCell, lo.Range) Is Nothing Then Exit Function

    colOffset = Application.ActiveCell.Column - lo.HeaderRowRange.Column + 1
    If colOffset < 1 Or colOffset > lo.ListColumns.Count Then Exit Function

    ActiveCellColumnName = lo.ListColumns(colOffset).Name
End Function


'@section Folder multi helpers — file loading by column type
'===============================================================================

'@Description("Load selected setup files into the setups column and apply per-row language validation.")
'@param lo ListObject. The T_Multi ListObject.
'@param io IOSFiles. The file picker with selected files.
Private Sub LoadSetupFiles(ByVal lo As ListObject, ByVal io As IOSFiles)
    Dim filePaths As BetterArray
    Dim setupBook As Workbook
    Dim tradSheet As Worksheet
    Dim langString As String
    Dim langCol As ListColumn
    Dim startRow As Long
    Dim currentRow As Long
    Dim filePath As String

    'Collect file paths into a BetterArray
    Set filePaths = New BetterArray
    filePaths.LowerBound = 1
    io.ResetFilesIterator
    Do While io.HasNextFile()
        filePaths.Push io.NextFile()
    Loop

    If filePaths.Length = 0 Then Exit Sub

    startRow = Application.ActiveCell.Row

    'Write all file paths into the setups column, extending the table as needed
    WriteFilesToColumn lo, COL_SETUPS, startRow, filePaths

    'Resolve the language of the dictionary column
    On Error Resume Next
    Set langCol = lo.ListColumns(COL_LANG_DICTIONARY)
    On Error GoTo 0

    If langCol Is Nothing Then Exit Sub

    'For each setup file, extract languages and apply per-row validation
    currentRow = startRow
    Dim idx As Long
    For idx = filePaths.LowerBound To filePaths.UpperBound
        filePath = CStr(filePaths.Item(idx))

        'Open the setup file read-only
        Set setupBook = Nothing
        On Error Resume Next
        Set setupBook = Workbooks.Open(filePath, ReadOnly:=True)
        On Error GoTo 0

        If setupBook Is Nothing Then
            currentRow = currentRow + 1
            GoTo ContinueSetup
        End If

        'Resolve the Translations worksheet
        Set tradSheet = Nothing
        On Error Resume Next
        Set tradSheet = setupBook.Worksheets(SHEET_TRANSLATIONS)
        On Error GoTo 0

        If Not tradSheet Is Nothing Then
            langString = ExtractLanguagesForRow(tradSheet)
            If LenB(langString) > 0 Then
                Dim langCell As Range
                Set langCell = lo.Parent.Cells(currentRow, langCol.Range.Column)
                ApplyDirectValidation langCell, langString
            End If
        End If

        setupBook.Close saveChanges:=False
        Set setupBook = Nothing

        currentRow = currentRow + 1
ContinueSetup:
    Next idx
End Sub

'@Description("Load selected geobase files into the geobases column.")
'@param lo ListObject. The T_Multi ListObject.
'@param io IOSFiles. The file picker with selected files.
Private Sub LoadGeobaseFiles(ByVal lo As ListObject, ByVal io As IOSFiles)
    Dim filePaths As BetterArray
    Dim startRow As Long

    Set filePaths = New BetterArray
    filePaths.LowerBound = 1
    io.ResetFilesIterator
    Do While io.HasNextFile()
        filePaths.Push io.NextFile()
    Loop

    If filePaths.Length = 0 Then Exit Sub

    startRow = Application.ActiveCell.Row
    WriteFilesToColumn lo, COL_GEOBASES, startRow, filePaths
End Sub

'@Description("Write a folder path into the output folders column at the active cell row.")
'@param lo ListObject. The T_Multi ListObject.
'@param io IOSFiles. The folder picker with selected folder.
Private Sub LoadOutputFolder(ByVal lo As ListObject, ByVal io As IOSFiles)
    Dim col As ListColumn
    Dim targetCell As Range

    On Error Resume Next
    Set col = lo.ListColumns(COL_OUTPUT_FOLDERS)
    On Error GoTo 0

    If col Is Nothing Then Exit Sub

    Set targetCell = lo.Parent.Cells(Application.ActiveCell.Row, col.Range.Column)
    targetCell.Value = io.Folder()
End Sub


'@section Language extraction and validation helpers
'===============================================================================

'@Description("Extract language names from a setup Translations sheet as a comma-separated string.")
'@param tradSheet Worksheet. The Translations worksheet of a setup workbook.
'@return String. Comma-separated language names, or vbNullString when none found.
Private Function ExtractLanguagesForRow(ByVal tradSheet As Worksheet) As String
    Dim store As IHiddenNames
    Dim langString As String
    Dim languages() As String
    Dim result As String
    Dim idx As Long
    Dim sep As String

    sep = Application.International(xlListSeparator)

    'Try HiddenNames first (same pattern as EventsDesignerAdvanced.ExtractAndUpdateLanguages)
    Set store = HiddenNames.Create(tradSheet)

    If store.HasName(SETUP_LANGUAGES_TAG) Then
        langString = store.ValueAsString(SETUP_LANGUAGES_TAG)
        If LenB(langString) > 0 Then
            languages = Split(langString, ";")
            For idx = LBound(languages) To UBound(languages)
                If LenB(Trim$(languages(idx))) > 0 Then
                    If LenB(result) > 0 Then result = result & sep
                    result = result & Trim$(languages(idx))
                End If
            Next idx
            ExtractLanguagesForRow = result
            Exit Function
        End If
    End If

    'Fallback: read column headers from the first ListObject
    If tradSheet.ListObjects.Count = 0 Then Exit Function

    Dim lo As ListObject
    Set lo = tradSheet.ListObjects(1)
    If lo.HeaderRowRange Is Nothing Then Exit Function

    Dim headerValues As Variant
    headerValues = lo.HeaderRowRange.Value

    If Not IsArray(headerValues) Then
        ExtractLanguagesForRow = CStr(headerValues)
        Exit Function
    End If

    Dim colIdx As Long
    For colIdx = LBound(headerValues, 2) To UBound(headerValues, 2)
        Dim headerVal As String
        headerVal = Trim$(CStr(headerValues(1, colIdx)))
        If LenB(headerVal) > 0 Then
            If LenB(result) > 0 Then result = result & sep
            result = result & headerVal
        End If
    Next colIdx

    ExtractLanguagesForRow = result
End Function

'@Description("Apply a direct list validation to a single cell using a comma-separated formula string.")
'@param cell Range. The cell to validate.
'@param listString String. Comma-separated list of valid values.
Private Sub ApplyDirectValidation(ByVal cell As Range, ByVal listString As String)
    If cell Is Nothing Then Exit Sub
    If LenB(listString) = 0 Then Exit Sub

    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertInformation, _
             Operator:=xlBetween, _
             Formula1:=listString
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'@Description("Write file paths into a column, adding rows to the table as needed.")
'@param lo ListObject. The T_Multi ListObject.
'@param colName String. Column header to write into.
'@param startRow Long. Worksheet row number to start writing from.
'@param filePaths BetterArray. File paths to write (1-based).
Private Sub WriteFilesToColumn(ByVal lo As ListObject, _
                               ByVal colName As String, _
                               ByVal startRow As Long, _
                               ByVal filePaths As BetterArray)
    Dim col As ListColumn
    Dim currentRow As Long
    Dim lastDataRow As Long
    Dim idx As Long

    On Error Resume Next
    Set col = lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then Exit Sub

    currentRow = startRow

    For idx = filePaths.LowerBound To filePaths.UpperBound
        'Ensure the row exists within the table
        lastDataRow = lo.HeaderRowRange.Row + lo.ListRows.Count
        If currentRow > lastDataRow Then
            lo.ListRows.Add
        End If

        lo.Parent.Cells(currentRow, col.Range.Column).Value = CStr(filePaths.Item(idx))
        currentRow = currentRow + 1
    Next idx
End Sub
