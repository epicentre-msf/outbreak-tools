Attribute VB_Name = "EventsDesignerMulti"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Ribbon callbacks for the Multi group on the designer workbook.")
'@depends CustomTable, ApplicationState, OSFiles, DropdownLists, BetterArray, EventsDesignerAdvanced
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Ribbon callbacks for the Multi group manage the T_Multi ListObject on
'the GenerateMultiple worksheet. Each callback follows the established
'pattern: show dialogs before entering busy state, wrap work in
'On Error GoTo Cleanup, and restore application state on exit.
'
'THE ID RULE
'-------------------------------------------------------------------------------
'An ID is written once. Adding rows, duplicating a row, resizing and
'importing all fill only the blank ID cells with the next free numbers,
'through EnsureRowIds. The per-row language dropdown is named after the
'row ID (<id>_lang), so a rewritten ID would detach its row from the
'dropdown the validation points at.
'
'A step that gets skipped is reported: the callbacks collect one line
'per skipped step and show them in one message after the busy state is
'restored.

Private Const SHEET_GENERATE_MULTIPLE As String = "GenerateMultiple"
Private Const TABLE_MULTI As String = "T_Multi"
Private Const PROMPT_TITLE As String = "Designer"

'The twelve columns of T_Multi. The callbacks here write the three path
'columns and wire the dictionary language dropdown; the build driver of
'the multi generation reads the output file, the two passwords and
'writes the result.
Private Const COL_ID As String = "ID"
Private Const COL_SETUPS As String = "setups"
Private Const COL_GEOBASES As String = "geobases"
Private Const COL_OUTPUT_FOLDERS As String = "output folders"
Private Const COL_OUTPUT_FILES As String = "output files"
Private Const COL_PASSWORD As String = "output file password"
Private Const COL_DEBUG_PASSWORD As String = "output file debugging password"
Private Const COL_LANG_DICTIONARY As String = "language of the dictionary"
Private Const COL_LANG_INTERFACE As String = "language of the interface"
Private Const COL_EPIWEEK_START As String = "epiweek start"
Private Const COL_DESIGN As String = "design"
Private Const COL_RESULT As String = "result"

'Setup language extraction
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const ID_PREFIX As String = "Operation-"

'Dropdown-based language validation
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const DROPDOWN_PREFIX As String = "dropdown_"
Private Const LANG_SUFFIX As String = "_lang"

Private Const MSG_PLACE_DATA As String = "Please place the cursor inside the table data area."


'@section Multi group callbacks
'===============================================================================

'@Description("Load files or folder into the active T_Multi column (setups, geobases, output folders).")
'@EntryPoint
Public Sub clickFolderMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim colName As String
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim skipped As BetterArray
    Dim startRow As Long

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    'The write lands at the cursor row, so the cursor has to sit on a
    'data cell. A cursor on the header row used to write the first file
    'path over the header text.
    If lo.DataBodyRange Is Nothing Then
        MsgBox MSG_PLACE_DATA, vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If
    If Intersect(Application.ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox MSG_PLACE_DATA, vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

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

    startRow = Application.ActiveCell.Row

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set skipped = New BetterArray
    skipped.LowerBound = 1

    Select Case LCase$(colName)
    Case LCase$(COL_SETUPS)
        LoadSetupFiles lo, CollectFiles(io), startRow, ResolveDropdownManager(), skipped
    Case LCase$(COL_GEOBASES)
        LoadGeobaseFiles lo, CollectFiles(io), startRow, skipped
    Case LCase$(COL_OUTPUT_FOLDERS)
        LoadOutputFolder lo, io.Folder(), startRow, skipped
    End Select

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickFolderMulti: "; errNumber; errDesc
        MsgBox "Unable to load files: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not skipped Is Nothing Then
        ShowSkipped skipped
    End If
End Sub

'@Description("Duplicate the active row in T_Multi with the same values and a fresh ID.")
'@EntryPoint
Public Sub clickDupMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim appScope As ApplicationState
    Dim relPos As Long
    Dim sourceRow As Range
    Dim destRow As Range
    Dim idCol As ListColumn
    Dim idMissing As Boolean

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    'Verify the active cell is inside the table data body
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Application.ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox MSG_PLACE_DATA, vbInformation + vbOKOnly, PROMPT_TITLE
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

    'The copy carried the source row's ID, and two rows sharing an ID
    'share the <id>_lang dropdown. The new row starts blank and gets the
    'next free number.
    On Error Resume Next
    Set idCol = lo.ListColumns(COL_ID)
    On Error GoTo Cleanup

    If idCol Is Nothing Then
        idMissing = True
    Else
        destRow.Cells(1, idCol.Index).Value = vbNullString
        EnsureRowIds lo
    End If

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickDupMulti: "; errNumber; errDesc
        MsgBox "Unable to duplicate row: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf idMissing Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Add rows to the T_Multi table. New rows get the next free IDs.")
'@EntryPoint
Public Sub clickAddRowsMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim table As CustomTable
    Dim appScope As ApplicationState
    Dim hasIdColumn As Boolean

    hasIdColumn = True

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set table = CustomTable.Create(lo)
    table.AddRows nbRows:=10, insertShift:=False, includeIds:=False
    hasIdColumn = EnsureRowIds(lo)

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickAddRowsMulti: "; errNumber; errDesc
        MsgBox "Unable to add rows: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not hasIdColumn Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Resize the T_Multi table by removing empty rows. Kept rows keep their IDs.")
'@EntryPoint
Public Sub clickResizeMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim table As CustomTable
    Dim appScope As ApplicationState
    Dim hasIdColumn As Boolean

    hasIdColumn = True

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set table = CustomTable.Create(lo)
    table.RemoveRows totalCount:=0, includeIds:=False, forceShift:=False
    hasIdColumn = EnsureRowIds(lo)

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickResizeMulti: "; errNumber; errDesc
        MsgBox "Unable to resize table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not hasIdColumn Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Import T_Multi data from another workbook. Blank IDs get the next free numbers.")
'@EntryPoint
Public Sub clickImpMulti(ByRef control As IRibbonControl)
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim importBook As Workbook
    Dim sourceLo As ListObject
    Dim targetLo As ListObject
    Dim sourceTable As CustomTable
    Dim targetTable As CustomTable
    Dim hasIdColumn As Boolean

    hasIdColumn = True

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
        ReportMissingTable
        GoTo Cleanup
    End If

    Set sourceTable = CustomTable.Create(sourceLo)
    Set targetTable = CustomTable.Create(targetLo)
    targetTable.Import sourceTable
    hasIdColumn = EnsureRowIds(targetLo)

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
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickImpMulti: "; errNumber; errDesc
        MsgBox "Unable to import table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not hasIdColumn Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Export the T_Multi table to a new workbook in a user-selected folder.")
'@EntryPoint
Public Sub clickExportMulti(ByRef control As IRibbonControl)
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim lo As ListObject
    Dim table As CustomTable
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
    If lo Is Nothing Then
        ReportMissingTable
        GoTo Cleanup
    End If

    Set table = CustomTable.Create(lo)

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
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickExportMulti: "; errNumber; errDesc
        MsgBox "Unable to export table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Generate one linelist per row of the multi table. The driver that walks the rows arrives in a later update; this stub keeps the ribbon button wired.")
'@EntryPoint
Public Sub clickGenerateMulti(ByRef control As IRibbonControl)
    MsgBox "Multi generation is under construction.", _
           vbInformation + vbOKOnly, PROMPT_TITLE
End Sub


'@section Table and dropdown resolution
'===============================================================================

'@Description("Resolve the T_Multi ListObject from the GenerateMultiple worksheet.")
'@param targetBook Optional Workbook. The workbook to resolve on. Defaults to this workbook.
'@return ListObject. The T_Multi ListObject, or Nothing when not found.
Public Function ResolveMultiTable(Optional ByVal targetBook As Workbook = Nothing) As ListObject
    Dim sh As Worksheet

    If targetBook Is Nothing Then Set targetBook = ThisWorkbook

    On Error Resume Next
    Set sh = targetBook.Worksheets(SHEET_GENERATE_MULTIPLE)
    On Error GoTo 0

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Set ResolveMultiTable = sh.ListObjects(TABLE_MULTI)
    On Error GoTo 0
End Function

'@Description("Resolve the DropdownLists manager on the __dropdowns worksheet.")
'@param targetBook Optional Workbook. The workbook to resolve on. Defaults to this workbook.
'@return DropdownLists. The dropdown manager, or Nothing when the sheet is missing.
Public Function ResolveDropdownManager(Optional ByVal targetBook As Workbook = Nothing) As DropdownLists
    Dim dropSheet As Worksheet

    If targetBook Is Nothing Then Set targetBook = ThisWorkbook

    On Error Resume Next
    Set dropSheet = targetBook.Worksheets(SHEET_DROPDOWNS)
    On Error GoTo 0

    If dropSheet Is Nothing Then Exit Function

    Set ResolveDropdownManager = DropdownLists.Create(dropSheet, DROPDOWN_PREFIX)
End Function

'@Description("Fill only the blank ID cells with the next free numbers.")
'@details
'An ID is written once. This scans the ID column for the largest number
'already written, then gives every blank cell the next numbers, so a row
'keeps the dropdown named after its ID for its whole life.
'@param lo ListObject. The T_Multi ListObject.
'@return Boolean. True when the ID column exists.
Public Function EnsureRowIds(ByVal lo As ListObject) As Boolean
    Dim idCol As ListColumn
    Dim idRange As Range
    Dim rowIdx As Long
    Dim nextNumber As Long
    Dim numberPart As Long
    Dim cellText As String

    On Error Resume Next
    Set idCol = lo.ListColumns(COL_ID)
    On Error GoTo 0

    If idCol Is Nothing Then Exit Function
    EnsureRowIds = True

    Set idRange = idCol.DataBodyRange
    If idRange Is Nothing Then Exit Function

    'Find the largest number already written
    For rowIdx = 1 To idRange.Rows.Count
        cellText = Trim$(CStr(idRange.Cells(rowIdx, 1).Value))
        If LenB(cellText) > 0 Then
            numberPart = TrailingNumber(cellText)
            If numberPart > nextNumber Then nextNumber = numberPart
        End If
    Next rowIdx

    'Fill the blank cells. The ID shape matches the one CustomTable
    'writes: the prefix, one space, the number.
    For rowIdx = 1 To idRange.Rows.Count
        cellText = Trim$(CStr(idRange.Cells(rowIdx, 1).Value))
        If LenB(cellText) = 0 Then
            nextNumber = nextNumber + 1
            idRange.Cells(rowIdx, 1).Value = ID_PREFIX & " " & CStr(nextNumber)
        End If
    Next rowIdx
End Function

'@Description("Read the number at the end of an ID value.")
'@param idText String. The ID cell text.
'@return Long. The trailing number, or 0 when the text ends without digits.
Private Function TrailingNumber(ByVal idText As String) As Long
    Dim charIdx As Long
    Dim oneChar As String
    Dim digits As String

    For charIdx = Len(idText) To 1 Step -1
        oneChar = Mid$(idText, charIdx, 1)
        If oneChar Like "[0-9]" Then
            digits = oneChar & digits
        Else
            Exit For
        End If
    Next charIdx

    If LenB(digits) > 0 Then TrailingNumber = CLng(digits)
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


'@section Folder multi helpers -- file loading by column type
'===============================================================================

'@Description("Write setup paths into the setups column and wire one language dropdown per row.")
'@param lo ListObject. The T_Multi ListObject.
'@param filePaths BetterArray. Setup file paths (1-based).
'@param startRow Long. Worksheet row number the first path lands on.
'@param drop DropdownLists. The dropdown manager of the host workbook.
'@param skipped BetterArray. Collects one line per skipped step.
Public Sub LoadSetupFiles(ByVal lo As ListObject, _
                          ByVal filePaths As BetterArray, _
                          ByVal startRow As Long, _
                          ByVal drop As DropdownLists, _
                          ByVal skipped As BetterArray)
    Dim setupBook As Workbook
    Dim tradSheet As Worksheet
    Dim langValues As BetterArray
    Dim langCol As ListColumn
    Dim idCol As ListColumn
    Dim currentRow As Long
    Dim filePath As String
    Dim dropName As String
    Dim rowId As String
    Dim langCell As Range
    Dim idx As Long

    If filePaths Is Nothing Then Exit Sub
    If filePaths.Length = 0 Then Exit Sub

    WriteFilesToColumn lo, COL_SETUPS, startRow, filePaths, skipped

    'The write may have extended the table; the new rows take their IDs
    'here so each row's dropdown has its name.
    If Not EnsureRowIds(lo) Then
        skipped.Push MissingIdMessage()
        Exit Sub
    End If

    On Error Resume Next
    Set langCol = lo.ListColumns(COL_LANG_DICTIONARY)
    Set idCol = lo.ListColumns(COL_ID)
    On Error GoTo 0

    If langCol Is Nothing Then
        skipped.Push MissingColumnMessage(COL_LANG_DICTIONARY) & _
                     " The language dropdowns were skipped."
        Exit Sub
    End If

    If drop Is Nothing Then
        skipped.Push "The " & SHEET_DROPDOWNS & _
                     " sheet is missing. The language dropdowns were skipped."
        Exit Sub
    End If

    'For each setup file, extract languages and wire a per-row dropdown
    currentRow = startRow
    For idx = filePaths.LowerBound To filePaths.UpperBound
        filePath = CStr(filePaths.Item(idx))

        'Open the setup file read-only
        Set setupBook = Nothing
        On Error Resume Next
        Set setupBook = Workbooks.Open(filePath, ReadOnly:=True)
        On Error GoTo 0

        If setupBook Is Nothing Then
            skipped.Push "This setup file failed to open: " & filePath
        Else
            'Resolve the Translations worksheet
            Set tradSheet = Nothing
            On Error Resume Next
            Set tradSheet = setupBook.Worksheets(SHEET_TRANSLATIONS)
            On Error GoTo 0

            If tradSheet Is Nothing Then
                skipped.Push "This setup file has no " & SHEET_TRANSLATIONS & _
                             " sheet: " & filePath
            Else
                Set langValues = EventsDesignerAdvanced.SetupLanguages(tradSheet)
                If langValues.Length = 0 Then
                    skipped.Push "No language was found in: " & filePath
                Else
                    'The dropdown is named after the row ID
                    rowId = CStr(lo.Parent.Cells(currentRow, idCol.Range.Column).Value)
                    dropName = rowId & LANG_SUFFIX

                    'Add or update the dropdown with extracted languages
                    If drop.Exists(dropName) Then
                        drop.Update langValues, dropName
                    Else
                        drop.Add langValues, dropName
                    End If

                    'Apply validation on the language cell using the dropdown
                    Set langCell = lo.Parent.Cells(currentRow, langCol.Range.Column)
                    drop.SetValidation langCell, dropName
                End If
            End If

            setupBook.Close saveChanges:=False
            Set setupBook = Nothing
        End If

        currentRow = currentRow + 1
    Next idx
End Sub

'@Description("Write geobase paths into the geobases column.")
'@param lo ListObject. The T_Multi ListObject.
'@param filePaths BetterArray. Geobase file paths (1-based).
'@param startRow Long. Worksheet row number the first path lands on.
'@param skipped BetterArray. Collects one line per skipped step.
Public Sub LoadGeobaseFiles(ByVal lo As ListObject, _
                            ByVal filePaths As BetterArray, _
                            ByVal startRow As Long, _
                            ByVal skipped As BetterArray)
    If filePaths Is Nothing Then Exit Sub
    If filePaths.Length = 0 Then Exit Sub

    WriteFilesToColumn lo, COL_GEOBASES, startRow, filePaths, skipped
    If Not EnsureRowIds(lo) Then skipped.Push MissingIdMessage()
End Sub

'@Description("Write a folder path into the output folders column at the given row.")
'@param lo ListObject. The T_Multi ListObject.
'@param folderPath String. The selected folder path.
'@param startRow Long. Worksheet row number to write on.
'@param skipped BetterArray. Collects one line per skipped step.
Public Sub LoadOutputFolder(ByVal lo As ListObject, _
                            ByVal folderPath As String, _
                            ByVal startRow As Long, _
                            ByVal skipped As BetterArray)
    Dim col As ListColumn

    On Error Resume Next
    Set col = lo.ListColumns(COL_OUTPUT_FOLDERS)
    On Error GoTo 0

    If col Is Nothing Then
        skipped.Push MissingColumnMessage(COL_OUTPUT_FOLDERS)
        Exit Sub
    End If

    lo.Parent.Cells(startRow, col.Range.Column).Value = folderPath
End Sub

'@Description("Write file paths into a column, adding rows to the table as needed.")
'@param lo ListObject. The T_Multi ListObject.
'@param colName String. Column header to write into.
'@param startRow Long. Worksheet row number to start writing from.
'@param filePaths BetterArray. File paths to write (1-based).
'@param skipped BetterArray. Collects one line per skipped step.
Private Sub WriteFilesToColumn(ByVal lo As ListObject, _
                               ByVal colName As String, _
                               ByVal startRow As Long, _
                               ByVal filePaths As BetterArray, _
                               ByVal skipped As BetterArray)
    Dim col As ListColumn
    Dim currentRow As Long
    Dim lastDataRow As Long
    Dim idx As Long

    On Error Resume Next
    Set col = lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then
        skipped.Push MissingColumnMessage(colName)
        Exit Sub
    End If

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


'@section Message helpers
'===============================================================================

'@Description("Collect the selected file paths into a BetterArray (1-based).")
'@param io OSFiles. The file picker with selected files.
'@return BetterArray. The selected file paths.
Private Function CollectFiles(ByVal io As OSFiles) As BetterArray
    Dim filePaths As BetterArray

    Set filePaths = New BetterArray
    filePaths.LowerBound = 1

    io.ResetFilesIterator
    Do While io.HasNextFile()
        filePaths.Push io.NextFile()
    Loop

    Set CollectFiles = filePaths
End Function

'@Description("Show the collected skip lines in one message.")
'@param skipped BetterArray. The skip lines pushed by the helpers.
Private Sub ShowSkipped(ByVal skipped As BetterArray)
    Dim message As String
    Dim idx As Long

    If skipped.Length = 0 Then Exit Sub

    For idx = skipped.LowerBound To skipped.UpperBound
        message = message & CStr(skipped.Item(idx)) & vbNewLine
    Next idx

    MsgBox "Some steps were skipped:" & vbNewLine & message, _
           vbExclamation + vbOKOnly, PROMPT_TITLE
End Sub

'@Description("Tell the user the multi table is missing.")
Private Sub ReportMissingTable()
    MsgBox "The " & TABLE_MULTI & " table was not found on the " & _
           SHEET_GENERATE_MULTIPLE & " sheet.", _
           vbExclamation + vbOKOnly, PROMPT_TITLE
End Sub

'@Description("Build the message for a missing column.")
'@param colName String. The missing column header.
'@return String. The message line.
Private Function MissingColumnMessage(ByVal colName As String) As String
    MissingColumnMessage = "The column " & Chr(34) & colName & Chr(34) & _
                           " is missing on " & TABLE_MULTI & "."
End Function

'@Description("Build the message for a missing ID column.")
'@return String. The message line.
Private Function MissingIdMessage() As String
    MissingIdMessage = "The " & TABLE_MULTI & " table has no " & COL_ID & _
                       " column, so the rows got no ID."
End Function
