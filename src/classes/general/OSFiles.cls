VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OSFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Importing files from the operating system"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Importing files from the operating system")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation


Option Explicit
Implements IOSFiles

'@class OSFiles
'@description
'Cross-platform file and folder selection helper. Wraps Windows
'FileDialog and macOS AppleScript pickers behind a unified API.
'Supports single and multi-selection for both files and folders,
'persists the last-used folder across sessions, and provides
'iterator-style enumeration over multi-selections.
'@version 1.0 (2026-02-09)
'@author OutbreakTools

Private Type TOSFiles
    'State bag for OS detection, selections, and cursor positions.
    currentOS As String
    currentLoadedFile As String
    currentLoadedFolder As String
    currentLoadedFiles As Variant
    currentLoadedFolders As Variant
    lastKnownFolder As String
    currentFileIndex As Long
    currentFolderIndex As Long
End Type

'Exposed methods
Private this As TOSFiles

'Workbook-level Name used to persist the last used folder path.
Private Const LAST_FOLDER_STORAGE_NAME As String = "OSFiles_LastFolder"

'@section Initialization
'===============================================================================
'Construction, state reset, and low-level array utilities.

'@label:class-initialize
'@sub-title Set default values for file and folder state
Private Sub Class_Initialize()
    'Ensure every new instance starts with clean selections and no cached folder.
    ResetFileState
    ResetFolderState
    this.lastKnownFolder = vbNullString
End Sub

'@label:reset-file-state
'@sub-title Clear file-related state and reset the iteration cursor
Private Sub ResetFileState()
    this.currentLoadedFile = vbNullString
    this.currentLoadedFiles = EmptyStringArray()
    this.currentFileIndex = 0
End Sub

'@label:reset-folder-state
'@sub-title Clear folder-related state and reset the iteration cursor
Private Sub ResetFolderState()
    this.currentLoadedFolder = vbNullString
    this.currentLoadedFolders = EmptyStringArray()
    this.currentFolderIndex = 0
End Sub

'@label:empty-string-array
'@sub-title Return an empty variant array
'@return Variant. Empty array.
Private Function EmptyStringArray() As Variant
    EmptyStringArray = Array()
End Function

'@label:select-folders-on-windows
'@sub-title Show the Windows multi-folder picker
'@return Variant. Zero-based string array of folder paths.
Private Function SelectFoldersOnWindows() As Variant

    Dim fDialog As Office.FileDialog
    Dim selectedCount As Long
    Dim results() As String
    Dim index As Long

    SelectFoldersOnWindows = EmptyStringArray()

    'Use the built-in folder picker to collect directories on Windows Excel.
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = True
        .title = "Chose your directories"        
        .filters.Clear

        If .Show = -1 Then
            'The SelectedItems collection can throw; swallow and bail gracefully.
            On Error Resume Next
            selectedCount = .SelectedItems.Count
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0

            If selectedCount > 0 Then
                ReDim results(0 To selectedCount - 1)
                For index = 1 To selectedCount
                    results(index - 1) = CStr(.SelectedItems(index))
                Next index
                'Return the normalized, zero-based results list.
                SelectFoldersOnWindows = results
            End If
        End If
    End With
End Function

'@label:try-get-array-bounds
'@sub-title Safely probe array bounds
Private Function TryGetArrayBounds(ByVal candidate As Variant, _
                                   ByRef lowerBound As Long, _
                                   ByRef upperBound As Long) As Boolean

    'Safely probe for bounds; many callers pass Empty or scalars.
    If Not IsArray(candidate) Then
        Exit Function
    End If

    On Error Resume Next
    lowerBound = LBound(candidate)
    upperBound = UBound(candidate)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If upperBound < lowerBound Then
        Exit Function
    End If

    'Only mark success once we have a valid span of indexes.
    TryGetArrayBounds = True
End Function

'@label:count-array-items
'@sub-title Count the number of items in a variant array
Private Function CountArrayItems(ByVal candidate As Variant) As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    'Translate array bounds into a count; returns 0 for unsupported inputs.
    If Not TryGetArrayBounds(candidate, lowerBound, upperBound) Then
        Exit Function
    End If

    CountArrayItems = (upperBound - lowerBound) + 1
End Function

'@label:to-string-array
'@sub-title Coerce a variant into a zero-based string array
Private Function ToStringArray(ByVal value As Variant) As Variant
    Dim normalized() As String
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim position As Long

    'Coerce the supplied value into a zero-based string array for storage.
    If IsArray(value) Then
        If Not TryGetArrayBounds(value, lowerBound, upperBound) Then
            ToStringArray = EmptyStringArray()
            Exit Function
        End If

        ReDim normalized(0 To (upperBound - lowerBound))
        'Walk the original bounds and copy each entry as text.
        For index = lowerBound To upperBound
            normalized(position) = CStr(value(index))
            position = position + 1
        Next index

        ToStringArray = normalized
        Exit Function
    End If

    If IsEmpty(value) Then
        ToStringArray = EmptyStringArray()
        Exit Function
    End If

    If IsNull(value) Then
        ToStringArray = EmptyStringArray()
        Exit Function
    End If

    If VarType(value) = vbString Then
        If value = vbNullString Then
            ToStringArray = EmptyStringArray()
        Else
            'Wrap single strings in a 1-element array to simplify callers.
            ReDim normalized(0 To 0)
            normalized(0) = CStr(value)
            ToStringArray = normalized
        End If
        Exit Function
    End If

    ReDim normalized(0 To 0)
    normalized(0) = CStr(value)
    ToStringArray = normalized
End Function

'@label:clone-string-array
'@sub-title Create a safe copy of a string array
Private Function CloneStringArray(ByVal source As Variant) As Variant
    'Use ToStringArray to guarantee a safe copy that call sites cannot mutate.
    CloneStringArray = ToStringArray(source)
End Function

'@label:has-next-array-item
'@sub-title Check whether the cursor is within array bounds
Private Function HasNextArrayItem(ByVal items As Variant, ByVal cursor As Long) As Boolean
    Dim lowerBound As Long
    Dim upperBound As Long

    'Return True when the cursor still points inside the normalized array span.
    If cursor < 0 Then
        Exit Function
    End If

    If Not TryGetArrayBounds(items, lowerBound, upperBound) Then
        Exit Function
    End If

    HasNextArrayItem = (cursor <= (upperBound - lowerBound))
End Function

'@label:get-next-array-item
'@sub-title Fetch the item at the cursor and advance it
Private Function GetNextArrayItem(ByVal items As Variant, _
                                  ByRef cursor As Long) As String
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim actualIndex As Long
    Dim resultValue As String

    'Fetch the item at the cursor and advance it, tolerating gaps and errors.
    If cursor < 0 Then
        cursor = 0
    End If

    If Not TryGetArrayBounds(items, lowerBound, upperBound) Then
        Exit Function
    End If

    actualIndex = lowerBound + cursor
    If actualIndex > upperBound Then
        Exit Function
    End If

    On Error Resume Next
    resultValue = CStr(items(actualIndex))
    If Err.Number <> 0 Then
        Err.Clear
        resultValue = vbNullString
    End If
    On Error GoTo 0

    cursor = cursor + 1
    GetNextArrayItem = resultValue
End Function

'@label:set-files-selection
'@sub-title Persist file selections and prime the iterator
Private Sub SetFilesSelection(ByVal selectedItems As Variant)
    Dim normalized As Variant
    normalized = ToStringArray(selectedItems)

    'Persist the user selection and prime the iterator against the first entry.
    this.currentLoadedFiles = normalized
    this.currentLoadedFile = FirstStringOrEmpty(normalized)
    this.currentFileIndex = 0
End Sub

'@label:set-folders-selection
'@sub-title Persist folder selections and prime the iterator
Private Sub SetFoldersSelection(ByVal selectedItems As Variant)
    Dim normalized As Variant
    normalized = ToStringArray(selectedItems)

    'Store folder selections and make the first path immediately available.
    this.currentLoadedFolders = normalized
    this.currentLoadedFolder = FirstStringOrEmpty(normalized)
    this.currentFolderIndex = 0
End Sub

'@label:first-string-or-empty
'@sub-title Return the first array entry or an empty string
Private Function FirstStringOrEmpty(ByVal source As Variant) As String
    Dim lowerBound As Long
    Dim upperBound As Long

    'Return the first array entry or an empty string when the selection is void.
    If Not TryGetArrayBounds(source, lowerBound, upperBound) Then
        FirstStringOrEmpty = vbNullString
        Exit Function
    End If

    If upperBound < lowerBound Then
        FirstStringOrEmpty = vbNullString
        Exit Function
    End If

    FirstStringOrEmpty = CStr(source(lowerBound))
End Function

'@label:normalize-folder-path
'@sub-title Trim trailing separators and clean the folder path
Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    Dim cleaned As String

    'Trim OS-specific quirks (extra separators, CR/LF) from folder paths.
    cleaned = Trim$(CStr(folderPath))
    If cleaned = vbNullString Then
        Exit Function
    End If

    cleaned = Replace(cleaned, vbCr, vbNullString)
    cleaned = Replace(cleaned, vbLf, vbNullString)

    Do While Len(cleaned) > 0
        Select Case Right$(cleaned, 1)
            Case "\", "/"
                If Len(cleaned) = 1 Then
                    Exit Do
                End If

                If Len(cleaned) = 3 And Mid$(cleaned, 2, 1) = ":" Then
                    Exit Do
                End If

                'Strip trailing separator while keeping drive roots intact.
                cleaned = Left$(cleaned, Len(cleaned) - 1)
            Case ":"
                If Len(cleaned) > 1 Then
                    cleaned = Left$(cleaned, Len(cleaned) - 1)
                Else
                    Exit Do
                End If
            Case Else
                Exit Do
        End Select
    Loop

    NormalizeFolderPath = cleaned
End Function

'@label:extract-folder-from-path
'@sub-title Identify the parent directory from a file path
Private Function ExtractFolderFromPath(ByVal filePath As String) As String
    Dim cleaned As String
    Dim slashPos As Long
    Dim backslashPos As Long
    Dim colonPos As Long
    Dim delimiterPos As Long

    'Identify the parent directory by looking for the last path delimiter.
    cleaned = Trim$(CStr(filePath))
    If cleaned = vbNullString Then
        Exit Function
    End If

    cleaned = Replace(cleaned, vbCr, vbNullString)
    cleaned = Replace(cleaned, vbLf, vbNullString)

    slashPos = InStrRev(cleaned, "/")
    backslashPos = InStrRev(cleaned, "\")
    colonPos = InStrRev(cleaned, ":")

    delimiterPos = slashPos
    If backslashPos > delimiterPos Then
        delimiterPos = backslashPos
    End If
    If colonPos > delimiterPos Then
        delimiterPos = colonPos
    End If

    If delimiterPos <= 0 Then
        Exit Function
    End If

    ExtractFolderFromPath = Left$(cleaned, delimiterPos - 1)
End Function

'@label:remember-folder
'@sub-title Cache the folder locally and in the workbook
Private Sub RememberFolder(ByVal folderPath As String)
    Dim normalized As String

    'Cache the folder locally and in the workbook so future pickers prefill it.
    normalized = NormalizeFolderPath(folderPath)
    If normalized = vbNullString Then
        Exit Sub
    End If

    this.lastKnownFolder = normalized
    WriteLastFolderToWorkbook normalized
End Sub

'@label:remember-folder-from-files
'@sub-title Persist the parent directory of the first selected file
Private Sub RememberFolderFromFiles(ByVal filesSelection As Variant)
    Dim firstFile As String

    'Persist the directory that contains the first selected file.
    firstFile = FirstStringOrEmpty(filesSelection)
    If firstFile = vbNullString Then
        Exit Sub
    End If

    RememberFolder ExtractFolderFromPath(firstFile)
End Sub

'@label:remember-folder-from-folders
'@sub-title Persist the first folder from a folder selection
Private Sub RememberFolderFromFolders(ByVal folderSelection As Variant)
    Dim firstFolder As String

    'Persist the first folder returned by the picker, if any.
    firstFolder = FirstStringOrEmpty(folderSelection)
    If firstFolder = vbNullString Then
        Exit Sub
    End If

    RememberFolder firstFolder
End Sub

'@label:resolve-last-known-folder
'@sub-title Resolve the last-used folder from cache or workbook storage
Private Function ResolveLastKnownFolder() As String
    'Favor the in-memory cache, falling back to workbook storage when needed.
    If this.lastKnownFolder <> vbNullString Then
        ResolveLastKnownFolder = this.lastKnownFolder
        Exit Function
    End If

    this.lastKnownFolder = NormalizeFolderPath(ReadLastFolderFromWorkbook())
    ResolveLastKnownFolder = this.lastKnownFolder
End Function

'@label:has-workbook-context
'@sub-title Check whether ThisWorkbook is accessible
Private Function HasWorkbookContext() As Boolean
    Dim workbookName As String

    'Some host contexts (e.g. unit tests) do not expose ThisWorkbook; detect it.
    On Error Resume Next
    workbookName = ThisWorkbook.Name
    HasWorkbookContext = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'@label:read-last-folder-from-workbook
'@sub-title Read the last-used folder from a workbook-level Name
Private Function ReadLastFolderFromWorkbook() As String
    Dim storedName As Name
    Dim refersToValue As String

    ReadLastFolderFromWorkbook = vbNullString

    'Only attempt to read the scoped Name when running inside Excel.
    If Not HasWorkbookContext() Then
        Exit Function
    End If

    On Error Resume Next
    Set storedName = ThisWorkbook.Names(LAST_FOLDER_STORAGE_NAME)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If storedName Is Nothing Then
        Exit Function
    End If

    'Names that survive workbook saves can still raise if the RefersTo is invalid.
    On Error Resume Next
    refersToValue = storedName.RefersTo
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ReadLastFolderFromWorkbook = DecodeWorkbookStorageValue(refersToValue)
End Function

'@label:write-last-folder-to-workbook
'@sub-title Persist the folder path in a hidden workbook Name
Private Sub WriteLastFolderToWorkbook(ByVal folderPath As String)
    Dim storedName As Name
    Dim storageValue As String

    If folderPath = vbNullString Then
        Exit Sub
    End If

    'Persist the folder path in a hidden Name so future sessions remember it.
    If Not HasWorkbookContext() Then
        Exit Sub
    End If

    storageValue = EncodeWorkbookStorageValue(folderPath)

    On Error Resume Next
    Set storedName = ThisWorkbook.Names(LAST_FOLDER_STORAGE_NAME)
    If Err.Number <> 0 Then
        Err.Clear
        Set storedName = Nothing
    End If
    On Error GoTo 0

    If storedName Is Nothing Then
        On Error Resume Next
        ThisWorkbook.Names.Add Name:=LAST_FOLDER_STORAGE_NAME, _
                               RefersTo:=storageValue, _
                               Visible:=False
        Err.Clear
        On Error GoTo 0
    Else
        On Error Resume Next
        storedName.RefersTo = storageValue
        storedName.Visible = False
        Err.Clear
        On Error GoTo 0
    End If
End Sub

'@label:encode-workbook-storage-value
'@sub-title Convert a folder path into a Name.RefersTo-safe string literal
Private Function EncodeWorkbookStorageValue(ByVal folderPath As String) As String
    Dim sanitized As String

    'Convert a raw folder path into a string literal safe for Name.RefersTo.
    sanitized = Replace(folderPath, vbCr, vbNullString)
    sanitized = Replace(sanitized, vbLf, vbNullString)
    sanitized = Replace(sanitized, Chr(34), Chr(34) & Chr(34))

    EncodeWorkbookStorageValue = "=" & Chr(34) & sanitized & Chr(34)
End Function

'@label:decode-workbook-storage-value
'@sub-title Undo Name serialization to recover the original folder path
Private Function DecodeWorkbookStorageValue(ByVal storedValue As String) As String
    Dim valueText As String

    'Undo the Name serialization, unwrapping ="..." and escaped quotes.
    valueText = Trim$(storedValue)
    If Len(valueText) >= 3 Then
        If Left$(valueText, 2) = "=" & Chr(34) And Right$(valueText, 1) = Chr(34) Then
            valueText = Mid$(valueText, 3, Len(valueText) - 3)
        End If
    End If

    valueText = Replace(valueText, Chr(34) & Chr(34), Chr(34))
    DecodeWorkbookStorageValue = valueText
End Function

'@label:wrap-apple-script-string
'@sub-title Escape a string for safe insertion into AppleScript
Private Function WrapAppleScriptString(ByVal value As String) As String
    Dim sanitized As String

    'Escape newline and quote characters for safe insertion into AppleScript.
    sanitized = Replace(value, vbCr, vbNullString)
    sanitized = Replace(sanitized, vbLf, vbNullString)
    sanitized = Replace(sanitized, Chr(34), Chr(92) & Chr(34))

    WrapAppleScriptString = Chr(34) & sanitized & Chr(34)
End Function

'@label:wrap-apple-script-alias
'@sub-title Build an AppleScript alias literal from a colon path
Private Function WrapAppleScriptAlias(ByVal aliasPath As String) As String
    'Build the alias literal that AppleScript expects for default location.
    WrapAppleScriptAlias = "alias " & WrapAppleScriptString(aliasPath)
End Function

'@label:build-apple-script-list
'@sub-title Translate a VBA array into an AppleScript list literal
Private Function BuildAppleScriptList(ByVal entries As Variant) As String
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim builder As String

    'Translate a VBA array into an AppleScript list literal.
    If Not TryGetArrayBounds(entries, lowerBound, upperBound) Then
        BuildAppleScriptList = "{}"
        Exit Function
    End If

    For index = lowerBound To upperBound
        If builder <> vbNullString Then
            builder = builder & ","
        End If
        builder = builder & WrapAppleScriptString(CStr(entries(index)))
    Next index

    BuildAppleScriptList = "{" & builder & "}"
End Function

'@label:build-mac-filter-list
'@sub-title Construct the UTI list for the AppleScript file picker
Private Function BuildMacFilterList(ByVal filters As String, _
                                    ByVal customMacFilter As String) As String
    Dim filterEntries As Variant

    'Construct the UTI list consumed by the AppleScript choose-file command.
    If customMacFilter <> vbNullString Then
        BuildMacFilterList = BuildAppleScriptList(Array(customMacFilter))
        Exit Function
    End If

    Select Case filters
        Case "*.xls"
            filterEntries = Array("com.microsoft.Excel.xls")
        Case "*.xlsx"
            filterEntries = Array("org.openxmlformats.spreadsheetml.sheet")
        Case "*.xlsb"
            filterEntries = Array("com.microsoft.Excel.sheet.binary.macroenabled")
        Case "*.xlsb, *.xlsx"
            filterEntries = Array("org.openxmlformats.spreadsheetml.sheet", _
                                  "com.microsoft.Excel.sheet.binary.macroenabled")
        Case Else
            filterEntries = Array("org.openxmlformats.spreadsheetml.sheet")
    End Select

    BuildMacFilterList = BuildAppleScriptList(filterEntries)
End Function

'@label:parse-mac-selection-result
'@sub-title Split newline-delimited POSIX paths from AppleScript output
Private Function ParseMacSelectionResult(ByVal response As String) As Variant
    Dim cleaned As String
    Dim parts() As String
    Dim results() As String
    Dim count As Long
    Dim index As Long
    Dim entry As String

    'MacScript returns newline-delimited POSIX paths; split and trim them.
    cleaned = Trim$(response)
    cleaned = Replace(cleaned, vbCr, vbNullString)

    If cleaned = vbNullString Then
        ParseMacSelectionResult = EmptyStringArray()
        Exit Function
    End If

    parts = Split(cleaned, Chr$(10))

    For index = LBound(parts) To UBound(parts)
        entry = Trim$(parts(index))
        If entry <> vbNullString Then
            If count = 0 Then
                ReDim results(0 To 0)
            Else
                ReDim Preserve results(0 To count)
            End If
            results(count) = entry
            count = count + 1
        End If
    Next index

    If count = 0 Then
        ParseMacSelectionResult = EmptyStringArray()
    Else
        ParseMacSelectionResult = results
    End If
End Function

'@label:normalize-folder-selections
'@sub-title Normalize picker results into clean folder paths
Private Function NormalizeFolderSelections(ByVal folderSelection As Variant) As Variant
    Dim selection As Variant
    Dim normalized() As String
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim index As Long
    Dim valueText As String
    Dim count As Long

    'Convert picker results into normalized folder paths and drop blanks.
    selection = ToStringArray(folderSelection)

    If Not TryGetArrayBounds(selection, lowerBound, upperBound) Then
        NormalizeFolderSelections = EmptyStringArray()
        Exit Function
    End If

    For index = lowerBound To upperBound
        valueText = NormalizeFolderPath(selection(index))
        If valueText <> vbNullString Then
            If count = 0 Then
                ReDim normalized(0 To 0)
            Else
                ReDim Preserve normalized(0 To count)
            End If
            normalized(count) = valueText
            count = count + 1
        End If
    Next index

    If count = 0 Then
        NormalizeFolderSelections = EmptyStringArray()
    Else
        NormalizeFolderSelections = normalized
    End If
End Function

'@label:convert-posix-to-colon-path
'@sub-title Translate a POSIX path to AppleScript colon-delimited format
Private Function ConvertPosixToColonPath(ByVal posixPath As String) As String
    Dim script As String
    Dim result As String

    'Translate a POSIX path into the colon-delimited format AppleScript prefers.
    If posixPath = vbNullString Then
        Exit Function
    End If

    script = "return POSIX file " & WrapAppleScriptString(posixPath) & " as string"

    On Error Resume Next
    result = MacScript(script)
    If Err.Number <> 0 Then
        Err.Clear
        result = vbNullString
    End If
    On Error GoTo 0

    ConvertPosixToColonPath = result
End Function

'@label:get-mac-default-folder-alias
'@sub-title Resolve the default folder alias for macOS pickers
Private Function GetMacDefaultFolderAlias() As String
    Dim storedFolder As String
    Dim aliasPath As String

    'Base the default location on the last selected folder when possible.
    storedFolder = ResolveLastKnownFolder()
    If storedFolder <> vbNullString Then
        If InStr(storedFolder, "/") > 0 Then
            aliasPath = ConvertPosixToColonPath(storedFolder)
            If aliasPath <> vbNullString Then
                GetMacDefaultFolderAlias = aliasPath
                Exit Function
            End If
        End If
    End If

    On Error Resume Next
    GetMacDefaultFolderAlias = MacScript("return (path to documents folder) as string")
    If Err.Number <> 0 Then
        Err.Clear
        GetMacDefaultFolderAlias = vbNullString
    End If
    On Error GoTo 0
End Function

'@label:build-mac-file-selection-script
'@sub-title Compose the AppleScript for file selection on macOS
Private Function BuildMacFileSelectionScript(ByVal filters As String, _
                                             ByVal customMacFilter As String, _
                                             ByVal defaultAlias As String, _
                                             ByVal allowMultiple As Boolean) As String
    Dim multipleFlag As String
    Dim defaultClause As String
    Dim filterList As String

    'Compose the AppleScript snippet that shows a file picker on macOS Excel.
    multipleFlag = IIf(allowMultiple, "true", "false")
    filterList = BuildMacFilterList(filters, customMacFilter)

    If defaultAlias <> vbNullString Then
        defaultClause = "default location " & WrapAppleScriptAlias(defaultAlias) & " "
    Else
        defaultClause = vbNullString
    End If

    BuildMacFileSelectionScript = _
        "set allowMultipleSelection to " & multipleFlag & vbNewLine & _
        "set chosenItems to (choose file of type " & filterList & _
        " with prompt " & WrapAppleScriptString("Please select a file or files") & " " & _
        defaultClause & "multiple selections allowed allowMultipleSelection)" & vbNewLine & _
        "if class of chosenItems is not list then set chosenItems to {chosenItems}" & vbNewLine & _
        "set posixPaths to {}" & vbNewLine & _
        "repeat with anItem in chosenItems" & vbNewLine & _
        "set end of posixPaths to POSIX path of anItem" & vbNewLine & _
        "end repeat" & vbNewLine & _
        "set applescript's text item delimiters to ASCII character 10" & vbNewLine & _
        "return posixPaths as string"
End Function

'@label:build-mac-folder-selection-script
'@sub-title Compose the AppleScript for folder selection on macOS
Private Function BuildMacFolderSelectionScript(ByVal defaultAlias As String, _
                                               ByVal allowMultiple As Boolean) As String
    Dim multipleFlag As String
    Dim defaultClause As String

    'Compose the AppleScript snippet that asks for one or more folders.
    multipleFlag = IIf(allowMultiple, "true", "false")

    If defaultAlias <> vbNullString Then
        defaultClause = "default location " & WrapAppleScriptAlias(defaultAlias) & " "
    Else
        defaultClause = vbNullString
    End If

    BuildMacFolderSelectionScript = _
        "set allowMultipleSelection to " & multipleFlag & vbNewLine & _
        "set chosenItems to (choose folder with prompt " & WrapAppleScriptString("Select the folder") & " " & _
        defaultClause & "multiple selections allowed allowMultipleSelection)" & vbNewLine & _
        "if class of chosenItems is not list then set chosenItems to {chosenItems}" & vbNewLine & _
        "set posixPaths to {}" & vbNewLine & _
        "repeat with anItem in chosenItems" & vbNewLine & _
        "set end of posixPaths to POSIX path of anItem" & vbNewLine & _
        "end repeat" & vbNewLine & _
        "set applescript's text item delimiters to ASCII character 10" & vbNewLine & _
        "return posixPaths as string"
End Function

'@label:run-mac-file-picker
'@sub-title Execute the AppleScript file picker and return POSIX paths
Private Function RunMacFilePicker(ByVal filters As String, _
                                  ByVal customMacFilter As String, _
                                  ByVal allowMultiple As Boolean) As Variant
    Dim script As String
    Dim scriptResult As String
    Dim defaultAlias As String

    'Execute the AppleScript picker and return the chosen POSIX paths.
    defaultAlias = GetMacDefaultFolderAlias()
    script = BuildMacFileSelectionScript(filters, customMacFilter, defaultAlias, allowMultiple)

    On Error Resume Next
    scriptResult = MacScript(script)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        RunMacFilePicker = EmptyStringArray()
        Exit Function
    End If
    On Error GoTo 0

    RunMacFilePicker = ToStringArray(ParseMacSelectionResult(scriptResult))
End Function

'@label:run-mac-folder-picker
'@sub-title Execute the AppleScript folder picker and return POSIX paths
Private Function RunMacFolderPicker(ByVal allowMultiple As Boolean) As Variant
    Dim script As String
    Dim scriptResult As String
    Dim defaultAlias As String

    'Use AppleScript to collect folder selections, normalized for VBA usage.
    defaultAlias = GetMacDefaultFolderAlias()
    script = BuildMacFolderSelectionScript(defaultAlias, allowMultiple)

    On Error Resume Next
    scriptResult = MacScript(script)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        RunMacFolderPicker = EmptyStringArray()
        Exit Function
    End If
    On Error GoTo 0

    RunMacFolderPicker = NormalizeFolderSelections(ParseMacSelectionResult(scriptResult))
End Function


'@label:create
'@sub-title Factory returning an initialised interface
'@return IOSFiles. Ready to use.
'@export
Public Function Create() As IOSFiles

    'Will run on Mac or windows operating system
    Dim currentOS As String
    'OperatingSystem string differs per platform; we only need a simple flag.
    currentOS = IIf(Application.OperatingSystem Like "*Mac*", "Mac", "Windows")
  
    With New OSFiles
        'Store the detected OS before handing back the interface implementation.
        .OS = currentOS
        Set Create = .Self
    End With
End Function

'@section Elements
'===============================================================================
'File and folder path accessors.

'@label:self
'@prop-title Self-reference as IOSFiles
'@return IOSFiles. Self-reference.
Public Property Get Self() As IOSFiles
    'Return the class instance as the interface the consumer expects.
    Set Self = Me
End Property

'@label:os
'@prop-title Current operating system flag
'@return String. "Mac" or "Windows".
Public Property Get OS() As String
    OS = this.currentOS
End Property

'@label:os-set
'@prop-title Override the detected OS flag
'@param currentOS String. "Mac" or "Windows".
Public Property Let OS(ByVal currentOS As String)
    this.currentOS = currentOS
End Property

'@label:file
'@prop-title Current loaded file path
'@return String. The first selected file path.
'@export
Private Property Get File() As String
    File = this.currentLoadedFile
End Property

'@label:files
'@prop-title All loaded file paths from a multi-selection
'@return Variant. Zero-based string array of file paths.
'@export
Private Property Get Files() As Variant
    Files = CloneStringArray(this.currentLoadedFiles)
End Property

'@label:folder
'@prop-title Current loaded folder path
'@return String. The first selected folder path.
'@export
Private Property Get Folder() As String
    Folder = this.currentLoadedFolder
End Property

'@label:folders
'@prop-title All loaded folder paths from a multi-selection
'@return Variant. Zero-based string array of folder paths.
'@export
Private Property Get Folders() As Variant
    Folders = CloneStringArray(this.currentLoadedFolders)
End Property

'@label:file-set
'@prop-title Assign a single file path through the selection pipeline
'@param ff String. The file path.
Private Property Let File(ByVal ff As String)
    SetFilesSelection ff
End Property

'@label:folder-set
'@prop-title Assign a single folder path through the selection pipeline
'@param ff String. The folder path.
Private Property Let Folder(ByVal ff As String)
    SetFoldersSelection ff
End Property

'@label:select-folder-on-mac
'@sub-title Show the macOS single-folder picker
Private Function SelectFolderOnMac() As String
    Dim selection As Variant

    'Ask macOS to pick a single folder, then unwrap the first choice.
    selection = RunMacFolderPicker(False)
    SelectFolderOnMac = FirstStringOrEmpty(selection)
End Function

'@label:select-folders-on-mac
'@sub-title Show the macOS multi-folder picker
Private Function SelectFoldersOnMac() As Variant
    'Allow macOS users to return multiple folders in one interaction.
    SelectFoldersOnMac = RunMacFolderPicker(True)
End Function

'@label:select-folder-on-windows
'@sub-title Show the Windows single-folder picker
Private Function SelectFolderOnWindows() As String

    Dim fDialog As Office.FileDialog

    SelectFolderOnWindows = vbNullString

    'Windows relies on Excel's FileDialog implementation to browse directories.
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .AllowMultiSelect = False
        .title = "Chose your directory"          'MSG_ChooseDir
        .filters.Clear

        If .Show = -1 Then
            SelectFolderOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'@label:select-files-on-windows
'@sub-title Show the Windows multi-file picker
Private Function SelectFilesOnWindows(ByVal filters As String) As Variant

    Dim fDialog As Office.FileDialog
    Dim selectedCount As Long
    Dim results() As String
    Dim index As Long

    SelectFilesOnWindows = EmptyStringArray()

    'Use the Windows file picker with multi-select capability enabled.
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = True
        .title = "Chose your files"              'MSG_ChooseFiles
        .filters.Clear
        .filters.Add "Excel Sheet", filters

        If .Show = True Then
            On Error Resume Next
            selectedCount = .SelectedItems.Count
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0

            If selectedCount > 0 Then
                ReDim results(0 To selectedCount - 1)
                For index = 1 To selectedCount
                    results(index - 1) = CStr(.SelectedItems(index))
                Next index
                'Copy the selections into a standard VBA array for consumers.
                SelectFilesOnWindows = results
            End If
        End If
    End With
End Function

'@label:select-file-on-mac
'@sub-title Show the macOS single-file picker
Private Function SelectFileOnMac(ByVal filters As String, _ 
                                 Optional ByVal customMacFilter As String = vbNullString) As String

    Dim selection As Variant

    'Leverage the macOS picker (with optional custom UTIs) to get one path.
    selection = RunMacFilePicker(filters, customMacFilter, False)
    SelectFileOnMac = FirstStringOrEmpty(selection)
End Function

'@label:select-files-on-mac
'@sub-title Show the macOS multi-file picker
Private Function SelectFilesOnMac(ByVal filters As String, _
                                  Optional ByVal customMacFilter As String = vbNullString) As Variant
    'Collect multiple POSIX file paths when the AppleScript picker allows it.
    SelectFilesOnMac = RunMacFilePicker(filters, customMacFilter, True)
End Function

'@label:select-file-on-windows
'@sub-title Show the Windows single-file picker
Private Function SelectFileOnWindows(ByVal filters As String) As String

    Dim fDialog As Office.FileDialog

    SelectFileOnWindows = vbNullString

    'Surface the Windows file picker configured with the supplied filters.
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .title = "Chose your file"               'MSG_ChooseFile
        .filters.Clear
        .filters.Add "Excel Sheet", filters      '"*.xlsx" ', *.xlsm, *.xlsb,  *.xls" 'MSG_ExcelFile

        If .Show = True Then
            SelectFileOnWindows = .SelectedItems(1)
        End If
    End With

End Function

'@section Selecting
'===============================================================================
'File and folder picker operations.

'@label:load-folder
'@sub-title Show a single-folder picker dialog
'@export
Private Sub LoadFolder()
    Dim selectedFolder As String

    'Delegate to the correct picker based on runtime OS and Excel version.
    If OS() = "Windows" Then
        'We are on windows DOS
        selectedFolder = SelectFolderOnWindows()
    ElseIf OS() = "Mac" Then
        'We are on Mac, need to test the version of excel running
        If val(Application.Version) > 14 Then
            selectedFolder = SelectFolderOnMac()
        End If
    End If

    Folder = selectedFolder
    'Store the resolved folder so subsequent pickers default to it.
    RememberFolderFromFolders this.currentLoadedFolders
End Sub


'@label:load-file
'@sub-title Show a single-file picker dialog
'@param filters String. File type filters (e.g. "*.xlsx").
'@param customMacFilter String. Optional UTI override for macOS. Defaults to vbNullString.
'@export
Private Sub LoadFile(ByVal filters As String, _ 
                     Optional ByVal customMacFilter As String = vbNullString)
    Dim selectedFile As String

    'Use the appropriate single-file picker to populate File().
    If OS() = "Windows" Then
        'We are on windows DOS
        selectedFile = SelectFileOnWindows(filters)
    ElseIf OS() = "Mac" Then
        'We are on Mac, need to test the version of excel running
        If val(Application.Version) > 14 Then
            selectedFile = SelectFileOnMac(filters, customMacFilter)
        End If
    End If

    File = selectedFile
    'Remember the parent folder for future browsing sessions.
    RememberFolderFromFiles this.currentLoadedFiles
End Sub

'@label:load-files
'@sub-title Show a multi-file picker dialog
'@param filters String. File type filters.
'@param customMacFilter String. Optional UTI override for macOS. Defaults to vbNullString.
'@export
Private Sub LoadFiles(ByVal filters As String, _
                      Optional ByVal customMacFilter As String = vbNullString)

    Dim selection As Variant
    selection = EmptyStringArray()

    'Call the multi-select picker that matches the host and capture all results.
    If OS() = "Windows" Then
        selection = SelectFilesOnWindows(filters)
    ElseIf OS() = "Mac" Then
        If val(Application.Version) > 14 Then
            selection = SelectFilesOnMac(filters, customMacFilter)
        End If
    End If

    SetFilesSelection selection
    'Persist the originating folder for convenience in subsequent sessions.
    RememberFolderFromFiles this.currentLoadedFiles
End Sub

'@label:load-folders
'@sub-title Show a multi-folder picker dialog
'@export
Private Sub LoadFolders()

    Dim selection As Variant
    selection = EmptyStringArray()

    'Fan out to the appropriate multi-folder picker and store the results.
    If OS() = "Windows" Then
        selection = SelectFoldersOnWindows()
    ElseIf OS() = "Mac" Then
        If val(Application.Version) > 14 Then
            selection = SelectFoldersOnMac()
        End If
    End If

    SetFoldersSelection selection
    'Persist the first folder chosen so the next run opens in that location.
    RememberFolderFromFolders this.currentLoadedFolders
End Sub

'@section Iteration
'===============================================================================
'Cursor-based enumeration over multi-selections.

'@label:reset-files-iterator
'@sub-title Reset the file iteration cursor to the beginning
'@export
Private Sub ResetFilesCursor()
    'Reset iteration so external callers can enumerate from the first file.
    this.currentFileIndex = 0
End Sub

'@label:reset-folders-iterator
'@sub-title Reset the folder iteration cursor to the beginning
'@export
Private Sub ResetFoldersCursor()
    'Reset iteration so external callers can enumerate from the first folder.
    this.currentFolderIndex = 0
End Sub

'@label:has-next-file
'@sub-title Check if another file is available in the iterator
'@export
Private Function HasNextFileValue() As Boolean
    'Expose whether the internal file iterator has another entry.
    HasNextFileValue = HasNextArrayItem(this.currentLoadedFiles, this.currentFileIndex)
End Function

'@label:has-next-folder
'@sub-title Check if another folder is available in the iterator
'@export
Private Function HasNextFolderValue() As Boolean
    'Expose whether the internal folder iterator has another entry.
    HasNextFolderValue = HasNextArrayItem(this.currentLoadedFolders, this.currentFolderIndex)
End Function

'@label:next-file
'@sub-title Return the next file path and advance the cursor
'@export
Private Function NextFileValue() As String
    'Return the next file path in sequence and advance the cursor.
    NextFileValue = GetNextArrayItem(this.currentLoadedFiles, this.currentFileIndex)
End Function

'@label:next-folder
'@sub-title Return the next folder path and advance the cursor
'@export
Private Function NextFolderValue() As String
    'Return the next folder path in sequence and advance the cursor.
    NextFolderValue = GetNextArrayItem(this.currentLoadedFolders, this.currentFolderIndex)
End Function

'@section Test Support
'===============================================================================
'Hooks for injecting selections from unit tests.

'@label:assign-files-for-testing
'@sub-title Inject file selections without showing a dialog
'@param selections Variant. File paths to assign.
'@TestHook
Public Sub AssignFilesForTesting(ByVal selections As Variant)
    'Bypass the dialogs and inject file selections from unit tests.
    SetFilesSelection selections
End Sub

'@label:assign-folders-for-testing
'@sub-title Inject folder selections without showing a dialog
'@param selections Variant. Folder paths to assign.
'@TestHook
Public Sub AssignFoldersForTesting(ByVal selections As Variant)
    'Bypass the dialogs and inject folder selections from unit tests.
    SetFoldersSelection selections
End Sub

'@section Validation
'===============================================================================
'Selection validity checks.

'@label:has-valid-file
'@sub-title Whether the loaded file path is valid
'@return Boolean. True when a non-empty file path exists.
'@export
Private Function HasValidFile() As Boolean
    'A simple sanity check used by clients before acting on File().
    HasValidFile = (File() <> vbNullString)
End Function

'@label:has-valid-files
'@sub-title Whether the loaded files array contains entries
'@return Boolean. True when at least one file was selected.
'@export
Private Function HasValidFiles() As Boolean
    'Report whether any file selections exist in the internal buffer.
    HasValidFiles = (CountArrayItems(this.currentLoadedFiles) > 0)
End Function

'@label:has-valid-folder
'@sub-title Whether the loaded folder path is valid
'@return Boolean. True when a non-empty folder path exists.
'@export
Private Function HasValidFolder() As Boolean
    'A simple sanity check used by clients before acting on Folder().
    HasValidFolder = (Folder() <> vbNullString)
End Function

'@label:has-valid-folders
'@sub-title Whether the loaded folders array contains entries
'@return Boolean. True when at least one folder was selected.
'@export
Private Function HasValidFolders() As Boolean
    'Report whether any folder selections exist in the internal buffer.
    HasValidFolders = (CountArrayItems(this.currentLoadedFolders) > 0)
End Function


'@section Interface
'===============================================================================
'IOSFiles delegation wiring.

Private Property Get IOSFiles_File() As String
    IOSFiles_File = File()
End Property

Private Property Get IOSFiles_Files() As Variant
    IOSFiles_Files = Files()
End Property

Private Property Get IOSFiles_Folder() As String
    IOSFiles_Folder = Folder()
End Property

Private Property Get IOSFiles_Folders() As Variant
    IOSFiles_Folders = Folders()
End Property

Private Sub IOSFiles_LoadFolder()
    LoadFolder
End Sub

Private Sub IOSFiles_LoadFolders()
    LoadFolders
End Sub

Private Sub IOSFiles_ResetFilesIterator()
    ResetFilesCursor
End Sub

Private Function IOSFiles_HasNextFile() As Boolean
    IOSFiles_HasNextFile = HasNextFileValue()
End Function

Private Function IOSFiles_NextFile() As String
    IOSFiles_NextFile = NextFileValue()
End Function

Private Sub IOSFiles_ResetFoldersIterator()
    ResetFoldersCursor
End Sub

Private Function IOSFiles_HasNextFolder() As Boolean
    IOSFiles_HasNextFolder = HasNextFolderValue()
End Function

Private Function IOSFiles_NextFolder() As String
    IOSFiles_NextFolder = NextFolderValue()
End Function

Private Sub IOSFiles_LoadFile( _
    ByVal filters As String, _
    Optional ByVal customMacFilter As String = vbNullString)
    LoadFile filters:=filters, customMacFilter:=customMacFilter
End Sub

Private Sub IOSFiles_LoadFiles( _
    ByVal filters As String, _
    Optional ByVal customMacFilter As String = vbNullString)
    LoadFiles filters:=filters, customMacFilter:=customMacFilter
End Sub

Private Function IOSFiles_HasValidFile() As Boolean
    IOSFiles_HasValidFile = HasValidFile()
End Function

Private Function IOSFiles_HasValidFiles() As Boolean
    IOSFiles_HasValidFiles = HasValidFiles()
End Function

Private Function IOSFiles_HasValidFolder() As Boolean
    IOSFiles_HasValidFolder = HasValidFolder()
End Function

Private Function IOSFiles_HasValidFolders() As Boolean
    IOSFiles_HasValidFolders = HasValidFolders()
End Function

