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

'@label:initialization
'@section Initialization helpers
'===============================================================================

Private Sub Class_Initialize()
    'Ensure every new instance starts with clean selections and no cached folder.
    ResetFileState
    ResetFolderState
    this.lastKnownFolder = vbNullString
End Sub

Private Sub ResetFileState()
    'Clear the current file-related state and reset the iteration cursor.
    this.currentLoadedFile = vbNullString
    this.currentLoadedFiles = EmptyStringArray()
    this.currentFileIndex = 0
End Sub

Private Sub ResetFolderState()
    'Clear the current folder-related state and reset the iteration cursor.
    this.currentLoadedFolder = vbNullString
    this.currentLoadedFolders = EmptyStringArray()
    this.currentFolderIndex = 0
End Sub

Private Function EmptyStringArray() As Variant
    'Return an empty variant array that VBA callers can treat like a string list.
    EmptyStringArray = Array()
End Function

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

Private Function CountArrayItems(ByVal candidate As Variant) As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    'Translate array bounds into a count; returns 0 for unsupported inputs.
    If Not TryGetArrayBounds(candidate, lowerBound, upperBound) Then
        Exit Function
    End If

    CountArrayItems = (upperBound - lowerBound) + 1
End Function

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

Private Function CloneStringArray(ByVal source As Variant) As Variant
    'Use ToStringArray to guarantee a safe copy that call sites cannot mutate.
    CloneStringArray = ToStringArray(source)
End Function

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

Private Sub SetFilesSelection(ByVal selectedItems As Variant)
    Dim normalized As Variant
    normalized = ToStringArray(selectedItems)

    'Persist the user selection and prime the iterator against the first entry.
    this.currentLoadedFiles = normalized
    this.currentLoadedFile = FirstStringOrEmpty(normalized)
    this.currentFileIndex = 0
End Sub

Private Sub SetFoldersSelection(ByVal selectedItems As Variant)
    Dim normalized As Variant
    normalized = ToStringArray(selectedItems)

    'Store folder selections and make the first path immediately available.
    this.currentLoadedFolders = normalized
    this.currentLoadedFolder = FirstStringOrEmpty(normalized)
    this.currentFolderIndex = 0
End Sub

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

Private Sub RememberFolderFromFiles(ByVal filesSelection As Variant)
    Dim firstFile As String

    'Persist the directory that contains the first selected file.
    firstFile = FirstStringOrEmpty(filesSelection)
    If firstFile = vbNullString Then
        Exit Sub
    End If

    RememberFolder ExtractFolderFromPath(firstFile)
End Sub

Private Sub RememberFolderFromFolders(ByVal folderSelection As Variant)
    Dim firstFolder As String

    'Persist the first folder returned by the picker, if any.
    firstFolder = FirstStringOrEmpty(folderSelection)
    If firstFolder = vbNullString Then
        Exit Sub
    End If

    RememberFolder firstFolder
End Sub

Private Function ResolveLastKnownFolder() As String
    'Favor the in-memory cache, falling back to workbook storage when needed.
    If this.lastKnownFolder <> vbNullString Then
        ResolveLastKnownFolder = this.lastKnownFolder
        Exit Function
    End If

    this.lastKnownFolder = NormalizeFolderPath(ReadLastFolderFromWorkbook())
    ResolveLastKnownFolder = this.lastKnownFolder
End Function

Private Function HasWorkbookContext() As Boolean
    Dim workbookName As String

    'Some host contexts (e.g. unit tests) do not expose ThisWorkbook; detect it.
    On Error Resume Next
    workbookName = ThisWorkbook.Name
    HasWorkbookContext = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

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

Private Function EncodeWorkbookStorageValue(ByVal folderPath As String) As String
    Dim sanitized As String

    'Convert a raw folder path into a string literal safe for Name.RefersTo.
    sanitized = Replace(folderPath, vbCr, vbNullString)
    sanitized = Replace(sanitized, vbLf, vbNullString)
    sanitized = Replace(sanitized, Chr(34), Chr(34) & Chr(34))

    EncodeWorkbookStorageValue = "=" & Chr(34) & sanitized & Chr(34)
End Function

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

Private Function WrapAppleScriptString(ByVal value As String) As String
    Dim sanitized As String

    'Escape newline and quote characters for safe insertion into AppleScript.
    sanitized = Replace(value, vbCr, vbNullString)
    sanitized = Replace(sanitized, vbLf, vbNullString)
    sanitized = Replace(sanitized, Chr(34), Chr(92) & Chr(34))

    WrapAppleScriptString = Chr(34) & sanitized & Chr(34)
End Function

Private Function WrapAppleScriptAlias(ByVal aliasPath As String) As String
    'Build the alias literal that AppleScript expects for default location.
    WrapAppleScriptAlias = "alias " & WrapAppleScriptString(aliasPath)
End Function

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


'@label:Create
'@sub-title Create an OSFile object
'@details
'OSFile is a class to import files from different operation systems 
'(Mac or Windows). 
'The main advantage of OSFiles is that it could work on both Macos and Windows
'The allowed filters are the following: "*.xls", "*.xlsx", "*.xlsb", "*.xlsb, *.xlsx"
'If you want to use more filters, please remember to update the customMacFilter 
'argument of `LoadFile` or `LoadFolder`. New filters like .csv may be accessible on 
'windows, but if the filter is not in the previous list it will default to xlsx files
'on Mac.
' Here is the list of filters you can use on Mac:
' - xls : com.microsoft.Excel.xls
' - xlsx : org.openxmlformats.spreadsheetml.sheet
' - xlsm : org.openxmlformats.spreadsheetml.sheet.macroenabled
' - xlsb : com.microsoft.Excel.sheet.binary.macroenabled
' - csv : public.comma-separated-values-text
' - doc : com.microsoft.word.doc
' - docx : org.openxmlformats.wordprocessingml.document
' - docm : org.openxmlformats.wordprocessingml.document.macroenabled
' - ppt : com.microsoft.powerpoint.ppt
' - pptx : org.openxmlformats.presentationml.presentation
' - pptm : org.openxmlformats.presentationml.presentation.macroenabled
' - txt : public.plain-text
' - pdf : com.adobe.pdf
' - jpg : public.jpeg
' - png : public.png
' - QIF : com.apple.traditional-mac-plain-text
' - htm : public.html
'
'Based on https://macexcel.com/examples/filesandfolders/selectfiles/.
'
'@return OSFile Object
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

'@label:elements
'@section OSFiles elements
'===============================================================================


Public Property Get Self() As IOSFiles
    'Return the class instance as the interface the consumer expects.
    Set Self = Me
End Property

Public Property Get OS() As String
    'Expose the current operating system flag (Mac or Windows).
    OS = this.currentOS
End Property

Public Property Let OS(ByVal currentOS As String)
    'Allow callers to override the detected OS, mainly for testing purposes.
    this.currentOS = currentOS
End Property

'@label:File
'@prop-title Get the current loaded file path
'@return A string, the file path
'@export
Private Property Get File() As String
    'Return the first file path captured by the last selection.
    File = this.currentLoadedFile
End Property

'@label:Files
'@prop-title Get the current loaded files paths
'@return A Variant array containing file paths
'@export
Private Property Get Files() As Variant
    'Return a copy of all file paths selected in the last picker session.
    Files = CloneStringArray(this.currentLoadedFiles)
End Property

'@label:Folder
'@prop-title Get the current loaded folder path
'@return A String, the folder path
'@export
Private Property Get Folder() As String
    'Return the first folder captured by the last selection.
    Folder = this.currentLoadedFolder
End Property

'@label:Folders
'@prop-title Get the current loaded folders paths
'@return A Variant array containing folder paths
'@export
Private Property Get Folders() As Variant
    'Return a copy of the selected folder paths without exposing internals.
    Folders = CloneStringArray(this.currentLoadedFolders)
End Property

Private Property Let File(ByVal ff As String)
    'Re-use the selection pipeline so validation remains consistent.
    SetFilesSelection ff
End Property

Private Property Let Folder(ByVal ff As String)
    'Normalize single folder assignments through the selection helper.
    SetFoldersSelection ff
End Property

'Functions to select a folder


'@prop-title Select a folder on MacOS
Private Function SelectFolderOnMac() As String
    Dim selection As Variant

    'Ask macOS to pick a single folder, then unwrap the first choice.
    selection = RunMacFolderPicker(False)
    SelectFolderOnMac = FirstStringOrEmpty(selection)
End Function

'@prop-title Select multiple folders on MacOS
'@return A Variant array containing POSIX folder paths
Private Function SelectFoldersOnMac() As Variant
    'Allow macOS users to return multiple folders in one interaction.
    SelectFoldersOnMac = RunMacFolderPicker(True)
End Function

'@prop-title Select a folder on windows OS
'@return A String, the path to the folder on windows
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

'@prop-title Select a file on MacOS
'@param filters A String, filters to use during selection. See Create
'@param customMacFilter Optional, a string. filter to use in case you want a specific filter on MacOS. See Create
'@return A String, the path to the selected file on MacOS
Private Function SelectFileOnMac(ByVal filters As String, _ 
                                 Optional ByVal customMacFilter As String = vbNullString) As String

    Dim selection As Variant

    'Leverage the macOS picker (with optional custom UTIs) to get one path.
    selection = RunMacFilePicker(filters, customMacFilter, False)
    SelectFileOnMac = FirstStringOrEmpty(selection)
End Function

'@prop-title Select multiple files on MacOS
'@return A Variant array containing POSIX file paths
Private Function SelectFilesOnMac(ByVal filters As String, _
                                  Optional ByVal customMacFilter As String = vbNullString) As Variant
    'Collect multiple POSIX file paths when the AppleScript picker allows it.
    SelectFilesOnMac = RunMacFilePicker(filters, customMacFilter, True)
End Function

'@prop-title Select a file on Windows OS
'@param filters A String, filters to use during selection. See Create
'@return A String, the path to the file on Windows
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

'@label:selecting
'@section selecting
'===============================================================================


'@label:LoadFolder
'@sub-title Select a folder
'@return A String, The file path to the folder
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


'@label:LoadFile
'@sub-title Select a file
'@details
'If the user aborts the selection process or something go wrong during the 
'interaction, the internal object will have an empty file. There is a property
'validate the selection process and make sur the file exists.
'
'@param filters A String, filters to use during selection. See Create
'@param customMacFilter Optional, a string. filter to use in case you want a specific filter on MacOS. See Create
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

'@label:LoadFiles
'@sub-title Select multiple files
'@details
'When multiple selection is enabled, the class stores all selected entries while
'preserving the first item for backward compatibility.
'@param filters A String, filters to use during selection.
'@param customMacFilter Optional, a string. Filter to use on macOS.
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

'@label:LoadFolders
'@sub-title Select multiple folders
'@details
'macOS support relies on AppleScript multi-selection; Windows uses FileDialog.
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

'@label:iteration
'@section Iteration helpers
'===============================================================================

Private Sub ResetFilesCursor()
    'Reset iteration so external callers can enumerate from the first file.
    this.currentFileIndex = 0
End Sub

Private Sub ResetFoldersCursor()
    'Reset iteration so external callers can enumerate from the first folder.
    this.currentFolderIndex = 0
End Sub

Private Function HasNextFileValue() As Boolean
    'Expose whether the internal file iterator has another entry.
    HasNextFileValue = HasNextArrayItem(this.currentLoadedFiles, this.currentFileIndex)
End Function

Private Function HasNextFolderValue() As Boolean
    'Expose whether the internal folder iterator has another entry.
    HasNextFolderValue = HasNextArrayItem(this.currentLoadedFolders, this.currentFolderIndex)
End Function

Private Function NextFileValue() As String
    'Return the next file path in sequence and advance the cursor.
    NextFileValue = GetNextArrayItem(this.currentLoadedFiles, this.currentFileIndex)
End Function

Private Function NextFolderValue() As String
    'Return the next folder path in sequence and advance the cursor.
    NextFolderValue = GetNextArrayItem(this.currentLoadedFolders, this.currentFolderIndex)
End Function

'@label:tests
'@section Test support
'===============================================================================

'@TestHook
Public Sub AssignFilesForTesting(ByVal selections As Variant)
    'Bypass the dialogs and inject file selections from unit tests.
    SetFilesSelection selections
End Sub

'@TestHook
Public Sub AssignFoldersForTesting(ByVal selections As Variant)
    'Bypass the dialogs and inject folder selections from unit tests.
    SetFoldersSelection selections
End Sub

'@label:validation
'@section Validation
'===============================================================================

'@label:HasValidFile
'@prop-title Test if the internal loaded file is valid
'@return A Boolean
'@export
Private Function HasValidFile() As Boolean
    'A simple sanity check used by clients before acting on File().
    HasValidFile = (File() <> vbNullString)
End Function

'@label:HasValidFiles
'@prop-title Test if the internal loaded files collection is valid
'@return A Boolean
'@export
Private Function HasValidFiles() As Boolean
    'Report whether any file selections exist in the internal buffer.
    HasValidFiles = (CountArrayItems(this.currentLoadedFiles) > 0)
End Function

'@label:HasValidFolder
'@prop-title Test if the internal loaded folder is valid
'@return A Boolean
'@export
Private Function HasValidFolder() As Boolean
    'A simple sanity check used by clients before acting on Folder().
    HasValidFolder = (Folder() <> vbNullString)
End Function

'@label:HasValidFolders
'@prop-title Test if the internal loaded folders collection is valid
'@return A Boolean
'@export
Private Function HasValidFolders() As Boolean
    'Report whether any folder selections exist in the internal buffer.
    HasValidFolders = (CountArrayItems(this.currentLoadedFolders) > 0)
End Function


'@jump:Create
'@section Interface
'===============================================================================

'These members fulfill the IOSFiles contract by delegating to the private logic.
'@jump:elements
'IOSFile elements

'@jump:File
Private Property Get IOSFiles_File() As String
    'IOSFiles: expose the single-file selection value.
    IOSFiles_File = File()
End Property

'@jump:Files
Private Property Get IOSFiles_Files() As Variant
    'IOSFiles: expose the multi-file selection collection.
    IOSFiles_Files = Files()
End Property

'@jump:Folder
Private Property Get IOSFiles_Folder() As String
    'IOSFiles: expose the single-folder selection value.
    IOSFiles_Folder = Folder()
End Property

'@jump:Folders
Private Property Get IOSFiles_Folders() As Variant
    'IOSFiles: expose the multi-folder selection collection.
    IOSFiles_Folders = Folders()
End Property

'@jump:selecting
'Selecting----------------------------------------------------------------------

'@jump:LoadFolder
Private Sub IOSFiles_LoadFolder()
    'IOSFiles: pick a single folder and track it internally.
    LoadFolder
End Sub

'@jump:LoadFolders
Private Sub IOSFiles_LoadFolders()
    'IOSFiles: pick multiple folders and track them internally.
    LoadFolders
End Sub

'@jump:ResetFilesIterator
Private Sub IOSFiles_ResetFilesIterator()
    'IOSFiles: allow consumers to restart file enumeration.
    ResetFilesCursor
End Sub

'@jump:HasNextFile
Private Function IOSFiles_HasNextFile() As Boolean
    'IOSFiles: check if another file path is queued.
    IOSFiles_HasNextFile = HasNextFileValue()
End Function

'@jump:NextFile
Private Function IOSFiles_NextFile() As String
    'IOSFiles: fetch the next file path from the buffer.
    IOSFiles_NextFile = NextFileValue()
End Function

'@jump:ResetFoldersIterator
Private Sub IOSFiles_ResetFoldersIterator()
    'IOSFiles: allow consumers to restart folder enumeration.
    ResetFoldersCursor
End Sub

'@jump:HasNextFolder
Private Function IOSFiles_HasNextFolder() As Boolean
    'IOSFiles: check if another folder path is queued.
    IOSFiles_HasNextFolder = HasNextFolderValue()
End Function

'@jump:NextFolder
Private Function IOSFiles_NextFolder() As String
    'IOSFiles: fetch the next folder path from the buffer.
    IOSFiles_NextFolder = NextFolderValue()
End Function

'@jump:LoadFile
Private Sub IOSFiles_LoadFile(ByVal filters As String, _ 
                             Optional ByVal customMacFilter As String = vbNullString)
    'IOSFiles: pick a single file using the supplied filters.
    LoadFile filters:=filters, customMacFilter:=customMacFilter
End Sub

'@jump:LoadFiles
Private Sub IOSFiles_LoadFiles(ByVal filters As String, _
                               Optional ByVal customMacFilter As String = vbNullString)
    'IOSFiles: pick multiple files using the supplied filters.
    LoadFiles filters:=filters, customMacFilter:=customMacFilter
End Sub

'@jump:validation
'Validate-----------------------------------------------------------------------

'@jump:HasValidFile
Private Function IOSFiles_HasValidFile() As Boolean
    'IOSFiles: mirror HasValidFile for interface consumers.
    IOSFiles_HasValidFile = HasValidFile()
End Function

'@jump:HasValidFiles
Private Function IOSFiles_HasValidFiles() As Boolean
    'IOSFiles: mirror HasValidFiles for interface consumers.
    IOSFiles_HasValidFiles = HasValidFiles()
End Function

'@jump:HasValidFolder
Private Function IOSFiles_HasValidFolder() As Boolean
    'IOSFiles: mirror HasValidFolder for interface consumers.
    IOSFiles_HasValidFolder = HasValidFolder()
End Function

'@jump:HasValidFolders
Private Function IOSFiles_HasValidFolders() As Boolean
    'IOSFiles: mirror HasValidFolders for interface consumers.
    IOSFiles_HasValidFolders = HasValidFolders()
End Function


 
