Attribute VB_Name = "TestHelpers"
Attribute VB_Description = "Utility helpers shared across tests"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@Folder("Tests")
'@ModuleDescription("Utility helpers shared across tests")

Private Const VBEXT_CT_STD_MODULE As Long = 1
Private Const VBEXT_CT_CLASS_MODULE As Long = 2
Private Const VBEXT_CT_DOCUMENT As Long = 100

'@section Application State
'===============================================================================

'@label BusyApp
'@sub-title Suspend heavy Excel UI features while tests manipulate workbooks.
'@details Suspend heavy Excel UI features while tests manipulate workbooks.
Public Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'@label RestoreApp
'@sub-title Restore the Excel UI to its default behaviour after BusyApp.
'@details Restore the Excel UI to its default behaviour after BusyApp.
Public Sub RestoreApp()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableAnimations = True
End Sub

'@section Workbooks
'===============================================================================

'@label NewWorkbook
'@fun-title Create a new workbook ready for test usage.
'@details Create a new workbook ready for test usage.
'@return Workbook freshly created.
Public Function NewWorkbook() As workbook
    BusyApp
    Set NewWorkbook = Workbooks.Add
    ActiveWindow.WindowState = xlMinimized
End Function

'@label DeleteWorkbook
'@sub-title Close and discard a workbook if it exists.
'@details Close and discard a workbook if it exists.
'@param wb Workbook or Object reference to close.
Public Sub DeleteWorkbook(ByVal wb As workbook)
    On Error Resume Next
        BusyApp
        wb.Close saveChanges:=False
    On Error GoTo 0
End Sub

'@section Worksheets
'===============================================================================

'@label EnsureWorksheet
'@fun-title Ensure a worksheet exists and is cleared.
'@details Ensure a worksheet exists and is cleared.
'@param sheetName String. Name of the worksheet to create/reset.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
'@return Worksheet ensured for use.
Public Function EnsureWorksheet(ByVal sheetName As String, _
                                Optional ByVal targetBook As workbook, _
                                Optional ByVal clearSheet As Boolean = True, _
                                Optional ByVal visibility As Long = xlSheetVisible) As Worksheet

    Dim wb As workbook
    Dim sh As Worksheet

    If (targetBook Is Nothing) Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    If sh Is Nothing Then
        BusyApp
        Set sh = wb.Worksheets.Add
        sh.Name = sheetName
    End If

    sh.Visible = visibility
    If clearSheet Then 
        ClearWorksheet sh
    End If
    
    Set EnsureWorksheet = sh
End Function

'@label NewWorksheet
'@sub-title Create a worksheet when missing and clear its cells.
'@details Create a worksheet when missing and clear its cells.
'@param sheetName String. Name of the worksheet to reset.
Public Sub NewWorksheet(ByVal sheetName As String)
    Call EnsureWorksheet(sheetName)
End Sub

'@label DeleteWorksheet
'@sub-title Delete a worksheet if it exists.
'@details Delete a worksheet if it exists.
'@param sheetName String. Worksheet to delete.
Public Sub DeleteWorksheet(ByVal sheetName As String)
    On Error Resume Next
        BusyApp
        ThisWorkbook.worksheets(sheetName).Visible = xlSheetVeryHidden
        ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
End Sub

'@label DeleteWorksheets
'@sub-title Delete several worksheets in a single call.
'@details Delete several worksheets in a single call.
'@param sheetNames ParamArray list of worksheet names.
Public Sub DeleteWorksheets(ParamArray sheetNames() As Variant)
    Dim idx As Long

    For idx = LBound(sheetNames) To UBound(sheetNames)
        DeleteWorksheet CStr(sheetNames(idx))
    Next idx
End Sub

'@label WorksheetExists
'@fun-title Test whether a worksheet exists in a workbook.
'@details Test whether a worksheet exists in a workbook.
'@param sheetName String. Name to look up.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
'@return Boolean indicating existence.
Public Function WorksheetExists(ByVal sheetName As String, _
                                Optional ByVal targetBook As workbook) As Boolean

    Dim wb As workbook
    Dim sh As Worksheet

    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    On Error Resume Next
        Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not (sh Is Nothing)
End Function

'@label ClearWorksheet
'@sub-title Remove data, tables, shapes and names from a worksheet.
'@details Remove data, tables, shapes and names from a worksheet.
'@param sh Worksheet to clear.
Public Sub ClearWorksheet(ByVal sh As Worksheet)

    Dim nm As Name

    If sh Is Nothing Then Exit Sub

    BusyApp

    On Error Resume Next
        Do While sh.ListObjects.Count > 0
            sh.ListObjects(1).Delete
        Loop

        Do While sh.Shapes.Count > 0
            sh.Shapes(1).Delete
        Loop

        For Each nm In sh.Names
            nm.Delete
        Next nm

        For Each nm In sh.Parent.Names
            If InStr(1, nm.RefersTo, "'" & sh.Name & "'!", vbTextCompare) > 0 Then nm.Delete
        Next nm

        sh.Cells.Clear
    On Error GoTo 0
End Sub

'@section Named Ranges
'===============================================================================

'@label NamedRangeExists
'@fun-title Determine whether a workbook or worksheet name exists.
'@details Determine whether a workbook or worksheet name exists.
'@param nameText String. Name to inspect.
'@param targetBook Optional Workbook. Defaults to ThisWorkbook.
'@return Boolean True when the name is present.
Public Function NamedRangeExists(ByVal nameText As String, _
                                 Optional ByVal targetBook As workbook) As Boolean

    Dim wb As workbook
    Dim nm As Name

    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    On Error Resume Next
        Set nm = wb.Names(nameText)
    On Error GoTo 0

    If Not (nm Is Nothing) Then
        NamedRangeExists = True
        Exit Function
    End If

    Dim sheetName As String
    sheetName = ParseSheetName(nameText)

    If sheetName <> vbNullString And WorksheetExists(sheetName, wb) Then
        On Error Resume Next
            Set nm = wb.Worksheets(sheetName).Names(nameText)
        On Error GoTo 0
        NamedRangeExists = Not (nm Is Nothing)
    End If
End Function

'@label ParseSheetName
'@fun-title Extract worksheet name from qualified references like Sheet1!Name.
'@details Extract worksheet name from qualified references like Sheet1!Name.
'@param qualifiedName String possibly containing '!'.
'@return Worksheet name or empty string.
Private Function ParseSheetName(ByVal qualifiedName As String) As String
    Dim bangPos As Long

    bangPos = InStr(qualifiedName, "!")
    If bangPos > 0 Then
        ParseSheetName = Replace(Left$(qualifiedName, bangPos - 1), "'", vbNullString)
    End If
End Function

'@section Range Writers
'===============================================================================

'@label WriteRow
'@sub-title Write a row of values to a target range.
'@details Write a row of values to a target range.
'@param target Range. Starting cell.
'@param values ParamArray values to write.
Public Sub WriteRow(ByVal target As Range, ParamArray values() As Variant)
    Dim idx As Long

    For idx = LBound(values) To UBound(values)
        target.Offset(0, idx - LBound(values)).value = values(idx)
    Next idx
End Sub

'@label WriteColumn
'@sub-title Write a column of values to a target range.
'@details Write a column of values to a target range.
'@param target Range. Starting cell.
'@param values ParamArray values to write.
Public Sub WriteColumn(ByVal target As Range, ParamArray values() As Variant)
    Dim idx As Long

    For idx = LBound(values) To UBound(values)
        target.Offset(idx - LBound(values), 0).value = values(idx)
    Next idx
End Sub

'@label SingleColumnRows
'@fun-title Convert a 1-D array into an array of single-column rows.
'@details Convert a 1-D array into an array where each element is wrapped in its own single-column row array.
'@param values Variant. One-dimensional array to convert.
'@return Variant array of row arrays or Empty when input is invalid.
Public Function SingleColumnRows(values As Variant) As Variant
    Dim result() As Variant
    Dim idx As Long
    Dim lower As Long
    Dim upper As Long

    If Not IsArray(values) Then Exit Function

    lower = LBound(values)
    upper = UBound(values)
    If upper < lower Then
        SingleColumnRows = Array()
        Exit Function
    End If
    ReDim result(0 To upper - lower)

    For idx = lower To upper
        result(idx - lower) = Array(values(idx))
    Next idx

    SingleColumnRows = result
End Function

'@label RowsToMatrix
'@fun-title Convert an array of row arrays into a 2D matrix.
'@details Convert an array of row arrays into a 2D matrix.
'@param rows Variant. Array of arrays to convert.
'@return Variant 2D matrix or Empty when invalid.
Public Function RowsToMatrix(rows As Variant) As Variant
    Dim rowLower As Long
    Dim rowUpper As Long
    Dim colLower As Long
    Dim colUpper As Long
    Dim r As Long
    Dim c As Long
    Dim matrix() As Variant

    If Not IsArray(rows) Then Exit Function

    rowLower = LBound(rows)
    rowUpper = UBound(rows)
    colLower = LBound(rows(rowLower))
    colUpper = UBound(rows(rowLower))

    ReDim matrix(1 To rowUpper - rowLower + 1, 1 To colUpper - colLower + 1)

    For r = rowLower To rowUpper
        For c = colLower To colUpper
            matrix(r - rowLower + 1, c - colLower + 1) = rows(r)(c)
        Next c
    Next r

    RowsToMatrix = matrix
End Function

'@label WriteMatrix
'@sub-title Write a 2D matrix into the supplied range.
'@details Write a 2D matrix into the supplied range.
'@param target Range. Upper-left cell for the matrix.
'@param matrix Variant. 2D array of values.
Public Sub WriteMatrix(ByVal target As Range, matrix As Variant)
    If IsEmpty(matrix) Then Exit Sub

    target.Resize(UBound(matrix, 1) - LBound(matrix, 1) + 1, _
                  UBound(matrix, 2) - LBound(matrix, 2) + 1).value = matrix
End Sub

'@section Data Builders
'===============================================================================

'@label CollectionToArray
'@fun-title Copy the items from a collection into a zero-based variant array.
'@details Copy the items from a collection into a zero-based variant array, returning an empty array when the input is Nothing or empty.
'@param items Collection containing the values to copy.
'@return Variant array containing the copied values.
Public Function CollectionToArray(ByVal items As Collection) As Variant
    Dim result() As Variant
    Dim idx As Long

    If items Is Nothing Then
        CollectionToArray = Array()
        Exit Function
    End If

    If items.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If

    ReDim result(0 To items.Count - 1)
    For idx = 1 To items.Count
        result(idx - 1) = items(idx)
    Next idx

    CollectionToArray = result
End Function

'@label BetterArrayFromList
'@fun-title Create a BetterArray with the supplied items.
'@details Create a BetterArray with the supplied items.
'@param items ParamArray values to push.
'@return BetterArray containing the items.
Public Function BetterArrayFromList(ParamArray items() As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    Set result = New BetterArray
    result.lowerBound = 0

    For idx = LBound(items) To UBound(items)
        result.Push items(idx)
    Next idx

    Set BetterArrayFromList = result
End Function

'@label BetterArrayFromVariant
'@fun-title Build a BetterArray from a 1-D Variant array.
'@details Build a BetterArray from a 1-D Variant array.
'@param values Variant array.
'@return BetterArray with copied values.
Public Function BetterArrayFromVariant(values As Variant) As BetterArray
    Dim result As BetterArray
    Dim idx As Long

    If Not IsArray(values) Then Exit Function

    Set result = New BetterArray
    result.lowerBound = 0

    For idx = LBound(values) To UBound(values)
        result.Push values(idx)
    Next idx

    Set BetterArrayFromVariant = result
End Function

'@section Assertions
'===============================================================================

'@label FailUnexpectedError
'@sub-title Fail the current test when unexpected errors surface.
'@details Fail the current test when unexpected errors surface.
'@param assertObj Rubberduck Assert object.
'@param routineName String. Name of the failing routine.
Public Sub FailUnexpectedError(assertObj As Object, ByVal routineName As String)
    On Error Resume Next
    assertObj.Fail "Unexpected error in " & routineName & ": " & Err.Number & " - " & Err.description
    On Error GoTo 0
End Sub

'@label CustomTestSetTitles
'@sub-title Configure the pending test title and subtitle for a CustomTest harness.
'@details Safely sets the next test name and subtitle when the harness reference is valid.
'@param harness ICustomTest harness instance.
'@param testName String title to assign.
'@param testSubtitle Optional String subtitle to assign.
Public Sub CustomTestSetTitles(ByVal harness As ICustomTest, _
                              ByVal testName As String, _
                              Optional ByVal testSubtitle As String = vbNullString)
    If harness Is Nothing Then Exit Sub
    harness.SetTestName testName
    harness.SetTestSubtitle testSubtitle
End Sub

'@label CustomTestLogFailure
'@sub-title Log a formatted failure message on a CustomTest harness.
'@details Builds a descriptive failure message containing the routine name and optional error info, then logs it.
'@param harness ICustomTest harness instance.
'@param routineName String name of the failing routine.
'@param errNumber Optional Long error number to include.
'@param errDescription Optional String error description to include.
Public Sub CustomTestLogFailure(ByVal harness As ICustomTest, _
                                ByVal routineName As String, _
                                Optional ByVal errNumber As Long = 0, _
                                Optional ByVal errDescription As String = vbNullString)
    Dim message As String
    Dim errorExplanation As String

    If harness Is Nothing Then Exit Sub
    message = routineName
    
    If errNumber <> 0 Or LenB(errDescription) > 0 Then

        Select Case errNumber
        Case 1001: errorExplanation =  "Invalid argument" 
        Case 1002: errorExplanation =  "Object not initialized" 
        Case 1004: errorExplanation =  "Unexpected state" 
        Case 1005: errorExplanation =  "Element should exists" 
        Case 1006: errorExplanation =  "Element should not exists" 
        Case 1007: errorExplanation =  "Element not found" 
        Case 1008: errorExplanation =  "Something went wrong"
        Case Else: errorExplanation = "Unkown error: (" & errNumber & ")" 
        End Select

        message = message & ": " & errorExplanation & " - " & errDescription
    End If

    harness.LogFailure message
End Sub

'@section VBProject helpers
'===============================================================================

'@label ResolveExportFolder
'@fun-title Determine a writable folder for exported test artifacts.
'@details Prefers the provided workbook path, falling back to ThisWorkbook or the current directory.
'@param referenceWorkbook Optional Workbook used to resolve the path context.
'@return String path guaranteed non-empty.
Public Function ResolveExportFolder(Optional ByVal referenceWorkbook As Workbook, _
                     Optional ByVal folderName As String = vbNullString) As String

    Dim folderPath As String

    If Not referenceWorkbook Is Nothing Then
        folderPath = referenceWorkbook.Path
    Else
        On Error Resume Next
            folderPath = ThisWorkbook.Path
        On Error GoTo 0
    End If

    If LenB(folderPath) = 0 Then folderPath = CurDir$
    If LenB(folderName) <> 0 Then folderPath = folderPath & Application.PathSeparator & folderName
    If Dir$(folderPath, vbDirectory) = vbNullString Then Mkdir folderPath
    ResolveExportFolder = folderPath
End Function

'@label BuildWorkbookPath
'@fun-title Construct a unique workbook path inside the export folder.
'@details Appends timestamp fragments to avoid collisions while keeping the requested prefix.
'@param exportFolder String target folder.
'@param filePrefix String prefix to apply to the generated filename.
'@param extension Optional String file extension including the leading dot. Defaults to .xlsb.
'@return Fully qualified path suitable for saving an Excel workbook.
Public Function BuildWorkbookPath(ByVal exportFolder As String, _
                                  ByVal filePrefix As String, _
                                  Optional ByVal extension As String = ".xlsb") As String

    Dim separatorChar As String
    Dim sanitizedExtension As String

    separatorChar = Application.PathSeparator
    If LenB(extension) = 0 Then
        sanitizedExtension = ".xlsb"
    ElseIf Left$(extension, 1) <> "." Then
        sanitizedExtension = "." & extension
    Else
        sanitizedExtension = extension
    End If

    BuildWorkbookPath = exportFolder & separatorChar & filePrefix & "_" & _
                        Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                        Format$(Timer, "000000") & sanitizedExtension
End Function

'@label ExportComponentToFolder
'@fun-title Export a VBComponent to disk and return its path.
'@details Removes any pre-existing file before exporting to guarantee fresh contents.
'@param sourceWorkbook Workbook hosting the component.
'@param componentName String component code name.
'@param exportFolder String destination folder.
'@return Fully qualified path to the exported component.
Public Function ExportComponentToFolder(ByVal sourceWorkbook As Workbook, _
                                        ByVal componentName As String, _
                                        ByVal exportFolder As String) As String

    Dim vbComp As Object
    Dim exportPath As String
    Dim separatorChar As String

    If sourceWorkbook Is Nothing Then
        Err.Raise vbObjectError + 512, "TestHelpers.ExportComponentToFolder", _
                  "Source workbook is required"
    End If

    separatorChar = Application.PathSeparator

    On Error GoTo MissingComponent
        Set vbComp = sourceWorkbook.VBProject.VBComponents(componentName)
    On Error GoTo 0

    exportPath = exportFolder & separatorChar & componentName & "_" & _
                 Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer, "000000") & _
                 ComponentExtensionName(vbComp.Type)

    On Error Resume Next
        If Dir$(exportPath) <> vbNullString Then Kill exportPath
    On Error GoTo 0

    vbComp.Export exportPath
    ExportComponentToFolder = exportPath
    Exit Function

MissingComponent:
    Err.Raise vbObjectError + 513, "TestHelpers.ExportComponentToFolder", _
              "Component '" & componentName & "' not found"
End Function

'@label CleanupExportedFiles
'@sub-title Delete exported component files captured during a test.
'@details Iterates through supplied collection paths, ignoring errors when files are already removed.
'@param exportedFiles Collection of file paths.
Public Sub CleanupExportedFiles(ByVal exportedFiles As Collection)
    Dim idx As Long

    If exportedFiles Is Nothing Then Exit Sub

    On Error Resume Next
        For idx = 1 To exportedFiles.Count
            If Dir$(CStr(exportedFiles(idx))) <> vbNullString Then
                Kill CStr(exportedFiles(idx))
            End If
        Next idx
    On Error GoTo 0
End Sub

Private Function ComponentExtensionName(ByVal componentType As Long) As String
    Select Case componentType
        Case VBEXT_CT_DOCUMENT, VBEXT_CT_CLASS_MODULE
            ComponentExtensionName = ".cls"
        Case VBEXT_CT_STD_MODULE
            ComponentExtensionName = ".bas"
        Case Else
            ComponentExtensionName = ".cls"
    End Select
End Function
