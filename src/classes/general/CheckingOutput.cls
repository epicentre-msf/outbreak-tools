VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckingOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Writes checking results to a worksheet"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Writes checking results to a worksheet")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation


'@section Overview
'@details Provides facilities to print checking results and filter them directly within the worksheet context.

Option Explicit
Implements ICheckingOutput

'@section State
'===============================================================================

Private Type TState
    targetSheet As Worksheet
End Type

Private Type TFormatting
    fontColour As Long
    fillColour As Long
    fontSize As Long
    borderColour As Long
End Type

Private this As TState

Private Const CLASSNAME As String = "CheckingOutput"
Private Const FIRST_OUTPUT_COLUMN As Long = 2 'Column B
Private Const OUTPUT_COLUMN_COUNT As Long = 3 'B:C:D
Private Const FILTER_CELL_ADDRESS As String = "C1"
Private Const FILTER_NAME As String = "CheckingOutputFilter"
Private Const EVENT_MARKER_NAME As String = "CheckingOutputEventInstalled"

'@section Factory
'===============================================================================

'@label:Create
'@sub-title Create an output writer bound to a worksheet
'@details Instantiates a new writer and associates it with the supplied worksheet.
'@param sh Worksheet that will receive the checking output.
'@return ICheckingOutput fully initialised instance.
Public Function Create(ByVal sh As Worksheet) As ICheckingOutput
    If sh Is Nothing Then
        RaiseError "Unable to create an instance of the class, please provide a valid worksheet"
    End If

    With New CheckingOutput
        Set .Worksheet = sh
        Set Create = .Self
    End With
End Function

'@section WorksheetBinding
'===============================================================================

'@label:Worksheet
'@prop-title Retrieve the worksheet backing this output writer
'@return Worksheet reference.
Public Property Get Worksheet() As Worksheet
    Set Worksheet = this.targetSheet
End Property

'@prop-title Define the worksheet that will receive the output
'@param sh Worksheet to use.
Public Property Set Worksheet(ByVal sh As Worksheet)
    Set this.targetSheet = sh
End Property

'@prop-title Convenience accessor so callers can chain Create().Self
'@return ICheckingOutput reference to Me.
Public Property Get Self() As ICheckingOutput
    Set Self = Me
End Property

'@section FormattingHelpers
'===============================================================================


'@sub-title Resolve formatting attributes for a given level and colour name
'@param level Byte styling level (1 = title, 2 = subtitle, 3 = detail)
'@param colourName String colour token (blue, red, orange, grey, default)
'@return TFormatting structure with font, fill, and border properties
Private Function ResolveFormatting(ByVal level As Byte, ByVal colourName As String) As TFormatting
    Dim formatInfo As TFormatting

    Select Case level
        Case 1
            formatInfo.fontSize = 14
            formatInfo.fontColour = RGB(41, 55, 73)
            formatInfo.fillColour = RGB(241, 244, 248)
            formatInfo.borderColour = RGB(191, 198, 207)
            ResolveFormatting = formatInfo
            Exit Function
        Case 2
            formatInfo.fontSize = 12
            formatInfo.fontColour = RGB(48, 84, 150)
            formatInfo.fillColour = RGB(235, 241, 255)
            formatInfo.borderColour = RGB(173, 196, 230)
            ResolveFormatting = formatInfo
            Exit Function
        Case Else
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(66, 66, 66)
            formatInfo.fillColour = RGB(248, 248, 248)
            formatInfo.borderColour = RGB(213, 214, 216)
    End Select

    Select Case LCase$(colourName)
        Case "purple"
            formatInfo.fontColour = RGB(112, 48, 160)
            formatInfo.fillColour = RGB(244, 236, 255)
            formatInfo.borderColour = RGB(203, 167, 226)
        Case "red"
            formatInfo.fontColour = RGB(192, 0, 0)
            formatInfo.fillColour = RGB(255, 235, 238)
            formatInfo.borderColour = RGB(224, 102, 102)
        Case "orange"
            formatInfo.fontColour = RGB(167, 106, 0)
            formatInfo.fillColour = RGB(255, 240, 220)
            formatInfo.borderColour = RGB(236, 184, 120)
        Case "grey"
            formatInfo.fontColour = RGB(96, 97, 100)
            formatInfo.fillColour = RGB(242, 243, 245)
            formatInfo.borderColour = RGB(205, 206, 210)
        Case Else
            ' keep default neutral palette
    End Select

    ResolveFormatting = formatInfo
End Function

'@sub-title Apply typography styling based on level and resolved formatting
'@param target Range receiving the formatting
'@param level Byte styling level (1 = title, 2 = subtitle, 3 = detail)
'@param formatting TFormatting structure containing font and fill attributes
'Decorates text with bold/underline emphasis for primary headings.
Private Sub ApplyTextDecorations(ByVal target As Range, _
                                 ByVal level As Byte, _
                                 ByVal formatting As TFormatting)

    With target
        .Font.Size = formatting.fontSize
        .Font.Color = formatting.fontColour
        .Interior.Color = formatting.fillColour

        Select Case level
            Case 1
                .Font.Bold = True
                .Font.Name = "Tahoma"
            Case 2
                .Font.Italic = True
                .Font.Name = "Tahoma"
            Case Else
                .Font.Name = "Consolas"
        End Select
    End With
End Sub

'@sub-title Apply border styling to the supplied range
'@param targetRange Range receiving the border formatting
'@param borderColour Long colour code for the border
'Applies dashed borders around and within the target range block.
Private Sub ApplyBorders(ByVal targetRange As Range, ByVal borderColour As Long)
    Dim borderIndex As Long

    For borderIndex = xlEdgeLeft To xlInsideHorizontal
        With targetRange.Borders(borderIndex)
            .LineStyle = xlDashDotDot
            .Weight = xlHairline
            .Color = borderColour
        End With
    Next borderIndex
End Sub

'Formats and writes a single output row using the supplied text fragments.
Private Sub WriteRow(ByVal rowIndex As Long, _
                     ByVal level As Byte, _
                     ByVal textValue As String, _
                     Optional ByVal colourName As String = "black")

    Dim target As Range
    Dim splitValues() As String
    Dim formatting As TFormatting

    Set target = Worksheet.Range(Worksheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN), _
                                 Worksheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN + OUTPUT_COLUMN_COUNT - 1))

    splitValues = Split(textValue, "--")
    target.Cells(1, 1).Value = ValueOrEmpty(splitValues, 0)
    target.Cells(1, 2).Value = ValueOrEmpty(splitValues, 1)
    target.Cells(1, 3).Value = ValueOrEmpty(splitValues, 2)

    formatting = ResolveFormatting(level, colourName)

    ApplyTextDecorations target, level, formatting
    target.EntireColumn.AutoFit

    ApplyBorders target, formatting.borderColour
End Sub

Private Function ValueOrEmpty(ByVal items() As String, ByVal index As Long) As String
    If index > UBound(items) Then
        ValueOrEmpty = vbNullString
    Else
        ValueOrEmpty = items(index)
    End If
End Function

'@section CheckingWriters
'===============================================================================

Private Sub WriteChecking(ByVal checking As IChecking, _
                          ByRef currentRow As Long, _
                          ByRef writtenTitles As Collection)
'Outputs headings and detail entries for an individual checking bundle.

    Dim title As String
    Dim subtitle As String
    Dim keys As BetterArray
    Dim keyIndex As Long
    Dim keyName As String
    Dim keyColour As String
    Dim keyLabel As String
    Dim keyType As String
    Dim fullLabel As String

    title = checking.Heading
    subtitle = checking.Heading(subTitle:=True)

    If Not TitleAlreadyWritten(title, writtenTitles) Then
        If currentRow > 2 Then currentRow = currentRow + 1
        WriteRow currentRow, 1, title
        writtenTitles.Add Item:=title, Key:=title
        currentRow = currentRow + 1
    End If

    If subtitle <> vbNullString Then
        WriteRow currentRow, 2, subtitle
        currentRow = currentRow + 1
    End If

    Set keys = checking.ListOfKeys
    For keyIndex = keys.LowerBound To keys.UpperBound
        keyName = CStr(keys.Item(keyIndex))
        keyColour = checking.ValueOf(keyName, checkingColor)
        keyLabel = checking.ValueOf(keyName, checkingLabel)
        keyType = checking.ValueOf(keyName, checkingType)
        fullLabel = keyType & "--" & keyLabel
        WriteRow currentRow, 3, fullLabel, keyColour
        currentRow = currentRow + 1
    Next keyIndex
End Sub

'Detects whether a section title has already been persisted in the worksheet.
Private Function TitleAlreadyWritten(ByVal title As String, ByVal writtenTitles As Collection) As Boolean
    On Error Resume Next
        TitleAlreadyWritten = Not IsEmpty(writtenTitles.Item(title))
    On Error GoTo 0
End Function

'@section WorksheetPreparation
'===============================================================================

'Resets the worksheet before re-rendering the checking output.
Private Sub InitialiseWorksheet(ByRef startingRow As Long)
    Worksheet.Cells.Clear
    Worksheet.Cells(1, 2).Value = "Show only:"
    EnsureFilterDropdown
    EnsureWorksheetChangeHandler
    startingRow = 2
End Sub

'Validates that the supplied BetterArray contains only IChecking entries.
Private Sub ValidateCheckTable(ByVal checkTable As BetterArray)
    Dim index As Long

    For index = checkTable.LowerBound To checkTable.UpperBound
        If Not TypeOf checkTable.Item(index) Is IChecking Then
            RaiseError "Items in PrintOutput must implement IChecking. Item " & CStr(index) & " does not."
        End If
    Next index
End Sub

'@section ErrorHandling
'===============================================================================

'Raises a class-specific error with a friendly message.
Private Sub RaiseError(ByVal message As String)
    Err.Raise vbObjectError + 516, CLASSNAME, message
End Sub

'@section Filtering
'===============================================================================
'@details Utilities that manage the drop-down UI and filtering algorithm.

'Keeps the filter drop-down present and styled on the worksheet header.
Private Sub EnsureFilterDropdown()
    Dim dropdownCell As Range

    Set dropdownCell = Worksheet.Range(FILTER_CELL_ADDRESS)

    On Error Resume Next
        Worksheet.Names(FILTER_NAME).Delete
    On Error GoTo 0

    Worksheet.Names.Add Name:=FILTER_NAME, RefersTo:=dropdownCell

    With dropdownCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              Formula1:="All,Errors,Warnings,Notes,Infos"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With dropdownCell
        .Font.Color = RGB(34, 66, 106)
        .Font.Bold = True
        .Interior.Color = RGB(230, 242, 255)
    End With

    If LenB(CStr(dropdownCell.Value)) = 0 Then dropdownCell.Value = "All"
End Sub

'@section WorksheetEvents
'===============================================================================
'@details Helpers that install and track the worksheet-level event code.

'Checks whether the worksheet already carries the event hook marker.
Private Function EventHandlerInstalled() As Boolean
    Dim prop As CustomProperty

    For Each prop In Worksheet.CustomProperties
        If StrComp(prop.Name, EVENT_MARKER_NAME, vbTextCompare) = 0 Then
            EventHandlerInstalled = True
            Exit Function
        End If
    Next prop
End Function

'Stores a custom property so we do not reinstall the worksheet event twice.
Private Sub MarkEventHandlerInstalled()
    On Error Resume Next
        Worksheet.CustomProperties(EVENT_MARKER_NAME).Delete
    On Error GoTo 0
    Worksheet.CustomProperties.Add EVENT_MARKER_NAME, "True"
End Sub

'Injects the Worksheet_Change handler and helper into the target module when absent.
Private Sub EnsureWorksheetChangeHandler()
    Dim codeModule As Object
    Dim moduleLines As Long
    Dim handlerCode As String

    On Error GoTo Handler
    If EventHandlerInstalled() Then Exit Sub

    Set codeModule = Worksheet.Parent.VBProject.VBComponents(Worksheet.CodeName).CodeModule
    moduleLines = codeModule.CountOfLines

    If moduleLines > 0 Then
        If codeModule.Find("Private Sub FilterCheckingOutputRows", 1, 1, moduleLines, 255) Then
            MarkEventHandlerInstalled
            Exit Sub
        End If
    End If

    handlerCode = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
    handlerCode = handlerCode & "    On Error GoTo ExitHandler" & vbCrLf
    handlerCode = handlerCode & "    If Intersect(Target, Me.Range(""" & FILTER_CELL_ADDRESS & """)) Is Nothing Then Exit Sub" & vbCrLf
    handlerCode = handlerCode & "    Application.EnableEvents = False" & vbCrLf
    handlerCode = handlerCode & "    FilterCheckingOutputRows" & vbCrLf
    handlerCode = handlerCode & "ExitHandler:" & vbCrLf
    handlerCode = handlerCode & "    Application.EnableEvents = True" & vbCrLf
    handlerCode = handlerCode & "End Sub" & vbCrLf
    handlerCode = handlerCode & vbCrLf
    handlerCode = handlerCode & "Private Sub FilterCheckingOutputRows()" & vbCrLf
    handlerCode = handlerCode & "    Dim selectionValue As String" & vbCrLf
    handlerCode = handlerCode & "    Dim matchType As String" & vbCrLf
    handlerCode = handlerCode & "    Dim lastRow As Long" & vbCrLf
    handlerCode = handlerCode & "    Dim rowIndex As Long" & vbCrLf
    handlerCode = handlerCode & "    Dim rowType As String" & vbCrLf
    handlerCode = handlerCode & "    Dim hasLabel As Boolean" & vbCrLf
    handlerCode = handlerCode & vbCrLf
    handlerCode = handlerCode & "    selectionValue = Trim$(CStr(Me.Range(""" & FILTER_CELL_ADDRESS & """).Value))" & vbCrLf
    handlerCode = handlerCode & "    Select Case UCase$(selectionValue)" & vbCrLf
    handlerCode = handlerCode & "        Case  ""ALL"" " & vbCrLf
    handlerCode = handlerCode & "            matchType = vbNullString" & vbCrLf
    handlerCode = handlerCode & "        Case ""ERROR"", ""ERRORS""" & vbCrLf
    handlerCode = handlerCode & "            matchType = ""Error""" & vbCrLf
    handlerCode = handlerCode & "        Case ""WARNING"", ""WARNINGS""" & vbCrLf
    handlerCode = handlerCode & "            matchType = ""Warning""" & vbCrLf
    handlerCode = handlerCode & "        Case ""NOTE"", ""NOTES""" & vbCrLf
    handlerCode = handlerCode & "            matchType = ""Note""" & vbCrLf
    handlerCode = handlerCode & "        Case ""INFO"", ""INFOS""" & vbCrLf
    handlerCode = handlerCode & "            matchType = ""Info""" & vbCrLf
    handlerCode = handlerCode & "        Case Else" & vbCrLf
    handlerCode = handlerCode & "            matchType = vbNullString" & vbCrLf
    handlerCode = handlerCode & "    End Select" & vbCrLf
    handlerCode = handlerCode & vbCrLf
    handlerCode = handlerCode & "    lastRow = Me.Cells(Me.Rows.Count, " & CStr(FIRST_OUTPUT_COLUMN) & ").End(xlUp).Row" & vbCrLf
    handlerCode = handlerCode & "    If lastRow < 2 Then Exit Sub" & vbCrLf
    handlerCode = handlerCode & vbCrLf
    handlerCode = handlerCode & "    For rowIndex = 2 To lastRow" & vbCrLf
    handlerCode = handlerCode & "        Me.Rows(rowIndex).Hidden = False" & vbCrLf
    handlerCode = handlerCode & "        hasLabel = LenB(CStr(Me.Cells(rowIndex, " & CStr(FIRST_OUTPUT_COLUMN + 1) & ").Value)) > 0" & vbCrLf
    handlerCode = handlerCode & vbCrLf
    handlerCode = handlerCode & "        If hasLabel And (matchType <> vbNullString) Then" & vbCrLf
    handlerCode = handlerCode & "            rowType = CStr(Me.Cells(rowIndex, " & CStr(FIRST_OUTPUT_COLUMN) & ").Value)" & vbCrLf
    handlerCode = handlerCode & "            If StrComp(rowType, matchType, vbTextCompare) <> 0 Then" & vbCrLf
    handlerCode = handlerCode & "                Me.Rows(rowIndex).Hidden = True" & vbCrLf
    handlerCode = handlerCode & "            End If" & vbCrLf
    handlerCode = handlerCode & "        End If" & vbCrLf
    handlerCode = handlerCode & "    Next rowIndex" & vbCrLf
    handlerCode = handlerCode & "End Sub" & vbCrLf

    codeModule.InsertLines moduleLines + 1, handlerCode

    MarkEventHandlerInstalled
    Exit Sub

Handler:
    RaiseError "Unable to install worksheet change handler: " & Err.Description
End Sub

'@label:FilterWorksheet
'@sub-title Apply filtering logic using the backing worksheet
'@details Allows consumers to reapply the worksheet filtering logic without relying on events.
'@param selectionValue Optional manual filter selection that overrides the worksheet drop-down value.
'@return Nothing
'Reapplies filtering logic using either the supplied value or the worksheet selection.
Public Sub FilterWorksheet(Optional ByVal selectionValue As String = vbNullString)
    If Worksheet Is Nothing Then Exit Sub
    ApplyWorksheetFilter Worksheet, selectionValue
End Sub

'@label:ApplyWorksheetFilter
'@sub-title Internal helper that performs filtering
'@details Centralises the filtering algorithm so both the public API and worksheet events reuse the same logic.
'@param targetSheet Worksheet hosting the checking output data.
'@param selectionValue String representing the desired filter; blank values use the worksheet cell.
'@return Nothing
'Executes the row-hiding algorithm based on the provided worksheet and selection.
Private Sub ApplyWorksheetFilter(ByVal targetSheet As Worksheet, ByVal selectionValue As String)
    Dim selectionValueLocal As String
    Dim matchType As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rowType As String
    Dim hasLabel As Boolean

    If targetSheet Is Nothing Then Exit Sub

    selectionValueLocal = Trim$(selectionValue)
    If LenB(selectionValueLocal) = 0 Then
        selectionValueLocal = Trim$(CStr(targetSheet.Range(FILTER_CELL_ADDRESS).Value))
    End If

    Select Case UCase$(selectionValueLocal)
        Case "", "ALL"
            matchType = vbNullString
        Case "ERROR", "ERRORS"
            matchType = "Error"
        Case "WARNING", "WARNINGS"
            matchType = "Warning"
        Case "NOTE", "NOTES"
            matchType = "Note"
        Case "INFO", "INFOS"
            matchType = "Info"
        Case Else
            matchType = vbNullString
    End Select

    lastRow = targetSheet.Cells(targetSheet.Rows.Count, FIRST_OUTPUT_COLUMN).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For rowIndex = 2 To lastRow
        targetSheet.Rows(rowIndex).Hidden = False
        hasLabel = LenB(CStr(targetSheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN + 1).Value)) > 0

        If hasLabel And (matchType <> vbNullString) Then
            rowType = CStr(targetSheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN).Value)
            If StrComp(rowType, matchType, vbTextCompare) <> 0 Then
                targetSheet.Rows(rowIndex).Hidden = True
            End If
        End If
    Next rowIndex
End Sub

'@section ICheckingOutput
'===============================================================================
'@details Interface members exposed to external callers.

'Coordinates the end-to-end rendering when invoked through the interface.
Private Sub ICheckingOutput_PrintOutput(ByVal checkTable As BetterArray)
    Dim currentRow As Long
    Dim writtenTitles As Collection
    Dim index As Long
    Dim previousEventState As Boolean

    If checkTable Is Nothing Then
        RaiseError "PrintOutput requires a BetterArray of IChecking instances"
    End If

    previousEventState = Application.EnableEvents
    On Error GoTo CleanExit
    Application.EnableEvents = False

    ValidateCheckTable checkTable
    InitialiseWorksheet currentRow
    Set writtenTitles = New Collection

    For index = checkTable.LowerBound To checkTable.UpperBound
        WriteChecking checkTable.Item(index), currentRow, writtenTitles
    Next index

CleanExit:
    Application.EnableEvents = previousEventState
    If Err.Number <> 0 Then RaiseError Err.Description
End Sub


'@jump:FilterWorksheetInterface
'@sub-title Interface shim delegating to FilterWorksheet
'@details Bridges the interface call to the public filtering routine so callers can filter without casting.
'@param selectionValue Optional String token matching the desired filter.
'@return Nothing
'Routes filter requests from the interface to the public method.
Private Sub ICheckingOutput_FilterWorksheet(Optional ByVal selectionValue As String = vbNullString)
    FilterWorksheet selectionValue
End Sub
