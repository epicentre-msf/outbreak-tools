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
    nm As String
End Type

Private Type TFormatting
    fontColour As Long
    fillColour As Long
    fontSize As Long
    borderColour As Long
    fontBold As Boolean
End Type

Private this As TState

Private Const HIDDEN_TITLE_COLUMN As Long = 2 'Column B (hidden titles)
Private Const FIRST_OUTPUT_COLUMN As Long = 3 'Column C (visible data)
Private Const OUTPUT_COLUMN_COUNT As Long = 3 'Columns C:E for visible output
Private Const FIRST_ROW_OUTPUT As Long = 4
Private Const FILTER_CELL_ADDRESS As String = "C1"
Private Const FILTER_NAME As String = "CheckingOutputFilter"
Private Const TITLE_FILTER_CELL_ADDRESS As String = "E1"
Private Const TITLE_FILTER_NAME As String = "CheckingOutputTitleFilter"
Private Const TITLE_STORE_ANCHOR As String = "Z2"
Private Const TITLE_RANGE_NAME As String = "CheckingOutputTitles"
Private Const TITLE_FILTER_DEFAULT As String = "All Titles"
Private Const EVENT_MARKER_NAME As String = "CheckingOutputEventInstalled"
Private Const DEFAULT_SEPARATOR As String = "--"
Private Const ROW_MARKER_NAME As String = "CheckingOutputStartingRow"
Private Const END_TAG As String = "__end__checkings___"

'@section Factory
'===============================================================================

'@label:Create
'@sub-title Create an output writer bound to a worksheet
'@details Instantiates a new writer and associates it with the supplied worksheet.
'@param sh Worksheet that will receive the checking output.
'@return ICheckingOutput fully initialised instance.
Public Function Create(ByVal sh As Worksheet, Optional ByVal title As String = vbNullString) As ICheckingOutput
    If sh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
         "Unable to create an instance of the class, please provide a valid worksheet"
    End If

    With New CheckingOutput
        Set .Wksh = sh
         .Name = title
        Set Create = .Self
    End With
End Function

'@section WorksheetBinding
'===============================================================================

'@label:Wksh
'@prop-title Retrieve the worksheet backing this output writer
'@return Worksheet reference.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.targetSheet
End Property

'@prop-title Define the worksheet that will receive the output
'@param sh Worksheet to use.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.targetSheet = sh
End Property

'@prop-title Convenience accessor so callers can chain Create().Self
'@return ICheckingOutput reference to Me.
Public Property Get Self() As ICheckingOutput
    Set Self = Me
End Property

Public Property Get Name() As String
    Name = this.nm
End Property

Public Property Let Name(ByVal nm As String)
    this.nm = nm
End Property

'@section FormattingHelpers
'===============================================================================


'@fun-title Resolve formatting attributes for a given level and colour name
'@details Create formatting base on level
'@param tag String colour or level type token (blue, red, orange, grey, green, "title", "subtitle")
'@return TFormatting structure with font, fill, and border properties
Private Function ResolveFormatting(Optional ByVal tag As String) As TFormatting
    Dim formatInfo As TFormatting


    Select Case LCase$(tag)
        Case "level 1"
            formatInfo.fontBold = True
            formatInfo.fontSize = 20
            formatInfo.fontColour = RGB(0, 0, 128)           ' Navy font
            formatInfo.fillColour = RGB(230, 235, 245)       ' Light navy fill
            formatInfo.borderColour = RGB(0, 0, 64)          ' Dark navy border
        Case "level 2"
            formatInfo.fontBold = True
            formatInfo.fontSize = 14
            formatInfo.fontColour = RGB(0, 0, 128)           ' Navy font
            formatInfo.fillColour = RGB(230, 235, 245)       ' Light navy fill
            formatInfo.borderColour = RGB(0, 0, 64)          ' Dark navy border
        Case "purple"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(112, 48, 160)
            formatInfo.fillColour = RGB(244, 236, 255)
            formatInfo.borderColour = RGB(203, 167, 226)
        Case "red"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(192, 0, 0)
            formatInfo.fillColour = RGB(255, 235, 238)
            formatInfo.borderColour = RGB(224, 102, 102)
        Case "orange"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(167, 106, 0)
            formatInfo.fillColour = RGB(255, 240, 220)
            formatInfo.borderColour = RGB(236, 184, 120)
        Case "grey"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(96, 97, 100)
            formatInfo.fillColour = RGB(242, 243, 245)
            formatInfo.borderColour = RGB(205, 206, 210)
        Case "green"
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(0, 120, 50)
            formatInfo.fillColour = RGB(220, 245, 226)
            formatInfo.borderColour = RGB(120, 200, 140)
        Case "label"
            'Default formatting
            formatInfo.fontSize = 10
            formatInfo.fontColour = RGB(0, 0, 128)           ' Greenish font
            formatInfo.fillColour = vbWhite                  ' Light green fill
            formatInfo.borderColour = RGB(0, 0, 64)   
        Case Else
            'Default formatting
            formatInfo.fontSize = 9
            formatInfo.fontColour = RGB(0, 0, 128)           ' Navy font
            formatInfo.fillColour = RGB(230, 235, 245)       ' Light navy fill
            formatInfo.borderColour = RGB(0, 0, 64) 
    End Select

    ResolveFormatting = formatInfo
End Function


'@sub-title Apply typography styling based on formatting
'@details Decorates text with bold/underline emphasis for primary headings.
'@param target Range receiving the formatting
'@param formatting TFormatting structure containing font and fill attributes
Private Sub ApplyTextDecorations(formatting As TFormatting, ByVal target As Range)

    With target
        .Font.Size = formatting.fontSize
        .Font.Color = formatting.fontColour
        .Interior.Color = formatting.fillColour
        .Font.Bold = formatting.fontBold
    End With
End Sub


'@sub-title Apply border styling to a range
'@details Applies dashed borders around and within the target range block.
'@param targetRange Range receiving the border formatting
'@param borderColour Long colour code for the border
Private Sub ApplyBorders(ByVal targetRange As Range, ByVal borderColour As Long)
    With targetRange.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .Weight = xlThin
        .Color = borderColour
    End With
End Sub


'@sub-title Formats and writes a single row
'@depends Use the supplied text fragments. Text fragments are splitted by --
'@param rowIndex Row where to apply the formatting, Long.
'@param textValue A String, what to write on the row
'@param formatTag, sep Strings; resp. Tag for formatting and split separator.
Private Sub WriteRow(ByVal rowIndex As Long, ByVal textValue As String, _
                     Optional ByVal formatTag As String, _
                     Optional ByVal sep As String = DEFAULT_SEPARATOR, _
                     Optional ByVal parentTitle As String = vbNullString)
    Dim target As Range
    Dim splitValues() As String
    Dim formatting As TFormatting
    Dim sh As Worksheet
    Dim hiddenCell As Range

    Set sh = Wksh()

    Set target = sh.Range(sh.Cells(rowIndex, FIRST_OUTPUT_COLUMN), _
                          sh.Cells(rowIndex, FIRST_OUTPUT_COLUMN + OUTPUT_COLUMN_COUNT - 1))

    Set hiddenCell = sh.Cells(rowIndex, HIDDEN_TITLE_COLUMN)
    hiddenCell.NumberFormat = "@"
    hiddenCell.Font.Color = vbWhite

    If LenB(parentTitle) = 0 Then
        hiddenCell.value = vbNullString
    Else
        hiddenCell.value = parentTitle
    End If

    splitValues = Split(textValue, sep)
    target.Cells(1, 1).value = ValueOrEmpty(splitValues, 0)
    target.Cells(1, 2).value = ValueOrEmpty(splitValues, 1)
    target.Cells(1, 3).value = ValueOrEmpty(splitValues, 2)

    formatting = ResolveFormatting(formatTag)

    ApplyTextDecorations formatting, target
    target.EntireColumn.AutoFit
    ApplyBorders target, formatting.borderColour
End Sub

'@label:ValueOrEmpty
'@fun-title Auto-generated title for ValueOrEmpty.
'@details
'@param items Auto-generated parameter documentation.
'@param index Auto-generated parameter documentation.
'@return Auto-generated return documentation.
Private Function ValueOrEmpty(items() As String, ByVal index As Long) As String
    If index > UBound(items) Then
        ValueOrEmpty = vbNullString
    Else
        ValueOrEmpty = items(index)
    End If
End Function

'@section CheckingWriters
'===============================================================================

'@label WriteChecking
'@sub-title Auto-generated title for WriteChecking.
'@details Auto-generated documentation for WriteChecking.
Private Sub WriteChecking(ByVal checking As IChecking, _
                          ByRef currentRow As Long, _
                          ByRef writtenTitles As Collection, _
                          Optional ByVal sep As String = DEFAULT_SEPARATOR)
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
    subtitle = checking.Heading(subtitle:=True)

    If Not TitleAlreadyWritten(title, writtenTitles) Then
        WriteRow rowIndex:=currentRow, textValue:=title, formatTag:="level 1", sep:=sep, parentTitle:=title
        If LenB(Trim$(title)) > 0 Then AppendTitleToStore title
        writtenTitles.Add Item:=title, key:=title
        currentRow = currentRow + 2
    End If

    If subtitle <> vbNullString Then
        currentRow = currentRow + 1
        WriteRow rowIndex:=currentRow, textValue:=subtitle, formatTag:="level 2", sep:=sep, parentTitle:=title
        currentRow = currentRow + 2
    End If

    Set keys = checking.ListOfKeys
    For keyIndex = keys.lowerBound To keys.upperBound
        keyName = CStr(keys.Item(keyIndex))
        keyColour = checking.ValueOf(keyName, checkingColor)
        keyLabel = checking.ValueOf(keyName, checkingLabel)
        keyType = checking.ValueOf(keyName, checkingType)
        fullLabel = keyType & sep & keyLabel
        WriteRow rowIndex:=currentRow, textValue:=fullLabel, formatTag:=keyColour, sep:=sep, parentTitle:=title
        currentRow = currentRow + 1
    Next keyIndex
End Sub


'@fun-title Whether a title exists
'@param title String, Title to look for
'@param writtenTitles Collection of titles
'@return Boolean
Private Function TitleAlreadyWritten(ByVal title As String, ByRef writtenTitles As Collection) As Boolean
    On Error Resume Next
        TitleAlreadyWritten = Not IsEmpty(writtenTitles.Item(title))
    On Error GoTo 0
End Function

'@section TitleStoreManagement
'===============================================================================

Private Sub EnsureTitleStore(Optional ByVal resetStore As Boolean = False)
    Dim sh As Worksheet

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    If resetStore Then
        ResetTitleStore sh
        Exit Sub
    End If

    EnsureTitleStoreAnchor sh

    If Not WorksheetHasNamedRange(sh, TITLE_RANGE_NAME) Then
        RefreshTitleNamedRange sh
    End If

End Sub

Private Sub ResetTitleStore(ByVal sh As Worksheet)
    Dim anchor As Range
    Dim clearRange As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    Set clearRange = sh.Range(anchor, sh.Cells(sh.Rows.Count, anchor.Column))
    clearRange.ClearContents

    anchor.NumberFormat = "@"
    anchor.value = TITLE_FILTER_DEFAULT
    anchor.Font.Color = vbWhite

    On Error Resume Next
        sh.Names(TITLE_RANGE_NAME).Delete
    On Error GoTo 0

    RefreshTitleNamedRange sh
End Sub

Private Sub EnsureTitleStoreAnchor(ByVal sh As Worksheet)
    Dim anchor As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    anchor.NumberFormat = "@"
    If LenB(Trim$(CStr(anchor.value))) = 0 Then
        anchor.value = TITLE_FILTER_DEFAULT
    End If

    anchor.Font.Color = vbWhite
End Sub

Private Sub AppendTitleToStore(ByVal titleValue As String)
    Dim sh As Worksheet
    Dim anchor As Range
    Dim target As Range
    Dim nextRow As Long

    titleValue = Trim$(titleValue)
    If LenB(titleValue) = 0 Then Exit Sub

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    EnsureTitleStoreAnchor sh

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    If TitleStoreContains(sh, titleValue) Then Exit Sub

    nextRow = NextTitleStoreRow(sh, anchor)
    Set target = sh.Cells(nextRow, anchor.Column)

    target.value = titleValue
    target.NumberFormat = "@"
    target.Font.Color = vbWhite

    RefreshTitleNamedRange sh
End Sub

Private Function TitleStoreContains(ByVal sh As Worksheet, ByVal titleValue As String) As Boolean
    Dim anchor As Range
    Dim lastRow As Long
    Dim titleRange As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    lastRow = sh.Cells(sh.Rows.Count, anchor.Column).End(xlUp).Row
    If lastRow < anchor.Row Then Exit Function

    Set titleRange = sh.Range(anchor, sh.Cells(lastRow, anchor.Column))

    On Error Resume Next
        TitleStoreContains = WorksheetFunction.CountIf(titleRange, titleValue) > 0
    On Error GoTo 0
End Function

Private Sub RefreshTitleNamedRange(ByVal sh As Worksheet)
    Dim anchor As Range
    Dim lastRow As Long
    Dim titleRange As Range

    Set anchor = sh.Range(TITLE_STORE_ANCHOR)

    lastRow = sh.Cells(sh.Rows.Count, anchor.Column).End(xlUp).Row
    If lastRow < anchor.Row Then lastRow = anchor.Row

    Set titleRange = sh.Range(anchor, sh.Cells(lastRow, anchor.Column))

    On Error Resume Next
        sh.Names(TITLE_RANGE_NAME).Delete
    On Error GoTo 0

    sh.Names.Add Name:=TITLE_RANGE_NAME, RefersTo:=titleRange
End Sub

Private Function NextTitleStoreRow(ByVal sh As Worksheet, ByVal anchor As Range) As Long
    Dim lastRow As Long

    lastRow = sh.Cells(sh.Rows.Count, anchor.Column).End(xlUp).Row
    If lastRow < anchor.Row Then
        NextTitleStoreRow = anchor.Row + 1
    Else
        NextTitleStoreRow = lastRow + 1
    End If
End Function

Private Function WorksheetHasNamedRange(ByVal targetSheet As Worksheet, ByVal rangeName As String) As Boolean
    Dim nm As Name

    On Error Resume Next
        Set nm = targetSheet.Names(rangeName)
        WorksheetHasNamedRange = Err.Number = 0 And Not nm Is Nothing
    On Error GoTo 0
End Function

Private Function WorksheetNameForFormula(ByVal targetSheet As Worksheet) As String
    WorksheetNameForFormula = Replace(targetSheet.Name, "'", "''")
End Function

'@sub-title Mask hidden helper columns so metadata remains invisible
'@details Ensures the parent-title column and backing title store columns keep a white font and hidden presentation.
'@param sh Worksheet to format
Private Sub ApplyHiddenColumnsFormatting(ByVal sh As Worksheet)
    Dim storeColumn As Long

    If sh Is Nothing Then Exit Sub

    storeColumn = sh.Range(TITLE_STORE_ANCHOR).Column

    With sh.Columns(HIDDEN_TITLE_COLUMN)
        .NumberFormat = "@"
        .Font.Color = vbWhite
        .Interior.Color = vbWhite
    End With

    With sh.Columns(storeColumn)
        .NumberFormat = "@"
        .Font.Color = vbWhite
        .Interior.Color = vbWhite
        .Hidden = True
    End With
End Sub

'@section WorksheetPreparation
'===============================================================================

'Resets the worksheet before re-rendering the checking output.
'@label InitialiseWorksheet
'@fun-title Auto-generated title for InitialiseWorksheet.
'@details Auto-generated documentation for InitialiseWorksheet.
'@param startingRow Auto-generated parameter documentation.
Private Function InitialiseWorksheet() As Long
    Dim sh As Worksheet
    Dim startingRow As Long
    Dim view As WorksheetView

    Set sh = Wksh()
    sh.Rows.Hidden = False
    
    If Not CustomPropertyExists(ROW_MARKER_NAME) Then
        sh.Cells.Clear
        sh.Cells.Font.Name = "Consolas"
        sh.rows(1).Font.Bold = True
        sh.rows(1).Font.Size = 14
        
        EnsureTitleStore True
        EnsureWorksheetChangeHandler
        ApplyHiddenColumnsFormatting sh
        EnsureFilterDropdown
        startingRow = FIRST_ROW_OUTPUT
        StoreCustomProperty ROW_MARKER_NAME, CStr(FIRST_ROW_OUTPUT)

        'Remove gridlines of the worksheet
        For Each view In sh.Parent.Windows(1).SheetViews
            If view.sheet.Name = sh.Name Then
                view.DisplayGridlines = False
                view.DisplayZeros = False
                Exit For
            End If
        Next view
    Else
        EnsureWorksheetChangeHandler
        startingRow = sh.Cells(sh.Rows.Count, FIRST_OUTPUT_COLUMN).End(xlUp).Row
        If startingRow < FIRST_ROW_OUTPUT Then
            startingRow = FIRST_ROW_OUTPUT
            EnsureTitleStore True
            EnsureFilterDropdown
        Else
            EnsureTitleStore False
            sh.Rows(startingRow).ClearContents
        End If
        ApplyHiddenColumnsFormatting sh
    End If

    InitialiseWorksheet = startingRow
End Function



'@section Filtering
'===============================================================================

'Keeps the filter drop-down present and styled on the worksheet header.
'@sub-title Auto-generated title for EnsureFilterDropdown.
'@label EnsureFilterDropdown
'@details Utilities that manage the drop-down UI and filtering algorithm
'@details Auto-generated documentation for EnsureFilterDropdown.
Private Sub EnsureFilterDropdown()
    Dim sh As Worksheet
    Dim statusCell As Range
    Dim titleCell As Range
    Dim statusLabel As Range
    Dim titleLabel As Range

    Set sh = Wksh()

    EnsureTitleStore False

    Set statusLabel = sh.Cells(1, HIDDEN_TITLE_COLUMN)
    statusLabel.value = "Status:"
    ApplyTextDecorations ResolveFormatting("label"), statusLabel

    Set statusCell = sh.Range(FILTER_CELL_ADDRESS)
    ApplyTextDecorations ResolveFormatting("dropdown"), statusCell

    On Error Resume Next
        sh.Names(FILTER_NAME).Delete
    On Error GoTo 0
    sh.Names.Add Name:=FILTER_NAME, RefersTo:=statusCell

    With statusCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              Formula1:="All,Errors,Warnings,Notes,Infos,Successes,Without Successes"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    If LenB(Trim$(CStr(statusCell.value))) = 0 Then statusCell.value = "All"

    Set titleLabel = sh.Cells(1, sh.Range(TITLE_FILTER_CELL_ADDRESS).Column - 1)
    titleLabel.value = "Title:"
    ApplyTextDecorations ResolveFormatting("label"), titleLabel

    Set titleCell = sh.Range(TITLE_FILTER_CELL_ADDRESS)
    ApplyTextDecorations ResolveFormatting("dropdown"), titleCell
    
    On Error Resume Next
        sh.Names(TITLE_FILTER_NAME).Delete
    On Error GoTo 0

    sh.Names.Add Name:=TITLE_FILTER_NAME, RefersTo:=titleCell

    RefreshTitleNamedRange sh

    With titleCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
              Formula1:="= " & TITLE_RANGE_NAME
        .IgnoreBlank = True
        .InCellDropdown = True
    End With


    If LenB(Trim$(CStr(titleCell.value))) = 0 Then titleCell.value = TITLE_FILTER_DEFAULT
End Sub

'@section WorksheetEvents
'===============================================================================

'Injects the Worksheet_Change handler and helper into the target module when absent.
'@label EnsureWorksheetChangeHandler
'@sub-title Auto-generated title for EnsureWorksheetChangeHandler.
'@details Auto-generated documentation for EnsureWorksheetChangeHandler.
Private Sub EnsureWorksheetChangeHandler()
    Dim codeModule As Object
    Dim moduleLines As Long
    Dim handlerCode As String

    On Error GoTo Handler
    If CustomPropertyExists(EVENT_MARKER_NAME) Then Exit Sub

    Set codeModule = Wksh.Parent.VBProject.VBComponents(Wksh.codeName).codeModule
    moduleLines = codeModule.CountOfLines

    If moduleLines > 0 Then
        Dim moduleText As String
        moduleText = codeModule.Lines(1, moduleLines)
        If ModuleContainsProcedure(moduleText, "FilterCheckingOutputRows") Then
            StoreCustomProperty EVENT_MARKER_NAME, "True"
            Exit Sub
        End If
    End If

    Dim dq As String

    dq = """"

    handlerCode = "Private Function StripIcons(ByVal txt As String) As String" & vbNewLine
    handlerCode = handlerCode & "    Dim stripped As String" & vbNewLine
    handlerCode = handlerCode & "    'Remove severity glyphs so comparisons remain text based" & vbNewLine
    handlerCode = handlerCode & "    stripped = Replace(txt, ChrW(10060), """")" & vbNewLine
    handlerCode = handlerCode & "    stripped = Replace(stripped, ChrW(9888), """")" & vbNewLine
    handlerCode = handlerCode & "    stripped = Replace(stripped, ChrW(8505), """")" & vbNewLine
    handlerCode = handlerCode & "    stripped = Replace(stripped, ChrW(9998), """")" & vbNewLine
    handlerCode = handlerCode & "    stripped = Replace(stripped, ChrW(10004), """")" & vbNewLine
    handlerCode = handlerCode & "    StripIcons = Trim$(stripped)" & vbNewLine
    handlerCode = handlerCode & "End Function" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "Private Function ResolveMatchType(ByVal candidate As String) As String" & vbNewLine
    handlerCode = handlerCode & "    'Normalise the requested status so the filtering block can branch quickly" & vbNewLine
    handlerCode = handlerCode & "    Select Case UCase$(Trim$(candidate))" & vbNewLine
    handlerCode = handlerCode & "        Case """", ""ALL""" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""ALL""" & vbNewLine
    handlerCode = handlerCode & "        Case ""ERROR"", ""ERRORS""" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""ERROR""" & vbNewLine
    handlerCode = handlerCode & "        Case ""WARNING"", ""WARNINGS""" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""WARNING""" & vbNewLine
    handlerCode = handlerCode & "        Case ""NOTE"", ""NOTES""" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""NOTE""" & vbNewLine
    handlerCode = handlerCode & "        Case ""INFO"", ""INFOS""" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""INFO""" & vbNewLine
    handlerCode = handlerCode & "        Case ""SUCCESS"", ""SUCCESSES""" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""SUCCESS""" & vbNewLine
    handlerCode = handlerCode & "        Case ""WITHOUT SUCCESS"", ""WITHOUT SUCCESSES""" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""WITHOUT_SUCCESS""" & vbNewLine
    handlerCode = handlerCode & "        Case Else" & vbNewLine
    handlerCode = handlerCode & "            ResolveMatchType = ""ALL""" & vbNewLine
    handlerCode = handlerCode & "    End Select" & vbNewLine
    handlerCode = handlerCode & "End Function" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "Private Function ShouldFilterByTitle(ByVal candidate As String) As Boolean" & vbNewLine
    handlerCode = handlerCode & "    Dim trimmedTitle As String" & vbNewLine
    handlerCode = handlerCode & "    trimmedTitle = Trim$(candidate)" & vbNewLine
    handlerCode = handlerCode & "    'The default item keeps all titles visible" & vbNewLine
    handlerCode = handlerCode & "    ShouldFilterByTitle = LenB(trimmedTitle) > 0 And StrComp(trimmedTitle, """ & TITLE_FILTER_DEFAULT & """, vbTextCompare) <> 0" & vbNewLine
    handlerCode = handlerCode & "End Function" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine
    handlerCode = handlerCode & "    On Error GoTo ExitHandler" & vbNewLine
    handlerCode = handlerCode & "    Dim statusCell As Range" & vbNewLine
    handlerCode = handlerCode & "    Dim titleCell As Range" & vbNewLine
    handlerCode = handlerCode & "    Dim statusChanged As Boolean" & vbNewLine
    handlerCode = handlerCode & "    Dim titleChanged As Boolean" & vbNewLine
    handlerCode = handlerCode & "    Set statusCell = Me.Range(" & dq & FILTER_CELL_ADDRESS & dq & ")" & vbNewLine
    handlerCode = handlerCode & "    Set titleCell = Me.Range(" & dq & TITLE_FILTER_CELL_ADDRESS & dq & ")" & vbNewLine
    handlerCode = handlerCode & "    statusChanged = Not Intersect(Target, statusCell) Is Nothing" & vbNewLine
    handlerCode = handlerCode & "    titleChanged = Not Intersect(Target, titleCell) Is Nothing" & vbNewLine
    handlerCode = handlerCode & "    'Ignore edits triggered outside of the filter cells" & vbNewLine
    handlerCode = handlerCode & "    If Not statusChanged And Not titleChanged Then GoTo ExitHandler" & vbNewLine
    handlerCode = handlerCode & "    Application.EnableEvents = False" & vbNewLine
    handlerCode = handlerCode & "    Application.ScreenUpdating = False" & vbNewLine
    handlerCode = handlerCode & "    Application.Calculation = xlCalculationManual" & vbNewLine
    handlerCode = handlerCode & "    'Reapply visibility with the refreshed selections" & vbNewLine
    handlerCode = handlerCode & "    FilterCheckingOutputRows statusCell.Value, titleCell.Value, statusChanged, titleChanged" & vbNewLine
    handlerCode = handlerCode & "ExitHandler:" & vbNewLine
    handlerCode = handlerCode & "    Application.EnableEvents = True" & vbNewLine
    handlerCode = handlerCode & "    Application.ScreenUpdating = True" & vbNewLine
    handlerCode = handlerCode & "    Application.Calculation = xlCalculationAutomatic" & vbNewLine
    handlerCode = handlerCode & "End Sub" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "Private Sub FilterCheckingOutputRows(ByVal statusValue As String, _" & vbNewLine
    handlerCode = handlerCode & "                                    ByVal titleValue As String, _" & vbNewLine
    handlerCode = handlerCode & "                                    ByVal statusChanged As Boolean, _" & vbNewLine
    handlerCode = handlerCode & "                                    ByVal titleChanged As Boolean)" & vbNewLine
    handlerCode = handlerCode & "    Dim effectiveStatus As String" & vbNewLine
    handlerCode = handlerCode & "    Dim matchType As String" & vbNewLine
    handlerCode = handlerCode & "    Dim applyTitleFilter As Boolean" & vbNewLine
    handlerCode = handlerCode & "    Dim lastRow As Long" & vbNewLine
    handlerCode = handlerCode & "    Dim rowIndex As Long" & vbNewLine
    handlerCode = handlerCode & "    Dim rowType As String" & vbNewLine
    handlerCode = handlerCode & "    Dim rowTitle As String" & vbNewLine
    handlerCode = handlerCode & "    Dim hasLabel As Boolean" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "    effectiveStatus = statusValue" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "    matchType = ResolveMatchType(effectiveStatus)" & vbNewLine
    handlerCode = handlerCode & "    applyTitleFilter = ShouldFilterByTitle(titleValue)" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "    lastRow = Me.Cells(Me.Rows.Count, " & FIRST_OUTPUT_COLUMN & ").End(xlUp).Row" & vbNewLine
    handlerCode = handlerCode & "    'Exit early when no printable rows are present" & vbNewLine
    handlerCode = handlerCode & "    If lastRow < " & FIRST_ROW_OUTPUT & " Then Exit Sub" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "    For rowIndex = " & FIRST_ROW_OUTPUT & " To lastRow" & vbNewLine
    handlerCode = handlerCode & "        'Start by revealing rows before applying filters" & vbNewLine
    handlerCode = handlerCode & "        Me.Rows(rowIndex).Hidden = False" & vbNewLine
    handlerCode = handlerCode & "        rowTitle = Trim$(CStr(Me.Cells(rowIndex, " & HIDDEN_TITLE_COLUMN & ").Value))" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "        If applyTitleFilter Then" & vbNewLine
    handlerCode = handlerCode & "            'Hide any row whose stored parent title differs from the selection" & vbNewLine
    handlerCode = handlerCode & "            If (StrComp(rowTitle, Trim$(titleValue), vbTextCompare) <> 0) And (rowTitle <> " & dq & END_TAG & dq & ") Then" & vbNewLine
    handlerCode = handlerCode & "                Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handlerCode = handlerCode & "                GoTo ContinueRow" & vbNewLine
    handlerCode = handlerCode & "            End If" & vbNewLine
    handlerCode = handlerCode & "        End If" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "        'Labels differentiate data rows from headers" & vbNewLine
    handlerCode = handlerCode & "        hasLabel = LenB(CStr(Me.Cells(rowIndex, " & (FIRST_OUTPUT_COLUMN + 1) & ").Value)) > 0" & vbNewLine
    handlerCode = handlerCode & vbNewLine
    handlerCode = handlerCode & "        If hasLabel Then" & vbNewLine
    handlerCode = handlerCode & "            rowType = StripIcons(CStr(Me.Cells(rowIndex, " & FIRST_OUTPUT_COLUMN & ").Value))" & vbNewLine
    handlerCode = handlerCode & "            'Apply severity matching only to data rows" & vbNewLine
    handlerCode = handlerCode & "            Select Case matchType" & vbNewLine
    handlerCode = handlerCode & "                Case ""ALL""" & vbNewLine
    handlerCode = handlerCode & "                    'No filtering required for ALL" & vbNewLine
    handlerCode = handlerCode & "                Case ""SUCCESS""" & vbNewLine
    handlerCode = handlerCode & "                    If StrComp(rowType, ""Success"", vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handlerCode = handlerCode & "                Case ""WITHOUT_SUCCESS""" & vbNewLine
    handlerCode = handlerCode & "                    If StrComp(rowType, ""Success"", vbTextCompare) = 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handlerCode = handlerCode & "                Case Else" & vbNewLine
    handlerCode = handlerCode & "                    If StrComp(rowType, matchType, vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handlerCode = handlerCode & "            End Select" & vbNewLine
    handlerCode = handlerCode & "        End If" & vbNewLine
    handlerCode = handlerCode & "ContinueRow:" & vbNewLine
    handlerCode = handlerCode & "    Next rowIndex" & vbNewLine
    handlerCode = handlerCode & "End Sub" & vbNewLine
    handlerCode = handlerCode & vbNewLine

    codeModule.InsertLines moduleLines + 1, handlerCode

    StoreCustomProperty EVENT_MARKER_NAME, "True"
    Exit Sub

Handler:
    ThrowError ProjectError.SomethingWentWrong, "Unable to install worksheet change handler: " & Err.description
End Sub

'@label:FilterWorksheet
'@sub-title Apply filtering logic using the backing worksheet
'@details Allows consumers to reapply the worksheet filtering logic without relying on events.
'@param statusSelection Optional manual filter selection that overrides the worksheet drop-down value.
'@param titleSelection Optional title filter that overrides the worksheet selection when provided.
'@return Nothing
'Reapplies filtering logic using either the supplied values or the worksheet selections.
Public Sub FilterWorksheet(Optional ByVal statusSelection As String = vbNullString, _
                          Optional ByVal titleSelection As String = vbNullString)
    If Wksh() Is Nothing Then Exit Sub
    ApplyWorksheetFilter Wksh, statusSelection, titleSelection
End Sub

'@label:ApplyWorksheetFilter
'@sub-title Internal helper that performs filtering
'@details Centralises the filtering algorithm so both the public API and worksheet events reuse the same logic.
'@param targetSheet Worksheet hosting the checking output data.
'@param statusSelection String representing the desired severity filter; blank values use the worksheet cell.
'@param titleSelection Optional title token overriding the worksheet drop-down when provided.
'@return Nothing
'Executes the row-hiding algorithm based on the provided worksheet and selection.
Private Sub ApplyWorksheetFilter(ByVal targetSheet As Worksheet, _
                                 Optional ByVal statusSelection As String = vbNullString, _
                                 Optional ByVal titleSelection As String = vbNullString)
    Dim statusValue As String
    Dim titleValue As String
    Dim matchMode As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rowType As String
    Dim hasLabel As Boolean
    Dim rowTitle As String
    Dim filterByTitle As Boolean
    Dim explicitStatusOverride As Boolean

    If targetSheet Is Nothing Then Exit Sub

    statusValue = Trim$(statusSelection)
    explicitStatusOverride = LenB(statusValue) > 0
    If LenB(statusValue) = 0 Then
        statusValue = Trim$(CStr(targetSheet.Range(FILTER_CELL_ADDRESS).value))
    End If

    titleValue = Trim$(titleSelection)
    If LenB(titleValue) = 0 Then
        titleValue = Trim$(CStr(targetSheet.Range(TITLE_FILTER_CELL_ADDRESS).value))
    End If

    filterByTitle = LenB(titleValue) > 0 And _
                    StrComp(titleValue, TITLE_FILTER_DEFAULT, vbTextCompare) <> 0

    Select Case UCase$(statusValue)
        Case "", "ALL"
            matchMode = "ALL"
        Case "ERROR", "ERRORS"
            matchMode = "ERROR"
        Case "WARNING", "WARNINGS"
            matchMode = "WARNING"
        Case "NOTE", "NOTES"
            matchMode = "NOTE"
        Case "INFO", "INFOS"
            matchMode = "INFO"
        Case "SUCCESS", "SUCCESSES"
            matchMode = "SUCCESS"
        Case "WITHOUT SUCCESS", "WITHOUT SUCCESSES"
            matchMode = "WITHOUT_SUCCESS"
        Case Else
            matchMode = "ALL"
    End Select

    If filterByTitle And Not explicitStatusOverride Then
        'Grant the title filter precedence when no status override is explicitly supplied
        matchMode = "ALL"
    End If

    lastRow = targetSheet.Cells(targetSheet.rows.Count, FIRST_OUTPUT_COLUMN).End(xlUp).row
    If lastRow < FIRST_ROW_OUTPUT Then Exit Sub

    For rowIndex = FIRST_ROW_OUTPUT To lastRow
        targetSheet.rows(rowIndex).Hidden = False
        rowTitle = Trim$(CStr(targetSheet.Cells(rowIndex, HIDDEN_TITLE_COLUMN).value))

        If filterByTitle Then
            If StrComp(rowTitle, titleValue, vbTextCompare) <> 0 Then
                targetSheet.rows(rowIndex).Hidden = True
                GoTo ContinueLoop
            End If
        End If

        hasLabel = LenB(CStr(targetSheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN + 1).value)) > 0

        If hasLabel Then
            rowType = NormaliseTypeLabel(CStr(targetSheet.Cells(rowIndex, FIRST_OUTPUT_COLUMN).value))
            Select Case matchMode
                Case "ALL"
                    ' no filtering
                Case "SUCCESS"
                    If StrComp(rowType, "Success", vbTextCompare) <> 0 Then targetSheet.rows(rowIndex).Hidden = True
                Case "WITHOUT_SUCCESS"
                    If StrComp(rowType, "Success", vbTextCompare) = 0 Then targetSheet.rows(rowIndex).Hidden = True
                Case Else
                    If StrComp(rowType, matchMode, vbTextCompare) <> 0 Then targetSheet.rows(rowIndex).Hidden = True
            End Select
        End If

ContinueLoop:
    Next rowIndex
End Sub



'@section ICheckingOutput
'===============================================================================
'@label Get
'@prop-title Worksheet of the checking output
'@return A Worksheet Object
Private Property Get ICheckingOutput_Wksh() As Worksheet
    Set ICheckingOutput_Wksh = Wksh()
End Property

Private Property Get ICheckingOutput_Name() As String
    ICheckingOutput_Name = Name()
End Property

'@details Interface members exposed to external callers.

'Coordinates the end-to-end rendering when invoked through the interface.
'@label ICheckingOutput_PrintOutput
'@sub-title Auto-generated title for ICheckingOutput_PrintOutput.
'@param checkTable A BetterArray of checkings to print out
Private Sub ICheckingOutput_PrintOutput(ByVal checkTable As BetterArray)
    Dim currentRow As Long
    Dim writtenTitles As Collection
    Dim index As Long
    Dim previousEventState As Boolean

    If checkTable Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "PrintOutput requires a BetterArray of IChecking instances"
    End If

    previousEventState = Application.EnableEvents
    Application.EnableEvents = False
    ValidateCheckTable checkTable
    
    On Error GoTo CleanExit
    
    currentRow = InitialiseWorksheet()
    Set writtenTitles = New Collection

    For index = checkTable.lowerBound To checkTable.upperBound
        WriteChecking checkTable.Item(index), currentRow, writtenTitles
    Next index

    '@Ignore AssignmentNotUsed
    currentRow = currentRow + 4
    WriteRow currentRow, formatTag:="level 4", textValue:="End of checkings " & Name(), parentTitle:=END_TAG
    StoreCustomProperty ROW_MARKER_NAME, CStr(currentRow)

    EnsureWorksheetChangeHandler

CleanExit:
    Application.EnableEvents = previousEventState
    If Err.Number <> 0 Then ThrowError Err.Number, Err.description
End Sub


'@jump:FilterWorksheetInterface
'@sub-title Interface shim delegating to FilterWorksheet
'@details Bridges the interface call to the public filtering routine so callers can filter without casting.
'Routes filter requests from the interface to the public method.
'@param selectionValue Optional String token matching the desired filter.
Private Sub ICheckingOutput_FilterWorksheet(Optional ByVal selectionValue As String = vbNullString, _
                                            Optional ByVal titleSelection As String = vbNullString)
    FilterWorksheet selectionValue, titleSelection
End Sub

'Ensure worksheet handler
Private Sub ICheckingOutput_EnsureWorksheetChangeHandler()
    EnsureWorksheetChangeHandler
End Sub

'@section Class private helpers
'===============================================================================

'@sub-title The BetterArray contains only IChecking entries
'@param checkTable BetterArray table of Checking
Private Sub ValidateCheckTable(ByVal checkTable As BetterArray)
    Dim index As Long
    For index = checkTable.lowerBound To checkTable.upperBound
        If Not TypeOf checkTable.Item(index) Is IChecking Then
           ThrowError ProjectError.InvalidArgument, _
            "Items in PrintOutput must implement IChecking. Item " & CStr(index) & " does not."
        End If
    Next index
End Sub


'@label:ThrowError
'@sub-title Raise a ProjectError-coded runtime error
'@param errNumb ProjectError enumeration value.
'@param errorMessage String human readable description.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "CheckingOutput", errorMessage
End Sub


'@fun-title Clean output checking label
'@param typeLabel A String, label to clean on a worksheet
'@return A String, cleaned label
Private Function NormaliseTypeLabel(ByVal typeLabel As String) As String
    Dim cleaned As String

    cleaned = typeLabel
    cleaned = Replace(cleaned, ChrW(10060), vbNullString)
    cleaned = Replace(cleaned, ChrW(9888), vbNullString)
    cleaned = Replace(cleaned, ChrW(8505), vbNullString)
    cleaned = Replace(cleaned, ChrW(9998), vbNullString)
    cleaned = Replace(cleaned, ChrW(10004), vbNullString)

    NormaliseTypeLabel = Trim$(cleaned)
End Function

'@fun-title Check whether a custom property exists
'@return A Boolean
Private Function CustomPropertyExists(ByVal propName As String) As Boolean
    Dim prop As CustomProperty
    Dim sh As Worksheet
    Set sh = Wksh()

    For Each prop In sh.CustomProperties
        If StrComp(prop.Name, propName, vbTextCompare) = 0 Then
            CustomPropertyExists = True
            Exit Function
        End If
    Next
End Function

'@sub-title Stores a custom property.
Private Sub StoreCustomProperty(propName As String, propValue As String)
    Dim prop As CustomProperty
    Dim sh As Worksheet
    Set sh = Wksh()
    For Each prop In sh.CustomProperties
        If prop.Name = propName Then
            prop.Value = propValue
            Exit Sub
        End If
    Next
    sh.CustomProperties.Add propName, propValue
End Sub

'@label:ModuleContainsProcedure
'@sub-title Determine whether the module text already declares the target procedure
'@param moduleText VBA module contents
'@param procedureName Name of the procedure to search for
'@return Boolean indicating whether the procedure exists
Private Function ModuleContainsProcedure(ByVal moduleText As String, _
                                         ByVal procedureName As String) As Boolean
    Dim subSignature As String
    Dim functionSignature As String

    If LenB(moduleText) = 0 Then Exit Function

    subSignature = "Sub " & procedureName & "("
    If InStr(1, moduleText, subSignature, vbTextCompare) > 0 Then
        ModuleContainsProcedure = True
        Exit Function
    End If

    functionSignature = "Function " & procedureName & "("
    ModuleContainsProcedure = (InStr(1, moduleText, functionSignature, vbTextCompare) > 0)
End Function
