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
Private Const VBIDE_PROC_KIND As Long = 0 'Matches vbext_pk_Proc to avoid VBIDE dependency

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
    
    currentRow = currentRow + 2
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
        Else
            EnsureTitleStore False
            sh.Rows(startingRow).ClearContents
        End If
        ApplyHiddenColumnsFormatting sh
        EnsureFilterDropdown

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
    Dim injectionCode As String
    Dim moduleLines As Long

    On Error GoTo Handler

    If Wksh() Is Nothing Then Exit Sub

    Set codeModule = Wksh.Parent.VBProject.VBComponents(Wksh.CodeName).CodeModule

    RemoveManagedWorksheetProcedures codeModule

    injectionCode = BuildWorksheetChangeHandlerCode()
    moduleLines = codeModule.CountOfLines

    If moduleLines > 0 Then
        codeModule.InsertLines moduleLines + 1, vbNewLine & injectionCode
    Else
        codeModule.InsertLines 1, injectionCode
    End If

    StoreCustomProperty EVENT_MARKER_NAME, "True"
    Exit Sub

Handler:
    ThrowError ProjectError.SomethingWentWrong, _
        "Unable to install worksheet change handler: " & Err.Description
End Sub

'@label BuildWorksheetChangeHandlerCode
'@sub-title Compose the worksheet filtering logic
'@details Generates the VBA code injected in the worksheet module so it can filter rows without depending on this class.
Private Function BuildWorksheetChangeHandlerCode() As String
    Dim handler As String
    Dim dq As String

    dq = """"

    handler = handler & "Private Function StripIcons(ByVal txt As String) As String" & vbNewLine
    handler = handler & "    Dim stripped As String" & vbNewLine
    handler = handler & "    stripped = Replace(txt, ChrW(10060), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(9888), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(8505), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(9998), """")" & vbNewLine
    handler = handler & "    stripped = Replace(stripped, ChrW(10004), """")" & vbNewLine
    handler = handler & "    StripIcons = Trim$(stripped)" & vbNewLine
    handler = handler & "End Function" & vbNewLine & vbNewLine

    handler = handler & "Private Function ResolveMatchType(ByVal candidate As String) As String" & vbNewLine
    handler = handler & "    Select Case UCase$(Trim$(candidate))" & vbNewLine
    handler = handler & "        Case """", ""ALL""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""ALL""" & vbNewLine
    handler = handler & "        Case ""ERROR"", ""ERRORS""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""ERROR""" & vbNewLine
    handler = handler & "        Case ""WARNING"", ""WARNINGS""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""WARNING""" & vbNewLine
    handler = handler & "        Case ""NOTE"", ""NOTES""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""NOTE""" & vbNewLine
    handler = handler & "        Case ""INFO"", ""INFOS""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""INFO""" & vbNewLine
    handler = handler & "        Case ""SUCCESS"", ""SUCCESSES""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""SUCCESS""" & vbNewLine
    handler = handler & "        Case ""WITHOUT SUCCESS"", ""WITHOUT SUCCESSES""" & vbNewLine
    handler = handler & "            ResolveMatchType = ""WITHOUT_SUCCESS""" & vbNewLine
    handler = handler & "        Case Else" & vbNewLine
    handler = handler & "            ResolveMatchType = ""ALL""" & vbNewLine
    handler = handler & "    End Select" & vbNewLine
    handler = handler & "End Function" & vbNewLine & vbNewLine

    handler = handler & "Private Function ShouldFilterByTitle(ByVal candidate As String) As Boolean" & vbNewLine
    handler = handler & "    Dim trimmedTitle As String" & vbNewLine
    handler = handler & "    trimmedTitle = Trim$(candidate)" & vbNewLine
    handler = handler & "    ShouldFilterByTitle = LenB(trimmedTitle) > 0 And StrComp(trimmedTitle, " & dq & TITLE_FILTER_DEFAULT & dq & ", vbTextCompare) <> 0" & vbNewLine
    handler = handler & "End Function" & vbNewLine & vbNewLine

    handler = handler & "Private Sub ResetUiState(ByVal previousEvents As Boolean, ByVal previousUpdating As Boolean, ByVal previousCalculation As XlCalculation)" & vbNewLine
    handler = handler & "    Application.EnableEvents = previousEvents" & vbNewLine
    handler = handler & "    Application.ScreenUpdating = previousUpdating" & vbNewLine
    handler = handler & "    Application.Calculation = previousCalculation" & vbNewLine
    handler = handler & "End Sub" & vbNewLine & vbNewLine

    handler = handler & "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine
    handler = handler & "    On Error GoTo Handler" & vbNewLine
    handler = handler & "    Dim statusCell As Range" & vbNewLine
    handler = handler & "    Dim titleCell As Range" & vbNewLine
    handler = handler & "    Dim statusChanged As Boolean" & vbNewLine
    handler = handler & "    Dim titleChanged As Boolean" & vbNewLine
    handler = handler & "    Dim previousEvents As Boolean" & vbNewLine
    handler = handler & "    Dim previousUpdating As Boolean" & vbNewLine
    handler = handler & "    Dim previousCalculation As XlCalculation" & vbNewLine
    handler = handler & "    Dim managedState As Boolean" & vbNewLine
    handler = handler & "    Dim errorNumber As Long" & vbNewLine
    handler = handler & "    Dim errorDescription As String" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    If Target Is Nothing Then Exit Sub" & vbNewLine
    handler = handler & "    If Target.CountLarge > 1 Then Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    Set statusCell = Me.Range(" & dq & FILTER_CELL_ADDRESS & dq & ")" & vbNewLine
    handler = handler & "    Set titleCell = Me.Range(" & dq & TITLE_FILTER_CELL_ADDRESS & dq & ")" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    statusChanged = Not Intersect(Target, statusCell) Is Nothing" & vbNewLine
    handler = handler & "    titleChanged = Not Intersect(Target, titleCell) Is Nothing" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    If Not statusChanged And Not titleChanged Then Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    previousEvents = Application.EnableEvents" & vbNewLine
    handler = handler & "    previousUpdating = Application.ScreenUpdating" & vbNewLine
    handler = handler & "    previousCalculation = Application.Calculation" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    Application.EnableEvents = False" & vbNewLine
    handler = handler & "    Application.ScreenUpdating = False" & vbNewLine
    handler = handler & "    Application.Calculation = xlCalculationManual" & vbNewLine
    handler = handler & "    managedState = True" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    FilterCheckingOutputRows CStr(statusCell.Value), CStr(titleCell.Value)" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "ExitHandler:" & vbNewLine
    handler = handler & "    If managedState Then" & vbNewLine
    handler = handler & "        ResetUiState previousEvents, previousUpdating, previousCalculation" & vbNewLine
    handler = handler & "    End If" & vbNewLine
    handler = handler & "    If errorNumber <> 0 Then" & vbNewLine
    handler = handler & "        Err.Raise errorNumber, " & dq & "CheckingOutputWorksheet" & dq & ", errorDescription" & vbNewLine
    handler = handler & "    End If" & vbNewLine
    handler = handler & "    Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "Handler:" & vbNewLine
    handler = handler & "    errorNumber = Err.Number" & vbNewLine
    handler = handler & "    errorDescription = Err.Description" & vbNewLine
    handler = handler & "    Resume ExitHandler" & vbNewLine
    handler = handler & "End Sub" & vbNewLine & vbNewLine

    handler = handler & "Private Sub FilterCheckingOutputRows(ByVal statusValue As String, ByVal titleValue As String)" & vbNewLine
    handler = handler & "    Dim matchType As String" & vbNewLine
    handler = handler & "    Dim applyTitleFilter As Boolean" & vbNewLine
    handler = handler & "    Dim lastRow As Long" & vbNewLine
    handler = handler & "    Dim rowIndex As Long" & vbNewLine
    handler = handler & "    Dim rowType As String" & vbNewLine
    handler = handler & "    Dim rowTitle As String" & vbNewLine
    handler = handler & "    Dim hasLabel As Boolean" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    matchType = ResolveMatchType(statusValue)" & vbNewLine
    handler = handler & "    applyTitleFilter = ShouldFilterByTitle(titleValue)" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    lastRow = Me.Cells(Me.Rows.Count, " & FIRST_OUTPUT_COLUMN & ").End(xlUp).Row" & vbNewLine
    handler = handler & "    If lastRow < " & FIRST_ROW_OUTPUT & " Then Exit Sub" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "    For rowIndex = " & FIRST_ROW_OUTPUT & " To lastRow" & vbNewLine
    handler = handler & "        Me.Rows(rowIndex).Hidden = False" & vbNewLine
    handler = handler & "        rowTitle = Trim$(CStr(Me.Cells(rowIndex, " & HIDDEN_TITLE_COLUMN & ").Value))" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "        If applyTitleFilter Then" & vbNewLine
    handler = handler & "            If (StrComp(rowTitle, Trim$(titleValue), vbTextCompare) <> 0) And (rowTitle <> " & dq & END_TAG & dq & ") Then" & vbNewLine
    handler = handler & "                Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "                GoTo ContinueRow" & vbNewLine
    handler = handler & "            End If" & vbNewLine
    handler = handler & "        End If" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "        hasLabel = LenB(CStr(Me.Cells(rowIndex, " & (FIRST_OUTPUT_COLUMN + 1) & ").Value)) > 0" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "        If hasLabel Then" & vbNewLine
    handler = handler & "            rowType = StripIcons(CStr(Me.Cells(rowIndex, " & FIRST_OUTPUT_COLUMN & ").Value))" & vbNewLine
    handler = handler & "            Select Case matchType" & vbNewLine
    handler = handler & "                Case ""ALL""" & vbNewLine
    handler = handler & "                    'No filtering required" & vbNewLine
    handler = handler & "                Case ""SUCCESS""" & vbNewLine
    handler = handler & "                    If StrComp(rowType, ""Success"", vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "                Case ""WITHOUT_SUCCESS""" & vbNewLine
    handler = handler & "                    If StrComp(rowType, ""Success"", vbTextCompare) = 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "                Case Else" & vbNewLine
    handler = handler & "                    If StrComp(rowType, matchType, vbTextCompare) <> 0 Then Me.Rows(rowIndex).Hidden = True" & vbNewLine
    handler = handler & "            End Select" & vbNewLine
    handler = handler & "        End If" & vbNewLine
    handler = handler & vbNewLine
    handler = handler & "ContinueRow:" & vbNewLine
    handler = handler & "    Next rowIndex" & vbNewLine
    handler = handler & "End Sub" & vbNewLine

    BuildWorksheetChangeHandlerCode = handler
End Function
'@section WorksheetEventHelpers
'===============================================================================

'@label RemoveManagedWorksheetProcedures
'@sub-title Strip previously injected worksheet routines
'@details Removes any instance of the managed procedures so a single fresh handler can be written.
Private Sub RemoveManagedWorksheetProcedures(ByVal codeModule As Object)
    DeleteProcedureIfExists codeModule, "Worksheet_Change"
    DeleteProcedureIfExists codeModule, "StripIcons"
    DeleteProcedureIfExists codeModule, "ResolveMatchType"
    DeleteProcedureIfExists codeModule, "ShouldFilterByTitle"
    DeleteProcedureIfExists codeModule, "ResetUiState"
    DeleteProcedureIfExists codeModule, "FilterCheckingOutputRows"
End Sub

'@label DeleteProcedureIfExists
'@sub-title Safely remove a procedure from a worksheet module
'@details Deletes the target procedure when present. Ignores missing procedures silently.
'@param codeModule Worksheet code module reference.
'@param procedureName Name of the procedure to delete.
Private Sub DeleteProcedureIfExists(ByVal codeModule As Object, ByVal procedureName As String)
    Dim startLine As Long
    Dim lineCount As Long

    On Error Resume Next
        startLine = codeModule.ProcStartLine(procedureName, VBIDE_PROC_KIND)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        lineCount = codeModule.ProcCountLines(procedureName, VBIDE_PROC_KIND)
        codeModule.DeleteLines startLine, lineCount
    On Error GoTo 0
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
