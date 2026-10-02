VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds disease worksheets with translated headers, validations, and formatting"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Builds disease worksheets with translated headers, validations, and formatting")
'@depends MasterSetupVariables, CustomTable, DropdownLists, HiddenNames, BetterArray
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'Builds a disease worksheet of the master setup: the table with translated
'headers, its validations, hidden names and formatting.

'@section Constants
'===============================================================================

Private Const NAME_SHEET_TAG As String = "sheetTag"
Private Const NAME_DISNAME As String = "__Var_DISNAME"
Private Const NAME_DISLANG As String = "__Var_DISLANG"
Private Const NAME_INDEX As String = "__Var_DISINDEX"
Private Const NAME_DISCODE As String = "__Var_DISCODE"


Private Const DEFAULT_ANCHOR_SHEET As String = "Variables"
Private Const DEFAULT_TABLE_PREFIX As String = "disTab_"
Private Const DEFAULT_DISEASE_CODE As String = "DISSHEET"


Private Const STATUS_DROPDOWN As String = "__var_status"
Private Const DISEASES_DROPDOWN_PROHIBITED As String = "__prohibited_diseases_list"
Private Const DISEASES_DROPDOWN As String = "__diseases_list"
Private Const CHOICES_DROPDOWN As String = "__lst_choices"
Private Const LANGUAGES_DROPDOWN As String = "__data_languages"

Private Const ANCHORROW As Long = 2
Private Const ANCHORCOLUMN As Long = 2
'A new disease table opens with this many lines, so the user picks before
'reaching for Add Rows.
Private Const DEFAULT_ROW_COUNT As Long = 15

'@section State
'===============================================================================

'The seven headers of a disease table, in the order the table carries them.
Private Type TCurrentHeaders
    varOrder As String
    varSection As String
    varName As String
    varLabel As String
    varChoice As String
    choiceVal As String
    varStatus As String
End Type

Private Type TBuilderState
    hostwb As Workbook
    dropdowns As DropdownLists
    initialised As Boolean
    hostSheet As Worksheet
    hostIndex As Long
    hostTable As CustomTable
    hostVariables As MasterSetupVariables
    hostLanguage As String
    defaultColumns As BetterArray       'The headers as a list, built once
    headers As TCurrentHeaders          'The same headers by name
End Type

Private this As TBuilderState

'@section Factory
'===============================================================================

'@sub-title Instantiate a sheet builder with provided dependencies.
Public Function Create(ByVal hostWorkbook As Workbook, _
                       ByVal dropdowns As DropdownLists, _
                       ByVal varObj As MasterSetupVariables) As DiseaseSheet

    Dim instance As DiseaseSheet
    Set instance = New DiseaseSheet
    instance.Initialise hostWorkbook, dropdowns, varObj
    Set Create = instance
End Function

'@section Configuration
'===============================================================================

Public Sub Initialise(ByVal hostWorkbook As Workbook, _
                      ByVal dropdowns As DropdownLists, _
                      ByVal varObj As MasterSetupVariables)

    ValidateInitialise hostWorkbook, dropdowns, varObj

    Set this.hostwb = hostWorkbook
    Set this.dropdowns = dropdowns
    Set this.hostVariables = varObj
    this.hostLanguage = vbNullString
    this.initialised = True
End Sub

'@section Worksheet Construction
'===============================================================================

Public Function Build(ByVal diseaseName As String, Optional ByVal languageTag As String = vbNullString) As Worksheet

    diseaseName = Trim$(diseaseName)
    ValidateBuildInputs diseaseName, languageTag
    this.hostIndex = ResolveNextDiseaseIndex()

    'create the disease worksheet
    CreateDiseaseWorksheet diseaseName



    'Add the table, validations and metadata
    WriteData

    
    
    'Format the disease worksheet, add lockings
    FormatDiseaseSheet

    Set Build = this.hostSheet
End Function

'@section Maintenance
'===============================================================================

'@sub-title Frame every cell of a disease table with the dotted gray border.
'@details The door for the two paths that grow a disease table after its
'build: the ribbon's Add Rows and the importer's append. A row added
'through ListRows.Add takes the vertical edges of the row above and no
'horizontal edge, and a table grown through ListObject.Resize takes no
'format at all, so the frame is painted again over the whole table. The
'call is stateless and runs on the predeclared instance.
'@param diseaseTable ListObject. The disease table to frame.
Public Sub FrameTable(ByVal diseaseTable As ListObject)
    If diseaseTable Is Nothing Then Exit Sub
    PaintTableBorders diseaseTable.Range
End Sub

Public Sub Reset()
    Set this.hostwb = Nothing
    Set this.dropdowns = Nothing
    this.initialised = False
    Set this.hostSheet = Nothing
    Set this.hostTable = Nothing
    Set this.hostVariables = Nothing
    this.hostLanguage = vbNullString
    Set this.defaultColumns = Nothing
End Sub

'@section Internal Helpers
'===============================================================================
Private Sub EnsureColumnsInitialized()

    If this.defaultColumns Is Nothing Then

        Set this.defaultColumns = New BetterArray
        
        this.defaultColumns.LowerBound = 1

        this.headers.varOrder    = "Variable Order"
        this.headers.varName     = "Variable Name"
        this.headers.varSection  = "Variable Section"
        this.headers.varLabel    = "Main Label"
        this.headers.varChoice   = "Choice"
        this.headers.choiceVal   = "Choice Values"
        this.headers.varStatus   = "Status"


        this.defaultColumns.Push  this.headers.varOrder, _
                                    this.headers.varName, _
                                    this.headers.varSection, _
                                    this.headers.varLabel, _
                                    this.headers.varChoice, _
                                    this.headers.choiceVal, _
                                    this.headers.varStatus
    End If
End Sub


Private Sub ValidateInitialise(ByVal hostWorkbook As Workbook, _
                               ByVal dropdowns As DropdownLists, _
                               ByVal varObj As MasterSetupVariables)

    If hostWorkbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook reference cannot be Nothing"
    End If

    If dropdowns Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Dropdown lists cannot be Nothing"
    End If

    If varObj Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Variables manager cannot be Nothing"
    End If
End Sub

Private Sub ValidateBuildInputs(ByVal diseaseName As String, Optional ByVal languageTag As String = vbNullString)
    Dim langDrop As BetterArray
    Dim temp As BetterArray
    Dim defaultLanguage As Variant

    If Not this.initialised Then
        ThrowError ProjectError.ObjectNotInitialized, "Builder has not been initialised"
    End If

    If LenB(diseaseName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Disease name cannot be empty"
    End If

    If Len(diseaseName) > 30 Then
        ThrowError ProjectError.InvalidArgument, "Disease name should be less than 30 characters long (including spaces)"
    End If

    Set langDrop = this.dropdowns.Values(LANGUAGES_DROPDOWN)
    If langDrop Is Nothing Or langDrop.Length = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                  "Language dropdown '" & LANGUAGES_DROPDOWN & "' has no available entries."
    End If

    If LenB(languageTag) > 0 Then
        If Not langDrop.Includes(languageTag) Then
            ThrowError ProjectError.InvalidArgument, _
                      "Language '" & languageTag & "' is not present in the list of languages."
        End If
        this.hostLanguage = languageTag
    Else
        defaultLanguage = langDrop.Item(langDrop.LowerBound)
        If LenB(CStr(defaultLanguage)) = 0 Then
            ThrowError ProjectError.InvalidArgument, _
                      "Language dropdown '" & LANGUAGES_DROPDOWN & "' does not contain a default value."
        End If
        this.hostLanguage = CStr(defaultLanguage)
    End If

    Set temp = this.dropdowns.Values(DISEASES_DROPDOWN_PROHIBITED)
    If Not temp Is Nothing And temp.Includes(diseaseName) Then
        ThrowError ProjectError.InvalidArgument, _
                  "The name " & diseaseName & " is incorrect because it is reserved. Please use another name."
    End If

    Set temp = this.dropdowns.Values(DISEASES_DROPDOWN)
    If Not temp Is Nothing And temp.Includes(diseaseName) Then
        ThrowError ProjectError.InvalidArgument, _
                  "The name " & diseaseName & " already exists, please choose another name."
    End If
End Sub


Private Sub CreateDiseaseWorksheet(ByVal diseaseName As String)
    Dim anchor As Worksheet
    Dim hostwb As Workbook
    Dim hostsheet As Worksheet
 
    Set hostwb = this.hostwb

    On Error Resume Next
    Set anchor = hostwb.Worksheets(DEFAULT_ANCHOR_SHEET)
    On Error GoTo 0

    If (anchor Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, _
         "Anchor worksheet " & DEFAULT_ANCHOR_SHEET & " not found"
    End If

    Set hostsheet = hostwb.Worksheets.Add(Before:=anchor)
    hostsheet.Name = diseaseName
    UpdateInternalElements hostsheet
End Sub

Private Sub UpdateInternalElements(ByVal hostsheet As Worksheet)
    
    Dim temp As BetterArray
    Set temp = this.dropdowns.Values(DISEASES_DROPDOWN)

    If temp Is Nothing Then
        Set temp = New BetterArray
        temp.LowerBound = 1
    End If

    temp.Push hostsheet.Name

    this.dropdowns.Update temp, DISEASES_DROPDOWN
    Set this.hostSheet = hostsheet
End Sub


Private Sub WriteValues()

    Dim anchorRng As Range

    If this.hostSheet Is Nothing Then Exit Sub
    
    Set anchorRng = this.hostSheet.Cells(ANCHORROW, ANCHORCOLUMN)

    'The language cell is the one visible control of the header line. Every
    'tag of the sheet lives in its hidden names (owner rule, 2026-08-25):
    'the code and the marker cells of the old master setup are gone.
    anchorRng.Value = this.hostLanguage

    EnsureColumnsInitialized
    Set anchorRng = anchorRng.Offset(2)
    this.defaultColumns.ToExcelRange anchorRng, TransposeValues:=True

    'The two formula columns are written once the table exists, over its
    'whole body: a new table opens with DEFAULT_ROW_COUNT lines, and a
    'formula put on the first line alone leaves the others blank.
End Sub

'@sub-title Put the label and the choice values formulas back on every line of a disease table.
'@details The build writes the two columns here once the table exists, an
'import writes values over them because a merge copies whole lines, and a
'table grown through Resize takes no formula on its new rows. The master
'setup answers those two columns itself, through the worksheet functions,
'so the formulas go on every line each time. The table has to be one this
'class built: the columns are found by the headers it writes.
'@param diseaseTable ListObject. The table of a disease worksheet.
Public Sub WriteLineFormulas(ByVal diseaseTable As ListObject)
    Dim body As Range
    Dim sheet As Worksheet
    Dim langAddress As String
    Dim mainLabIndex As Long
    Dim choiValIndex As Long
    Dim varNameIndex As Long
    Dim varChoiceIndex As Long

    If diseaseTable Is Nothing Then Exit Sub
    Set body = diseaseTable.DataBodyRange
    If body Is Nothing Then Exit Sub

    Set sheet = diseaseTable.Parent
    langAddress = sheet.Cells(ANCHORROW, ANCHORCOLUMN).Address(True, True)

    mainLabIndex = ColumnIndexFor("Main Label")
    choiValIndex = ColumnIndexFor("Choice Values")
    varNameIndex = ColumnIndexFor("Variable Name")
    varChoiceIndex = ColumnIndexFor("Choice")
    If varNameIndex <= 0 Then Exit Sub

    'A formula written on a whole column keeps its row references relative,
    'so every line reads its own name and choice cells.
    If mainLabIndex > 0 Then
        body.Columns(mainLabIndex).Formula = LabelFormula(body.Cells(1, varNameIndex), langAddress)
    End If

    If choiValIndex > 0 And varChoiceIndex > 0 Then
        body.Columns(choiValIndex).Formula = ChoiceValuesFormula(body.Cells(1, varChoiceIndex), langAddress)
    End If

    'Under manual calculation the imported values would stand until
    'something else recalculated the sheet.
    body.Calculate
End Sub

'@sub-title The label formula of a line: the worksheet function over the name cell.
Private Function LabelFormula(ByVal nameCell As Range, ByVal langAddress As String) As String
    LabelFormula = "=MainLabelValue(" & nameCell.Address(RowAbsolute:=False) & ", " & langAddress & ")"
End Function

'@sub-title The choice values formula of a line: the worksheet function over the choice cell.
Private Function ChoiceValuesFormula(ByVal choiceCell As Range, ByVal langAddress As String) As String
    ChoiceValuesFormula = "=ChoiceValues(" & choiceCell.Address(RowAbsolute:=False) & ", " & langAddress & ")"
End Function

Private Sub CreateDiseaseTable()

    Dim anchorRng As Range 
    Dim tableRange As Range
    Dim dataEnd As Range
    Dim table As ListObject

    If this.hostSheet Is Nothing Then Exit Sub
    
    EnsureColumnsInitialized
    Set anchorRng = this.hostSheet.Cells(ANCHORROW, ANCHORCOLUMN).Offset(2)
    Set dataEnd = anchorRng.Offset(DEFAULT_ROW_COUNT, this.defaultColumns.Length - 1)
    Set tableRange = this.hostSheet.Range(anchorRng, dataEnd)
    Set table = this.hostSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    table.Name = DEFAULT_TABLE_PREFIX & Format$(this.hostIndex, "000")
    table.TableStyle = ""
    Set this.hostTable = CustomTable.Create(table)

End Sub


'Add validations on the disease worksheet
Private Sub AddValidations()
    
    Dim anchorRng As Range

    If (this.hostTable Is Nothing) Then Exit Sub
    If (this.hostVariables Is Nothing) Then Exit Sub

    EnsureColumnsInitialized

    'set validation
    this.hostVariables.SetVariableValidation this.hostTable.DataRange(this.headers.varName)
    this.hostTable.SetValidation colName:=this.headers.varStatus, drop:=this.dropdowns, dropName:=STATUS_DROPDOWN
    this.hostTable.SetValidation colName:=this.headers.varChoice, drop:=this.dropdowns, dropName:=CHOICES_DROPDOWN
    
    Set anchorRng = this.hostSheet.Cells(ANCHORROW, ANCHORCOLUMN)

    this.dropdowns.SetValidation anchorRng, LANGUAGES_DROPDOWN
End Sub


Private Sub WriteData()

    WriteValues
    CreateDiseaseTable
    If Not this.hostTable Is Nothing Then WriteLineFormulas this.hostTable.Table
    AddValidations
    WriteMetadata

End Sub

Private Sub WriteMetadata()

    Dim storeSheet As HiddenNames
    Dim storeBook As HiddenNames

    If this.hostSheet Is Nothing Then Exit Sub

    Set storeSheet = HiddenNames.Create(this.hostSheet)
    Set storeBook = HiddenNames.Create(this.hostSheet.Parent)

    'Write the sheet tag
    storeSheet.EnsureName NAME_SHEET_TAG, "disease", HiddenNameTypeString

    'Write the disease index
    storeSheet.EnsureName NAME_INDEX, this.hostIndex , HiddenNameTypeLong


    'Write the language of the disease
    storeSheet.EnsureName NAME_DISLANG, this.hostLanguage, HiddenNameTypeString

    'Add the disease name to the worksheet
    storeSheet.EnsureName NAME_DISNAME, this.hostSheet.Name, HiddenNameTypeString

    'Write the disease code beside the other tags. Exports read it from here;
    'no cell carries it.
    storeSheet.EnsureName NAME_DISCODE, DiseaseCodeName(), HiddenNameTypeString

    'Add the disease code name to the workbook
    storeBook.EnsureName DiseaseCodeName(), this.hostSheet.Name, HiddenNameTypeString
End Sub


'The sheet is white. The language cell and the table headers carry the house
'blue with white text, so the two things the user reads first stand out. The
'table is framed in dotted gray, because the gridlines are off and the body
'would otherwise read as nothing past the header row.
Private Sub FormatDiseaseSheet()
    Dim languageCell As Range

    RemoveGridlines this.hostSheet

    With this.hostSheet
        .Cells.Font.Name = "Calibri"
        .Cells.Interior.Color = vbWhite
    End With

    Set languageCell = this.hostSheet.Cells(ANCHORROW, ANCHORCOLUMN)
    With languageCell
        .Interior.Color = HouseBlue()
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With

    'The frame goes on before the header fill, so the header keeps its blue.
    FrameTable this.hostTable.Table

    With this.hostTable.HeaderRange
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = HouseBlue()
        .Cells.HorizontalAlignment = xlHAlignLeft
        .Cells.VerticalAlignment = xlVAlignCenter
    End With

    'The sheet ships protected, and every cell of a worksheet starts locked.
    'The language cell and the table body are unlocked so the user can pick
    'in them; the two formula columns lock back, because the label and the
    'choice values are answered by the worksheet functions and never typed.
    languageCell.Locked = False
    this.hostTable.Table.DataBodyRange.Locked = False
    this.hostTable.DataRange(this.headers.choiceVal).Locked = True
    this.hostTable.DataRange(this.headers.varLabel).Locked = True

End Sub

'RGB is a call, so the colour cannot be a Const.
Private Function HouseBlue() As Long
    HouseBlue = RGB(0, 82, 155)
End Function

Private Function BorderGray() As Long
    BorderGray = RGB(166, 166, 166)
End Function

'Frame every cell of the table with a thin dotted gray border. Rows added
'later through ManageRows inherit the format of the last row, so the frame
'grows with the table.
'
'The four edges are painted cell by cell. Painting the edges of the whole
'table range leaves the bottom edge of the last table row unreadable: the
'cell shows the line, and Borders(xlEdgeBottom) still answers xlNone, so
'the frame could not be proved. Per cell, every edge reads back.
Private Sub PaintTableBorders(ByVal target As Range)
    Dim cell As Range

    For Each cell In target.Cells
        PaintDottedBorder cell, xlEdgeTop
        PaintDottedBorder cell, xlEdgeBottom
        PaintDottedBorder cell, xlEdgeLeft
        PaintDottedBorder cell, xlEdgeRight
    Next cell
End Sub

'Paint one edge of a range as a thin dotted gray line.
Private Sub PaintDottedBorder(ByVal target As Range, ByVal edge As Long)
    With target.Borders(edge)
        .LineStyle = xlDot
        .Weight = xlThin
        .Color = BorderGray()
    End With
End Sub

Private Sub RemoveGridlines(ByVal sheet As Worksheet)
    Dim view As WorksheetView
    For Each view In sheet.Parent.Windows(1).SheetViews
        If view.Sheet.Name = sheet.Name Then
            view.DisplayGridlines = False
            Exit For
        End If
    Next view
End Sub

Private Function ColumnIndexFor(ByVal header As String) As Long
    EnsureColumnsInitialized
    ColumnIndexFor = this.defaultColumns.IndexOf(header)
End Function

Private Function FindWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet

    If wb Is Nothing Then Exit Function

    For Each sh In wb.Worksheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sh
            Exit Function
        End If
    Next sh
End Function

Private Sub EnsureNameDeleted(ByVal targetBook As Workbook, ByVal nameValue As String)
    Dim namedRange As Name

    For Each namedRange In targetBook.Names
        If StrComp(namedRange.Name, nameValue, vbTextCompare) = 0 Then
            namedRange.Delete
            Exit For
        End If
    Next namedRange
End Sub

Private Sub AssignName(ByVal targetBook As Workbook, ByVal target As Range, ByVal nameValue As String)
    EnsureNameDeleted targetBook, nameValue
    targetBook.Names.Add Name:=nameValue, RefersTo:=target
End Sub
Private Function ResolveNextDiseaseIndex() As Long
    Dim highestIndex As Long
    Dim fallbackCount As Long
    Dim sh As Worksheet
    Dim sheetIndex As Long

    If this.hostwb Is Nothing Then Exit Function

    For Each sh In this.hostwb.Worksheets
        If WorksheetIsDisease(sh) Then
            sheetIndex = DiseaseSheetIndex(sh)
            If sheetIndex > 0 Then
                If sheetIndex > highestIndex Then highestIndex = sheetIndex
            Else
                fallbackCount = fallbackCount + 1
            End If
        End If
    Next sh

    If highestIndex > 0 Then
        ResolveNextDiseaseIndex = highestIndex + 1
    Else
        ResolveNextDiseaseIndex = fallbackCount + 1
    End If
End Function

Private Function WorksheetIsDisease(ByVal targetSheet As Worksheet) As Boolean
    Dim store As HiddenNames
    Dim tagValue As String

    If targetSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set store = HiddenNames.Create(targetSheet)
    On Error GoTo 0

    If Not store Is Nothing Then
        tagValue = Trim$(store.ValueAsString(NAME_SHEET_TAG))
        If LenB(tagValue) > 0 Then
            WorksheetIsDisease = (StrComp(tagValue, "disease", vbTextCompare) = 0)
            If WorksheetIsDisease Then Exit Function
        End If
    End If
End Function

Private Function DiseaseSheetIndex(ByVal targetSheet As Worksheet) As Long
    Dim store As HiddenNames

    If targetSheet Is Nothing Then Exit Function

    On Error Resume Next
        Set store = HiddenNames.Create(targetSheet)
    On Error GoTo 0

    If store Is Nothing Then Exit Function

    DiseaseSheetIndex = store.ValueAsLong(NAME_INDEX, 0)
End Function

Private Function DiseaseCodeName() As String
    DiseaseCodeName = DEFAULT_DISEASE_CODE & Format$(this.hostIndex, "000")
End Function

'@section ErrorHandling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseSheet", errMessage
End Sub
