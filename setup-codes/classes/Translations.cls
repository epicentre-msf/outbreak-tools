VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Translations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage translations of the current setup"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Manage translations of the current setup")
'@IgnoreModule ImplicitActiveSheetReference : I will use Range method without referring to a worksheet.

Option Explicit
Implements ITranslations

Private Type TTrans
  sh As Worksheet
  loname As String
End Type

Private this As TTrans
Private Const CLASSNAME As String = "Translations"

Public Function Create(ByVal sh As Worksheet, loname As String) As ITranslations
  If sh Is Nothing Then ThrowError "Please provide a Worksheet for translation"
  If loname = vbNullString Then ThrowError "Please provide a listobject name for translations"
  With New Translations
    Set .Wksh = sh
    .TableName = loname
    Set Create = .Self()
  End With
End Function

Private Property Get Table() As ListObject
    Dim sh As Worksheet
    Dim loname As String
    Set sh = Wksh()
    loname = TableName()
    Set Table = sh.ListObjects(loname)
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Get TableName() As String
    TableName = this.loname
End Property

Public Property Let TableName(ByVal txt As String)
    this.loname = txt
End Property

Public Property Get Self() As ITranslations
    Set Self = Me
End Property

'where to start writing elements in the listObject
Public Property Get StartRange() As Range
  Dim cellRng As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Set cellRng = Lo.Range.Cells(1, 1)
  Do While cellRng.Value <> vbNullString
    Set cellRng = cellRng.Offset(1)
  Loop
  Set StartRange = cellRng
End Property

'anchor range which is the column range of the table
Private Property Get KeyLangRange() As Range
  Dim rng As Range
  Dim Lo As ListObject

  Set Lo = Table()
  If (Lo.DataBodyRange Is Nothing) Then
    Set rng = Lo.Range.Cells(2, 1)
  Else
    Set rng = Lo.ListColumns(1).DataBodyRange
  End If
  Set KeyLangRange = rng
End Property

Private Property Get NumberTimes() As Long
  Dim nbTimes As Long
  Dim sh As Worksheet

  Set sh = Wksh()
  nbTimes = 0
  On Error Resume Next
  nbTimes = CLng(sh.Cells(1, 1).Value)
  On Error GoTo 0
  NumberTimes = nbTimes
End Property

Private Sub AddNumberOfTimes()
  Dim sh As Worksheet
  Dim nbTimes As Long
  Set sh = Wksh()
  nbTimes = NumberTimes()
  sh.Cells(1, 1).Value = nbTimes + 1
End Sub

'Add a chunk (a word or a phrase in one cell to the translation table)
Private Sub AddChunk(ByVal lab As String, ByVal rngTag As String)
  Dim insertRng As Range 'where to insert the label
  'Test if the element to include Exists

  'range Tag will have two elements
  'updated to know if the range have been updated and
  'number of times to know the number of times in translations
  If chunkExists(lab) Then
    Set insertRng = chunkRange(lab)
  Else
    Set insertRng = StartRange()
    insertRng.Cells(1, 1).Value = lab
  End If
  insertRng.Cells(1, 1).Offset(, -1).Value = rngTag
End Sub

'Add a formula chunk to the table
Private Sub AddFormulaChunk(ByVal lab As String, ByVal rngTag As String)

  Dim cleanedFormula As String
  Dim innerChunkStart As Long
  Dim counter As Long

  'remove all empty strings in the formula
  cleanedFormula = Replace(lab, Chr(34) & Chr(34), vbNullString)

  'If you enconter a character string, extract the text within
  If InStr(1, cleanedFormula, Chr(34), 1) > 0 Then
      For counter = 1 To Len(cleanedFormula)
          If Mid(cleanedFormula, counter, 1) = Chr(34) Then
              If innerChunkStart = 0 Then
                  innerChunkStart = counter + 1
              Else
                  AddChunk Mid(cleanedFormula, innerChunkStart, counter - innerChunkStart), rngTag
                  innerChunkStart = 0
              End If
          End If
      Next
  End If
End Sub

'Delete unwanted Rows
Private Sub DeleteRows(ByVal rngSh As Worksheet)
  Dim Lo As ListObject
  Dim cellRng As Range
  Dim colRng As Range
  Dim actualNbTimes As Long
  Dim rngTag As String
  Dim delRow As Boolean
  Dim nbTag As Long
  Dim savedRng As Range
  Dim rngName As String

  Set Lo = Table()
  Set colRng = Lo.ListColumns(1).Range
  Set cellRng = colRng.Cells(colRng.Rows.Count, 1)
  Set savedRng = colRng.Cells(colRng.Rows.Count + 1, 1)

  actualNbTimes = NumberTimes()

  Do While cellRng.Row > colRng.Row
    rngTag = cellRng.Offset(, -1).Value
    If rngTag = vbNullString Then
      delRow = True
    Else
      nbTag = 0
      On Error Resume Next
        rngName = Split(rngTag, "--")(0)
        nbTag = CLng(Split(rngTag, "--")(1))
      On Error GoTo 0
      delRow = (nbTag < actualNbTimes) And RngHasUpdate(rngName, rngSh)
    End If

    'Now delete the row
    If delRow Then
      cellRng.EntireRow.Delete
      Set cellRng = savedRng.Offset(-1)
    Else
      Set cellRng = cellRng.Offset(-1)
      Set savedRng = savedRng.Offset(-1)
    End If
  Loop
End Sub

Private Function RngHasUpdate(ByVal rngName As String, ByVal rngSh As Worksheet) As Boolean
  Dim cellRng As Range
  Set cellRng = rngSh.Cells.Find(What:=rngName, lookAt:=xlWhole, MatchCase:=True)
  If Not (cellRng Is Nothing) Then _
  RngHasUpdate = (cellRng.Cells(1, 2).Value = "yes")
End Function

'Logics to update the columns with translation values
Private Sub UpdateTrans(ByVal rngSh As Worksheet)

  Const NAMEDRNGCOLNAME As String = "rngname" 'range of listColumn names
  Const TRANSLATEASTEXT As String = "translate as text"
  Const TRANSLATEASFORMULA As String = "translate as formula"

  Dim Lo As ListObject
  Dim upRng As Range
  Dim rngTag As String
  Dim rngName As String
  Dim rng As Range
  Dim counter As Long
  Dim nbTimes As Long
  Dim hasUpdate As Boolean
  Dim translateTag As String
  Dim tradCounter As Long
  Dim chunk As String

  'add a 1 to the number of times
  AddNumberOfTimes
  'add eventual languages if there is one.
  AddTransLang
  nbTimes = NumberTimes()

  For Each Lo In rngSh.ListObjects
    Set rng = Lo.ListColumns(NAMEDRNGCOLNAME).Range
    For counter = 1 To rng.Rows.Count
      rngName = rng.Cells(counter, 1).Value
      hasUpdate = (rng.Cells(counter, 2).Value = "yes")
      rngTag = rngName & "--" & nbTimes

      If RangeExists(rngName) And (hasUpdate Or (nbTimes = 1)) Then
        translateTag = rng.Cells(counter, 3).Value
        Set upRng = Range(rngName)
        If Not IsEmpty(upRng) Then
            Select Case translateTag
            Case TRANSLATEASTEXT
              For tradCounter = 1 To upRng.Rows.Count
                chunk = upRng.Cells(tradCounter, 1).Value
                AddChunk chunk, rngTag
              Next
            Case TRANSLATEASFORMULA
              For tradCounter = 1 To upRng.Rows.Count
                chunk = upRng.Cells(tradCounter, 1).Value
                AddFormulaChunk chunk, rngTag
              Next
            End Select
         End If
      End If
    Next
  Next
  'Now delete rows
  DeleteRows rngSh
  'Sort the listObject
  SortTransLo
  'Format the listObject
  FormatTable
  'Add condtional formatting to some of the columns
  AddConditionalFormatting
  'Count the number of elements and print a message
  PrintNumberOfMissing
End Sub

Private Sub FormatTable()
  Dim Lo As ListObject
  Dim hRng As Range
  Dim loRng As Range
  Dim sh As Worksheet
  Dim csTab As ICustomTable

  Set Lo = Table()
  Set hRng = Lo.HeaderRowRange()
  Set loRng = Lo.Range
  Set sh = Wksh()

  Lo.Range.Interior.color = vbWhite
  Lo.Range.Font.color = vbBlack
  Lo.Range.EntireRow.RowHeight = 30

  hRng.Font.color = RGB(34, 96, 162)
  hRng.Interior.color = RGB(221, 235, 247)
  hRng.Locked = False

  With Lo.Range
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .color = RGB(34, 96, 162)
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .color = RGB(34, 96, 162)
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .color = RGB(34, 96, 162)
    End With
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .color = RGB(34, 96, 162)
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .color = RGB(34, 96, 162)
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .color = RGB(34, 96, 162)
    End With
  End With

  'Unlock range for other languages
  sh.Range(hRng.Cells(2, 2), loRng.Cells(loRng.Rows.Count, hRng.Columns.Count)).Locked = False

  'resize the current table in the worksheet
  On Error Resume Next
  Set csTab = CustomTable.Create(Lo)
  csTab.RemoveRows totalCount:=1
  On Error GoTo 0
End Sub

'Add validation to the listObject
Private Sub AddConditionalFormatting()

  Dim Lo As ListObject
  Dim colRng As Range
  Dim counter As Long
  Dim sh As Worksheet

  Set Lo = Table()
  Set sh = Wksh()

  On Error Resume Next
  sh.Cells.FormatConditions.Delete
  On Error GoTo 0

  'Start with the second column of the listobject
  Set colRng = Lo.ListColumns(1).Range.Offset(1, 1)
  'Highlighting duplicates
  Do While Not (Intersect(colRng, Lo.Range) Is Nothing)
      'Delete previous conditional formatting before adding new one
      colRng.FormatConditions.AddUniqueValues
      With colRng.FormatConditions(colRng.FormatConditions.Count)
          .DupeUnique = xlDuplicate
          .Interior.color = vbRed
      End With
    Set colRng = colRng.Offset(, 1)
  Loop
End Sub

'Message with the number of missing values
Private Sub PrintNumberOfMissing()
  Dim Lo As ListObject
  Dim hRng As Range
  Dim colRng As Range
  Dim langName As String
  Dim nbMissing As Long
  Dim counter As Long
  Dim messageMissing As String

  messageMissing = "Translation Updated!" & Chr(10)
  Set Lo = Table()
  Set hRng = Lo.HeaderRowRange
  If (Not Lo.DataBodyRange Is Nothing) Then
    For counter = 1 To hRng.Columns.Count
      langName = hRng.Cells(1, counter).Value
      Set colRng = Lo.ListColumns(langName).DataBodyRange
      nbMissing = Application.WorksheetFunction.CountBlank(colRng)
      If nbMissing > 0 Then
        messageMissing = messageMissing & nbMissing & _
         " labels are missing for column " & langName & "." & Chr(10)
      End If
    Next
  End If

  MsgBox messageMissing, vbInformation + vbOkOnly, "Done!"
End Sub

Private Sub SortTransLo()
  Dim Lo As ListObject
  Dim loRng As Range
  Dim sh As Worksheet
  Dim sortRng As Range
  Dim keyColRng As Range
  Dim savedName As String

  Set Lo = Table()
  Set sh = Wksh()
  Set loRng = Lo.Range
  Set keyColRng = Lo.ListColumns(1).Range
  keyColRng.Cells(1, 0).Value = "TranslationTag"
  Set sortRng = sh.Range(keyColRng.Cells(1, 0), loRng.Cells(loRng.Rows.Count, loRng.Columns.Count))
  'Need to unlist the listobject before redefining it so that I can sort on all the sortRng.
  savedName = Lo.Name
  Lo.Unlist
  sortRng.Sort key1:=keyColRng, header:=xlYes, Orientation:=xlTopToBottom
  'Now redefine the listObject
  sh.ListObjects.Add(xlSrcRange, loRng, , xlYes).Name = savedName
  sh.ListObjects(1).TableStyle = vbNullString
End Sub


Private Function chunkExists(ByVal lab As String) As Boolean
  Dim langRng As Range
  Set langRng = KeyLangRange()

  If lab = vbNullString Then Exit Function

  chunkExists = Not (langRng.Find(What:=lab, lookAt:=xlWhole, MatchCase:=True) _
                     Is Nothing)
End Function

Private Property Get chunkRange(ByVal lab As String) As Range
  Dim langRng As Range
  Set langRng = KeyLangRange()
  Set chunkRange = langRng.Find(What:=lab, lookAt:=xlWhole, MatchCase:=True)
End Property

'Add a new language to the worksheet
Private Sub AddTransLang(Optional ByVal langVal As String = vbNullString)

  Dim sh As Worksheet
  Dim Lo As ListObject
  Dim hRng As Range
  Dim prevLoRng As Range
  Dim newLoRng As Range
  Dim cellRng As Range
  Dim startColRng As Range 'Start column for adding validations
  Dim langList As BetterArray


  Set Lo = Table()
  Set sh = Wksh()

  Set hRng = Lo.HeaderRowRange
  'Previous range of the listObject
  Set prevLoRng = Lo.Range

  If (langVal <> vbNullString) Then
    Set langList = New BetterArray
    langList.Items = Split(langVal, ";")
    langList.ToExcelRange hRng.Cells(1, hRng.Columns.Count + 1), TransposeValues:=True
  End If

  Set cellRng = hRng.Cells(1, hRng.Columns.Count + 1)

  If cellRng.Value <> vbNullString Then
    'Stop where there is no new language
    Do While (cellRng.Value <> vbNullString)
      Set cellRng = cellRng.Offset(, 1)
      'The last cellRng is the one with empty values
    Loop
    'Resize with new range of the listObject
    'now offset from number of lines of the listObject
    'since the last cellRng is the one with empty values, we need to go back
    Set cellRng = cellRng.Offset(prevLoRng.Rows.Count - 1, -1)
    Set newLoRng = sh.Range(hRng.Cells(1, 1), cellRng.Cells(1, 1))
    newLoRng.EntireColumn.ColumnWidth = 60
    newLoRng.EntireRow.RowHeight = 30
    Lo.Resize newLoRng
    'Format the new table
    FormatTable
    'Add conditional Formatting
    AddConditionalFormatting
  End If
End Sub

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Test if a range exists
Private Function RangeExists(ByVal rngName As String) As Boolean
  Dim rng As Range
  On Error Resume Next
  Set rng = Range(rngName)
  On Error GoTo 0
  RangeExists = (Not rng Is Nothing)
End Function

'Exposed methods
Private Sub ITranslations_UpdateTrans(ByVal rngSh As Worksheet)
  UpdateTrans rngSh
End Sub

Private Sub ITranslations_AddTransLang(ByVal langVal As String)
  AddTransLang langVal
End Sub
