VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A wrapper of ListObject"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("A wrapper of ListObject")
'@IgnoreModule ReadOnlyPropertyAssignment, ExcelMemberMayReturnNothing, IIfSideEffect

Option Explicit
Implements ICustomTable
Private Type TCustomTable
  Lo As ListObject
  idCol As String
  idPrefix As String
End Type

Private this As TCustomTable
Private Const CLASSNAME As String = "ICustomTable"
Private Const NBROWSADD As Long = 5 'Number of Rows to add when resizing
Private Const TABTRANSLATIONS As String = "Tab_Translations"
Private Const TABEXPORT As String = "Tab_Export"


Public Function Create(ByVal Lo As ListObject, _
                      Optional ByVal idCol As String = vbNullString, _
                      Optional ByVal idPrefix As String = vbNullString) As ICustomTable
  If Lo Is Nothing Then ThrowError "Unable to initialize the class, no listobject provided"

  With New CustomTable
    Set .Table = Lo
    .IdValue = idCol
    .PrefixValue = idPrefix
    Set Create = .Self()
  End With
End Function


Public Property Get Self() As ICustomTable
  Set Self = Me
End Property

'Worksheet of the listObject
Private Function Wksh() As Worksheet
  Dim Lo As ListObject
  Set Lo = Table()
  Set Wksh = Lo.Parent
End Function


'Table listObject
Public Property Get Table() As ListObject
  Set Table = this.Lo
End Property

Public Property Set Table(ByVal Lo As ListObject)
  Set this.Lo = Lo
End Property

Public Property Get IdValue() As String
  IdValue = this.idCol
End Property

Public Property Get PrefixValue() As String
  PrefixValue = this.idPrefix
End Property

Public Property Let IdValue(ByVal idCol As String)
  this.idCol = idCol
End Property

Public Property Let PrefixValue(ByVal idPrefix As String)
  this.idPrefix = idPrefix
End Property

'HeaderRowRange
Private Property Get HeaderRowRange() As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Set HeaderRowRange = Lo.HeaderRowRange()
End Property

'Resize a table
Private Sub Resize(Optional ByVal AddRows As Boolean = False, _ 
                   Optional ByVal totalRowCount As Long = 0, _ 
                   Optional ByVal nbRows As Long = NBROWSADD)
  Dim Lo As ListObject
  'Begining of the tables
  Dim loRowHeader As Long
  Dim delStop As Long 'where to stop when deleting rows of a table
  Dim loColHeader  As Long
  Dim rowCounter As Long
  'End of the listobject table
  Dim loRowsEnd As Long
  Dim loColsEnd As Long
  
  Dim sh As Worksheet

  Set sh = Wksh()
  Set Lo = Table()

  'Remove unwanted rows from each listObject
  With sh
    'Rows and columns at the begining of the table to resize
    loRowHeader = Lo.Range.Row
    loColHeader = Lo.Range.Column
   'Rows and Columns at the end of the Table to resize
    loRowsEnd = loRowHeader + Lo.Range.Rows.Count - 1
    loColsEnd = loColHeader + Lo.Range.Columns.Count - 1
    rowCounter = loRowsEnd
    If AddRows Then
      loRowsEnd = loRowsEnd + 1 'Start at the bottom of the table
      For rowCounter = 1 To  nbRows + 1
            .Rows(loRowsEnd).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      Next
      loRowsEnd = loRowsEnd + nbRows
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
      AddIds
    Else
      'On analysis sheets, there are formulas, you should stop/avoid the first row
      delStop = IIf(sh.Name = "Analysis", loRowHeader + 1, loRowHeader)
      Do While (rowCounter > delStop)
        If (Application.WorksheetFunction.CountA(.Rows(rowCounter)) <= totalRowCount) Then
          On Error Resume Next
          .Rows(rowCounter).EntireRow.Delete
          On Error GoTo 0
          'update the end rows
          loRowsEnd = loRowsEnd - 1
        End If
        rowCounter = rowCounter - 1
      Loop
      loRowsEnd = Application.WorksheetFunction.Max(loRowsEnd, loRowHeader + 1)
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
      'Add the IDs after resizing
      AddIds
    End If
  End With
End Sub

Private Sub AddIds()
  Dim rng As Range
  Dim counter As Long
  Dim idPref As String
  Dim ID As String

  ID = IdValue()
  If (ID = vbNullString) Then Exit Sub
  If Not (ColumnExists(ID)) Then Exit Sub
  idPref = PrefixValue()
  Set rng = DataRange(ID)
  For counter = 1 To rng.Rows.Count
    rng.Cells(counter, 1).Value = Application.WorksheetFunction.Trim(idPref & " " & counter)
  Next
  rng.Locked = rng.Cells(1, 1).Locked
End Sub


Private Sub SetValidation(ByVal colName As String, _
                          ByVal drop As IDropdownLists, _
                          ByVal dropName As String, _
                          Optional ByVal alertType As String = "info")

    Dim rng As Range

    'Test if the colname exists before seting the validation
    If Not ColumnExists(colName) Then Exit Sub
    Set rng = DataRange(colName)
    'set validation for each cell in the datarange
    drop.SetValidation cellRng:=rng, listName:=dropName, _
                       alertType:=alertType
End Sub


'Column Range and column Index of a custom table
Private Function ColumnExists(ByVal colName As String, _
                              Optional ByVal strictSearch As Boolean = False) As Boolean
    Dim exist As Boolean
    Dim hRng As Range
    Set hRng = HeaderRowRange
    exist = False
    On Error Resume Next
    exist = (Not hRng.Find(What:=colName, _
                           lookAt:=IIf(strictSearch, xlWhole, xlPart), _
                           MatchCase:=strictSearch, _
                           After:=hRng.Cells(1, hRng.Columns.Count)) Is Nothing)
    On Error GoTo 0
    ColumnExists = exist
End Function

Private Function ColumnIndex(ByVal colName As String, _
                             Optional ByVal strictSearch As Boolean = False) As Long
    Dim Index As Long
    Dim hRng As Range

    Set hRng = HeaderRowRange()
    If ColumnExists(colName) Then
        Index = hRng.Find(What:=colName, _
                         lookAt:=IIf(strictSearch, xlWhole, xlPart), _
                         MatchCase:=strictSearch, _
                         After:=hRng.Cells(1, hRng.Columns.Count) _
                         ).Column - hRng.Column + 1
    Else
        Index = -1
    End If
    ColumnIndex = Index
End Function

Private Property Get DataRange(ByVal colName As String, _
                               Optional ByVal includeHeaders As Boolean = False, _
                               Optional ByVal strictSearch As Boolean = False) As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Dim rng As Range

  If Not ColumnExists(colName, strictSearch) Then Exit Property
  If includeHeaders Then
    Set rng = Lo.ListColumns(ColumnIndex(colName, strictSearch)).Range
  Else
    'DataBody range could be Nothing
    On Error Resume Next
    Set rng = Lo.ListColumns(ColumnIndex(colName, strictSearch)).DataBodyRange
    On Error GoTo 0
    'You can test here if it is Nothing, in that case the table is Empty, just return the second line
    If rng Is Nothing Then Set rng = Lo.ListColumns(ColumnIndex(colName, strictSearch)).Range.Cells(2, 1)
    Set DataRange = rng
  End If
End Property

Private Property Get HeaderRange() As Range
  Dim Lo As ListObject
  Set Lo = Table()
  Set HeaderRange = Lo.HeaderRowRange()
End Property

Private Property Get CellRange(ByVal colName As String, ByVal lineNum As Long) As Range
  Dim cellPos As Long
  Dim rng As Range

  If Not ColumnExists(colName) Then Exit Property
  Set rng = DataRange(colName)
  cellPos = lineNum - rng.Row + 1
  If cellPos < 0 Then Exit Property
  Set CellRange = rng.Cells(cellPos, 1)
End Property

  Private Property Get Value(ByVal colName As String, ByVal keyName As String) As String

    Dim idRng As Range
    Dim colRng As Range
    Dim retValue As String
  
    'Iterations for id and return ranges
    Dim cellidRng As Range
    Dim cellRetRng As Range
  
    If IdValue = vbNullString Then Exit Property
  
    If (Not ColumnExists(colName)) Or (Not ColumnExists(IdValue)) Then Exit Property
    Set idRng = DataRange(IdValue)
    Set colRng = DataRange(colName)
    Set cellidRng = idRng.Cells(1, 1)
    Set cellRetRng = colRng.Cells(1, 1)
  
    Do while (cellidRng.Value <> keyName) And _ 
             (cellidRng.Row <= idRng.Cells(idRng.Rows.Count, 1).Row)
      Set cellidRng = cellidRng.Offset(1)
      Set cellRetRng = cellRetRng.Offset(1)
    Loop

    retValue = vbNullString
    On Error Resume Next
      If cellidRng.Row <= idRng.Cells(idRng.Rows.Count, 1).Row Then
        retValue = CStr(cellRetRng.Value)
      End If
    On Error GoTo 0
  
    Value = retValue
  End Property

Private Property Get Name() As String
  Name = Table.Name
End Property

'Clean every columns of the listObject and avoid formula columns
Private Sub Clean(ByVal excludeNames As BetterArray)
  Dim rng As Range
  Dim hRng As Range
  Dim Lo As ListObject
  Dim headerValue As String
  Dim counter As Long

  Set Lo = Table()
  Set hRng = Lo.HeaderRowRange

  For counter = 1 To hRng.Columns.Count
    headerValue = hRng.Cells(1, counter).Value
    Set rng = DataRange(headerValue)

    If (Not rng.Cells(1, 1).HasFormula) And _ 
       (Not excludeNames.Includes(headerValue)) Then 
      rng.ClearContents
    End If
  Next

  'For translation table, delete the data bodyrange
  If (Name() = TABTRANSLATIONS) Then
    On Error Resume Next
    Lo.DataBodyRange.Delete
    On Error GoTo 0
  End If
End Sub

Private Sub ImportAll(ByVal impTab As ICustomTable)
  Dim currhRng As Range
  Dim imphRng As Range
  Dim currLo As ListObject
  Dim sh As Worksheet
  Dim impTabList As BetterArray 'Data of imported Table
  Dim bodyRng As Range
  Dim cellRng As Range
  Dim excludeNames As BetterArray

  'I want to import by keeping all the headers of the imported tab.
  'First I need to clean the current tab and then I copy and paste
  'the data from the imported table

  If Name() <> impTab.Name() Then Exit Sub
  Set excludeNames = New BetterArray

  Clean excludeNames:=excludeNames
  Set currLo = Table()
  Set currhRng = currLo.HeaderRowRange()
  Set imphRng = impTab.HeaderRange()
  Set impTabList = New BetterArray
  Set sh = Wksh()

  'Get the data from imported listObject
  impTabList.FromExcelRange imphRng.Cells(1, 1), _ 
                            DetectLastRow:=True, DetectLastColumn:=True

  'AddRows to the current Table
  Resize AddRows:=True, nbRows:=impTabList.Length
  'Paste data with betterarray
  impTabList.ToExcelRange currhRng.Cells(1, 1)
  'Resize the new listObject with imported elements
  With sh
    Set bodyRng = .Range(currhRng.Cells(1, 1), currhRng.Cells(impTabList.Length, imphRng.Columns.Count))
    currLo.Resize bodyRng
  End With

  'Clear elements after the end of the table (in case imported table has
  'less columns compare to current table)
  Set cellRng = currhRng.Cells(1, imphRng.Columns.Count + 1)

  Do while Not IsEmpty(cellRng)
    cellRng.EntireColumn.Cells.Clear
    cellRng.EntireColumn.Locked = False
    Set cellRng = cellRng.Offset(, 1)
  Loop

End Sub


'Import another custom table to this one

Private Sub Import(ByVal impTab As ICustomTable)

  Dim imphRng As Range
  Dim headerValue As String 
  Dim newHeaderValue As String
  Dim counter As Long
  Dim formatCounter As Long
  Dim nbImpRows As Long
  Dim impDataRng As Range
  Dim currDataRng As Range
  Dim rng As Range
  Dim sh As Worksheet
  Dim formatHeaders As BetterArray
  Dim hasComment As Boolean
  Dim hasCommentThreaded As Boolean
  Dim commentText As String
  Dim commentThreadedText As String

  'Test if the custom tables have the same name, otherwise exit.
  If Name() <> impTab.Name() Then Exit Sub

  Set imphRng = impTab.HeaderRange()
  Set sh = Wksh()
  'List of headers where we should import comments and formatting.
  Set formatHeaders = New BetterArray
  formatHeaders.Push "Variable Name", "Main Label", "Dev Comments", _ 
                     "Sub Label", "Formatting Values", "Control", _ 
                     "Control Details"

  For counter = 1 To imphRng.Columns.Count
    headerValue = imphRng.Cells(1, counter).Value
    'Add code for some header values for importing from old setup
    newHeaderValue = Switch(headerValue = "Type", "Variable Type", _
                         headerValue = "Format", "Variable Format", _
                         headerValue = "Export", "Export Number", _
                         True, headerValue)

    If ColumnExists(newHeaderValue) Then
      Set impDataRng = impTab.DataRange(colName:=headerValue, strictSearch:=True)
      Set currDataRng = DataRange(newHeaderValue)
      nbImpRows = impDataRng.Rows.Count
      'Resize to fit the number of required rows
      Do While (nbImpRows > currDataRng.Rows.Count)
        Resize AddRows:=True
        Set currDataRng = DataRange(newHeaderValue)
      Loop
      'Add values if the range does not contains formulas. If the range contains formulas, do nothing
      Set rng = sh.Range(currDataRng.Cells(1, 1), currDataRng.Cells(nbImpRows, 1))

      If Not (rng.Cells(1, 1).HasFormula) Then
        currDataRng.Cells.ClearContents
        rng.Value = impDataRng.Value
      End If

      'Import the columns formattings for the conditional formatting and variable Name
      If formatHeaders.Includes(newHeaderValue) Then
        
        'Clear all eventual comments on the range
        rng.ClearComments

        For formatCounter = 1 To impDataRng.Rows.Count

          'keep the comment if there is one
          hasComment = False
          hasCommentThreaded = False
          
          'On some excel version there is no CommentThread. So it would
          'result in Error.
          On Error Resume Next
            hasCommentThreaded = (impDataRng.Cells(formatCounter, 1).CommentThreaded.Text <> vbNullString)
            hasComment = (impDataRng.Cells(formatCounter, 1).Comment.Text <> vbNullString)
            commentThreadedText = impDataRng.Cells(formatCounter, 1).CommentThreaded.Text
            commentText = impDataRng.Cells(formatCounter, 1).Comment.Text
          On Error GoTo 0

          If hasComment Then
            On Error Resume Next
            rng.Cells(formatCounter, 1).AddComment
            rng.Cells(formatCounter, 1).Comment.Text Text:=commentText
            rng.Cells(formatCounter, 1).Comment.Visible = False
            On Error GoTo 0
          End If

          If hasCommentThreaded Then
            On Error Resume Next
              rng.Cells(formatCounter, 1).AddComment
              rng.Cells(formatCounter, 1).Comment.Text Text:=commentThreadedText
            On Error GoTo 0
          End If

          'Add interiror color
          If (impDataRng.Cells(formatCounter, 1).Interior.color <> vbWhite) Then 
            rng.Cells(formatCounter, 1).Interior.color = impDataRng.Cells(formatCounter, 1).Interior.color
          End If
          
          'Add Font color
          If (impDataRng.Cells(formatCounter, 1) <> vbBlack) Then
            rng.Cells(formatCounter, 1).Font.color = impDataRng.Cells(formatCounter, 1).Font.color
          End If

          'Add Font weight and decoration
          rng.Cells(formatCounter, 1).Font.Bold = impDataRng.Cells(formatCounter, 1).Font.Bold
          rng.Cells(formatCounter, 1).Font.Italic = impDataRng.Cells(formatCounter, 1).Font.Italic
        Next
      End If
    End If
  Next
End Sub

'Sort a listObject on one column (on first element found)
Private Sub SortOnFirst(ByVal colName As String)

  Dim numRng As Range 'Number range
  Dim colRng As Range 'Colum range to sort on
  Dim loRng As Range 'ListObject Range
  Dim hRng As Range 'Header Range of the list Object
  Dim counter As Long
  Dim trackData As BetterArray
  Dim sh As Worksheet
  Dim sortedValue As String
  Dim num As Long
  Dim numData As BetterArray
  Dim resRng As Range
  Dim Lo As ListObject

  'Sort on values only If there is data
  Set Lo = Table()
  Set loRng = Lo.DataBodyRange
  Set colRng = DataRange(colName)

  If (loRng Is Nothing) Or IsEmpty(loRng) Or (colRng Is Nothing) Or _
     (IsEmpty(colRng)) Then Exit Sub

  Set trackData = New BetterArray
  Set numData = New BetterArray
  Set sh = Wksh()
  Set hRng = Lo.HeaderRowRange

  'Test values progressively
  'If the trackData does not contains the values,
  'Update the number (where you will sort)
  hRng.Cells(1, hRng.Cells.Count + 1).Value = "__number"
  Set numRng = Lo.ListColumns(hRng.Cells.Count + 1).DataBodyRange
  'Seems like sometimes the databodyrange of a new column could be empty in that case exit
  If numRng Is Nothing Then
    Lo.ListColumns(hRng.Cells.Count + 1).Delete
    Exit Sub
  End If
  num = 1 'Number to sort later on
  'Resize the listObject to be sure it includes the number added
  With sh
      Set resRng = .Range(hRng.Cells(1, 1), loRng.Cells(loRng.Rows.Count, loRng.Columns.Count + 1))
      Lo.Resize resRng
  End With

  For counter = 1 To colRng.Cells.Count
     sortedValue = colRng.Cells(counter, 1).Value
    If Not trackData.Includes(sortedValue) Then
      trackData.Push sortedValue
      numData.Push num
      numRng.Cells(counter, 1).Value = num
      num = num + 1
    Else
      numRng.Cells(counter, 1) = numData.Items(trackData.IndexOf(sortedValue))
    End If
  Next
  'Now sort the listObject on the __number column
  resRng.Sort key1:=numRng, order1:=xlAscending, header:=xlYes
  'Remove the __number column
  Lo.ListColumns(hRng.Cells.Count + 1).Delete
End Sub

'Sucessively sort the listObject on multiple elements (in a BetterArray with colname values)
Private Sub SortElements(ByVal els As Variant)
  Dim colName As String
  Dim counter As Long

  If (TypeName(els) <> "String") And (TypeName(els) <> "BetterArray") Then Exit Sub

  If (TypeName(els) = "String") Then
    colName = els
    SortOnFirst colName
    Exit Sub
  End If

  For counter = els.LowerBound To els.UpperBound
    colName = els.Item(counter)
    SortOnFirst colName
  Next
End Sub

Private Sub ICustomTable_AddRows(Optional ByVal nbRows As Long = NBROWSADD)
  Resize AddRows:=True, nbRows:=nbRows
End Sub

Private Sub ICustomTable_RemoveRows(Optional ByVal totalCount As Long = 0)
  Resize AddRows:=False, totalRowCount:=totalCount
End Sub

Private Sub ICustomTable_SetValidation(ByVal colName As String, ByVal drop As IDropdownLists, _
                                      ByVal dropName As String, Optional ByVal alertType As String = "info")
    SetValidation colName, drop, dropName, alertType
End Sub

Private Property Get ICustomTable_DataRange(ByVal colName As String, _
                                          Optional ByVal includeHeaders As Boolean = False, _
                                          Optional ByVal strictSearch As Boolean = False) As Range
  Set ICustomTable_DataRange = DataRange(colName, includeHeaders, strictSearch)
End Property

Private Property Get ICustomTable_Value(ByVal colName As String, ByVal keyName As String) As String
  ICustomTable_Value = Value(colName, keyName)
End Property

Private Property Get ICustomTable_CellRange(ByVal colName As String, ByVal lineNum As Long) As Range
  Set ICustomTable_CellRange = CellRange(colName, lineNum)
End Property

Private Sub ICustomTable_Import(ByVal impTab As ICustomTable, Optional ByVal keepImportHeaders As Boolean = False)

  'If you want to keep import headers, use import all instead.
  If keepImportHeaders Then
    ImportAll impTab
  Else
    Import impTab
  End If

End Sub

Private Property Get ICustomTable_HeaderRange() As Range
  Set ICustomTable_HeaderRange = HeaderRange()
End Property

Private Property Get ICustomTable_Name() As String
  ICustomTable_Name = Name()
End Property

Private Sub ICustomTable_Clean(ByVal excludeNames As BetterArray)
  Clean excludeNames
End Sub

Private Sub ICustomTable_Sort(ByVal els As Variant)
  SortElements els
End Sub

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub
