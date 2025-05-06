VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Wrapper for Listobjects manipulations"

'@IgnoreModule ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Wrapper for Listobjects manipulations")


Option Explicit
Implements ICustomTable
Private Type TCustomTable
  Lo As listObject
  idCol As String
  idPrefix As String
  hasImport As Boolean
  hasChecks As Boolean
End Type

Private this As TCustomTable
Private Const CLASSNAME As String = "ICustomTable"
Private NOTIMPORTEDCOLUMNS As BetterArray
Private HIDDENCOLUMNSLIST As BetterArray

'checking objects
Private checkCounter As Integer
Private internalChecks As IChecking

'@section Instanciation

'@description
'The custom data table is created just with one listObject. It is a wrapper arround
'the listobject class. Which basically add some functionalities to ease the
'data-manipulation process. A custom data table can have a "id" column. This 
'is a column where values are not duplicated, used to locate lines. When adding
'new rows to a table, one can decide to automatically number the id column in 
'a sequential way (1, 2, 3, ...) or a a prefix before those numbers. ie "line 1", "line 2".
'The prefix and the id column are provided at instanciation.
'
'
'@param Lo A ListObject
'@param idCol A String, the ID column of the custom table. Default is empty string
'@param idPrefix A String. A prefix to add to each values of the ID column when adding rows.
'       Default is empty string.        
Public Function Create(ByVal Lo As listObject, _
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

'@section Custom table Elements

'@prop-title Worksheet of the listObject
Private Property Get Wksh() As Worksheet
  Dim Lo As listObject
  Set Lo = Table()
  Set Wksh = Lo.Parent
End Property

'Table listObject
Public Property Get Table() As listObject
  Set Table = this.Lo
End Property

Public Property Set Table(ByVal Lo As listObject)
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
  Dim Lo As listObject
  Set Lo = Table()
  Set HeaderRowRange = Lo.HeaderRowRange()
End Property

'Resize a table
Private Sub Resize( Optional ByVal AddRows As Boolean = False, _
                    Optional ByVal insertShift As Boolean = True, _
                    Optional ByVal totalRowCount As Long = 0, _
                    Optional ByVal nbRows As Long = 5, _ 
                    Optional ByVal includeIds As Boolean = True)
  Dim Lo As listObject
  'Begining of the tables
  Dim loRowHeader As Long
  Dim loColHeader  As Long
  Dim rowCounter As Long
  'End of the listobject table
  Dim loRowsEnd As Long
  Dim loColsEnd As Long
  Dim sh As Worksheet
  Dim delRng As Range
  Dim hRng As Range

  Set sh = Wksh()
  Set Lo = Table()
  Set hRng = HeaderRowRange()

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
      If insertShift Then 
        loRowsEnd = loRowsEnd + 1 'Start at the bottom of the table
        For rowCounter = 1 To nbRows + 1
              .Rows(loRowsEnd).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next
      End If
      loRowsEnd = loRowsEnd + nbRows
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))

    Else

      'Tend to delete formulas in the first line when every thing is empty.
      'I will let the user sort and then delete afterward instead.
      Do While (rowCounter > loRowHeader + 1)
        If (Application.WorksheetFunction.COUNTA(.Rows(rowCounter)) <= totalRowCount) Then

          If (delRng Is Nothing) Then
            Set delRng = hRng.Offset(rowCounter - loRowHeader)
          Else
            Set delRng = Application.Union(delRng,  hRng.Offset(rowCounter - loRowHeader))
          End If
          'update the end rows
          loRowsEnd = loRowsEnd - 1
        End If
        rowCounter = rowCounter - 1
      Loop
      'Delete the range if necessary
      If Not (delRng Is Nothing) Then delRng.Delete
      'Resize the new listObject
      loRowsEnd = Application.WorksheetFunction.Max(loRowsEnd, loRowHeader + 1)
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
    End If
  End With

  'Add Ids if it is required by the user
  If includeIds Then AddIds
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
    Dim Exist As Boolean
    Dim hRng As Range
    Set hRng = HeaderRowRange
 
    On Error Resume Next
    Exist = (Not hRng.Find(What:=colName, _
                           lookAt:=IIf(strictSearch, xlWhole, xlPart), _
                           MatchCase:=strictSearch, _
                           after:=hRng.Cells(1, hRng.Columns.Count)) Is Nothing)
    On Error GoTo 0
    ColumnExists = Exist
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
                         after:=hRng.Cells(1, hRng.Columns.Count) _
                         ).Column - hRng.Column + 1
    Else
        Index = -1
    End If
    ColumnIndex = Index
End Function

Private Property Get DataRange(Optional ByVal colName As String = "__all__", _
                               Optional ByVal includeHeaders As Boolean = False, _
                               Optional ByVal strictSearch As Boolean = False) As Range
  Dim Lo As listObject
  Set Lo = Table()
  Dim rng As Range

  If colName = "__all__" Then
    If includeHeaders Then
      Set rng = Lo.Range()
    Else
      Set rng = Lo.DataBodyRange()
      If rng Is Nothing Then Set rng = Lo.HeaderRowRange.Offset(1)
    End If
  Else
    If Not ColumnExists(colName, strictSearch) Then Exit Property
    
    If includeHeaders Then
      Set rng = Lo.ListColumns(ColumnIndex(colName, strictSearch)).Range
    Else
      'DataBody range could be Nothing
      On Error Resume Next
      Set rng = Lo.ListColumns(ColumnIndex(colName, strictSearch)).DataBodyRange
      On Error GoTo 0
      'You can test here if it is Nothing, in that case the table is Empty,
      'just return the second line
      If rng Is Nothing Then _
       Set rng = Lo.ListColumns(ColumnIndex(colName, strictSearch)).Range.Cells(2, 1)
    End If
  
  End If

  Set DataRange = rng
End Property

Private Property Get HeaderRange() As Range
  Dim Lo As listObject
  Set Lo = Table()
  Set HeaderRange = Lo.HeaderRowRange()
End Property

Private Property Get CellRange(ByVal colName As String, _
                               ByVal lineNum As Long) As Range
  Dim cellPos As Long
  Dim rng As Range

  If Not ColumnExists(colName) Then Exit Property
  Set rng = DataRange(colName)
  cellPos = lineNum - rng.Row + 1
  If cellPos < 0 Then Exit Property
  Set CellRange = rng.Cells(cellPos, 1)
End Property

Private Function ValueRange(ByVal colName As String, ByVal keyName As String) As Range

  Dim idRng As Range
  Dim colRng As Range

  'Iterations for id and return ranges
  Dim cellidRng As Range
  Dim cellRetRng As Range

  If IdValue = vbNullString Then Exit Function

  If (Not ColumnExists(colName)) Or  (Not ColumnExists(IdValue)) Then Exit Function
  Set idRng = DataRange(IdValue)
  Set colRng = DataRange(colName)
  Set cellidRng = idRng.Cells(1, 1)
  Set cellRetRng = colRng.Cells(1, 1)

  Do While (cellidRng.Value <> keyName) And _
           (cellidRng.Row <= idRng.Cells(idRng.Rows.Count, 1).Row)
    Set cellidRng = cellidRng.Offset(1)
    Set cellRetRng = cellRetRng.Offset(1)
  Loop

  If cellidRng.Row <= idRng.Cells(idRng.Rows.Count, 1).Row Then
    Set ValueRange = cellRetRng
  End If
End Function

Private Property Get Value(ByVal colName As String, _
                           ByVal keyName As String) As String
  Dim retValue As String
  Dim cellRetRng As Range

  'Get the value Range and retrieve (try) it value
  'On error (unable to convert to string for ex) nothing is done, returned value is null string
  On Error Resume Next
    Set cellRetRng = ValueRange(colName, keyName)
    retValue = CStr(cellRetRng.Value)
  On Error GoTo 0
  Value = retValue
End Property

Private Sub SetValue(ByVal colName As String, ByVal keyName As String, ByVal newValue As String)

  Dim cellRetRng As Range

  'Get the value Range and assign (try) it the newValue
  'On error (cellRetRng is Nothing for ex, nothing is done)
  On Error Resume Next
    Set cellRetRng = ValueRange(colName, keyName)
    cellRetRng.Value = newValue
  On Error GoTo 0
End Sub

Private Property Get Name() As String
  Name = Table.Name
End Property

'Clean every columns of the listObject and avoid formula columns
Private Sub Clean(Optional ByVal excludeNames As Object = Nothing, _ 
                  Optional ByVal totalCount As Long = 0, _
                  Optional ByVal clearContent As Boolean = False, _
                  Optional ByVal addNewRows As Boolean = False)
  
  Dim rng As Range
  Dim hRng As Range
  Dim Lo As listObject
  Dim headerValue As String
  Dim counter As Long
  Dim toClear As Boolean

  Set Lo = Table()
  Set hRng = Lo.HeaderRowRange

  If clearContent Then
    On Error Resume Next
    Lo.DataBodyRange.Delete
    On Error GoTo 0
  Else
    For counter = 1 To hRng.Columns.Count  
      headerValue = hRng.Cells(1, counter).Value
      Set rng = DataRange(headerValue)

      'On formula colunms, no clearing is done, clear only "non formula" columns
      toClear = (Not rng.Cells(1, 1).HasFormula)

      If (TypeName(excludeNames) = "BetterArray") Then toClear =  toClear And (Not excludeNames.Includes(headerValue))

      'clear the content if needed
      If toClear Then rng.ClearContents
    Next
  End If

  'The user want to add new rows after cleaning
  If addNewRows Then 
    'Now resize the listObject  
    Resize AddRows:=False, totalRowCount:=totalCount
  End If
End Sub


'Import another custom table to this one
'
'pasteAtBottom: paste the data at the bottom of the current table
'strictColumnSearch: Use strict rules when searching column names in the current table (case sensitive etc.)
'insertShift: Insert some lines when extending the table (Default is true)
'formatHeaders: A BetterArray of headers names where column format should be kept

Private Sub Import(ByVal impTab As Object, _
                   Optional ByVal pasteAtBottom As Boolean = False, _
                   Optional ByVal strictColumnSearch As Boolean = False, _ 
                   Optional ByVal insertShift As Boolean = True, _ 
                   Optional ByVal formatHeaders As Object = Nothing)

  Dim imphRng As Range
  Dim headerValue As String
  Dim counter As Long
  Dim nbImpRows As Long
  Dim impDataRng As Range
  Dim currDataRng As Range
  Dim rng As Range
  Dim sh As Worksheet
  Dim startPastingRow As Long

  If ((TypeName(impTab) <> "DataSheet") And (TypeName(impTab) <> "CustomTable")) Or _ 
     ((TypeName(impTab) = "CustomTable") And (LCase(Name()) <> LCase(impTab.Name()))) Then
    Exit Sub
  End If

  'Initialize and Clear the private variable NOTIMPORTEDCOLUMNS
  If NOTIMPORTEDCOLUMNS Is Nothing Then Set NOTIMPORTEDCOLUMNS = New BetterArray

  NOTIMPORTEDCOLUMNS.Clear

  'Unhide hidden columns before import
  UnhideHiddenColumns

  Set imphRng = impTab.HeaderRange()
  Set sh = Wksh()

  'This is where to start pasting the data in the custom table.
  'I can decide to paste a the end of the table, or clear the data
  'and replace with new one
  startPastingRow = 1

  If pasteAtBottom Then
    'DataBodyRange could be Nothing, in that case, the startPastingRow is 1.
    On Error Resume Next
    startPastingRow = Table.DataBodyRange.Rows.Count + 1
    On Error GoTo 0
    Resize AddRows:=True, nbRows:=impTab.DataRange.Rows.Count, insertShift:=insertShift
  End If

  For counter = 1 To imphRng.Columns.Count
    headerValue = imphRng.Cells(1, counter).Value

    If ColumnExists(headerValue, strictColumnSearch) Then
      If (TypeName(impTab) = "CustomTable") Then
        Set impDataRng = impTab.DataRange(colName:=headerValue, strictSearch:=strictColumnSearch)
      Else
        Set impDataRng = impTab.DataRange(headerValue)
      End If
      Set currDataRng = DataRange(colName:=headerValue, strictSearch:=strictColumnSearch)
      nbImpRows = impDataRng.Rows.Count
      
      'Resize to fit the number of required rows if needed
      If (nbImpRows > currDataRng.Rows.Count) And Not (pasteAtBottom) Then
        Resize AddRows:=True, nbRows:=(nbImpRows - currDataRng.Rows.Count), insertShift:=insertShift
        Set currDataRng = DataRange(colName:=headerValue, strictSearch:=strictColumnSearch)
      End If

      'Add values if the range does not contains formulas. If the range contains formulas, do nothing
      Set rng = sh.Range(currDataRng.Cells(startPastingRow, 1), _
                         currDataRng.Cells(startPastingRow + nbImpRows - 1, 1))

      If (Not rng.Cells(1, 1).HasFormula) Then
        'Clear the cell content only and only there is no paste
        'pasteAtBottom.
        If Not pasteAtBottom Then currDataRng.Cells.ClearContents
        rng.Value = impDataRng.Value
      End If

      'Import the format/Comment from some ranges
      If TypeName(formatHeaders) = "BetterArray" Then
        If formatHeaders.Includes(newHeaderValue) Then ImportFormat impRng:=impDataRng, currRng:=rng
      End If
    Else
      NOTIMPORTEDCOLUMNS.Push headerValue
    End If
  Next
  
  this.hasImport = True

  'return back hidden columns
  ReturnBackHiddenColumns

  'Format Headers are
End Sub

'Import All is the Sub that imports by keeping the imported object as the master
'one. It is assumed that all the columns of the imported Object SHOULD be imported, and
'will replace all the current columns WITH all the imported ones

Private Sub ImportAll(ByVal impTab As Object)

  Dim currhRng As Range
  Dim imphRng As Range
  Dim currLo As ListObject
  Dim sh As Worksheet
  Dim impTabList As BetterArray 'Data of imported Table
  Dim bodyRng As Range
  Dim cellRng As Range

  'I want to import by keeping all the headers of the imported tab.
  'First I need to clean the current tab and then I copy and paste
  'the data from the imported table
  'I would avoid doing it with DataSheet, just to make sure we are importing
  'the same "elements", but it is still possible do import All elements of a DataSheet

  If ((TypeName(impTab) <> "DataSheet") And (TypeName(impTab) <> "CustomTable")) Or _ 
      ((TypeName(impTab) = "CustomTable") And (LCase(Name()) <> LCase(impTab.Name()))) Then
    Exit Sub
  End If

  'Initialize and Clear the private variable NOTIMPORTEDCOLUMNS
  If NOTIMPORTEDCOLUMNS Is Nothing Then Set NOTIMPORTEDCOLUMNS = New BetterArray
  NOTIMPORTEDCOLUMNS.Clear

  'Unhide previously hidden columns in the custom table. Those will not be
  'returned back after importAll because of the importing mecanism
  UnhideHiddenColumns

  'Clean the actual customTable object
  Clean
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

  

  'Typically in this case all the columns from the imported Object are imported,
  'so normally the NOTIMPORTEDCOLUMNS object is empty.
  this.hasImport = True
End Sub

'This sub is an internal sub to import columns formats (If required)

'impRng: The importing range
'currRng the current range in the current CustomTable

Private Sub ImportFormat(impRng As Range, ByVal currRng As Range)

  Dim formatCounter As Long
  Dim hasComment As Boolean
  Dim hasCommentThreaded As Boolean
  Dim commentText As String
  Dim commentThreadedText As String
      
  'Clear all eventual comments on the range
  currRng.ClearComments

  For formatCounter = 1 To impDataRng.Rows.Count
    'keep the comment if there is one
    hasComment = False
    hasCommentThreaded = False
        
    'On some excel version there is no CommentThread. So it would
    'result in Error.
    On Error Resume Next
      hasCommentThreaded = (impRng.Cells(formatCounter, 1).CommentThreaded.Text <> vbNullString)
      hasComment = (impRng.Cells(formatCounter, 1).Comment.Text <> vbNullString)
      commentThreadedText = impRng.Cells(formatCounter, 1).CommentThreaded.Text
      commentText = impRng.Cells(formatCounter, 1).Comment.Text
    On Error GoTo 0

    'Add comment and commentThread are converted to comments. No threads is 
    'added here

    If hasComment Then
      On Error Resume Next
      currRng.Cells(formatCounter, 1).AddComment
      currRng.Cells(formatCounter, 1).Comment.Text Text:=commentText
      currRng.Cells(formatCounter, 1).Comment.Visible = False
      On Error GoTo 0
    End If

    If hasCommentThreaded Then
      On Error Resume Next
        currRng.Cells(formatCounter, 1).AddComment
        currRng.Cells(formatCounter, 1).Comment.Text Text:=commentThreadedText
      On Error GoTo 0
    End If
        'Add interiror color
    If (impRng.Cells(formatCounter, 1).Interior.color <> vbWhite) Then 
      currRng.Cells(formatCounter, 1).Interior.color = impRng.Cells(formatCounter, 1).Interior.color
    End If
        
    'Add Font color
    If (impRng.Cells(formatCounter, 1) <> vbBlack) Then
      currRng.Cells(formatCounter, 1).Font.color = impRng.Cells(formatCounter, 1).Font.color
    End If
    'Add Font weight and decoration
    currRng.Cells(formatCounter, 1).Font.Bold = impRng.Cells(formatCounter, 1).Font.Bold
    currRng.Cells(formatCounter, 1).Font.Italic = impRng.Cells(formatCounter, 1).Font.Italic
  Next
End Sub

Private Property Get ImportColumnsNotFound() As BetterArray
  Set ImportColumnsNotFound = NOTIMPORTEDCOLUMNS.Clone()
End Property

'Test if there are imports in the custom table
Private Property Get HasColumnsNotImported() As Boolean
  If this.hasImport Then _
  HasColumnsNotImported = (NOTIMPORTEDCOLUMNS.Length >= 1)
End Property

'Those functions are for Unhiding and hiding back columns before Import or Exports
'Sometimes custom table contains hidden columns, range.Find method does not include
'hidden rows or column when searching, so we need to unhide all columns of
'the worksheet before Import/Export and to return back the state of hidden columns
'after import/export.

Private Sub UnhideHiddenColumns()
  Dim counter As Long
  Dim hRng As Range

  Set hRng = HeaderRange()
  
  Set HIDDENCOLUMNSLIST = New BetterArray
  HIDDENCOLUMNSLIST.Clear

  For counter = 1 To hRng.Columns.Count
    If hRng.Cells(1, counter).EntireColumn.HIDDEN Then
      On Error Resume Next
      hRng.Cells(1, counter).EntireColumn.HIDDEN = False
      On Error GoTo 0
      HIDDENCOLUMNSLIST.Push counter
    End If
  Next
End Sub

Private Sub ReturnBackHiddenColumns()
  Dim counter As Long
  Dim hRng As Range

  Set hRng = HeaderRange()
  'We should keep in mind that a returnBack shoud be AFTER a unhide.
  'For safeguard purposes, I can work toward ignoring errors (variable not set for ex.)
  On Error Resume Next
    For counter = HIDDENCOLUMNSLIST.LowerBound To HIDDENCOLUMNSLIST.UpperBound
      hRng.Cells(1, HIDDENCOLUMNSLIST.Item(counter)).EntireColumn.HIDDEN = True
    Next
  On Error GoTo 0
End Sub


'Export a Custom Table to a Worksheet
'startLine: Where to start exporting
Private Sub Export(ByVal sh As Worksheet, ByVal headersList As BetterArray, Optional ByVal startLine As Long = 1)

  Dim counter As Long
  Dim inpRng As Range
  Dim outrng As Range
  Dim headerName As String
  Dim headerCounter As Long

  If sh Is Nothing Then ThrowError _
   "Unable to export custom table, sheet not initialized"
  If headersList Is Nothing Then ThrowError _
   "Unable to export custom table, headers list not initialized"
  If headersList.Length = 0 Then ThrowError _
  "Unable to export custom table, I don't know wich column to export (length of headers = 0)"

  sh.Cells.Clear
  headerCounter = 1

  'This will unhide all hidden columns and keep track of them
  UnhideHiddenColumns

  For counter = headersList.LowerBound To headersList.UpperBound

    headerName = headersList.Item(counter)
    
    If ColumnExists(headerName, strictSearch:=True) Then
      Set inpRng = DataRange(headerName, strictSearch:=True, includeHeaders:=True)
      inpRng.Calculate
      
      With sh
        Set outrng = .Range(.Cells(startLine, headerCounter), .Cells(inpRng.Rows.Count + startLine - 1, headerCounter))
        outrng.Value = inpRng.Value
      End With

      headerCounter = headerCounter + 1
    End If
    
  Next

  'Return back the previously hidden columns
  ReturnBackHiddenColumns
End Sub

'Sort a listObject on one column (on first element found)
Private Sub SortOnFirst(ByVal colName As String)

  Dim numRng As Range 'Number range
  Dim colRng As Range 'Colum range to sort on
  Dim LoRng As Range 'ListObject Range
  Dim hRng As Range 'Header Range of the list Object
  Dim counter As Long
  Dim trackData As BetterArray
  Dim sh As Worksheet
  Dim sortedValue As String
  Dim num As Long
  Dim numData As BetterArray
  Dim resRng As Range
  Dim Lo As listObject

  'Sort on values only If there is data
  Set Lo = Table()
  Set LoRng = Lo.DataBodyRange
  Set colRng = DataRange(colName, strictSearch:=True)

  If (LoRng Is Nothing) Or IsEmpty(LoRng) Or (colRng Is Nothing) Or _
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
  'Seems like sometimes the databodyrange of a new column
  'could be empty in that case exit
  If numRng Is Nothing Then
    Lo.ListColumns(hRng.Cells.Count + 1).Delete
    Exit Sub
  End If
  num = 1 'Number to sort later on
  'Resize the listObject to be sure it includes the number added
  With sh
      Set resRng = .Range(hRng.Cells(1, 1), LoRng.Cells(LoRng.Rows.Count, LoRng.Columns.Count + 1))
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
  resRng.Sort key1:=numRng, Order1:=xlAscending, Header:=xlYes
  'Remove the __number column
  Lo.ListColumns(hRng.Cells.Count + 1).Delete
End Sub

Private Sub SortSimple(ByVal colName As String)

  Dim colRng As Range
  Dim LoRng As Range
  Dim Lo As ListObject

  Set Lo = Table()
  Set LoRng = Lo.Range
  Set colRng = DataRange(colName, strictSearch:=True)

  If (LoRng Is Nothing) Or IsEmpty(LoRng) Or (colRng Is Nothing) Or _
      (IsEmpty(colRng)) Then Exit Sub

  LoRng.Sort key1:=colRng, Order1:=xlAscending, Header:=xlYes

End Sub

'Sucessively sort the listObject on multiple elements (in a BetterArray with colname values)
Private Sub SortElements(ByVal els As Variant, Optional ByVal onFirst As Boolean = True)
  Dim colName As String
  Dim counter As Long

  If (TypeName(els) <> "String") And (TypeName(els) <> "BetterArray") Then Exit Sub

  If (TypeName(els) = "String") Then
    colName = els
    If onFirst Then 
      SortOnFirst colName
    Else
      SortSimple colName
    End If
    Exit Sub
  End If

  For counter = els.LowerBound To els.UpperBound
    colName = els.Item(counter)
    If onFirst Then
      SortOnFirst colName
    Else
      SortSimple colName
    End If
  Next
End Sub

'@section Interface

'On Linelist, there are not stacked listObjects on one worksheet,
'So there is no need to insert shifts like in analysis custom table. Inserting
'shifts is time consuming.
Private Sub ICustomTable_AddRows(Optional ByVal nbRows As Long = 5, _ 
                                 Optional Byval insertShift As Boolean = False, _ 
                                 Optional ByVal includeIds As Boolean = True)

  Resize AddRows:=True, nbRows:=nbRows, insertShift:=insertShift, includeIds:=includeIds
End Sub

Private Sub ICustomTable_RemoveRows(Optional ByVal totalCount As Long = 0)
  Resize AddRows:=False, totalRowCount:=totalCount, includeIds:=includeIds
End Sub

Private Sub ICustomTable_SetValidation(ByVal colName As String, ByVal drop As IDropdownLists, _
                                      ByVal dropName As String, Optional ByVal alertType As String = "info")
    SetValidation colName, drop, dropName, alertType
End Sub

Private Property Get ICustomTable_DataRange( _
          Optional ByVal colName As String = "__all__", _
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

Private Property Get ICustomTable_HeaderRange() As Range
  Set ICustomTable_HeaderRange = HeaderRange()
End Property

Private Property Get ICustomTable_Name() As String
  ICustomTable_Name = Name()
End Property

Private Sub ICustomTable_Clean()
  Clean
End Sub

Private Sub ICustomTable_Sort(ByVal els As Variant, Optional ByVal onFirst As Boolean = True)
  SortElements els, onFirst
End Sub

Private Sub ICustomTable_Export(ByVal sh As Worksheet, _
                               ByVal headersList As BetterArray, _
                               Optional ByVal startLine As Long = 1)
  Export sh, headersList, startLine
End Sub

'Imports Functions
Private Sub ICustomTable_Import( _
              ByVal impTab As Object, _
              Optional ByVal pasteAtBottom As Boolean = False, _
              Optional ByVal strictColumnSearch As Boolean = False, _
              Optional ByVal insertShift As Boolean = True, _ 
              Optional ByVal keepImportHeaders As Boolean = False)

  'keepImportHeaders means the import Object is the master
  If keepImportHeaders Then
    ImportAll impTab
  Else
    Import impTab, pasteAtBottom, strictColumnSearch, insertShift
  End If
End Sub

Private Sub ICustomTable_SetValue(ByVal colName As String, ByVal keyName As String, _ 
                           ByVal newValue As String)
  SetValue colName, keyName, newValue
End Sub

Private Property Get ICustomTable_ImportColumnsNotFound() As BetterArray
  Set ICustomTable_ImportColumnsNotFound = ImportColumnsNotFound()
End Property

'Test if there are imports in the custom table
Private Property Get ICustomTable_HasColumnsNotImported() As Boolean
    ICustomTable_HasColumnsNotImported = HasColumnsNotImported()
End Property

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub
