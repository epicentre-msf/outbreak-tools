VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Wrapper for Listobjects manipulations")


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

Private Sub BusyApp()
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Application.EnableAnimations = False
  Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  Application.EnableAnimations = True
  Application.Calculation = xlCalculationAutomatic
End Sub

'Resize a table
Private Sub Resize(Optional ByVal AddRows As Boolean = False, _ 
                  Optional ByVal totalRowCount As Long = 0, _ 
                  Optional ByVal nbRows As Long = NBROWSADD)
  Dim Lo As ListObject
  'Begining of the tables
  Dim loRowHeader As Long
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
      For rowCounter = 1 To nbRows + 1
            .Rows(loRowsEnd).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      Next
      loRowsEnd = loRowsEnd + nbRows
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
      AddIds
    Else
      Do While (rowCounter > loRowHeader + 1)
        If (Application.WorksheetFunction.COUNTA(.Rows(rowCounter)) <= totalRowCount) Then
          .Rows(rowCounter).EntireRow.Delete
          'update the end rows
          loRowsEnd = loRowsEnd - 1
        End If
        rowCounter = rowCounter - 1
      Loop
      Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
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
End Sub


Private Sub SetValidation(ByVal colName As String, _
                          ByVal drop As IDropdownLists, _
                          ByVal dropName As String, _
                          Optional ByVal alertType As String = "info")

    Dim rng As Range
    Dim sh As Worksheet

    Set sh = Wksh()
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
  Dim sh As Worksheet
  Dim FUN As WorksheetFunction
  Dim retValue As Variant

  'Iterations for id and return ranges
  Dim cellidRng As Range
  Dim cellRetRng As Range

  If IdValue = vbNullString Then Exit Property

  Set FUN = Application.WorksheetFunction
  Set sh = Wksh()
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

  On Error Resume Next
    If cellidRng.Row <= idRng.Cells(idRng.Rows.Count, 1).Row Then
      retValue = cellRetRng.Value
    End If
  On Error GoTo 0

  If TypeName(retValue) = "String" Then Value = retValue
End Property

Private Property Get Name() As String
  Name = Table.Name
End Property

'Clean every columns of the listObject and avoid formula columns
Private Sub Clean()
  Dim rng As Range
  Dim hRng As Range
  Dim Lo As ListObject
  Dim headerValue As String
  Dim counter As Long
  Dim toClear As Boolean
  Dim totalCount As Long

  Set Lo = Table()
  Set hRng = Lo.HeaderRowRange
  totalCount = 0

  For counter = 1 To hRng.Columns.Count
    headerValue = hRng.Cells(1, counter).Value
    Set rng = DataRange(headerValue)
    toClear = False
    On Error Resume Next
    toClear = (CStr(rng.Cells(1, 1).formula) = CStr(rng.Cells(1, 1).Value))
    On Error GoTo 0
    If toClear Then rng.ClearContents
  Next

  If Name() = "Tab_Translations" Then
    On Error Resume Next
    Lo.DataBodyRange.ClearContents
    On Error GoTo 0
    totalCount = 1
  End If
  'Now resize the listObject
  If Name() <> "Tab_Export" Then Resize AddRows:=False, totalRowCount:=totalCount
End Sub

'Import another custom table to this one

Private Sub Import(ByVal impTab As Object)

  Dim imphRng As Range
  Dim headerValue As String
  Dim counter As Long
  Dim nbImpRows As Long
  Dim impDataRng As Range
  Dim currDataRng As Range
  Dim rng As Range
  Dim sh As Worksheet
  Dim toClear As Boolean

  If (TypeName(impTab) <> "DataSheet") And (TypeName(impTab) <> "CustomTable") Then _
    Exit Sub

  If (TypeName(impTab) = "CustomTable") Then
    'Test if the custom tables have the same name, otherwise exit.
    If Name() <> impTab.Name() Then Exit Sub
  End If

  Set imphRng = impTab.HeaderRange()
  Set sh = Wksh()

  For counter = 1 To imphRng.Columns.Count
    headerValue = imphRng.Cells(1, counter).Value

    If ColumnExists(headerValue) Then
      If (TypeName(impTab) = "CustomTable") Then
        Set impDataRng = impTab.DataRange(colName:=headerValue, strictSearch:=True)
      Else
        Set impDataRng = impTab.DataRange(headerValue)
      End If
      Set currDataRng = DataRange(headerValue)
      nbImpRows = impDataRng.Rows.Count
      'Resize to fit the number of required rows
      Do While (nbImpRows > currDataRng.Rows.Count)
        Resize AddRows:=True
        Set currDataRng = DataRange(headerValue)
      Loop
      'Add values if the range does not contains formulas. If the range contains formulas, do nothing
      Set rng = sh.Range(currDataRng.Cells(1, 1), currDataRng.Cells(nbImpRows, 1))
      toClear = False

      On Error Resume Next
      toClear = (CStr(rng.Cells(1, 1).formula) = CStr(rng.Cells(1, 1).Value))
      On Error GoTo 0

      If toClear Then
        currDataRng.Cells.ClearContents
        rng.Value = impDataRng.Value
      End If
    End If
  Next
End Sub

'Export a Custom Table to a Worksheet

Private Sub Export(ByVal sh As Worksheet, ByVal headersList As BetterArray, _ 
                   Optional ByVal useLabels As Boolean = False)

  Dim counter As Long
  Dim inpRng As Range
  Dim outRng As Range
  Dim headerName As String
  Dim headerCounter As Long

  If sh Is Nothing Then ThrowError "Unable to export custom table, sheet not initialized"
  If headersList Is Nothing Then ThrowError "Unable to export custom table, headers list not initialized"
  If headersList.Length = 0 Then ThrowError "Unable to export custom table, I don't know wich column to export (length of headers = 0)"

  sh.Cells.Clear
  headerCounter = 1

  For counter = headersList.LowerBound to headersList.UpperBound

    headerName = headersList.Item(counter)
    
    If ColumnExists(headerName, strictSearch:=True) Then
      Set inpRng = DataRange(headerName, strictSearch:=True, includeHeaders:=True)
      
      With sh
        Set outRng = .Range(.Cells(1, headerCounter), .Cells(inpRng.Rows.Count, headerCounter))
        outRng.Value = inpRng.Value
      End With
        
      'assuming there could be two headers/ top header for labels, second header for variable names
      On Error Resume Next
        'Enter is for notes, if you want to use labels, just remove notes (after the Enter)
        If useLabels Then outRng.Cells(1, 1).Value = Split(inpRng.Cells(0, 1).Value, Chr(10))(0)
      On Error GoTo 0

      headerCounter = headerCounter + 1
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

Public Sub ICustomTable_AddRows(Optional ByVal nbRows As Long = 100)
  Resize AddRows:=True, nbRows:=nbRows
End Sub

Public Sub ICustomTable_RemoveRows(Optional ByVal totalCount As Long = 0)
  Resize AddRows:=False, totalRowCount:=totalCount
End Sub

Public Sub ICustomTable_SetValidation(ByVal colName As String, ByVal drop As IDropdownLists, _
                                      ByVal dropName As String, Optional ByVal alertType As String = "info")
    SetValidation colName, drop, dropName, alertType
End Sub

Public Property Get ICustomTable_DataRange(ByVal colName As String, _
                                          Optional ByVal includeHeaders As Boolean = False, _
                                          Optional ByVal strictSearch As Boolean = False) As Range
  Set ICustomTable_DataRange = DataRange(colName, includeHeaders, strictSearch)
End Property

Public Property Get ICustomTable_Value(ByVal colName As String, ByVal keyName As String) As String
  ICustomTable_Value = Value(colName, keyName)
End Property

Public Property Get ICustomTable_CellRange(ByVal colName As String, ByVal lineNum As Long) As Range
  Set ICustomTable_CellRange = CellRange(colName, lineNum)
End Property

Public Sub ICustomTable_Import(ByVal impTab As Object)
  Import impTab
End Sub

Public Property Get ICustomTable_HeaderRange() As Range
  Set ICustomTable_HeaderRange = HeaderRange()
End Property

Public Property Get ICustomTable_Name() As String
  ICustomTable_Name = Name()
End Property

Public Sub ICustomTable_Clean()
  Clean
End Sub

Public Sub ICustomTable_Sort(ByVal els As Variant)
  SortElements els
End Sub

Private Sub ICustomTable_Export(ByVal sh As Worksheet, _ 
                               ByVal headersList As BetterArray, _ 
                               Optional ByVal useLabels As Boolean = False)
  Export sh, headersList, useLabels
End Sub

'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub
