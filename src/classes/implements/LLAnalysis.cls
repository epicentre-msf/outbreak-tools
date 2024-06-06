VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Analysis Class"

'@Folder("Analysis")
'@ModuleDescription("Analysis Class")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLAnalysis

Private Type TLLAnalysis
    anash As Worksheet
    lData As ILinelistSpecs
End Type

Private this As TLLAnalysis
Private Const CLASSNAME As String = "LLAnalysis"
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const TITLEGRAPHTIMESERIES As String = "Tab_Label_TSGraph"
Private Const SPATIAL As String = "Tab_Spatial_Analysis"
Private Const SPATIOTEMPORAL As String = "Tab_SpatioTemporal_Analysis"
Private Const SPATIOTEMPSPECS As String = "Tab_SpatioTemporal_Specs"

'Exposed methods
Public Function Create(ByVal anaWksh As Worksheet) As ILLAnalysis
    If anaWksh Is Nothing Then ThrowErrorUnableInstanciate "Worksheet object not defined"
    'There is no cheking at instanciation since the worksheet could be empty
    With New LLAnalysis
        Set .Wksh = anaWksh
        Set Create = .Self
    End With
End Function

Public Property Get Wksh() As Worksheet
    Set Wksh = this.anash
End Property

Public Property Set Wksh(ByVal anash As Worksheet)
    Set this.anash = anash
End Property

Public Property Get Self() As ILLAnalysis
    Set Self = Me
End Property

Private Sub CheckTablesRequirements()
    Dim loTable As BetterArray
    Dim counter As Long
    Dim Lo As listObject
    Dim loName As String

    Set loTable = New BetterArray

    loTable.Push GLOBALSUMMARY, UNIVARIATE, BIVARIATE, TIMESERIES, _
                 GRAPHTIMESERIES, TITLEGRAPHTIMESERIES, SPATIAL, SPATIOTEMPORAL, _ 
                 SPATIOTEMPSPECS

    For counter = loTable.LowerBound To loTable.UpperBound
        loName = loTable.Items(counter)
        On Error Resume Next
        Set Lo = Wksh.ListObjects(loName)
        On Error GoTo 0
        If Lo Is Nothing Then ThrowError " Listobject " & loName & " does not exist in analysis worksheet "
    Next
End Sub

'Find if a column Exists in a listObject of the analysis
Private Function ColumnExists(ByVal Lo As listObject, ByVal colName As String) As Boolean
    Dim hRng As Range
    Set hRng = Lo.HeaderRowRange

    ColumnExists = (Not hRng.Find(What:=colName, lookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function

Private Function ColumnIndex(ByVal Lo As listObject, ByVal colName) As Long
    Dim Index As Long
    Dim hRng As Range

    Set hRng = Lo.HeaderRowRange
    If ColumnExists(Lo, colName) Then
        Index = hRng.Find(What:=colName, lookAt:=xlWhole, MatchCase:=False).Column - hRng.Column + 1
    Else
        Index = -1
    End If
    ColumnIndex = Index
End Function

Private Function DataRange(ByVal Lo As listObject, ByVal colName As String) As Range
    If ColumnExists(Lo, colName) Then Set DataRange = Lo.ListColumns(ColumnIndex(Lo, colName)).Range
End Function

Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

Private Sub Build(ByVal ll As Linelist)
    Dim anaout As IAnalysisOutput
    Dim sh As Worksheet
    Set sh = Wksh()
    Set anaout = AnalysisOutput.Create(sh, ll)
    anaout.WriteAnalysis
End Sub

'Import listobjects in analysis
Private Sub Import(ByVal impWksh As Worksheet)
    Dim anash As Worksheet
    Dim impRng As Range
    Dim actRng As Range
    Dim impAddress As String
    Dim impLo As listObject

    Set anash = Wksh()

    'Test if the sheet is empty
    If impWksh Is Nothing Then Exit Sub          'Add trace for import

    anash.Cells.Clear

    On Error GoTo LoFail
    Set impLo = impWksh.ListObjects(1)
    On Error GoTo 0


    For Each impLo In impWksh.ListObjects

        Set impRng = impLo.Range
        impAddress = impRng.Address
        Set actRng = anash.Range(impAddress)
        actRng.Value = impRng.Value

        actRng.Cells(-1, 1).Value = impRng.Cells(-1, 1).Value
        actRng.Cells(-1, 1).Font.Size = impRng.Cells(-1, 1).Font.Size
        actRng.Cells(-1, 1).Font.color = impRng.Cells(-1, 1).Font.color
        actRng.Cells(-1, 1).Font.Bold = True

        'Define the listobject
        With anash
            .ListObjects.Add(xlSrcRange, actRng, , xlYes).Name = impLo.Name
            .ListObjects(impLo.Name).TableStyle = impLo.TableStyle
        End With
    Next

    'Be sure all the tables exist after import
    CheckTablesRequirements

    anash.Cells(1, 1).Value = "analysis"
    anash.Cells(1, 1).Font.color = vbBlue

    'Sort the analysis worksheet on each one of the tables
    SortAnalysis

    Exit Sub
LoFail:
    'add trace for no list objects
End Sub


Private Sub SortListOnFirst(ByVal Lo As listObject, ByVal colIndex As Long)

    Dim numRng As Range 'Number range
    Dim colRng As Range 'Colum range to sort on
    Dim LoRng As Range 'ListObject Range
    Dim hRng As Range 'Header Range of the list Object
    Dim resRng As Range 'Will resize on this range
    Dim counter As Long
    Dim trackData As BetterArray
    Dim sh As Worksheet
    Dim sortedValue As String
    Dim num As Long
    Dim numData As BetterArray

    'Sort on values only If there is data
    If Lo.DataBodyRange Is Nothing Then Exit Sub
    If IsEmpty(Lo.DataBodyRange) Then Exit Sub

    Set trackData = New BetterArray
    Set numData = New BetterArray
    Set sh = Wksh()

    Set LoRng = Lo.DataBodyRange
    Set hRng = Lo.HeaderRowRange
    Set colRng = Lo.ListColumns(colIndex).DataBodyRange

    'Test values progressively
    'If the trackData does not contains the values,
    'Update the number (where you will sort)

    hRng.Cells(1, hRng.Cells.Count + 1).Value = "__number"

    With sh
        Set resRng = .Range(hRng.Cells(1, 1), LoRng.Cells(LoRng.Rows.Count, LoRng.Columns.Count + 1))
        Lo.Resize resRng
    End With

    Set numRng = Lo.ListColumns(hRng.Cells.Count + 1).DataBodyRange

    'Seems like sometimes the databodyrange of a new column could be empty in that case exit
    If numRng Is Nothing Then
        Lo.ListColumns(hRng.Cells.Count + 1).Delete
       Exit Sub
   End If

    num = 1 'Number to sort later on

    'Resize the listObject to be sure it includes the number added
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
     'Now sort the listObject on the column
     LoRng.Sort key1:=numRng, Order1:=xlAscending, Header:=xlYes
     'Remove the _number column
     Lo.ListColumns(hRng.Cells.Count + 1).Delete
End Sub

'Resize a listobject to resize every listObject after loading.
Private Sub ResizeLo(ByVal Lo As listObject, Optional ByVal totalRowCount As Long = 0)

    'Begining of the tables
    Dim loRowHeader As Long
    Dim loColHeader  As Long
    Dim rowCounter As Long

    'End of the listobject table
    Dim loRowsEnd As Long
    Dim loColsEnd As Long
    Dim sh As Worksheet

    Set sh = Wksh()

    'Remove unwanted rows from each listObject
    With sh
        'Rows and columns at the begining of the table to resize
        loRowHeader = Lo.Range.Row
        loColHeader = Lo.Range.Column

        'Rows and Columns at the end of the Table to resize
        loRowsEnd = loRowHeader + Lo.Range.Rows.Count - 1
        loColsEnd = loColHeader + Lo.Range.Columns.Count - 1
        rowCounter = loRowsEnd
        Do While (rowCounter > loRowHeader + 1)
            If (Application.WorksheetFunction.COUNTA(.Rows(rowCounter)) <= totalRowCount) Then
                .Rows(rowCounter).EntireRow.Delete
                'update the end rows
                loRowsEnd = loRowsEnd - 1
            End If
            rowCounter = rowCounter - 1
        Loop
            Lo.Resize .Range(.Cells(loRowHeader, loColHeader), .Cells(loRowsEnd, loColsEnd))
    End With
End Sub


Private Sub SortAnalysis()

    Dim sh As Worksheet
    Dim Lo As listObject
    Dim colIndex As Long
    Dim tabl As BetterArray 'Will contain the list of listObjects to sort
    Dim counter As Long
    Dim loName As String
    Dim rowCount As Long

    Set tabl = New BetterArray
    tabl.LowerBound = 1

    tabl.Push UNIVARIATE, BIVARIATE, TIMESERIES, SPATIAL, GRAPHTIMESERIES, TITLEGRAPHTIMESERIES

    Set sh = Wksh()

    For counter = 1 To tabl.Length

        loName = tabl.Items(counter)
        Set Lo = sh.ListObjects(loName)

        'Row count is the number of elements in the row that triggers deletion of the
        'entire row of the listObject
        rowCount = Switch(loName = TIMESERIES, 2, loName = GRAPHTIMESERIES, 4, True, 0)
        
        If (loName = GRAPHTIMESERIES Or loName = TITLEGRAPHTIMESERIES) Then 
            colIndex = ColumnIndex(Lo, "graph id")
        ElseIf (loName = TIMESERIES) Then
            colIndex = ColumnIndex(Lo, "table order")
        Else
            colIndex = ColumnIndex(Lo, "section")
        End If

        SortListOnFirst Lo, colIndex

        'Sort on table order for timeseries after sorting on table order
        If (loName = TIMESERIES) Then SortListOnFirst Lo, ColumnIndex(Lo, "section")

        'resize the listObject
        ResizeLo Lo, totalRowCount:=rowCount
    Next

End Sub

'Translate
Private Sub Translate(ByVal TransObject As ITranslation)
    Dim transCols As BetterArray
    Dim anash As Worksheet
    Dim counter As Long
    Dim Lo As listObject
    Dim colRng As Range

    'TransObject is the translation object (from Translation.cls). It can
    'translate ranges, and formula ranges.

    Set transCols = New BetterArray
    Set anash = Wksh()
    transCols.Push "section", "table title", "graph title", "summary label", _
                   "choices", "label", "section (select)"
 
    'choices is for the time series graphs translation

    For Each Lo In anash.ListObjects
        For counter = transCols.LowerBound To transCols.UpperBound
            If ColumnExists(Lo, transCols.Item(counter)) Then
                Set colRng = DataRange(Lo, transCols.Item(counter))
                TransObject.TranslateRange colRng
            End If
        Next

        'Formulas
        If ColumnExists(Lo, "summary function") Then
            Set colRng = DataRange(Lo, "summary function")
            TransObject.TranslateRange colRng, containsFormula:=True
        End If
    Next
End Sub

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, errorMessage
End Sub

'Interface
Public Sub ILLAnalysis_Import(ByVal impWksh As Worksheet)
    Import impWksh
End Sub

Public Sub ILLAnalysis_Translate(ByVal TransObject As ITranslation)
    Translate TransObject
End Sub

Public Property Get ILLAnalysis_Wksh() As Worksheet
    Set ILLAnalysis_Wksh = Wksh()
End Property

Public Sub ILLAnalysis_Build(ll As ILinelist)
    Build ll
End Sub
