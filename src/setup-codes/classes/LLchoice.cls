VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLchoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist choice class"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Linelist choice class")

Option Explicit
Implements ILLchoice


Private Type Tchoice
    choiceDataSheet As IDataSheet
End Type

Private this As Tchoice

Private Const LISTNAME As String = "List Name"
Private Const LABEL   As String = "Label"
Private Const LABELSHORT As String = "Short Label"
Private Const ORDER   As String = "Ordering List"

'Instanciation
Public Function Create(ByVal choiceWorksheet As Worksheet, ByVal choiceStartRow As Long, ByVal choiceStartColumn As Long) As ILLchoice

    Dim customDataSheet As IDataSheet
    Set customDataSheet = DataSheet.Create(choiceWorksheet, choiceStartRow, choiceStartColumn)

    With New LLchoice
        Set .ChoiceData = customDataSheet
        Set Create = .Self
    End With
End Function

'Self
Public Property Get Self() As LLchoice
    Set Self = Me
End Property

'Worksheet
Public Property Get Wksh() As Worksheet
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    Set Wksh = dataObject.Wksh()
End Property

'StartRow
Private Property Get StartRow() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    StartRow = dataObject.StartRow()
End Property

'StartColumn
Private Property Get StartColumn() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    StartColumn = dataObject.StartColumn()
End Property

Private Property Get ChoiceData() As IDataSheet
    Set ChoiceData = this.choiceDataSheet
End Property

Public Property Set ChoiceData(ByVal customDataSheet As IDataSheet)
    Set this.choiceDataSheet = customDataSheet
End Property

'Get the data range
Private Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = True) As Range
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData

    Set DataRange = dataObject.DataRange(colName, includeHeaders)
End Property

'Test if a choice Exists in the choice sheet
Private Property Get ChoiceExists(ByVal choiceName As String) As Boolean
    Dim listRng As Range
    Dim exist As Boolean
    Dim dataObject As IDataSheet

    exist = False

    Set dataObject = ChoiceData
    If dataObject.ColumnExists(LISTNAME) Then
        Set listRng = DataRange(LISTNAME, includeHeaders:=False)
        exist = Not (listRng.Find(What:=choiceName, lookAt:=xlWhole, _
                                  MatchCase:=True, After:=listRng.Cells(listRng.Rows.Count, 1)) Is Nothing)
    End If
    ChoiceExists = exist
End Property

'Get the categories
Private Property Get Categories(ByVal choiceName As String) As BetterArray
    Dim counter As Long
    Dim rngName As Range
    Dim rngLabel As Range
    Dim lst As BetterArray
    Set lst = New BetterArray

    If ChoiceExists(choiceName) Then
        Set rngName = DataRange(LISTNAME, includeHeaders:=False)
        Set rngLabel = DataRange(LABEL, includeHeaders:=False)
        For counter = 1 To rngName.Rows.Count
            If rngName.Cells(counter, 1).Value = choiceName Then _
             lst.Push rngLabel.Cells(counter, 1).Value
        Next
    End If
    Set Categories = lst.Clone()
End Property

'Sort the Choices sheet
Private Sub Sort()

    Dim dataObject As IDataSheet
    Dim orderRng As Range
    Dim listRng As Range
    Dim choiceRng As Range

    Set dataObject = ChoiceData

    If dataObject.ColumnExists(LISTNAME) And dataObject.ColumnExists(ORDER) Then
        Set listRng = DataRange(LISTNAME, includeHeaders:=False)
        Set orderRng = DataRange(ORDER, includeHeaders:=False)
        Set choiceRng = DataRange(includeHeaders:=False)

        choiceRng.Sort key1:=listRng, order1:=xlAscending, key2:=orderRng, order2:=xlAscending
    End If
End Sub

Private Property Get AllChoices() As BetterArray
    Dim lstChoi As BetterArray
    Dim lstRng As Range
    Dim counter As Long 'counter for the number of rows of one range

    Set lstChoi = New BetterArray
    Set lstRng = DataRange(LISTNAME, includeHeaders:=False)
    For counter = 1 To lstRng.Rows.Count
        If Not lstChoi.Includes(lstRng.Cells(counter, 1).Value) Then _
            lstChoi.Push lstRng.Cells(counter, 1).Value
    Next
    
    Set AllChoices = lstChoi.Clone()
End Property

'Add New Categories
Private Sub AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray)

    Dim listRng As Range
    Dim labelRng As Range
    Dim orderRng As Range
    Dim choicewksh As Worksheet
    Dim EndRow As Long
    Dim endRowlabel As Long
    Dim counter As Long
    Dim checkCategories As BetterArray

    If CategoriesData.Length = 0 Then Exit Sub

    'Test if the choice already exists
    Set checkCategories = Categories(choiceName)
    If checkCategories.Length > 0 Then Exit Sub  'Will probably trace log.

    Set listRng = DataRange(LISTNAME)
    Set labelRng = DataRange(LABEL)
    Set orderRng = DataRange(ORDER)
    Set choicewksh = Wksh()

    EndRow = StartRow + listRng.Rows.Count
    endRowlabel = StartRow + labelRng.Rows.Count

    With choicewksh
        If (EndRow = endRowlabel) Then
            CategoriesData.ToExcelRange .Cells(EndRow, labelRng.Column)
            For counter = 0 To (CategoriesData.Length - 1)
                .Cells(EndRow + counter, listRng.Column).Value = choiceName
                .Cells(EndRow + counter, orderRng.Column).Value = counter + 1
            Next
        End If
    End With
End Sub

'Add / Remove Rows
Private Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim tabl As ICustomTable

    Set sh = Wksh()

    On Error Resume Next
    Set Lo = sh.ListObjects(1)
    On Error GoTo 0
    If Lo Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(Lo)
    If del Then
        tabl.RemoveRows
    Else
        tabl.AddRows
    End If
End Sub

'Properties
Private Property Get ILLchoice_Wksh() As Worksheet
    Set ILLchoice_Wksh = Wksh()
End Property

Private Property Get ILLchoice_StartRow() As Long
    ILLchoice_StartRow = StartRow()
End Property

Private Property Get ILLchoice_StartColumn() As Long
    ILLchoice_StartColumn = StartColumn()
End Property

Private Property Get ILLchoice_ChoiceData() As IDataSheet
    Set ILLchoice_ChoiceData = ChoiceData
End Property

Private Property Get ILLchoice_DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = True) As Range
    Set ILLchoice_DataRange = DataRange(colName, includeHeaders)
End Property

Private Property Get ILLchoice_Categories(ByVal choiceName As String) As BetterArray
    Set ILLchoice_Categories = Categories(choiceName)
End Property

'@Ignore ProcedureNotUsed
Private Property Get ILLChoice_ChoiceExists(ByVal choiceName As String) As Boolean
    ILLChoice_ChoiceExists = ChoiceExists(choiceName)
End Property

'Void methods
Private Sub ILLchoice_Sort()
    Sort
End Sub

Private Sub ILLchoice_AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray)
    AddChoice choiceName:=choiceName, CategoriesData:=CategoriesData
End Sub

Private Sub ILLchoice_AddRows()
    ManageRows
End Sub

Private Sub ILLchoice_RemoveRows()
    ManageRows del:=True
End Sub

Private Property Get ILLchoice_AllChoices() As BetterArray
    Set ILLchoice_AllChoices = AllChoices()
End Property

