VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLchoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Linelist choice Class")
'@IgnoreModule

Option Explicit
Implements ILLchoice


Private Type Tchoice
    choiceDataSheet As IDataSheet
End Type

Private this As Tchoice

Private Const ListName As String = "list name"
Private Const LABEL   As String = "label"
Private Const ORDER   As String = "ordering list"

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
Public Property Get StartRow() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    StartRow = dataObject.StartRow()
End Property

'StartColumn
Public Property Get StartColumn() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    StartColumn = dataObject.StartColumn()
End Property

Private Property Get ChoiceData() As IDataSheet
    Set ChoiceData = this.choiceDataSheet
End Property

Friend Property Set ChoiceData(ByVal customDataSheet As IDataSheet)
    Set this.choiceDataSheet = customDataSheet
End Property

'Get the data range
Public Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders = True) As Range
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData

    Set DataRange = dataObject.DataRange(colName, includeHeaders)
End Property

'Get the categories
Public Property Get Categories(ByVal choiceName As String) As BetterArray
    Dim dataObject As IDataSheet

    Set dataObject = ChoiceData

    Set Categories = dataObject.FilterData(varName:=ListName, _
                                           criteriaName:=choiceName, returnedColumnName:=LABEL, _
                                           includeHeaders:=False)
End Property

'Sort the Choices sheet
Private Sub Sort()

    Dim dataObject As IDataSheet
    Dim orderRng As Range
    Dim listRng As Range
    Dim choiceRng As Range

    Set dataObject = ChoiceData

    If dataObject.ColumnExists(ListName) And dataObject.ColumnExists(ORDER) Then
        Set listRng = DataRange(ListName, includeHeaders:=False)
        Set orderRng = DataRange(ORDER, includeHeaders:=False)
        Set choiceRng = DataRange(includeHeaders:=False)

        choiceRng.Sort key1:=listRng, order1:=xlAscending, key2:=orderRng, order2:=xlAscending
    End If
End Sub

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
    Set checkCategories = New BetterArray
    Set checkCategories = Categories(choiceName)
    If checkCategories.Length > 0 Then Exit Sub  'Will probably trace log.


    Set listRng = DataRange(ListName)
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
        Else
            'Do something (trace log)

        End If
    End With
End Sub

'Import and export Sub
Private Sub Export(wkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Dim choiData As IDataSheet
    Set choiData = ChoiceData()
    ChoiceData.Export wkb, Hide:=Hide
End Sub

Private Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Dim choiData As IDataSheet
    Set choiData = ChoiceData()
    choiData.Import fromWksh, fromStartRow, fromStartcol
End Sub

'Translation
Private Sub Translate(ByVal TransObject As ITranslation)
    Dim lblRng As Range
    Set lblRng = DataRange("label")
    TransObject.TranslateRange lblRng
End Sub

'Properties
Public Property Get ILLchoice_Wksh() As Worksheet
    Set ILLchoice_Wksh = Wksh()
End Property

Public Property Get ILLchoice_StartRow() As Long
    ILLchoice_StartRow = StartRow()
End Property

Public Property Get ILLchoice_StartColumn() As Long
    ILLchoice_StartColumn = StartColumn()
End Property

Public Property Get ILLchoice_ChoiceData() As IDataSheet
    Set ILLchoice_ChoiceData = ChoiceData
End Property

Public Property Get ILLchoice_DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders = True) As Range
    Set ILLchoice_DataRange = DataRange(colName, includeHeaders)
End Property

Public Property Get ILLchoice_Categories(ByVal choiceName As String) As BetterArray
    Set ILLchoice_Categories = Categories(choiceName)
End Property

'Void methods
Public Sub ILLchoice_Sort()
    Sort
End Sub

Public Sub ILLchoice_AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray)
    AddChoice choiceName:=choiceName, CategoriesData:=CategoriesData
End Sub

'Export and Import
Public Sub ILLchoice_Export(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, Hide:=Hide
End Sub

Public Sub ILLchoice_Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Import fromWksh, fromStartRow, fromStartcol
End Sub

Public Sub ILLchoice_Translate(ByVal TransObject As ITranslation)
    Translate TransObject
End Sub


