VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of Choice class")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'@section Translations of headers
Public Property Get HasTranslation() As Boolean: End Property
Public Property Get TradObject() As ITranslation: End Property

'@section Associated DataSheet Object properties
Public Property Get Wksh() As Worksheet: End Property
Public Property Get StartRow() As Long: End Property
Public Property Get StartColumn() As Long: End Property
Public Property Get ChoiceData() As IDataSheet: End Property
Public Property Get DataRange( _ 
          Optional ByVal colName As String = "__all__", _
          Optional ByVal includeHeaders = True) As Range
End Property

'@section Working with Choices

Public Property Get ChoiceExists(ByVal choiceName As String) As Boolean
End Property

'Sort the choices in the choice worksheet
Public Sub Sort(): End Sub

'Add choice
Public Sub AddChoice(ByVal choiceName As String, _ 
                     ByVal longLabelData As BetterArray, _
                     Optional ByVal shortLabelData As Object = Nothing)
End Sub

'Remove choice
Public Sub RemoveChoice(ByVal choiceNames As String, Optional ByVal sep As String = ";")
End Sub

'Remove Rows in the choice worksheet if there is a custom table attached
Public Sub AddRows(Optional ByVal nbRows As Long = 5, Optional ByVal insertShift As Boolean = False)
End Sub

Public Sub RemoveRows(): End Sub

Public Sub Translate(ByVal trads As ITranslation): End Sub

'Retrieve all the choices in the choice object
Public Property Get AllChoices() As BetterArray: End Property

'@section Working with categories

Public Property Get Categories( _ 
  ByVal choiceName As String, _
  Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Property

Public Property Get ConcatenateCategories( _ 
        ByVal choiceName As String, _ 
        Optional ByVal separator As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String
End Property

'@section Input/output (Import/Export)

'Import and exports
'Export the choice object to a workbook
Public Sub Export(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
End Sub

'Import a choice object to a new workbook
Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, _ 
                  ByVal fromStartcol As Long)
End Sub

'@section Checkings

'Check if the current choice has some checkings
Public Property Get HasCheckings() As Boolean: End Property

'Errors and checkings
Public Property Get CheckingValues() As Object: End Property
