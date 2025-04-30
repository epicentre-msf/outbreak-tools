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
Public Property Get HasCheckings() As Boolean: End Property

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



Public Property Get Categories( _ 
  ByVal choiceName As String, _
  Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Property




'Exposed methods
Public Property Get TraceValues() As Object: End Property




'Import and exports
Public Sub Export(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
End Sub

Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
End Sub

Public Sub Translate(ByVal TransObject As ITranslation): End Sub


'Add and remove rows from the choice listObject
Public Sub AddRows(): End Sub
Public Sub RemoveRows(): End Sub