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
'@Folder("GeneralInterface")
'@ModuleDescription("Interface of Choice class")
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@jump:Create

'@section Translations of headers

'@jump:HasTranslation
'Check if the choices have translation of headers
Public Property Get HasTranslation() As Boolean: End Property

'@jump:TradObject
'Translation object used for translating the headers
Public Property Get TradObject() As ITranslation: End Property

'@section Associated DataSheet Object properties

'@jump:ChoiceData
'Associated DataSheet Object
Public Property Get ChoiceData() As IDataSheet: End Property

'@jump:Wksh
'Worksheet of the choices
Public Property Get Wksh() As Worksheet: End Property

'@section Working with Choices

'@jump:Name
'Name of the choice object
Public Property Get Name() As String : End Property

'@jump:ChoiceExists
'Test if a choice Exists
Public Property Get ChoiceExists(ByVal choiceName As String) As Boolean: End Property

'@jump:Sort
'Sort the choices in the choice worksheet
Public Sub Sort(): End Sub

'@jump:AddChoice
'Add new choices to the choice worksheet
Public Sub AddChoice(ByVal choiceName As String, _ 
                     ByVal longLabelData As BetterArray, _
                     Optional ByVal shortLabelData As Object = Nothing)
End Sub

'@jump:RemoveChoice
'Remove one or multiple choices at once, separated by a separator
Public Sub RemoveChoice(ByVal choiceNames As String, Optional ByVal sep As String = ";")
End Sub

'@jump:Translate
'Translate the current choice worksheet
Public Sub Translate(ByVal trads As ITranslation): End Sub

'@jump:AddRows
'Add Rows in the choice associated custom table
Public Sub AddRows(Optional ByVal nbRows As Long = 5, _ 
                   Optional ByVal insertShift As Boolean = False)
End Sub

'@jump:RemoveRows
'Remove Rows in the choice associated custom table
Public Sub RemoveRows(): End Sub

'@jump:AllChoices
'Get all the choices in a BetterArray
Public Property Get AllChoices() As BetterArray: End Property

'@section Working with categories

'@jump:Categories
'Get the categories of one choice
Public Property Get Categories( _ 
  ByVal choiceName As String, _
  Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Property

'@jump:ConcatenateCategories
'Concatenate all categories of one choice
Public Property Get ConcatenateCategories( _ 
        ByVal choiceName As String, _ 
        Optional ByVal sep As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String
End Property

'@section Export and Import

'@jump:AddHeaders
'Add the headers to the choice object
Public Sub AddHeaders() : End Sub

'@jump:Export
'Export the choice object to a workbook
Public Sub Export(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
End Sub

'@jump:Import
'Import a choice object to a new workbook
Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, _ 
                  ByVal fromStartcol As Long)
End Sub

'@jump:ImportChoice
'Import from a choice object
Public Sub ImportChoice(ByVal choiObj As ILLChoices): End Sub

'@section Checkings

'@jump:HasCheckings
'Check if the current choice has some checkings
Public Property Get HasCheckings() As Boolean: End Property

'Errors and checkings
Public Property Get CheckingValues() As Object: End Property
