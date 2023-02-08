VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLchoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of Choice class")
'@IgnoreModule

Option Explicit

'Properties
Public Property Get Wksh() As Worksheet: End Property
Public Property Get StartRow() As Long: End Property

Public Property Get StartColumn() As Long: End Property
Public Property Get ChoiceData() As IDataSheet: End Property

Public Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders = True) As Range: End Property
Public Property Get Categories(ByVal choiceName As String) As BetterArray: End Property

'Exposed methods

'Void methods
'Sort the choices sheet
Public Sub Sort(): End Sub

'Add choice
Public Sub AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray): End Sub

'Import and exports
Public Sub Export(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
End Sub

Public Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
End Sub

Public Sub Translate(ByVal TransObject As ITranslation)
End Sub

