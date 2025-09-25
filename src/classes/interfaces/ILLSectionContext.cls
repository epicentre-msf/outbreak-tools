VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSectionContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Shared dictionary context for linelist sections and subsections")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'@InterfaceDescription("Provides cached metadata and range helpers for section writers")

Option Explicit

'@section Lifecycle
'===============================================================================
Public Sub Initialise(ByVal ll As ILinelist, ByVal startRow As Long)
End Sub

Public Sub ResetCaches()
End Sub

Public Property Get Linelist() As ILinelist
End Property

Public Property Get StartRow() As Long
End Property

Public Property Let StartRow(ByVal newValue As Long)
End Property

'@section Dictionary helpers
'===============================================================================
Public Function Dictionary() As ILLdictionary
End Function

Public Function Specs() As ILinelistSpecs
End Function

Public Function Design() As ILLFormat
End Function

Public Function Translation() As ITranslationObject
End Function

Public Function DataEndRow() As Long
End Function

'@section Lookups
'===============================================================================
Public Function Value(ByVal columnName As String, _
                      Optional ByVal rowOffset As Long = 0) As String
End Function

Public Function EndRow(ByVal columnNames As Variant) As Long
End Function

Public Function HasNext(ByVal columnNames As Variant) As Boolean
End Function

Public Function NextStartRow(ByVal columnNames As Variant) As Long
End Function

Public Function VariableNames(ByVal firstRow As Long, _
                              ByVal lastRow As Long) As BetterArray
End Function

