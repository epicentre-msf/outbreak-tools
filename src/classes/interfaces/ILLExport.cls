VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of export object"


'@Interface
'@Folder("Linelist")
'@ModuleDescription("Interface of export object")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

Public Enum ExportScope
  ExportAna = 998
  ExportAll = 999
End Enum

'Exposed methods
'Get the value of one column for one expor type
Public Property Get Value(ByVal colName As String, ByVal exportNumber As Integer) As String
End Property

'Get the total number of exports
Public Property Get NumberOfExports() As Long
End Property


'In the linelist, get the fileName of an export
Public Property Get ExportFileName(ByVal scope As Byte, _
                                   ByVal dict As ILLdictionary, _
                                  ByVal pass As ILLPasswords) As String
End Property

'Test if an export is active
Public Property Get IsActive(ByVal exportNumber As Integer) As Boolean
End Property

Public Sub Translate(ByVal trans As ITranslation): End Sub

'Import/export specifications from the export worksheet to / in another worksheet
Public Sub ExportSpecs(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden): End Sub
Public Sub ImportSpecs(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long): End Sub
