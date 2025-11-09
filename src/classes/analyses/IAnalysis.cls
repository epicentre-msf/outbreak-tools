VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Core analysis management capabilities"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Core analysis management capabilities")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation

Option Explicit

'@jump:elements
'@section Elements
'===============================================================================

'@jump:wksh
'@prop-title Worksheet hosting the analysis tables
'@return Worksheet
Public Property Get Wksh() As Worksheet: End Property

'@jump:rowsmanagements
'@section Row management
'===============================================================================

'@jump:AddRows
'@sub-title Add rows to analysis tables based on the selection header
Public Sub AddRows(): End Sub

'@jump:InsertRows
'@sub-title Insert rows based on a worksheet selection
'@param targetCell Range anchor describing the rows to mirror.
'@param insertShift Optional Boolean. When True worksheet rows are inserted to keep stacked tables aligned.
Public Sub InsertRows(ByVal targetCell As Range, _
                      Optional ByVal insertShift As Boolean = True)
End Sub

'@jump:DeleteRows
Public Sub DeleteRows(ByVal targetCell As Range, _
                      Optional ByVal includeIds As Boolean = True, _
                      Optional ByVal forceShift As Boolean = True)
End Sub

'@jump:RemoveRows
'@sub-title Remove rows from analysis tables while preserving required rows
Public Sub RemoveRows(): End Sub

Public Sub ManageRows(Optional ByVal del As Boolean = False): End Sub
'@jump:outside
'@section Operate with outsilde world
'===============================================================================
'@jump:Import
'@sub-title Import analysis listobjects from another worksheet
'@param sourchesheet Worksheet providing the source analysis tables
Public Sub Import(ByVal sourchesheet As Worksheet): End Sub

'@jump:Export
'@sub-title Export analysis worksheet to another workbook
'@param sourcewb Workbook receiving the exported worksheet
Public Sub Export(ByVal sourcewb As Workbook, Optional ByVal Hide As Long = xlSheetHidden): End Sub

'@jump:Translate
'@sub-title Translate analysis jumps using a translation object
'@param TransObject ITranslationObject used for translation
Public Sub Translate(ByVal translator As ITranslationObject): End Sub

'@jump:Sort
'@sub-title Sort the analysis worksheet
Public Sub Sort(): End Sub

'@jump:Clean
'@sub-title Clean the analysis worksheet
Public Sub Clean(): End Sub

'@jump:EnsureAnalysisTables
'@sub-title Make sure analysis tables are present; raise error if not
Public Sub EnsureAnalysisTables(): End Sub

'@jump:checkings
'@section Checkings
'===============================================================================

'@jump:HasCheckings
'@prop-title Indicates whether trace information is available
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'@prop-title Retrieve the collected trace information
Public Property Get CheckingValues() As Object: End Property
