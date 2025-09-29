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

'@jump:RemoveRows
'@sub-title Remove rows from analysis tables while preserving required rows
Public Sub RemoveRows(): End Sub


'@jump:outside
'@section Operate with outsilde world
'===============================================================================
'@jump:Import
'@sub-title Import analysis listobjects from another worksheet
'@param fromWksh Worksheet providing the source analysis tables
Public Sub Import(ByVal fromWksh As Worksheet): End Sub

'@jump:Export
'@sub-title Export analysis worksheet to another workbook
'@param toWkb Workbook receiving the exported worksheet
Public Sub Export(ByVal toWkb As Workbook): End Sub

'@jump:Translate
'@sub-title Translate analysis jumps using a translation object
'@param TransObject ITranslationObject used for translation
Public Sub Translate(ByVal TransObject As ITranslationObject): End Sub

'@jump:Sort
'@sub-title Sort the analysis worksheet
Public Sub Sort(): End Sub

'@jump:checkings
'@section Checkings
'===============================================================================

'@jump:HasCheckings
'@prop-title Indicates whether trace information is available
Public Property Get HasCheckings() As Boolean: End Property

'@jump:CheckingValues
'@prop-title Retrieve the collected trace information
Public Property Get CheckingValues() As Object: End Property
