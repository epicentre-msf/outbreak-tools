VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnalysisTableIterationItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Describes the result of enumerating a single analysis table specification")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Return the raw TableSpecs reference created for the current row.
Public Property Get Specification() As ITablesSpecs
End Property

'Indicates whether the current row begins a new analysis section.
Public Property Get IsNewSection() As Boolean
End Property

'Expose the worksheet range containing column headers for the specification list.
Public Property Get HeaderRange() As Range
End Property

'Expose the worksheet range tied to the current specification row.
Public Property Get RowRange() As Range
End Property

'Expose the list row reference used to generate the entry.
Public Property Get ListRow() As ListRow
End Property

'Return the specification that precedes the current row, when available.
Public Property Get PreviousSpecification() As ITablesSpecs
End Property

'Return the specification that follows the current row, when available.
Public Property Get NextSpecification() As ITablesSpecs
End Property

