VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimeSeriesTableSpecsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis.Graphs")
'@ModuleDescription("Factory abstraction creating ITablesSpecs instances for time-series graph rows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Create a specification object for the supplied time-series list row.
Public Function Create(ByVal headerRange As Range, _
                       ByVal listRow As ListRow, _
                       ByVal linelist As ILinelistSpecs) As ITablesSpecs
End Function

'Expose the instance for fluent usage.
Public Property Get Self() As ITimeSeriesTableSpecsFactory
End Property

