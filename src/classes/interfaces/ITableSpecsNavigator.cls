VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableSpecsNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Navigation helper for TableSpecs rows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Factory
'===============================================================================
'Initialise the navigator with the active table context.
Public Function Create(ByVal headerRange As Range, _
                       ByVal currentRange As Range, _
                       ByVal linelistSpecs As ILinelistSpecs, _
                       ByVal columnMap As ITableSpecsColumnMap, _
                       ByVal builder As ITableSpecsBuilder, _
                       ByVal tableTypeCode As Byte) As ITableSpecsNavigator
End Function

'@section Navigation
'===============================================================================
'Return the previous valid table specification.
Public Function PreviousSpec() As ITablesSpecs
End Function

'Return the next valid table specification limited by the provided anchor range.
Public Function NextSpec(ByVal anchorRange As Range) As ITablesSpecs
End Function

'Determine whether the current row starts a new section.
Public Function IsNewSection() As Boolean
End Function

'@section Cache management
'===============================================================================
'Invalidate any cached column map data so subsequent operations refresh values.
Public Sub Invalidate()
End Sub

