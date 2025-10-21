VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListBuildContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist.Shared")
'@ModuleDescription("Provides high-level metadata required when building linelist worksheets")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Metadata
'===============================================================================
'Name of the worksheet currently being processed.
Public Property Get WorksheetName() As String
End Property

'Table-type discriminator used to select layout and business rules.
Public Property Get TableTypeCode() As Byte
End Property

'Optional hook exposing the underlying linelist object when additional context is required.
Public Property Get Linelist() As Object
End Property

