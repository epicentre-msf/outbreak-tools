VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListContextCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Linelist.Shared")
'@ModuleDescription("Caches frequently accessed linelist collaborators (dictionary, variables, linelist)")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@section Attachment
'===============================================================================
Public Sub AttachLinelist(ByVal linelist As Object)
End Sub

Public Sub AttachDictionary(ByVal dictionary As Object)
End Sub

Public Sub AttachVariables(ByVal variables As Object)
End Sub

'@section Accessors
'===============================================================================
Public Function Linelist() As Object
End Function

Public Function Dictionary() As Object
End Function

Public Function Variables() As Object
End Function

'@section Maintenance
'===============================================================================
Public Sub Invalidate()
End Sub

