VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListSectionDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Concrete IListSectionDescriptor storing section identity and anchors")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IListSectionDescriptor

Private Type TSectionState
    sectionName As String
    anchorRow As Long
    anchorColumn As Long
End Type

Private this As TSectionState
Private Const CLASSNAME As String = "ListSectionDescriptor"

'@section Factory
'===============================================================================
'Create a descriptor with the supplied identity and anchor coordinates.
Public Function Create(ByVal sectionName As String, _
                       ByVal anchorRow As Long, _
                       ByVal anchorColumn As Long) As IListSectionDescriptor

    ValidateSectionName sectionName
    ValidateAnchor anchorRow, "anchorRow"
    ValidateAnchor anchorColumn, "anchorColumn"

    With New ListSectionDescriptor
        .Initialise sectionName, anchorRow, anchorColumn
        Set Create = .Self
    End With
End Function

'Expose the instance as its interface type.
Public Property Get Self() As IListSectionDescriptor
    Set Self = Me
End Property

'Initialise instance state; internal helper used by the factory.
Public Sub Initialise(ByVal sectionName As String, _
                      ByVal anchorRow As Long, _
                      ByVal anchorColumn As Long)

    this.sectionName = Trim$(sectionName)
    this.anchorRow = anchorRow
    this.anchorColumn = anchorColumn
End Sub

'@section Validation helpers
'===============================================================================
Private Sub ValidateSectionName(ByVal sectionName As String)
    If Trim$(sectionName) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Section name cannot be empty when building descriptors"
    End If
End Sub

Private Sub ValidateAnchor(ByVal anchorValue As Long, ByVal argumentName As String)
    If anchorValue <= 0 Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  argumentName & " must be a positive value"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IListSectionDescriptor_SectionName() As String
    IListSectionDescriptor_SectionName = this.sectionName
End Property

Private Property Get IListSectionDescriptor_AnchorRow() As Long
    IListSectionDescriptor_AnchorRow = this.anchorRow
End Property

Private Property Get IListSectionDescriptor_AnchorColumn() As Long
    IListSectionDescriptor_AnchorColumn = this.anchorColumn
End Property

