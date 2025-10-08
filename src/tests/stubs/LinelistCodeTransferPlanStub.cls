VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistCodeTransferPlanStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Configurable ILinelistCodeTransferPlan stub for verifying transfer orchestration")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ILinelistCodeTransferPlan

Private Type TPlan
    classModules As BetterArray
    standardModules As BetterArray
    forms As BetterArray
    workbookModuleName As String
End Type

Private this As TPlan

Public Sub Initialise()
    Set this.classModules = New BetterArray
    this.classModules.LowerBound = 1

    Set this.standardModules = New BetterArray
    this.standardModules.LowerBound = 1

    Set this.forms = New BetterArray
    this.forms.LowerBound = 1

    this.workbookModuleName = vbNullString
End Sub

Public Sub SetClassModules(ParamArray names() As Variant)
    PopulateArray this.classModules, names
End Sub

Public Sub SetStandardModules(ParamArray names() As Variant)
    PopulateArray this.standardModules, names
End Sub

Public Sub SetForms(ParamArray names() As Variant)
    PopulateArray this.forms, names
End Sub

Public Property Let WorkbookModuleName(ByVal value As String)
    this.workbookModuleName = value
End Property

Private Sub PopulateArray(ByVal target As BetterArray, ByVal values() As Variant)
    Dim index As Long
    Dim lower As Long
    Dim upper As Long

    target.Clear
    target.LowerBound = 1

    If Not HasValues(values, lower, upper) Then Exit Sub

    For index = lower To upper
        target.Push CStr(values(index))
    Next index
End Sub

Private Function HasValues(ByRef values() As Variant, ByRef lower As Long, ByRef upper As Long) As Boolean
    On Error GoTo EmptyArray
        lower = LBound(values)
        upper = UBound(values)
        HasValues = (upper >= lower)
        Exit Function
EmptyArray:
    Err.Clear
    lower = 0
    upper = -1
    HasValues = False
End Function

Private Function ILinelistCodeTransferPlan_ClassComponentNames() As BetterArray
    Set ILinelistCodeTransferPlan_ClassComponentNames = this.classModules
End Function

Private Function ILinelistCodeTransferPlan_ModuleComponentNames() As BetterArray
    Set ILinelistCodeTransferPlan_ModuleComponentNames = this.standardModules
End Function

Private Function ILinelistCodeTransferPlan_FormNames() As BetterArray
    Set ILinelistCodeTransferPlan_FormNames = this.forms
End Function

Private Property Get ILinelistCodeTransferPlan_WorkbookModuleName() As String
    ILinelistCodeTransferPlan_WorkbookModuleName = this.workbookModuleName
End Property
