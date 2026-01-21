VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DesignerPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Prepares designer workbooks and ribbon settings"

'@PredeclaredId
'@Folder("Designer")
'@ModuleDescription("Prepares designer workbooks and persists ribbon settings through hidden names.")
'@depends ProjectError, HiddenNames, IHiddenNames, IDevelopment
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ParameterNotUsed

Option Explicit
Implements IDesignerPreparation

'@section State
'===============================================================================
Private Type TDesignerPreparation
    hostBook As Workbook       'Designer workbook being prepared
    store As IHiddenNames      'Workbook-level hidden names manager
End Type

Private this As TDesignerPreparation

'@section Constants
'===============================================================================
Private Const CLASS_NAME As String = "DesignerPreparation"
Private Const FLAG_ALERT As String = "chkAlert"
Private Const FLAG_INSTRUCT As String = "chkInstruct"
Private Const FLAG_ON As String = "Yes"
Private Const FLAG_OFF As String = "No"


'@section Factory
'===============================================================================
'@Description("Create and configure a designer preparation helper.")
Public Function Create(ByVal hostBook As Workbook) As IDesignerPreparation
    Dim instance As DesignerPreparation

    Set instance = New DesignerPreparation
    instance.Configure hostBook
    Set Create = instance.Self
End Function

'@Description("Expose the current instance as an interface reference.")
Public Property Get Self() As IDesignerPreparation
    Set Self = Me
End Property


'@section Public API
'===============================================================================
'@Description("Bind the helper to the designer workbook.")
Public Sub Configure(ByVal hostBook As Workbook)
    If hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook is required."
    End If

    Set this.hostBook = hostBook
    Set this.store = Nothing
End Sub

'@Description("Prepare workbook-level settings and ensure ribbon flags exist.")
Public Sub Prepare(ByRef devManager As IDevelopment)
    EnsureHiddenStore
    EnsureDefaultFlags

    'Keep the parameter consumed to satisfy interface parity with setup preparation.
    If Not devManager Is Nothing Then
        'No designer-specific dev work required yet.
    End If
End Sub

'@Description("Persist a ribbon checkbox flag as Yes/No in hidden names.")
Public Sub SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
    EnsureHiddenStore
    EnsureDefaultFlags
    this.store.SetValue nameId, IIf(enabled, FLAG_ON, FLAG_OFF)
End Sub

'@Description("Read a ribbon checkbox flag as a boolean from hidden names.")
Public Function GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim rawValue As String
    EnsureHiddenStore
    EnsureDefaultFlags

    rawValue = this.store.ValueAsString(nameId, IIf(defaultValue, FLAG_ON, FLAG_OFF))
    GetFlag = (StrComp(rawValue, FLAG_ON, vbTextCompare) = 0)
End Function

'@Description("Expose the configured host workbook.")
Public Property Get HostWorkbook() As Workbook
    If this.hostBook Is Nothing Then
        RaiseError ProjectError.ObjectNotInitialized, "Host workbook was not configured."
    End If
    Set HostWorkbook = this.hostBook
End Property

'@Description("Expose the workbook-level hidden names manager.")
Public Property Get HiddenStore() As IHiddenNames
    EnsureHiddenStore
    Set HiddenStore = this.store
End Property


'@section Internal helpers
'===============================================================================
Private Sub EnsureHiddenStore()
    If this.store Is Nothing Then
        Set this.store = HiddenNames.Create(HostWorkbook)
    End If
End Sub

Private Sub EnsureDefaultFlags()
    this.store.EnsureName FLAG_ALERT, FLAG_OFF, HiddenNameTypeString
    this.store.EnsureName FLAG_INSTRUCT, FLAG_OFF, HiddenNameTypeString
End Sub

Private Sub RaiseError(ByVal errorCode As ProjectError, ByVal messageText As String)
    Err.Raise CLng(errorCode), CLASS_NAME, messageText
End Sub


'@section Interface implementation
'===============================================================================
Private Sub IDesignerPreparation_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub

Private Function IDesignerPreparation_GetFlag(ByVal nameId As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    IDesignerPreparation_GetFlag = GetFlag(nameId, defaultValue)
End Function

Private Property Get IDesignerPreparation_HiddenStore() As IHiddenNames
    Set IDesignerPreparation_HiddenStore = HiddenStore
End Property

Private Property Get IDesignerPreparation_HostWorkbook() As Workbook
    Set IDesignerPreparation_HostWorkbook = HostWorkbook
End Property

Private Sub IDesignerPreparation_Prepare(ByRef devManager As IDevelopment)
    Prepare devManager
End Sub

Private Sub IDesignerPreparation_SetFlag(ByVal nameId As String, ByVal enabled As Boolean)
    SetFlag nameId, enabled
End Sub
