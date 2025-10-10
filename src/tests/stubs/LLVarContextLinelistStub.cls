VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarContextLinelistStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests")
'@ModuleDescription("Test double implementing the ILinelist contract for LLVarContext tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILinelist

Private Type TLinelistStub
    hostWorkbook As Workbook
    specs As ILinelistSpecs
    dictionary As ILLdictionary
    dropdownDefault As IDropdownLists
    dropdownCustom As IDropdownLists
    printedSheets As Collection
End Type

Private this As TLinelistStub

Private Sub Class_Initialize()
    Set this.hostWorkbook = ThisWorkbook
    Set this.printedSheets = New Collection
End Sub

Public Sub UseWorkbook(ByVal wb As Workbook)
    Set this.hostWorkbook = wb
End Sub

Public Sub UseSpecs(ByVal specs As ILinelistSpecs)
    Set this.specs = specs
End Sub

Public Sub UseDictionary(ByVal dict As ILLdictionary)
    Set this.dictionary = dict
End Sub

Public Sub UseDropdowns(ByVal defaultDrop As IDropdownLists, _
                        Optional ByVal customDrop As IDropdownLists = Nothing)
    Set this.dropdownDefault = defaultDrop
    Set this.dropdownCustom = customDrop
End Sub

Public Sub UsePrintedSheet(ByVal sheetName As String, ByVal printedSheet As Worksheet)
    Dim key As String

    key = LCase$(Trim$(sheetName))

    On Error Resume Next
        this.printedSheets.Remove key
    On Error GoTo 0

    On Error Resume Next
        this.printedSheets.Add printedSheet, key
    On Error GoTo 0
End Sub

Private Function ResolveSheet(ByVal sheetName As String) As Worksheet
    Dim target As Worksheet
    On Error Resume Next
        Set target = this.hostWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    Set ResolveSheet = target
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim target As Worksheet

    Set target = ResolveSheet(sheetName)
    If target Is Nothing Then
        Set target = TestHelpers.EnsureWorksheet(sheetName, this.hostWorkbook)
    End If

    Set EnsureSheet = target
End Function

Private Sub ThrowNotImplemented(ByVal memberName As String)
    Err.Raise vbObjectError + 513, "LLVarContextLinelistStub." & memberName, _
              "Member not implemented in stub"
End Sub

'@section ILinelist implementation
'===============================================================================
Private Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = ResolveSheet(sheetName)
End Property

Private Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Dim key As String
    Dim printed As Worksheet

    key = LCase$(Trim$(sheetName))

    On Error Resume Next
        Set printed = this.printedSheets.Item(key)
    On Error GoTo 0

    If printed Is Nothing Then
        Set printed = ResolveSheet(sheetName)
    End If

    Set ILinelist_PrintWksh = printed
End Property

Private Property Get ILinelist_CRFWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_CRFWksh = ResolveSheet(sheetName)
End Property

Private Property Get ILinelist_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Select Case scope
    Case 1
        Set ILinelist_Dropdown = this.dropdownDefault
    Case Else
        If Not this.dropdownCustom Is Nothing Then
            Set ILinelist_Dropdown = this.dropdownCustom
        Else
            Set ILinelist_Dropdown = this.dropdownDefault
        End If
    End Select
End Property

Private Property Get ILinelist_Pivots() As ICustomPivotTables
    ThrowNotImplemented "Pivots"
End Property

Private Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = this.specs
End Property

Private Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = this.dictionary
End Property

Private Sub ILinelist_Prepare()
    'Not required for stub
End Sub

Private Function ILinelist_sheetExists(ByVal sheetName As String) As Boolean
    ILinelist_sheetExists = Not (ResolveSheet(sheetName) Is Nothing)
End Function

Private Sub ILinelist_AddOutputSheet(ByVal sheetName As String, _
                                     Optional ByVal Hide As Long = xlSheetVeryHidden, _
                                     Optional ByVal sheetScope As Byte = 1)
    Dim target As Worksheet

    Set target = EnsureSheet(sheetName)

    On Error Resume Next
        target.Visible = Hide
    On Error GoTo 0
End Sub

Private Sub ILinelist_SaveLL()
    ThrowNotImplemented "SaveLL"
End Sub

Private Sub ILinelist_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    ThrowNotImplemented "ErrorManage"
End Sub
