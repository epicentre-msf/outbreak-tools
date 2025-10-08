VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistRecordingAccessorStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("ILinelist implementation that records sheet creation order while exposing workbook accessor semantics")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ILinelist
Implements ILinelistWorkbookAccessor
Implements ILinelistWorkbookReleaser

Private Type TLinelist
    dictionary As ILLdictionary
    specs As ILinelistSpecs
    workbook As Workbook
    createdSheets As BetterArray
    clearCount As Long
End Type

Private this As TLinelist

Public Sub Initialise(ByVal dictionary As ILLdictionary, _
                      ByVal specs As ILinelistSpecs, _
                      Optional ByVal workbook As Workbook)
    Set this.dictionary = dictionary
    Set this.specs = specs
    Set this.workbook = workbook

    Set this.createdSheets = New BetterArray
    this.createdSheets.LowerBound = 1
    this.clearCount = 0
End Sub

Public Property Get CreatedSheets() As BetterArray
    Set CreatedSheets = this.createdSheets
End Property

Public Property Get ClearCount() As Long
    ClearCount = this.clearCount
End Property

Private Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Nothing
End Property

Private Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_PrintWksh = Nothing
End Property

Private Property Get ILinelist_CRFWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_CRFWksh = Nothing
End Property

Private Property Get ILinelist_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Set ILinelist_Dropdown = Nothing
End Property

Private Property Get ILinelist_Pivots() As ICustomPivotTables
    Set ILinelist_Pivots = Nothing
End Property

Private Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = this.specs
End Property

Private Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = this.dictionary
End Property

Private Sub ILinelist_Prepare()
End Sub

Private Function ILinelist_sheetExists(ByVal sheetName As String) As Boolean
    Dim index As Long

    For index = this.createdSheets.LowerBound To this.createdSheets.UpperBound
        If StrComp(CStr(this.createdSheets.Item(index)), sheetName, vbTextCompare) = 0 Then
            ILinelist_sheetExists = True
            Exit Function
        End If
    Next index

    ILinelist_sheetExists = False
End Function

Private Sub ILinelist_AddOutputSheet(ByVal sheetName As String, Optional ByVal Hide As Long = xlSheetVeryHidden, Optional ByVal sheetScope As Byte = 1)
    this.createdSheets.Push sheetName
End Sub

Private Sub ILinelist_SaveLL()
End Sub

Private Sub ILinelist_ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub

Private Property Get ILinelistWorkbookAccessor_OutputWorkbook() As Workbook
    Set ILinelistWorkbookAccessor_OutputWorkbook = this.workbook
End Property

Private Sub ILinelistWorkbookReleaser_ClearOutputWorkbook()
    this.clearCount = this.clearCount + 1
    Set this.workbook = Nothing
End Sub

