VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Final linelist workbook")
'@IgnoreModule

Option Explicit
Implements ILinelist
Private Type TLinelist
    outwkb As Workbook
    data As ILinelistSpecs
End Type

Private this As TLinelist
Private Const CLASSNAME  As String = "Linelist"

Public Function Create(lData As ILinelistSpecs) As ILinelist

    If lData Is Nothing Then ThrowError "Linelist specifications unfound"
    With New Linelist
        Set .linelistData = lData
        Set Create = .Self()
    End With

End Function

Public Property Get Self() As ILinelist
    Set Self = Me
End Property

Public Property Get linelistData() As ILinelistSpecs
    Set linelistData = this.data
End Property

Public Property Set linelistData(ByVal lData As ILinelistSpecs)
    Set this.data = lData
End Property

'Dictionary
Public Property Get Dictionary()
    Dim dict As ILLdictionary
    Dim lData As ILinelistSpecs

    Set lData = linelistData()
    Set dict = lData.Dictionary()
End Property

'Workbook
Public Property Get Wkb() As Workbook
    If this.outwkb Is Nothing Then
        Set this.outwkb = Workbooks.Add()
    End If
    Set Wkb = this.outwkb
End Property

'Test if a worksheet Exists
Private Function sheetExists(ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    Dim outwkb As Workbook

    Set outwkb = Wkb()
    On Error Resume Next
    Set sh = outwkb.Worksheets(sheetName)
    On Error GoTo 0
    sheetExists = (Not sh Is Nothing)
End Function

'Get the dropdown object
Private Property Get Dropdown() As IDropdownLists
    Dim sh As Worksheet
    Dim listSheetName As String
    Dim lData As ILinelistSpecs
    Dim drop As IDropdownLists

    Set lData = linelistData()
    listSheetName = lData.TemporarySheetName(TemporarySheetList)
    Set sh = Wksh(listSheetName)

    Set drop = DropdownLists.Create(sh)
    Set Dropdown = drop
End Property

'Add worksheets to workbook
Private Sub AddOutputSheet(ByVal sheetName As String, Optional ByVal Hide As Long = xlSheetVeryHidden)
    Dim outwkb As Workbook
    Set outwkb = Wkb()

    If sheetExists(sheetName) Then ThrowError "The sheet " & sheetName & "already exists in the output workbook"
    With outwkb
        .Worksheets.Add.Name = sheetName
        .Worksheets(sheetName).Visible = Hide
    End With
End Sub

'Get a Worksheet in the output workbook
Private Property Get Wksh(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet
    Dim outwkb As Workbook

    'Test if the sheet Exists
    If Not sheetExists(sheetName) Then ThrowError "The sheet " & sheetName & "does not exists in current workbook"
    Set outwkb = Wkb()
    Set sh = outwkb.Worksheets(sheetName)
    Set Wksh = sh
End Property

'Prepare the Linelist by adding required worksheets
Private Sub Prepare()
    Dim dict As ILLdictionary
    Dim trans As ITranslation
    Dim sheetsList As BetterArray
    Dim lData As ILinelistSpecs
    Dim counter As Long
    Dim sheetName As String
    Dim outwkb As Workbook
    
    Set outwkb = Wkb()
    
    Set lData = linelistData()
    Set sheetsList = New BetterArray
    Set trans = lData.TransObject()
    Set dict = lData.Dictionary()


    'temporary sheet Names
    sheetsList.Push lData.TemporarySheetName(TemporarySingle), lData.TemporarySheetName(TemporarySheetList), _
        lData.TemporarySheetName(TemporarySheetDebug), lData.TemporarySheetName(TemporarySheetMetadata), _
        lData.TemporarySheetName(TemporarySheetAnalysis), lData.TemporarySheetName(TemporarySheetImport)

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVeryHidden
    Next
    
    'Unique sheet Names (HList or VList)
    sheetsList.Clear
    Set sheetsList = dict.UniqueValues("sheet name")
    sheetsList.Reverse

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlsheetvisible
    Next
    
    'Admin sheet Name
    AddOutputSheet trans.TranslatedValue("LLSHEET_Admin"), xlsheetvisible
    
    'Delete the first worksheet with name sheet1
     outwkb.Worksheets(1).Delete
End Sub

Private Sub Build()
    'Prepare the workbook
    Prepare

    '
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Properties
Public Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Wksh(sheetName)
End Property

Public Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = linelistData()
End Property

Public Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = Dictionary()
End Property

Public Property Get ILinelist_Dropdown() As IDropdownLists
    Set ILinelist_Dropdown = Dropdown()
End Property

'Exposed methods
Public Sub ILinelist_Prepare()
    Prepare
End Sub

Public Function ILinelist_SheetExists(ByVal sheetName As String) As Boolean
    ILinelist_SheetExists = sheetExists(sheetName)
End Function

Public Sub ILinelist_Build()
    Build
End Sub


