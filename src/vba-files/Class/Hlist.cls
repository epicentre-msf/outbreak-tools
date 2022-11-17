VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("All HList variables")
'@IgnoreModule

Option Explicit

Implements IHlist
Private Type THList
    sh As Worksheet
    data As ILinelistSpecs
    outll As ILinelist
End Type

Private this As THList
Private Const CLASSNAME As String = "HList"

Public Function Create(ByVal sheetName As String, lData As ILinelistSpecs, outputLL As ILinelist) As IHlist

    If outputLL Is Nothing Then ThrowError "The output worksheet does not exists"
    If lData Is Nothing Then ThrowError "Linelist specifications unfound"
    If Not outputLL.sheetExists(sheetName) Then ThrowError "Worksheet " & sheetName & " does not exists in the output workbook"

    With New Hlist
        .Name = sheetName
        Set .linelistData = lData
        Set .OutputLinelist = outputLL
        Set Create = .Self
    End With

End Function

Public Property Get Self() As IHlist
    Set Self = Me
End Property

Public Property Let Name(ByVal sheetName As String)
    this.Name = sheetName
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Get linelistData() As ILinelistSpecs
    Set linelistData = this.data
End Property

Public Property Set linelistData(ByVal lData As ILinelistSpecs)
    Set this.data = lData
End Property

Public Property Get OutputLinelist() As ILinelist
    Set OutputLinelist = this.outll
End Property

Public Property Set OutputLinelist(ByVal outll As ILinelist)
    Set this.outll = outll
End Property

Public Property Get Wksh() As Worksheet
    Dim outll As ILinelist

    Set outll = OutputLinelist()

    Set Wksh = outll.Wksh(Name)
End Property

Public Property Get Variables() As BetterArray
    Dim varlists As BetterArray
    Dim dict As ILLdictionary
    Dim lData As ILinelistSpecs
    Dim drop As IDropdownLists

    Set varlists = New BetterArray
    Set lData = linelistData()
    Set dict = lData.Dictionary()
    Set varlists = dict.data.FilterData("sheet name", Name, "variable name")

    Set Variables = varlists.Clone()
End Property

Public Sub Build()
    Dim varlists As BetterArray
    Dim columnsIndex As BetterArray
    Dim sectionsList As BetterArray
    Dim subSectionsList As BetterArray



    Dim counter As Long
    Dim varName As String
    Dim varObject As IHListVars
    Dim drop As IDropdownLists
    Dim outll As ILinelist
    Dim sh As Worksheet
    Dim lData As linelistData

    Set varlists = Variables()
    Set outll = OutputLinelist()
    Set sh = Wksh()
    Set drop = outll.Dropdown()
    Set lData = linelistData()

    For counter = varlists.LowerBound To varlists.UpperBound
        varName = varlists.Item(counter)
        Set varObject = HListVars.Create(varName, sh, lData, drop)
        varObject.WriteInfo
    Next

    'Add Sections

    'Add SubSections

    'Add dropdown for Goto sections

End Sub

'Error
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Public Sub IHList_Build()
    Build
End Sub

