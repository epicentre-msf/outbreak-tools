VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HListSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Sections of Horizontal worksheets of the linelist"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Sections of Horizontal worksheets of the linelist")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements IHListSections

Private Type THListSections
    strt As Long
    endrw As Long
    ll As ILinelist
End Type

Private this As THListSections
Private Const SECROW As Long = 5                 'Row for the sub section
Private Const CLASSNAME As String = "HListSections"

Public Function Create(ByVal strt As Long, ll As ILinelist) As IHListSections
    CheckRequirements strt, ll
    With New HListSections
        .StartRow = strt
        Set .LinelistObject = ll
        Set Create = .Self()
    End With

End Function

Public Property Get Self() As IHListSections
    Set Self = Me
End Property

Public Property Get StartRow() As Long           'Start Row of the subsection in the dictionary
    StartRow = this.strt
End Property

Public Property Let StartRow(ByVal strt As Long)
    this.strt = strt
    this.endrw = 0
End Property

Public Property Get LinelistObject() As ILinelist
    Set LinelistObject = this.ll
End Property

Public Property Set LinelistObject(ByVal ll As ILinelist)
    Set this.ll = ll
End Property

Public Property Get Dictionary() As ILLdictionary
    Dim ll As ILinelist

    Set ll = LinelistObject()
    Set Dictionary = ll.Dictionary()
End Property

Private Property Get ActualSectionName() As String
    Dim dict As ILLdictionary
    Dim strt As Long
    Dim sectionRng As Range
  
    Set dict = Dictionary()
    strt = StartRow()
    Set sectionRng = dict.DataRange("main section")

    ActualSectionName = sectionRng.Cells(strt, 1).Value
End Property

Private Property Get ActualSheetName() As String
    Dim dict As ILLdictionary
    Dim strt As Long
    Dim sheetRng As Range
  
    Set dict = Dictionary()
    strt = StartRow()
    Set sheetRng = dict.DataRange("sheet name")

    ActualSheetName = sheetRng.Cells(strt, 1).Value
End Property

Public Property Get EndRow() As Long             'End Row of the subsection in the dictionary (Not in the worksheet)

    Dim dict As ILLdictionary
    Dim rowNumber As Long
    Dim sectionRng As Range
    Dim sheetRng As Range
    Dim sectionName As String
    Dim sheetName As String
    Dim sectionEnd As Boolean

    If (this.endrw = 0) Then
        rowNumber = StartRow()
        Set dict = Dictionary()
        Set sectionRng = dict.DataRange("main section")
        Set sheetRng = dict.DataRange("sheet name")
        sectionName = ActualSectionName()
        sheetName = ActualSheetName()
        sectionEnd = False

        Do While Not sectionEnd
            rowNumber = rowNumber + 1
            sectionEnd = (sectionRng.Cells(rowNumber, 1).Value <> sectionName) Or _
                                                                               (sheetRng.Cells(rowNumber, 1).Value <> sheetName)
        Loop
        this.endrw = rowNumber - 1
    End If

    EndRow = this.endrw
End Property

Private Property Get HasNext() As Boolean
    Dim endrw As Long
    Dim dict As ILLdictionary
    Dim sheetRng As Range
    Dim sheetName As String

    sheetName = ActualSheetName()
    endrw = EndRow() + 1
    Set dict = Dictionary()
    Set sheetRng = dict.DataRange("sheet name")

    HasNext = (endrw <= dict.Data.DataEndRow()) And (sheetRng.Cells(endrw, 1).Value = sheetName)
End Property

Private Property Get NextSection() As IHListSections
    Dim nextSec As IHListSections
    Dim ll As ILinelist
    Dim strt As Long

    If Not HasNext() Then ThrowError "no new sections, I can't provide next section"

    Set ll = LinelistObject()
    strt = EndRow() + 1
    Set nextSec = HListSections.Create(strt, ll)
    Set NextSection = nextSec
End Property

Private Sub Build()
    'Build subsequent sections
    Dim subsec As IHListSubSections
    Dim ll As ILinelist
    Dim strt As Long

    Set ll = LinelistObject()
    strt = StartRow()
    Set subsec = HListSubSections.Create(strt, ll)
    subsec.Build
    subsec.Format
  
    Do While subsec.HasNext()
        Set subsec = subsec.NextSubSection()
        subsec.Build
        subsec.Format
    Loop
  
End Sub

Private Sub Format()
  
    Dim rng As Range
    Dim strtcol As Long
    Dim strtrw As Long
    Dim endrw As Long
    Dim endCol As Long
    Dim ll As ILinelist
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim sectionName As String
    Dim sheetName As String
    Dim varRng As Range
    Dim varName As String
    Dim sh As Worksheet
    Dim design As ILLFormat

    Set ll = LinelistObject()
    Set design = ll.LinelistData.DesignFormat()
    Set dict = Dictionary()
    Set varRng = dict.DataRange("variable name")
    Set vars = LLVariables.Create(dict)
    strtrw = StartRow()                          'End and start row of the section in the dictionary
    endrw = EndRow()


    'Compute the startColumn
    varName = varRng.Cells(strtrw, 1).Value
    strtcol = vars.Value(varName:=varName, colName:="column index")

    'Compute the endColumn
    varName = varRng.Cells(endrw, 1).Value
    endCol = vars.Value(varName:=varName, colName:="column index")

    'sheet name
    sheetName = ActualSheetName()
    sectionName = ActualSectionName()
    Set sh = ll.Wksh(sheetName)

    With sh
        Set rng = .Range(.Cells(SECROW, strtcol), .Cells(SECROW, endCol))
        rng.Cells(1, 1).Value = sectionName
    End With

    design.ApplyFormat rng, scope:=HListSection

    'Same for Printed sheet
    Set sh = ll.PrintWksh(sheetName)
    With sh
        Set rng = .Range(.Cells(SECROW, strtcol), .Cells(SECROW, endCol))
        rng.Cells(1, 1).Value = sectionName
    End With

    design.ApplyFormat rng, scope:=HListSection

    'For CRF sheet sections
    Set sh = ll.CRFWksh(sheetName)
    With sh
        Set rng = .Cells(.Cells(.Rows.Count, 1).End(xlUp).Row + 4, 1)
        Set rng = .Range(rng, rng.Cells(1, 7))
        rng.Cells(1, 1).Value = sectionName
    End With

    design.ApplyFormat rng, scope:=HListCRFSection
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Execution Error : " & errorMessage
End Sub

'requirements for creating an objec
Private Sub CheckRequirements(ByVal strt As Long, ll As ILinelist)
    If ll Is Nothing Then ThrowError "Please provide a linelist object, no linelist object found"
    If strt < 0 Or strt > ll.Dictionary.Data.DataEndRow() Then ThrowError "Start row should be a positive integer less or equal than end row of the dictionary"
End Sub

'Exposed methods

Private Property Get IHListSections_HasNext() As Boolean
    IHListSections_HasNext = HasNext()
End Property

Private Property Get IHListSections_NextSection() As IHListSections
    Set IHListSections_NextSection = NextSection()
End Property

'Build the sub sections
Private Sub IHListSections_Build()
    Build
End Sub

'Format the subSection
Private Sub IHListSections_Format()
    Format
End Sub

'Add section Name
Private Property Get IHListSections_Name() As String
    IHListSections_Name = ActualSectionName()
End Property
