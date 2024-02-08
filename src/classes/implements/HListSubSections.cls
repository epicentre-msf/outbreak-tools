VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HListSubSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Subsections of a Horizontal variable sheet"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Subsections of a Horizontal variable sheet")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements IHListSubSections
'Exposed methods

Private Type THListSubSections
    strt As Long
    endrw As Long
    ll As ILinelist
    varLists As BetterArray
End Type

Private this As THListSubSections
Private Const SUBSECROW As Long = 6              'Row for the sub section
Private Const CLASSNAME As String = "HListSubSections"

Public Function Create(ByVal strt As Long, ll As ILinelist) As IHListSubSections

    CheckRequirements strt, ll

    With New HListSubSections
        .StartRow = strt
        Set .LinelistObject = ll
        Set Create = .Self()
    End With

End Function

Public Property Get Self() As IHListSubSections
    Set Self = Me
End Property

Public Property Get StartRow() As Long           'Start Row of the subsection in the dictionary
    StartRow = this.strt
End Property

Public Property Let StartRow(ByVal strt As Long)
    this.strt = strt
    this.endrw = 0
End Property

Public Property Get EndRow() As Long             'End Row of the subsection in the dictionary (Not in the worksheet)

    Dim dict As ILLdictionary
    Dim rowNumber As Long
    Dim strt As Long
    Dim subSectionRng As Range
    Dim sectionRng As Range
    Dim sheetRng As Range
    Dim subSectionEnd As Boolean
    Dim subSectionName As String
    Dim sectionName As String
    Dim sheetName As String

    If (this.endrw = 0) Then
        rowNumber = StartRow()
        strt = StartRow()

        Set dict = Dictionary()
        Set subSectionRng = dict.DataRange("sub section")
        Set sectionRng = dict.DataRange("main section")
        Set sheetRng = dict.DataRange("sheet name")

        subSectionName = subSectionRng.Cells(strt, 1).Value
        sectionName = sectionRng.Cells(strt, 1).Value
        sheetName = sheetRng.Cells(strt, 1).Value

        subSectionEnd = False

        Do While Not subSectionEnd
            rowNumber = rowNumber + 1
            subSectionEnd = (subSectionRng.Cells(rowNumber, 1).Value <> subSectionName) Or _
                            (sectionRng.Cells(rowNumber, 1).Value <> sectionName) Or _ 
                            (sheetRng.Cells(rowNumber, 1).Value <> sheetName)
        Loop

        this.endrw = rowNumber - 1
    End If

    EndRow = this.endrw
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

Private Property Get ActualSubSection() As String
    Dim dict As ILLdictionary
    Dim strt As Long
    Dim subSectionRng As Range

    Set dict = Dictionary()
    strt = StartRow()
    Set subSectionRng = dict.DataRange("sub section")
    ActualSubSection = subSectionRng.Cells(strt, 1).Value
End Property

'List of all variables in one sub section
Private Property Get VariableLists() As BetterArray
    Dim varl As BetterArray
    Dim counter As Long
    Dim strt As Long
    Dim endrw As Long
    Dim varRng As Range
    Dim dict As ILLdictionary

    If this.varLists Is Nothing Then
        Set dict = Dictionary()
        Set varl = New BetterArray
        Set varRng = dict.DataRange("variable name")
        strt = StartRow()
        endrw = EndRow()

        For counter = strt To endrw
            varl.Push varRng.Cells(counter, 1).Value
        Next

        Set this.varLists = New BetterArray
        Set this.varLists = varl.Clone()
    End If

    Set VariableLists = this.varLists.Clone()
End Property

Private Property Get HasNext() As Boolean
    Dim endrw As Long
    Dim strt As Long
    Dim dict As ILLdictionary
    Dim sectRng As Range
    Dim sectName As String
    Dim sheetName As String
    Dim sheetRng As Range

    endrw = EndRow() + 1
    strt = StartRow()
    Set dict = Dictionary()
    Set sectRng = dict.DataRange("main section")
    Set sheetRng = dict.DataRange("sheet name")
    sectName = sectRng.Cells(strt, 1).Value
    sheetName = sheetRng.Cells(strt, 1).Value
    HasNext = (sectRng.Cells(endrw, 1).Value = sectName) And _ 
              (endrw <= dict.Data.DataEndRow()) And _ 
              (sheetRng.Cells(endrw, 1).Value = sheetName)
End Property

Private Property Get NextSubSection() As IHListSubSections
    Dim nextSubSec As IHListSubSections
    Dim ll As ILinelist
    Dim strt As Long

    If Not HasNext() Then ThrowError "no subsections in actual section, I can't provide next subsection"

    Set ll = LinelistObject()
    strt = EndRow() + 1
    Set nextSubSec = HListSubSections.Create(strt, ll)
    Set NextSubSection = nextSubSec
End Property

'Build variables in a subsection
Private Sub Build()
    Dim vars As IHListVars
    Dim crfv As ICRFVars
    Dim varLists As BetterArray
    Dim counter As Long
    Dim varName As String
    Dim ll As ILinelist
    Set ll = LinelistObject()
    Set varLists = VariableLists()
    'For each variable in one subSection build that sub section
    For counter = varLists.LowerBound To varLists.UpperBound
        varName = varLists.Item(counter)
        Set vars = HListVars.Create(varName, ll)
        vars.WriteInfo
        'Write informations for the crf variable
        Set crfv = CRFVars.Create(varName, ll)
        crfv.WriteInfo
    Next
End Sub

'Format the sub section
Private Sub Format()
    Dim rng As Range
    Dim varLists As BetterArray
    Dim strtcol As Long                          'Start Column of the subsecction
    Dim endCol As Long                           'End column of the sub section
    Dim ll As ILinelist
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim varName As String
    Dim subSectionName As String
    Dim sheetName As String
    Dim sh As Worksheet
    Dim design As ILLFormat
    Dim crfLine As Long

    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)
    Set varLists = VariableLists()
    Set ll = LinelistObject()
    Set design = ll.LinelistData.DesignFormat()

    'compute the endColumn
    varName = varLists.Item(varLists.UpperBound)
    endCol = CLng(vars.Value(colName:="column index", varName:=varName))

    'compute the start column
    varName = varLists.Item(varLists.LowerBound)
    strtcol = CLng(vars.Value(colName:="column index", varName:=varName))
    crfLine = CLng(vars.Value(colName:="crf index", varName:=varName))

    'Actual sub section
    subSectionName = ActualSubSection()
    sheetName = vars.Value(colName:="sheet name", varName:=varName)

    Set sh = ll.Wksh(sheetName)
    With sh
        Set rng = .Range(.Cells(SUBSECROW, strtcol), .Cells(SUBSECROW, endCol))
        rng.Cells(1, 1).Value = subSectionName
    End With

    design.ApplyFormat rng, scope:=HListSubSection

    'Do the same for the printed Worksheek
    Set sh = ll.PrintWksh(sheetName)
    With sh
        Set rng = .Range(.Cells(SUBSECROW, strtcol), .Cells(SUBSECROW, endCol))
        rng.Cells(1, 1).Value = subSectionName
    End With

    design.ApplyFormat rng, scope:=HListSubSection

    'Add subsections on crf worksheet
    'For CRF sheet sub sections sections
    Set sh = ll.CRFWksh(sheetName)
    With sh
        Set rng = .Range(.Cells(crfLine - 2, 1), .Cells(crfLine - 2, 7))
        rng.Cells(1, 1).Value = subSectionName
    End With

    design.ApplyFormat rng, scope:=HListCRFSubSection
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

'Interface
Private Property Get IHListSubSections_HasNext() As Boolean
    IHListSubSections_HasNext = HasNext()
End Property

Private Property Get IHListSubSections_NextSubSection() As IHListSubSections
    Set IHListSubSections_NextSubSection = NextSubSection()
End Property

Private Sub IHListSubSections_Build()
    Build
End Sub

Private Sub IHListSubSections_Format()
    Format
End Sub
