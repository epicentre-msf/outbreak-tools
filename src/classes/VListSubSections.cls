VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VListSubSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Vlist sub sections")
'@IgnoreModule

Option Explicit
Implements IVListSubSections
'Exposed methods

Private Type TVListSubIVListSubSections
    strt As Long
    endrw As Long
    ll As ILinelist
    Varlists As BetterArray
End Type

Private this As TVListSubIVListSubSections
Private Const SUBSECCOL As Long = 3              'Row for the sub section
Private Const CLASSNAME As String = "IVListSubSections"

Public Function Create(ByVal strt As Long, ll As ILinelist) As IVListSubSections

    CheckRequirements strt, ll

    With New VListSubSections
        .StartRow = strt
        Set .LinelistObject = ll
        Set Create = .Self()
    End With

End Function

Public Property Get Self() As IVListSubSections
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
    Dim subSectionRng As Range
    Dim sectionRng As Range
    Dim subSectionEnd As Boolean
    Dim subSectionName As String
    Dim sectionName As String

    If this.endrw = 0 Then
        rowNumber = StartRow()
        Set dict = Dictionary()
        Set subSectionRng = dict.DataRange("sub section")
        Set sectionRng = dict.DataRange("main section")
        subSectionName = ActualSubSection()
        sectionName = ActualSection()
        subSectionEnd = False

        Do While Not subSectionEnd
            rowNumber = rowNumber + 1
            subSectionEnd = (subSectionRng.Cells(rowNumber, 1).Value <> subSectionName) Or _
                                                                                        (sectionRng.Cells(rowNumber, 1).Value <> sectionName)
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

Private Property Get ActualSection() As String
    Dim dict As ILLdictionary
    Dim strt As Long
    Dim sectionRng As Range
  
    Set dict = Dictionary()
    strt = StartRow()
    Set sectionRng = dict.DataRange("main section")

    ActualSection = sectionRng.Cells(strt, 1).Value
End Property

Private Property Get HasNext() As Boolean
    Dim endrw As Long
    Dim dict As ILLdictionary
    Dim sectRng As Range
    Dim sectName As String
    
    endrw = EndRow() + 1
    sectName = ActualSection()
    Set dict = Dictionary()
    Set sectRng = dict.DataRange("main section")
    HasNext = (sectRng.Cells(endrw, 1).Value = sectName) And (endrw <= dict.Data.DataEndRow)
End Property

Private Property Get NextSubSection() As IVListSubSections
    Dim nextSubSec As IVListSubSections
    Dim ll As ILinelist
    Dim strt As Long
  
    If Not HasNext() Then ThrowError "no subsections in actual section, I can't provide next subsection"
  
    Set ll = LinelistObject()
    strt = EndRow() + 1
    Set nextSubSec = VListSubSections.Create(strt, ll)
    Set NextSubSection = nextSubSec
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

    If this.Varlists Is Nothing Then
        Set dict = Dictionary()
        Set varl = New BetterArray
        Set varRng = dict.DataRange("variable name")
        strt = StartRow()
        endrw = EndRow()

        For counter = strt To endrw
            varl.Push varRng.Cells(counter, 1).Value
        Next

        Set this.Varlists = New BetterArray
        Set this.Varlists = varl.Clone()
    End If

    Set VariableLists = this.Varlists.Clone()
End Property

'Build variables in a subsection
Private Sub Build()
    Dim vars As IVListVars
    Dim Varlists As BetterArray
    Dim counter As Long
    Dim varName As String
    Dim ll As ILinelist
    Set ll = LinelistObject()
    Set Varlists = VariableLists()
    'For each variable in one subSection build that sub section
    For counter = Varlists.LowerBound To Varlists.UpperBound
        varName = Varlists.Item(counter)
        Set vars = VListVars.Create(varName, ll)
        vars.WriteInfo
    Next
End Sub

'Format the sub section
Private Sub Format()
    Dim rng As Range
    Dim Varlists As BetterArray
    Dim strtrw As Long                           'Start Row of the subsecction (in vlist linelist)
    Dim endrw As Long                            'End Row of the sub section
    Dim ll As ILinelist
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim varName As String
    Dim subSectionName As String
    Dim sheetName As String
    Dim sh As Worksheet
    Dim design As ILLFormat


    Set dict = Dictionary()
    Set vars = LLVariables.Create(dict)
    Set Varlists = VariableLists()
    Set ll = LinelistObject()
    Set design = ll.LinelistData.DesignFormat()
  
    'compute the endColumn
    varName = Varlists.Item(Varlists.UpperBound)
    endrw = CLng(vars.Value(colName:="column index", varName:=varName))
  
    'compute the start column
    varName = Varlists.Item(Varlists.LowerBound)
    strtrw = CLng(vars.Value(colName:="column index", varName:=varName))
  
    'Actual sub section
    subSectionName = ActualSubSection()
    sheetName = vars.Value(colName:="sheet name", varName:=varName)

    Set sh = ll.Wksh(sheetName)
  
    With sh
    
        Set rng = .Range(.Cells(strtrw, SUBSECCOL), .Cells(endrw, SUBSECCOL))
        rng.Cells(1, 1).Value = subSectionName
    End With

    design.ApplyFormat rng, scope:=VListSubSection

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

'Build the sub sections
Public Sub IVListSubSections_Build()
    Build
End Sub

'Format the subSection
Public Sub IVListSubSections_Format()
    Format
End Sub

Public Property Get IVListSubSections_HasNext() As Boolean
    IVListSubSections_HasNext = HasNext()
End Property

Public Property Get IVListSubSections_NextSubSection() As IVListSubSections
    Set IVListSubSections_NextSubSection = NextSubSection()
End Property


