VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropdownListsStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Tests.Stubs")
'@ModuleDescription("Minimal IDropdownLists implementation used to track dropdown wiring")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDropdownLists

Private Type TDropdownListsStub
    hostSheet As Worksheet
    addCalls As Collection
    validationCalls As Collection
    labelRanges As Collection
    returnLinkCalls As Collection
    hyperlinkCalls As Collection
End Type

Private this As TDropdownListsStub

Public Sub Initialise(ByVal hostSheet As Worksheet)
    Set this.hostSheet = hostSheet
    Set this.addCalls = New Collection
    Set this.validationCalls = New Collection
    Set this.labelRanges = New Collection
    Set this.returnLinkCalls = New Collection
    Set this.hyperlinkCalls = New Collection
End Sub

Public Function LastValidationShowError() As Boolean
    Dim record As Variant
    If this.validationCalls.Count = 0 Then Exit Function
    record = this.validationCalls(this.validationCalls.Count)
    LastValidationShowError = CBool(record(3))
End Function

Public Function LastAddedListName() As String
    Dim record As Variant
    If this.addCalls.Count = 0 Then Exit Function
    record = this.addCalls(this.addCalls.Count)
    LastAddedListName = CStr(record(1))
End Function

Public Function ReturnLinkCount() As Long
    ReturnLinkCount = this.returnLinkCalls.Count
End Function

Public Function ValidationCallCount() As Long
    ValidationCallCount = this.validationCalls.Count
End Function

Public Function ValidationListName(ByVal index As Long) As String
    Dim record As Variant

    If index < 1 Then Exit Function
    If index > this.validationCalls.Count Then Exit Function

    record = this.validationCalls(index)
    ValidationListName = CStr(record(1))
End Function

Public Function ValidationTargetAddress(ByVal index As Long) As String
    Dim record As Variant
    Dim target As Range

    If index < 1 Then Exit Function
    If index > this.validationCalls.Count Then Exit Function

    record = this.validationCalls(index)
    Set target = record(0)
    If target Is Nothing Then Exit Function
    ValidationTargetAddress = target.Address(False, False)
End Function

Public Function ValidationShowError(ByVal index As Long) As Boolean
    Dim record As Variant

    If index < 1 Then Exit Function
    If index > this.validationCalls.Count Then Exit Function

    record = this.validationCalls(index)
    ValidationShowError = CBool(record(3))
End Function

Private Function EnsureHostSheet() As Worksheet
    If this.hostSheet Is Nothing Then
        Err.Raise vbObjectError + 512, "DropdownListsStub", "Host worksheet not supplied"
    End If
    Set EnsureHostSheet = this.hostSheet
End Function

Private Function NextLabelCell() As Range
    Dim host As Worksheet
    Dim nextRow As Long

    Set host = EnsureHostSheet()
    nextRow = this.labelRanges.Count + 1
    Set NextLabelCell = host.Cells(nextRow, 26) 'Column Z keeps the labels away from the test range
End Function

Private Sub StoreLabelRange(ByVal listName As String, ByVal labelCell As Range)
    Dim entry(1) As Variant
    entry(0) = LCase$(Trim$(listName))
    Set entry(1) = labelCell
    this.labelRanges.Add entry
End Sub

Private Function ResolveLabelRange(ByVal listName As String) As Range
    Dim idx As Long
    Dim entry As Variant

    For idx = 1 To this.labelRanges.Count
        entry = this.labelRanges(idx)
        If StrComp(CStr(entry(0)), LCase$(Trim$(listName)), vbTextCompare) = 0 Then
            Set ResolveLabelRange = entry(1)
            Exit Function
        End If
    Next idx
End Function

'@section IDropdownLists implementation
'===============================================================================
Private Property Get IDropdownLists_Exists(ByVal listName As String) As Boolean
    IDropdownLists_Exists = Not (ResolveLabelRange(listName) Is Nothing)
End Property

Private Property Get IDropdownLists_Name() As String
    IDropdownLists_Name = "DropdownListsStub"
End Property

Private Property Get IDropdownLists_LabelRange(ByVal listName As String) As Range
    Set IDropdownLists_LabelRange = ResolveLabelRange(listName)
End Property

Private Property Get IDropdownLists_Values(ByVal listName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
    Dim result As BetterArray
    Set result = New BetterArray
    result.LowerBound = 1
    Set IDropdownLists_Values = result
End Property

Private Property Get IDropdownLists_Length() As Long
    IDropdownLists_Length = this.addCalls.Count
End Property

Private Sub IDropdownLists_Add(ByVal listValues As BetterArray, ByVal listName As String, _
                               Optional ByVal addLabel As Boolean = False, _
                               Optional ByVal counterPrefix As String = vbNullString, _
                               Optional ByVal tableStyle As String = "TableStyleLight13")
    Dim labelCell As Range
    Dim labelText As String
    Dim record(4) As Variant

    record(0) = listValues
    record(1) = listName
    record(2) = addLabel
    record(3) = counterPrefix
    record(4) = tableStyle
    this.addCalls.Add record

    Set labelCell = NextLabelCell()

    If addLabel Then
        If LenB(counterPrefix) > 0 Then
            labelText = counterPrefix & " 1"
        Else
            labelText = "choice 1"
        End If
    ElseIf Not listValues Is Nothing Then
        If listValues.Length > 0 Then
            labelText = CStr(listValues.Item(listValues.LowerBound))
        End If
    End If

    labelCell.Value = labelText
    StoreLabelRange listName, labelCell
End Sub

Private Sub IDropdownLists_Remove(ByVal listName As String)
    'Not required for tests
End Sub

Private Sub IDropdownLists_ClearList(ByVal listName As String)
    'Not required for tests
End Sub

Private Sub IDropdownLists_Update(ByVal UpdateData As BetterArray, ByVal listName As String, _
                                  Optional ByVal pasteAtBottom As Boolean = False, _
                                  Optional ByVal removeDuplicates As Boolean = True)
    'Not required for tests
End Sub

Private Sub IDropdownLists_Sort(ByVal listName As String, Optional ByVal sortOrder As Long = xlAscending)
    'Not required for tests
End Sub

Private Sub IDropdownLists_SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                                         Optional ByVal alertType As String = "info", _
                                         Optional ByVal message As String = vbNullString, _
                                         Optional ByVal showError As Boolean = True)
    Dim record(4) As Variant
    Set record(0) = cellRng
    record(1) = listName
    record(2) = alertType
    record(3) = showError
    record(4) = message
    this.validationCalls.Add record
End Sub

Private Sub IDropdownLists_AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    Dim record(1) As Variant
    record(0) = listName
    Set record(1) = returnRng
    this.returnLinkCalls.Add record
End Sub

Private Sub IDropdownLists_AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
    Dim record(1) As Variant
    record(0) = listName
    Set record(1) = linkRng
    this.hyperlinkCalls.Add record
End Sub

Private Property Get IDropdownLists_HasCheckings() As Boolean
    IDropdownLists_HasCheckings = False
End Property

Private Property Get IDropdownLists_CheckingValues() As Object
    Set IDropdownLists_CheckingValues = Nothing
End Property
*** End Patch
