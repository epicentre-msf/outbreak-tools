VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Compares two disease tables and answers the differences as Checking bundles"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Compares two disease tables and answers the differences as Checking bundles")
'@depends BetterArray, Checking
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class DiseaseComparison
'@description
'Reads two disease ListObjects and answers five Checking bundles: the
'variables only in the first disease, the variables only in the second, the
'shared variables whose choice sets differ, the shared variables whose choice
'sets match, and the counts over the two tables.
'
'WHAT IS COMPARED
'-------------------------------------------------------------------------------
'A disease table carries Variable Order, Variable Name, Variable Section,
'Main Label, Choice, Choice Values and Status. The name column keys the
'comparison. Names are matched without regard to case and surrounding spaces.
'The Choice Values cell is split on " | " and the pieces are compared as a
'set, without regard to case. A variable whose Choice cell is empty carries
'no choices, so two such variables are equivalent.
'
'LOOKUPS
'-------------------------------------------------------------------------------
'The name index of each table is a keyed Collection. Dictionary is missing on
'Mac Excel. The lists are BetterArray based at 1.

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "DiseaseComparison"

Private Const HEADER_NAME As String = "Variable Name"
Private Const HEADER_SECTION As String = "Variable Section"
Private Const HEADER_LABEL As String = "Main Label"
Private Const HEADER_CHOICE As String = "Choice"
Private Const HEADER_CHOICE_VALUES As String = "Choice Values"

Private Const CHOICE_SEPARATOR As String = " | "
Private Const REPORT_SUBTITLE As String = "Type--Where?--Details"

'@section State
'===============================================================================

'One disease table read into memory: its names in order, an index from the
'normalised name to the row, and the section, label and choices of each row.
Private Type TDiseaseData
    displayName As String
    rowCount As Long
    names As BetterArray
    keys As BetterArray
    sections As BetterArray
    labels As BetterArray
    choiceSets As BetterArray
    index As Collection
End Type

Private Type TComparisonState
    firstTable As ListObject
    secondTable As ListObject
    first As TDiseaseData
    second As TDiseaseData
    compared As Boolean
    onlyInFirst As Checking
    onlyInSecond As Checking
    differing As Checking
    equivalent As Checking
    statistics As Checking
    sharedCount As Long
    matchingCount As Long
    differingCount As Long
End Type

Private this As TComparisonState

'@section Factory
'===============================================================================

'@sub-title Build a comparison over two disease tables.
'@param firstTable ListObject. The disease table of the first disease.
'@param secondTable ListObject. The disease table of the second disease.
'@param firstName Optional String. Name shown for the first disease. Defaults to "Disease 1".
'@param secondName Optional String. Name shown for the second disease. Defaults to "Disease 2".
'@return DiseaseComparison ready to compare.
Public Function Create(ByVal firstTable As ListObject, ByVal secondTable As ListObject, _
                       Optional ByVal firstName As String = "Disease 1", _
                       Optional ByVal secondName As String = "Disease 2") As DiseaseComparison
    Dim instance As DiseaseComparison

    Set instance = New DiseaseComparison
    instance.Initialise firstTable, secondTable, firstName, secondName
    Set Create = instance
End Function

Public Sub Initialise(ByVal firstTable As ListObject, ByVal secondTable As ListObject, _
                      ByVal firstName As String, ByVal secondName As String)
    ValidateTable firstTable, "firstTable"
    ValidateTable secondTable, "secondTable"

    Set this.firstTable = firstTable
    Set this.secondTable = secondTable
    this.first.displayName = DisplayNameOrDefault(firstName, "Disease 1")
    this.second.displayName = DisplayNameOrDefault(secondName, "Disease 2")
    this.compared = False
End Sub

'@section Comparison
'===============================================================================

'@sub-title Read both tables and build the five bundles.
'@details
'Runs once. Every accessor calls it, so a caller can read the bundles
'straight after Create.
Public Sub Compare()
    If this.compared Then Exit Sub

    ReadTable this.firstTable, this.first
    ReadTable this.secondTable, this.second

    Set this.onlyInFirst = Checking.Create("Variables only in " & this.first.displayName, REPORT_SUBTITLE)
    Set this.onlyInSecond = Checking.Create("Variables only in " & this.second.displayName, REPORT_SUBTITLE)
    Set this.differing = Checking.Create("Shared variables whose choices differ", REPORT_SUBTITLE)
    Set this.equivalent = Checking.Create("Equivalent variables", REPORT_SUBTITLE)
    Set this.statistics = Checking.Create("Comparison statistics", REPORT_SUBTITLE)

    this.sharedCount = 0
    this.matchingCount = 0
    this.differingCount = 0

    CollectMissing this.first, this.second, this.onlyInFirst, "only-in-1"
    CollectMissing this.second, this.first, this.onlyInSecond, "only-in-2"
    CollectShared
    CollectStatistics

    this.compared = True
End Sub

'@section Bundles
'===============================================================================

'@prop-title Variables of the first disease the second has no row for.
Public Property Get OnlyInFirst() As Checking
    Compare
    Set OnlyInFirst = this.onlyInFirst
End Property

'@prop-title Variables of the second disease the first has no row for.
Public Property Get OnlyInSecond() As Checking
    Compare
    Set OnlyInSecond = this.onlyInSecond
End Property

'@prop-title Shared variables whose choice sets differ, with the values each side lacks.
Public Property Get DifferingChoices() As Checking
    Compare
    Set DifferingChoices = this.differing
End Property

'@prop-title Shared variables whose choice sets match.
Public Property Get EquivalentVariables() As Checking
    Compare
    Set EquivalentVariables = this.equivalent
End Property

'@prop-title The counts over the two tables.
Public Property Get Statistics() As Checking
    Compare
    Set Statistics = this.statistics
End Property

'@prop-title The five bundles in report order, based at 1.
'@details Only in 1, only in 2, differing choices, equivalent, statistics.
Public Property Get Reports() As BetterArray
    Dim bundles As BetterArray

    Compare

    Set bundles = New BetterArray
    bundles.LowerBound = 1
    bundles.Push this.onlyInFirst, this.onlyInSecond, this.differing, this.equivalent, this.statistics
    Set Reports = bundles
End Property

'@section Counts
'===============================================================================

'@prop-title Number of variables in the first disease.
Public Property Get FirstCount() As Long
    Compare
    FirstCount = this.first.rowCount
End Property

'@prop-title Number of variables in the second disease.
Public Property Get SecondCount() As Long
    Compare
    SecondCount = this.second.rowCount
End Property

'@prop-title Number of variable names both diseases carry.
Public Property Get SharedCount() As Long
    Compare
    SharedCount = this.sharedCount
End Property

'@prop-title Number of variables only the first disease carries.
Public Property Get OnlyInFirstCount() As Long
    Compare
    OnlyInFirstCount = this.first.rowCount - this.sharedCount
End Property

'@prop-title Number of variables only the second disease carries.
Public Property Get OnlyInSecondCount() As Long
    Compare
    OnlyInSecondCount = this.second.rowCount - this.sharedCount
End Property

'@prop-title Number of shared variables whose choice sets match.
Public Property Get MatchingCount() As Long
    Compare
    MatchingCount = this.matchingCount
End Property

'@prop-title Number of shared variables whose choice sets differ.
Public Property Get DifferingCount() As Long
    Compare
    DifferingCount = this.differingCount
End Property

'@prop-title Share of the shared variables that are equivalent, between 0 and 1.
'@details Answers 0 when the two diseases share no variable.
Public Property Get EquivalentShare() As Double
    Compare
    If this.sharedCount = 0 Then Exit Property
    EquivalentShare = this.matchingCount / this.sharedCount
End Property

'@section Reading
'===============================================================================

'@sub-title Read the columns a comparison needs into memory.
Private Sub ReadTable(ByVal table As ListObject, ByRef data As TDiseaseData)
    Dim nameValues As Variant
    Dim sectionValues As Variant
    Dim labelValues As Variant
    Dim choiceValues As Variant
    Dim choiceListValues As Variant
    Dim rowIndex As Long
    Dim rowTotal As Long
    Dim variableName As String
    Dim key As String

    Set data.names = NewList()
    Set data.keys = NewList()
    Set data.sections = NewList()
    Set data.labels = NewList()
    Set data.choiceSets = NewList()
    Set data.index = New Collection
    data.rowCount = 0

    rowTotal = table.ListRows.Count
    If rowTotal = 0 Then Exit Sub

    nameValues = ColumnValues(table, HEADER_NAME, rowTotal)
    sectionValues = ColumnValues(table, HEADER_SECTION, rowTotal)
    labelValues = ColumnValues(table, HEADER_LABEL, rowTotal)
    choiceValues = ColumnValues(table, HEADER_CHOICE, rowTotal)
    choiceListValues = ColumnValues(table, HEADER_CHOICE_VALUES, rowTotal)

    For rowIndex = 1 To rowTotal
        variableName = Trim$(CStr(nameValues(rowIndex, 1)))
        key = LCase$(variableName)
        'A blank name is an empty line of the table. A name already seen is a
        'duplicate; the first row keeps it, the same as the importer does.
        If LenB(key) > 0 Then
            If Not IndexHas(data.index, key) Then
                data.rowCount = data.rowCount + 1
                data.index.Add data.rowCount, key
                data.names.Push variableName
                data.keys.Push key
                data.sections.Push Trim$(CStr(sectionValues(rowIndex, 1)))
                data.labels.Push Trim$(CStr(labelValues(rowIndex, 1)))
                data.choiceSets.Push SplitChoices(CStr(choiceValues(rowIndex, 1)), CStr(choiceListValues(rowIndex, 1)))
            End If
        End If
    Next rowIndex
End Sub

'@sub-title Read one column of the table as a 2D array based at 1.
Private Function ColumnValues(ByVal table As ListObject, ByVal header As String, ByVal rowTotal As Long) As Variant
    Dim tableColumn As ListColumn
    Dim values As Variant
    Dim singleCell(1 To 1, 1 To 1) As Variant

    Set tableColumn = FindColumn(table, header)
    If tableColumn Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "The table " & table.Name & " has no column " & header
    End If

    values = tableColumn.DataBodyRange.Value
    If rowTotal = 1 Then
        singleCell(1, 1) = values
        ColumnValues = singleCell
    Else
        ColumnValues = values
    End If
End Function

'@sub-title Find a column by header without regard to case.
Private Function FindColumn(ByVal table As ListObject, ByVal header As String) As ListColumn
    Dim tableColumn As ListColumn

    For Each tableColumn In table.ListColumns
        If StrComp(Trim$(tableColumn.Name), header, vbTextCompare) = 0 Then
            Set FindColumn = tableColumn
            Exit Function
        End If
    Next tableColumn
End Function

'@sub-title Split a Choice Values cell into a list of trimmed values.
'@details
'A variable whose Choice cell is empty carries no choices, whatever the
'values cell holds. Empty pieces are dropped. The list is based at 1.
Private Function SplitChoices(ByVal choiceName As String, ByVal choiceList As String) As BetterArray
    Dim pieces() As String
    Dim pieceIndex As Long
    Dim piece As String
    Dim values As BetterArray

    Set values = NewList()

    If LenB(Trim$(choiceName)) = 0 Then
        Set SplitChoices = values
        Exit Function
    End If

    If LenB(Trim$(choiceList)) = 0 Then
        Set SplitChoices = values
        Exit Function
    End If

    pieces = Split(choiceList, CHOICE_SEPARATOR)
    For pieceIndex = LBound(pieces) To UBound(pieces)
        piece = Trim$(pieces(pieceIndex))
        If LenB(piece) > 0 Then values.Push piece
    Next pieceIndex

    Set SplitChoices = values
End Function

'@section Collecting
'===============================================================================

'@sub-title Add every variable of one disease the other has no row for.
Private Sub CollectMissing(ByRef owner As TDiseaseData, ByRef other As TDiseaseData, _
                           ByVal bundle As Checking, ByVal keyPrefix As String)
    Dim rowIndex As Long
    Dim counter As Long

    For rowIndex = 1 To owner.rowCount
        If Not IndexHas(other.index, owner.keys.Item(rowIndex)) Then
            counter = counter + 1
            bundle.Add keyPrefix & "-" & counter, _
                       VariableWhere(owner, rowIndex) & "--" & _
                       "Missing from " & other.displayName & ". Label: " & owner.labels.Item(rowIndex), _
                       checkingError
        End If
    Next rowIndex
End Sub

'@sub-title Walk the shared names and sort each into differing or equivalent.
Private Sub CollectShared()
    Dim rowIndex As Long
    Dim otherRow As Long
    Dim firstChoices As BetterArray
    Dim secondChoices As BetterArray
    Dim missingFromSecond As BetterArray
    Dim missingFromFirst As BetterArray
    Dim differingCounter As Long
    Dim equivalentCounter As Long

    For rowIndex = 1 To this.first.rowCount
        otherRow = IndexRow(this.second.index, this.first.keys.Item(rowIndex))
        If otherRow > 0 Then
            this.sharedCount = this.sharedCount + 1
            Set firstChoices = this.first.choiceSets.Item(rowIndex)
            Set secondChoices = this.second.choiceSets.Item(otherRow)
            Set missingFromSecond = ValuesMissing(firstChoices, secondChoices)
            Set missingFromFirst = ValuesMissing(secondChoices, firstChoices)

            If missingFromSecond.Length = 0 And missingFromFirst.Length = 0 Then
                this.matchingCount = this.matchingCount + 1
                equivalentCounter = equivalentCounter + 1
                this.equivalent.Add "equivalent-" & equivalentCounter, _
                                    VariableWhere(this.first, rowIndex) & "--" & _
                                    "Same choices in both diseases (" & firstChoices.Length & " values).", _
                                    checkingNote
            Else
                this.differingCount = this.differingCount + 1
                differingCounter = differingCounter + 1
                AddDifference rowIndex, differingCounter, missingFromSecond, missingFromFirst
            End If
        End If
    Next rowIndex
End Sub

'@sub-title Add the entries of one differing variable.
'@details
'One entry names the variable, then one entry per value in the first disease
'and missing from the second, then one per value in the second and missing
'from the first.
Private Sub AddDifference(ByVal rowIndex As Long, ByVal differingCounter As Long, _
                          ByVal missingFromSecond As BetterArray, ByVal missingFromFirst As BetterArray)
    Dim valueIndex As Long
    Dim keyBase As String
    Dim whereText As String

    keyBase = "differing-" & differingCounter
    whereText = VariableWhere(this.first, rowIndex)

    this.differing.Add keyBase, _
                       whereText & "--" & _
                       "The choices differ: " & missingFromSecond.Length & " value(s) only in " & _
                       this.first.displayName & ", " & missingFromFirst.Length & " value(s) only in " & _
                       this.second.displayName & ".", _
                       checkingWarning

    For valueIndex = 1 To missingFromSecond.Length
        this.differing.Add keyBase & "-first-" & valueIndex, _
                           whereText & "--" & _
                           "Value """ & missingFromSecond.Item(valueIndex) & """ is in " & _
                           this.first.displayName & " and missing from " & this.second.displayName & ".", _
                           checkingWarning
    Next valueIndex

    For valueIndex = 1 To missingFromFirst.Length
        this.differing.Add keyBase & "-second-" & valueIndex, _
                           whereText & "--" & _
                           "Value """ & missingFromFirst.Item(valueIndex) & """ is in " & _
                           this.second.displayName & " and missing from " & this.first.displayName & ".", _
                           checkingWarning
    Next valueIndex
End Sub

'@sub-title Write the counts as one bundle.
'@details
'The counts are read from the state here. The public count properties call
'Compare, and Compare is still running when this sub is called.
Private Sub CollectStatistics()
    Dim onlyFirst As Long
    Dim onlySecond As Long
    Dim share As Double

    onlyFirst = this.first.rowCount - this.sharedCount
    onlySecond = this.second.rowCount - this.sharedCount
    If this.sharedCount > 0 Then share = this.matchingCount / this.sharedCount

    this.statistics.Add "stat-count-1", _
                        this.first.displayName & "--" & this.first.rowCount & " variable(s).", checkingInfo
    this.statistics.Add "stat-count-2", _
                        this.second.displayName & "--" & this.second.rowCount & " variable(s).", checkingInfo
    this.statistics.Add "stat-shared", _
                        "Both diseases--" & this.sharedCount & " shared variable(s).", checkingInfo
    this.statistics.Add "stat-only-1", _
                        this.first.displayName & "--" & onlyFirst & " variable(s) only here.", checkingInfo
    this.statistics.Add "stat-only-2", _
                        this.second.displayName & "--" & onlySecond & " variable(s) only here.", checkingInfo
    this.statistics.Add "stat-matching", _
                        "Shared variables--" & this.matchingCount & " with matching choices.", checkingInfo
    this.statistics.Add "stat-differing", _
                        "Shared variables--" & this.differingCount & " with differing choices.", checkingInfo
    this.statistics.Add "stat-share", _
                        "Shared variables--" & Format$(share, "0.0%") & " are equivalent.", checkingInfo
End Sub

'@section Set Helpers
'===============================================================================

'@sub-title Values of the first list the second list lacks, without regard to case.
Private Function ValuesMissing(ByVal source As BetterArray, ByVal other As BetterArray) As BetterArray
    Dim missing As BetterArray
    Dim otherIndex As Collection
    Dim valueIndex As Long
    Dim value As String

    Set missing = NewList()
    Set otherIndex = New Collection

    For valueIndex = 1 To other.Length
        AddKeyIfMissing otherIndex, LCase$(CStr(other.Item(valueIndex)))
    Next valueIndex

    For valueIndex = 1 To source.Length
        value = CStr(source.Item(valueIndex))
        If Not IndexHas(otherIndex, LCase$(value)) Then
            missing.Push value
        End If
    Next valueIndex

    Set ValuesMissing = missing
End Function

'@sub-title The "Where?" text of a variable: its name and its section.
Private Function VariableWhere(ByRef data As TDiseaseData, ByVal rowIndex As Long) As String
    Dim sectionText As String

    sectionText = data.sections.Item(rowIndex)
    If LenB(sectionText) = 0 Then sectionText = "(no section)"
    VariableWhere = data.names.Item(rowIndex) & " (" & sectionText & ")"
End Function

'@section Collection Helpers
'===============================================================================

'@sub-title Add a key to a collection when it is missing.
Private Sub AddKeyIfMissing(ByVal index As Collection, ByVal key As String)
    'A duplicate key raises; the first one stays.
    On Error Resume Next
        index.Add True, key
    On Error GoTo 0
End Sub

'@sub-title Test whether a collection carries a key.
Private Function IndexHas(ByVal index As Collection, ByVal key As String) As Boolean
    'A missing key raises; that raise is the answer.
    On Error Resume Next
        index.Item key
        IndexHas = (Err.Number = 0)
        Err.Clear
    On Error GoTo 0
End Function

'@sub-title Read the row stored under a key, 0 when the key is missing.
Private Function IndexRow(ByVal index As Collection, ByVal key As String) As Long
    'A missing key raises; 0 is the answer then.
    On Error Resume Next
        IndexRow = CLng(index.Item(key))
        If Err.Number <> 0 Then
            IndexRow = 0
            Err.Clear
        End If
    On Error GoTo 0
End Function

'@sub-title A BetterArray based at 1.
'@details Every list here takes pushed values, so it is based at 1 first.
Private Function NewList() As BetterArray
    Dim list As BetterArray
    Set list = New BetterArray
    list.LowerBound = 1
    Set NewList = list
End Function

'@section Validation
'===============================================================================

Private Sub ValidateTable(ByVal table As ListObject, ByVal argumentName As String)
    If table Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, argumentName & " cannot be Nothing"
    End If
    If table.ListColumns.Count = 0 Then
        ThrowError ProjectError.InvalidArgument, argumentName & " must contain at least one column"
    End If
End Sub

Private Function DisplayNameOrDefault(ByVal candidate As String, ByVal fallback As String) As String
    DisplayNameOrDefault = Trim$(candidate)
    If LenB(DisplayNameOrDefault) = 0 Then DisplayNameOrDefault = fallback
End Function

'@section ErrorHandling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, CLASS_NAME, errMessage
End Sub
