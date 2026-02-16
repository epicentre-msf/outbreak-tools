VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist choice Class"


'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Linelist choice Class")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLChoices

'@class LLChoices
'@description
'Manages worksheet-based choice lists used for dropdowns and cascading
'selectors. Each choice is identified by a list name and carries long
'labels, short labels, and an ordering column. Built on top of a
'DataSheet for column-oriented data access.
'@depends DataSheet, CustomTable, Checking, BetterArray, ITranslationObject
'@version 1.0 (2026-02-09)


'@section State
'===============================================================================
'Backing fields and constants for the choice object.

Private Type Tchoice
    choiceDataSheet As IDataSheet
    hasChecks As Boolean
    choiceName As String
End Type


Private this As Tchoice
Private Const LISTNAMECODE As String = "list name"
Private Const TRANSLABELCODE As String = "translated label"
Private Const LABELCODE   As String = "label"
Private Const LABELSHORTCODE As String = "short label"
Private Const ORDERINLISTCODE   As String = "ordering list"

'Those are for log tracing in case
Private checkCounter As Long
Private internalChecks As IChecking

'constant / keys for tracing something to the checks class

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new LLChoices instance
'@details
'Wraps the given worksheet in a DataSheet and returns a configured
'ILLChoices instance. The default headers are "list name", "label",
'"short label", and "ordering list".
'@param choiceWorksheet Worksheet. The worksheet hosting the choices.
'@param choiceStartRow Long. First row of choice data.
'@param choiceStartColumn Long. First column of choice data.
'@param choiceName Optional String. Name for the choice object. Defaults to vbNullString.
'@return ILLChoices. The new instance.
'@export
Public Function Create(ByVal choiceWorksheet As Worksheet, _ 
                       ByVal choiceStartRow As Long, _
                       ByVal choiceStartColumn As Long, _ 
                       Optional ByVal choiceName As String = vbNullString) As _ 
                       ILLChoices

    Dim customDataSheet As IDataSheet
    Dim instance As LLChoices

    Set customDataSheet = DataSheet.Create(choiceWorksheet, choiceStartRow, choiceStartColumn)
    Set instance = New LLChoices

    Set instance.ChoiceData = customDataSheet
    instance.Name = choiceName
    Set Create = instance
End Function


'@label:choice-label
'@prop-title Resolve the label column name
'@details
'Returns the translated label column when available, otherwise falls
'back to the default label column.
'@param keepDefault Optional Boolean. When True, returns the default. Defaults to True.
'@return String. The resolved column name.
Private Property Get ChoiceLabel(Optional ByVal keepDefault As Boolean = True) As String

    If (Not keepDefault) Then
        Dim dataObject As IDataSheet
        
        Set dataObject = ChoiceData()
        If dataObject.ColumnExists(TRANSLABELCODE, matchCase:=False) Then
            ChoiceLabel = TRANSLABELCODE
            Exit Property
        End If
    End If
       
    ChoiceLabel = LABELCODE
End Property

'@section Elements
'===============================================================================

'@label:choice-data
'@prop-title Associated DataSheet object
'@return IDataSheet. The backing DataSheet.
'@export
Public Property Get ChoiceData() As IDataSheet
    Set ChoiceData = this.choiceDataSheet
End Property

'@label:choice-data-set
'@prop-title Assign the backing DataSheet
'@param customDataSheet IDataSheet. The DataSheet to assign.
Public Property Set ChoiceData(ByVal customDataSheet As IDataSheet)
    Set this.choiceDataSheet = customDataSheet
End Property

'@label:wksh
'@prop-title Worksheet hosting the choice lists
'@return Worksheet. The host worksheet.
'@export
Public Property Get Wksh() As Worksheet
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    Set Wksh = dataObject.Wksh()
End Property

'@label:start-row
'@prop-title First row of choice data
'@return Long. The start row.
Public Property Get StartRow() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()
    StartRow = dataObject.DataStartRow()
End Property

'@label:start-column
'@prop-title First column of choice data
'@return Long. The start column.
Public Property Get StartColumn() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()
    StartColumn = dataObject.DataStartColumn()
End Property


'@label:data-range
'@prop-title Data range from the backing DataSheet
'@param colName Optional String. Column name. Defaults to "__all__".
'@param includeHeaders Optional Boolean. When True, includes headers. Defaults to True.
Private Property Get DataRange( _ 
    Optional ByVal colName As String = "__all__", _
    Optional ByVal includeHeaders As Boolean = True) As Range

    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()

    'The range is obtained without case matching, because it can
    'change depending on the context (title in setup, lowercase in designer, etc.)
    Set DataRange = dataObject.DataRange( _ 
                        colName, includeHeaders, _
                         matchCase:=False)
End Property

'@section Validation
'===============================================================================

'@label:has-valid-headers
'@prop-title Check whether all required headers exist
'@return Boolean. True when all four required columns are found.
Private Property Get HasValidHeaders() As Boolean

    Dim dataObject As IDataSheet
    
    Set dataObject = ChoiceData()
    HasValidHeaders = dataObject.ColumnExists(LISTNAMECODE, matchCase:=False) And _ 
                      dataObject.ColumnExists(ChoiceLabel(), matchCase:=False) And _
                      dataObject.ColumnExists(LABELSHORTCODE, matchCase:=False) And _ 
                      dataObject.ColumnExists(ORDERINLISTCODE, matchCase:=False)

End Property

'@section Choices
'===============================================================================

'@label:name
'@prop-title Name of the choice object
'@return String. The choice object name.
'@export
Public Property Get Name() As String

    If LenB(this.choiceName) = 0 Then
        On Error Resume Next
            this.choiceName = "LLCHOICE"
            this.choiceName = Wksh.Name
        On Error GoTo 0
    End If
    
    Name = this.choiceName
End Property

'@label:name-set
'@prop-title Assign the choice object name
'@param choiceName String. The name to assign.
Public Property Let Name(ByVal choiceName As String)
    this.choiceName = choiceName
End Property

'@label:choice-exists
'@sub-title Check whether a choice exists
'@param choiceName String. Choice name to search for.
'@return Boolean. True when the choice is found.
'@export
Private Property Get ChoiceExists(ByVal choiceName As String) As Boolean
    Dim listRng As Range
    Dim Exist As Boolean
    Dim dataObject As IDataSheet


    Set dataObject = ChoiceData()
    
    'If the column listName() is not found, return false
    If Not dataObject.ColumnExists(LISTNAMECODE, matchCase:=False) Then GoTo ReturnExist

    'Empty choicename should return a False
    If LenB(choiceName) = 0 Then GoTo ReturnExist
    
    Set listRng = DataRange(LISTNAMECODE, includeHeaders:=False)
    Exist = Not ( _ 
    listRng.Find( _ 
        What:=choiceName, _ 
        lookAt:=xlWhole, _
        MatchCase:=True, _
        After:=listRng.Cells(listRng.Rows.Count, 1) _ 
        ) _ 
        Is Nothing _
    )
ReturnExist:
    ChoiceExists = Exist
End Property

'@label:sort
'@sub-title Sort choices in the worksheet
'@export
Private Sub Sort()

    Dim orderRng As Range
    Dim listRng As Range
    Dim choiceRng As Range

    'If the columns listName and Orders are not found, exit
    If Not HasValidHeaders() Then  Exit Sub

    Set listRng = DataRange(LISTNAMECODE, includeHeaders:=False)
    Set orderRng = DataRange(ORDERINLISTCODE, includeHeaders:=False)
    Set choiceRng = DataRange(includeHeaders:=True)

    choiceRng.Sort key1:=listRng, Order1:=xlAscending, key2:=orderRng, order2:=xlAscending, Header:=xlYes
End Sub

'@label:add-choice
'@sub-title Add a new choice with its categories
'@details
'Appends a new choice entry with long labels and optional short labels.
'Silently exits when the choice already exists or the data is empty.
'@param choiceName String. Name of the choice to add.
'@param longLabelData BetterArray. Long label values for the categories.
'@param shortLabelData Optional Object. BetterArray of short label values. Defaults to Nothing.
'@export
Private Sub AddChoice(ByVal choiceName As String, _
                      ByVal longLabelData As BetterArray, _ 
                      Optional ByVal shortLabelData As Object = Nothing)

    Dim listRng As Range
    Dim labelRng As Range
    Dim orderRng As Range
    Dim shortLabelRng As Range
    Dim choicewksh As Worksheet
    Dim EndRow As Long
    Dim endRowlabel As Long
    Dim counter As Long
    Dim hasShortLabel As Boolean

    hasShortLabel = False

    'Test if the choice already exists
    If ChoiceExists(choiceName) Then
         LogInfo "The choice " & choiceName & "already exists and is not added"
         Exit Sub
    End If

    'Test if categories table is empty
    If longLabelData.Length = 0 Then
        LogInfo "The provided choice " & choiceName & " does not have any categories and is not added"
        Exit Sub
    End If

    If TypeName(shortLabelData) = "BetterArray" Then
        If (longLabelData.Length <> shortLabelData.Length) Or _ 
           (longLabelData.LowerBound <> shortLabelData.LowerBound) Then
            LogInfo "long and short label tables should have the same length,  no categories added for choice " & choiceName
            Exit Sub
        Else
            hasShortLabel = True
        End If
    End If

    'Test if the headers (converted or not exists)
    If Not HasValidHeaders() Then 
        LogInfo "Some Headers of the choice " & Name() & " Object are missing", _ 
                scope:=checkingWarning
        Exit Sub
    End If

    'Get the ranges of the 
    Set listRng = DataRange(LISTNAMECODE)
    Set labelRng = DataRange(ChoiceLabel(False))
    Set orderRng = DataRange(ORDERINLISTCODE)
    Set shortLabelRng = DataRange(LABELSHORTCODE)
    Set choicewksh = Wksh()

    EndRow = StartRow + listRng.Rows.Count
    endRowlabel = StartRow + labelRng.Rows.Count

    'Add the new choice if we are really at the end of the choice Objet worksheet
    With choicewksh
        If (EndRow = endRowlabel) Then
            longLabelData.ToExcelRange .Cells(EndRow, labelRng.Column)
            If hasShortLabel Then shortLabelData.ToExcelRange .Cells(EndRow, shortLabelRng.Column)
            'Add the orders and the choice name
            For counter = 0 To (longLabelData.Length - 1)
                .Cells(EndRow + counter, listRng.Column).Value = choiceName
                .Cells(EndRow + counter, orderRng.Column).Value = counter + 1
            Next
        End If
    End With
End Sub

'@label:remove-one-choice
'@sub-title Remove a single choice by name
'@param choiceName String. The choice to remove.
Private Sub RemoveOneChoice(ByVal choiceName As String)

    Dim listRng As Range
    Dim cellRng As Range

    'exit if the choice does not exists
    If Not ChoiceExists(choiceName) Then
        LogInfo "Choice " & choiceName & " can not be deleted, it does not exists"
        ShowDebug
        Exit Sub
    End If

    If Not HasValidHeaders() Then
        LogInfo "Choice " & choiceName & " can not be deleted, some headers are missing", _ 
                 scope:=checkingWarning
        ShowDebug
        Exit Sub
    End If
    
    'Look for the choiceName in the listRange
    'Sort the actual choice object before going through
    Sort

    Set listRng = DataRange(LISTNAMECODE, includeHeaders:=False)

    'I am sure here that the choice exists, so cellRng Is not Nothing
    Set cellRng = listRng.Find( _ 
            What:=choiceName, _ 
            lookAt:=xlWhole, _
            MatchCase:=True, _
            After:=listRng.Cells(listRng.Rows.Count, 1) _ 
        )

    If (cellRng Is Nothing) Then 
        LogInfo "Choice " & choiceName & " Not deleted, something went wrong"
        ShowDebug
        Exit Sub
    End If

    Do While (cellRng.value = choiceName)
        Set cellRng = cellRng.Offset(1)
        'Move one cell upward for deleting
        cellRng.Offset(-1).EntireRow.Delete
    Loop
End Sub

'@label:remove-choice
'@sub-title Remove one or multiple choices
'@param choiceNames String. Choice names separated by sep.
'@param sep Optional String. Delimiter. Defaults to ";".
'@export
Private Sub RemoveChoice(Byval choiceNames As String, Optional ByVal sep As String = ";")
    Dim namesData As BetterArray
    Dim counter As Long
    Dim choiceValue As String

    Set namesData = New BetterArray
    namesData.Items = Split(choiceNames, sep)

    For counter = namesData.LowerBound to namesData.UpperBound
        choiceValue = namesData.Item(counter)
        RemoveOneChoice choiceName:=choiceValue
    Next
End Sub

'@label:translate
'@sub-title Translate the choice worksheet labels
'@param trads ITranslationObject. Translation provider.
'@export
Private Sub Translate(ByVal trads As ITranslationObject)

    Dim lblRng As Range

    If Not HasValidHeaders() Then
        LogInfo "Some Headers of the choice Object are missing, translation not done"
        Exit Sub
    End If

    'Translate the label
    Set lblRng = DataRange(ChoiceLabel(False))
    trads.TranslateRange lblRng

    'Translate the short label
    Set lblRng = DataRange(LABELSHORTCODE)
    trads.TranslateRange lblRng
End Sub

'@label:manage-rows
'@sub-title Add or remove rows via a single toggle
'@param del Optional Boolean. When True, removes rows. Defaults to False.
'@export
Private Sub ManageRows(Optional ByVal del As Boolean = False)
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim tabl As ICustomTable

    Set sh = Wksh()

    On Error Resume Next
    Set Lo = sh.ListObjects(1)
    On Error GoTo 0
    If Lo Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(Lo)
    If del Then
        tabl.RemoveRows totalCount:=1
    Else
        tabl.AddRows
    End If
End Sub

'@label:add-rows
'@sub-title Add rows to the choice table
'@export
Private Sub AddRows()
    ManageRows del:=False
End Sub

'@label:insert-rows
'@sub-title Insert rows at the selected position
'@details
'Mirrors the selection height inside the choices ListObject. When the
'selection is invalid the attempt is logged and ignored.
'@param targetCell Range. Selection anchoring the insertion point.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
'@export
Private Sub InsertRows(ByVal targetCell As Range, _
                       Optional ByVal insertShift As Boolean = False)
    Dim lo As ListObject
    Dim tabl As ICustomTable

    Set lo = ResolveRowTarget(targetCell, "insert")
    If lo Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(lo)
    tabl.InsertRowsAt targetCell:=targetCell, insertShift:=insertShift
End Sub

'@label:delete-rows
'@sub-title Delete rows intersecting the selection
'@param targetCell Range. Selection identifying rows to delete.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
'@export
Private Sub DeleteRows(ByVal targetCell As Range, _
                       Optional ByVal includeIds As Boolean = True, _
                       Optional ByVal forceShift As Boolean = False)
    Dim lo As ListObject
    Dim tabl As ICustomTable

    Set lo = ResolveRowTarget(targetCell, "delete")
    If lo Is Nothing Then Exit Sub

    Set tabl = CustomTable.Create(lo)
    tabl.DeleteRowsAt targetCell:=targetCell, _
                      includeIds:=includeIds, _
                      forceShift:=forceShift
End Sub

'@label:remove-rows
'@sub-title Remove empty rows from the choice table
'@export
Private Sub RemoveRows()
    ManageRows del:=True
End Sub

'@label:resolve-row-target
'@sub-title Validate a selection and resolve its ListObject
'@param targetCell Range. The user-selected cell to validate.
'@param actionLabel String. Human-readable action name used in warning messages.
'@return ListObject. The ListObject, or Nothing when invalid.
Private Function ResolveRowTarget(ByVal targetCell As Range, ByVal actionLabel As String) As ListObject
    'Validate that the provided selection belongs to the choices worksheet and intersects a listobject.
    Dim hostSheet As Worksheet
    Dim selectionSheet As Worksheet
    Dim normalizedAction As String

    normalizedAction = LCase$(actionLabel)
    Set hostSheet = Wksh()

    If targetCell Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the target selection is missing", checkingWarning
        Exit Function
    End If

    On Error Resume Next
        Set selectionSheet = targetCell.Worksheet
    On Error GoTo 0

    If selectionSheet Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the worksheet could not be resolved", checkingWarning
        Exit Function
    End If

    If Not (selectionSheet Is hostSheet) Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection is on " & selectionSheet.Name, checkingWarning
        Exit Function
    End If

    On Error Resume Next
        Set ResolveRowTarget = targetCell.ListObject
    On Error GoTo 0

    If ResolveRowTarget Is Nothing Then
        LogInfo "Unable to " & normalizedAction & " rows because the selection is outside the choices table", checkingWarning
    End If
End Function

'@label:all-choices
'@prop-title All distinct choice names
'@return BetterArray. The distinct choice names.
'@export
Private Property Get AllChoices() As BetterArray
    Dim lstChoi As BetterArray
    Dim lstRng As Range
    Dim counter As Long

    Set lstChoi = New BetterArray


    If Not HasValidHeaders() Then
        LogInfo "Some Headers of the choice Object are missing", checkingWarning
        ShowDebug
        Exit Property
    End If

    Set lstRng = DataRange(LISTNAMECODE, includeHeaders:=False)
    For counter = 1 To lstRng.Rows.Count
        If Not lstChoi.Includes(lstRng.Cells(counter, 1).Value) Then
            lstChoi.Push lstRng.Cells(counter, 1).Value
        End If
    Next
    
    Set AllChoices = lstChoi.Clone()
End Property


'@section Categories
'===============================================================================

'@label:categories
'@prop-title Categories of a choice
'@details
'Returns the long label categories for the specified choice. When
'useShortlabels is True, short labels replace long labels where
'available. Falls back to long labels for missing short entries.
'@param choiceName String. The choice to retrieve categories from.
'@param useShortlabels Optional Boolean. When True, uses short labels. Defaults to False.
'@return BetterArray. The category values.
'@export
Public Property Get Categories(ByVal choiceName As String, _
                               Optional ByVal useShortlabels As Boolean = False) _ 
                               As BetterArray
    
    'Variables
    Dim dataObject As IDataSheet 
    Dim lngcat As BetterArray 'Long and short betterarrays returned
    Dim shrtcat As BetterArray
    Dim counter As Long

    'Initializations
    Set dataObject = ChoiceData()

    'In case the choice does not exist, return empty Betterarray
    If Not ChoiceExists(choiceName) Then GoTo ReturnLongCat

    'Filter the datasheet to get the choices
    Set lngcat = dataObject.FilterData( _ 
                            varName:=LISTNAMECODE, _
                            criteriaName:=choiceName, _ 
                            returnedColumnName:=ChoiceLabel(), _
                            matchCase:=False, _
                            includeHeaders:=False)
    If lngcat Is Nothing Then Set lngcat = New BetterArray
    'Only look for shortlabels in case there is a need
    If useShortlabels Then
        'Short Categories values
        Set shrtcat = dataObject.FilterData( _ 
            varName:=LISTNAMECODE, _
            criteriaName:=choiceName, _
            returnedColumnName:=LABELSHORTCODE, _
            matchCase:=False, _
            includeHeaders:=False _
        )
        If shrtcat Is Nothing Then Set shrtcat = New BetterArray

        'In case the long and short categories don't have the same 
        'length, return the long one, because of a probable error
        If lngcat.Length <> shrtcat.Length Then GoTo ReturnLongCat
        
        'short categories and long categories have the same length, replace some
        'of the long values by the short one (non empty short ones)
        For counter = lngcat.LowerBound To lngCat.UpperBound
            If shrtcat.Item(counter) <> vbNullString Then
                lngcat.Item(counter) = shrtcat.Item(counter)
            End If
        Next
    End If

ReturnLongCat:
    If lngcat Is Nothing Then
        Dim emptyResult As BetterArray
        Set emptyResult = New BetterArray
        Set Categories = emptyResult.Clone()
    Else
        Set Categories = lngcat.Clone()
    End If
End Property


'@label:concatenate-categories
'@prop-title Concatenate all categories of a choice
'@details
'Joins the categories with a separator. Optionally translates each
'category via the supplied TranslationObject.
'@param choiceName String. The choice whose categories to concatenate.
'@param sep Optional String. Separator. Defaults to " | ".
'@param trads Optional Object. A TranslationObject. Defaults to Nothing.
'@return String. The concatenated categories.
'@export
Private Property Get ConcatenateCategories( _
        ByVal choiceName As String, _ 
        Optional ByVal sep As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String

    Dim lst As BetterArray
    Dim counter As Long

    Set lst = Categories(choiceName)

    If lst.Length = 0 Then Exit Property

    If Not trads Is Nothing Then
        If StrComp(TypeName(trads), "TranslationObject", vbTextCompare) = 0 Then
            For counter = lst.LowerBound To lst.UpperBound
                lst.Item(counter) = trads.TranslatedValue(lst.Item(counter))
            Next
        End If
    End If

    ConcatenateCategories = lst.ToString( _
                     separator:=sep, _
                     OpeningDelimiter:=vbNullString, _
                     ClosingDelimiter:=vbNullString, _
                     QuoteStrings:=False)
End Property

'@section DataExchange
'===============================================================================
'Persist the choices sheet or hydrate from external sources.

'@label:add-headers
'@sub-title Write the default headers to the choice worksheet
'@details
'Writes the four standard column headers (list name, ordering,
'long label, short label) starting at the configured origin cell.
'@export
Private Sub AddHeaders()

    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = Wksh()
    Set cellRng = sh.Cells(StartRow(), StartColumn())

    cellRng.Value = LISTNAMECODE
    cellRng.Offset(, 1).Value = ORDERINLISTCODE
    cellRng.Offset(, 2).Value = ChoiceLabel()
    cellRng.Offset(, 3).Value = LABELSHORTCODE
End Sub

'@label:export
'@sub-title Export the current choice object to a workbook
'@details
'Delegates to DataSheet.Export, including hidden names.
'@param Wkb Workbook. Destination workbook.
'@param Hide Long. Worksheet visibility constant. Defaults to xlSheetHidden.
'@export
Private Sub Export(ByVal Wkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    ChoiceData.Export Wkb, Hide:=Hide, includeNames:=True
End Sub

'@label:import
'@sub-title Import choices from a worksheet
'@details
'Optionally clears the sheet and rewrites headers before importing.
'Creates a temporary DataSheet from the source range and delegates to
'DataSheet.Import.
'@param fromWksh Worksheet. Source worksheet containing choices.
'@param fromStartRow Long. First data row in the source.
'@param fromStartcol Long. First data column in the source.
'@param clearSheet Boolean. When True, clears the sheet before import. Defaults to False.
'@export
Private Sub Import(ByVal fromWksh As Worksheet, _
                   ByVal fromStartRow As Long, _
                   ByVal fromStartcol As Long, _ 
                   Optional ByVal clearSheet As Boolean = False)
    
    Dim choiData As IDataSheet
    Dim impChoiData As IDataSheet
    Dim sh As Worksheet


    'Initialize headers before import
    Set sh = Wksh()
    Set choiData = ChoiceData()

    'clear the choices sheet before imports
    If clearSheet Then 
        sh.Cells.Clear
        AddHeaders
    End If

    'Import the data
    Set impChoiData = DataSheet.Create(fromWksh, fromStartRow, fromStartcol)

    choiData.Import impChoiData
End Sub


'@label:import-choice
'@sub-title Import choices from another ILLChoices object
'@details
'Iterates all choices in the source object, retrieves both long and
'short label categories, and adds each choice via AddChoice.
'@param choiObj ILLChoices. Source choice object to import from.
'@export
Private Sub ImportChoice(ByVal choiObj As ILLChoices)

    Dim lstCats As BetterArray
    Dim lstShortCats As BetterArray
    Dim lstChoi As BetterArray
    Dim counter As Long
    Dim choiceName As String


    'Get the list of categories
    Set lstChoi = choiObj.AllChoices()

    For counter = lstChoi.LowerBound to lstChoi.UpperBound
        choiceName = lstChoi.Item(counter)
        Set lstCats = choiObj.Categories(choiceName)
        Set lstShortCats = choiObj.Categories(choiceName, useShortLabels:=True)
        
        AddChoice choiceName:=choiceName, _
                  longLabelData:=lstCats, _
                  shortLabelData:=lstShortCats

        lstCats.Clear
        lstShortCats.Clear
    Next
End Sub

'@section Checkings
'===============================================================================
'Internal logging helpers for the checking subsystem.

'@label:log-info
'@sub-title Append a diagnostic entry to the internal checking object
'@details
'Lazily creates the Checking instance on first call. Increments the
'internal counter used as the entry key.
'@param label String. The message to record.
'@param scope Byte. Severity level. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _ 
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "Choices " & Name() & " object informations"

    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    
    'Add new informations to the checking object
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:show-debug
'@sub-title Print diagnostic summary to the Immediate window
Private Sub ShowDebug()
    Debug.Print "Check the informations of CustomTable object " & Name()
End Sub

'@label:has-checkings
'@prop-title Whether the choice object has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@details
'Returns the internal IChecking object containing all logged messages.
'Returns Nothing when no entries have been recorded.
'@return Object. An IChecking instance, or Nothing.
'@export
Private Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property

'@section Interface
'===============================================================================
'ILLChoices delegation wiring.

Private Property Get ILLChoices_ChoiceData() As IDataSheet
    Set ILLChoices_ChoiceData = ChoiceData
End Property

Private Property Get ILLChoices_Wksh() As Worksheet
    Set ILLChoices_Wksh = Wksh()
End Property

Private Property Get ILLChoices_Name() As String
    ILLChoices_Name = Name()
End Property

Private Property Get ILLChoices_ChoiceExists(ByVal choiceName As String) As Boolean
    ILLChoices_ChoiceExists = ChoiceExists(choiceName)
End Property

Private Sub ILLChoices_Sort()
    Sort
End Sub

Private Sub ILLChoices_AddChoice(ByVal choiceName As String, _
                                 ByVal longLabelData As BetterArray, _
                                 Optional ByVal shortLabelData As Object = Nothing)

    AddChoice choiceName:=choiceName, _
             longLabelData:=longLabelData, _
             shortLabelData:=shortLabelData
End Sub

Private Sub ILLChoices_RemoveChoice(ByVal choiceNames As String, _
                                    Optional ByVal sep As String = ";")
    RemoveChoice choiceNames:=choiceNames, sep:=sep
End Sub

Private Sub ILLChoices_Translate(ByVal trads As ITranslationObject)
    Translate trads
End Sub

Private Sub ILLChoices_AddRows()
    AddRows
End Sub

Private Sub ILLChoices_InsertRows(ByVal targetCell As Range, _
                                  Optional ByVal insertShift As Boolean = False)
    InsertRows targetCell, insertShift
End Sub

Private Sub ILLChoices_DeleteRows(ByVal targetCell As Range, _
                                  Optional ByVal includeIds As Boolean = True, _
                                  Optional ByVal forceShift As Boolean = False)
    DeleteRows targetCell, includeIds, forceShift
End Sub

Private Sub ILLChoices_RemoveRows()
   RemoveRows
End Sub

Private Sub ILLChoices_ManageRows(Optional ByVal del As Boolean = False)
    ManageRows del
End Sub

Private Property Get ILLChoices_AllChoices() As BetterArray
    Set ILLChoices_AllChoices = AllChoices()
End Property

Private Property Get ILLChoices_Categories(ByVal choiceName As String, _
                                        Optional ByVal useShortlabels As Boolean = False) _
                                        As BetterArray
    Set ILLChoices_Categories = Categories(choiceName, useShortlabels)
End Property

Private Property Get ILLChoices_ConcatenateCategories( _
        ByVal choiceName As String, Optional ByVal sep As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String

    ILLChoices_ConcatenateCategories = ConcatenateCategories(choiceName, sep, trads)

End Property

Private Sub ILLChoices_AddHeaders()
    AddHeaders
End Sub

Private Sub ILLChoices_Export(ByVal toWkb As Workbook, _
                              Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, Hide:=Hide
End Sub

Private Sub ILLChoices_Import(ByVal fromWksh As Worksheet, _
                              ByVal fromStartRow As Long, _
                              ByVal fromStartcol As Long, _
                              Optional ByVal clearSheet As Boolean = False)
    Import fromWksh, fromStartRow, fromStartcol, clearSheet
End Sub

Private Sub ILLChoices_ImportChoice(ByVal choiObj As ILLChoices)
    ImportChoice choiObj
End Sub

Private Property Get ILLChoices_HasCheckings() As Boolean
    ILLChoices_HasCheckings = HasCheckings()
End Property

Private Property Get ILLChoices_CheckingValues() As Object
     Set ILLChoices_CheckingValues = CheckingValues()
End Property
