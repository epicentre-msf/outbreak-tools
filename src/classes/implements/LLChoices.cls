VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Linelist choice Class")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLChoices

'CHOICES OBJECT

Private Type Tchoice
    choiceDataSheet As IDataSheet
    listNameValue As String
    labelValue As String
    labelShortValue As String
    orderValue As String
    hasTrads As Boolean
    trads As ITranslation
    hasChecks As Boolean
    choiceName As String
End Type


Private this As Tchoice
Private Const LISTNAMECODE As String = "list name"
Private Const LABELCODE   As String = "label"
Private Const LABELSHORTCODE As String = "short label"
Private Const ORDERINLISTCODE   As String = "ordering list"

'Those are for log tracing in case
Private checkCounter As Integer
Private internalChecks As IChecking

'constant / keys for tracing something to the checks class

'@label:Create
'@section Instanciation
' Create a choice object
'
'@description
'A choice Object is a worksheet which contains the definitions of the choices
'used in dropdowns and multiple dropdown.
'
'@details
'The choice object is bulit on top of a DataSheet, so it needs the startrow and
' the starcolumn. In some cases, there is a translation object 
'that goes with the choices. Translation are used to translate the headers of
'choice worksheet. The default headers of a choice object are "list name", "label",
'"short label", "ordering list". In case the headers are not in English, you
'should use a translation when creating the object.
'
'@param choiceWorksheet Worksheet object. The worksheet where the choices are
'@param choiceStartRow Long. start row of the choices
'@param trads Translation Object. Translation used for choices headers
'@export
'@return A Choice Object
Public Function Create(ByVal choiceWorksheet As Worksheet, _ 
                       ByVal choiceStartRow As Long, _
                       ByVal choiceStartColumn As Long, _ 
                       Optional ByVal trads As ITranslation = Nothing,
                       Optional ByVal choiceName As String = vbNullString) As _ 
                       ILLChoices

    Dim customDataSheet As IDataSheet

    Set customDataSheet = DataSheet.Create(choiceWorksheet, choiceStartRow, choiceStartColumn)
    
    With New LLChoices
        Set .ChoiceData = customDataSheet
        If  (Not trads Is Nothing) Then
            If TypeName(trads) = "Translation" Then
                'Set the translation object if required
                .HasTranslation = True
                Set .TradObject = trads
            End If
        End If
        .Name = choiceName
        Set Create = .Self
    End With
End Function


'Self
Public Property Get Self() As ILLChoices
    Set Self = Me
End Property

'@section Translations of headers

'@prop-title Set translation property
Public Property Let HasTranslation(ByVal hastrads As Boolean)
    this.hasTrads = hastrads
End Property


'@label:HasTranslation
'@prop-title Check if the choices have a translation
'@return A Boolean
'@export
Public Property Get HasTranslation() As Boolean
    HasTranslation = this.hasTrads
End Property


'@label:TradObject
'@prop-title Translation object used for translating the headers
'@return Translation Object for the headers. Could be Nothing if no translation available.
'@export
Public Property Get TradObject() As ITranslation
    Set TradObject = this.trads
End Property

Public Property Set TradObject(ByVal trads As ITranslation)
    Set this.trads = trads
End Property


'@note Sub/properties for Getting the headers ---

'Depending on whether there is a translation or NOT attached to the choice object
'Get the label, the list name, the short label and other values
Private Property Get ListName() As String
    Dim trads As ITranslation

    If Not HasTranslation() Then this.listNameValue = LISTNAMECODE

    If this.listNameValue = vbNullString Then
        Set trads = TradObject()
        this.listNameValue = trads.TranslatedValue(LISTNAMECODE)
    End If

    ListName = this.listNameValue
End Property

Private Property Get ChoiceLabel() As String
    Dim trads As ITranslation

    If Not HasTranslation() Then this.labelValue = LABELCODE

    If this.labelValue = vbNullString Then
        Set trads = TradObject()
        this.labelValue = trads.TranslatedValue(LABELCODE)
    End If
    
    ChoiceLabel = this.labelValue
End Property

Private Property Get LabelShort() As String
    Dim trads As ITranslation

    If Not HasTranslation() Then this.labelShortValue = LABELSHORTCODE

    If this.labelShortValue = vbNullString Then
        Set trads = TradObject()
        this.labelShortValue = trads.TranslatedValue(LABELSHORTCODE)
    End If

    LabelShort = this.labelShortValue
End Property

Private Property Get OrderingList() As String
    Dim trads As ITranslation

    If Not HasTranslation() Then this.orderValue = ORDERINLISTCODE

    If this.orderValue = vbNullString Then
        Set trads = TradObject()
        this.orderValue = trads.TranslatedValue(ORDERINLISTCODE)
    End If

    OrderingList = this.orderValue
End Property

'@section Associated DataSheet Object properties

'@label:ChoiceData
'@prop-title Get the associated DataSheet Object
'@return A DataSheet Object
'@export
Private Property Get ChoiceData() As IDataSheet
    Set ChoiceData = this.choiceDataSheet
End Property

'@prop-title Set the associated DataSheet object
'@param customDataSheet DataSheet object
Private Property Set ChoiceData(ByVal customDataSheet As IDataSheet)
    Set this.choiceDataSheet = customDataSheet
End Property

'@label:Wksh
'@prop-title Worksheet of the choice
'@return A Worksheet Object
'@export
Public Property Get Wksh() As Worksheet
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    Set Wksh = dataObject.Wksh()
End Property

'@prop-title StartRow
'@return A Long. The row where the choices start
Public Property Get StartRow() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()
    StartRow = dataObject.StartRow()
End Property

'@prop-title StartColumn
'@return A Long. The column where the choices start
Public Property Get StartColumn() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()
    StartColumn = dataObject.StartColumn()
End Property


'@prop-title Get one data Range of the DataSheet
'
'@param colName Optional, A String. The column name of the range to return. 
'       Default is all column names
'@param includeHeaders. A Boolean. Include Headers in the Data Range?
Private Property Get DataRange( _ 
    Optional ByVal colName As String = "__all__", _
    Optional ByVal includeHeaders As Boolean = True) As Range

    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()

    'The range is obtained without case matching, because it can
    'change depending on the context (title in setup, lowercase in designer, etc.)
    Set DataRange = dataObject.DataRange(_ 
                        colName, includeHeaders, _
                         matchCase:=False)
End Property

'@prop-title Make sure all the headers exists (with or without translation)
'@return Boolean
Private Property Get HasValidHeaders() As Boolean

    Dim dataObject As IDataSheet
    
    Set dataObject = ChoiceData()
    HasValidHeaders = dataObject.ColumnExists(ListName()) And _ 
                      dataObject.ColumnExists(ChoiceLabel()) And _
                      dataObject.ColumnExists(LabelShort()) And _ 
                      dataObject.ColumnExists(OrderingList())

End Property

'@section Working with Choices

'@label:Name
'@prop-title Name of the choice object
'@return String
'@export
Public Property Get Name() As String
    Name = this.choiceName
End Property

Public Property Let Name(ByVal choiceName As String)
    this.choiceName = choiceName
End Property

'@label:ChoiceExists
'@prop-title Test if a choice Exists
'@param choiceName String. One choice in the worksheet
'@export
Private Property Get ChoiceExists(ByVal choiceName As String) As Boolean
    Dim listRng As Range
    Dim exist As Boolean
    Dim dataObject As IDataSheet

    Exist = False

    Set dataObject = ChoiceData()
    
    'If the column listName() is not found, return false
    If Not dataObject.ColumnExists(listName()) Then GoTo ReturnExist

    'Empty choicename should return a False
    If choiceName = vbNullString Then GoTo ReturnExist
    
    Set listRng = DataRange(listName(), includeHeaders:=False)
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

'@label:Sort
'@sub-title Sort Choices in the choice Worksheet
'@export
Private Sub Sort()

    Dim orderRng As Range
    Dim listRng As Range
    Dim choiceRng As Range

    Set dataObject = ChoiceData()

    'If the columns listName and Orders are not found, exit
    If Not HasValidHeaders() Then  Exit Sub

    Set listRng = DataRange(ListName(), includeHeaders:=False)
    Set orderRng = DataRange(OrderingList(), includeHeaders:=False)
    Set choiceRng = DataRange(includeHeaders:=False)

    choiceRng.Sort key1:=listRng, Order1:=xlAscending, key2:=orderRng, order2:=xlAscending
End Sub

'@label:AddChoice
'@sub-title Add new choices to the choice worksheet
'@param choiceName String the name of the choice to add
'@param longLabelData BetterArray. A table of all the long labels values of the choice
'@param shortLabelData A BetterArray (Optional). A table of all the short labels values of the choice
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
        Else
            hasShortLabel = True
        Exit Sub
        End If
    End If

    'Test if the headers (converted or not exists)
    If Not HeadersExists() Then 
        LogInfo "Some Headers of the choice " & Name() & " Object are missing", _ 
                scope:=checkingWarning
        Exit Sub
    End If

    'Get the ranges of the 
    Set listRng = DataRange(ListName())
    Set labelRng = DataRange(ChoiceLabel())
    Set orderRng = DataRange(OrderingList())
    Set shortLabelRng = DataRange(LabelShort())
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

'@sub-title Remove one choice
'@param choiceName String, the choice to remove
Private Sub RemoveOneChoice(ByVal choiceName As String)

    Dim listRng As Range
    Dim cellRng As Range

    'exit if the choice does not exists
    If Not ChoiceExists(choiceName) Then
        LogInfo "Choice " & choiceName & " can not be deleted, it does not exists"
        ShowDebug
        Exit Sub
    End If

    If Not HeadersExists() Then
        LogInfo "Choice " & choiceName " can not be deleted, some headers are missing", _ 
                 scope:=checkingWarning
        ShowDebug
        Exit Sub
    End If
    
    'Look for the choiceName in the listRange
    'Sort the actual choice object before going through
    Sort

    Set listRng = DataRange(ListName(), includeHeaders:=False)

    'I am sure here that the choice exists, so cellRng Is not Nothing
    Set cellRng = listRng.Find( _ 
            What:=choiceName, _ 
            lookAt:=xlWhole, _
            MatchCase:=True, _
            After:=listRng.Cells(listRng.Rows.Count, 1) _ 
        )

    If (cellRng Is Nothing) Then 
        LogInfo "Choice " & choiceName " Not deleted, something went wrong"
        ShowDebug
        Exit Sub
    End If

    Do While (cellRng.value = choiceName)
        Set cellRng = cellRng.Offset(1)
        'Move one cell upward for deleting
        cellRng.Offset(-1).EntireRow.Delete
    Loop
End Sub

'@label:RemoveChoice
'@sub-title Remove one or multiple choices at once, separated by a separator
'@param choiNames String. A list of choice names separated by separator.
'@param sep String, separator for choices. Default is semi-colon ";"
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

'@label:Translate
'@sub-title Translate the current Choice worksheet
'@param trads ITranslation object used for translations
'@export
Private Sub Translate(ByVal trads As ITranslation)

    Dim lblRng As Range

    If Not HeadersExists() Then
        LogInfo "Some Headers of the choice Object are missing, translation not done"
        Exit Sub
    End If

    'Translate the label
    Set lblRng = DataRange(ChoiceLabel())
    trads.TranslateRange lblRng

    'Translate the short label
    Set lblRng = DataRange(LabelShort())
    trads.TranslateRange lblRng
End Sub


'@sub-title Resize the custom table in the worksheet
'@param del Boolean. Delete or Add new rows
'@param nbRows Long, number of rows to add. Default is 5
'@param insertShift Boolean. Insert new lines when Adding new rows (Default is False)
'Add / Remove Rows
Private Sub Resize(Optional ByVal del As Boolean = False, _ 
                   Optional ByVal nbRows As Long = 5, _ 
                   Optional ByVal insertShift As Boolean = False)
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
        tabl.RemoveRows
    Else
        tabl.AddRows nbRows:=nbRows, insertShift:=insertShift
    End If
End Sub

'@label:AddRows
'@sub-title Add Rows to the choice associated custom table
'@param nbRows Optional, Long. Number of Rows to add. Default is 5
'@param insertShift Optional, Boolean. Insert new lines when Adding new rows. 
'       Default is False
'@export
Private Sub AddRows(Optional ByVal nbRows As Long = 5, _ 
                    Optional ByVal insertShift As Boolean = False)
 Resize del:=False, insertShift:=insertShift, nbRows:=nbRows
End Sub

'@label:RemoveRows
'@sub-title Remove Rows in the choice associated custom table
'@export
Private Sub RemoveRows()
    Resize del:=True
End Sub

'@label:AllChoices
'@prop-title Get all the choices in a BetterArray
'@return A BetterArray
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

    Set lstRng = DataRange(ListName(), includeHeaders:=False)
    For counter = 1 To lstRng.Rows.Count
        If Not lstChoi.Includes(lstRng.Cells(counter, 1).Value) Then
            lstChoi.Push lstRng.Cells(counter, 1).Value
        End If
    Next
    
    Set AllChoices = lstChoi.Clone()
End Property


'@section Working with Categories

'@label:Categories
'@prop-title Categories of a choice
'
'@description
'Given a choice, get the list of all categories as a BetterArray Object.
'If useShortLabels is True, the short label categories are used instead.
'
'In case there is no short label for one specific category, long labels are used
'instead. So for one specific choice, you can end up with a mix of short/long
'labels if short labels are missing for some values of long labels.
'
'@pram choiceName String. The name of choice where we want categories from
'@param useShortlabels Boolean. Return the long or the Short labels. Default is False
'@return A BetterArray with all the categories
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
    Set lngcat = New BetterArray
    Set shrtcat = New BetterArray

    'In case the choice does not exist, return empty Betterarray
    If Not ChoiceExists(choiceName) Then GoTo ReturnLongCat

    'Filter the datasheet to get the choices
    Set lngcat = dataObject.FilterData( _ 
                            varName:=listName(), _
                            criteriaName:=choiceName, _ 
                            returnedColumnName:=ChoiceLabel(), _
                            includeHeaders:=False)
    'Only look for shortlabels in case there is a need
    If useShortlabels Then
        'Short Categories values
        Set shrtcat = dataObject.FilterData( _ 
            varName:=listName(), _
            criteriaName:=choiceName, _
            returnedColumnName:=LabelShort(), _
            includeHeaders:=False _
        )

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
    Set Categories = lngcat.Clone()
End Property


'@label:ConcatenateCategories
'@prop-title Concatenate all the categories of one choice
'
'@description
'Retrieve the categories, concatenated with a separator. The default separator 
'is " | ". It is possible to get categories translated, if a translation object
'is provided. In case the translation object is nothing, it is defaulted to the
'simple categories.
'
'
'@param choiceName String. The choice
'@param sep String. The separator to use for concatenation
'@param trads A Translation Object. Optional, see description.
'@return String
'@export A concatenated String, translated if a translation object is provided
Private Property Get ConcatenateCategories( _ 
        ByVal choiceName As String, _ 
        Optional ByVal sep As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String

    Dim lst As BetterArray
    Dim counter As Long

    Set lst = Categories(choiceName)

    If lst.Length = 0 Then Exit Property

    If (TypeName(trads) = "Translation") Then
        For counter = lst.LowerBound To lst.UpperBound
            lst.Item(counter) = trads.TranslatedValue(lst.Item(counter))
        Next
    End If

    ConcatenateCategories = lst.ToString( _
                     separator:=sep, _
                     OpeningDelimiter:=vbNullString, _
                     ClosingDelimiter:=vbNullString, _
                     QuoteStrings:=False)
End Property

'@section Export and Import

'@label:AddHeaders
'@sub-title Add the headers to the current choice object
'@export
Private Sub AddHeaders()

    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = Wksh()
    Set cellRng = sh.Cells(StartRow(), StartColumn())

    cellRng.Value = ListName()
    cellRng.Offset(, 1).Value = OrderingList()
    cellRng.Offset(, 2).Value = ChoiceLabel()
    cellRng.Offset(, 3).Value = LabelShort()
End Sub

'@label:Export
'@sub-title Export the current choice object to a workbook
'@param wkb Workbook. Output workbook
'@param Hide. Optional, Long. Sheet hidden attribute. Default is xlSheetHidden. 
'@export
Private Sub Export(ByVal Wkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Dim choiData As IDataSheet
    Set choiData = ChoiceData()
    ChoiceData.Export Wkb, Hide:=Hide
End Sub

'@label:Import
'@sub-title Import elements of a worksheet into the current choice object
'@param fromWksh A Worksheet containing the choices to import
'@param fromStartRow, fromStartCol. Long, start row and column of the choice in 
'        the imported worksheet
'@export
Private Sub Import(ByVal fromWksh As Worksheet, _ 
                   ByVal fromStartRow As Long, _
                   ByVal fromStartcol As Long)
    
    Dim choiData As IDataSheet
    Dim impChoiData As IDataSheet
    Dim sh As Worksheet
    Dim cellRng As Range


    'Initialize headers before import
    Set sh = Wksh()
    Set choiData = ChoiceData()

    'clear the choices sheet before imports
    sh.Cells.Clear
    AddHeaders

    'Import the data
    Set impChoiData = DataSheet.Create(fromWksh, fromStartRow, fromStartcol)

    choiData.Import impChoiData
End Sub


'@label:ImportChoice
'@sub-title Import from a choice object
'@param choiObj A Choice object to import 
'@export
Private Sub ImportChoice(ByVal choiObj As ILLChoices)

    Dim lstCats As BetterArray
    Dim lstShortCats As BetterArray
    Dim lstChoi As BetterArray
    Dim counter As Long
    Dim choiceName As String


    Set lstCats = New BetterArray
    Set lstChoi = New BetterArray

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

'@section checkings

'@sub-title Add elements to the checking object for future notes
'
'@param label. String, the message to add.
'@param scope, Byte. The scope of the checking object
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

Private Sub ShowDebug()
    Debug.Print "Check the informations of choice object " & Name()
End Sub

'@label:HasCheckings
'@prop-title Check if the object has checkings 
'@export
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Get the cheking values
'@export
Private Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property

'@jump:Create
'@section Interfaces

'Translations of headers -------------------------------------------------------

'@jump:HasTranslation
'Check if the choices have translation
Private Property Get ILLChoices_HasTranslation() As Boolean
    ILLChoices_HasTranslation = HasTranslation()
End Property

'@jump:TradObject
'Translation object used for translating the headers
Private Property Get ILLChoices_TradObject() As ITranslation
    Set ILLChoices_TradObject = TradObject()
End Property

'Associated DataSheet Object properties ----------------------------------------

'@jump:ChoiceData
'Associated DataSheet Object
Private Property Get ILLChoices_ChoiceData() As IDataSheet
    Set ILLChoices_ChoiceData = ChoiceData
End Property

'@jump:Wksh
'Worksheet of the choices
Private Property Get ILLChoices_Wksh() As Worksheet
    Set ILLChoices_Wksh = Wksh()
End Property

'Working with choices ----------------------------------------------------------

'@jump:Name
'Name of the choice object
Private Property Get ILLChoices_Name() As String
    ILLChoices_Name = Name()
End Property

'@jump:ChoiceExists
'Test if a choice Exists
Private Property Get ILLChoices_ChoiceExists(ByVal choiceName As String) As Boolean
    ILLChoices_ChoiceExists = ChoiceExists(choiceName)
End Property

'@jump:Sort
'Sort the choices in the choice worksheet
Private Sub ILLChoices_Sort()
    Sort
End Sub

'@jump:AddChoice
'Add new choices to the choice worksheet
Private Sub ILLChoices_AddChoice(ByVal choiceName As String, _ 
                                 ByVal longLabelData As BetterArray, _
                                 Optional ByVal shortLabelData As Object = Nothing)

    AddChoice choiceName:=choiceName, _ 
             longLabelData:=longLabelData, _ 
             shortLabelData:=shortLabelData
End Sub

'@jump:RemoveChoice
'Remove one or multiple choices at once, separated by a separator
Private Sub ILLChoices_RemoveChoice(ByVal choiceNames As String, _ 
                                    Optional ByVal sep As String = ";")
    RemoveChoice choiceNames:=choiceNames, sep:=sep
End Sub

'@jump:Translate
'Translate the current choice worksheet
Private Sub ILLChoices_Translate(ByVal trads As ITranslation)
    Translate trads
End Sub

'@jump:AddRows
'Add Rows to the choice associated custom table
Private Sub ILLChoices_AddRows(Optional ByVal nbRows As Long = 5, _ 
                              Optional ByVal insertShift As Boolean = False)
    AddRows nbRows:=nbRows, insertShift:=insertShift
End Sub


'@jump:RemoveRows
'Remove Rows in the choice associated custom table
Private Sub ILLChoices_RemoveRows()
   RemoveRows
End Sub

'@jump:AllChoices
'Get all the choices in a BetterArray
Private Property Get ILLChoices_AllChoices() As BetterArray
    Set ILLChoices_AllChoices = AllChoices()
End Property


'Working with categories -------------------------------------------------------

'@jump:Categories
'Get the categories of one choice
Private Property Get ILLChoices_Categories(ByVal choiceName As String, _ 
                                        Optional ByVal useShortlabels As Boolean = False) _ 
                                        As BetterArray
    Set ILLChoices_Categories = Categories(choiceName, useShortlabels)
End Property


'@jump:ConcatenateCategories
'Concatenate all categories of one choice
Private Property Get ILLChoices_ConcatenateCategories( _ 
        ByVal choiceName As String, Optional ByVal sep As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String

    ILLChoices_ConcatenateCategories = ConcatenateCategories(choiceName, sep, trads)

End Property

'Export and Import -------------------------------------------------------------

'@jump:AddHeaders
'add the headers to the choice object
Private Sub ILLChoices_AddHeaders()
    AddHeaders
End Sub

'@jump:Export
'export the choice object to another workbook
Private Sub ILLChoices_Export(ByVal toWkb As Workbook, _ 
                              Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, Hide:=Hide
End Sub

'@jump:Import
'import the choice object
Private Sub ILLChoices_Import(ByVal fromWksh As Worksheet, _ 
                              ByVal fromStartRow As Long, _ 
                              ByVal fromStartcol As Long)
    Import fromWksh, fromStartRow, fromStartcol
End Sub

'@jump:ImportChoice
'Import from a choice object
Private Sub ILLChoices_ImportChoice(ByVal choiObj As ILLChoices)
    ImportChoice choiObj
End Sub

'Checkings ---------------------------------------------------------------------

'@jump:HasCheckings
'Check if the current choice has some checkings
Private Property Get ILLChoices_HasCheckings() As Boolean
    ILLChoices_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
'Get the checking object
Private Property Get ILLChoices_CheckingValues() As Object
     Set ILLChoices_CheckingValues = CheckingValues()
End Property
