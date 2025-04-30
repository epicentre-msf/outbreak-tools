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
                       Optional ByVal trads As ITranslation = Nothing) As _ 
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

'@prop-title Check if the choices have a translation
'@return A Boolean
'@export
Public Property Get HasTranslation() As Boolean
    HasTranslation = this.hasTrads
End Property

'@prop-title Setting the ITranslation object used for translating the headers
'@return ITranslation Object
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

Private Property Get Label() As String
    Dim trads As ITranslation

    If Not HasTranslation() Then this.labelValue = LABELCODE

    If this.labelValue = vbNullString Then
        Set trads = TradObject()
        this.labelValue = trads.TranslatedValue(LABELCODE)
    End If
    
    Label = this.labelValue
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

'@prop-title Get the associated DataSheet Object
Private Property Get ChoiceData() As IDataSheet
    Set ChoiceData = this.choiceDataSheet
End Property

'@prop-title Set the associated DataSheet object
'@param customDataSheet DataSheet object
Private Property Set ChoiceData(ByVal customDataSheet As IDataSheet)
    Set this.choiceDataSheet = customDataSheet
End Property

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
'@export
Public Property Get StartRow() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()
    StartRow = dataObject.StartRow()
End Property

'@prop-title StartColumn
'@return A Long. The column where the choices start
'@export
Public Property Get StartColumn() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()
    StartColumn = dataObject.StartColumn()
End Property

'@prop-title Get one data Range of the DataSheet
'
'@param colName A String. The column name of the range
'@param includeHeaders. A Boolean. Include Headers in the Data Range?
'@return A Range Object
'@export
Private Property Get DataRange( _ 
    Optional ByVal colName As String = "__all__", _
    Optional ByVal includeHeaders As Boolean = True) As Range

    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData()

    Set DataRange = dataObject.DataRange(colName, includeHeaders)
End Property

'@prop-title Make sure all the headers exists (with or without translation)
Private Property Get HasValidHeaders() As Boolean

    Dim dataObject As IDataSheet
    
    Set dataObject = ChoiceData()
    HasValidHeaders = dataObject.ColumnExists(ListName()) And _ 
                      dataObject.ColumnExists(Label()) And _
                      dataObject.ColumnExists(LabelShort()) And _ 
                      dataObject.ColumnExists(OrderingList())

End Property

'@section Working with Choices

'@prop-title Test if a choice Exists
'
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
            LogInfo "long and short label tables should have the same name,  no categories added for choice " & choiceName
        Else
            hasShortLabel = True
        Exit Sub
        End If
    End If

    'Test if the headers (converted or not exists)
    If Not HeadersExists() Then 
        LogInfo "Some Headers of the choice Object are missing"
        Exit Sub
    End If

    'Get the ranges of the 
    Set listRng = DataRange(ListName())
    Set labelRng = DataRange(Label())
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

'@sub-title Translate the current Choice worksheet
'@param trads ITranslation object used for translations
Private Sub Translate(ByVal trads As ITranslation)

    Dim lblRng As Range

    If Not HeadersExists() Then
        LogInfo "Some Headers of the choice Object are missing, translation not done"
        Exit Sub
    End If

    'Translate the label
    Set lblRng = DataRange(Label())
    trads.TranslateRange lblRng

    'Translate the short label
    Set lblRng = DataRange(LabelShort())
    trads.TranslateRange lblRng
End Sub


'@note Working with Categories ===

'Get the categories of a choice
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
Public Property Get Categories( _ 
                                ByVal choiceName As String, _ 
                                Optional ByVal useShortlabels As Boolean = False _ 
                    ) As BetterArray
    
    'Variables
    Dim dataObject As IDataSheet 
    Dim lngcat As BetterArray 'Long and short betterarrays returned
    Dim shrtcat As BetterArray
    Dim counter As Long

    'Initializations
    Set dataObject = ChoiceData
    Set lngcat = New BetterArray
    Set shrtcat = New BetterArray

    'In case the choice does not exist, return empty Betterarray
    If Not ChoiceExists(choiceName) Then GoTo ReturnLongCat

    'Filter the datasheet to get the choices
    Set lngcat = dataObject.FilterData( _ 
                            varName:=listName(), _
                            criteriaName:=choiceName, _ 
                            returnedColumnName:=label(), _
                            includeHeaders:=False)
    'Only look for shortlabels in case there is a need
    If useShortlabels Then
        'Short Categories values
        Set shrtcat = dataObject.FilterData( _ 
            varName:=listName(), _
            criteriaName:=choiceName, _
            returnedColumnName:=labelShort(), _
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

'Concatenate all the categories of one choice, with a separator
'
'@param choiName String. The choice
'@param separator String. The separator to use for concatenation
'@return String
'@export
Private Property Get ConcatenateCategories( _ 
        ByVal choiName As String, _ 
        Optional ByVal separator As String = " | ") As String

    Dim lst As BetterArray

    Set lst = Categories(choiName)

    ConcatenateCategories = lst.ToString( _
                     separator:=separator, _
                     OpeningDelimiter:=vbNullString, _
                     ClosingDelimiter:=vbNullString, _
                     QuoteStrings:=False)
End Property


'Sub / Properties for checking and tracing ===
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'
Private Sub AddTrace(ByVal label As String, Optional ByVal scope As Byte = checkingNote)

    Const CHECKNAME = "Choices object trace"
    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(CHECKNAME)
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    If Not this.hasChecks Then this.hasChecks = True
End Sub

Private Property Get TraceValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then 
        Exit Property
    End If

    Set TraceValues = internalChecks
End Property



Private Sub AddHeaders()

    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = Wksh()
    Set cellRng = sh.Cells(StartRow(), StartColumn())

    cellRng.Value = listName()
    cellRng.Offset(, 1).Value = OrderingList()
    cellRng.Offset(, 2).Value = Label()
    cellRng.Offset(, 3).Value = LabelShort()

End Sub


'Get the categories

'Import and export Sub
Private Sub Export(ByVal Wkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Dim choiData As IDataSheet
    Set choiData = ChoiceData()
    ChoiceData.Export Wkb, Hide:=Hide
End Sub

Private Sub Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Dim choiData As IDataSheet
    Dim impChoiData As IDataSheet
    Dim sh As Worksheet
    Dim cellRng As Range
    Set choiData = ChoiceData()

    'Initialize headers before import
    Set sh = Wksh()
    'clear the choices sheet before imports
    sh.Cells.Clear
    Set cellRng = sh.Cells(StartRow, StartColumn)
    cellRng.Value = LISTNAME
    cellRng.Offset(, 1).Value = ORDER
    cellRng.Offset(, 2).Value = LABEL
    cellRng.Offset(, 3).Value = LABELSHORT

    'Import the data
    Set impChoiData = DataSheet.Create(fromWksh, fromStartRow, fromStartcol)

    choiData.Import impChoiData
End Sub



'@note Interface Properties ====

'Translations of headers ---
Private Property Get ILLChoices_HasTranslation() As Boolean
    ILLChoices_HasTranslation = HasTranslation()
End Property

Private Property Get ILLChoices_TradObject() As ITranslation
    Set ILLChoices_TradObject = TradObject()
End Property

'Elements of the choice Object ---

Private Property Get ILLChoices_Wksh() As Worksheet
    Set ILLChoices_Wksh = Wksh()
End Property

Private Property Get ILLChoices_StartRow() As Long
    ILLChoices_StartRow = StartRow()
End Property

Private Property Get ILLChoices_StartColumn() As Long
    ILLChoices_StartColumn = StartColumn()
End Property

Private Property Get ILLChoices_ChoiceData() As IDataSheet
    Set ILLChoices_ChoiceData = ChoiceData
End Property

Private Property Get ILLChoices_DataRange( _ 
                                        Optional ByVal colName As String = "__all__", _ 
                                        Optional ByVal includeHeaders = True) _ 
                                        As Range
    Set ILLChoices_DataRange = DataRange(colName, includeHeaders)
End Property

'Working with choices ---

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


Private Property Get ILLChoices_Categories( _ 
                                        ByVal choiceName As String, _ 
                                        Optional ByVal useShortlabels As Boolean = False) _ 
                                        As BetterArray
    Set ILLChoices_Categories = Categories(choiceName, useShortlabels)
End Property


Private Property Get ILLChoices_HasCheckings() As Boolean
    ILLChoices_HasCheckings = HasCheckings()
End Property

Private Property Get ILLChoices_TraceValues() As Object
     Set ILLChoices_TraceValues = TraceValues()
End Property



'Void methods




'Export and Import
Private Sub ILLChoices_Export(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
    Export toWkb, Hide:=Hide
End Sub

Private Sub ILLChoices_Import(ByVal fromWksh As Worksheet, ByVal fromStartRow As Long, ByVal fromStartcol As Long)
    Import fromWksh, fromStartRow, fromStartcol
End Sub



