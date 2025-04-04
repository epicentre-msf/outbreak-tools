VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Choices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist choice class"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Linelist choice class")
'@IgnoreModule ConstantNotUsed

Option Explicit
Implements IChoices


Private Type Tchoice
    choiceDataSheet As IDataSheet
    trads As ITranslation
    listNameValue As String
    labelValue As String
    labelShortValue As String
    orderValue As String
End Type

Private this As Tchoice

Private Const LISTNAMECODE As String = "ListName"
Private Const LABELCODE   As String = "Label"
Private Const LABELSHORTCODE As String = "ShortLabel"
Private Const ORDERINLISTCODE   As String = "OrderingList"
Private Const DEFAULTCHOICESEPARATOR As String = " | "

'Instanciation
Public Function Create(ByVal choiceWorksheet As Worksheet, _ 
                       ByVal choiceStartRow As Long, _ 
                       ByVal choiceStartColumn As Long, _
                       ByVal trads As ITranslation) As IChoices

    Dim customDataSheet As IDataSheet
    Set customDataSheet = DataSheet.Create(choiceWorksheet, choiceStartRow, choiceStartColumn)

    With New Choices
        Set .ChoiceData = customDataSheet
        Set .TradsObject = trads
        Set Create = .Self
    End With
End Function

'Self
Public Property Get Self() As IChoices
    Set Self = Me
End Property

'Worksheet
Public Property Get Wksh() As Worksheet
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    Set Wksh = dataObject.Wksh()
End Property

Public Property Get TradObject() As ITranslation
    Set TradObject = this.trads
End Property

Public Property Set TradObject(ByVal trads As ITranslation)
    Set this.trads = trads
End Property

Private Property Get ListName() As String
    Dim trads As ITranslation
    If this.listNameValue = VbNullString Then 
        Set trads = TradObject()
        this.listNameValue = trads.TranslatedValue(LISTNAMECODE)
    End If

    ListName = this.listNameValue
End Property

Private Property Get Label() As String
    Dim trads As ITranslation

    If this.labelValue = vbNullString Then
        Set trads = TradObject()
        this.labelValue = trads.TranslatedValue(LABELCODE)
    End If

End Property

Private Property Get LabelShort() As String
    Dim trads As ITranslation

    If this.labelShortValue = vbNullString Then
        Set trads = TradObject()
        this.labelShortValue = trads.TranslatedValue(LABELSHORTCODE)
    End If

    LabelShort = this.labelShortValue
End Property

Private Property Get OrderingList() As String
    Dim trads As ITranslation

    If this.orderValue = vbNullString Then
        Set trads = TradObject()
        this.orderValue = trads.TranslatedValue(ORDERINLISTCODE)
    End If

    OrderingList = this.orderValue
End Property

'StartRow
Private Property Get StartRow() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    StartRow = dataObject.StartRow()
End Property

'StartColumn
Private Property Get StartColumn() As Long
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData
    StartColumn = dataObject.StartColumn()
End Property

Private Sub AddHeaders()

    Dim sh As Worksheet
    Dim cellRng As Range

    Set sh = Wksh()
    Set cellRng = sh.Cells(StartRow(), StartColumn())

    cellRng.Value = ListName()
    cellRng.Offset(, 1).Value = OrderingList()
    cellRng.Offset(, 2).Value = Label()
    cellRng.Offset(, 3).Value = LabelShort()
End Sub

Private Property Get ChoiceData() As IDataSheet
    Set ChoiceData = this.choiceDataSheet
End Property

Public Property Set ChoiceData(ByVal customDataSheet As IDataSheet)
    Set this.choiceDataSheet = customDataSheet
End Property

'Get the data range
Private Property Get DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = True) As Range
    Dim dataObject As IDataSheet
    Set dataObject = ChoiceData

    Set DataRange = dataObject.DataRange(colName, includeHeaders)
End Property

'Test if a choice Exists in the choice sheet
Private Property Get ChoiceExists(ByVal choiceName As String) As Boolean
    Dim listRng As Range
    Dim exist As Boolean
    Dim dataObject As IDataSheet

    exist = False

    Set dataObject = ChoiceData
    If dataObject.ColumnExists(ListName()) Then
        Set listRng = DataRange(ListName(), includeHeaders:=False)
        exist = Not (listRng.Find(What:=choiceName, lookAt:=xlWhole, _
                                  MatchCase:=True, After:=listRng.Cells(listRng.Rows.Count, 1)) Is Nothing)
    End If
    ChoiceExists = exist
End Property

'Get the categories
Private Property Get Categories(ByVal choiceName As String) As BetterArray
    Dim counter As Long
    Dim rngName As Range
    Dim rngLabel As Range
    Dim lst As BetterArray
    Set lst = New BetterArray

    If ChoiceExists(choiceName) Then
        Set rngName = DataRange(ListName(), includeHeaders:=False)
        Set rngLabel = DataRange(Label(), includeHeaders:=False)
        For counter = 1 To rngName.Rows.Count
            If rngName.Cells(counter, 1).Value = choiceName Then _
             lst.Push rngLabel.Cells(counter, 1).Value
        Next
    End If
    Set Categories = lst.Clone()
End Property


Private Property Get PasteTranslatedCategories(ByVal choiceName As String, ByVal trads As ITranslation, _
                                    Optional ByVal separator As String = " | ") As String
    Dim lst As BetterArray
    Dim counter As Long
    Dim pasteValue As String

    If Not ChoiceExists(choiceName) Then Exit Property
    Set lst = Categories(choiceName)
    pasteValue = trads.TranslatedValue(lst.Item(lst.LowerBound))
    For counter = lst.LowerBound + 1 To lst.UpperBound
        pasteValue = pasteValue & separator & trads.TranslatedValue(lst.Item(counter))
    Next

    PasteTranslatedCategories = pasteValue
    
End Property


Private Property Get PasteCategories(ByVal choiName As String, Optional ByVal separator As String = " | ") As String

    Dim lst As BetterArray

    If Not ChoiceExists(choiName) Then Exit Property
    Set lst = Categories(choiName)

    PasteCategories = lst.ToString( _ 
                     Separator:=separator, _ 
                     OpeningDelimiter:=vbNullString, _ 
                     ClosingDelimiter:=vbNullString, _
                     QuoteString:=False)
End Property

'Sort the Choices sheet
Private Sub Sort()

    Dim dataObject As IDataSheet
    Dim orderRng As Range
    Dim listRng As Range
    Dim choiceRng As Range

    Set dataObject = ChoiceData

    If dataObject.ColumnExists(ListName()) And dataObject.ColumnExists(OrderingList()) Then
        Set listRng = DataRange(ListName(), includeHeaders:=False)
        Set orderRng = DataRange(OrderingList(), includeHeaders:=False)
        Set choiceRng = DataRange(includeHeaders:=False)

        choiceRng.Sort key1:=listRng, order1:=xlAscending, key2:=orderRng, order2:=xlAscending
    End If
End Sub

Private Property Get AllChoices() As BetterArray
    Dim lstChoi As BetterArray
    Dim lstRng As Range
    Dim counter As Long 'counter for the number of rows of one range

    Set lstChoi = New BetterArray
    Set lstRng = DataRange(ListName(), includeHeaders:=False)
    For counter = 1 To lstRng.Rows.Count
        If Not lstChoi.Includes(lstRng.Cells(counter, 1).Value) Then _
            lstChoi.Push lstRng.Cells(counter, 1).Value
    Next
    
    Set AllChoices = lstChoi.Clone()
End Property

'Add New Categories
Private Sub AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray)

    Dim listRng As Range
    Dim labelRng As Range
    Dim orderRng As Range
    Dim choicewksh As Worksheet
    Dim EndRow As Long
    Dim endRowlabel As Long
    Dim counter As Long
    Dim checkCategories As BetterArray

    If CategoriesData.Length = 0 Then Exit Sub

    'Test if the choice already exists
    Set checkCategories = Categories(choiceName)
    If checkCategories.Length > 0 Then Exit Sub  'Will probably trace log.

    Set listRng = DataRange(ListName())
    Set labelRng = DataRange(Label())
    Set orderRng = DataRange(OrderingList())
    Set choicewksh = Wksh()

    EndRow = StartRow + listRng.Rows.Count
    endRowlabel = StartRow + labelRng.Rows.Count

    With choicewksh
        If (EndRow = endRowlabel) Then
            CategoriesData.ToExcelRange .Cells(EndRow, labelRng.Column)
            For counter = 0 To (CategoriesData.Length - 1)
                .Cells(EndRow + counter, listRng.Column).Value = choiceName
                .Cells(EndRow + counter, orderRng.Column).Value = counter + 1
            Next
        End If
    End With
End Sub

'Add / Remove Rows
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
        tabl.RemoveRows
    Else
        tabl.AddRows
    End If
End Sub

'Properties
Private Property Get IChoices_Wksh() As Worksheet
    Set IChoices_Wksh = Wksh()
End Property

Private Property Get IChoices_StartRow() As Long
    IChoices_StartRow = StartRow()
End Property

Private Property Get IChoices_StartColumn() As Long
    IChoices_StartColumn = StartColumn()
End Property

Private Property Get IChoices_ChoiceData() As IDataSheet
    Set IChoices_ChoiceData = ChoiceData
End Property

Private Property Get IChoices_DataRange(Optional ByVal colName As String = "__all__", Optional ByVal includeHeaders As Boolean = True) As Range
    Set IChoices_DataRange = DataRange(colName, includeHeaders)
End Property

Private Property Get IChoices_Categories(ByVal choiceName As String) As BetterArray
    Set IChoices_Categories = Categories(choiceName)
End Property

'@Ignore ProcedureNotUsed
Private Property Get IChoices_ChoiceExists(ByVal choiceName As String) As Boolean
    IChoices_ChoiceExists = ChoiceExists(choiceName)
End Property

'Void methods
Private Sub IChoices_Sort()
    Sort
End Sub

Private Sub IChoices_AddChoice(ByVal choiceName As String, ByVal CategoriesData As BetterArray)
    AddChoice choiceName:=choiceName, CategoriesData:=CategoriesData
End Sub

Private Sub IChoices_AddRows()
    ManageRows
End Sub

Private Sub IChoices_RemoveRows()
    ManageRows del:=True
End Sub

Private Property Get IChoices_AllChoices() As BetterArray
    Set IChoices_AllChoices = AllChoices()
End Property

Private Property Get IChoices_PasteTranslatedCategories(ByVal choiceName As String, ByVal trads As ITranslation, _
                                    Optional ByVal separator As String = DEFAULTCHOICESEPARATOR) As String
    IChoices_PasteTranslatedCategories = PasteTranslatedCategories(choiceName, trads, separator)
End Property


Private Property Get IChoices_PasteCategories(ByVal choiceName As String, _ 
                                              Optional ByVal separator As String = DEFAULTCHOICESEPARATOR) As String
    IChoices_PasteCategories = PasteCategories(choiceName, separator)
End Property

Private Sub IChoices_AddHeaders()
    AddHeaders
End Sub