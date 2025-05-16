VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropdownLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A wrapper for dropdowns"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("A wrapper for dropdowns")

Option Explicit
Implements IDropdownLists

Private Type TDropdownLists
    sh As Worksheet
    hprefix As String
    hasChecks As Boolean
End Type

Private this As TDropdownLists
Private Const CLASSNAME As String = "DropdownLists"
Private Const DEFAULTNAMEPREFIX As String = "choice_"

'increment by default each dropdown counter by 1. This is the number of
'dropdowns in the current dropdown worksheet.
Private Const DROPDOWNCOUNTERINCREMENT As Byte = 1 

'Where all the dropdowns should start, sometimes there could be a need
'for header for the dropdown (like in custom dropdowns, so we shift one line)
Private Const DROPDOWNSTARTROW As Integer = 2


'checking objects
Private checkCounter As Long
Private internalChecks As IChecking

'@section Instanciation


'@label:Create
'@prop-title Create a Dropdownlist Object
'
'@details
'A dropdownlist object is a worksheet with listObjects of one column. Those
'listObjects are used as dropdowns. The dropdowns can be used as data validation
' or data entry tools. Using ListObjects of one columns can help with dynamically
'extending the dropdown entries. The prefix can help with creating multiple
'dropdown lists in the same workbook as prefixes are used when creating names
'for references.
'
'@param outsh A Worksheet Object. The worksheet of the listobjects
'@hprefix Optional, String. The prefix used when creating the dropdown tables. Default is "dropdown_"
'@returned A DropdowLists Object
'@export
Public Function Create(ByVal outsh As Worksheet, _
                       Optional ByVal hprefix As String = "dropdown_") _
                       As IDropdownLists
    If outsh Is Nothing Then ThrowError "Dropdown worksheet not initialised"
  
    With New DropdownLists
        Set .Wksh = outsh
        Let .headerPrefix = hprefix
        Set Create = .Self()
    End With
End Function

'@label:elements
'@section Dropdown Elements

Public Property Get Self() As IDropdownLists
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@prop-title Worksheet of the dropdown list
Public Property Set Wksh(ByVal outsh As Worksheet)
    'Format the worksheet before creating dropdowns
    RemoveGridLines outsh
    outsh.Cells(1, 1).Font.Color = vbWhite
    outsh.Cells(1, 1).Interior.Color = vbWhite
    Set this.sh = outsh
End Property

'@prop-title Helper to remove Gridlines
'@param sh A Worksheet Object.
'@param DisplayZeros Optional, Boolean. Display zeros in the worksheet. Default is False
Private Sub RemoveGridLines(ByVal sh As Worksheet, Optional DisplayZeros As Boolean = False)
    Dim View As WorksheetView
    For Each View In sh.Parent.Windows(1).SheetViews
        If View.Sheet.Name = sh.Name Then
            View.DisplayGridlines = False
            View.DisplayZeros = DisplayZeros
            Exit Sub
        End If
    Next
End Sub

'@prop-title Header Prefix
Public Property Get headerPrefix() As String
    headerPrefix = this.hprefix
End Property

Public Property Let headerPrefix(ByVal hprefix As String)
    this.hprefix = hprefix
End Property

'@prop-title Name of the dropdown object
'@details
'The name of the dropdown Object is the same as the name of the worksheet
'where the dropdown lists are created.
'@export
'@returned As String
Private Property Get Name() As String
    Name = Wksh.Name
End Property

'@prop-title Find the lastColumn of the dropdown list
'@details
'The last column is the column where to start building new listObject. This 
'is used when adding new dropdowns.
'@returned A Long
Private Property Get LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()

    With sh
        colIndex = .Cells(DROPDOWNSTARTROW, .Columns.Count).End(xlToLeft).Column
    End With
    LastColumn = colIndex + 2
End Property

'@prop-title Name of a listobject related to one dropdown
'@param listName A String. The name of the dropdown
'@returned A String
Private Property Get CurrentListName(ByVal listName As String) As String
    CurrentListName = Replace(Application.WorksheetFunction.Trim(listName), " ", "_")
End Property

'@prop-title Name of the first column of one specific dropdown
'@param listName A String. The name of the dropdown
'@returned As String
Private Property Get CurrentHeaderName(ByVal listName As String) As String
    CurrentHeaderName = headerPrefix() & CurrentListName(listName)
End Property

'@prop-title Name of the name of dropdown used in reference
'
'@details
'In references, the name used is a combination of the prefix and the name
'of the first column. The default name is with choice_[headerName] in case there is no
'prefix.
'
'@param listName A String. The name of the dropdown
'@returned A String
Private Property Get CurrentDropdownName(ByVal listName As String) As String
    Dim headerName As String

    headerName = CurrentHeaderName(listName)
    CurrentDropdownName = IIf(headerPrefix() = vbNullString, _
                              DEFAULTNAMEPREFIX & headerName, _
                              headerName)
End Property

'@prop-title Name of the listObject related to one dropdown
'@param listName String. Dropdown name
'@returned A String. Name of the corresponding listobject
Private Property Get CurrentLoName(ByVal listName As String) As String
    CurrentLoName = "list_" & CurrentListName(listName)
End Property

'Test if a listobject exists in a sheet
Private Function Exists(ByVal listName As String) As Boolean
    Dim Lo As listObject
    Dim loName As String
    Dim sh As Worksheet
    Set sh = Wksh()
    loName = CurrentLoName(listName)
    On Error Resume Next
    Set Lo = sh.ListObjects(loName)
    On Error GoTo 0
    Exists = (Not Lo Is Nothing)
End Function

'Dropdown label using the dropdown counter
Private Function CurrentDropdownLabel(ByVal labelPrefix As String) As String
    Dim dropCounter As Long
    Dim sh As Worksheet

    Set sh = Wksh()
    dropCounter = 0
    On Error Resume Next
        dropCounter = CLng(sh.Cells(1, 1).Value)
    On Error GoTo 0
    CurrentDropdownLabel = labelPrefix & " " & dropCounter
End Function

'Update the counter of the number of dropdowns in the worksheet (separate from retrieving the current dropdown label)
Private Sub UpdateDropdownCounter()
    Dim sh As Worksheet
    Dim dropCounter As Long
    Dim counterRng As Range

    Set sh = Wksh()
    Set counterRng = sh.Cells(1, 1)

    On Error Resume Next
        dropCounter = CLng(counterRng.Value)
    On Error GoTo 0

    dropCounter = dropCounter + DROPDOWNCOUNTERINCREMENT
    counterRng.Value = dropCounter
End Sub

'Add New list to the dropdown sheet
Private Sub Add(ByVal listValues As BetterArray, ByVal listName As String, _ 
                Optional ByVal addLabel As Boolean = False, _
                Optional ByVal counterPrefix As String = vbNullString, _ 
                Optional ByVal tableStyle As String = "TableStyleLight13")

  
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim headerName As String              'header Name
    Dim loName As String                  'List object name
    Dim rng As Range
    Dim Wkb As Workbook
    Dim dropdownName As String            'dropdown name (used in named ranges)
    Dim dropdownLabel As String

    'Do not add a new dropdown if there is already another one
    If Exists(listName) Then Exit Sub

    loName = CurrentLoName(listName)
    headerName = CurrentHeaderName(listName)
    dropdownName = CurrentDropdownName(listName)
    colIndex = LastColumn()
    Set sh = Wksh()
    'dropdown startRow is where to add the ListObject
    With sh
        listValues.ToExcelRange .Cells(DROPDOWNSTARTROW + 1, colIndex)
        .Cells(DROPDOWNSTARTROW, colIndex).Value = headerName
        Set rng = .Range(.Cells(DROPDOWNSTARTROW, colIndex), .Cells(DROPDOWNSTARTROW + listValues.Length, colIndex))
        'Add A list object table
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loName
        'Add a style to the listObject
        On Error Resume Next
        .ListObjects(loName).TableStyle = tableStyle
        On Error GoTo 0
        Set Wkb = .Parent
        Wkb.Names.Add Name:=dropdownName, RefersToR1C1:="=" & loName & "[" & headerName & "]"

        'Add the label, with a prefix if needed
        If addLabel Then
            
            'Add an increment to the dropdown counter
            UpdateDropdownCounter

            .Cells(DROPDOWNSTARTROW - 1, colIndex).Value = CurrentDropdownLabel(counterPrefix)
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Interior.Color = vbWhite
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Color = RGB(91, 155, 213)
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Bold = True
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Italic = True

        End If

        'Auto fit dropdown column, and increment the dropdown counter
        rng.EntireColumn.AutoFit
    End With
End Sub

'Retrieve the range of label of a dropdown
Private Property Get LabelRange(ByVal listName As String) As Range
    Dim cellRng As Range
    Dim loName As String
    Dim sh As Worksheet

    If Not Exists(listName) Then Exit Property
    Set sh = Wksh()
    loName = CurrentLoName(listName)
    'Top header of the range of the listObject
    Set LabelRange = sh.ListObjects(loName).Range.Cells(0, 1)
End Property

'Remove the dropdown
Private Sub Remove(ByVal listName As String)
    Dim sh As Worksheet
    Dim Wkb As Workbook
    Dim rng As Range
    Dim loName As String
    Dim dropdownName As String

    If Not Exists(listName) Then Exit Sub
    
    loName = CurrentLoName(listName)
    dropdownName = CurrentDropdownName(listName)

    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range
    Set Wkb = sh.Parent

    On Error Resume Next
        Wkb.NAMES(dropdownName).Delete
    On Error GoTo 0
    
    If rng.Column = 1 Then
        rng.EntireColumn.Delete
    Else
        rng.EntireColumn.Delete
        rng.Offset(, 1).EntireColumn.Delete
    End If
End Sub

'Clear a dropdown values
Private Sub ClearList(ByVal listName As String)

    Dim sh As Worksheet
    Dim loName As String

    If Not Exists(listName) Then Exit Sub

    loName = CurrentLoName(listName)
    Set sh = Wksh()
    On Error Resume Next
    sh.ListObjects(loName).DataBodyRange.Delete
    On Error GoTo 0
End Sub

'Update a list with new values
Private Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String, _ 
                   Optional ByVal pasteAtBottom As Boolean = False)
    
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim rng As Range
    Dim loRng As Range
    Dim loName As String
    Dim Lo As listObject
    Dim tabLength As Long

    If Not Exists(listName) Then Exit Sub
    Set sh = Wksh()
    loName = CurrentLoName(listName)
    Set Lo = sh.ListObjects(loName)
    Set loRng = Lo.Range

    If (Not pasteAtBottom) Then
        On Error Resume Next
            Lo.DataBodyRange.ClearContents
        On Error GoTo 0
        Set cellRng = loRng.Cells(2, 1)
        tabLength = UpdateData.Length + 1
    Else
        tabLength = loRng.Rows.Count + UpdateData.Length
        Set cellRng = loRng.Cells(loRng.Rows.Count + 1, 1)
    End If
    
    'put the new data in the range
    UpdateData.ToExcelRange cellRng
    
    'Resize the listObject to the new Range
    Set rng = sh.Range(loRng.Cells(1, 1), loRng.Cells(tabLength, 1))
    Lo.Resize rng

    'Remove duplicates if pasteAtBottom (the loRng has changed with a resize)
    If pasteAtBottom Then
        On Error Resume Next
            Lo.Range.RemoveDuplicates Columns:=1, Header:=xlYes
        On Error GoTo 0
    End If
End Sub

'Set a validation on a range
Private Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                          Optional ByVal alertType As String = "info", _
                          Optional ByVal message As String = vbNullString, _ 
                          Optional ByVal showError As Boolean = True)

    Dim dropdownName As String

    If Not Exists(listName) Then ThrowError "There is no available dropdown for " & listName

    'dropdown names can change. I add the tag
    dropdownName = "= " & CurrentDropdownName(listName)
    
    With cellRng.validation
        .Delete
        Select Case alertType
        Case "error"                             '"error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=dropdownName
        Case "warning"                           '"warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, Formula1:=dropdownName
        Case Else            'for all the others, add an information alert
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
            Operator:=xlBetween, Formula1:=dropdownName
        End Select

        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .errorTitle = vbNullString
        .InputMessage = vbNullString
        .errorMessage = Left(message, 254)
        .ShowInput = showError
        .ShowError = showError
    End With
End Sub

'Add Return link on a label (to return back to the variable)
Private Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    
    Dim loName As String
    Dim sh As Worksheet
    Dim rng As Range
    Dim targetAddress As String
    Dim savedValue As String
    Dim savedInteriorColor As Long
    Dim savedFontColor As Long


    If Not Exists(listName) Then Exit Sub
    If (returnRng Is Nothing) Then Exit Sub

    loName = CurrentLoName(listName)
    Set sh = Wksh()
    'Top header of the listobject (label Range)
    Set rng = sh.ListObjects(loName).Range.Cells(0, 1)
    targetAddress = "'" & returnRng.Parent.Name & "'" & "!" & returnRng.Address

    savedValue = rng.Value
    savedInteriorColor = rng.Interior.Color
    savedFontColor = rng.Font.Color

    'Adding hyperlinks remove the writings (sub label and so on, so need to set it back)
    With sh
        .Hyperlinks.Add _ 
        Anchor:=rng, _ 
        Address:="", _
        SubAddress:= targetAddress, _
        TextToDisplay:= ""
    End With

    'Adding links remove the formatting, set it back
    With rng
        .Value = savedValue
        .Font.UnderLine = xlUnderlineStyleNone
        .Interior.Color = savedInteriorColor
        .Font.Color = savedFontColor
        .Font.Bold = True
        .Font.Italic = True
    End With
End Sub

'Sort a dropdown listObject
Private Sub Sort(ByVal listName As String, _
                 Optional ByVal sortOrder As Long = xlAscending)

    Dim loName As String
    Dim rng As Range
    Dim sh As Worksheet
    
    If Not Exists(listName) Then Exit Sub
    
    loName = CurrentLoName(listName)
    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range

    On Error Resume Next
    rng.Sort key1:=rng, Order1:=sortOrder, Header:=xlYes
    On Error GoTo 0
End Sub


'@label:Checkings
'@section Checkings

'@sub-title Add elements to the checking object for future notes
'@param Label String, the message to add.
'@param scope Byte. The scope of the checking object.
'@export
Private Sub LogInfo(ByVal label As String, _ 
                    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "DropdownList " & Name() & " object trace"

    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    
    'Add new informations to the checking object
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:HasChecking
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

Private Sub ShowDebug()
    Debug.Print "Check the informations of choice object " & Name()
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'@jump:Create
'@section Interface

'@jump:elements
'DropdownList Elements --------------------------------------------------------
Private Property Get IDropdownLists_Exists(ByVal listName As String) As Boolean
    IDropdownLists_Exists = Exists(listName)
End Property




Private Sub IDropdownLists_Add(ByVal listValues As BetterArray, ByVal listName As String, _ 
                            Optional ByVal addLabel As Boolean = False, _
                            Optional ByVal counterPrefix As String = vbNullString, _ 
                            Optional ByVal tableStyle As String = "TableStyleLight13")

    Add listValues:=listValues, listName:=listName, _ 
        addLabel:=addLabel, counterPrefix:=counterPrefix, tableStyle:=tableStyle
End Sub

Private Sub IDropdownLists_Remove(ByVal listName As String)
    Remove listName:=listName
End Sub

Private Property Get IDropdownLists_LabelRange(ByVal listName As String) As Range
    Set IDropdownLists_LabelRange = LabelRange(listName)
End Property

Private Sub IDropdownLists_SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                                        Optional ByVal alertType As String = "info", _
                                        Optional ByVal message As String = vbNullString, _ 
                                        Optional ByVal showError As Boolean = True)

    SetValidation cellRng:=cellRng, listName:=listName, _ 
                  alertType:=alertType, message:=message, _ 
                  showError:=showError
End Sub


Private Sub IDropdownLists_Update(ByVal UpdateData As BetterArray, _ 
                                  ByVal listName As String, _ 
                                  Optional ByVal pasteAtBottom As Boolean = False)
    Update UpdateData:=UpdateData, _ 
           listName:=listName, _ 
           pasteAtBottom:=pasteAtBottom
End Sub

Private Sub IDropdownLists_ClearList(ByVal listName As String)
    ClearList listName
End Sub

Private Sub IDropdownLists_Sort(ByVal listName As String, _
                                Optional ByVal sortOrder As Long = xlAscending)
    Sort listName, sortOrder
End Sub

Private Sub IDropdownLists_AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    AddReturnLink listName, returnRng
End Sub
