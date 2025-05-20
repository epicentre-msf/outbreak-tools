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
'@IgnoreModule SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation


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

'increment by default the counter on the number of dropdowns by 1. Everytime a
'dropdown is added to the current worksheet. The counter is saved somewhere the
'first cell fo the worksheet. It starts at 1.
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
'A dropdownlist object is a worksheet with multiple listObjects of one column. Those
'listObjects are used as dropdowns. The dropdowns can be used as data validation
'or data entry tools elsewhere in the workbook. Using ListObjects of one columns
'can help with dynamically extending the dropdown entries. The prefix can help 
'with creating multiple dropdown lists in the same workbook as prefixes are used
'when creating names for references.
'
'@param outsh A Worksheet Object. The worksheet of the listobjects
'@hprefix Optional, String. The prefix used when creating the dropdown tables. Default is "dropdown_"
'@returned A DropdowLists Object
'@export
Public Function Create(ByVal outsh As Worksheet, _
                       Optional ByVal hprefix As String = "dropdown_") _
                       As IDropdownLists
    If outsh Is Nothing Then ThrowError "Dropdown worksheet not found"
  
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

'@label:Name
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
'of the first column. The default name is with choice_[headerName] in case there 
'is no prefix.
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

'@prop-title Label of the dropdown
'@details
'The dropdown label is the label written at the top of the dropdown listObject.
'It is for reading purposes and is only for user readability. In the linelist
'for example, the label is "Custom dropdown"
'@return String
Private Function CurrentDropdownLabel(ByVal labelPrefix As String) As String
    Dim dropCounter As Long
    Dim sh As Worksheet

    Set sh = Wksh()

    On Error Resume Next
        dropCounter = CLng(sh.Cells(1, 1).Value)
    On Error GoTo 0
    CurrentDropdownLabel = labelPrefix & " " & dropCounter
End Function

'@label:exists
'@prop-title Test if a dropdown is present in the dropdown worksheets
'@param listName A String. The name of dropdown searched
'@export
'@returned A Boolean
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

'@label:LabelRange
'prop-title Retrieve the Range of one specific dropdown
'@param listName A String, the name of the dropdown
'@export
'@return A Range object. The Label Range of the dropdown, generally the cell above the ListObject
Private Property Get LabelRange(ByVal listName As String) As Range

    Dim loName As String
    Dim sh As Worksheet

    If Not Exists(listName) Then Exit Property
    Set sh = Wksh()
    loName = CurrentLoName(listName)
    'Top header of the range of the listObject
    Set LabelRange = sh.ListObjects(loName).Range.Cells(0, 1)
End Property

'@label:modify
'@section Modify the dropdownlist

'@prop-title Update the counters of the dropdown
'@details
'The counter on the number of dropdown in the current dropdown worksheet
'is store in the first cell of the worksheet. The UpdateDropdownCounter sub
'adds the default increment to the counter stored in the first cell of the
'worksheet.
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

'@label:Add
'@sub-title Add table values to the dropdown list
'@details
'Given a onedimensional BetterArray, of strings, add the items of the betterArray 
'to a new listObject and create references to the listObject for future use. Label
'is added using the prefix provided by counterPrefix. For example (Custom Dropdown 1,
' Custom Dropdown 2, etc.). If the table style does not exists,
'the default table style when creating listObjects in the current excel version
'is used.
'
'@param listValues BetterArray Object. Table of string values to fill the listObject
'@param addLabel Optional, Boolean. Add label on top of the dropdown listobject. Default is False
'@param conterPrefix Optional, String. Prefix to use when adding label. Default is vbNullString
'@param tableStyle Optional, String. Table style to use for the listObject. Default is "TableStyleLight13".
'@export
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

    'Do not add a new dropdown if there is already another one
    If Exists(listName) Then
        LogInfo "Dropdown " & listName & " already exits: not added"
        ShowDebug
        Exit Sub
    End If

    loName = CurrentLoName(listName)
    headerName = CurrentHeaderName(listName)
    dropdownName = CurrentDropdownName(listName)
    colIndex = LastColumn()
    Set sh = Wksh()

    'dropdown startRow is where to add the ListObject
    With sh
        listValues.ToExcelRange .Cells(DROPDOWNSTARTROW + 1, colIndex)
        .Cells(DROPDOWNSTARTROW, colIndex).Value = headerName
        Set rng = .Range(.Cells(DROPDOWNSTARTROW, colIndex), _ 
                         .Cells(DROPDOWNSTARTROW + listValues.Length, colIndex))

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

'@label:Remove
'@sub-title Remove a dropdown that was added
'@details
'Use with caution as it will delete the listObject of the dropdown as well as 
'any references in the workbook
'@param listName A String, the dropdown name
'@return
Private Sub Remove(ByVal listName As String)
    Dim sh As Worksheet
    Dim Wkb As Workbook
    Dim rng As Range
    Dim loName As String
    Dim dropdownName As String

    If Not Exists(listName) Then
        LogInfo "Can not delete a dropdown that does not exits: " & listName & "Not found"
        ShowDebug
        Exit Sub
    End If
    
    loName = CurrentLoName(listName)
    dropdownName = CurrentDropdownName(listName)

    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range
    Set Wkb = sh.Parent

    On Error Resume Next
        Wkb.NAMES(dropdownName).Delete
    On Error GoTo 0
    
    'On the first column delete only the listObject
    If rng.Column = 1 Then
        rng.EntireColumn.Delete
    Else
        rng.EntireColumn.Delete
        rng.Offset(, 1).EntireColumn.Delete
    End If
End Sub

'@label:ClearList
'@sub-title Clear values in a dropdown table
'@details
'Given one dropdown, clear the entries and remove all values, without deleting
'the dropdown itself.
'@param listName A String. The name of the dropdown
'@export
Private Sub ClearList(ByVal listName As String)

    Dim sh As Worksheet
    Dim loName As String

    If Not Exists(listName) Then
        LogInfo "Can not clear a dropdown that does not exits: " & listName & " Not found"
        ShowDebug
        Exit Sub
    End If

    loName = CurrentLoName(listName)
    Set sh = Wksh()
    On Error Resume Next
    sh.ListObjects(loName).DataBodyRange.Delete
    On Error GoTo 0
End Sub

'@label:Update
'@sub-title Update values in a dropdown table
'@details
'Either replace all the values in a dropdown with new values, or add 
'values at the bottom of the table. In the last case, if `pasteAtBottom=True`,
'duplicate values are removed.
'@param UpdateData A BetterArray, with new values to add
'@param listName A String, the name of the dropdown to update
'@param pasteAtBottom Boolean. Update by pasting new values at the bottom of the table. Default is False
'@export
Private Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String, _ 
                   Optional ByVal pasteAtBottom As Boolean = False)
    
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim rng As Range
    Dim loRng As Range
    Dim loName As String
    Dim Lo As listObject
    Dim tabLength As Long

    If Not Exists(listName) Then 
        LogInfo "Can not Update a dropdown that does not exists: " & listName & " Not found"
        ShowDebug 
        Exit Sub
    End If

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

'@label:Sort
'@sub-title Sort a dropdown listObject
'@param listName A String, the dropdown name
'@param sortOrder A Long, the sortOrder as a `xlSortOder` enumeration. Default is xlAscending
Private Sub Sort(ByVal listName As String, _
                 Optional ByVal sortOrder As Long = xlAscending)

    Dim loName As String
    Dim rng As Range
    Dim sh As Worksheet
    
    If Not Exists(listName) Then
        LogInfo "Can not sort a dropdown that does not exits: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If
    
    loName = CurrentLoName(listName)
    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range

    On Error Resume Next
    rng.Sort key1:=rng, Order1:=sortOrder, Header:=xlYes
    On Error GoTo 0
End Sub


'@label:operate
'@section Operate with the outside world

'@label:SetValidation
'@sub-title Use a dropdown to set validation on a Range
'@details
'
'alertType can take three values: "info" for raising information in case the entered values
'are not present in the dropdown, "warning" for raising a warning and "error" for raising error.
'"error" prevent entering values outside the dropdown values.
' 
'@param cellRng A Range object, where to se the validation
'@param listName A String, the name of the dropdown to use for validation
'@param alertType Optional, A String. Se details
'@param message Optional, A String. Message to show the user in case validation failed. Default is vbNullString
'@param showError Optional, Boolean. Show validation error. Default is True.
'@export
Private Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                          Optional ByVal alertType As String = "info", _
                          Optional ByVal message As String = vbNullString, _ 
                          Optional ByVal showError As Boolean = True)

    Dim dropdownName As String

    If Not Exists(listName) Then 
        LogInfo "There is no available dropdown for " & listName
        ShowDebug
        Exit Sub
    End If

    If cellRng Is Nothing Then
        LogInfo "There is no range to set the validation of " & listName & " on."
        ShowDebug
        Exit Sub
    End If

    'dropdown names can change. I add the tag
    dropdownName = "= " & CurrentDropdownName(listName)
    
    With cellRng.validation
        .Delete
        Select Case alertType
        Case "error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=dropdownName
        Case "warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, Formula1:=dropdownName
        Case Else
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


'@label:AddReturnLink
'@sub-title Add hyperlink on a range that targets the dropdown
'@param listName A String, the name of the dropdown targeted
'@param returnRng A Range. The Range where you want to return when you clik on the label.
'@return
Private Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    
    Dim loName As String
    Dim sh As Worksheet
    Dim rng As Range
    Dim targetAddress As String
    Dim savedValue As String
    Dim savedInteriorColor As Long
    Dim savedFontColor As Long


    If Not Exists(listName) Then
        LogInfo "Can not add a returnlink on a non existing dropdown: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    If (returnRng Is Nothing) Then 
        LogInfo "Can not add a returnlink to a non existing Range: " & listName & " not found." 
        ShowDebug
        Exit Sub
    End If

    loName = CurrentLoName(listName)
    Set sh = Wksh()
    'Top header of the listobject (label Range)
    Set rng = sh.ListObjects(loName).Range.Cells(0, 1)
    targetAddress = "'" & returnRng.Parent.Name & "'" & "!" & returnRng.Address

    savedValue = rng.Value
    savedInteriorColor = rng.Interior.Color
    savedFontColor = rng.Font.Color

    'Adding hyperlinks remove the writings 
    '(sub label and so on, so need to set it back)
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

'AddHyperLink to one specific Range
'@label:AddHyperLink
'@details
'The link will target one dropdown in the dropdownlist. The targeted adress is
'the top of the dropdown, the label Range.
'@param listName A String, the name of dropdown to target
'@param linkRng A Range. Where to add the link
'@export
Private Sub AddHyperLink(ByVal listName As String, ByVal linkRng As Range)

    Dim labRng As Range
    Dim targetAddress As String
    Dim sh As Worksheet
    Dim currsh As Worksheet

    If Not Exists(listName) Then
        LogInfo "Can not add a hyperlink with a non existing dropdown: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    If (linkRng Is Nothing) Then 
        LogInfo "Can not add a hyperlink to a non existing Range: " & listName & " not found." 
        ShowDebug
        Exit Sub
    End If

    'The worksheet is the one of the linkRng
    Set sh = linkRng.Parent
    Set currsh = Wksh()
    Set labRng = LabelRange(listName)

    targetAddress = "'" & currsh.Name & "'" & "!" & labRng.Address

    'Add the HyperLink
    With sh
        .Hyperlinks.Add _ 
        Anchor:=linkRng, _ 
        Address:="", _
        SubAddress:= targetAddress, _
        TextToDisplay:= ""
    End With
End Sub


'@label:checkings
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

'@jump:exists
'Test if a dropdown is present in the worksheet
Private Property Get IDropdownLists_Exists(ByVal listName As String) As Boolean
    IDropdownLists_Exists = Exists(listName)
End Property

'@jump:Name
'Get the name of the dropdown object
Private Property Get IDropdownLists_Name() As String
    IDropdownLists_Name = Name()
End Property

'@jump:LabelRange
Private Property Get IDropdownLists_LabelRange(ByVal listName As String) As Range
    Set IDropdownLists_LabelRange = LabelRange(listName)
End Property

'@jump:modify
'Modify the dropdownlist -------------------------------------------------------

'@jump:Add
Private Sub IDropdownLists_Add(ByVal listValues As BetterArray, _ 
                               ByVal listName As String, _ 
                               Optional ByVal addLabel As Boolean = False, _
                               Optional ByVal counterPrefix As String = vbNullString, _ 
                               Optional ByVal tableStyle As String = "TableStyleLight13")

    Add listValues:=listValues, listName:=listName, _ 
        addLabel:=addLabel, counterPrefix:=counterPrefix, tableStyle:=tableStyle
End Sub

'@jump:Remove
Private Sub IDropdownLists_Remove(ByVal listName As String)
    Remove listName:=listName
End Sub

'@jump:ClearList
Private Sub IDropdownLists_ClearList(ByVal listName As String)
    ClearList listName
End Sub

'@jump:Update
Private Sub IDropdownLists_Update(ByVal UpdateData As BetterArray, _ 
                                  ByVal listName As String, _ 
                                  Optional ByVal pasteAtBottom As Boolean = False)
    Update UpdateData:=UpdateData, _ 
           listName:=listName, _ 
           pasteAtBottom:=pasteAtBottom
End Sub

'@jump:Sort
Private Sub IDropdownLists_Sort(ByVal listName As String, _
                                Optional ByVal sortOrder As Long = xlAscending)
    Sort listName, sortOrder
End Sub

'@jump:operate
'Operate with the outside world-------------------------------------------------

'@jump:SetValidation
Private Sub IDropdownLists_SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                                        Optional ByVal alertType As String = "info", _
                                        Optional ByVal message As String = vbNullString, _ 
                                        Optional ByVal showError As Boolean = True)

    SetValidation cellRng:=cellRng, listName:=listName, _ 
                  alertType:=alertType, message:=message, _ 
                  showError:=showError
End Sub

'@jump:AddReturnLink
Private Sub IDropdownLists_AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    AddReturnLink listName, returnRng
End Sub


'@jump:AddHyperLink
Private Sub IDropdownLists_AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
    AddHyperLink listName, linkRng
End Sub

'@jump:checkings
'Checkings----------------------------------------------------------------------

'@jump:HasChekings
'Test if the object has checkings
Private Property Get IDropdownLists_HasCheckings() As Boolean
 IDropdownLists_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
'Get the checking values
Private Property Get IDropdownLists_CheckingValues() As Object
    Set IDropdownLists_CheckingValues = CheckingValues()
End Property