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
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument,   ExcelMemberMayReturnNothing,  SuperfluousAnnotationArgument, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation
'@depends HiddenNames, BetterArray, Checking

Option Explicit
Implements IDropdownLists

Private Type TDropdownLists
    sh As Worksheet
    hprefix As String
    hasChecks As Boolean
    initialised As Boolean
    wbStored As IHiddenNames
    shStored As IHiddenNames
End Type

Private this As TDropdownLists
Private Const CLASSNAME As String = "DropdownLists"
Private Const DEFAULTNAMEPREFIX As String = "choice_"
Private Const WORKBOOKCOUNTERNAME As String = "__Var__WBDROPCOUNTER"  'Workbook-level dropdown counter
Private Const WORKSHEETCOUNTERNAME As String = "__Var__SHDROPCOUNTER" 'Worksheet-level dropdown counter
Private Const DROPDOWNINITIALSEDNAME As String = "__Var__DROPINIT"
Private Const LISTNAMECOLUMNHEADER As String =  "__listName__"
Private Const LONAMECOLUMNHEADER As String = "__loName__"
Private Const WORKSHEETLANGUAGENAME As String = "__Var__LANGNAME"

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
'===============================================================================
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
'@depends Checkings
'@export
Public Function Create(ByVal outsh As Worksheet, Optional ByVal hprefix As String = "dropdown_") As IDropdownLists

    If outsh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Dropdown worksheet not found"
    End If

    With New DropdownLists
        Set .Wksh = outsh
        Let .headerPrefix = hprefix
        Set Create = .Self()
    End With
End Function

'@label:elements
'@section Dropdown Elements
'===============================================================================

Public Property Get Self() As IDropdownLists
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property


'@prop-title Worksheet of the dropdown list
Public Property Set Wksh(ByVal outsh As Worksheet)
    
    Dim store As IHiddenNames
    Set store = HiddenNames.Create(outsh)
    this.initialised = (store.ValueAsString(DROPDOWNINITIALSEDNAME) = "Yes")

    If (Not this.initialised) Then

       'Format the worksheet before creating dropdowns
        RemoveGridLines outsh

        outsh.Range("A:D").Font.Color = vbWhite
        outsh.Range("A:D").Interior.Color = vbWhite

        outsh.Cells(DROPDOWNSTARTROW, 3).Value = LISTNAMECOLUMNHEADER
        outsh.Cells(DROPDOWNSTARTROW, 4).Value = LONAMECOLUMNHEADER

        this.initialised = True
        store.EnsureName DROPDOWNINITIALSEDNAME, "Yes", HiddenNameTypeString
    End If

    Set this.sh = outsh
    Set this.shStored = Nothing
    Set this.wbStored = Nothing
End Property


Private Property Get AllDropdowns() As BetterArray
    Dim sh As Worksheet
    Dim temp As BetterArray
    Dim headerCell As Range
    Dim cursor As Range

    Set temp = New BetterArray
    temp.LowerBound = 1

    Set sh = Wksh()
    If sh Is Nothing Then
        Set AllDropdowns = temp
        Exit Property
    End If

    Set headerCell = sh.Cells(DROPDOWNSTARTROW, 3)
    Set cursor = sh.Cells(sh.Rows.Count, 3).End(xlUp)

    If cursor.Row <= headerCell.Row Then
        Set AllDropdowns = temp
        Exit Property
    End If

    Do While cursor.Row > headerCell.Row
        If LenB(Trim$(CStr(cursor.Value))) > 0 Then
            temp.Push CStr(cursor.Value)
        End If
        Set cursor = cursor.Offset(-1)
    Loop

    temp.Reverse
    Set AllDropdowns = temp.Clone()
End Property


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


'@section Workbook counter management
'===============================================================================
Private Property Get InternalStore(Optional ByVal sheetLevel As Boolean = False) As IHiddenNames
    Dim targetSheet As Worksheet
    Dim targetWb As Workbook

    If sheetLevel Then
        If this.shStored Is Nothing Then
            Set targetSheet = Wksh
            If targetSheet Is Nothing Then
                ThrowError ProjectError.ObjectNotInitialized, "DropdownLists requires a worksheet before accessing sheet storage."
            End If
            Set this.shStored = HiddenNames.Create(targetSheet)
        End If
        Set InternalStore = this.shStored
    Else
        If this.wbStored Is Nothing Then
            Set targetSheet = Wksh
            If targetSheet Is Nothing Then
                ThrowError ProjectError.ObjectNotInitialized, "DropdownLists requires a worksheet before accessing workbook storage."
            End If
            Set targetWb = targetSheet.Parent
            Set this.wbStored = HiddenNames.Create(targetWb)
        End If
        Set InternalStore = this.wbStored
    End If
End Property


Private Sub EnsureWorkbookCounter()
    Dim store As IHiddenNames

    Set store = InternalStore(False)
    If Not store.HasName(WORKBOOKCOUNTERNAME) Then
        store.EnsureName WORKBOOKCOUNTERNAME, 1, HiddenNameTypeLong
    End If
End Sub

Private Function WorkbookCounterValue() As Long
    Dim store As IHiddenNames

    EnsureWorkbookCounter
    Set store = InternalStore(False)
    WorkbookCounterValue = store.ValueAsLong(WORKBOOKCOUNTERNAME, 1)
End Function

Private Sub SetWorkbookCounter(ByVal newValue As Long)
    Dim store As IHiddenNames

    EnsureWorkbookCounter
    Set store = InternalStore(False)
    store.SetValue WORKBOOKCOUNTERNAME, newValue
End Sub

Private Sub IncrementWorkbookCounter()
    Dim nextValue As Long
    nextValue = WorkbookCounterValue() + DROPDOWNCOUNTERINCREMENT
    SetWorkbookCounter nextValue
End Sub

Private Sub DecrementWorkbookCounter()
     Dim nextValue As Long
    nextValue = WorkbookCounterValue() - DROPDOWNCOUNTERINCREMENT
    SetWorkbookCounter nextValue
End Sub
'@section Worksheet Counter
'===============================================================================

Private Sub EnsureWorksheetCounter()
    Dim store As IHiddenNames

    Set store = InternalStore(True)
    If Not store.HasName(WORKSHEETCOUNTERNAME) Then
        store.EnsureName WORKSHEETCOUNTERNAME, 0, HiddenNameTypeLong
    End If
End Sub


Private Function SheetCounterValue() As Long
    Dim store As IHiddenNames

    EnsureWorksheetCounter
    Set store = InternalStore(True)
    SheetCounterValue = store.ValueAsLong(WORKSHEETCOUNTERNAME, 0)
End Function

Private Sub SetSheetCounter(ByVal newValue As Long)
    Dim store As IHiddenNames

    EnsureWorksheetCounter
    If newValue < 0 Then newValue = 0

    Set store = InternalStore(True)
    store.SetValue WORKSHEETCOUNTERNAME, newValue
End Sub

Private Sub IncrementSheetCounter()
    Dim nextValue As Long
    nextValue = SheetCounterValue() + DROPDOWNCOUNTERINCREMENT
    SetSheetCounter nextValue
End Sub

Private Sub DecrementSheetCounter()
    Dim nextValue As Long
    nextValue = SheetCounterValue() - DROPDOWNCOUNTERINCREMENT
    SetSheetCounter nextValue
End Sub


'@prop-title Update the counters of the dropdown
'@details
'The counter on the number of dropdown in the current dropdown worksheet
'is store in the first cell of the worksheet. The UpdateDropdownCounter sub
'adds the default increment to the counter stored in the first cell of the
'worksheet.
Private Sub IncrementCounters()
    IncrementSheetCounter
    IncrementWorkbookCounter
End Sub

Private Sub DecrementCounters()
    DecrementSheetCounter
    DecrementWorkbookCounter
End Sub

'@section Private class Helpers
'===============================================================================

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


'@prop-title Find the lastColumn of the dropdown list
'@details
'The last column is the column where to start building new listObject. This 
'is used when adding new dropdowns.
'@returned A Long
Private Function LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()

    With sh
        colIndex = .Cells(DROPDOWNSTARTROW, .Columns.Count).End(xlToLeft).Column
    End With

    LastColumn = colIndex + 2
End Function


'@fun-title Name of a listobject related to one dropdown
'@param listName A String. The name of the dropdown
'@returned A String
Private Function SanitizedListName(ByVal listName As String) As String
    SanitizedListName = Replace(Application.WorksheetFunction.Trim(listName), " ", "_")
End Function

'@fun-title Name of the first column of one specific dropdown
'@param listName A String. The name of the dropdown
'@returned As String
Private Function ListObjectHeaderName(ByVal listName As String) As String
    ListObjectHeaderName = headerPrefix() & SanitizedListName(listName)
End Function

'@prop-title Name of the name of dropdown used in reference
'
'@details
'In references, the name used is a combination of the prefix and the name
'of the first column. The default name is with choice_[headerName] in case there 
'is no prefix.
'
'@param listName A String. The name of the dropdown
'@returned A String
Private Property Get CorrespondingWorkbookName(ByVal listName As String) As String
    Dim headerName As String

    headerName = ListObjectHeaderName(listName)
    CorrespondingWorkbookName = IIf(LenB(headerPrefix()) = 0, DEFAULTNAMEPREFIX & headerName, headerName)
End Property

'@prop-title Name of the listObject related to one dropdown
'@param listName String. Dropdown name
'@returned A String. Name of the corresponding listobject
Private Property Get ListObjectName(ByVal listName As String) As String
    Dim loName As String

    If Not Exists(listName) Then
        loName = "list_" & SanitizedListName(listName) & WorkbookCounterValue()
        AddLoName listName, loName
    Else
        loName = RetrieveLoName(listName)
    End If

    ListObjectName = loName
End Property

'Add and Retrieve the listObject Name for a specific listName
'@sub-title Add listObject Name
'@param lst A String, the name of the dropdown
'@param  lo A String, the name of the corresponding listObject
Private Sub AddLoName(ByVal listName As String, ByVal loName As String)

    Dim cellRng As Range
    Dim sh As Worksheet

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    Do While (LenB(CStr(cellRng.Value)) <> 0)
        Set cellRng = cellRng.Offset(1)
    Loop

    cellRng.Value = listName
    cellRng.Offset(, 1).Value = loName
End Sub


'@prop-title Retrieve a list object name of a dropdown
'@param  lst A String, the name of the dropdown
'@return A String
Private Property Get RetrieveLoName(ByVal lst As String) As String
    Dim cellRng As Range
    Dim sh As Worksheet
    Dim nbElmts As Long

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    For nbElmts = 0 To Length()
        If cellRng.Value = lst Then
            RetrieveLoName = cellRng.Offset(, 1).Value
            Exit Property
        End If
        Set cellRng = cellRng.Offset(1)
    Next
End Property

'@sub-title Remove the list object name
Private Sub RemoveLoName(ByVal lst As String)
    Dim cellRng As Range
    Dim sh As Worksheet
    Dim nbElmts As Long

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    For nbElmts = 0 To Length()
        If cellRng.Value = lst Then
            cellRng.ClearContents
            cellRng.Offset(, 1).ClearContents
            Exit Sub
        End If
        Set cellRng = cellRng.Offset(1)
    Next
End Sub


'@label:Length
'@prop-title Total Number of dropdowns in the worksheet
'@return
'@export
Private Property Get Length() As Long
    Length = SheetCounterValue()
End Property


'@prop-title Label of the dropdown
'@details
'The dropdown label is the label written at the top of the dropdown listObject.
'It is for reading purposes and is only for user readability. In the linelist
'for example, the label is "Custom dropdown"
'@return String
Private Function CurrentDropdownLabel(ByVal labelPrefix As String) As String
    Dim dropCounter As Long

    dropCounter = SheetCounterValue()
    CurrentDropdownLabel = labelPrefix & " " & dropCounter
End Function

'@label:exists
'@prop-title Test if a dropdown is present in the dropdown worksheets
'@param listName A String. The name of dropdown searched
'@export
'@returned A Boolean
Private Function Exists(ByVal listName As String) As Boolean

    Dim sh As Worksheet
    Dim cellRng As Range
    Dim count As Long

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    For count = 0 To Length()
        If cellRng.Value = listName Then
            Exists = True
            Exit Function
        End If
        Set cellRng = cellRng.Offset(1)
    Next

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
    loName = ListObjectName(listName)
    'Top header of the range of the listObject
    Set LabelRange = sh.ListObjects(loName).Range.Cells(0, 1)
End Property


'@label:Values
'@prop-title Retrieve values of a dropdown
'@param listName A String, the name of the dropdown
'@export
'@return A BetterArray Object, elements of the dropdown
Private Property Get Values(ByVal listName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray

    Dim lstVal As BetterArray
    Dim sh As Worksheet
    Dim rng As Range
    Dim loName As String
    Dim lo As ListObject

    Set lstVal = New BetterArray
    lstVal.LowerBound = 1

    If Not Exists(listName) Then
        LogInfo "Can not retrive values of a non existing dropdown: " & listName & " not found."
        ShowDebug
    Else
        Set sh = Wksh()
        loName = ListObjectName(listName)
        Set Lo = sh.ListObjects(loName)
        Set rng = Lo.Range

        If Not includeHeaders Then
            Set rng = sh.Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
        End If

        lstVal.FromExcelRange rng, DetectLastColumn:=False, DetectLastRow:=False
    End If

    Set Values = lstVal.Clone()
End Property

'@label:modify
'@section Modify the dropdownlist
'===============================================================================

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
    Dim store As IHiddenNames
    Dim dropdownLo As ListObject

    'Do not add a new dropdown if there is already another one
    If Exists(listName) Then
        LogInfo "Dropdown " & listName & " already exits: not added"
        Exit Sub
    End If

    If (listName = LISTNAMECOLUMNHEADER) Then
        LogInfo "Dropdown " & listName & " is an invalid name already used for programming purposes, please change it."
        Exit Sub
    End If

    loName = ListObjectName(listName)
    headerName = ListObjectHeaderName(listName)

    colIndex = LastColumn()
    Set sh = Wksh()
    
    'Add an increment to the dropdown counter
    IncrementCounters


    'dropdown startRow is where to add the ListObject
    With sh
        listValues.ToExcelRange .Cells(DROPDOWNSTARTROW + 1, colIndex)
        .Cells(DROPDOWNSTARTROW, colIndex).Value = headerName
        Set rng = .Range(.Cells(DROPDOWNSTARTROW, colIndex), .Cells(DROPDOWNSTARTROW + listValues.Length, colIndex))

        'Add A list object table
        '@Ignore ReadOnlyPropertyAssignment
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loName

        'Add the label, with a prefix if needed
        If addLabel Then
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Value = CurrentDropdownLabel(counterPrefix)
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Interior.Color = vbWhite
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Color = RGB(91, 155, 213)
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Bold = True
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Italic = True
        End If
        
        Set dropdownLo = .ListObjects(loName)
    End With

    'Add a style to the listObject
    On Error Resume Next
    dropdownLo.TableStyle = tableStyle
    On Error GoTo 0
    

    'Add the corresponding name at workbook level
    Set store = InternalStore(False)
    store.SetListObjectHeader CorrespondingWorkbookName(listName), dropdownLo, headerName

    'Auto fit dropdown column, and increment the dropdown counter
    rng.EntireColumn.AutoFit
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
    Dim rng As Range
    Dim loName As String
    Dim store As IHiddenNames

    If Not Exists(listName) Then
        LogInfo "Can not delete a dropdown that does not exits: " & listName & "Not found"
        Exit Sub
    End If


    Set store = InternalStore(False)
    store.RemoveName CorrespondingWorkbookName(listName)

    'Remove the listObject
    loName = ListObjectName(listName)
    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range

    'Important to delete the range at the left before deleting the
    'range itself
    rng.Offset(, 1).EntireColumn.Delete
    rng.EntireColumn.Delete

    'Remove the name from the stored name in the worksheet
    RemoveLoName listName
    
    'Decrement counters
    DecrementCounters
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

    loName = ListObjectName(listName)
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
    Optional ByVal pasteAtBottom As Boolean = False, _
    Optional ByVal removeDuplicates As Boolean = True)

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

    If  (UpdateData.Length = 0) Then
        LogInfo "Can not Update because updated data is empty"
        Exit Sub
    End If

    Set sh = Wksh()
    loName = ListObjectName(listName)
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
    If pasteAtBottom Or removeDuplicates Then
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

    loName = ListObjectName(listName)
    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range

    On Error Resume Next
    rng.Sort key1:=rng, Order1:=sortOrder, Header:=xlYes
    On Error GoTo 0
End Sub

'@label:Translate
Private Sub Translate(ByVal trads As Object, Optional ByVal force As Boolean = False, Optional ByVal exclude As String)
    Dim sh As Worksheet
    Dim headerCell As Range
    Dim cursor As Range
    Dim loName As String
    Dim lo As ListObject
    Dim store As IHiddenNames

    If trads Is Nothing Then Exit Sub
    If TypeName(trads) <> "TranslationObject" Then Exit Sub

    'check language before moving to translation
    Set store = InternalStore()
    
    If (trads.Language = store.ValueAsString(WORKSHEETLANGUAGENAME)) And (Not force) Then
        Exit Sub
    End If

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    Set headerCell = sh.Cells(DROPDOWNSTARTROW, 4)
    Set cursor = sh.Cells(sh.Rows.Count, 4).End(xlUp)

    If cursor.Row <= headerCell.Row Then
        Exit Sub
    End If

    Do While cursor.Row > headerCell.Row
        loName = CStr(cursor.Value) 
        If LenB(loName) = 0 Then GoTo NextLo
        If LenB(exclude) > 0 Then
            If InStr(1, loName, exclude, vbTextCompare) > 0 Then GoTo NextLo
        End If

        On Error Resume Next
            Set lo = sh.ListObjects(loName)
        On Error GoTo 0 
        
        If Not (lo Is Nothing) Then trads.TranslateRange lo.DataBodyRange   
NextLo:    
        Set lo = Nothing
        Set cursor = cursor.Offset(-1)
    Loop

    'update language value
    store.EnsureName WORKSHEETLANGUAGENAME, trads.Language, HiddenNameTypeString
End Sub

'@label:operate
'@section Operate with the outside world
'===============================================================================

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
    dropdownName = "= " & CorrespondingWorkbookName(listName)

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


'@sub-title Add hyperlink on a range that targets or returns to the dropdown
'@param listName A String, the name of the dropdown targeted
'@param targetRng A Range. The Range where you want to return when you clik on the label.
Private Sub AddLink(ByVal listName As String, ByVal targetRng As Range, _
    Optional ByVal returnLink As Boolean = False)

    Dim loName As String
    Dim sh As Worksheet
    Dim anchorRng As Range
    Dim headerRng As Range
    Dim targetAddress As String
    Dim savedValue As String
    Dim savedInteriorColor As Long
    Dim savedFontColor As Long


    If Not Exists(listName) Then
        LogInfo "Can not add a returnlink on a non existing dropdown: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    If (targetRng Is Nothing) Then
        LogInfo "Can not add a returnlink to a non existing Range: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    loName = ListObjectName(listName)
    Set headerRng = Wksh.ListObjects(loName).Range.Cells(0, 1)

    If returnLink Then
        Set sh =  Wksh()
        Set anchorRng = headerRng
        targetAddress = "'" & targetRng.Parent.Name & "'" & "!" & targetRng.Address
    Else
        Set sh = targetRng.Parent
        Set anchorRng = targetRng
        targetAddress = "'" & headerRng.Parent.Name & "'" & "!" & headerRng.Address
    End If


    savedValue = anchorRng.Value
    savedInteriorColor = anchorRng.Interior.Color
    savedFontColor = anchorRng.Font.Color

    'Adding hyperlinks remove the writings 
    '(sub label and so on, so need to set it back)
    With sh
        .Hyperlinks.Add _
            Anchor:=anchorRng, _
            Address:="", _
            SubAddress:= targetAddress, _
            TextToDisplay:= ""
    End With

    'Adding links remove the formatting, set it back
    With anchorRng
        If savedValue <> vbNullString Then .Value = savedValue
        .Font.UnderLine = xlUnderlineStyleNone
        .Interior.Color = savedInteriorColor
        .Font.Color = savedFontColor
        .Font.Bold = True
        .Font.Italic = True
    End With
End Sub

'@label:AddReturnLink
'@sub-title Add hyperlink on a dropdown header that targets one cell
'@details
'The link is on the top of the header of one dropdown in the dropdownlist. The targeted adress is
'the cell that links to the current dropdown.
'@param listName A String, the name of dropdown to target
'@param returnRng A Range. Where to add the link
'@export
Private Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    AddLink listName:=listName, targetRng:=returnRng, returnLink:=True
End Sub


'@label:AddHyperLink
'@sub-title Add hyperlink on a cell header that targets one dropdown
'@details
'The link will target one dropdown in the dropdownlist. The targeted adress is
'the top of the dropdown, the label Range.
'@param listName A String, the name of dropdown to target
'@param linkRng A Range. Where to add the link
'@export
Private Sub AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
    AddLink listName:=listName, targetRng:=linkRng, returnLink:=False
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

Private Sub ShowDebug()
    Debug.Print "Check the informations of dropdownlist object " & Name()
End Sub

'Errors
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error : " & Chr(34) & errorMessage
End Sub

'@jump:Create
'@section Interface
'===============================================================================

'@jump:elements
'DropdownList Elements ---------------------------------------------------------

'@jump:Wksh
Private Property Get IDropdownLists_Wksh() As Worksheet
    Set IDropdownLists_Wksh = Wksh()
End Property

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

Private Property Get IDropdownLists_Values(ByVal listName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
    Set IDropdownLists_Values = Values(listName, includeHeaders)
End Property

'@jump:Length
Private Property Get IDropdownLists_Length() As Long
    IDropdownLists_Length = Length()
End Property

Private Property Get IDropdownLists_AllDropdowns() As BetterArray
    Set IDropdownLists_AllDropdowns = AllDropdowns
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
    Optional ByVal pasteAtBottom As Boolean = False, _
    Optional ByVal removeDuplicates As Boolean = True)
    Update UpdateData:=UpdateData, _
        listName:=listName, _
        pasteAtBottom:=pasteAtBottom, _
        removeDuplicates:=removeDuplicates
End Sub

'@jump:Sort
Private Sub IDropdownLists_Sort(ByVal listName As String, _
    Optional ByVal sortOrder As Long = xlAscending)
    Sort listName, sortOrder
End Sub

'@jump:Translate
Private Sub IDropdownLists_Translate(ByVal trads As Object, Optional ByVal force As Boolean = False, Optional ByVal exclude As String)
    Translate trads, force
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

'@jump:HasCheckings
'Test if the object has checkings
Private Property Get IDropdownLists_HasCheckings() As Boolean
    IDropdownLists_HasCheckings = HasCheckings()
End Property

'@jump:CheckingValues
'Get the checking values
Private Property Get IDropdownLists_CheckingValues() As Object
    Set IDropdownLists_CheckingValues = CheckingValues()
End Property
