VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropdownLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Add, remove and twist dropdowns"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Add, remove and twist dropdowns")

Option Explicit
Implements IDropdownLists

Private Type TDropdownLists
    sh As Worksheet
End Type

Private this As TDropdownLists
Private Const CLASSNAME As String = "DropdownLists"

Public Function Create(ByVal outsh As Worksheet) As IDropdownLists
    'Errors
    If outsh Is Nothing Then ThrowError "Dropdown worksheet not initialised"

    With New DropdownLists
        Set .Wksh = outsh
        Set Create = .Self
    End With

End Function

Public Property Get Self() As IDropdownLists
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal outsh As Worksheet)
    Set this.sh = outsh
End Property

Private Property Get LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()
    With sh
        colIndex = .Cells(1, .Columns.Count).End(xlToLeft).Column
        colIndex = IIf(colIndex = 1 And IsEmpty(.Cells(1, 1)), colIndex, colIndex + 2)
    End With
    LastColumn = colIndex
End Property

Private Property Get CurrentListName(ByVal listName As String) As String
    CurrentListName = Replace(Application.WorksheetFunction.Trim(listName), " ", "_")
End Property

Private Property Get CurrentHeaderName(ByVal listName As String) As String
    CurrentHeaderName = "dropdown_" & CurrentListName(listName)
End Property

Private Property Get CurrentLoName(ByVal listName As String) As String
    CurrentLoName = "list_" & CurrentListName(listName)
End Property

'Test if a listobject exists in a sheet
Public Function Exists(ByVal listName As String) As Boolean
    Dim Lo As ListObject
    Dim loname As String
    Dim sh As Worksheet
    Set sh = Wksh()
    loname = CurrentLoName(listName)
    On Error Resume Next
    Set Lo = sh.ListObjects(loname)
    On Error GoTo 0
    Exists = (Not Lo Is Nothing)
End Function

'Add New list to the dropdown sheet
Private Sub Add(ByVal listValues As BetterArray, ByVal listName As String)

    Dim sh As Worksheet
    Dim colIndex As Long
    Dim headerName As String                     'header Name
    Dim loname As String                         'List object name
    Dim rng As Range
    Dim Wkb As Workbook


    If (Not Exists(listName)) Then
        loname = CurrentLoName(listName)
        headerName = CurrentHeaderName(listName)
        colIndex = LastColumn()
        Set sh = Wksh()

        With sh
            listValues.ToExcelRange .Cells(2, colIndex)
            .Cells(1, colIndex).Value = headerName
            Set rng = .Range(.Cells(1, colIndex), .Cells(listValues.Length + 1, colIndex))

            'Add A list object table
            .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loname
            Set Wkb = .Parent
        End With

        Wkb.Names.Add Name:=headerName, RefersToR1C1:="=" & loname & "[" & headerName & "]"
    End If
End Sub

'Remove the validation
Private Sub Remove(ByVal listName As String)
    Dim sh As Worksheet
    Dim Wkb As Workbook
    Dim rng As Range
    Dim loname As String
    Dim headerName As String

    If Exists(listName) Then

        loname = CurrentLoName(listName)
        headerName = CurrentHeaderName(listName)

        Set sh = Wksh()
        Set rng = sh.ListObjects(loname).Range
        Set Wkb = sh.Parent

        Wkb.Names(headerName).Delete

        If rng.Column = 1 Then

            rng.EntireColumn.Delete

        Else

            rng.EntireColumn.Delete
            rng.Offset(, 1).EntireColumn.Delete

        End If

    End If
End Sub

'Update a list with new values
Private Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String)

    Dim sh As Worksheet
    Dim cellRng As Range
    Dim rng As Range
    Dim loname As String
    Dim Lo As ListObject

    If Not Exists(listName) Then Exit Sub
    Set sh = Wksh()
    loname = CurrentLoName(listName)
    Set Lo = sh.ListObjects(loname)
    Set cellRng = Lo.Range.Cells(2, 1)

    'put the new data in the range
    UpdateData.ToExcelRange cellRng

    'Resize the listObject to the new Range
    Set rng = sh.Range(cellRng.Cells(0, 1), cellRng.Cells(UpdateData.Length, 1))
    Lo.Resize rng
End Sub

'Set a validation on a range
Private Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                          Optional ByVal alertType As String = "info", _
                          Optional ByVal message As String = vbNullString, _
                          Optional ByVal ignoreBlank As Boolean = True)

    Dim headerName As String

    If Not Exists(listName) Then ThrowError "There is no available dropdown for " & listName

    headerName = "= " & CurrentHeaderName(listName)

    With cellRng.Validation
        .Delete
        Select Case alertType
        Case "error"                             '"error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=headerName
        Case "warning"                           '"warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=headerName
        Case Else                                'for all the others, add an information alert
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=headerName
        End Select

        .ignoreBlank = ignoreBlank
        .InCellDropdown = True
        .InputTitle = vbNullString
        .errorTitle = vbNullString
        .InputMessage = vbNullString
        .errorMessage = message
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Private Property Get IDropdownLists_Exists(ByVal listName As String) As Boolean
    IDropdownLists_Exists = Exists(listName)
End Property

Private Sub IDropdownLists_Add(ByVal listValues As BetterArray, ByVal listName As String)
    Add listValues:=listValues, listName:=listName
End Sub

Private Sub IDropdownLists_Remove(ByVal listName As String)
    Remove listName:=listName
End Sub

Private Sub IDropdownLists_SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                                        Optional ByVal alertType As String = "info", _
                                        Optional ByVal message As String = vbNullString, _
                                        Optional ByVal ignoreBlank As Boolean = True)

    SetValidation cellRng:=cellRng, listName:=listName, alertType:=alertType, message:=message, ignoreBlank:=ignoreBlank
End Sub


Private Sub IDropdownLists_Update(ByVal UpdateData As BetterArray, ByVal listName As String)
    Update UpdateData:=UpdateData, listName:=listName
End Sub
