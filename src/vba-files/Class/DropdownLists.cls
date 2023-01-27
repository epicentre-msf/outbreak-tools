VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropdownLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Add dropdowns to the linelist")
'@IgnoreModule

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
        colIndex = IIf(colIndex = 1 And .Cells(1, 1).Value <> "", colIndex, colIndex + 2)
    End With
    LastColumn = colIndex
End Property

Private Property Get CurrentListName(ByVal ListName As String) As String
    CurrentListName = Replace(Application.WorksheetFunction.Trim(ListName), " ", "_")
End Property

Private Property Get CurrentHeaderName(ByVal ListName As String) As String
    CurrentHeaderName = "dropdown_" & CurrentListName(ListName)
End Property

Private Property Get CurrentLoName(ByVal ListName As String) As String
    CurrentLoName = "list_" & CurrentListName(ListName)
End Property

'Test if a listobject exists in a sheet
Public Function Exists(ByVal ListName As String) As Boolean
    Dim Lo As ListObject
    Dim loName As String
    Dim sh As Worksheet
    Set sh = Wksh()
    loName = CurrentLoName(ListName)
    On Error Resume Next
    Set Lo = sh.ListObjects(loName)
    On Error GoTo 0
    Exists = (Not Lo Is Nothing)
End Function

'Add New list to the dropdown sheet
Private Sub Add(ByVal listValues As BetterArray, ByVal ListName As String)
  
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim headerName As String                     'header Name
    Dim loName As String                         'List object name
    Dim rng As Range
    Dim wkb As Workbook


    If (Not Exists(ListName)) Then
        loName = CurrentLoName(ListName)
        headerName = CurrentHeaderName(ListName)
        colIndex = LastColumn()
        Set sh = Wksh()

        With sh
            listValues.ToExcelRange .Cells(2, colIndex)
            .Cells(1, colIndex).Value = headerName
            Set rng = .Range(.Cells(1, colIndex), .Cells(listValues.Length + 1, colIndex))

            'Add A list object table
            .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loName
            Set wkb = .Parent
        End With

        wkb.NAMES.Add Name:=headerName, RefersToR1C1:="=" & loName & "[" & headerName & "]"
    End If
End Sub

'Remove the validation
Private Sub Remove(ByVal ListName As String)
    Dim sh As Worksheet
    Dim wkb As Workbook
    Dim rng As Range
    Dim loName As String
    Dim headerName As String

    If Exists(ListName) Then
    
        loName = CurrentLoName(ListName)
        headerName = CurrentHeaderName(ListName)

        Set sh = Wksh()
        Set rng = sh.ListObjects(loName).Range
        Set wkb = sh.Parent
    
        wkb.NAMES(headerName).Delete
    
        If rng.Column = 1 Then
    
            rng.EntireColumn.Delete
    
        Else
    
            rng.EntireColumn.Delete
            rng.Offset(, 1).EntireColumn.Delete
    
        End If

    End If
End Sub

'Update a list with new values
Private Sub Update(ByVal UpdateData As BetterArray, ByVal ListName As String)
    
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim rng As Range
    Dim loName As String
    Dim Lo As ListObject

    If Not Exists(ListName) Then Exit Sub
    Set sh = Wksh()
    loName = CurrentLoName(ListName)
    Set Lo = sh.ListObjects(loName)
    Set cellRng = Lo.Range.Cells(2, 1)
    
    'put the new data in the range
    UpdateData.ToExcelRange cellRng
    
    'Resize the listObject to the new Range
    Set rng = sh.Range(cellRng.Cells(0, 1), cellRng.Cells(UpdateData.Length, 1))
    Lo.Resize rng
End Sub

'Set a validation on a range
Private Sub SetValidation(ByVal cellRng As Range, ByVal ListName As String, _
                          Optional ByVal alertType As String = "info", _
                          Optional ByVal message As String = vbNullString)

    Dim headerName As String

    If Not Exists(ListName) Then ThrowError "There is no available dropdown for " & ListName

    headerName = "= " & CurrentHeaderName(ListName)
    
    With cellRng.validation
        .Delete
        Select Case alertType
        Case "error"                             '"error"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=headerName
        Case "warning"                           '"warning"
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=headerName
        Case Else                                'for all the others, add an information alert
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=headerName
        End Select

        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .errorTitle = ""
        .InputMessage = ""
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
Public Property Get IDropdownLists_Exists(ByVal ListName As String) As Boolean
    IDropdownLists_Exists = Exists(ListName)
End Property

Public Sub IDropdownLists_Add(ByVal listValues As BetterArray, ByVal ListName As String)
    Add listValues:=listValues, ListName:=ListName
End Sub

Public Sub IDropdownLists_Remove(ByVal ListName As String)
    Remove ListName:=ListName
End Sub

Public Sub IDropdownLists_SetValidation(ByVal cellRng As Range, ByVal ListName As String, _
                                        Optional ByVal alertType As String = "info", _
                                        Optional ByVal message As String = vbNullString)

    SetValidation cellRng:=cellRng, ListName:=ListName, alertType:=alertType, message:=message
End Sub


Public Sub IDropdownLists_Update(ByVal UpdateData As BetterArray, ByVal ListName As String)
    Update UpdateData:=UpdateData, ListName:=ListName
End Sub
