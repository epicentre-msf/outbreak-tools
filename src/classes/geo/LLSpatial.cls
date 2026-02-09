VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Spatial analysis table management"

'@PredeclaredId
'@Folder("Geo")
'@ModuleDescription("Spatial analysis table management")
'@IgnoreModule UnrecognizedAnnotation, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLSpatial

'@section State
'===============================================================================

Private Type TLLSpatial
    sh As Worksheet
End Type

Private this As TLLSpatial
Private Const CLASS_NAME As String = "LLSpatial"
Private Const SPATIAL_SHEET As String = "spatial_tables__"
Private Const ADMIN As String = "adm"
Private Const CONCAT As String = "concat_"
Private Const HF As String = "hf"

'@section Factory
'===============================================================================

' @description Create a new LLSpatial instance from the spatial tables worksheet.
' @param sh The spatial tables worksheet (must be named "spatial_tables__")
' @return ILLSpatial interface pointer
Public Function Create(ByVal sh As Worksheet) As ILLSpatial
    CheckRequirements sh

    With New LLSpatial
        Set .Wksh = sh
        Set Create = .Self
    End With
End Function

Public Property Get Self() As ILLSpatial
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

'@section Naming Conventions
'===============================================================================

' @description Build the header named range name for a spatial table.
Private Property Get CurrentHeaderName(ByVal varName As String, _
                                       ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

' @description Build the ListObject name for a spatial table.
Private Property Get CurrentLoName(ByVal varName As String, _
                                   ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'@section Exists
'===============================================================================

' @description Check if a variable name exists in the spatial variables list.
Private Function Exists(ByVal varName As String) As Boolean
    Dim listVarsRng As Range
    Set listVarsRng = Wksh().ListObjects("listofgeovars").Range
    Exists = Not (listVarsRng.Find(What:=varName, lookAt:=xlPart, _
                                   MatchCase:=True) Is Nothing)
End Function

' @description Check if a ListObject exists on the spatial sheet.
Private Function LoExists(ByVal loName As String) As Boolean
    Dim Lo As ListObject

    On Error Resume Next
    Set Lo = Wksh().ListObjects(loName)
    On Error GoTo 0

    LoExists = Not (Lo Is Nothing)
End Function

'@section Update
'===============================================================================

' @description Update all spatial tables from HList filtered data.
'              Iterates all HList worksheets, finds geo1/hf control columns,
'              and populates the corresponding spatial ListObjects.
Private Sub Update()
    Dim actsh As Worksheet
    Dim sh As Worksheet
    Dim filtsh As Worksheet
    Dim wb As Workbook
    Dim rng As Range
    Dim tableName As String
    Dim varName As String
    Dim cellRng As Range
    Dim controlValue As String
    Dim Lo As ListObject

    Set actsh = Wksh()
    Set wb = actsh.Parent

    For Each sh In wb.Worksheets

        If sh.Cells(1, 3).Value <> "HList" Then GoTo ContinueSheet

        On Error Resume Next
        Set filtsh = wb.Worksheets(sh.Cells(1, 5).Value)
        On Error GoTo 0

        If filtsh Is Nothing Then GoTo ContinueSheet

        Set Lo = Nothing
        On Error Resume Next
        Set Lo = filtsh.ListObjects(1)
        On Error GoTo 0

        If Lo Is Nothing Then GoTo ContinueSheet
        If Lo.DataBodyRange Is Nothing Then GoTo ContinueSheet

        tableName = sh.Cells(1, 4).Value
        Set cellRng = sh.Range(tableName & "_START")

        Do While Not IsEmpty(cellRng)
            controlValue = cellRng.Offset(-4).Value

            If controlValue = "geo1" Then
                varName = Replace(cellRng.Value, ADMIN & "1_", vbNullString)

                If Exists(varName) Then
                    On Error Resume Next
                    Set rng = Lo.ListColumns(CONCAT & ADMIN & "1_" & varName).DataBodyRange()
                    On Error GoTo 0

                    If Not rng Is Nothing Then
                        UpdateSpatialValue rng, varName
                    End If
                End If

                'Skip geo2-4, pcodes, and concat columns (4+4+3 = 11)
                Set cellRng = cellRng.Offset(, 11)

            ElseIf controlValue = "hf" Then
                varName = Replace(cellRng.Value, HF & "_", vbNullString)

                If Exists(varName) Then
                    On Error Resume Next
                    Set rng = Lo.ListColumns(HF & "_" & varName).DataBodyRange()
                    On Error GoTo 0

                    If Not rng Is Nothing Then
                        UpdateSpatialValue rng, varName, isHF:=True
                    End If
                End If

                'Skip HF pcode column
                Set cellRng = cellRng.Offset(, 1)
            End If

            Set cellRng = cellRng.Offset(, 1)
        Loop

ContinueSheet:
        Set filtsh = Nothing
    Next

    actsh.UsedRange.Calculate
End Sub

'@section List of Variable Names
'===============================================================================

' @description Build a list of all spatial variable names matching a base name.
'              Searches the "listofgeovars" ListObject for partial matches.
Private Function ListOfExistingVarNames(ByVal varName As String) As BetterArray
    Dim listRng As Range
    Dim cellRng As Range
    Dim sh As Worksheet
    Dim listTab As BetterArray

    Set sh = Wksh()
    Set listTab = New BetterArray
    Set listRng = sh.ListObjects("listofgeovars").Range

    Do While Not (listRng.Find(What:=varName, lookAt:=xlPart, _
                               MatchCase:=True) Is Nothing)

        Set cellRng = listRng.Find(What:=varName, lookAt:=xlPart, _
                                   MatchCase:=True, _
                                   After:=listRng.Cells(listRng.Rows.Count, 1))
        listTab.Push cellRng.Value

        If listRng.Rows.Count = 1 Then Exit Do
        If cellRng.Cells(2, 1).Row > listRng.Cells(listRng.Rows.Count, 1).Row Then Exit Do

        Set listRng = sh.Range(cellRng.Cells(2, 1), _
                               listRng.Cells(listRng.Rows.Count, 1))
    Loop

    Set ListOfExistingVarNames = listTab.Clone()
End Function

'@section Update Spatial Values
'===============================================================================

' @description Populate spatial ListObjects from a filtered input range.
'              Extracts unique values, pastes into the spatial table, removes
'              duplicates, resizes, and sorts by formula column.
Private Sub UpdateSpatialValue(ByVal inputRng As Range, _
                               ByVal inputVarName As String, _
                               Optional ByVal isHF As Boolean = False)

    Dim pasteRng As Range
    Dim counter As Long
    Dim adminCounter As Long
    Dim sh As Worksheet
    Dim cellRng As Range
    Dim rng As Range
    Dim Lo As ListObject
    Dim loName As String
    Dim adminName As String
    Dim sortRng As Range
    Dim nbRows As Long
    Dim LoRng As Range
    Dim listVars As BetterArray
    Dim varName As String
    Dim endCounter As Long

    Set sh = Wksh()
    Set listVars = ListOfExistingVarNames(inputVarName)

    endCounter = IIf(isHF, 1, 4)

    For counter = listVars.LowerBound To listVars.UpperBound
        varName = listVars.Item(counter)
        Set rng = inputRng

        For adminCounter = 1 To endCounter
            nbRows = rng.Cells.Count

            If isHF Then
                loName = CurrentLoName(varName, HF)
                adminName = HF
            Else
                loName = CurrentLoName(varName, ADMIN & adminCounter)
                adminName = ADMIN & adminCounter
            End If

            If Not LoExists(loName) Then GoTo ContinueAdmin

            Set Lo = sh.ListObjects(loName)

            'Paste values to the scratch column, remove duplicates
            Set cellRng = sh.Range("RNG_PastingCol")
            Set pasteRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nbRows, 1))
            pasteRng.Value = rng.Value

            On Error Resume Next
            pasteRng.RemoveDuplicates Columns:=1, Header:=xlNo
            On Error GoTo 0

            'Resize after dedup
            nbRows = sh.Cells(sh.Rows.Count, pasteRng.Column).End(xlUp).Row
            Set pasteRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nbRows, 1))

            'Clear existing data and paste into the ListObject
            Set cellRng = Lo.Range.Cells(2, 1)

            On Error Resume Next
            Lo.ListColumns(1).DataBodyRange.Cells.Clear
            On Error GoTo 0

            Set LoRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nbRows, 1))
            LoRng.Value = pasteRng.Value
            pasteRng.Cells.Clear

            'Resize the ListObject (including Attack Rate columns)
            Set LoRng = sh.Range(cellRng.Cells(0, 1), cellRng.Cells(nbRows, 4))
            Lo.Resize LoRng

            'Sort by formula column
            On Error Resume Next
            Set sortRng = Lo.ListColumns("formula_" & adminName).DataBodyRange
            sortRng.Calculate
            Lo.ListColumns(1).DataBodyRange.Sort key1:=sortRng, Order1:=xlDescending
            On Error GoTo 0

ContinueAdmin:
            Set rng = rng.Offset(, 1)
        Next
    Next
End Sub

'@section Sort
'===============================================================================

' @description Sort spatial tables by formula value or attack rate for a given table id.
Private Sub Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False)
    Dim sortRng As Range
    Dim Lo As ListObject
    Dim sh As Worksheet
    Dim adminCounter As Long
    Dim sortRngName As String
    Dim varName As String
    Dim listRng As Range
    Dim cellRng As Range

    Set sh = Wksh()
    Set listRng = sh.ListObjects("listofgeovars").Range

    Set cellRng = listRng.Find(What:=tabId, lookAt:=xlPart, MatchCase:=True)
    If cellRng Is Nothing Then Exit Sub

    varName = cellRng.Value

    For adminCounter = 1 To 4
        Set Lo = Nothing
        On Error Resume Next
        Set Lo = sh.ListObjects(CurrentLoName(varName, ADMIN & adminCounter))
        On Error GoTo 0

        If Lo Is Nothing Then GoTo ContinueSort
        If Lo.DataBodyRange Is Nothing Then GoTo ContinueSort
        If IsEmpty(Lo.ListColumns(1).DataBodyRange) Then GoTo ContinueSort

        sortRngName = IIf(onAR, "attack_rate_", "formula_") & ADMIN & adminCounter

        On Error Resume Next
        Set sortRng = Lo.ListColumns(sortRngName).Range
        Lo.ListColumns(1).Range.Sort key1:=sortRng, Order1:=xlDescending, Header:=xlYes
        On Error GoTo 0

ContinueSort:
    Next
End Sub

'@section TopGeoValue / TopHFValue
'===============================================================================

' @description Get the value at a given rank from a sorted spatial geo table.
Private Property Get TopGeoValue(ByVal adminName As String, _
                                 ByVal orderValue As Long, _
                                 ByVal inputVarName As String, _
                                 Optional ByVal tabId As String = vbNullString, _
                                 Optional ByVal concatValue As Byte = 1) As String
    Dim Lo As ListObject
    Dim loName As String
    Dim varName As String
    Dim rng As Range

    varName = inputVarName & "_" & tabId
    loName = CurrentLoName(varName, adminName)

    If Not LoExists(loName) Then Exit Property

    Set Lo = Wksh().ListObjects(loName)
    If Lo.DataBodyRange Is Nothing Then Exit Property

    If concatValue = 1 Then
        Set rng = Lo.ListColumns(1).DataBodyRange()
    Else
        Set rng = Lo.ListColumns(3).DataBodyRange()
    End If

    If orderValue > rng.Rows.Count Then Exit Property

    TopGeoValue = CStr(rng.Cells(orderValue, 1).Value)
End Property

' @description Get the value at a given rank from a sorted spatial HF table.
Private Property Get TopHFValue(ByVal orderValue As Long, _
                                ByVal inputVarName As String, _
                                Optional ByVal tabId As String = vbNullString) As String
    Dim Lo As ListObject
    Dim loName As String
    Dim varName As String
    Dim rng As Range

    varName = inputVarName & "_" & tabId
    loName = CurrentLoName(varName, HF)

    If Not LoExists(loName) Then Exit Property

    Set Lo = Wksh().ListObjects(loName)
    If Lo.DataBodyRange Is Nothing Then Exit Property

    Set rng = Lo.ListColumns(1).DataBodyRange()

    If orderValue > rng.Rows.Count Then Exit Property

    TopHFValue = CStr(rng.Cells(orderValue, 1).Value)
End Property

'@section Validation
'===============================================================================

' @description Validate that the worksheet is the spatial tables sheet.
Private Sub CheckRequirements(ByVal sh As Worksheet)
    If sh Is Nothing Then
        ThrowError ObjectNotInitialized, "Worksheet not provided"
    End If

    If sh.Name <> SPATIAL_SHEET Then
        ThrowError InvalidArgument, _
            "Expected worksheet named " & SPATIAL_SHEET & " but got " & sh.Name
    End If
End Sub

'@section Error Handling
'===============================================================================

' @description Raise a ProjectError-based exception.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================

Private Property Get ILLSpatial_Exists(ByVal varName As String) As Boolean
    ILLSpatial_Exists = Exists(varName)
End Property

Private Sub ILLSpatial_Update()
    Update
End Sub

Private Property Get ILLSpatial_TopGeoValue(ByVal adminName As String, _
                                            ByVal orderValue As Long, _
                                            ByVal varName As String, _
                                            Optional ByVal tabId As String = vbNullString, _
                                            Optional ByVal concatValue As Byte = 1) As String
    ILLSpatial_TopGeoValue = TopGeoValue(adminName, orderValue, varName, tabId, concatValue)
End Property

Private Property Get ILLSpatial_TopHFValue(ByVal orderValue As Long, _
                                           ByVal inputVarName As String, _
                                           Optional ByVal tabId As String = vbNullString) As String
    ILLSpatial_TopHFValue = TopHFValue(orderValue, inputVarName, tabId)
End Property

Private Sub ILLSpatial_Sort(ByVal tabId As String, Optional ByVal onAR As Boolean = False)
    Sort tabId, onAR
End Sub
