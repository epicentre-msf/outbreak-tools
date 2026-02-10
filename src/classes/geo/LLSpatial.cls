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

'@class LLSpatial
'@description
'Manages the spatial analysis tables that aggregate filtered linelist data
'by geographic admin level and health facility. The spatial sheet contains
'a "listofgeovars" ListObject that lists which variables have spatial
'tables, and one ListObject per admin level per variable (named
'"spatial_adm1_varname" etc.) holding aggregated values, formula columns,
'and attack rate columns. This class populates those tables from filtered
'companion sheets, sorts them by formula or attack rate, and retrieves
'ranked values for use in analyses and graphs.
'@depends BetterArray


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

'@label:create
'@sub-title Create a new LLSpatial instance from the spatial tables worksheet
'@details
'Factory method that validates the supplied worksheet meets requirements
'(must be named "spatial_tables__" and contain the "listofgeovars"
'ListObject), then returns a new instance through the ILLSpatial
'interface. Called on the predeclared instance.
'@param sh Worksheet. The spatial tables worksheet.
'@return ILLSpatial. A fully initialised instance ready for use.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.InvalidArgument When the worksheet name is not "spatial_tables__".
'@export
Public Function Create(ByVal sh As Worksheet) As ILLSpatial
    CheckRequirements sh

    With New LLSpatial
        Set .Wksh = sh
        Set Create = .Self
    End With
End Function

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor so consuming code can fluently retrieve the
'interface reference from the predeclared Create method.
'@return ILLSpatial. The current instance cast to the interface.
Public Property Get Self() As ILLSpatial
    Set Self = Me
End Property

'@label:wksh
'@prop-title Spatial tables worksheet
'@return Worksheet. The backing spatial tables worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the spatial tables worksheet
'@param sh Worksheet. The worksheet to store.
Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

'@section Naming Conventions
'===============================================================================

'@label:current-header-name
'@prop-title Named range name for a spatial table header
'@details
'Builds the conventional named range name used to locate the header row
'of a spatial table. The pattern is "tabl_{adminName}_{varName}".
'@param varName String. The variable name component.
'@param adminName String. The admin level component (e.g. "adm1", "hf").
'@return String. The constructed named range name.
Private Property Get CurrentHeaderName(ByVal varName As String, _
                                       ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

'@label:current-lo-name
'@prop-title ListObject name for a spatial table
'@details
'Builds the conventional ListObject name for a spatial table. The pattern
'is "spatial_{adminName}_{varName}".
'@param varName String. The variable name component.
'@param adminName String. The admin level component (e.g. "adm1", "hf").
'@return String. The constructed ListObject name.
Private Property Get CurrentLoName(ByVal varName As String, _
                                   ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'@section Exists
'===============================================================================

'@label:exists
'@sub-title Check if a variable name exists in the spatial variables list
'@details
'Searches the "listofgeovars" ListObject range for a partial match of the
'supplied variable name using xlPart matching. Returns True when at least
'one cell contains the variable name as a substring, indicating that
'spatial analysis tables exist for this variable.
'@param varName String. Variable name or partial name to look up.
'@return Boolean. True when the variable is found in the list.
'@export
Private Function Exists(ByVal varName As String) As Boolean
    Dim listVarsRng As Range
    Set listVarsRng = Wksh().ListObjects("listofgeovars").Range
    Exists = Not (listVarsRng.Find(What:=varName, lookAt:=xlPart, _
                                   MatchCase:=True) Is Nothing)
End Function

'@label:lo-exists
'@sub-title Check if a ListObject exists on the spatial sheet
'@details
'Attempts to retrieve the named ListObject from the spatial worksheet.
'Returns True when the object is found, False otherwise. Uses error
'trapping to handle the case where the name does not exist.
'@param loName String. The ListObject name to check.
'@return Boolean. True when the ListObject exists.
Private Function LoExists(ByVal loName As String) As Boolean
    Dim Lo As ListObject

    On Error Resume Next
    Set Lo = Wksh().ListObjects(loName)
    On Error GoTo 0

    LoExists = Not (Lo Is Nothing)
End Function

'@section Update
'===============================================================================

'@label:update
'@sub-title Update all spatial tables from HList filtered data
'@details
'Iterates all worksheets in the workbook looking for HList sheets. For
'each HList sheet, locates its filtered companion sheet (from cell(1,5)),
'then walks the variable columns looking for geo1 and hf control values.
'When a geo1 column is found, extracts the concat column data from the
'filtered companion and calls UpdateSpatialValue to populate the spatial
'tables for all four admin levels. When an hf column is found, does the
'same for the HF spatial table. After processing all sheets, recalculates
'the spatial sheet so formula columns update.
'@export
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

'@label:list-of-existing-var-names
'@sub-title Build a list of all spatial variable names matching a base name
'@details
'Searches the "listofgeovars" ListObject for all cells that partially
'match the supplied variable name. Returns a BetterArray of the full
'variable names found. The search progresses through the range by
'narrowing it after each match to avoid finding the same cell again.
'@param varName String. Base variable name to search for.
'@return BetterArray. All matching spatial variable names.
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

'@label:update-spatial-value
'@sub-title Populate spatial ListObjects from a filtered input range
'@details
'For each matching variable name returned by ListOfExistingVarNames,
'iterates admin levels 1 through 4 (or just HF when isHF is True). For
'each admin level, pastes the input range values to a scratch column
'(RNG_PastingCol), removes duplicates, clears the existing spatial table
'data, pastes the deduplicated values into the ListObject first column,
'resizes the table to fit, and sorts by the formula column in descending
'order. The input range is shifted one column right for each admin level
'to cascade through the concat columns.
'@param inputRng Range. The filtered data range to extract values from.
'@param inputVarName String. Base variable name to look up in the spatial list.
'@param isHF Optional Boolean. True to update HF tables only, False for admin1-4. Defaults to False.
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

'@label:sort
'@sub-title Sort spatial tables by formula value or attack rate for a given table id
'@details
'Locates the spatial variable whose name partially matches tabId in the
'"listofgeovars" list, then iterates admin1 through admin4 and sorts each
'corresponding ListObject by either the formula column or the attack_rate
'column in descending order. Silently skips tables that do not exist, have
'no data body, or whose first column is empty.
'@param tabId String. Table identifier used to locate the variable name via partial match.
'@param onAR Optional Boolean. True to sort by attack_rate column, False for formula column. Defaults to False.
'@export
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

'@label:top-geo-value
'@prop-title Value at a given rank from a sorted spatial geo table
'@details
'Constructs the ListObject name from the admin name and variable name
'(with tabId suffix), then retrieves the value at the requested rank
'position. The concatValue parameter selects the column: 1 returns
'the name column (first column), any other value returns the concat
'column (third column). Returns an empty string when the table does
'not exist, has no data, or the rank exceeds available rows.
'@param adminName String. Admin level name ("adm1", "adm2", "adm3", or "adm4").
'@param orderValue Long. Rank position in the sorted table (1-based).
'@param inputVarName String. Base variable name.
'@param tabId Optional String. Table identifier suffix appended to the variable name. Defaults to vbNullString.
'@param concatValue Optional Byte. 1 for the name column, other values for the concat column. Defaults to 1.
'@return String. The value at the requested rank, or empty string if not found.
'@export
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

'@label:top-hf-value
'@prop-title Value at a given rank from a sorted spatial HF table
'@details
'Constructs the HF ListObject name from the variable name (with tabId
'suffix), then retrieves the value at the requested rank position from
'the first column. Returns an empty string when the table does not exist,
'has no data, or the rank exceeds available rows.
'@param orderValue Long. Rank position in the sorted table (1-based).
'@param inputVarName String. Base variable name.
'@param tabId Optional String. Table identifier suffix appended to the variable name. Defaults to vbNullString.
'@return String. The value at the requested rank, or empty string if not found.
'@export
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

'@label:check-requirements
'@sub-title Validate that the worksheet is the spatial tables sheet
'@details
'Guards the factory Create method by ensuring the worksheet is not
'Nothing and its name matches the expected "spatial_tables__" constant.
'Raises an error when either condition fails.
'@param sh Worksheet. The worksheet to validate.
'@throws ProjectError.ObjectNotInitialized When the worksheet is Nothing.
'@throws ProjectError.InvalidArgument When the worksheet name does not match "spatial_tables__".
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

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME,
'providing a consistent stack trace across all methods in this class.
'@param errNumber Long. The ProjectError error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ILLSpatial contract.
'See the corresponding members above for full documentation.

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
