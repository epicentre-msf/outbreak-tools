VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventLinelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Workbook-level linelist event coordinator"
Option Explicit


'@PredeclaredId
'@Folder("Linelist Events")
'@ModuleDescription("Encapsulates workbook-level linelist events and lazily-cached domain managers")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Implements IEventLinelist

'@class EventLinelist
'@description
'Encapsulates workbook-level linelist events and lazily-cached domain managers.
'This class is the single entry point for all event handling in the output
'linelist workbook. It reacts to sheet change (HList geo cascades, editable
'labels, list auto, go-to-section, multiple choices; VList editable labels and
'multiple choices; analysis go-to-section and spatial dispatch), sheet deactivate
'(list auto update), selection change (geo dropdown refresh),
'double-click (spatio-temporal geo form), filter table updates, list-auto
'dropdown maintenance, and geo/spatial UDF delegation.
'@depends BetterArray, HiddenNames, DropdownLists, LLGeo, LLSpatial, LLdictionary, LLVariables, LLTranslation, TranslationObject, Passwords, ProjectError


'@section State
'===============================================================================
Private Type TEventLinelist
    hostBook        As Workbook           'Linelist workbook
    tradmess        As ITranslationObject 'Cached message translations
    lltrads         As ILLTranslation     'Cached translation helper
    wkbNames        As IHiddenNames       'Cached workbook-level hidden names
    drop            As IDropdownLists     'Cached dropdown manager
    geo             As ILLGeo             'Cached geo helper
    sp              As ILLSpatial         'Cached spatial helper
    dict            As ILLdictionary      'Cached dictionary helper
    vars            As ILLVariables       'Cached variable lookup helper
    pass            As IPasswords         'Cached password manager
End Type

Private this As TEventLinelist


'@section Constants
'===============================================================================
Private Const CLASSNAME         As String = "EventLinelist"
Private Const GOTOSECCODE       As String = "go_to_section"
Private Const DROPDOWNSHEET     As String = "dropdown_lists__"
Private Const GEOSHEET          As String = "Geo"
Private Const SPATIALSHEET      As String = "spatial_tables__"
Private Const DICTSHEET         As String = "Dictionary"
Private Const LLSHEET           As String = "LinelistTranslation"
Private Const PASSSHEET         As String = "__pass"
Private Const PRINTPREFIX       As String = "print_"
Private Const CRFPREFIX         As String = "crf_"


'@section Factory
'===============================================================================

'@sub-title Build a configured EventLinelist instance bound to a host workbook.
'@param hostBook Workbook. The linelist workbook.
'@return IEventLinelist. Configured service instance.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@export
Public Function Create(ByVal hostBook As Workbook) As IEventLinelist
    Dim instance As EventLinelist

    Set instance = New EventLinelist
    instance.Configure hostBook
    Set Create = instance.Self
End Function

'@sub-title Current object instance cast to the interface.
'@return IEventLinelist. The current instance.
Public Property Get Self() As IEventLinelist
    Set Self = Me
End Property


'@section IEventLinelist implementation
'===============================================================================

'@sub-title Configure the service with its host workbook context.
'@param hostBook Workbook. The linelist workbook.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
    If (hostBook Is Nothing) Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook is required to configure EventLinelist"
    End If

    Set this.hostBook = hostBook
    ResetInternalCaches
End Sub

Private Sub IEventLinelist_Configure(ByVal hostBook As Workbook)
    Configure hostBook
End Sub


'@sub-title React to workbook open notifications.
'@export
Public Sub OnWorkbookOpen()
    On Error Resume Next
    Application.OnKey "^+g", "ClickGeoApp"
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = True
    Application.FormatStaleValues = False
    On Error GoTo 0
End Sub

Private Sub IEventLinelist_OnWorkbookOpen()
    OnWorkbookOpen
End Sub


'@sub-title Handle worksheet change events.
'@param sh Worksheet. The worksheet raising the change event.
'@param target Range. The edited cells.
'@export
Public Sub OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    Dim shHn As IHiddenNames
    Set shHn = HiddenNames.Create(sh)

    Dim shType As String
    shType = shHn.ValueAsString("sheet_type")

    Select Case shType
        Case "HList"
            HandleHListChange sh, target, shHn
        Case "VList"
            HandleVListChange sh, target, shHn
        Case "Uni-Bi-Analysis", "TS-Analysis", "SPT-Analysis", "SP-Analysis"
            HandleAnalysisChange sh, target, shType
    End Select
End Sub

Private Sub IEventLinelist_OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    OnSheetChange sh, target
End Sub


'@sub-title Handle worksheet deactivation events.
'@param prevSheetName String. Name of the sheet that was deactivated.
'@export
Public Sub OnSheetDeactivate(ByVal prevSheetName As String)
    Dim prevsh As Worksheet

    On Error GoTo ErrHand

    EnsureTranslation

    'Update the listAuto only if the flag was set during editing
    If this.wkbNames Is Nothing Then Exit Sub
    If this.wkbNames.ValueAsString("RNG_UpdateListAuto") <> "yes" Then Exit Sub

    Set prevsh = SheetOrNothing(prevSheetName)
    If prevsh Is Nothing Then Exit Sub

    UpdateListAutoSheet prevsh
    this.wkbNames.SetValue "RNG_UpdateListAuto", "no"

    Exit Sub
ErrHand:
End Sub

Private Sub IEventLinelist_OnSheetDeactivate(ByVal prevSheetName As String)
    OnSheetDeactivate prevSheetName
End Sub


'@sub-title Handle selection change events on HList sheets.
'@param sh Worksheet. The worksheet where the selection changed.
'@param target Range. The newly selected cell(s).
'@export
Public Sub OnSelectionChange(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    Dim shHn As IHiddenNames
    Set shHn = HiddenNames.Create(sh)

    Dim shType As String
    shType = shHn.ValueAsString("sheet_type")
    If shType <> "HList" Then Exit Sub

    HandleSelectionHList sh, target, shHn
End Sub

Private Sub IEventLinelist_OnSelectionChange(ByVal sh As Worksheet, ByVal target As Range)
    OnSelectionChange sh, target
End Sub


'@sub-title Handle double-click events on analysis sheets.
'@param sh Worksheet. The worksheet where the double-click occurred.
'@param target Range. The double-clicked cell.
'@export
Public Sub OnDoubleClick(ByVal sh As Worksheet, ByVal target As Range)
    Dim rngName As String
    Dim shType As String

    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    On Error Resume Next
    rngName = target.Name.Name
    On Error GoTo 0

    shType = SheetTag(sh)
    If shType <> "SPT-Analysis" Then Exit Sub

    If (InStr(1, rngName, "INPUTSPTGEO_") > 0) Then
        LoadGeo 0
    ElseIf (InStr(1, rngName, "INPUTSPTHF_") > 0) Then
        LoadGeo 1
    End If
End Sub

Private Sub IEventLinelist_OnDoubleClick(ByVal sh As Worksheet, ByVal target As Range)
    OnDoubleClick sh, target
End Sub


'@sub-title Update filter tables used by analysis sheets.
'@param calculate Optional Boolean. When True (default), recalculates the active sheet.
'@export
Public Sub UpdateFilterTables(Optional ByVal calculate As Boolean = True)
    Dim sh As Worksheet
    Dim filtsh As Worksheet
    Dim lo As ListObject
    Dim destRng As Range
    Dim delRng As Range
    Dim loRng As Range
    Dim rowCounter As Long
    Dim filtLoHrng As Range
    Dim wsFunc As WorksheetFunction
    Dim shHn As IHiddenNames

    On Error GoTo ErrUpdate

    EnsureTranslation
    Set wsFunc = Application.WorksheetFunction

    For Each sh In this.hostBook.Worksheets
        If SheetTag(sh) = "HList" Then

            Set lo = sh.ListObjects(1)

            If Not lo.DataBodyRange Is Nothing Then
                Set loRng = lo.DataBodyRange

                Set shHn = HiddenNames.Create(sh)
                Set filtsh = this.hostBook.Worksheets(shHn.ValueAsString("filtered_sheet"))

                rowCounter = loRng.Rows.Count

                With filtsh
                    On Error Resume Next
                    .ListObjects(1).DataBodyRange.Delete
                    On Error GoTo ErrUpdate

                    .ListObjects(1).Resize .Range(lo.Range.Address)
                    Set destRng = .Range(loRng.Address)
                    Set filtLoHrng = .ListObjects(1).HeaderRowRange
                    Set delRng = Nothing
                End With

                destRng.Value = loRng.Value

                Do While rowCounter >= 1
                    If (loRng.Cells(rowCounter, 1).EntireRow.Hidden) Or _
                        (wsFunc.CountA(filtLoHrng.Offset(rowCounter)) = 0) Then
                        If delRng Is Nothing Then
                            Set delRng = filtLoHrng.Offset(rowCounter)
                        Else
                            Set delRng = Application.Union(delRng, filtLoHrng.Offset(rowCounter))
                        End If
                    End If
                    rowCounter = rowCounter - 1
                Loop

                If Not (delRng Is Nothing) Then delRng.Delete
            End If
        End If
    Next

    If calculate Then
        ActiveSheet.Calculate
        ActiveSheet.UsedRange.Calculate
        ActiveSheet.Columns("A:E").Calculate
    End If

    Exit Sub

ErrUpdate:
    If Not this.tradmess Is Nothing Then
        MsgBox this.tradmess.TranslatedValue("MSG_ErrUpdate") & ": " & Err.Description, _
               vbCritical + vbOKOnly
    End If
End Sub

Private Sub IEventLinelist_UpdateFilterTables(Optional ByVal calculate As Boolean = True)
    UpdateFilterTables calculate
End Sub


'@sub-title Update all list-auto dropdowns across all HList sheets.
'@export
Public Sub UpdateAllListAuto()
    Dim sh As Worksheet

    If this.hostBook Is Nothing Then Exit Sub

    For Each sh In this.hostBook.Worksheets
        If SheetTag(sh) = "HList" Then UpdateListAutoSheet sh
    Next
End Sub

Private Sub IEventLinelist_UpdateAllListAuto()
    UpdateAllListAuto
End Sub


'@sub-title Clear cached resources.
'@export
Public Sub ResetCaches()
    ResetInternalCaches
End Sub

Private Sub IEventLinelist_ResetCaches()
    ResetInternalCaches
End Sub


'@section Geo / Spatial UDF delegation
'===============================================================================

'@sub-title Concatenate geo cells up to a given admin level.
'@param cellRng Range. The first geo cell.
'@param geoLevel Byte. Number of admin levels to concatenate (1-4).
'@return String. Pipe-delimited concatenation of non-empty admin values.
'@export
Public Function GeoConcat(ByVal cellRng As Range, ByVal geoLevel As Byte) As String
    Dim concatValue As String
    Dim nonEmptyValue As Boolean

    Select Case geoLevel

    Case 1
        concatValue = cellRng.Value

    Case 2
        nonEmptyValue = (Not IsEmpty(cellRng)) And _
                        (Not IsEmpty(cellRng.Offset(, 1)))
        concatValue = IIf(nonEmptyValue, _
                          cellRng.Value & " | " & cellRng.Offset(, 1).Value, _
                          vbNullString)

    Case 3
        nonEmptyValue = (Not IsEmpty(cellRng)) And _
                        (Not IsEmpty(cellRng.Offset(, 1))) And _
                        (Not IsEmpty(cellRng.Offset(, 2)))
        concatValue = IIf(nonEmptyValue, _
                          cellRng.Value & " | " & cellRng.Offset(, 1).Value & _
                          " | " & cellRng.Offset(, 2).Value, _
                          vbNullString)

    Case 4
        nonEmptyValue = (Not IsEmpty(cellRng)) And _
                        (Not IsEmpty(cellRng.Offset(, 1))) And _
                        (Not IsEmpty(cellRng.Offset(, 2))) And _
                        (Not IsEmpty(cellRng.Offset(, 3)))
        concatValue = IIf(nonEmptyValue, _
                          cellRng.Value & " | " & cellRng.Offset(, 1).Value & _
                          " | " & cellRng.Offset(, 2).Value & _
                          " | " & cellRng.Offset(, 3).Value, _
                          vbNullString)

    Case Else
        concatValue = cellRng.Value

    End Select

    GeoConcat = concatValue
End Function

Private Function IEventLinelist_GeoConcat(ByVal cellRng As Range, _
                                          ByVal geoLevel As Byte) As String
    IEventLinelist_GeoConcat = GeoConcat(cellRng, geoLevel)
End Function


'@sub-title Return the top admin name for a spatial variable.
'@param adminLevel String. Admin level label (e.g. "adm1").
'@param adminOrder Integer. Rank order (1 = top, 2 = second, etc.).
'@param varName String. Variable name (with prefix, e.g. "adm1_concat_adm1_myvar").
'@param tabId Optional String. Spatial table identifier.
'@return String. Admin name at the given rank.
'@export
Public Function TopAdmin(ByVal adminLevel As String, _
                         ByVal adminOrder As Integer, _
                         ByVal varName As String, _
                         Optional ByVal tabId As String = vbNullString) As String
    Dim adminName As String
    Dim actualVarName As String

    EnsureGeo
    EnsureSpatial

    If this.geo Is Nothing Then Exit Function
    If this.sp Is Nothing Then Exit Function

    actualVarName = Split(varName, "_")(2)
    adminName = this.geo.AdminCode(adminLevel)
    TopAdmin = this.sp.TopGeoValue(adminName, adminOrder, actualVarName, tabId)
End Function

Private Function IEventLinelist_TopAdmin(ByVal adminLevel As String, _
                                         ByVal adminOrder As Integer, _
                                         ByVal varName As String, _
                                         Optional ByVal tabId As String = vbNullString) As String
    IEventLinelist_TopAdmin = TopAdmin(adminLevel, adminOrder, varName, tabId)
End Function


'@sub-title Return the population for the top admin of a spatial variable.
'@param adminLevel String. Admin level label (e.g. "adm1").
'@param adminOrder Integer. Rank order (1 = top, 2 = second, etc.).
'@param varName String. Variable name (with prefix "concat_adm1_").
'@param tabId Optional String. Spatial table identifier.
'@return Long. Population count at the given rank.
'@export
Public Function TopPop(ByVal adminLevel As String, _
                       ByVal adminOrder As Integer, _
                       ByVal varName As String, _
                       Optional ByVal tabId As String = vbNullString) As Long
    Dim adminName As String
    Dim actualVarName As String
    Dim pop As Long

    EnsureGeo
    EnsureSpatial

    If this.geo Is Nothing Then Exit Function
    If this.sp Is Nothing Then Exit Function

    ' 12 is the length of the "concat_adm1_" prefix
    actualVarName = Right(varName, (Len(varName) - 12))
    adminName = this.geo.AdminCode(adminLevel)

    pop = 0
    On Error Resume Next
    pop = CLng(this.sp.TopGeoValue(adminName, adminOrder, actualVarName, tabId, 2))
    On Error GoTo 0

    TopPop = pop
End Function

Private Function IEventLinelist_TopPop(ByVal adminLevel As String, _
                                       ByVal adminOrder As Integer, _
                                       ByVal varName As String, _
                                       Optional ByVal tabId As String = vbNullString) As Long
    IEventLinelist_TopPop = TopPop(adminLevel, adminOrder, varName, tabId)
End Function


'@sub-title Return the top health facility name for a spatial variable.
'@param adminOrder Integer. Rank order (1 = top, 2 = second, etc.).
'@param varName String. Variable name (with "hf_" prefix).
'@param tabId Optional String. Spatial table identifier.
'@return String. Health facility name at the given rank.
'@export
Public Function TopHF(ByVal adminOrder As Integer, _
                      ByVal varName As String, _
                      Optional ByVal tabId As String = vbNullString) As String
    Dim actualVarName As String

    EnsureSpatial

    If this.sp Is Nothing Then Exit Function

    ' 3 is the length of the "hf_" prefix
    actualVarName = Right(varName, (Len(varName) - 3))
    TopHF = this.sp.TopHFValue(adminOrder, actualVarName, tabId)
End Function

Private Function IEventLinelist_TopHF(ByVal adminOrder As Integer, _
                                      ByVal varName As String, _
                                      Optional ByVal tabId As String = vbNullString) As String
    IEventLinelist_TopHF = TopHF(adminOrder, varName, tabId)
End Function


'@section Private — HList change handler
'===============================================================================

'@sub-title Handle value changes on HList sheets.
'@details
'Dispatches based on the target cell's position relative to the table:
'geo cascades (geo1-geo4 controls), editable labels (label row),
'list_auto flagging, go-to-section navigation, header protection
'(header row), and multiple choice updates.
'@param sh Worksheet. The HList worksheet.
'@param target Range. The edited cell(s).
'@param shHn IHiddenNames. The worksheet-level hidden names (shared from caller).
Private Sub HandleHListChange(ByRef sh As Worksheet, ByRef target As Range, _
                              ByRef shHn As IHiddenNames)
    Dim adminTable As BetterArray
    Dim adminNames As BetterArray
    Dim hRng As Range
    Dim calcRng As Range
    Dim rng As Range
    Dim cellRng As Range
    Dim varControl As String
    Dim varLabel As String
    Dim varName As String
    Dim varSubLabel As String
    Dim tableName As String
    Dim varEditable As String
    Dim sectionName As String
    Dim targetColumn As Long
    Dim startLine As Long
    Dim nbOffset As Long
    Dim choiSep As String

    On Error GoTo ErrHand

    Set hRng = sh.ListObjects(1).HeaderRowRange
    Set adminNames = New BetterArray
    adminNames.LowerBound = 1
    Set adminTable = New BetterArray

    EnsureTranslation

    tableName = shHn.ValueAsString("table_name")
    targetColumn = target.Column
    startLine = hRng.Row + 1
    varName = sh.Cells(hRng.Row, targetColumn).Value
    varControl = shHn.ValueAsString(varName & " -- control")

    If (target.Row >= startLine) Then

        nbOffset = target.Row - hRng.Row
        Set calcRng = hRng.Offset(nbOffset)
        calcRng.Calculate

        'Geo variables
        If (varControl = "geo1") Or _
           (varControl = "geo2") Or _
           (varControl = "geo3") Or _
           (varControl = "geo4") Then

            If (target.Value = vbNullString) Then Exit Sub

            EnsureDropdown
            EnsureGeo

            If this.drop Is Nothing Then Exit Sub
            If this.geo Is Nothing Then Exit Sub

            Select Case varControl

            Case "geo1"
                this.drop.ClearList "admin2"
                target.Offset(, 1).Value = vbNullString
                this.drop.ClearList "admin3"
                target.Offset(, 2).Value = vbNullString
                this.drop.ClearList "admin4"
                target.Offset(, 3).Value = vbNullString

                Set adminTable = this.geo.GeoLevel(LevelAdmin2, _
                                                   GeoScopeAdmin, _
                                                   target.Value)
                this.drop.Update adminTable, "admin2"

            Case "geo2"
                this.drop.ClearList "admin3"
                target.Offset(, 1).Value = vbNullString
                this.drop.ClearList "admin4"
                target.Offset(, 2).Value = vbNullString
                adminNames.Push target.Offset(, -1).Value, target.Value
                Set adminTable = this.geo.GeoLevel(LevelAdmin3, _
                                                   GeoScopeAdmin, _
                                                   adminNames)
                this.drop.Update adminTable, "admin3"

            Case "geo3"
                this.drop.ClearList "admin4"
                target.Offset(, 1).Value = vbNullString
                adminNames.Push target.Offset(, -2).Value, _
                               target.Offset(, -1).Value, _
                               target.Value
                Set adminTable = this.geo.GeoLevel(LevelAdmin4, _
                                                   GeoScopeAdmin, _
                                                   adminNames)
                this.drop.Update adminTable, "admin4"

            End Select

            Exit Sub
        End If
    End If

    'Update variables with editable column (label row = header - 1)
    If (target.Row = hRng.Row - 1) Then

        varEditable = shHn.ValueAsString(varName & " -- editable")
        If (varEditable <> "yes") Then Exit Sub

        EnsureDictionary
        EnsureVariables
        If this.vars Is Nothing Then Exit Sub

        varSubLabel = this.vars.Value(varName:=varName, colName:="sub label")
        varLabel = Replace(target.Value, varSubLabel, vbNullString)
        varLabel = Replace(varLabel, Chr(10), vbNullString)

        this.vars.SetValue varName:=varName, colName:="main label", _
                           newValue:=varLabel
        Exit Sub
    End If

    'Update the list auto flag
    If (target.Row >= startLine) And _
       (shHn.ValueAsString(varName & " -- listauto") <> vbNullString) Then

        If Not this.wkbNames Is Nothing Then
            this.wkbNames.SetValue "RNG_UpdateListAuto", "yes"
        End If
        Exit Sub
    End If

    'GoTo section
    Set rng = sh.Range(tableName & "_" & GOTOSECCODE)

    If Not (Intersect(target, rng) Is Nothing) Then
        If Not this.wkbNames Is Nothing Then
            sectionName = Replace(target.Value, _
                             this.wkbNames.ValueAsString("RNG_GoToSection") & ": ", _
                             vbNullString)
        End If

        Set hRng = hRng.Offset(-3)
        Set cellRng = hRng.Find(What:=sectionName, LookAt:=xlWhole, MatchCase:=True)
        If Not cellRng Is Nothing Then cellRng.Activate
        Exit Sub
    End If

    'Avoid modifying headers of the table (with variable names)
    If (target.Row = hRng.Row) Then
        target.Value = target.Offset(-1).Name.Name
        If Not this.tradmess Is Nothing Then
            MsgBox this.tradmess.TranslatedValue("MSG_NotModify"), _
                   vbOKOnly + vbCritical, _
                   this.tradmess.TranslatedValue("MSG_Error")
        End If
        Exit Sub
    End If

    'Update multiple choices dropdown
    If InStr(1, varControl, "choice_multiple") = 1 Then
        On Error Resume Next
        choiSep = Split(varControl, "(" & Chr(34))(1)
        choiSep = Replace(choiSep, Chr(34) & ")", vbNullString)
        On Error GoTo 0

        If (choiSep = vbNullString) Or _
           (InStr(1, choiSep, "choice_multiple") = 1) Then _
            choiSep = ", "

        UpdateMultipleChoice target, choiSep
        Exit Sub
    End If

ErrHand:
End Sub


'@section Private — VList change handler
'===============================================================================

'@sub-title Handle value changes on VList sheets.
'@details
'Handles editable label updates, multiple choice updates, and go-to-section
'navigation on vertical list sheets.
'@param sh Worksheet. The VList worksheet.
'@param target Range. The edited cell(s).
'@param shHn IHiddenNames. The worksheet-level hidden names (shared from caller).
Private Sub HandleVListChange(ByRef sh As Worksheet, ByRef target As Range, _
                              ByRef shHn As IHiddenNames)
    Dim rng As Range
    Dim rngLook As Range
    Dim varLabel As String
    Dim tableName As String
    Dim varSubLabel As String
    Dim varEditable As String
    Dim varName As String
    Dim varControl As String
    Dim choiSep As String

    On Error GoTo ErrHandler

    EnsureTranslation

    tableName = shHn.ValueAsString("table_name")

    'Calculate the range where the values are entered
    Set rng = sh.Range(tableName & "_" & "PLAGEVALUES")
    rng.Calculate

    'Test if editable Labels have been modified
    Set rng = sh.Range(tableName & "_START")

    'Update variables with editable column
    If (target.Column = rng.Offset(, -1).Column) Then
        varName = (target.Offset(, 1).Name.Name)
        varEditable = shHn.ValueAsString(varName & " -- editable")
        If (varEditable <> "yes") Then Exit Sub

        EnsureDictionary
        EnsureVariables
        If this.vars Is Nothing Then Exit Sub

        varSubLabel = this.vars.Value(varName:=varName, colName:="sub label")
        varLabel = Replace(target.Value, varSubLabel, vbNullString)
        varLabel = Replace(varLabel, Chr(10), vbNullString)
        this.vars.SetValue varName:=varName, colName:="main label", _
                           newValue:=varLabel
        Exit Sub
    End If

    'Update the choice multiple
    If target.Column = rng.Column Then
        varControl = target.Offset(, 1).Value
        If InStr(1, varControl, "choice_multiple") = 1 Then
            On Error Resume Next
            choiSep = Split(varControl, "(" & Chr(34))(1)
            choiSep = Replace(choiSep, Chr(34) & ")", vbNullString)
            On Error GoTo 0

            If (choiSep = vbNullString) Or _
               (InStr(1, choiSep, "choice_multiple") = 1) Then _
                choiSep = ", "

            UpdateMultipleChoice target, choiSep
            Exit Sub
        End If
    End If

    'Jump to GoTo section
    Set rng = sh.Range(tableName & "_" & GOTOSECCODE)
    If Not Intersect(target, rng) Is Nothing Then
        If Not this.wkbNames Is Nothing Then
            varLabel = Replace(target.Value, _
                               this.wkbNames.ValueAsString("RNG_GoToSection") & ": ", _
                               vbNullString)
        End If
        Set rngLook = sh.Cells.Find(What:=varLabel, LookAt:=xlWhole, MatchCase:=True)
        If Not rngLook Is Nothing Then rngLook.Activate
    End If

    Exit Sub
ErrHandler:
End Sub


'@section Private — Analysis change handler
'===============================================================================

'@sub-title Handle value changes on analysis sheets.
'@details
'Processes go-to-section navigation for all analysis types. For spatial
'analysis sheets, dispatches to local spatial table update methods
'(UpdateSingleSpTable, DevideByPopulation, FormatDevidePop)
'based on named range patterns.
'@param sh Worksheet. The analysis worksheet.
'@param target Range. The edited cell(s).
'@param analysisType String. The sheet type tag (Uni-Bi-Analysis, TS-Analysis, etc.).
Private Sub HandleAnalysisChange(ByRef sh As Worksheet, _
                                  ByRef target As Range, _
                                  ByRef analysisType As String)
    Dim rng As Range
    Dim rngLook As Range
    Dim sLabel As String
    Dim goToSection As String
    Dim goToHeader As String
    Dim goToGraph As String
    Dim rngName As String

    EnsureTranslation

    'Range name if it exists
    On Error Resume Next
    rngName = target.Name.Name
    On Error GoTo 0

    On Error GoTo ErrHandler

    Select Case analysisType

    Case "Uni-Bi-Analysis"
        Set rng = sh.Range("ua_go_to_section")

    Case "TS-Analysis", "SPT-Analysis"
        sh.Calculate
        sh.UsedRange.Calculate
        sh.Columns("A:E").Calculate
        If InStr(1, rngName, "ts_go_to_section") > 0 Then Set rng = target
        If InStr(1, rngName, "spt_go_to_section") > 0 Then Set rng = target

    Case "SP-Analysis"
        Set rng = sh.Range("sp_go_to_section")
        'Dispatch spatial events to local private methods
        If InStr(1, rngName, "ADM_DROPDOWN_") > 0 Then UpdateSingleSpTable rngName
        If InStr(1, rngName, "POPFACT_") > 0 Then DevideByPopulation rngName
        If InStr(1, rngName, "DEVIDEPOP_") > 0 Then FormatDevidePop rngName

    End Select

    If rng Is Nothing Then Exit Sub
    If Intersect(target, rng) Is Nothing Then Exit Sub

    If Not this.wkbNames Is Nothing Then
        goToSection = this.wkbNames.ValueAsString("RNG_GoToSection")
        goToHeader = this.wkbNames.ValueAsString("RNG_GoToHeader")
        goToGraph = this.wkbNames.ValueAsString("RNG_GoToGraph")
    End If

    sLabel = Replace(target.Value, goToSection & ": ", vbNullString)
    sLabel = Replace(sLabel, goToHeader & ": ", vbNullString)
    sLabel = Replace(sLabel, goToGraph & ": ", vbNullString)

    Set rngLook = ActiveSheet.Cells.Find(What:=sLabel, LookIn:=xlValues, LookAt:=xlWhole, _
                                         MatchCase:=True, SearchFormat:=False)

    If Not rngLook Is Nothing Then rngLook.Activate

    Exit Sub
ErrHandler:
End Sub


'@section Private — Spatial table update handlers
'===============================================================================

'@sub-title Update formulas in a single spatial table when user changes admin level.
'@details
'Replaces the admin reference in all formulas within the table's outer values
'range. Uses cached geo/passwords instances instead of creating new ones.
'Busy state is managed by the calling LinelistEventsManager.
'@param rngName String. Named range of the admin dropdown (e.g. "ADM_DROPDOWN_sp1").
Private Sub UpdateSingleSpTable(ByVal rngName As String)
    Dim tabId As String
    Dim adminName As String
    Dim selectedAdmin As String
    Dim formulaValue As String
    Dim prevAdmValue As String
    Dim cellRng As Range
    Dim rng As Range
    Dim sh As Worksheet
    Dim hasFormula As Boolean

    EnsurePasswords
    EnsureGeo

    If this.pass Is Nothing Then Exit Sub
    If this.geo Is Nothing Then Exit Sub

    this.pass.UnProtect "_active"

    Set sh = ActiveSheet
    selectedAdmin = sh.Range(rngName).Value
    adminName = this.geo.AdminCode(selectedAdmin)

    tabId = Replace(rngName, "ADM_DROPDOWN_", vbNullString)
    prevAdmValue = sh.Range("PREVIOUS_ADM_" & tabId).Value

    Set rng = sh.Range("OUTER_VALUES_" & tabId)

    For Each cellRng In rng
        hasFormula = False
        formulaValue = cellRng.FormulaArray

        If formulaValue = vbNullString Then
            formulaValue = cellRng.Formula
            hasFormula = True
        End If

        If InStr(1, formulaValue, "concat_" & prevAdmValue) > 0 Then
            formulaValue = Replace(formulaValue, "concat_" & prevAdmValue, _
                                   "concat_" & adminName)

            If hasFormula Then
                cellRng.Formula = formulaValue
            Else
                cellRng.FormulaArray = formulaValue
            End If
        End If
    Next

    sh.Range("PREVIOUS_ADM_" & tabId).Value = adminName
    sh.Columns("C").EntireColumn.AutoFit
    rng.Calculate

    this.pass.Protect "_active", True
End Sub


'@sub-title Divide or revert all computed values by the population.
'@details
'Iterates through the outer values range of a spatial table and modifies
'formulas to include or remove population division. Uses cached
'geo/spatial/passwords instances.
'@param rngName String. Named range of the population factor (e.g. "POPFACT_sp1").
'@param revertBack Optional Boolean. True to remove the division and restore original formulas.
Private Sub DevideByPopulation(ByVal rngName As String, _
                               Optional ByVal revertBack As Boolean = False)

    Dim sh As Worksheet
    Dim hasFormula As Boolean
    Dim factorMult As Long
    Dim prevFact As Long
    Dim rng As Range
    Dim cellRng As Range
    Dim formulaValue As String
    Dim adminName As String
    Dim rowRng As Range
    Dim selectedAdmin As String
    Dim popValue As String
    Dim tabId As String

    EnsurePasswords
    EnsureGeo
    EnsureSpatial

    If this.pass Is Nothing Then Exit Sub
    If this.geo Is Nothing Then Exit Sub
    If this.sp Is Nothing Then Exit Sub

    Set sh = ActiveSheet
    this.pass.UnProtect "_active"

    tabId = Replace(rngName, "POPFACT_", vbNullString)
    prevFact = sh.Range("POPPREVFACT_" & tabId).Value

    factorMult = 100
    On Error Resume Next
    factorMult = CLng(Application.WorksheetFunction.Trim(sh.Range(rngName).Value))
    On Error GoTo ErrDivide

    selectedAdmin = sh.Range("ADM_DROPDOWN_" & tabId).Value
    adminName = this.geo.AdminCode(selectedAdmin)

    Set rng = sh.Range("OUTER_VALUES_" & tabId)
    Set rowRng = sh.Range("ROW_CATEGORIES_" & tabId)

    this.sp.Sort tabId:=tabId, onAR:=(Not revertBack)

    For Each cellRng In rng
        hasFormula = False
        formulaValue = cellRng.FormulaArray

        If formulaValue = vbNullString Then
            formulaValue = cellRng.Formula
            hasFormula = True
        End If

        If (InStr(1, formulaValue, "concat_" & adminName) > 0) And _
           (cellRng.Column > rowRng.Column) Then

            popValue = sh.Cells(cellRng.Row, rowRng.Column - 1).Address

            If (Not revertBack) And (prevFact = 0) Then
                formulaValue = Replace(formulaValue, "=", vbNullString)
                formulaValue = "= " & factorMult & "*" & formulaValue & "/" & popValue
            ElseIf (Not revertBack) And (prevFact <> 0) Then
                formulaValue = Replace(formulaValue, CStr(prevFact), CStr(factorMult))
            ElseIf prevFact <> 0 Then
                formulaValue = Replace(formulaValue, CStr(prevFact) & "*", vbNullString)
                formulaValue = Replace(formulaValue, "/" & popValue, vbNullString)
            End If

            If hasFormula Then
                cellRng.Formula = formulaValue
            Else
                cellRng.FormulaArray = formulaValue
            End If
        End If
    Next

    If revertBack Then
        sh.Range("POPPREVFACT_" & tabId).Value = 0
    Else
        sh.Range("POPPREVFACT_" & tabId).Value = factorMult
    End If

    rng.Calculate
    this.pass.Protect "_active"

ErrDivide:
End Sub


'@sub-title Toggle population division formatting.
'@details
'Shows or hides the population factor cell and label based on the
'divide/no-divide dropdown selection, then delegates to DevideByPopulation
'to apply or remove the formula modifications.
'@param rngName String. Named range of the divide/no-divide dropdown.
Private Sub FormatDevidePop(ByVal rngName As String)
    Dim sh As Worksheet
    Dim tabId As String

    EnsurePasswords
    EnsureTranslation

    If this.pass Is Nothing Then Exit Sub
    If this.wkbNames Is Nothing Then Exit Sub

    Set sh = ActiveSheet
    this.pass.UnProtect "_active"

    tabId = Replace(rngName, "DEVIDEPOP_", vbNullString)

    If sh.Range(rngName).Value = this.wkbNames.ValueAsString("RNG_NoDevide") Then
        sh.Range("POPFACT_" & tabId).Font.Color = vbWhite
        sh.Range("POPFACT_" & tabId).Locked = True
        sh.Range("POPFACT_" & tabId).FormulaHidden = True
        sh.Range("POPFACTLABEL_" & tabId).Font.Color = vbWhite
        sh.Range("POPFACTLABEL_" & tabId).Locked = True

        DevideByPopulation rngName:="POPFACT_" & tabId, revertBack:=True

    ElseIf sh.Range(rngName).Value = this.wkbNames.ValueAsString("RNG_Devide") Then
        sh.Range("POPFACT_" & tabId).Font.Color = vbBlack
        sh.Range("POPFACT_" & tabId).Locked = False
        sh.Range("POPFACT_" & tabId).FormulaHidden = False
        sh.Range("POPFACTLABEL_" & tabId).Font.Color = vbBlack
        sh.Range("POPFACTLABEL_" & tabId).Locked = False

        DevideByPopulation "POPFACT_" & tabId
    End If

    this.pass.Protect "_active", True
End Sub


'@section Private — Selection change handler
'===============================================================================

'@sub-title Handle selection changes on HList sheets.
'@details
'Updates geo dropdown validation lists when the user navigates to a geo2, geo3,
'or geo4 cell.
'@param sh Worksheet. The HList worksheet.
'@param target Range. The newly selected cell(s).
'@param shHn IHiddenNames. The worksheet-level hidden names (shared from caller).
Private Sub HandleSelectionHList(ByRef sh As Worksheet, ByRef target As Range, _
                                ByRef shHn As IHiddenNames)
    Dim targetColumn As Long
    Dim nbOffset As Long
    Dim hRng As Range
    Dim calcRng As Range
    Dim startLine As Long
    Dim varControl As String
    Dim varName As String
    Dim adminTable As BetterArray
    Dim adminNames As BetterArray

    On Error GoTo ErrHand

    Set hRng = sh.ListObjects(1).HeaderRowRange
    startLine = hRng.Row + 1

    'First test if we are on a valid line
    If (target.Row < startLine) Then Exit Sub

    'Calculate the line
    nbOffset = target.Row - hRng.Row
    Set calcRng = hRng.Offset(nbOffset)
    calcRng.Calculate

    'Test for geo control (Exit if not the case)
    targetColumn = target.Column
    varName = sh.Cells(hRng.Row, targetColumn).Value
    varControl = shHn.ValueAsString(varName & " -- control")

    If (varControl <> "geo2") And _
       (varControl <> "geo3") And _
       (varControl <> "geo4") Then _
        Exit Sub

    EnsureTranslation
    EnsureGeo
    EnsureDropdown

    If this.geo Is Nothing Then Exit Sub
    If this.drop Is Nothing Then Exit Sub

    Set adminNames = New BetterArray
    adminNames.LowerBound = 1

    Select Case varControl

    Case "geo2"
        this.drop.ClearList "admin2"
        Set adminTable = this.geo.GeoLevel(LevelAdmin2, GeoScopeAdmin, target.Offset(, -1).Value)
        this.drop.Update adminTable, "admin2"

    Case "geo3"
        this.drop.ClearList "admin3"
        adminNames.Push target.Offset(, -2).Value, target.Offset(, -1).Value
        Set adminTable = this.geo.GeoLevel(LevelAdmin3, GeoScopeAdmin, adminNames)
        this.drop.Update adminTable, "admin3"

    Case "geo4"
        this.drop.ClearList "admin4"
        adminNames.Push target.Offset(, -3).Value, target.Offset(, -2).Value, target.Offset(, -1).Value
        Set adminTable = this.geo.GeoLevel(LevelAdmin4, GeoScopeAdmin, adminNames)
        this.drop.Update adminTable, "admin4"

    End Select

ErrHand:
End Sub


'@section Private — List auto helpers
'===============================================================================

'@sub-title Update the list-auto dropdowns for a single sheet.
'@details
'Checks the has_listauto flag to skip sheets without list-auto variables.
'Iterates header columns and updates dropdown lists from unique data values.
'@param sh Worksheet. The HList sheet to update.
Private Sub UpdateListAutoSheet(ByRef sh As Worksheet)
    Dim varName As String
    Dim arrTable As BetterArray
    Dim hRng As Range
    Dim shHn As IHiddenNames
    Dim colIdx As Long

    Set shHn = HiddenNames.Create(sh)

    'Early exit if this sheet has no listauto variables
    If shHn.ValueAsString("has_listauto") <> "yes" Then Exit Sub

    Set arrTable = New BetterArray
    EnsureDropdown
    If this.drop Is Nothing Then Exit Sub

    Set hRng = sh.ListObjects(1).HeaderRowRange

    For colIdx = 1 To hRng.Columns.Count
        varName = hRng.Cells(1, colIdx).Value
        If varName = vbNullString Then GoTo NextCol

        If shHn.ValueAsString(varName & " -- listauto") <> vbNullString Then
            If this.drop.Exists(varName) Then
                arrTable.FromExcelRange hRng.Cells(1, colIdx).Offset(1), _
                                        DetectLastColumn:=False, _
                                        DetectLastRow:=True
                Set arrTable = GetUniqueBA(arrTable)
                this.drop.ClearList varName
                this.drop.Update arrTable, varName
                this.drop.Sort varName, xlDescending
            End If
        End If
NextCol:
    Next colIdx
End Sub


'@section Private — Multiple choice helper
'===============================================================================

'@sub-title Handle multiple choice dropdown value toggling.
'@param target Range. The edited cell.
'@param choiSep String. Separator between choices (e.g. ", ").
Private Sub UpdateMultipleChoice(ByRef target As Range, ByRef choiSep As String)
    Dim actualValue As String
    Dim prevValue As String
    Dim prevTab As BetterArray
    Dim actTab As BetterArray

    Set prevTab = New BetterArray
    Set actTab = New BetterArray

    If IsEmpty(target) Then Exit Sub

    actualValue = target.Value

    On Error Resume Next
    Application.Undo
    On Error GoTo ErrHandler

    If IsEmpty(target) Then GoTo KeepValue

    prevValue = target.Value
    prevTab.Items = Split(prevValue, choiSep)
    actTab.Items = Split(actualValue, choiSep)

    'Length reduction, keep values
    If (actTab.Length > 1) And (actTab.Length < prevTab.Length) Then GoTo KeepValue

    'There is no length reduction. Test for presence of actual element
    If prevTab.Includes(actualValue) Then
        target.Value = prevValue
        Exit Sub
    End If

    'Length of actual Tab is One, previous Tab does not include new value
    target.Value = prevValue & choiSep & actualValue
    Exit Sub

KeepValue:
    target.Value = actualValue

ErrHandler:
End Sub


'@section Private — Unique values helper
'===============================================================================

'@sub-title Return unique values from a BetterArray.
'@param inputTable BetterArray. Source array.
'@param Sort Optional Boolean. Whether to sort the result.
'@return BetterArray. Array of unique values.
Private Function GetUniqueBA(ByRef inputTable As BetterArray, _
                             Optional ByVal Sort As Boolean = False) As BetterArray
    Dim tableValue As String
    Dim counter As Long
    Dim outTable As BetterArray

    Set outTable = New BetterArray
    outTable.LowerBound = 1

    If inputTable.Length > 0 Then
        For counter = inputTable.LowerBound To inputTable.UpperBound
            On Error Resume Next
            tableValue = Application.WorksheetFunction.Trim( _
                             inputTable.Item(counter))
            On Error GoTo 0

            If (tableValue <> vbNullString) And _
               (Not outTable.Includes(tableValue)) Then _
                outTable.Push tableValue
        Next
    End If

    If Sort Then outTable.Sort
    Set GetUniqueBA = outTable.Clone()
End Function


'@section Private — Sheet metadata helpers
'===============================================================================

'@sub-title Get the sheet type tag from worksheet-level HiddenNames.
'@param sh Worksheet. The worksheet to query.
'@return String. The sheet type tag (HList, VList, Uni-Bi-Analysis, etc.).
Private Function SheetTag(ByRef sh As Worksheet) As String
    Dim shHn As IHiddenNames
    Set shHn = HiddenNames.Create(sh)
    SheetTag = shHn.ValueAsString("sheet_type")
End Function


'@section Private — Lazy initialisation
'===============================================================================

'@sub-title Lazily initialise translation, message translator, and hidden names.
Private Sub EnsureTranslation()
    If Not (this.lltrads Is Nothing) Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = SheetOrNothing(LLSHEET)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.lltrads = LLTranslation.Create(sh)
    If Not this.lltrads Is Nothing Then
        Set this.tradmess = this.lltrads.TransObject()
    End If
    Set this.wkbNames = HiddenNames.Create(this.hostBook)
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the dropdown manager.
Private Sub EnsureDropdown()
    If Not (this.drop Is Nothing) Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = SheetOrNothing(DROPDOWNSHEET)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.drop = DropdownLists.Create(sh)
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the geo helper.
Private Sub EnsureGeo()
    If Not (this.geo Is Nothing) Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = SheetOrNothing(GEOSHEET)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.geo = LLGeo.Create(sh)
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the spatial helper.
Private Sub EnsureSpatial()
    If Not (this.sp Is Nothing) Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = SheetOrNothing(SPATIALSHEET)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.sp = LLSpatial.Create(sh)
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the password manager.
Private Sub EnsurePasswords()
    If Not (this.pass Is Nothing) Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = SheetOrNothing(PASSSHEET)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.pass = Passwords.Create(sh)
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the dictionary and variables helpers.
Private Sub EnsureDictionary()
    If Not (this.dict Is Nothing) Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = SheetOrNothing(DICTSHEET)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.dict = LLdictionary.Create(sh, 1, 1)
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the variables helper from the cached dictionary.
Private Sub EnsureVariables()
    If Not (this.vars Is Nothing) Then Exit Sub

    EnsureDictionary
    If this.dict Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.vars = LLVariables.Create(this.dict)
    On Error GoTo 0
End Sub

'@sub-title Retrieve a worksheet from the host workbook or Nothing.
'@param sheetName String. Name of the worksheet to retrieve.
'@return Worksheet. The worksheet, or Nothing when not found.
Private Function SheetOrNothing(ByRef sheetName As String) As Worksheet
    Dim sh As Worksheet

    If this.hostBook Is Nothing Then Exit Function

    On Error Resume Next
    Set sh = this.hostBook.Worksheets(sheetName)
    On Error GoTo 0

    Set SheetOrNothing = sh
End Function


'@section Private — Cache management
'===============================================================================

'@sub-title Clear all lazily-cached domain managers.
Private Sub ResetInternalCaches()
    Set this.tradmess = Nothing
    Set this.lltrads = Nothing
    Set this.wkbNames = Nothing
    Set this.drop = Nothing
    Set this.geo = Nothing
    Set this.sp = Nothing
    Set this.dict = Nothing
    Set this.vars = Nothing
    Set this.pass = Nothing
End Sub


'@section Private — Error handling
'===============================================================================

'@sub-title Raise a ProjectError-based exception.
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errNumber As Long, ByRef message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub
