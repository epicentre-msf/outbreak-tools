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
'Each manager below has a "tried" flag beside it. A manager is built once per
'session, and a build that fails leaves its field Nothing. Without the flag the
'guard "the field is Nothing" reads as "never tried", so every keystroke rebuilt
'and re-failed the same manager -- for the translation helper that meant a walk
'of the workbook Names collection per event. The flag says the attempt was made.
'ResetInternalCaches clears the fields and the flags together, so a caller that
'fixed the workbook gets a fresh attempt.
Private Type TEventLinelist
    hostBook        As Workbook           'Linelist workbook
    tradmess        As TranslationObject 'Cached message translations
    lltrads         As LLTranslation     'Cached translation helper
    wkbNames        As HiddenNames       'Cached workbook-level hidden names
    sheetStores     As Collection        'Worksheet hidden names, keyed on sheet name
    drop            As DropdownLists     'Cached dropdown manager
    geo             As LLGeo             'Cached geo helper
    sp              As LLSpatial          'Cached spatial helper
    dict            As LLdictionary      'Cached dictionary helper
    vars            As LLVariables       'Cached variable lookup helper
    pass            As Passwords         'Cached password manager
    triedTranslation As Boolean
    triedWkbNames    As Boolean
    triedDropdown    As Boolean
    triedGeo         As Boolean
    triedSpatial     As Boolean
    triedDictionary  As Boolean
    triedVariables   As Boolean
    triedPasswords   As Boolean
    lastEditKey      As String  'Sheet and address of the last single-cell selection or edit
    lastEditValue    As String  'What that cell held at that moment
    lastEditKnown    As Boolean 'The two fields above hold a real reading
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

' The message code every box title is read from.
Private Const MSGTITLE          As String = "MSG_Error"

' What OnDoubleClick answers when the click asked for no geo picker. Every
' other answer it gives is a GeoScope value, and those start at 0, so a caller
' reads any answer below zero as "nothing to open".
Private Const GEOSCOPENONE      As Long = -1


'@section Factory
'===============================================================================

'@sub-title Build a configured EventLinelist instance bound to a host workbook.
'@param hostBook Workbook. The linelist workbook.
'@return EventLinelist. Configured service instance.
'@throws ProjectError.ObjectNotInitialized When hostBook is Nothing.
'@export
Public Function Create(ByVal hostBook As Workbook) As EventLinelist
    Dim instance As EventLinelist

    Set instance = New EventLinelist
    instance.Configure hostBook
    Set Create = instance
End Function


'@section Public operations
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


'@sub-title Handle worksheet change events.
'@param sh Worksheet. The worksheet raising the change event.
'@param target Range. The edited cells.
'@export
Public Sub OnSheetChange(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    'An edit can add or rename a worksheet name, so the held store of the sheet
    'that raised the change is dropped and read again here.
    DropSheetStore sh.Name

    Dim shHn As HiddenNames
    Set shHn = SheetStore(sh)
    If shHn Is Nothing Then Exit Sub

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

    'What the cell holds once its handlers have run is the previous value of
    'the next edit to the same cell. The multiple choice toggle reads it back.
    RememberCellValue target
End Sub


'@sub-title Handle worksheet deactivation events.
'@details
'The handler is written into the code module of each HList sheet, so the sheet
'that was left arrives as an object and one flag is all this reads. It wants the
'workbook hidden names alone: the labels of the translation stack have no part
'in it.
'@param prevsh Worksheet. The sheet that was deactivated.
'@export
Public Sub OnSheetDeactivate(ByVal prevsh As Worksheet)
    On Error GoTo ErrHand

    If prevsh Is Nothing Then Exit Sub

    EnsureWorkbookNames

    'Update the listAuto only if the flag was set during editing
    If this.wkbNames Is Nothing Then Exit Sub
    If this.wkbNames.ValueAsString("RNG_UpdateListAuto") <> "yes" Then Exit Sub

    UpdateListAutoSheet prevsh
    this.wkbNames.SetValue "RNG_UpdateListAuto", "no"

    Exit Sub
ErrHand:
End Sub


'@sub-title Handle selection change events on HList sheets.
'@details
'The sheet is an HList sheet by construction. Worksheet_SelectionChange is
'written into the code module of the HList sheets alone, so a "sheet_type" read
'here can only ever answer HList, and it cost a walk of the sheet Names
'collection on every arrow key. The handler below wants the store, and reads the
'held one.
'@param sh Worksheet. The worksheet where the selection changed.
'@param target Range. The newly selected cell(s).
'@export
Public Sub OnSelectionChange(ByVal sh As Worksheet, ByVal target As Range)
    If (sh Is Nothing) Or (target Is Nothing) Then Exit Sub

    'The value under the selection is the previous value of the edit that may
    'follow. The multiple choice toggle reads it back.
    RememberCellValue target

    Dim shHn As HiddenNames
    Set shHn = SheetStore(sh)
    If shHn Is Nothing Then Exit Sub

    HandleSelectionHList sh, target, shHn
End Sub


'@sub-title Handle double-click events on analysis sheets.
'@details
'Answers the geo scope whose picker the double-click asked for, and
'GEOSCOPENONE when the click asked for no picker at all. The caller is what
'opens the picker: the form lives in the workbook and a class cannot reach it.
'@param sh Worksheet. The worksheet where the double-click occurred.
'@param target Range. The double-clicked cell.
'@return Long. GeoScopeAdmin, GeoScopeHF, or GEOSCOPENONE.
'@export
Public Function OnDoubleClick(ByVal sh As Worksheet, ByVal target As Range) As Long
    Dim rngName As String
    Dim shType As String

    OnDoubleClick = GEOSCOPENONE

    If (sh Is Nothing) Or (target Is Nothing) Then Exit Function

    On Error Resume Next
    rngName = target.Name.Name
    On Error GoTo 0

    shType = SheetTag(sh)
    If shType <> "SPT-Analysis" Then Exit Function

    If (InStr(1, rngName, "INPUTSPTGEO_") > 0) Then
        OnDoubleClick = GeoScopeAdmin
    ElseIf (InStr(1, rngName, "INPUTSPTHF_") > 0) Then
        OnDoubleClick = GeoScopeHF
    End If
End Function


'@sub-title Update filter tables used by analysis sheets.
'@details
'Every filtered table is rewritten from its source table. A row is kept when
'its source row is visible and holds at least one value. The old shape copied
'the whole data body across, then walked every row from the bottom reading
'EntireRow.Hidden and CountA and deleted the rejects -- two COM crossings per
'row and an Application.Union that degrades as it grows, on tables that
'ClickAddRows pads with 199 blank rows at a time. The recalculation at the
'foot is one pass: the sheet pass covers the UsedRange and column passes that
'used to follow it.
'@param calculate Optional Boolean. When True (default), recalculates the active sheet.
'@export
Public Sub UpdateFilterTables(Optional ByVal calculate As Boolean = True)
    Dim sh As Worksheet
    Dim filtsh As Worksheet
    Dim lo As ListObject
    Dim shHn As HiddenNames

    On Error GoTo ErrUpdate

    For Each sh In this.hostBook.Worksheets
        If SheetTag(sh) = "HList" Then

            Set lo = sh.ListObjects(1)

            If Not lo.DataBodyRange Is Nothing Then
                Set shHn = SheetStore(sh)
                Set filtsh = this.hostBook.Worksheets(shHn.ValueAsString("filtered_sheet"))

                RewriteFilteredTable lo, filtsh
            End If
        End If
    Next

    If calculate Then ActiveSheet.Calculate

    Exit Sub

ErrUpdate:
    Fail "MSG_ErrUpdate", Err.Description
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


'@sub-title The one translation helper of the workbook.
'@details
'The button, ribbon and geo modules each used to build their own. LLTranslation
'validates all five tables on Create and caches every scope it is asked for
'until Refresh, so one instance answers all four callers for the price of the
'first build, and Refresh then has one place to be called from. A module reaches
'this through LinelistEventsManager.EventLinelistService.
'@return LLTranslation. The held helper, or Nothing when the workbook carries no
'usable translation sheet.
'@export
Public Function Translation() As LLTranslation
    EnsureTranslation
    Set Translation = this.lltrads
End Function


'@sub-title The hidden names of the workbook, built once and held.
'@details
'The button module used to build its own on every click, and a workbook store
'costs a walk of the whole Names collection.
'@return HiddenNames. The held store, or Nothing when the build failed.
'@export
Public Function WorkbookNames() As HiddenNames
    EnsureWorkbookNames
    Set WorkbookNames = this.wkbNames
End Function


'@sub-title The hidden names of one worksheet, built once and held.
'@details
'One store serves the event handlers and the buttons together, and the class
'drops it when that sheet raises a change. The button module read a sheet tag,
'a table name and a blank row count through three separate stores.
'@param sh Worksheet. The worksheet whose names are wanted.
'@return HiddenNames. The held store, or Nothing when the build failed.
'@export
Public Function SheetNames(ByVal sh As Worksheet) As HiddenNames
    Set SheetNames = SheetStore(sh)
End Function


'@sub-title The password manager of the workbook, built once and held.
'@return Passwords. The held manager, or Nothing when the build failed.
'@export
Public Function PasswordManager() As Passwords
    EnsurePasswords
    Set PasswordManager = this.pass
End Function


'@sub-title Clear cached resources.
'@export
Public Sub ResetCaches()
    ResetInternalCaches
End Sub


'@section Messages
'===============================================================================
'One warning and one failure box for the whole linelist event surface. The
'button, ribbon and geo modules reach these through
'LinelistEventsManager.EventLinelistService. Four files used to write their own
'MsgBox: some passed a title and some left it off, and the same class of message
'came up with vbCritical in one file and vbExclamation in another.

'@sub-title Tell the user why a button refused to act.
'@param msgCode String. Code of the message in the message translation table.
'@export
Public Sub Warn(ByVal msgCode As String)
    ShowMessage msgCode, vbOKOnly + vbExclamation, vbNullString, vbNullString
End Sub


'@sub-title Tell the user an operation failed.
'@param msgCode String. Code of the message in the message translation table.
'@param detail String. Optional. Follows the message, for an Err.Description.
'@param fallback String. Optional. Shown when the workbook carries no usable
'translation sheet. Leave it off to stay quiet in that state.
'@export
Public Sub Fail(ByVal msgCode As String, _
                Optional ByVal detail As String = vbNullString, _
                Optional ByVal fallback As String = vbNullString)
    ShowMessage msgCode, vbOKOnly + vbCritical, detail, fallback
End Sub


'@sub-title Show one box, translated when a translator is there.
'@details
'A workbook with no usable translation sheet has no text to show, and the
'callers that still have something to say pass a fallback. Every other caller
'stays quiet, which is what the two boxes this replaced did.
Private Sub ShowMessage(ByVal msgCode As String, _
                        ByVal boxStyle As Long, _
                        ByVal detail As String, _
                        ByVal fallback As String)
    Dim message As String
    Dim title As String

    EnsureTranslation

    If this.tradmess Is Nothing Then
        If LenB(fallback) = 0 Then Exit Sub
        message = fallback
        title = "Error"
    Else
        message = this.tradmess.TranslatedValue(msgCode)
        title = this.tradmess.TranslatedValue(MSGTITLE)
    End If

    If LenB(detail) > 0 Then message = message & ": " & detail

    MsgBox message, boxStyle, title
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


'@section Private -- HList change handler
'===============================================================================

'@sub-title Handle value changes on HList sheets.
'@details
'Dispatches on the position of the edited cell, cheapest and most specific test
'first, and every branch exits. An edit inside the table recalculates its own
'row and goes to the geo cascade, the list-auto flag or the multiple choice
'toggle; an edit on the label row updates the editable label; an edit on the
'go-to cell navigates to the chosen section; an edit on the header row is
'refused and the cell restored.
'
'THE GO-TO TEST READS THE FIRST CELL AND THE SHEET STORE
'-------------------------------------------------------------------------------
'The builder puts the go-to dropdown in the first cell of the sheet and stores
'its caption in the sheet store under <table>_go_to_section. Resolving that id
'as a range name raised 1004 on every sheet the builder makes, because the id
'names a stored string; the raise landed on the error label, so the header
'protection and the multiple choice toggle below the old block never ran.
'@param sh Worksheet. The HList worksheet.
'@param target Range. The edited cell(s).
'@param shHn HiddenNames. The worksheet-level hidden names (shared from caller).
Private Sub HandleHListChange(ByRef sh As Worksheet, ByRef target As Range, _
                              ByRef shHn As HiddenNames)
    Dim adminTable As BetterArray
    Dim adminNames As BetterArray
    Dim hRng As Range
    Dim calcRng As Range
    Dim cellRng As Range
    Dim varControl As String
    Dim varLabel As String
    Dim varName As String
    Dim varSubLabel As String
    Dim tableName As String
    Dim varEditable As String
    Dim goToCaption As String
    Dim sectionName As String
    Dim targetColumn As Long
    Dim startLine As Long
    Dim nbOffset As Long
    Dim choiSep As String

    On Error GoTo ErrHand

    Set hRng = sh.ListObjects(1).HeaderRowRange
    startLine = hRng.Row + 1
    targetColumn = target.Column

    'The edit sits in the table: one row is recalculated, and the control
    'string decides the branch. The tests below this block all sit above the
    'table, so the block exits whatever the control says.
    If (target.Row >= startLine) Then

        nbOffset = target.Row - hRng.Row
        Set calcRng = hRng.Offset(nbOffset)
        calcRng.Calculate

        varName = sh.Cells(hRng.Row, targetColumn).Value
        varControl = shHn.ValueAsString(varName & " -- control")

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

            'The two arrays serve this branch alone
            Set adminNames = New BetterArray
            adminNames.LowerBound = 1

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

        'Update the list auto flag
        If shHn.ValueAsString(varName & " -- listauto") <> vbNullString Then
            EnsureWorkbookNames
            If Not this.wkbNames Is Nothing Then
                this.wkbNames.SetValue "RNG_UpdateListAuto", "yes"
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
        End If

        Exit Sub
    End If

    'Update variables with editable column (label row = header - 1)
    If (target.Row = hRng.Row - 1) Then
        varName = sh.Cells(hRng.Row, targetColumn).Value
        varEditable = shHn.ValueAsString(varName & " -- editable")
        If (varEditable <> "yes") Then Exit Sub

        EnsureVariables
        If this.vars Is Nothing Then Exit Sub

        varSubLabel = this.vars.Value(varName:=varName, colName:="sub label")
        varLabel = Replace(target.Value, varSubLabel, vbNullString)
        varLabel = Replace(varLabel, Chr(10), vbNullString)

        this.vars.SetValue varName:=varName, colName:="main label", _
                           newValue:=varLabel
        Exit Sub
    End If

    'GoTo section: the dropdown is the first cell of the sheet, and the caption
    'held in the sheet store is the prefix of every entry it offers
    If Not (Intersect(target, sh.Cells(1, 1)) Is Nothing) Then
        tableName = shHn.ValueAsString("table_name")
        goToCaption = shHn.ValueAsString(tableName & "_" & GOTOSECCODE)

        If LenB(goToCaption) > 0 Then
            sectionName = Replace(target.Value, goToCaption & ": ", vbNullString)

            Set hRng = hRng.Offset(-3)
            Set cellRng = hRng.Find(What:=sectionName, LookAt:=xlWhole, MatchCase:=True)
            If Not cellRng Is Nothing Then cellRng.Activate
            Exit Sub
        End If
    End If

    'Avoid modifying headers of the table (with variable names)
    If (target.Row = hRng.Row) Then
        target.Value = target.Offset(-1).Name.Name
        Warn "MSG_NotModify"
        Exit Sub
    End If

    Exit Sub
ErrHand:
End Sub


'@section Private -- VList change handler
'===============================================================================

'@sub-title Handle value changes on VList sheets.
'@details
'Dispatches on the column of the edited cell and every branch exits. An edit in
'the label column updates the editable label; an edit in the value column
'recalculates the value block once and runs the multiple choice toggle when the
'control asks for it; an edit on the go-to cell navigates to the chosen label.
'The value block used to be recalculated on every edit of the sheet, above the
'dispatch, so a label edit and a navigation both paid for it.
'
'The go-to test reads the first cell of the sheet and the caption the sheet
'store carries, the same way the HList handler does.
'@param sh Worksheet. The VList worksheet.
'@param target Range. The edited cell(s).
'@param shHn HiddenNames. The worksheet-level hidden names (shared from caller).
Private Sub HandleVListChange(ByRef sh As Worksheet, ByRef target As Range, _
                              ByRef shHn As HiddenNames)
    Dim rng As Range
    Dim rngLook As Range
    Dim varLabel As String
    Dim tableName As String
    Dim varSubLabel As String
    Dim varEditable As String
    Dim varName As String
    Dim varControl As String
    Dim choiSep As String
    Dim goToCaption As String

    On Error GoTo ErrHandler

    tableName = shHn.ValueAsString("table_name")
    Set rng = sh.Range(tableName & "_START")

    'Update variables with editable column
    If (target.Column = rng.Offset(, -1).Column) Then
        varName = (target.Offset(, 1).Name.Name)
        varEditable = shHn.ValueAsString(varName & " -- editable")
        If (varEditable <> "yes") Then Exit Sub

        EnsureVariables
        If this.vars Is Nothing Then Exit Sub

        varSubLabel = this.vars.Value(varName:=varName, colName:="sub label")
        varLabel = Replace(target.Value, varSubLabel, vbNullString)
        varLabel = Replace(varLabel, Chr(10), vbNullString)
        this.vars.SetValue varName:=varName, colName:="main label", _
                           newValue:=varLabel
        Exit Sub
    End If

    'An edit in the value column: the value block is recalculated once, and the
    'multiple choice toggle runs when the control asks for it
    If target.Column = rng.Column Then
        sh.Range(tableName & "_" & "PLAGEVALUES").Calculate

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
        End If
        Exit Sub
    End If

    'Jump to GoTo section
    If Not (Intersect(target, sh.Cells(1, 1)) Is Nothing) Then
        goToCaption = shHn.ValueAsString(tableName & "_" & GOTOSECCODE)

        If LenB(goToCaption) > 0 Then
            varLabel = Replace(target.Value, goToCaption & ": ", vbNullString)
            Set rngLook = sh.Cells.Find(What:=varLabel, LookAt:=xlWhole, MatchCase:=True)
            If Not rngLook Is Nothing Then rngLook.Activate
        End If
    End If

    Exit Sub
ErrHandler:
End Sub


'@section Private -- Analysis change handler
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

    'Range name if it exists
    On Error Resume Next
    rngName = target.Name.Name
    On Error GoTo 0

    On Error GoTo ErrHandler

    Select Case analysisType

    Case "Uni-Bi-Analysis"
        Set rng = sh.Range("ua_go_to_section")

    Case "TS-Analysis", "SPT-Analysis"
        'One recalculation covers the sheet. These sheets carry array formulas
        'over the whole linelist, and two further passes here were what a user
        'felt as the workbook freezing while typing in a filter.
        sh.Calculate
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

    EnsureWorkbookNames
    If Not this.wkbNames Is Nothing Then
        goToSection = this.wkbNames.ValueAsString("RNG_GoToSection")
        goToHeader = this.wkbNames.ValueAsString("RNG_GoToHeader")
        goToGraph = this.wkbNames.ValueAsString("RNG_GoToGraph")
    End If

    sLabel = Replace(target.Value, goToSection & ": ", vbNullString)
    sLabel = Replace(sLabel, goToHeader & ": ", vbNullString)
    sLabel = Replace(sLabel, goToGraph & ": ", vbNullString)

    'The search runs on the sheet that raised the change. ActiveSheet is the
    'same sheet on a user edit and a different one on a programmatic edit.
    Set rngLook = sh.Cells.Find(What:=sLabel, LookIn:=xlValues, LookAt:=xlWhole, _
                                MatchCase:=True, SearchFormat:=False)

    If Not rngLook Is Nothing Then rngLook.Activate

    Exit Sub
ErrHandler:
End Sub


'@section Private -- Spatial table update handlers
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
    EnsureWorkbookNames

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


'@section Private -- Selection change handler
'===============================================================================

'@sub-title Handle selection changes on HList sheets.
'@details
'Updates geo dropdown validation lists when the user navigates to a geo2, geo3,
'or geo4 cell.
'@param sh Worksheet. The HList worksheet.
'@param target Range. The newly selected cell(s).
'@param shHn HiddenNames. The worksheet-level hidden names (shared from caller).
Private Sub HandleSelectionHList(ByRef sh As Worksheet, ByRef target As Range, _
                                ByRef shHn As HiddenNames)
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


'@section Private -- List auto helpers
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
    Dim shHn As HiddenNames
    Dim colIdx As Long

    Set shHn = SheetStore(sh)
    If shHn Is Nothing Then Exit Sub

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


'@section Private -- Multiple choice helper
'===============================================================================

'@sub-title Handle multiple choice dropdown value toggling.
'@details
'The toggle needs the value the cell held before the edit. That comes from the
'held copy RememberCellValue keeps, filled when the user lands on the cell and
'refreshed after every edit. The undo stack is the fallback for a cell the
'session has no reading for -- it costs a round trip per call, which on a
'multiple choice column is a round trip per keystroke.
'@param target Range. The edited cell.
'@param choiSep String. Separator between choices (e.g. ", ").
Private Sub UpdateMultipleChoice(ByRef target As Range, ByRef choiSep As String)
    Dim actualValue As String
    Dim prevValue As String
    Dim prevTab As BetterArray
    Dim actTab As BetterArray

    If IsEmpty(target) Then Exit Sub

    actualValue = target.Value

    If this.lastEditKnown And (this.lastEditKey = CellKey(target)) Then
        'The cell still holds what the user just entered on this path, so
        'keeping the entry needs no write.
        On Error GoTo ErrHandler
        prevValue = this.lastEditValue
        If prevValue = vbNullString Then Exit Sub
    Else
        On Error Resume Next
        Application.Undo
        On Error GoTo ErrHandler

        If IsEmpty(target) Then GoTo KeepValue
        prevValue = target.Value
    End If

    Set prevTab = New BetterArray
    Set actTab = New BetterArray

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


'@section Private -- Previous value capture
'===============================================================================

'@sub-title Hold the value of one cell for the next edit to it.
'@details
'OnSelectionChange stores the value found under the selection, and OnSheetChange
'stores what the cell holds once its handlers have run. A selection of several
'cells drops the held reading, because one value can stand for one cell only.
'@param target Range. The selected or edited cells.
Private Sub RememberCellValue(ByRef target As Range)
    this.lastEditKnown = False

    If target Is Nothing Then Exit Sub
    If target.Cells.Count > 1 Then Exit Sub

    On Error Resume Next
    this.lastEditKey = CellKey(target)
    this.lastEditValue = CStr(target.Value)
    this.lastEditKnown = (Err.Number = 0)
    On Error GoTo 0
End Sub

'@sub-title The identity of one cell, as sheet name and address.
'@param target Range. The cell.
'@return String. Sheet name and address joined with an exclamation mark.
Private Function CellKey(ByRef target As Range) As String
    CellKey = target.Worksheet.Name & "!" & target.Address
End Function


'@section Private -- Unique values helper
'===============================================================================

'@sub-title Return unique values from a BetterArray.
'@details
'The array built here goes to the caller as it stands. Each call used to answer
'a Clone, which copied every unique value once per list-auto column per sheet
'deactivate.
'@param inputTable BetterArray. Source array.
'@return BetterArray. Array of unique values.
Private Function GetUniqueBA(ByRef inputTable As BetterArray) As BetterArray
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

    Set GetUniqueBA = outTable
End Function


'@section Private -- Filtered table copy
'===============================================================================

'@sub-title Rewrite one filtered table from the visible rows of its source.
'@details
'A row is kept when it is visible and holds at least one value. Visibility is
'read off the source rows in one SpecialCells call. The read spans the whole
'row on purpose: a column hidden by show/hide keeps its values in the filtered
'table, and a row leaves when a filter or the user hides it. The values travel
'through memory -- one read of the whole block, the blank test on the array,
'and one write of the kept rows. The two tables sit at the same address on
'their two sheets, which is the layout the builder writes.
'@param sourceTable ListObject. The data entry table the rows come from.
'@param filtsh Worksheet. The sheet holding the filtered table.
Private Sub RewriteFilteredTable(ByVal sourceTable As ListObject, _
                                 ByVal filtsh As Worksheet)
    Dim bodyRng As Range
    Dim visibleCells As Range
    Dim visibleArea As Range
    Dim filtLo As ListObject
    Dim sourceValues As Variant
    Dim singleValue As Variant
    Dim keptValues() As Variant
    Dim keepRow() As Boolean
    Dim rowCount As Long
    Dim columnCount As Long
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim keptCount As Long
    Dim firstRow As Long
    Dim keptIdx As Long

    Set bodyRng = sourceTable.DataBodyRange
    rowCount = bodyRng.Rows.Count
    columnCount = bodyRng.Columns.Count

    sourceValues = bodyRng.Value
    If (rowCount = 1) And (columnCount = 1) Then
        'A single cell answers a scalar, and the walk below wants a block.
        singleValue = sourceValues
        ReDim sourceValues(1 To 1, 1 To 1)
        sourceValues(1, 1) = singleValue
    End If

    'SpecialCells raises 1004 when every row is hidden, and an empty answer is
    'what that state means. The suppression covers that one call.
    ReDim keepRow(1 To rowCount)
    On Error Resume Next
    Set visibleCells = bodyRng.EntireRow.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleCells Is Nothing Then
        For Each visibleArea In visibleCells.Areas
            firstRow = visibleArea.Row - bodyRng.Row + 1
            For rowIdx = firstRow To firstRow + visibleArea.Rows.Count - 1
                If (rowIdx >= 1) And (rowIdx <= rowCount) Then keepRow(rowIdx) = True
            Next rowIdx
        Next visibleArea
    End If

    For rowIdx = 1 To rowCount
        If keepRow(rowIdx) Then
            If RowHoldsAValue(sourceValues, rowIdx, columnCount) Then
                keptCount = keptCount + 1
            Else
                keepRow(rowIdx) = False
            End If
        End If
    Next rowIdx

    If keptCount > 0 Then
        ReDim keptValues(1 To keptCount, 1 To columnCount)
        For rowIdx = 1 To rowCount
            If keepRow(rowIdx) Then
                keptIdx = keptIdx + 1
                For colIdx = 1 To columnCount
                    keptValues(keptIdx, colIdx) = sourceValues(rowIdx, colIdx)
                Next colIdx
            End If
        Next rowIdx
    End If

    Set filtLo = filtsh.ListObjects(1)

    'A table already down to its header row has nothing to delete.
    On Error Resume Next
    filtLo.DataBodyRange.Delete
    On Error GoTo 0

    If keptCount > 0 Then
        filtLo.Resize filtsh.Range(sourceTable.HeaderRowRange.Resize(keptCount + 1).Address)
        filtLo.DataBodyRange.Value = keptValues
    End If
End Sub


'@sub-title Say whether one row of a value block holds anything.
'@details
'Mirrors what CountA answered on the old copied row: an empty cell and an
'empty string count as nothing, and every other value -- a zero, a date, an
'error -- counts as data.
'@param blockValues Variant. The value block, 1-based on both sides.
'@param rowIdx Long. The row to read.
'@param columnCount Long. The width of the block.
'@return Boolean. True when at least one cell of the row holds a value.
Private Function RowHoldsAValue(ByRef blockValues As Variant, _
                                ByVal rowIdx As Long, _
                                ByVal columnCount As Long) As Boolean
    Dim colIdx As Long
    Dim cellValue As Variant

    For colIdx = 1 To columnCount
        cellValue = blockValues(rowIdx, colIdx)
        If Not IsEmpty(cellValue) Then
            If VarType(cellValue) = vbString Then
                If LenB(cellValue) > 0 Then
                    RowHoldsAValue = True
                    Exit Function
                End If
            Else
                RowHoldsAValue = True
                Exit Function
            End If
        End If
    Next colIdx
End Function


'@section Private -- Sheet metadata helpers
'===============================================================================

'@sub-title Get the sheet type tag from worksheet-level HiddenNames.
'@details
'The quick read is what this wants. It used to build a whole HiddenNames store
'for one value, and UpdateFilterTables and UpdateAllListAuto both call it once
'per worksheet of the workbook, so every sheet in the loop paid a full walk of
'its own Names collection.
'@param sh Worksheet. The worksheet to query.
'@return String. The sheet type tag (HList, VList, Uni-Bi-Analysis, etc.).
Private Function SheetTag(ByRef sh As Worksheet) As String
    SheetTag = HiddenNames.QuickValue(sh, "sheet_type")
End Function


'@section Private -- Lazy initialisation
'===============================================================================

'@sub-title Lazily initialise the workbook-level hidden names.
'@details
'This used to sit at the foot of EnsureTranslation, below two early exits. A
'workbook with no LinelistTranslation sheet took the exit at that sheet lookup
'and never reached the line, so this.wkbNames stayed Nothing for the whole
'session: RNG_GoToSection was never read and go-to-section did nothing on every
'sheet, with no message. The store has nothing to do with the translation
'tables, so it is built on its own.
Private Sub EnsureWorkbookNames()
    If Not (this.wkbNames Is Nothing) Then Exit Sub
    If this.triedWkbNames Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedWkbNames = True

    On Error Resume Next
    Set this.wkbNames = HiddenNames.Create(this.hostBook)
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the translation helper and the message translator.
'@details
'A translation sheet missing a table, or a workbook whose language code is
'empty, leaves the two fields Nothing and every event that needs them does
'nothing. The suppression is kept for that reason: an event handler has nothing
'above it to catch a raise. A workbook in that state costs the events their
'labels and leaves everything else working.
Private Sub EnsureTranslation()
    If Not (this.lltrads Is Nothing) Then Exit Sub
    If this.triedTranslation Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedTranslation = True

    Dim sh As Worksheet
    Set sh = SheetOrNothing(LLSHEET)
    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set this.lltrads = LLTranslation.Create(sh)
    If Not this.lltrads Is Nothing Then
        Set this.tradmess = this.lltrads.TransObject()
    End If
    On Error GoTo 0
End Sub

'@sub-title Lazily initialise the dropdown manager.
Private Sub EnsureDropdown()
    If Not (this.drop Is Nothing) Then Exit Sub
    If this.triedDropdown Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedDropdown = True

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
    If this.triedGeo Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedGeo = True

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
    If this.triedSpatial Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedSpatial = True

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
    If this.triedPasswords Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedPasswords = True

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
    If this.triedDictionary Then Exit Sub
    If this.hostBook Is Nothing Then Exit Sub

    this.triedDictionary = True

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
    If this.triedVariables Then Exit Sub

    EnsureDictionary
    If this.dict Is Nothing Then Exit Sub

    this.triedVariables = True

    On Error Resume Next
    Set this.vars = LLVariables.Create(this.dict)
    On Error GoTo 0
End Sub

'@sub-title The hidden names of one worksheet, built once and held.
'@details
'Keyed on the sheet name in a Collection. Scripting.Dictionary is missing on Mac
'Excel and LLFormat holds its three caches the same shape. The store of a sheet
'costs a walk of that sheet's Names collection, and the selection handler asks
'for one on every arrow key.
'@param sh Worksheet. The worksheet whose names are wanted.
'@return HiddenNames. The held store, or Nothing when the build failed.
Private Function SheetStore(ByRef sh As Worksheet) As HiddenNames
    Dim store As HiddenNames

    If sh Is Nothing Then Exit Function
    If this.sheetStores Is Nothing Then Set this.sheetStores = New Collection

    'A Collection has no "does this key exist" test, so the miss is read off Err.
    'Err is cleared first, because On Error Resume Next and On Error GoTo 0 both
    'leave it alone and a stale number from earlier code would read as a hit.
    On Error Resume Next
        Err.Clear
        Set store = this.sheetStores.Item(sh.Name)
        Err.Clear
    On Error GoTo 0

    If store Is Nothing Then
        On Error Resume Next
        Set store = HiddenNames.Create(sh)
        On Error GoTo 0
        If store Is Nothing Then Exit Function
        this.sheetStores.Add store, sh.Name
    End If

    Set SheetStore = store
End Function

'@sub-title Forget the held store of one worksheet.
'@param sheetName String. Name of the worksheet whose store is dropped.
Private Sub DropSheetStore(ByRef sheetName As String)
    If this.sheetStores Is Nothing Then Exit Sub

    On Error Resume Next
    this.sheetStores.Remove sheetName
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


'@section Private -- Cache management
'===============================================================================

'@sub-title Clear all lazily-cached domain managers.
'@details
'The tried flags go with the fields. A flag left standing turns this into a
'permanent refusal to build the manager again.
Private Sub ResetInternalCaches()
    Set this.tradmess = Nothing
    Set this.lltrads = Nothing
    Set this.wkbNames = Nothing
    Set this.sheetStores = Nothing
    Set this.drop = Nothing
    Set this.geo = Nothing
    Set this.sp = Nothing
    Set this.dict = Nothing
    Set this.vars = Nothing
    Set this.pass = Nothing

    this.triedTranslation = False
    this.triedWkbNames = False
    this.triedDropdown = False
    this.triedGeo = False
    this.triedSpatial = False
    this.triedDictionary = False
    this.triedVariables = False
    this.triedPasswords = False

    this.lastEditKey = vbNullString
    this.lastEditValue = vbNullString
    this.lastEditKnown = False
End Sub


'@section Private -- Error handling
'===============================================================================

'@sub-title Raise a ProjectError-based exception.
'@param errNumber Long. The error code to raise.
'@param message String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errNumber As Long, ByRef message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub
