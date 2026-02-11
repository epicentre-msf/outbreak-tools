VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist data import operations"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("DataIO")
'@ModuleDescription("Linelist data import operations")

Option Explicit
Implements ILLImporter

'@class LLImporter
'@description
'Imports linelist data from external workbooks into the source linelist.
'Supports importing HList/VList data, custom dropdowns, show/hide state,
'editable labels, single metadata values, and geobase data. Tracks which
'sheets and variables were imported, and builds a report of items not
'imported or not touched. Consumers interact through the ILLImporter
'interface.
'@depends ILLdictionary, IPasswords, ICustomTable, IDataSheet
'@depends IShowHideManager, ILLGeo, ILLChoices, ILLVariables, IDropdownLists, IHiddenNames


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "LLImporter"

' Source workbook sheet names
Private Const DICTIONARYSHEET As String = "Dictionary"
Private Const CHOICESSHEET As String = "Choices"
Private Const DICTTRANSLATIONSHEET As String = "Translations"
Private Const METADATASHEET As String = "Metadata"
Private Const PASSWORDSHEET As String = "__pass"
Private Const SHOWHIDESHEET As String = "show_hide__"
Private Const GEOSHEET As String = "Geo"
Private Const RNGEPIWEEKSTART As String = "RNG_EpiWeekStart"


'@section State
'===============================================================================

Private Type TLLImporter
    sourceWkb As Workbook
    hasSameLanguage As Boolean
    needReport As Boolean
    ' Report data
    sheetsNotImported As BetterArray
    sheetsNotTouched As BetterArray
    varsNotImported As BetterArray
    varsNotTouched As BetterArray
    ' Tracking
    importedSheets As BetterArray
    importedVars As BetterArray
End Type

Private this As TLLImporter


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create an importer from the source linelist workbook
'@details
'Factory method that validates the workbook is not Nothing, then constructs
'and returns a new LLImporter instance through the ILLImporter interface.
'Called on the predeclared instance.
'@param sourceWkb Workbook. The linelist workbook that will receive imported data.
'@return ILLImporter. A fully initialised instance ready for import operations.
'@throws vbObjectError When the workbook is Nothing.
'@export
Public Function Create(ByVal sourceWkb As Workbook) As ILLImporter

    If sourceWkb Is Nothing Then
        ThrowError "LLImporter requires a source workbook"
    End If

    With New LLImporter
        Set .SourceWorkbook = sourceWkb
        Set Create = .Self
    End With
End Function


'@section Public Accessors
'===============================================================================

'@label:self
'@prop-title Current object instance
'@return ILLImporter. The current instance cast to the interface.
Public Property Get Self() As ILLImporter
    Set Self = Me
End Property

'@label:source-workbook
'@prop-title Source linelist workbook
'@return Workbook. The linelist workbook bound to this importer.
Public Property Get SourceWorkbook() As Workbook
    Set SourceWorkbook = this.sourceWkb
End Property

'@label:source-workbook-set
'@prop-title Assign the source linelist workbook
'@param wb Workbook. The workbook to store.
Public Property Set SourceWorkbook(ByVal wb As Workbook)
    Set this.sourceWkb = wb
End Property


'@section Data Checks
'===============================================================================

'@label:has-data
'@prop-title Whether the workbook contains user-entered data
'@details
'Scans all HList sheets for non-empty rows beyond the blank count
'threshold stored in cell(1,6). Returns True as soon as a single row
'has user data. Used to warn the user before clearing or overwriting.
'@return Boolean. True when at least one HList row has user data.
'@export
Public Property Get HasData() As Boolean

    Dim sh As Worksheet
    Dim lo As ListObject
    Dim nbBlanks As Long
    Dim counter As Long

    For Each sh In this.sourceWkb.Worksheets
        If CStr(sh.Cells(1, 3).Value) = "HList" Then
            nbBlanks = CLng(sh.Cells(1, 6).Value)
            Set lo = sh.ListObjects(1)

            If Not lo.DataBodyRange Is Nothing Then
                For counter = 1 To lo.DataBodyRange.Rows.Count
                    If Application.WorksheetFunction.CountA( _
                        lo.ListRows(counter).Range) > nbBlanks Then
                        HasData = True
                        Exit Property
                    End If
                Next counter
            End If
        End If
    Next sh
End Property

'@label:has-same-language
'@sub-title Check whether the import workbook language matches
'@details
'Reads the language tag from the import workbook Metadata sheet and
'compares it with the source workbook dictionary language via
'LLTranslations. Stores the result internally so that ImportShowHide,
'ImportEditableLabels, and ImportSingleValues silently skip when
'languages differ. Returns False when the Metadata sheet is missing.
'@param impwb Workbook. The opened import workbook.
'@return Boolean. True when languages match.
'@export
Public Function HasSameLanguage(ByVal impwb As Workbook) As Boolean

    Dim metaTab As IDataSheet
    Dim varRng As Range
    Dim cellRng As Range
    Dim impLang As String
    Dim currLang As String
    Dim wkbNames As IHiddenNames

    this.hasSameLanguage = False

    ' Metadata sheet must exist
    If Not SheetExists(impwb, METADATASHEET) Then
        HasSameLanguage = False
        Exit Function
    End If

    ' Read import language from metadata
    Set metaTab = DataSheet.Create(impwb.Worksheets(METADATASHEET), 1, 1)
    Set varRng = metaTab.DataRange("variable")
    impLang = vbNullString

    On Error Resume Next
    Set cellRng = varRng.Find(What:="language", lookAt:=xlWhole, MatchCase:=False)
    impLang = CStr(cellRng.Offset(, 1).Value)
    On Error GoTo 0

    If LenB(impLang) = 0 Then
        HasSameLanguage = False
        Exit Function
    End If

    ' Compare with current language
    Set wkbNames = HiddenNames.Create(this.sourceWkb)
    currLang = wkbNames.ValueAsString("RNG_DictionaryLanguage")

    this.hasSameLanguage = (impLang = currLang)
    HasSameLanguage = this.hasSameLanguage
End Function


'@section Data Operations
'===============================================================================

'@label:clear-data
'@sub-title Clear all user-entered data from HList and VList sheets
'@details
'Iterates over all worksheets in the source workbook. For HList sheets,
'clears non-formula column values from the ListObject data body. For
'VList sheets, clears non-formula named range values. Unprotects and
're-protects each sheet via IPasswords.
'@export
Public Sub ClearData()

    Dim sh As Worksheet
    Dim pass As IPasswords
    Dim lo As ListObject
    Dim valueRng As Range
    Dim labelRng As Range
    Dim counter As Long
    Dim actsh As Worksheet

    Set pass = PasswordsObject()
    Set actsh = ActiveSheet

    For Each sh In this.sourceWkb.Worksheets

        Select Case CStr(sh.Cells(1, 3).Value)

        Case "HList"
            pass.UnProtect sh.Name
            Set lo = sh.ListObjects(1)

            If Not lo.DataBodyRange Is Nothing Then
                For counter = 1 To lo.DataBodyRange.Columns.Count
                    Set valueRng = lo.ListColumns(counter).DataBodyRange
                    If Not valueRng.Cells(1, 1).HasFormula Then
                        valueRng.ClearContents
                    End If
                Next counter
                lo.Range.Calculate
            End If

            pass.Protect sh.Name

        Case "VList"
            pass.UnProtect sh.Name
            Set labelRng = sh.Cells(4, 4)

            Do While Not IsEmpty(labelRng)
                Set valueRng = labelRng.Offset(, 1)
                If Not valueRng.HasFormula Then valueRng.ClearContents
                Set labelRng = labelRng.Offset(1)
            Loop

            sh.UsedRange.Calculate
            pass.Protect sh.Name

        End Select
    Next sh

    actsh.Activate
End Sub

'@label:import-data
'@sub-title Import data from HList and VList worksheets
'@details
'Iterates over the import workbook worksheets, matching them to source
'workbook sheets by name. For HList sheets, delegates to
'ImportHListSheet. For VList sheets, delegates to ImportVListSheet.
'Tracks imported sheets and variables, and builds "not imported" report
'entries for sheets and variables not found in the current workbook.
'@param impwb Workbook. The opened import workbook.
'@param pasteAtBottom Boolean. When True, appends data below existing rows.
'@export
Public Sub ImportData(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)

    Dim dict As ILLdictionary
    Dim sheetsList As BetterArray
    Dim pass As IPasswords
    Dim impsh As Worksheet
    Dim currsh As Worksheet

    Set dict = SourceDictionary()
    Set sheetsList = dict.UniqueValues("sheet name")
    Set pass = PasswordsObject()

    ' Initialize tracking
    InitializeTracking

    For Each impsh In impwb.Worksheets

        ' Track imported sheets (non-reserved)
        If Not IsReservedSheet(impsh.Name) Then
            this.importedSheets.Push impsh.Name
        End If

        If sheetsList.Includes(impsh.Name) Then
            Set currsh = this.sourceWkb.Worksheets(impsh.Name)

            Select Case CStr(currsh.Cells(1, 3).Value)

            Case "VList"
                pass.UnProtect currsh.Name
                ImportVListSheet currsh, impsh
                pass.Protect currsh.Name

            Case "HList"
                pass.UnProtect currsh.Name
                ImportHListSheet currsh, impsh, pasteAtBottom
                pass.Protect currsh.Name

            End Select

            ' Recalculate after import
            On Error Resume Next
            currsh.UsedRange.Calculate
            On Error GoTo 0

        ElseIf Not IsReservedSheet(impsh.Name) Then
            ' Sheet from import not found in current workbook
            this.sheetsNotImported.Push impsh.Name
            this.needReport = True
        End If

    Next impsh
End Sub

'@label:import-custom-dropdown
'@sub-title Import custom dropdown choices from the import workbook
'@details
'Reads custom dropdown ListObjects from the source workbook custom
'dropdown sheet, matches them to choices in the import workbook Choices
'sheet, and updates each dropdown with the imported values. Silently
'exits when the import workbook has no Choices sheet or the custom
'dropdown sheet does not exist.
'@param impwb Workbook. The opened import workbook.
'@param pasteAtBottom Boolean. When True, appends choices below existing entries.
'@export
Public Sub ImportCustomDropdown(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)

    Dim wkbNames As IHiddenNames
    Dim choi As ILLChoices
    Dim lo As ListObject
    Dim drop As IDropdownLists
    Dim dropsh As Worksheet
    Dim listName As String
    Dim choiceName As String
    Dim updateData As BetterArray

    If Not SheetExists(impwb, CHOICESSHEET) Then Exit Sub

    Set wkbNames = HiddenNames.Create(this.sourceWkb)

    On Error Resume Next
    Set dropsh = this.sourceWkb.Worksheets(wkbNames.ValueAsString("RNG_CustomDrop"))
    On Error GoTo 0
    If dropsh Is Nothing Then Exit Sub

    Set choi = LLChoices.Create(impwb.Worksheets(CHOICESSHEET), 1, 1)
    Set drop = DropdownLists.Create(dropsh, vbNullString)
    Set updateData = New BetterArray

    For Each lo In dropsh.ListObjects

        updateData.Clear
        listName = lo.Name
        listName = Right$(listName, Len(listName) - 5)
        choiceName = "__choice_custom_" & listName

        On Error Resume Next
        Set updateData = choi.Categories(choiceName)
        On Error GoTo 0

        If updateData.Length > 0 Then
            drop.Update updateData, listName, pasteAtBottom
        End If

    Next lo
End Sub


'@section Migration Metadata
'===============================================================================

'@label:import-show-hide
'@sub-title Import show/hide visibility state from a migration workbook
'@details
'Reads the show_hide_migration ListObject from the import workbook and
'applies visibility and column/row sizes to each data sheet via
'ShowHideManager.ImportForMigration. Silently skips when HasSameLanguage
'returned False, when the import workbook lacks the required Metadata or
'show_hide__ sheets, or when the metadata tag is not set to "update on
'import".
'@param impwb Workbook. The opened migration workbook.
'@export
Public Sub ImportShowHide(ByVal impwb As Workbook)

    Dim metaTab As IDataSheet
    Dim varRng As Range
    Dim cellRng As Range
    Dim updateShowHide As Boolean
    Dim dict As ILLdictionary
    Dim sheetNames As BetterArray
    Dim sheetName As String
    Dim sheetType As String
    Dim sh As Worksheet
    Dim printSh As Worksheet
    Dim lo As ListObject
    Dim shm As IShowHideManager
    Dim impsh As Worksheet
    Dim counter As Long

    ' Prerequisites
    If Not this.hasSameLanguage Then Exit Sub
    If Not SheetExists(impwb, METADATASHEET) Then Exit Sub
    If Not SheetExists(impwb, SHOWHIDESHEET) Then Exit Sub

    ' Check metadata tag
    Set metaTab = DataSheet.Create(impwb.Worksheets(METADATASHEET), 1, 1)
    Set varRng = metaTab.DataRange("variable")
    updateShowHide = False

    On Error Resume Next
    Set cellRng = varRng.Find(What:="show_hide", lookAt:=xlWhole, MatchCase:=False)
    updateShowHide = (CStr(cellRng.Offset(, 1).Value) = "update on import")
    On Error GoTo 0

    If Not updateShowHide Then Exit Sub

    ' Get show/hide ListObject from import workbook
    Set impsh = impwb.Worksheets(SHOWHIDESHEET)
    If impsh.ListObjects.Count = 0 Then Exit Sub
    Set lo = impsh.ListObjects(1)

    ' Build dictionary for ShowHideManager creation
    Set dict = SourceDictionary()
    Set sheetNames = dict.UniqueValues("sheet name")

    ' For each data sheet, apply imported visibility
    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)

        Set sh = Nothing
        On Error Resume Next
        Set sh = this.sourceWkb.Worksheets(sheetName)
        On Error GoTo 0
        If sh Is Nothing Then GoTo NextShowHideSheet

        sheetType = CStr(sh.Cells(1, 3).Value)

        If sheetType = "HList" Then
            ' HList layer
            Set shm = ShowHideManager.Create(dict, ShowHideLayerHList, sheetName)
            shm.ImportForMigration lo, sh

            ' Printed layer (companion print sheet)
            Set printSh = Nothing
            On Error Resume Next
            Set printSh = this.sourceWkb.Worksheets("print_" & sheetName)
            On Error GoTo 0

            If Not printSh Is Nothing Then
                Set shm = ShowHideManager.Create(dict, ShowHideLayerPrinted, sheetName)
                shm.ImportForMigration lo, printSh
            End If

        ElseIf sheetType = "VList" Then
            ' VList layer
            Set shm = ShowHideManager.Create(dict, ShowHideLayerVList, sheetName)
            shm.ImportForMigration lo, sh
        End If

NextShowHideSheet:
    Next counter
End Sub

'@label:import-editable-labels
'@sub-title Import editable label values from a migration workbook
'@details
'Reads the import workbook dictionary for editable label values and
'updates both the worksheet cells and dictionary entries in the source
'workbook. Silently skips when HasSameLanguage returned False, the import
'dictionary is not prepared, or the metadata tag is not set to "update on
'import".
'@param impwb Workbook. The opened migration workbook.
'@export
Public Sub ImportEditableLabels(ByVal impwb As Workbook)

    Dim metaTab As IDataSheet
    Dim varRng As Range
    Dim cellRng As Range
    Dim updateLabels As Boolean
    Dim dict As ILLdictionary
    Dim inpDict As ILLdictionary
    Dim currVars As ILLVariables
    Dim inpVars As ILLVariables
    Dim counter As Long
    Dim varName As String
    Dim newLabel As String
    Dim sheetName As String
    Dim sheetInfo As String
    Dim sh As Worksheet
    Dim colIndex As Long

    ' Prerequisites
    If Not this.hasSameLanguage Then Exit Sub
    If Not SheetExists(impwb, METADATASHEET) Then Exit Sub
    If Not SheetExists(impwb, DICTIONARYSHEET) Then Exit Sub

    ' Check imported dictionary is prepared
    Set inpDict = LLdictionary.Create(impwb.Worksheets(DICTIONARYSHEET), 1, 1)
    If Not inpDict.Prepared Then Exit Sub

    ' Check metadata tag
    Set metaTab = DataSheet.Create(impwb.Worksheets(METADATASHEET), 1, 1)
    Set varRng = metaTab.DataRange("variable")
    updateLabels = False

    On Error Resume Next
    Set cellRng = varRng.Find(What:="editable_labels", lookAt:=xlWhole, MatchCase:=False)
    updateLabels = (CStr(cellRng.Offset(, 1).Value) = "update on import")
    On Error GoTo 0

    If Not updateLabels Then Exit Sub

    ' Build variables objects
    Set dict = SourceDictionary()
    Set currVars = LLVariables.Create(dict)
    Set inpVars = LLVariables.Create(inpDict)
    Set varRng = dict.DataRange("variable name")

    For counter = 1 To varRng.Rows.Count
        varName = CStr(varRng.Cells(counter, 1).Value)

        ' Skip variables not in import dictionary
        If Not inpVars.Contains(varName) Then GoTo NextLabel
        ' Skip non-editable labels
        If currVars.Value(varName:=varName, colName:="editable label") <> "yes" Then GoTo NextLabel

        ' Get the new label from import
        newLabel = inpVars.Value(varName:=varName, colName:="main label")

        ' Get column/row index
        colIndex = 0
        On Error Resume Next
        colIndex = CLng(currVars.Value(varName:=varName, colName:="column index"))
        On Error GoTo 0
        If colIndex = 0 Then GoTo NextLabel

        ' Get sheet and type
        sheetName = currVars.SheetName(varName)
        sheetInfo = currVars.Value(varName:=varName, colName:="sheet type")

        Set sh = Nothing
        On Error Resume Next
        Set sh = this.sourceWkb.Worksheets(sheetName)
        On Error GoTo 0
        If sh Is Nothing Then GoTo NextLabel

        ' Update label in worksheet
        Set cellRng = Nothing
        On Error Resume Next
        Set cellRng = sh.Range(CStr(sh.Cells(1, 4).Value) & "_START")
        On Error GoTo 0
        If cellRng Is Nothing Then GoTo NextLabel

        If sheetInfo = "hlist2D" Then
            sh.Cells(cellRng.Offset(-1).Row, colIndex).Value = newLabel
        ElseIf sheetInfo = "vlist1D" Then
            sh.Cells(colIndex, cellRng.Offset(, -1).Column).Value = newLabel
        End If

        ' Update label in dictionary
        currVars.SetValue varName:=varName, colName:="main label", newValue:=newLabel

NextLabel:
    Next counter
End Sub

'@label:import-single-values
'@sub-title Import single metadata values from a migration workbook
'@details
'Reads the epiweek start value from the import workbook Metadata sheet
'and updates the RNG_EpiWeekStart named range in the source workbook.
'Silently skips when HasSameLanguage returned False or the Metadata
'sheet is missing.
'@param impwb Workbook. The opened migration workbook.
'@export
Public Sub ImportSingleValues(ByVal impwb As Workbook)

    Dim metaTab As IDataSheet
    Dim varRng As Range
    Dim cellRng As Range
    Dim varValue As String

    If Not this.hasSameLanguage Then Exit Sub
    If Not SheetExists(impwb, METADATASHEET) Then Exit Sub

    Set metaTab = DataSheet.Create(impwb.Worksheets(METADATASHEET), 1, 1)
    Set varRng = metaTab.DataRange("variable")

    ' Retrieve and update epiweek start
    varValue = vbNullString
    On Error Resume Next
    Set cellRng = varRng.Find(What:="epiweekstart", lookAt:=xlWhole, MatchCase:=False)
    varValue = CStr(cellRng.Offset(, 1).Value)
    On Error GoTo 0

    If LenB(varValue) > 0 Then
        On Error Resume Next
        this.sourceWkb.Names(RNGEPIWEEKSTART).RefersToRange.Value = varValue
        On Error GoTo 0
    End If
End Sub

'@label:import-geobase
'@sub-title Import geobase data from a workbook
'@details
'Delegates to LLGeo.Import to load geobase data. When not histoOnly,
'also updates geobase headers in HList worksheets via geo.Update,
'dictionary geo labels via geo.UpdateDict, and show/hide tables for geo
'variables via geo.UpdateShowHide.
'@param impwb Workbook. The opened geobase workbook.
'@param histoOnly Optional Boolean. When True, imports only historic geobase data. Defaults to False.
'@export
Public Sub ImportGeobase(ByVal impwb As Workbook, Optional ByVal histoOnly As Boolean = False)

    Dim geo As ILLGeo
    Dim dict As ILLdictionary
    Dim pass As IPasswords
    Dim showHideSh As Worksheet

    Set geo = GeoObject()
    geo.Import impwb, histoOnly

    If Not histoOnly Then
        Set pass = PasswordsObject()
        Set dict = SourceDictionary()

        ' Update geobase headers in HList worksheets
        geo.Update pass
        ' Update dictionary geo variable labels
        geo.UpdateDict dict

        ' Update show/hide geo variables
        Set showHideSh = Nothing
        On Error Resume Next
        Set showHideSh = this.sourceWkb.Worksheets(SHOWHIDESHEET)
        On Error GoTo 0

        If Not showHideSh Is Nothing Then
            geo.UpdateShowHide showHideSh, dict
        End If
    End If
End Sub


'@section Import Report
'===============================================================================

'@label:need-report
'@prop-title Whether the last import had issues
'@details
'Returns True when at least one report entry exists, meaning some sheets
'or variables were not imported or not touched during the last import.
'@return Boolean. True when the import report has entries.
'@export
Public Property Get NeedReport() As Boolean
    NeedReport = this.needReport
End Property

'@label:finalize-report
'@sub-title Build the "not touched" part of the import report
'@details
'Compares source workbook sheets and variables against the imported items
'tracked during ImportData. Identifies sheets not touched and variables
'not touched, skipping formula-type variables (case_when, formula,
'choice_formula). Call after ImportData to complete the report.
'@export
Public Sub FinalizeReport()

    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim sheetsList As BetterArray
    Dim currVarsList As BetterArray
    Dim sheetName As String
    Dim varName As String
    Dim varControl As String
    Dim counter As Long
    Dim varCounter As Long

    Set dict = SourceDictionary()
    Set vars = LLVariables.Create(dict)
    Set sheetsList = dict.UniqueValues("sheet name")

    Set this.sheetsNotTouched = New BetterArray
    Set this.varsNotTouched = New BetterArray

    ' Sheets not touched: in current but not imported
    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        If Not this.importedSheets.Includes(sheetName) Then
            this.sheetsNotTouched.Push sheetName
            this.needReport = True
        End If
    Next counter

    ' Variables not touched: in current but not imported (skip formulas)
    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)

        ' Only check variables on sheets that were touched
        If this.sheetsNotTouched.Includes(sheetName) Then GoTo NextReportSheet

        Set currVarsList = dict.Data.FilterData("sheet name", sheetName, "variable name")

        For varCounter = currVarsList.LowerBound To currVarsList.UpperBound
            varName = CStr(currVarsList.Item(varCounter))
            varControl = vars.ControlType(varName)

            ' Skip formulas (they are recomputed)
            If varControl <> "case_when" And varControl <> "formula" And _
               varControl <> "choice_formula" Then
                If Not this.importedVars.Includes(varName) Then
                    this.varsNotTouched.Push Array(varName, sheetName)
                    this.needReport = True
                End If
            End If
        Next varCounter

NextReportSheet:
    Next counter
End Sub

'@label:report-sheets
'@sub-title Report sheet names for the given scope
'@details
'Returns a BetterArray of sheet names that were either not imported
'(ImportReportNotImported) or not touched (ImportReportNotTouched).
'Returns an empty BetterArray when no entries exist for the scope.
'@param scope Byte. ImportReportNotImported or ImportReportNotTouched.
'@return BetterArray. Sheet names matching the scope, or empty.
'@export
Public Function ReportSheets(ByVal scope As Byte) As BetterArray

    Select Case scope

    Case ImportReportNotImported
        If this.sheetsNotImported Is Nothing Then
            Set ReportSheets = New BetterArray
        Else
            Set ReportSheets = this.sheetsNotImported
        End If

    Case ImportReportNotTouched
        If this.sheetsNotTouched Is Nothing Then
            Set ReportSheets = New BetterArray
        Else
            Set ReportSheets = this.sheetsNotTouched
        End If

    Case Else
        Set ReportSheets = New BetterArray

    End Select
End Function

'@label:report-variables
'@sub-title Report variable entries for the given scope
'@details
'Returns a BetterArray where each entry is a Variant array of
'Array(variableName, sheetName). Returns an empty BetterArray when no
'entries exist for the scope.
'@param scope Byte. ImportReportNotImported or ImportReportNotTouched.
'@return BetterArray. Variant arrays of (variableName, sheetName), or empty.
'@export
Public Function ReportVariables(ByVal scope As Byte) As BetterArray

    Select Case scope

    Case ImportReportNotImported
        If this.varsNotImported Is Nothing Then
            Set ReportVariables = New BetterArray
        Else
            Set ReportVariables = this.varsNotImported
        End If

    Case ImportReportNotTouched
        If this.varsNotTouched Is Nothing Then
            Set ReportVariables = New BetterArray
        Else
            Set ReportVariables = this.varsNotTouched
        End If

    Case Else
        Set ReportVariables = New BetterArray

    End Select
End Function

'@label:clear-report
'@sub-title Reset report state
'@details
'Clears all report tracking arrays and resets the needReport flag. Call
'before starting a new import session to ensure a clean report.
'@export
Public Sub ClearReport()
    this.needReport = False
    Set this.sheetsNotImported = New BetterArray
    Set this.sheetsNotTouched = New BetterArray
    Set this.varsNotImported = New BetterArray
    Set this.varsNotTouched = New BetterArray
    Set this.importedSheets = New BetterArray
    Set this.importedVars = New BetterArray
End Sub


'@section Private - Source Object Accessors
'===============================================================================

'@label:source-dictionary
'@sub-title Get the dictionary from the source workbook
'@return ILLdictionary. Dictionary instance from the Dictionary sheet.
Private Function SourceDictionary() As ILLdictionary
    Set SourceDictionary = LLdictionary.Create( _
        this.sourceWkb.Worksheets(DICTIONARYSHEET), 1, 1)
End Function

'@label:passwords-object
'@sub-title Get passwords from the source workbook
'@return IPasswords. Passwords instance from the __pass sheet.
Private Function PasswordsObject() As IPasswords
    Set PasswordsObject = Passwords.Create(this.sourceWkb.Worksheets(PASSWORDSHEET))
End Function

'@label:geo-object
'@sub-title Get the geo object from the source workbook
'@return ILLGeo. Geo instance from the Geo sheet.
Private Function GeoObject() As ILLGeo
    Set GeoObject = LLGeo.Create(this.sourceWkb.Worksheets(GEOSHEET))
End Function


'@section Private - Import Helpers
'===============================================================================

'@label:initialize-tracking
'@sub-title Initialize tracking state for a new import session
Private Sub InitializeTracking()
    this.needReport = False
    Set this.sheetsNotImported = New BetterArray
    Set this.varsNotImported = New BetterArray
    Set this.importedSheets = New BetterArray
    Set this.importedVars = New BetterArray
End Sub

'@label:is-reserved-sheet
'@sub-title Check if a sheet name is reserved (metadata/infrastructure)
'@param sheetName String. The sheet name to check.
'@return Boolean. True when the name matches a reserved sheet.
Private Function IsReservedSheet(ByVal sheetName As String) As Boolean
    IsReservedSheet = (sheetName = DICTIONARYSHEET) Or _
                      (sheetName = DICTTRANSLATIONSHEET) Or _
                      (sheetName = CHOICESSHEET) Or _
                      (sheetName = METADATASHEET) Or _
                      (sheetName = SHOWHIDESHEET)
End Function

'@label:import-vlist-sheet
'@sub-title Import VList data from a single import sheet to the current sheet
'@details
'Iterates over the import sheet used range rows, resolving each variable
'name to a named range on the current sheet. Copies the value when the
'target cell has no formula. Tracks variables not found as "not imported"
'report entries.
'@param currsh Worksheet. The current source VList worksheet.
'@param impsh Worksheet. The import VList worksheet.
Private Sub ImportVListSheet(ByVal currsh As Worksheet, ByVal impsh As Worksheet)

    Dim usedRng As Range
    Dim vlistRng As Range
    Dim varName As String
    Dim hasFormula As Boolean
    Dim counter As Long

    Set usedRng = impsh.UsedRange

    For counter = 2 To usedRng.Rows.Count
        varName = CStr(usedRng.Cells(counter, 1).Value)

        Set vlistRng = Nothing
        hasFormula = True

        On Error Resume Next
        Set vlistRng = currsh.Range(varName)
        hasFormula = (CStr(vlistRng.Formula) <> CStr(vlistRng.Value))
        On Error GoTo 0

        If vlistRng Is Nothing Then
            ' Variable from import not found in current sheet
            this.varsNotImported.Push Array(varName, impsh.Name)
            this.needReport = True
        Else
            If Not hasFormula Then
                vlistRng.Value = usedRng.Cells(counter, 2).Value
            End If
        End If

        this.importedVars.Push varName
    Next counter
End Sub

'@label:import-hlist-sheet
'@sub-title Import HList data from a single import sheet to the current sheet
'@details
'Creates a CustomTable from the current sheet ListObject and a DataSheet
'from the import sheet, then delegates to CustomTable.Import. Tracks
'columns not found in the current table as "not imported" report entries.
'@param currsh Worksheet. The current source HList worksheet.
'@param impsh Worksheet. The import HList worksheet.
'@param pasteAtBottom Boolean. When True, appends data below existing rows.
Private Sub ImportHListSheet(ByVal currsh As Worksheet, ByVal impsh As Worksheet, _
                             ByVal pasteAtBottom As Boolean)

    Dim currTab As ICustomTable
    Dim impTab As IDataSheet
    Dim varsNotImpList As BetterArray
    Dim impHeaders As BetterArray
    Dim counter As Long

    Set currTab = CustomTable.Create(currsh.ListObjects(1))
    Set impTab = DataSheet.Create(impsh, 1, 1, True)

    currTab.Import impTab, pasteAtBottom, True, False

    ' Track variables not imported (columns not found in current table)
    If currTab.HasColumnsNotImported Then
        Set varsNotImpList = currTab.ImportColumnsNotFound()
        For counter = varsNotImpList.LowerBound To varsNotImpList.UpperBound
            this.varsNotImported.Push Array(CStr(varsNotImpList.Item(counter)), currsh.Name)
        Next counter
        this.needReport = True
    End If

    ' Track all imported variable names
    Set impHeaders = New BetterArray
    impHeaders.FromExcelRange impTab.HeaderRange()
    this.importedVars.Concat impHeaders.Items
End Sub


'@section Private - Helpers
'===============================================================================

'@label:sheet-exists
'@sub-title Check if a worksheet exists in a workbook
'@param wb Workbook. The workbook to search.
'@param sheetName String. The sheet name to look for.
'@return Boolean. True when the sheet exists.
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = (Not sh Is Nothing)
End Function


'@section Private - Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a class-level exception
'@details
'Raises a VBA error using vbObjectError + 516, standardising the source
'to CLASS_NAME.
'@param message String. Human-readable description of the failure.
'@throws vbObjectError Always raises the error.
Private Sub ThrowError(ByVal message As String)
    Err.Raise VBA.vbObjectError + 516, CLASS_NAME, message
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ILLImporter contract.
'See the corresponding members above for full documentation.

Private Property Get ILLImporter_HasData() As Boolean
    ILLImporter_HasData = HasData
End Property

Private Function ILLImporter_HasSameLanguage(ByVal impwb As Workbook) As Boolean
    ILLImporter_HasSameLanguage = HasSameLanguage(impwb)
End Function

Private Sub ILLImporter_ClearData()
    ClearData
End Sub

Private Sub ILLImporter_ImportData(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)
    ImportData impwb, pasteAtBottom
End Sub

Private Sub ILLImporter_ImportCustomDropdown(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)
    ImportCustomDropdown impwb, pasteAtBottom
End Sub

Private Sub ILLImporter_FinalizeReport()
    FinalizeReport
End Sub

Private Sub ILLImporter_ImportShowHide(ByVal impwb As Workbook)
    ImportShowHide impwb
End Sub

Private Sub ILLImporter_ImportEditableLabels(ByVal impwb As Workbook)
    ImportEditableLabels impwb
End Sub

Private Sub ILLImporter_ImportSingleValues(ByVal impwb As Workbook)
    ImportSingleValues impwb
End Sub

Private Sub ILLImporter_ImportGeobase(ByVal impwb As Workbook, Optional ByVal histoOnly As Boolean = False)
    ImportGeobase impwb, histoOnly
End Sub

Private Property Get ILLImporter_NeedReport() As Boolean
    ILLImporter_NeedReport = NeedReport
End Property

Private Function ILLImporter_ReportSheets(ByVal scope As Byte) As BetterArray
    Set ILLImporter_ReportSheets = ReportSheets(scope)
End Function

Private Function ILLImporter_ReportVariables(ByVal scope As Byte) As BetterArray
    Set ILLImporter_ReportVariables = ReportVariables(scope)
End Function

Private Sub ILLImporter_ClearReport()
    ClearReport
End Sub
