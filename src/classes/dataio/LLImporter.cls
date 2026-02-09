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
'@depends ILLImporter, ILLdictionary, IPasswords, ILLTranslations
'@depends ICustomTable, IDataSheet, IShowHideManager, ILLGeo, ILLChoices
'@depends ILLVariables, IDropdownLists

Option Explicit
Implements ILLImporter


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "LLImporter"

' Source workbook sheet names
Private Const DICTIONARYSHEET As String = "Dictionary"
Private Const CHOICESSHEET As String = "Choices"
Private Const DICTTRANSLATIONSHEET As String = "Translations"
Private Const METADATASHEET As String = "Metadata"
Private Const MESSTRANSLATIONSHEET As String = "LinelistTranslation"
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

' @description Create an importer from the source linelist workbook.
' @param sourceWkb Workbook. The linelist workbook that will receive imported data.
' @return ILLImporter ready for import operations.
' @export
Public Function Create(ByVal sourceWkb As Workbook) As ILLImporter

    If sourceWkb Is Nothing Then
        ThrowError "LLImporter requires a source workbook"
    End If

    With New LLImporter
        Set .SourceWorkbook = sourceWkb
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

Public Property Get Self() As ILLImporter
    Set Self = Me
End Property

Public Property Get SourceWorkbook() As Workbook
    Set SourceWorkbook = this.sourceWkb
End Property

Public Property Set SourceWorkbook(ByVal wb As Workbook)
    Set this.sourceWkb = wb
End Property


'@section Data checks
'===============================================================================

' @description Check whether the workbook contains any user-entered data.
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

' @description Check whether the import workbook language matches.
Public Function HasSameLanguage(ByVal impwb As Workbook) As Boolean

    Dim metaTab As IDataSheet
    Dim varRng As Range
    Dim cellRng As Range
    Dim impLang As String
    Dim currLang As String
    Dim lltrads As ILLTranslations

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
    Set lltrads = TranslationsObject()
    currLang = lltrads.Value("dictionarylanguage")

    this.hasSameLanguage = (impLang = currLang)
    HasSameLanguage = this.hasSameLanguage
End Function


'@section Data operations
'===============================================================================

' @description Clear all user-entered data from HList and VList sheets.
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

' @description Import data from HList and VList worksheets.
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

' @description Import custom dropdown choices.
Public Sub ImportCustomDropdown(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)

    Dim lltrads As ILLTranslations
    Dim choi As ILLChoices
    Dim lo As ListObject
    Dim drop As IDropdownLists
    Dim dropsh As Worksheet
    Dim listName As String
    Dim choiceName As String
    Dim updateData As BetterArray

    If Not SheetExists(impwb, CHOICESSHEET) Then Exit Sub

    Set lltrads = TranslationsObject()

    On Error Resume Next
    Set dropsh = this.sourceWkb.Worksheets(lltrads.Value("customdropdown"))
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


'@section Migration metadata
'===============================================================================

' @description Import show/hide visibility state and column widths.
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

' @description Import editable labels from a migration workbook.
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

' @description Import single values (epiweek start) from a migration workbook.
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

' @description Import geobase data.
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


'@section Import report
'===============================================================================

' @description Whether the last import had issues.
Public Property Get NeedReport() As Boolean
    NeedReport = this.needReport
End Property

' @description Build the "not touched" part of the import report.
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

' @description Report sheet names for the given scope.
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

' @description Report variable entries for the given scope.
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

' @description Reset report state.
Public Sub ClearReport()
    this.needReport = False
    Set this.sheetsNotImported = New BetterArray
    Set this.sheetsNotTouched = New BetterArray
    Set this.varsNotImported = New BetterArray
    Set this.varsNotTouched = New BetterArray
    Set this.importedSheets = New BetterArray
    Set this.importedVars = New BetterArray
End Sub


'@section Private - Source object accessors
'===============================================================================

' @description Get the dictionary from the source workbook.
Private Function SourceDictionary() As ILLdictionary
    Set SourceDictionary = LLdictionary.Create( _
        this.sourceWkb.Worksheets(DICTIONARYSHEET), 1, 1)
End Function

' @description Get passwords from the source workbook.
Private Function PasswordsObject() As IPasswords
    Set PasswordsObject = Passwords.Create(this.sourceWkb.Worksheets(PASSWORDSHEET))
End Function

' @description Get LLTranslations from the source workbook.
Private Function TranslationsObject() As ILLTranslations
    Set TranslationsObject = LLTranslations.Create( _
        this.sourceWkb.Worksheets(MESSTRANSLATIONSHEET), _
        this.sourceWkb.Worksheets(DICTTRANSLATIONSHEET))
End Function

' @description Get the geo object from the source workbook.
Private Function GeoObject() As ILLGeo
    Set GeoObject = LLGeo.Create(this.sourceWkb.Worksheets(GEOSHEET))
End Function


'@section Private - Import helpers
'===============================================================================

' @description Initialize tracking state for a new import session.
Private Sub InitializeTracking()
    this.needReport = False
    Set this.sheetsNotImported = New BetterArray
    Set this.varsNotImported = New BetterArray
    Set this.importedSheets = New BetterArray
    Set this.importedVars = New BetterArray
End Sub

' @description Check if a sheet name is reserved (metadata/infrastructure).
Private Function IsReservedSheet(ByVal sheetName As String) As Boolean
    IsReservedSheet = (sheetName = DICTIONARYSHEET) Or _
                      (sheetName = DICTTRANSLATIONSHEET) Or _
                      (sheetName = CHOICESSHEET) Or _
                      (sheetName = METADATASHEET) Or _
                      (sheetName = SHOWHIDESHEET)
End Function

' @description Import VList data from a single import sheet to the current sheet.
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

' @description Import HList data from a single import sheet to the current sheet.
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

' @description Check if a worksheet exists in a workbook.
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = (Not sh Is Nothing)
End Function


'@section Private - Error handling
'===============================================================================

' @description Raise a class-level exception.
Private Sub ThrowError(ByVal message As String)
    Err.Raise VBA.vbObjectError + 516, CLASS_NAME, message
End Sub


'@section Interface implementation
'===============================================================================

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
