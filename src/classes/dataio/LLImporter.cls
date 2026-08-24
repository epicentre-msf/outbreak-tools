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

'@class LLImporter
'@description
'Imports the data of an external workbook into the linelist that holds this
'object. CheckImportFile reads the file and says whether it can be imported,
'HasSameLanguage compares the two interface languages, and CurrentLanguage
'answers the language of this linelist. ImportData brings the HList and VList
'data in, ImportCustomDropdown the custom dropdowns, ImportShowHide the
'show/hide choices, ImportEditableLabels the labels a user may edit,
'ImportSingleValues the single metadata values and ImportGeobase the geobase.
'ClearData empties the linelist first when the caller asks for it, and HasData
'says whether the linelist already carries rows.
'
'THE REPORT
'-------------------------------------------------------------------------------
'CompareWithImportFile records which sheets and which variables the two files
'have in common and which they have on their own. NeedReport says whether
'anything was recorded, FinalizeReport writes the entries into ImportReport,
'ReportSheets and ReportVariables read them back, and ClearReport empties them.
'Entries filed along the way leave through HasCheckings and CheckingValues.
'@depends LLdictionary, Passwords, CustomTable, DataSheet
'@depends LLGeo, LLChoices, LLVariables, DropdownLists, HiddenNames
'@depends ShowHide, ShowHideLayout, ShowHideStore, ChoiceKeys, ImportMetadata
'@depends Checking, BetterArray, ImportReport


'@section Enumerations
'===============================================================================
'@description Scopes for import report entries, used to query which items were missed.

Public Enum ImportReportScope
    ImportReportNotImported = 1
    ImportReportNotTouched = 2
End Enum


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "LLImporter"

' Source workbook sheet names
Private Const DICTIONARYSHEET As String = "Dictionary"
Private Const CHOICESSHEET As String = "Choices"
Private Const DICTTRANSLATIONSHEET As String = "Translations"
Private Const LINELISTTRANSLATIONSHEET As String = "LinelistTranslation"
Private Const METADATASHEET As String = "Metadata"
Private Const PASSWORDSHEET As String = "__pass"
Private Const SHOWHIDESHEET As String = "__show_hide"
'The name a linelist generated before the internal-sheet rename gave that sheet
Private Const SHOWHIDESHEET_OLD As String = "show_hide__"
Private Const GEOSHEET As String = "Geo"

' Prefixes of the two companion sheets a data entry sheet can have
Private Const PRINTPREFIX As String = "print_"
Private Const CRFPREFIX As String = "crf_"
Private Const RNGEPIWEEKSTART As String = "RNG_EpiWeekStart"


'@section State
'===============================================================================

Private Type TLLImporter
    sourceWkb As Workbook
    needReport As Boolean
    ' Report data
    sheetsNotImported As BetterArray
    sheetsNotTouched As BetterArray
    varsNotImported As BetterArray
    varsNotTouched As BetterArray
    ' Tracking
    importedSheets As BetterArray
    importedVars As BetterArray
    ' Diagnostics
    checks As Checking
    sealed As Boolean
End Type

Private this As TLLImporter


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create an importer from the source linelist workbook
'@details
'Factory method that validates the workbook is not Nothing, then constructs
'and returns a new LLImporter instance. Called on the predeclared instance.
'@param sourceWkb Workbook. The linelist workbook that will receive imported data.
'@return LLImporter. A fully initialised instance ready for import operations.
'@throws vbObjectError When the workbook is Nothing.
'@export
Public Function Create(ByVal sourceWkb As Workbook) As LLImporter

    If sourceWkb Is Nothing Then
        ThrowError "LLImporter requires a source workbook"
    End If

    Dim newImporter As LLImporter
    Set newImporter = New LLImporter
    Set newImporter.SourceWorkbook = sourceWkb
    newImporter.Seal
    Set Create = newImporter
End Function


'@section Public Accessors
'===============================================================================

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
    GuardNotSealed "SourceWorkbook"
    Set this.sourceWkb = wb
End Property

'@label:seal
'@prop-title Prevent further changes to setup-only properties
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Raise when a setup-only property is written after sealing
'@param propName String. The name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError propName & " is set at creation only"
End Sub


'@section Diagnostics
'===============================================================================
'@description What the import found and what it decided to skip. Every entry is
'in English and the worksheet the caller writes them to is what the user reads.

'@label:has-checkings
'@prop-title Whether the import filed anything worth reading
'@return Boolean. True when at least one entry was filed.
'@export
Public Property Get HasCheckings() As Boolean
    If this.checks Is Nothing Then Exit Property
    HasCheckings = (this.checks.Length > 0)
End Property

'@label:checking-values
'@prop-title The entries the import filed
'@return Checking. The store, ready to hand to a CheckingOutput.
'@export
Public Property Get CheckingValues() As Checking
    Set CheckingValues = ImportCheckings()
End Property

'@label:import-checkings
'@sub-title The checking store, built on first use
'@return Checking. The store this import files into.
Private Function ImportCheckings() As Checking
    If this.checks Is Nothing Then _
        Set this.checks = Checking.Create("Import report", "What the import found")
    Set ImportCheckings = this.checks
End Function

'@label:file-checking
'@sub-title File one entry, and never raise on a key already used
'@details
'Checking.Add raises on a key it already holds. An import walks a dictionary,
'so a key built from a variable name can repeat when the same variable is on two
'sheets. The counter suffix keeps every entry and keeps the key unique.
'@param keyName String. The key of the entry.
'@param label String. What to tell the user.
'@param scope Byte. A CheckingScope value.
Private Sub FileChecking(ByVal keyName As String, ByVal label As String, _
                         ByVal scope As Byte)

    Dim uniqueKey As String
    Dim counter As Long

    uniqueKey = keyName
    Do While ImportCheckings().KeyExists(uniqueKey)
        counter = counter + 1
        uniqueKey = keyName & " (" & CStr(counter) & ")"
    Loop

    ImportCheckings().Add uniqueKey, label, scope
End Sub


'@section Data Checks
'===============================================================================

'@label:has-data
'@prop-title Whether the workbook contains user-entered data
'@details
'Scans all HList sheets for rows carrying more filled cells than an untouched
'row does. LLDataEntry records that figure per sheet under the hidden name
'blank_row_count when it makes the table. Returns True as soon as a single row
'has user data. Used to warn the user before clearing or overwriting.
'@return Boolean. True when at least one HList row has user data.
'@export
Public Property Get HasData() As Boolean

    Dim sh As Worksheet
    Dim lo As ListObject
    Dim nbBlanks As Long
    Dim counter As Long
    Dim shHnData As HiddenNames

    For Each sh In this.sourceWkb.Worksheets
        If SheetTagOf(sh) = "HList" Then
            Set shHnData = HiddenNames.Create(sh)
            nbBlanks = shHnData.ValueAsLong("blank_row_count")
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

'@label:current-language
'@prop-title The language this linelist was built in
'@details
'The workbook-level RNG_DictionaryLanguage hidden name, which the build set to
'the dictionary language. The Metadata sheet of an export carries the same fact
'under `language`, so the two compare like with like.
'@return String. The language code, empty when the workbook carries none.
'@export
Public Property Get CurrentLanguage() As String
    Dim wkbNames As HiddenNames
    Set wkbNames = HiddenNames.Create(this.sourceWkb)
    CurrentLanguage = wkbNames.ValueAsString("RNG_DictionaryLanguage")
End Property

'@label:has-same-language
'@sub-title Check whether the file being imported is in this linelist's language
'@details
'Compares the `language` tag of the file with CurrentLanguage. A file naming no
'language answers False, which is also the answer for a file carrying no
'Metadata sheet at all. The caller tells those two apart through the metadata
'object and says which happened.
'
'The answer used to be kept on the class and read by the three metadata imports,
'so a caller that skipped this step got False by default and all three did
'nothing with no diagnostic. The three take the metadata object now and the
'caller decides whether to call them.
'@param meta ImportMetadata. The metadata of the file being imported.
'@return Boolean. True when the languages match.
'@export
Public Function HasSameLanguage(ByVal meta As ImportMetadata) As Boolean

    Dim impLang As String

    If meta Is Nothing Then Exit Function

    impLang = meta.Language
    If LenB(impLang) = 0 Then Exit Function

    HasSameLanguage = (StrComp(impLang, CurrentLanguage, vbTextCompare) = 0)
End Function


'@label:check-import-file
'@sub-title Read the file over before importing it, and file what is wrong
'@details
'Every worksheet an export is supposed to carry is looked for, and the export
'options are read. The dictionary and the choices are errors when they are
'absent, because the labels and the category lists cannot be compared without
'them; the other metadata sheets are warnings, because the data still lands.
'
'ONE THING STOPS AN IMPORT
'-------------------------------------------------------------------------------
'A custom export written with the labels as its header row carries no variable
'name anywhere, so no column of it can be matched to a column of this linelist.
'Every row of such a file used to land in "not imported" with nothing said. It
'is refused here, with the reason.
'@param impwb Workbook. The opened import workbook.
'@param meta ImportMetadata. What that workbook says about itself.
'@return Boolean. True when the import may go on.
'@export
Public Function CheckImportFile(ByVal impwb As Workbook, _
                                ByVal meta As ImportMetadata) As Boolean

    CheckImportFile = True

    If Not meta.Exists Then
        FileChecking "metadata sheet", _
                     "The file carries no Metadata worksheet, so nothing is " & _
                     "known about how it was written. It is read as a file " & _
                     "with the variable names on the first row.", _
                     checkingWarning
    End If

    If Not SheetExists(impwb, DICTIONARYSHEET) Then
        FileChecking "dictionary sheet", _
                     "The file carries no Dictionary worksheet. The labels and " & _
                     "the choice lists of this linelist cannot be compared with " & _
                     "the ones the file was written from.", _
                     checkingError
    End If

    If Not SheetExists(impwb, CHOICESSHEET) Then
        FileChecking "choices sheet", _
                     "The file carries no Choices worksheet. The custom " & _
                     "dropdowns are left as they are and the category lists " & _
                     "cannot be compared.", _
                     checkingError
    End If

    If ShowHideSheetOf(impwb) Is Nothing Then
        FileChecking "show hide sheet", _
                     "The file carries no show/hide worksheet, so the " & _
                     "hidden and shown state of this linelist is left as it is.", _
                     checkingWarning
    End If

    If Not SheetExists(impwb, LINELISTTRANSLATIONSHEET) Then
        FileChecking "linelist translation sheet", _
                     "The file carries no LinelistTranslation worksheet.", _
                     checkingWarning
    End If

    If Not SheetExists(impwb, DICTTRANSLATIONSHEET) Then
        FileChecking "dictionary translation sheet", _
                     "The file carries no Translations worksheet.", _
                     checkingWarning
    End If

    If Not meta.CarriesVariableNames Then
        FileChecking "header row", _
                     "This file was exported with the variable labels as its " & _
                     "header row (" & meta.ExportType & "), so it carries no " & _
                     "variable name at all and no column of it can be matched " & _
                     "to a column of this linelist. Export it again with a " & _
                     "header format that keeps the variable names.", _
                     checkingError
        CheckImportFile = False
    End If
End Function


'@section Data Operations
'===============================================================================

'@label:clear-data
'@sub-title Clear all user-entered data from HList and VList sheets
'@details
'Iterates over all worksheets in the source workbook. For HList sheets,
'clears non-formula column values from the ListObject data body. For
'VList sheets, clears non-formula named range values. Unprotects and
're-protects each sheet via Passwords.
'@export
Public Sub ClearData()

    Dim sh As Worksheet
    Dim pass As Passwords
    Dim lo As ListObject
    Dim valueRng As Range
    Dim labelRng As Range
    Dim counter As Long
    Dim actsh As Worksheet

    Set pass = PasswordsObject()
    Set actsh = ActiveSheet

    For Each sh In this.sourceWkb.Worksheets

        Select Case SheetTagOf(sh)

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
'@param meta ImportMetadata. The metadata of that workbook. It says which row of
'  a data sheet holds the variable names, which is row 2 for a file written with
'  the labels above them.
'@export
Public Sub ImportData(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean, _
                      ByVal meta As ImportMetadata)

    Dim dict As LLdictionary
    Dim sheetsList As BetterArray
    Dim pass As Passwords
    Dim impsh As Worksheet
    Dim currsh As Worksheet
    Dim headerRow As Long

    headerRow = 1
    If Not meta Is Nothing Then headerRow = meta.HeaderRow

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

            Select Case SheetTagOf(currsh)

            Case "VList"
                pass.UnProtect currsh.Name
                ImportVListSheet currsh, impsh
                pass.Protect currsh.Name

            Case "HList"
                pass.UnProtect currsh.Name
                ImportHListSheet currsh, impsh, pasteAtBottom, headerRow
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
'Walks the dropdown registry of the source workbook, reads the values the
'import workbook holds for each dropdown, and writes them back. Silently exits
'when the import workbook has no Choices sheet or the custom dropdown sheet does
'not exist.
'
'The walk goes through DropdownLists.AllDropdowns, so each name is the one the
'registry holds and the one Update takes. Reading the ListObject names off the
'worksheet instead is what used to break this: the generated name ends with the
'workbook counter, cutting the five characters of "list_" off the front left
'that counter attached, and the lookup missed for every dropdown ever created.
'ChoiceKeys builds the key both ends of the round trip use.
'@param impwb Workbook. The opened import workbook.
'@param pasteAtBottom Boolean. When True, appends choices below existing entries.
'@export
Public Sub ImportCustomDropdown(ByVal impwb As Workbook, ByVal pasteAtBottom As Boolean)

    Dim choi As LLChoices
    Dim drop As DropdownLists
    Dim keys As ChoiceKeys
    Dim dropNames As BetterArray
    Dim updateData As BetterArray
    Dim listName As String
    Dim counter As Long

    If Not SheetExists(impwb, CHOICESSHEET) Then
        FileChecking "custom dropdowns", _
                     "The file carries no Choices worksheet, so the custom " & _
                     "dropdowns of this linelist are left as they are.", _
                     checkingWarning
        Exit Sub
    End If

    Set drop = CustomDropdowns()
    If drop Is Nothing Then Exit Sub

    Set choi = LLChoices.Create(impwb.Worksheets(CHOICESSHEET), 1, 1)
    Set keys = ChoiceKeys.Create()
    Set dropNames = drop.AllDropdowns

    For counter = dropNames.LowerBound To dropNames.UpperBound

        listName = CStr(dropNames.Item(counter))
        Set updateData = choi.Categories(keys.CustomChoiceName(listName))

        If updateData.Length > 0 Then
            drop.Update updateData, listName, pasteAtBottom
        End If

    Next counter
End Sub


'@section Migration Metadata
'===============================================================================

'@label:import-show-hide
'@sub-title Import show/hide choices from a migration workbook
'@details
'Reads the show/hide table of the import workbook and puts each data sheet of
'the running workbook in the state it describes: visibility, sizes and, on a
'printed sheet, the direction of each header. Every layer a sheet has is
'covered - HList and CRF for a data entry sheet, Printed for its companion,
'VList for an info sheet.
'
'Two faults used to sit here. The visibility read out of the file was written
'into memory and never reached a sheet, so nothing was ever hidden; and a blank
'width cell was read as a width of zero and written to the sheet, which is how
'Excel hides a column, so entries vanished for the opposite reason.
'
'Skips when the file carries no show/hide sheet, and when its `show_hide` tag
'asks for no update.
'@param impwb Workbook. The opened migration workbook.
'@param meta ImportMetadata. The metadata of that workbook.
'@export
Public Sub ImportShowHide(ByVal impwb As Workbook, ByVal meta As ImportMetadata)

    Dim dict As LLdictionary
    Dim sheetNames As BetterArray
    Dim sheetName As String
    Dim sheetType As String
    Dim sh As Worksheet
    Dim printSh As Worksheet
    Dim crfSh As Worksheet
    Dim impsh As Worksheet
    Dim impStore As ShowHideStore
    Dim localStore As ShowHideStore
    Dim skippedNames As BetterArray
    Dim counter As Long

    ' Prerequisites
    If meta Is Nothing Then Exit Sub
    If Not meta.UpdatesOnImport("show_hide") Then Exit Sub

    ' The show/hide table of the file being imported
    Set impsh = ShowHideSheetOf(impwb)
    If impsh Is Nothing Then Exit Sub
    Set impStore = ShowHideStore.CreateForRead(impsh)
    If Not impStore.HasTable Then Exit Sub

    Set dict = SourceDictionary()
    Set sheetNames = dict.UniqueValues("sheet name")

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)

        Set sh = Nothing
        On Error Resume Next
        Set sh = this.sourceWkb.Worksheets(sheetName)
        On Error GoTo 0
        If sh Is Nothing Then GoTo NextShowHideSheet

        sheetType = SheetTagOf(sh)

        If sheetType = "HList" Then
            ApplyOneShowHideLayer impStore, dict, ShowHideLayerHList, sheetName, sh

            Set printSh = CompanionSheet(PRINTPREFIX & sheetName)
            If Not printSh Is Nothing Then
                ApplyOneShowHideLayer impStore, dict, ShowHideLayerPrinted, sheetName, printSh
            End If

            Set crfSh = CompanionSheet(CRFPREFIX & sheetName)
            If Not crfSh Is Nothing Then
                ApplyOneShowHideLayer impStore, dict, ShowHideLayerCRF, sheetName, crfSh
            End If

        ElseIf sheetType = "VList" Then
            ApplyOneShowHideLayer impStore, dict, ShowHideLayerVList, sheetName, sh
        End If

NextShowHideSheet:
    Next counter

    'The file's saved layouts come across too. A name both stores hold is
    'taken from the file; a new name past the cap is skipped and reported.
    On Error Resume Next
    Set localStore = ShowHideStore.Create(this.sourceWkb)
    On Error GoTo 0

    If Not localStore Is Nothing Then
        Set skippedNames = localStore.MergeLayoutsFrom(impStore)

        If skippedNames.Length = 0 Then Exit Sub
        For counter = skippedNames.LowerBound To skippedNames.UpperBound
            FileChecking "show hide layout " & CStr(skippedNames.Item(counter)), _
                         "The file carries a saved layout named '" & _
                         CStr(skippedNames.Item(counter)) & "', and this " & _
                         "linelist already holds " & _
                         CStr(localStore.MaxSavedLayouts) & " saved layouts, " & _
                         "so it was left out. Delete a saved layout and " & _
                         "import again to bring it in.", _
                         checkingWarning
        Next counter
    End If
End Sub

'@label:apply-one-show-hide-layer
'@sub-title Put one layer of one sheet in the state the file describes
'@details
'Load carries the sizes and the header directions onto the sheet, and Apply is
'what actually hides and shows. The choices are written into this workbook's own
'store as well, so they survive the next time it is opened.
'@param impStore ShowHideStore. The store of the import workbook.
'@param dict LLdictionary. The dictionary of the running workbook.
'@param layer Byte. A ShowHideWorksheetLayer value. Which layer to apply.
'@param sheetName String. The base sheet name.
'@param targetSh Worksheet. The sheet of that layer in the running workbook.
Private Sub ApplyOneShowHideLayer(ByVal impStore As ShowHideStore, _
                                  ByVal dict As LLdictionary, _
                                  ByVal layer As Byte, _
                                  ByVal sheetName As String, _
                                  ByVal targetSh As Worksheet)

    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim ownStore As ShowHideStore
    Dim tableName As String

    Set entries = ShowHide.Create(dict, layer, sheetName)
    If entries.EntryCount = 0 Then Exit Sub

    tableName = BaseTableNameOf(targetSh)
    Set layout = ShowHideLayout.Create(targetSh, layer, PasswordsObject(), tableName)

    impStore.Load entries, layout
    entries.Apply layout

    Set ownStore = OwnShowHideStore()
    If Not ownStore Is Nothing Then ownStore.Save entries
End Sub

'@label:own-show-hide-store
'@sub-title The show/hide store of the workbook being written into
'@return ShowHideStore. The store, or Nothing when the workbook has no
'  __show_hide worksheet.
Private Function OwnShowHideStore() As ShowHideStore
    On Error Resume Next
    Set OwnShowHideStore = ShowHideStore.Create(this.sourceWkb)
    On Error GoTo 0
End Function

'@label:companion-sheet
'@sub-title Find a prefixed companion sheet of the running workbook
'@param sheetName String. The full sheet name, prefix included.
'@return Worksheet. The sheet, or Nothing when the workbook has none.
Private Function CompanionSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set CompanionSheet = this.sourceWkb.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:base-table-name-of
'@sub-title The table name a sheet's PRINTSTART anchor is named after
'@details
'A printed companion stores its table name with the print_ prefix in front and
'the anchor is named after the base one, so the prefix comes off here.
'@param sh Worksheet. The sheet to read.
'@return String. The table name with no print_ prefix.
Private Function BaseTableNameOf(ByVal sh As Worksheet) As String
    Dim tableName As String

    tableName = SheetTagOf(sh, "table_name")
    If InStr(1, tableName, PRINTPREFIX, vbTextCompare) = 1 Then
        tableName = Mid$(tableName, Len(PRINTPREFIX) + 1)
    End If

    BaseTableNameOf = tableName
End Function

'@label:import-editable-labels
'@sub-title Import editable label values from a migration workbook
'@details
'Reads the import workbook dictionary for editable label values and
'updates both the worksheet cells and dictionary entries in the source
'workbook. Skips when the file carries no Dictionary sheet, when that
'dictionary is not prepared, and when the `editable_labels` tag asks for no
'update.
'@param impwb Workbook. The opened migration workbook.
'@param meta ImportMetadata. The metadata of that workbook.
'@export
Public Sub ImportEditableLabels(ByVal impwb As Workbook, ByVal meta As ImportMetadata)

    Dim varRng As Range
    Dim cellRng As Range
    Dim dict As LLdictionary
    Dim inpDict As LLdictionary
    Dim currVars As LLVariables
    Dim inpVars As LLVariables
    Dim counter As Long
    Dim varName As String
    Dim newLabel As String
    Dim sheetName As String
    Dim sheetInfo As String
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim shHn As HiddenNames

    ' Prerequisites
    If meta Is Nothing Then Exit Sub
    If Not meta.UpdatesOnImport("editable_labels") Then Exit Sub
    If Not SheetExists(impwb, DICTIONARYSHEET) Then Exit Sub

    ' Check imported dictionary is prepared
    Set inpDict = LLdictionary.Create(impwb.Worksheets(DICTIONARYSHEET), 1, 1)
    If Not inpDict.Prepared Then Exit Sub

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
        Set shHn = HiddenNames.Create(sh)
        Set cellRng = Nothing
        On Error Resume Next
        Set cellRng = sh.Range(shHn.ValueAsString("table_name") & "_START")
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
'Takes the epiweek start value the file carries and writes it into the
'RNG_EpiWeekStart hidden name of the source workbook, then recalculates every
'sheet that carries an epiweek formula. Skips when the file names no value.
'
'THE NAME HAS NO RANGE BEHIND IT
'-------------------------------------------------------------------------------
'RNG_EpiWeekStart is a HiddenNames entry and HiddenNames writes a string as a
'quoted literal in the name definition. So the name exists, Names(...) finds it,
'and asking that name for a range raises 1004. The write used to go through
'Names(RNG_EpiWeekStart).RefersToRange.Value under a swallow, which means the
'value has never once landed. The store is what answers it, the same way
'FormLogicEpiWeek and CustomLinelistFunctions already read and write it.
'@param meta ImportMetadata. The metadata of the workbook being imported.
'@export
Public Sub ImportSingleValues(ByVal meta As ImportMetadata)

    Dim wkbNames As HiddenNames
    Dim varValue As String

    If meta Is Nothing Then Exit Sub

    varValue = meta.EpiWeekStart
    If LenB(varValue) = 0 Then Exit Sub

    Set wkbNames = HiddenNames.Create(this.sourceWkb)
    wkbNames.SetValue RNGEPIWEEKSTART, varValue

    RecalculateEpiWeekSheets
End Sub

'@label:recalculate-epi-week-sheets
'@sub-title Recompute every sheet whose formulas read the epiweek start
'@details
'Each epiweek formula in the workbook reads RNG_EpiWeekStart, so changing that
'value leaves every one of them showing the old week until the sheet is
'recalculated. FormLogicEpiWeek does the same after the user picks a day, and
'this is the same list of sheet tags.
Private Sub RecalculateEpiWeekSheets()

    Dim sh As Worksheet
    Dim tagValues As BetterArray

    Set tagValues = New BetterArray
    tagValues.Push "HList", "VList", "TS-Analysis", "SP-Analysis", _
                   "Uni-Bi-Analysis", "SPT-Analysis"

    For Each sh In this.sourceWkb.Worksheets
        If tagValues.Includes(SheetTagOf(sh)) Then
            On Error Resume Next
            sh.UsedRange.Calculate
            On Error GoTo 0
        End If
    Next sh
End Sub

'@label:import-geobase
'@sub-title Import geobase data from a workbook
'@details
'Delegates to LLGeo.Import to load geobase data. When not histoOnly,
'also updates geobase headers in HList worksheets via geo.Update and
'dictionary geo labels via geo.UpdateDict.
'
'THE SHOW/HIDE SHEET IS NOT TOUCHED
'-------------------------------------------------------------------------------
'A third call used to write the geo labels into per-sheet tables on the
'show/hide worksheet. Those tables belonged to the first show/hide design and
'ShowHideStore deletes them on sight now, so the call read a sheet carrying
'only the store''s own table and raised 91 on every import. The label reaches
'the form on its own: ShowHide composes it from the dictionary main label that
'geo.UpdateDict has just rewritten, every time the form is built.
'@param impwb Workbook. The opened geobase workbook.
'@param histoOnly Optional Boolean. When True, imports only historic geobase data. Defaults to False.
'@export
Public Sub ImportGeobase(ByVal impwb As Workbook, Optional ByVal histoOnly As Boolean = False)

    Dim geo As LLGeo
    Dim dict As LLdictionary
    Dim pass As Passwords

    Set geo = GeoObject()
    geo.Import impwb, histoOnly

    If Not histoOnly Then
        Set pass = PasswordsObject()
        Set dict = SourceDictionary()

        ' Update geobase headers in HList worksheets
        geo.Update pass
        ' Update dictionary geo variable labels
        geo.UpdateDict dict
    End If
End Sub


'@section Comparing the two files
'===============================================================================

'@label:compare-with-import-file
'@sub-title Compare the labels and the choice lists of the two files
'@details
'Walks the dictionary of this linelist and, for every variable the file also
'carries, files a warning when the two disagree about the label or about the
'values a choice variable may take.
'
'WHY THE CHOICE LISTS MATTER MORE THAN THE LABELS
'-------------------------------------------------------------------------------
'A data validation dropdown does not block a programmatic write. CustomTable
'writes the imported column straight into the cells, so a value the receiving
'linelist does not offer lands in the cell and stays there: the cell shows
'something no dropdown can produce, the next hand edit of it is rejected, and
'every analysis grouping by that variable gains a category the setup never
'declared. None of that used to appear anywhere.
'
'Run this after ImportData, so the sheets that were touched are known.
'@param impwb Workbook. The opened import workbook.
'@export
Public Sub CompareWithImportFile(ByVal impwb As Workbook)

    Dim inpDict As LLdictionary
    Dim currVars As LLVariables
    Dim inpVars As LLVariables
    Dim currChoi As LLChoices
    Dim impChoi As LLChoices
    Dim drop As DropdownLists
    Dim varNames As BetterArray
    Dim varName As String
    Dim counter As Long

    If Not SheetExists(impwb, DICTIONARYSHEET) Then Exit Sub

    Set inpDict = LLdictionary.Create(impwb.Worksheets(DICTIONARYSHEET), 1, 1)
    If Not inpDict.Prepared Then
        FileChecking "import dictionary", _
                     "The Dictionary worksheet of the file has never been " & _
                     "prepared, so the labels and the choice lists of the two " & _
                     "files cannot be compared.", _
                     checkingWarning
        Exit Sub
    End If

    Set currVars = LLVariables.Create(SourceDictionary())
    Set inpVars = LLVariables.Create(inpDict)
    Set varNames = currVars.VariableNames()

    Set currChoi = ChoicesOf(this.sourceWkb)
    Set impChoi = ChoicesOf(impwb)
    Set drop = CustomDropdowns()

    For counter = varNames.LowerBound To varNames.UpperBound
        varName = CStr(varNames.Item(counter))

        If inpVars.Contains(varName) Then
            If SheetWasImported(currVars.SheetName(varName)) Then
                CompareOneLabel currVars, inpVars, varName
                CompareOneChoiceList currVars, inpVars, currChoi, impChoi, drop, varName
            End If
        End If
    Next counter
End Sub

'@label:sheet-was-imported
'@sub-title Whether one sheet of this linelist took data from the file
'@details
'A comparison run on its own, with no import before it, has no list of touched
'sheets and answers True for every sheet.
'@param sheetName String. The sheet to test.
'@return Boolean. True when the sheet is worth comparing.
Private Function SheetWasImported(ByVal sheetName As String) As Boolean
    If this.importedSheets Is Nothing Then
        SheetWasImported = True
        Exit Function
    End If
    If this.importedSheets.Length = 0 Then
        SheetWasImported = True
        Exit Function
    End If

    SheetWasImported = this.importedSheets.Includes(sheetName)
End Function

'@label:compare-one-label
'@sub-title File a warning when the two files give a variable different labels
'@param currVars LLVariables. The variables of this linelist.
'@param inpVars LLVariables. The variables of the file.
'@param varName String. The variable to compare.
Private Sub CompareOneLabel(ByVal currVars As LLVariables, _
                            ByVal inpVars As LLVariables, _
                            ByVal varName As String)

    Dim currLabel As String
    Dim impLabel As String

    currLabel = currVars.Value(colName:="main label", varName:=varName)
    impLabel = inpVars.Value(colName:="main label", varName:=varName)

    If LenB(impLabel) = 0 Then Exit Sub
    If StrComp(Trim$(currLabel), Trim$(impLabel), vbBinaryCompare) = 0 Then Exit Sub

    FileChecking "label of " & varName, _
                 "The label of " & varName & " differs. This linelist reads " & _
                 Chr$(34) & currLabel & Chr$(34) & " and the file was written " & _
                 "with " & Chr$(34) & impLabel & Chr$(34) & ".", _
                 checkingWarning
End Sub

'@label:compare-one-choice-list
'@sub-title File a warning when the two files offer a choice variable different values
'@details
'`choice_formula` is left alone: its list name is computed from a formula and
'the routine that reads that formula lives in the setup, which a linelist does
'not carry.
'@param currVars LLVariables. The variables of this linelist.
'@param inpVars LLVariables. The variables of the file.
'@param currChoi LLChoices. The Choices worksheet of this linelist.
'@param impChoi LLChoices. The Choices worksheet of the file.
'@param drop DropdownLists. The custom dropdowns of this linelist.
'@param varName String. The variable to compare.
Private Sub CompareOneChoiceList(ByVal currVars As LLVariables, _
                                 ByVal inpVars As LLVariables, _
                                 ByVal currChoi As LLChoices, _
                                 ByVal impChoi As LLChoices, _
                                 ByVal drop As DropdownLists, _
                                 ByVal varName As String)

    Dim controlValue As String
    Dim currList As String
    Dim impList As String
    Dim currCats As BetterArray
    Dim impCats As BetterArray
    Dim keys As ChoiceKeys
    Dim onlyImport As String
    Dim onlyCurrent As String

    If impChoi Is Nothing Then Exit Sub

    controlValue = LCase$(Trim$(currVars.ControlType(varName)))
    If controlValue <> "choice_manual" Then
        If controlValue <> "choice_multiple" Then
            If controlValue <> "choice_custom" Then Exit Sub
        End If
    End If

    currList = currVars.Value(colName:="control details", varName:=varName)
    impList = inpVars.Value(colName:="control details", varName:=varName)
    If LenB(currList) = 0 Then Exit Sub
    If LenB(impList) = 0 Then Exit Sub

    Set keys = ChoiceKeys.Create()

    If controlValue = "choice_custom" Then
        ' The values of a custom dropdown live on the dropdown worksheet in this
        ' linelist and on the Choices worksheet in the file.
        If drop Is Nothing Then Exit Sub
        Set currCats = drop.Values(currList)
        Set impCats = impChoi.Categories(keys.CustomChoiceName(impList))
    Else
        If currChoi Is Nothing Then Exit Sub
        Set currCats = currChoi.Categories(currList)
        Set impCats = impChoi.Categories(impList)
    End If

    If currCats.Length = 0 Then Exit Sub
    If impCats.Length = 0 Then Exit Sub

    onlyImport = ValuesMissingFrom(impCats, currCats)
    onlyCurrent = ValuesMissingFrom(currCats, impCats)

    If LenB(onlyImport) > 0 Then
        FileChecking "choices of " & varName, _
                     "The file offers values for " & varName & " that this " & _
                     "linelist does not: " & onlyImport & ". Imported rows can " & _
                     "carry them, and no dropdown of this linelist can produce " & _
                     "them.", _
                     checkingWarning
    End If

    If LenB(onlyCurrent) > 0 Then
        FileChecking "choices of " & varName & " missing from the file", _
                     "This linelist offers values for " & varName & " that the " & _
                     "file does not: " & onlyCurrent & ".", _
                     checkingWarning
    End If
End Sub

'@label:values-missing-from
'@sub-title The values of one list that the other does not hold
'@details
'Values are matched without regard to case and to surrounding spaces, and they
'are shown as they were written, so "Yes" against "yes" is reported as a
'difference a person can act on. The list is capped, because a user reads a
'handful and the count is what tells them the size.
'@param source BetterArray. The list to walk.
'@param other BetterArray. The list to look each value up in.
'@return String. The missing values, joined, empty when there are none.
Private Function ValuesMissingFrom(ByVal source As BetterArray, _
                                   ByVal other As BetterArray) As String

    Const MAX_SHOWN As Long = 8

    Dim missing As BetterArray
    Dim oneValue As String
    Dim counter As Long
    Dim shown As Long
    Dim answer As String

    Set missing = New BetterArray
    missing.LowerBound = 1

    For counter = source.LowerBound To source.UpperBound
        oneValue = Trim$(CStr(source.Item(counter)))
        If LenB(oneValue) > 0 Then
            If Not HoldsValue(other, oneValue) Then missing.Push oneValue
        End If
    Next counter

    If missing.Length = 0 Then Exit Function

    shown = missing.Length
    If shown > MAX_SHOWN Then shown = MAX_SHOWN

    For counter = missing.LowerBound To missing.LowerBound + shown - 1
        If LenB(answer) > 0 Then answer = answer & ", "
        answer = answer & Chr$(34) & CStr(missing.Item(counter)) & Chr$(34)
    Next counter

    If missing.Length > shown Then
        answer = answer & " and " & CStr(missing.Length - shown) & " more"
    End If

    ValuesMissingFrom = answer
End Function

'@label:holds-value
'@sub-title Whether a list holds one value, ignoring case and outer spaces
'@param listValues BetterArray. The list to search.
'@param oneValue String. The value to look for.
'@return Boolean. True when the list holds it.
Private Function HoldsValue(ByVal listValues As BetterArray, _
                            ByVal oneValue As String) As Boolean

    Dim counter As Long

    For counter = listValues.LowerBound To listValues.UpperBound
        If StrComp(Trim$(CStr(listValues.Item(counter))), oneValue, vbTextCompare) = 0 Then
            HoldsValue = True
            Exit Function
        End If
    Next counter
End Function

'@label:choices-of
'@sub-title The Choices worksheet of one workbook
'@param wb Workbook. The workbook to read.
'@return LLChoices. The choices, or Nothing when the workbook carries none.
Private Function ChoicesOf(ByVal wb As Workbook) As LLChoices
    If Not SheetExists(wb, CHOICESSHEET) Then Exit Function
    Set ChoicesOf = LLChoices.Create(wb.Worksheets(CHOICESSHEET), 1, 1)
End Function

'@label:custom-dropdowns
'@sub-title The custom dropdown manager of this linelist
'@return DropdownLists. The manager, or Nothing when the workbook has no custom
'  dropdown worksheet.
Private Function CustomDropdowns() As DropdownLists

    Dim wkbNames As HiddenNames
    Dim dropsh As Worksheet

    Set wkbNames = HiddenNames.Create(this.sourceWkb)

    On Error Resume Next
    Set dropsh = this.sourceWkb.Worksheets(wkbNames.ValueAsString("RNG_CustomDrop"))
    On Error GoTo 0
    If dropsh Is Nothing Then Exit Function

    Set CustomDropdowns = DropdownLists.Create(dropsh, vbNullString)
End Function


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

    Dim dict As LLdictionary
    Dim vars As LLVariables
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

        'matchCase defaults to True on FilterData, and a dictionary whose
        'headers carry the setup file's own capitals then raises
        'ElementNotFound out of a column that is plainly there. A header name
        'is a contract resolved by name; its case decides nothing.
        Set currVarsList = dict.Data.FilterData("sheet name", sheetName, _
                                                "variable name", matchCase:=False)

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

    FileReportCheckings
    SaveReport
End Sub

'@label:save-report
'@sub-title Keep the report where it survives the workbook closing
'@details
'The four lists used to live in memory alone and die with this object at the end
'of the import, while the four ListObjects they belong on sat on __import_rep
'with nothing writing to them. They are written now, so F_ImportRep can be
'opened again days later from the advanced form.
'
'A workbook that will not take the store - anything other than a generated
'linelist - leaves the lists in memory and says so.
Private Sub SaveReport()

    Dim store As ImportReport
    Dim errNumber As Long
    Dim errMessage As String

    Set store = Nothing

    'Keeping the report is best effort from end to end. A workbook that will
    'not take the store, a worksheet that will not take a ListObject, a table
    'that will not resize - none of them is a reason to fail an import whose
    'data has already landed. What went wrong is filed for the user instead.
    On Error Resume Next
        Set store = ImportReport.Create(this.sourceWkb)

        If Not store Is Nothing Then
            store.SaveSheets ImportReportNotImported, this.sheetsNotImported
            store.SaveSheets ImportReportNotTouched, this.sheetsNotTouched
            store.SaveVariables ImportReportNotImported, this.varsNotImported
            store.SaveVariables ImportReportNotTouched, this.varsNotTouched
        End If

        errNumber = Err.Number
        errMessage = Err.Description
    On Error GoTo 0

    If store Is Nothing Or errNumber <> 0 Then
        FileChecking "import report store", _
                     "The report of this import could not be kept in the " & _
                     "workbook, so it is gone once this workbook closes. " & _
                     errMessage, _
                     checkingWarning
    End If
End Sub

'@label:file-report-checkings
'@sub-title Put the four report lists where the user can read them
'@details
'The four lists are the report the form shows. They are filed as checkings as
'well, so one worksheet carries everything the import found and the user can
'read it without opening a form.
Private Sub FileReportCheckings()

    Dim counter As Long
    Dim entry As Variant

    If this.sheetsNotImported Is Nothing Then Set this.sheetsNotImported = New BetterArray
    If this.sheetsNotTouched Is Nothing Then Set this.sheetsNotTouched = New BetterArray
    If this.varsNotImported Is Nothing Then Set this.varsNotImported = New BetterArray
    If this.varsNotTouched Is Nothing Then Set this.varsNotTouched = New BetterArray

    For counter = this.sheetsNotImported.LowerBound To this.sheetsNotImported.UpperBound
        FileChecking "sheet not imported: " & CStr(this.sheetsNotImported.Item(counter)), _
                     "The file carries a worksheet named " & _
                     CStr(this.sheetsNotImported.Item(counter)) & _
                     " and this linelist has none, so none of its data was read.", _
                     checkingWarning
    Next counter

    For counter = this.sheetsNotTouched.LowerBound To this.sheetsNotTouched.UpperBound
        FileChecking "sheet not touched: " & CStr(this.sheetsNotTouched.Item(counter)), _
                     "The worksheet " & CStr(this.sheetsNotTouched.Item(counter)) & _
                     " of this linelist took no data, because the file carries " & _
                     "no worksheet of that name.", _
                     checkingWarning
    Next counter

    For counter = this.varsNotImported.LowerBound To this.varsNotImported.UpperBound
        entry = this.varsNotImported.Item(counter)
        FileChecking "variable not imported: " & CStr(entry(LBound(entry))), _
                     "The file carries a variable named " & _
                     CStr(entry(LBound(entry))) & " on " & _
                     CStr(entry(LBound(entry) + 1)) & " and this linelist has " & _
                     "none, so its values were dropped.", _
                     checkingWarning
    Next counter

    For counter = this.varsNotTouched.LowerBound To this.varsNotTouched.UpperBound
        entry = this.varsNotTouched.Item(counter)
        FileChecking "variable not touched: " & CStr(entry(LBound(entry))), _
                     "The variable " & CStr(entry(LBound(entry))) & " on " & _
                     CStr(entry(LBound(entry) + 1)) & " took no data, because " & _
                     "the file carries no variable of that name.", _
                     checkingWarning
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
'Clears all report tracking arrays, drops the diagnostics, and resets the
'needReport flag. Call before starting a new import session to ensure a clean
'report.
'@export
Public Sub ClearReport()
    this.needReport = False
    Set this.sheetsNotImported = New BetterArray
    Set this.sheetsNotTouched = New BetterArray
    Set this.varsNotImported = New BetterArray
    Set this.varsNotTouched = New BetterArray
    Set this.importedSheets = New BetterArray
    Set this.importedVars = New BetterArray
    Set this.checks = Nothing
End Sub


'@section Private - Source Object Accessors
'===============================================================================

'@label:source-dictionary
'@sub-title Get the dictionary from the source workbook
'@return LLdictionary. Dictionary instance from the Dictionary sheet.
Private Function SourceDictionary() As LLdictionary
    Set SourceDictionary = LLdictionary.Create( _
        this.sourceWkb.Worksheets(DICTIONARYSHEET), 1, 1)
End Function

'@label:passwords-object
'@sub-title Get passwords from the source workbook
'@return Passwords. Passwords instance from the __pass sheet.
Private Function PasswordsObject() As Passwords
    Set PasswordsObject = Passwords.Create(this.sourceWkb.Worksheets(PASSWORDSHEET))
End Function

'@label:geo-object
'@sub-title Get the geo object from the source workbook
'@return LLGeo. Geo instance from the Geo sheet.
Private Function GeoObject() As LLGeo
    Set GeoObject = LLGeo.Create(this.sourceWkb.Worksheets(GEOSHEET))
End Function


'@section Private - Import Helpers
'===============================================================================

'@label:initialize-tracking
'@sub-title Initialize tracking state for a new import session
'@details
'The two "not touched" lists are left alone because FinalizeReport replaces
'them, and the diagnostics are left alone because CheckImportFile runs before
'this and its entries belong to the same import.
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
                      (sheetName = LINELISTTRANSLATIONSHEET) Or _
                      (sheetName = CHOICESSHEET) Or _
                      (sheetName = METADATASHEET) Or _
                      (sheetName = SHOWHIDESHEET) Or _
                      (sheetName = SHOWHIDESHEET_OLD)
End Function

'@label:show-hide-sheet-of
'@sub-title The show/hide worksheet of a workbook, or Nothing
'@details
'A file written before the internal-sheet rename carries the old trailing
'name, so the new name is looked for first and the old one answers when it
'is absent.
'@param wb Workbook. The workbook to look through.
'@return Worksheet. The show/hide worksheet, or Nothing when the workbook has none.
Private Function ShowHideSheetOf(ByVal wb As Workbook) As Worksheet
    Dim sh As Worksheet

    'A missing sheet raises on the read, and Nothing is the answer read back.
    On Error Resume Next
    Set sh = wb.Worksheets(SHOWHIDESHEET)
    If sh Is Nothing Then Set sh = wb.Worksheets(SHOWHIDESHEET_OLD)
    On Error GoTo 0

    Set ShowHideSheetOf = sh
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
    Dim counter As Long

    Set usedRng = impsh.UsedRange

    For counter = 2 To usedRng.Rows.Count
        varName = CStr(usedRng.Cells(counter, 1).Value)

        Set vlistRng = Nothing
        On Error Resume Next
        Set vlistRng = currsh.Range(varName)
        On Error GoTo 0

        If vlistRng Is Nothing Then
            ' Variable from import not found in current sheet
            this.varsNotImported.Push Array(varName, impsh.Name)
            this.needReport = True
        Else
            If Not vlistRng.HasFormula Then
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
'@param headerRow Long. The row of the import sheet that holds the variable
'  names. It was read as row 1 always, so a file written with the labels above
'  the names had its label row read as the header row, and every column of it
'  landed in "not imported".
Private Sub ImportHListSheet(ByVal currsh As Worksheet, ByVal impsh As Worksheet, _
                             ByVal pasteAtBottom As Boolean, ByVal headerRow As Long)

    Dim currTab As CustomTable
    Dim impTab As DataSheet
    Dim varsNotImpList As BetterArray
    Dim impHeaders As BetterArray
    Dim counter As Long

    Set currTab = CustomTable.Create(currsh.ListObjects(1))
    Set impTab = DataSheet.Create(impsh, headerRow, 1, True)

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


'@sub-title Get the sheet type tag from worksheet-level HiddenNames.
'@param sh Worksheet. The worksheet to query.
'@return String. The sheet type tag (HList, VList, etc.).
'@param nameKey Optional String. Which worksheet-level name to read.
Private Function SheetTagOf(ByVal sh As Worksheet, _
                            Optional ByVal nameKey As String = "sheet_type") As String
    Dim shHn As HiddenNames
    Set shHn = HiddenNames.Create(sh)
    SheetTagOf = shHn.ValueAsString(nameKey)
End Function

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
